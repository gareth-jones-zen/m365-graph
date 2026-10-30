"""A small Microsoft Graph client: sign in once, cache the token, make requests that don't hang.

    from m365_graph import Client
    graph = Client(cache_path=..., login_hint="you@example.com")
    me = graph.get("/me")
    for message in graph.pages("/teams/{id}/channels/{id}/messages", scopes=["ChannelMessage.Read.All"]):
        ...

What it's for, from running several unattended tools against Graph:
- **One sign-in per process and per cache file.** Every Client built on the same cache file shares the
  MSAL app and its tokens, so a tool that makes fifty calls signs in once.
- **No browser when nobody's there.** A run under Task Scheduler has no console; a browser sign-in
  would open where nobody sees it and wait forever. Unattended, a missing sign-in raises
  SignInNeeded instead, and the caller can notify someone.
- **Scopes per call.** Ask only for what a call needs, so a new consent for one tool never breaks another.
- **A timeout and retries** for stalls, 429s and 5xx. A stalled connection once hung a run for 30 minutes.
- **IPv4 first** (net.py), which fixed sign-ins that took minutes.
"""

from __future__ import annotations

import json
import os
import pathlib
import sys
import time
import urllib.error
import urllib.request

import msal

GRAPH = "https://graph.microsoft.com/v1.0"
# Microsoft's own public "Microsoft Graph Command Line Tools" app. Works in most tenants without
# registering an app; an admin may still have to approve scopes that need admin consent.
GRAPH_CLI_CLIENT = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
DEFAULT_CACHE = pathlib.Path(os.environ.get("LOCALAPPDATA", pathlib.Path.home())) / "m365-graph" / "msal-cache.json"
TIMEOUT = 60
RETRIES = 3
SLOW = 15


class SignInNeeded(PermissionError, RuntimeError):
    """No usable sign-in for these scopes, and none can be done now. A PermissionError (what callers
    of the earlier code caught) and a RuntimeError (what teams-transcripts caught)."""


def unattended() -> bool:
    """Under Task Scheduler (pythonw) there's no console, and nobody to complete a browser sign-in."""
    return sys.stdin is None or not sys.stdin.isatty()


_APPS: dict = {}     # (cache file, client, tenant) -> (MSAL app, its token cache, {scopes: (token, expires)})


class Client:
    def __init__(self, cache_path: os.PathLike | str | None = None, tenant: str = "organizations",
                 client_id: str = GRAPH_CLI_CLIENT, login_hint: str | None = None,
                 scopes: list | tuple = ("User.Read",), interactive: bool | None = None,
                 sign_in_hint: str = "Sign in again at a console: python -m m365_graph sign-in",
                 log=None):
        """interactive: True may open a browser; False never does (raises SignInNeeded); None decides
        by whether anyone is at the console. sign_in_hint is shown in SignInNeeded, so each tool can
        say how its own users sign in. log(str) receives slow-call and retry notices (default stderr)."""
        self.cache_path = pathlib.Path(cache_path) if cache_path else DEFAULT_CACHE
        self.tenant, self.client_id, self.login_hint = tenant, client_id, login_hint
        self.scopes, self.interactive, self.sign_in_hint = list(scopes), interactive, sign_in_hint
        self.log = log or (lambda text: print(text, file=sys.stderr, flush=True))

    # ---- sign-in ------------------------------------------------------------------------------

    def _app(self):
        key = (str(self.cache_path.resolve()).lower(), self.client_id, self.tenant)
        if key not in _APPS:
            cache = msal.SerializableTokenCache()
            if self.cache_path.exists():
                cache.deserialize(self.cache_path.read_text(encoding="utf-8"))
            app = msal.PublicClientApplication(
                self.client_id, authority=f"https://login.microsoftonline.com/{self.tenant}",
                token_cache=cache, timeout=TIMEOUT)
            _APPS[key] = (app, cache, {})
        return _APPS[key]

    def token(self, scopes: list | tuple | None = None, interactive: bool | None = None) -> str:
        scopes = sorted(scopes or self.scopes)
        app, cache, tokens = self._app()
        held = tokens.get(tuple(scopes))
        if held and held[1] > time.time() + 300:
            return held[0]
        if interactive is None:
            interactive = self.interactive
        if interactive is None:
            interactive = not unattended()

        result = None
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            # A scope never consented to comes back as an error dict, not None.
            if result and "access_token" not in result:
                result = None
        if not result:
            if not interactive:
                raise SignInNeeded(f"No cached sign-in for {', '.join(scopes)}, or it has expired. {self.sign_in_hint}")
            self.log("Opening a browser to sign in to Microsoft...")
            result = app.acquire_token_interactive(scopes, login_hint=self.login_hint)
        if cache.has_state_changed:
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            self.cache_path.write_text(cache.serialize(), encoding="utf-8")
            cache.has_state_changed = False
        if "access_token" not in result:
            raise SignInNeeded(f"Sign-in failed: {result.get('error')}: {result.get('error_description', '')[:300]}")
        tokens[tuple(scopes)] = (result["access_token"], time.time() + int(result.get("expires_in", 3600)))
        return result["access_token"]

    # ---- requests -----------------------------------------------------------------------------

    def request(self, url: str, method: str = "GET", body: dict | None = None, scopes: list | tuple | None = None,
                headers: dict | None = None, interactive: bool | None = None) -> dict:
        """One Graph request. url may be relative ("/me") or a full URL (such as an @odata.nextLink)."""
        if not url.startswith("http"):
            url = GRAPH + url
        data = json.dumps(body).encode("utf-8") if body is not None else None
        for attempt in range(1, RETRIES + 1):
            request = urllib.request.Request(url, data=data, method=method, headers={
                "Authorization": f"Bearer {self.token(scopes, interactive)}",
                "Content-Type": "application/json", **(headers or {})})
            started = time.time()
            try:
                with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
                    raw = response.read()
                if time.time() - started > SLOW:
                    self.log(f"slow Graph call: {time.time() - started:.0f}s {method} {url[:120]}")
                return json.loads(raw) if raw else {}
            except urllib.error.HTTPError as error:
                if error.code not in (429, 500, 502, 503, 504) or attempt == RETRIES:
                    raise
                wait = int(error.headers.get("Retry-After") or 5 * attempt)
            except (urllib.error.URLError, TimeoutError):
                if attempt == RETRIES:
                    raise
                wait = 5 * attempt
            self.log(f"Graph retry {attempt} in {wait}s: {method} {url[:120]}")
            time.sleep(wait)

    def get(self, url: str, scopes: list | tuple | None = None, headers: dict | None = None) -> dict:
        return self.request(url, scopes=scopes, headers=headers)

    def pages(self, url: str, scopes: list | tuple | None = None, headers: dict | None = None, stop=None):
        """Every item across @odata.nextLink pages. stop(item) -> True ends after the current page."""
        while url:
            body = self.request(url, scopes=scopes, headers=headers)
            done = False
            for item in body.get("value", []):
                if stop and stop(item):
                    done = True
                yield item
            url = None if done else body.get("@odata.nextLink")
