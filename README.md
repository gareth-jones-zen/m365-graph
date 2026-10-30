# m365-graph

A small Microsoft Graph client for Python tools: sign in once, cache the token, and make
requests that don't hang. Shared by [teams-transcripts](https://github.com/gareth-jones-zen/teams-transcripts)
and [teams-channel-notes](https://github.com/gareth-jones-zen/teams-channel-notes), so a fix
lands in one place.

```python
from m365_graph import Client, SignInNeeded

graph = Client(login_hint="you@example.com")          # cache in %LOCALAPPDATA%\m365-graph
me = graph.get("/me")
for message in graph.pages(f"/teams/{team}/channels/{channel}/messages",
                           scopes=["ChannelMessage.Read.All"]):
    ...
```

## Install

```
pip install git+https://github.com/gareth-jones-zen/m365-graph
```

Or, from a checkout, `pip install -e .` so changes apply at once. Needs Python 3.11+ and `msal`.

## What it does, and why

Each point came from running unattended tools against Graph.

- **It signs in once per process.** Every `Client` on the same cache file shares one MSAL app and its tokens, so a tool that makes fifty calls signs in once.
- **No browser when nobody's at the console.** Under Task Scheduler there's no console. A browser sign-in would open where nobody sees it and wait for ever, so instead a missing sign-in raises **`SignInNeeded`**, and the tool can tell someone. `interactive=True` always allows a browser, `False` never does, and `None` (the default) decides by whether there's a console.
- **Scopes are per call.** A call asks only for the permissions it needs, so consenting to something new for one tool never breaks another.
- **It times out and retries.** Each request has a 60-second timeout. Stalls, 429s and 5xx get three tries, honouring `Retry-After`, and calls over 15 seconds are logged. A stalled connection once hung a run for 30 minutes.
- **It prefers IPv4.** Importing the package sorts IPv4 addresses first, process-wide. `login.microsoftonline.com` lists eight IPv6 addresses first, and where IPv6 to Microsoft goes nowhere, Python waited the full timeout on each. Sign-ins took minutes until this; browsers and curl race the two, which hid it.

**`SignInNeeded`** is both a `PermissionError` and a `RuntimeError`, so code written against either earlier convention still catches it. Its message ends with `sign_in_hint`, which each tool sets to say how its own users sign in.

## The app it signs in as

By default it signs in as **Microsoft's own public "Microsoft Graph Command Line Tools" app**
(`14d82eec-204b-4c2f-b7e8-296a70dab67e`), so nothing needs registering.
- Scopes that need admin consent show **"Approval required"** at sign-in. Request it there, and an admin approves it.
- Some tenants grant a scope tenant-wide for this app (`ChannelMessage.Read.All` was at one). A token can then be issued with no prompt at all.
- Pass `client_id` and `tenant` to use your own app registration instead.

## Signing in by hand

```
python -m m365_graph sign-in                          # User.Read
python -m m365_graph sign-in ChannelMessage.Read.All  # approve another permission
python -m m365_graph whoami
python -m m365_graph sign-in --cache <path> --tenant <id> --login-hint you@example.com
```

Run it in a real console: a browser opens.
