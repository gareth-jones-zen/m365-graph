"""Sign in, or check who you're signed in as.

    python -m m365_graph sign-in                         # User.Read, into the default cache
    python -m m365_graph sign-in ChannelMessage.Read.All # approve another scope
    python -m m365_graph whoami
    python -m m365_graph sign-in --cache C:\\path\\msal-cache.json --tenant <tenant id> --login-hint you@example.com

Run it in a real console: a browser opens to sign in or approve a permission. A permission that
needs admin consent shows "Approval required"; request it there, and an admin approves it.
"""

import argparse
import sys

from .client import DEFAULT_CACHE, Client, SignInNeeded


def main() -> int:
    parser = argparse.ArgumentParser(prog="python -m m365_graph", description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("command", choices=["sign-in", "whoami"])
    parser.add_argument("scopes", nargs="*", help="permissions to sign in for (default User.Read)")
    parser.add_argument("--cache", default=str(DEFAULT_CACHE), help=f"token cache file (default {DEFAULT_CACHE})")
    parser.add_argument("--tenant", default="organizations")
    parser.add_argument("--login-hint")
    args = parser.parse_args()

    client = Client(cache_path=args.cache, tenant=args.tenant, login_hint=args.login_hint,
                    scopes=args.scopes or ["User.Read"], interactive=(args.command == "sign-in"))
    try:
        if args.command == "sign-in":
            client.token()
            print(f"Signed in for {', '.join(args.scopes or ['User.Read'])}. Cache: {args.cache}")
        me = client.get("/me?$select=displayName,userPrincipalName", scopes=["User.Read"])
        print(f"{me.get('displayName')} <{me.get('userPrincipalName')}>")
        return 0
    except SignInNeeded as error:
        print(error, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
