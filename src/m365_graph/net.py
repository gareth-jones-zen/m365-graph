"""Prefer IPv4 when connecting.

Python tries a host's addresses one at a time, in the order DNS gives them, and waits the full
timeout on each. login.microsoftonline.com returns eight IPv6 addresses before its IPv4 ones, so on
a network where IPv6 to Microsoft goes nowhere, every sign-in stalled for minutes before reaching
IPv4 (22 Oct 2026: a fetch hung for over 25 minutes in MSAL's tenant discovery). Browsers and curl
race the two families, which is why nothing else showed it.

Sorting IPv4 first keeps IPv6 as the fallback, so an IPv6-only host still works. Importing
m365_graph applies it, process-wide.
"""

import socket

_getaddrinfo = socket.getaddrinfo


def _ipv4_first(*args, **kwargs):
    return sorted(_getaddrinfo(*args, **kwargs), key=lambda info: info[0] != socket.AF_INET)


def prefer_ipv4() -> None:
    if getattr(socket.getaddrinfo, "__name__", "") != "_ipv4_first":
        socket.getaddrinfo = _ipv4_first


prefer_ipv4()
