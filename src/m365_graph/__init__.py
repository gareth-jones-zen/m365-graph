"""m365-graph: a small Microsoft Graph client shared by teams-transcripts, teams-channel-notes and others.

Importing it makes connections prefer IPv4 (see net.py).
"""

from . import net  # noqa: F401  (prefer IPv4, process-wide)
from .client import DEFAULT_CACHE, GRAPH, GRAPH_CLI_CLIENT, Client, SignInNeeded, unattended

__version__ = "0.1.0"
__all__ = ["Client", "SignInNeeded", "unattended", "GRAPH", "GRAPH_CLI_CLIENT", "DEFAULT_CACHE"]
