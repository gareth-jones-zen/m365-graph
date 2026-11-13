"""A write that may have reached Graph is never sent twice (13 Nov 2026, during a Microsoft outage)."""

import io
import socket
import unittest
import urllib.error
from unittest import mock

import m365_graph
from m365_graph import client as graph_client


def http_error(code: int) -> urllib.error.HTTPError:
    return urllib.error.HTTPError("https://graph", code, "x", {}, io.BytesIO(b""))


class Response:
    def __init__(self, body: bytes = b'{"ok": true}'):
        self.body = body

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def read(self):
        return self.body


class Fake:
    """Answers each urlopen call in turn: an exception to raise, or a body to return."""

    def __init__(self, *answers):
        self.answers = list(answers)
        self.calls = []

    def __call__(self, request, timeout=None):
        self.calls.append(request.get_method())
        answer = self.answers.pop(0)
        if isinstance(answer, BaseException):
            raise answer
        return Response(answer)


def run(fake, url="/me/events", method="POST", **kwargs):
    graph = graph_client.Client.__new__(graph_client.Client)
    graph.log = lambda message: None
    graph.token = lambda scopes=None, interactive=None: "token"
    with mock.patch.object(graph_client.urllib.request, "urlopen", fake), \
         mock.patch.object(graph_client.time, "sleep", lambda seconds: None):
        return graph.request(url, method=method, body={} if method != "GET" else None, **kwargs)


class Rules(unittest.TestCase):
    def test_idempotent(self):
        for method, url in [("GET", "/me"), ("DELETE", "/me/events/1"), ("PUT", "/x"),
                            ("PATCH", "/me/events/1"), ("PATCH", "/me/messages/1"),
                            ("POST", "/me/calendar/getSchedule"), ("POST", "https://graph/v1.0/me/findMeetingTimes?x=1")]:
            self.assertTrue(m365_graph.idempotent(method, url), (method, url))
        for method, url in [("POST", "/me/events"), ("POST", "/me/events/1/accept"), ("POST", "/me/messages/1/move"),
                            ("PATCH", "/users/a/onenote/pages/1/content"), ("POST", "/me/onenote/sections/1/pages")]:
            self.assertFalse(m365_graph.idempotent(method, url), (method, url))

    def test_never_sent(self):
        self.assertTrue(m365_graph.never_sent(urllib.error.URLError(ConnectionRefusedError())))
        self.assertTrue(m365_graph.never_sent(urllib.error.URLError(socket.gaierror())))
        self.assertFalse(m365_graph.never_sent(urllib.error.URLError(ConnectionResetError())))
        self.assertFalse(m365_graph.never_sent(TimeoutError()))
        self.assertFalse(m365_graph.never_sent(http_error(502)))


class Requests(unittest.TestCase):
    def test_get_retried_on_502(self):
        fake = Fake(http_error(502), http_error(504), b'{"v": 1}')
        self.assertEqual(run(fake, "/me", "GET"), {"v": 1})
        self.assertEqual(len(fake.calls), 3)

    def test_post_not_resent_after_502(self):
        for error in (http_error(502), http_error(504), http_error(500), http_error(503),
                      TimeoutError(), urllib.error.URLError(ConnectionResetError())):
            fake = Fake(error, b"{}")
            with self.assertRaises(type(error)):
                run(fake)
            self.assertEqual(len(fake.calls), 1, error)

    def test_post_resent_after_429(self):
        fake = Fake(http_error(429), b'{"id": "e"}')
        self.assertEqual(run(fake), {"id": "e"})
        self.assertEqual(len(fake.calls), 2)

    def test_post_resent_when_never_connected(self):
        fake = Fake(urllib.error.URLError(ConnectionRefusedError()), urllib.error.URLError(socket.gaierror()), b"{}")
        self.assertEqual(run(fake), {})
        self.assertEqual(len(fake.calls), 3)

    def test_post_429_then_502_stops(self):
        fake = Fake(http_error(429), http_error(502), b"{}")
        with self.assertRaises(urllib.error.HTTPError):
            run(fake)
        self.assertEqual(len(fake.calls), 2)

    def test_read_only_post_retried(self):
        fake = Fake(http_error(502), b'{"value": []}')
        self.assertEqual(run(fake, "/me/calendar/getSchedule"), {"value": []})

    def test_patch_retried_but_not_onenote_content(self):
        fake = Fake(http_error(504), b"{}")
        self.assertEqual(run(fake, "/me/events/1", "PATCH"), {})
        fake = Fake(http_error(504), b"{}")
        with self.assertRaises(urllib.error.HTTPError):
            run(fake, "/me/onenote/pages/1/content", "PATCH")

    def test_delete_404_after_uncertain_attempt_is_done(self):
        fake = Fake(http_error(504), http_error(404))
        self.assertEqual(run(fake, "/me/events/1", "DELETE"), {})

    def test_delete_404_first_time_still_raises(self):
        with self.assertRaises(urllib.error.HTTPError):
            run(Fake(http_error(404)), "/me/events/1", "DELETE")

    def test_override(self):
        fake = Fake(http_error(502), b"{}")
        self.assertEqual(run(fake, retry_writes=True), {})
        fake = Fake(http_error(502), b"{}")
        with self.assertRaises(urllib.error.HTTPError):
            run(fake, "/me/events/1", "PATCH", retry_writes=False)

    def test_last_attempt_still_raises(self):
        with self.assertRaises(urllib.error.HTTPError):
            run(Fake(http_error(502), http_error(502), http_error(502)), "/me", "GET")


if __name__ == "__main__":
    unittest.main()
