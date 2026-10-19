"""Child interpreter for the gate exit test.

argv: marker_path mode. "deadline": a target that never returns under a
0.2 second deadline. "no-deadline": a body that finishes after its caller
has stopped waiting. No COM: the apartment initializer is replaced.
"""
import asyncio
import atexit
import sys
import threading
import time

marker, mode = sys.argv[1], sys.argv[2]

from msaccess_vcs_mcp import access_gate  # noqa: E402

access_gate._init_com_apartment = lambda: None
access_gate.EXIT_GRACE_SEC = 0.5


def note(text):
    with open(marker, "a", encoding="utf-8") as handle:
        handle.write(text + "\n")


atexit.register(lambda: note("atexit-ran"))
access_gate._record_abandoned = lambda attempt: note(
    "abandoned " + str(attempt.snapshot().get("dispatch"))
)


class App:
    def Run(self, *_args):
        threading.Event().wait()


def deadline_body():
    attempt = access_gate.current_attempt()
    attempt.begin_validation()
    permit = attempt.permit_dispatch()
    permit.run(App(), "Target")
    note("body-finished")


def slow_body():
    time.sleep(0.5)
    note("body-finished")


async def main():
    gate = access_gate.get_access_gate()
    if mode == "deadline":
        reply = await gate.run_exclusive("vcs_call_vba", None, deadline_body, False, deadline_seconds=0.2)
        note("returned " + reply["error_pattern"])
    else:
        try:
            await asyncio.wait_for(gate.run_exclusive("vcs_export_database", None, slow_body, False), 0.05)
        except TimeoutError:
            note("caller-left")


asyncio.run(main())
