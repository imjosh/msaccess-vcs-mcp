# Setup acknowledgments

`vcs_run_tests` dispatches `RunFilteredTests` only after `SetOption(DefaultTestFilter)` returns JSON `success: true`. This applies to both callback and synchronous paths. With no filter, it still calls `SetOption` with an empty string to clear a previous filter and run all tests.

When a session ID is present, `vcs_set_option` writes an option only after `RegisterSession` returns JSON `success: true`. Setup failures preserve the add-in error and refusal fields. Empty, malformed, missing, and non-boolean acknowledgments fail without dispatching the dependent action.

Shutdown-time `EndSession` and `RegisterSession` calls ignore their results by design: shutdown cleanup is best effort, with no dependent user action to authorize. The current shutdown implementation (`main._cleanup_session`) calls only `EndSession`; there is no shutdown-time `RegisterSession` call in the current tree. This remains distinct from registration before an option write, which requires acknowledgment.
