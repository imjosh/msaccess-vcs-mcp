# Setup acknowledgments

`vcs_run_tests` dispatches `RunFilteredTests` only after `SetOption(DefaultTestFilter)` returns JSON `success: true`. This applies to both callback and synchronous paths. With no filter, it still calls `SetOption` with an empty string to clear a previous filter and run all tests.

When a session ID is present, `vcs_set_option` writes an option only after `RegisterSession` returns JSON `success: true`. Setup failures preserve the add-in error and refusal fields. Empty, malformed, missing, and non-boolean acknowledgments fail without dispatching the dependent action.

Shutdown-time `EndSession` and `RegisterSession` calls ignore their results by design: shutdown cleanup is best effort, with no dependent user action to authorize. The current shutdown implementation (`main._cleanup_session`) calls only `EndSession`; there is no shutdown-time `RegisterSession` call in the current tree. This remains distinct from registration before an option write, which requires acknowledgment.

## Request policy migration

Policies are trailing operation arguments. MCP does not set or clear a persistent
policy or select process-wide interaction mode before dispatch. An omitted policy
retains legacy behavior. Existing option and compatibility-session acknowledgments
above remain required. `APIAsync` can refuse synchronously; consume that refusal
without waiting for a callback or retrying through synchronous dispatch.

Production also requires the admitted identity's request-policy marker before
sending the new argument slot. See [the migration gate](DIALOGS.md#production-request-policy-migration-gate).
This guard preserves existing version and session checks. It does not claim
that the current production add-in implements the revised API.

Synchronous MCP operation methods use an independent public operation ID through
`APIExecuteMCP` once the admitted add-in advertises the request-policy marker. Raw `vcs_call_vba` calls
to the configured add-in's `API` use that operation entry point too. Add-in API
dispatch accepts the method plus three arguments. Arbitrary user-function calls
retain their existing three-argument limit. A pre-request-policy installed library requires
the original baseline client or an explicitly owned external harness to bootstrap
the supported rebuild. Metadata and inspection remain available through their
existing entry points. No synchronous operation is replayed after a COM exception.
