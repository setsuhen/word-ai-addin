# Real-time Inline AI Audit (Word Online + Word Desktop)

## Scope
Client-side add-in code review focused on initialization, eventing, API robustness, insertion behavior, and diagnostics.

## Key Findings

1. **No Word document event subscriptions are configured**
   - `Office.onReady` wires only UI DOM events and does not register `DocumentSelectionChanged` or content/paragraph change handlers.
   - This means the add-in is pull-based only (user click) and cannot react to rapid typing, cursor moves, or external edits in near real time.

2. **API calls have no resilience controls**
   - `callAI` executes a single `fetch` with no timeout, retry, or 429 backoff logic.
   - Provider parsers are optimistic; malformed provider payloads can silently degrade behavior without structured telemetry.

3. **Text edit operations are global and string-search based**
   - `replace_text`/`delete_text` search the entire body and replace all matches, which can modify unintended ranges when repeated text exists.
   - `insert_text` supports `replace_selection`, but there is no tracked-range strategy to preserve exact insertion anchors over async delays.

4. **Concurrency control is missing for rapid interactions**
   - `runAction` and chat loops can be started repeatedly; there is no cancellation token / run-id guard to drop stale responses.
   - Under rapid typing and repeated requests, delayed AI/tool responses can apply out of order and appear “misaligned.”

5. **Logging is mostly UI-only and not correlated**
   - Errors are surfaced in chat/status UI; only `execTool` has `console.error`.
   - No structured run metadata (provider, latency, tool count, retries, host platform), making Word Online vs Desktop drift hard to diagnose.

## Recommended Fix Plan (high impact first)

1. Add host event bridge with debounce/throttle
   - Register selection/content-related handlers once after `Office.onReady`.
   - Debounce to ~200–350ms and enqueue context snapshot updates.

2. Harden API transport
   - Add request timeout (e.g., 20s), retry policy for 429/5xx with exponential backoff + jitter, and `Retry-After` support.
   - Normalize parse failures into a consistent error shape with provider name.

3. Make edits range-scoped where possible
   - Prefer selection/range-based edits for inline workflows; avoid whole-document `search` for ambiguous snippets.
   - If using search, require stronger matching options and confirm unique hit count before replacement.

4. Add run coordination primitives
   - Maintain `activeRunId`; ignore stale responses if a newer run begins.
   - Optional: provide a “Stop” action using `AbortController` to cancel in-flight network calls.

5. Add structured diagnostics
   - Emit per-run telemetry object: `{ runId, host, platform, provider, promptChars, responseMs, retries, toolCalls, errors[] }`.
   - Keep a rolling in-memory log and optional export button for support tickets.

