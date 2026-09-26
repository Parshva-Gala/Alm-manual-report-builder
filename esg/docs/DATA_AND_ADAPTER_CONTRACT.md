# Data and local adapter contract

`web/lib/offline-store.ts` retains the cloud interface's data semantics while replacing network persistence with local IndexedDB. Paths such as `/api/workspace` name adapter operations; they are not remote API calls in the APK. The adapter rejects absolute off-origin operations, unknown paths and unsupported methods.

## Workspace envelope

| Field | Meaning |
|---|---|
| `modelVersion`, `name` | Compatibility marker and bank/workspace display name. |
| `scenario`, `country`, `taxonomySelection` | Selected stress, reporting profile and comparison context. |
| `actionsEnabled`, `carbonScale`, `physicalScale`, `actionOverride` | Explicit forecast controls; action year, amounts, approval and evidence. |
| `inputs` | Shared financial settings, borrowers, facilities, sites, bank assumptions, paths and action schedule. |
| `esg` | Actors, E&S assessments, taxonomy assessments, criterion tests, actions and readiness records. |
| `emissions` | Inventory entities, activity records, factors, operations and bank Scope 3 register. |

The saved revision and timestamp belong to the persistence envelope, not the workspace model. Derived engine results are recalculated; imported or submitted summaries are not authoritative inputs.

Database `esg-banking-offline-v1`, schema version 1, has three object stores:

| Store | Records |
|---|---|
| `meta` | `workspace`: `{id,formatVersion,revision,updatedAt,state}`; `sequence`: `{id,value}`. |
| `runs` | `{id,label,saved_at,state,summary,sequence}`. |
| `events` | `{id,created_at,action,details,revision,sequence}`. |

All stores use `id` as key. A monotonically increasing transaction sequence orders captures/history even if the device clock changes. A missing workspace in a nonempty database is treated as corruption, not as permission to reset. Reads check the saved model and compare each retrieved run's summary with a fresh calculation of its immutable inputs. Unsupported/corrupt storage returns a recovery error and remains intact.

Important join keys are borrower `id`, facility `borrowerId` and `emissionsEntityId`, inventory `id`/`borrowerId`, activity `entityId`/`siteId`/`factorId`, taxonomy `facilityId`/`borrowerId`, and criterion `assessmentId`. Evidence references are entered strings, not attached file contents. Review actors are local demonstration identities.

Reporting/financial dates use Excel serial days in shared financial inputs. Emissions and ESG records also accept their documented ISO date form. Do not convert missing fields to zero or invent dates to pass validation. JSON does not preserve `NaN`, infinity or JavaScript `undefined` as valid model values.

## Operations

| Operation | Input and result |
|---|---|
| `GET /api/workspace` | Initialize a deep-cloned default only if no saved workspace exists. Return `{state,revision,updatedAt,runs,events}`; initial revision is 0. |
| `PUT /api/workspace` | `{state,revision,description}` → `{state,revision,updatedAt}`. Validate, carry stored reviews, compare persisted revision, then atomically save state and one history event at revision+1. |
| `POST /api/runs` | `{state,label,revision}` → `{id,label,saved_at,summary}`. Capture an immutable draft snapshot and recalculated financial summary with an event. Does not save the draft or advance workspace revision. |
| `GET /api/runs?id=…` | Return `{id,label,saved_at,state,summary}`. Loading this result into the UI creates an unsaved draft. |
| `POST /api/review` | `{state,revision,kind,id}` → `{state,revision,updatedAt}`. Review one target, recalculate acceptance gates and atomically save the entire submitted workspace plus its event. |

Run labels are 1–100 characters after trimming. Save descriptions are limited to 500 characters. The workspace listing shows the newest 30 captures and 50 history events; these are presentation limits, not authorization to delete older data.

Adapter errors retain a numeric status: 400 for malformed requests, 404 for unknown records, 405 for unsupported methods, 409 for stale revisions, 413 for oversized requests, 422 for invalid data or a rejected review, and 503 for unavailable/corrupt storage. Request envelopes are limited to 1,800,000 UTF-8 bytes. Storage failure rejects the operation rather than reporting a false successful save. Local error status codes do not imply an HTTP server.

## Revision and review semantics

Revision comparison must happen against persisted data in the write transaction. Workspace changes and their events commit together; captures and their events commit together. A failed or stale operation must not leave a new stamp, event or partial state. Caller-owned drafts and stored snapshots must not share mutable object references.

Ordinary Save and Capture cannot manufacture review stamps. For each incoming criterion, taxonomy, action or readiness ID, copy the previously stored stamp and record type; remove submitted review fields where no stored stamp exists. A relevant edit then fails the old fingerprint. Only explicit Review may create a new stamp, and only for its requested target.

| Review kind | Required derived acceptance |
|---|---|
| `criterion` | `complete === true` |
| `action` | `decisionState === 'Closed'` |
| `taxonomy` | `claimState` is `Approved` or `Activity only` |
| `readiness` | `controlReadiness === 'Verified'` |

The ESG calculator checks the evidence, dates, owner/reviewer separation and decision gates. Review evaluates the assessment's stored scheme rather than granting approval from the top-level taxonomy comparison selection. Explicit Review saves the whole submitted draft, not only the visible record. Offline checks discourage accidental or imported approval changes; they do not establish a tamper-proof system against someone who controls the device or app data.

## Import and output conventions

CSV imports update existing borrower/facility IDs, with an explicit preview and Apply step. Unknown or duplicate IDs, unsupported columns, malformed quoting and blank numeric values are rejected. They do not create new lending records. Workspace JSON restore loads a compatible complete input state as a draft; the user must save it to replace the saved workspace.

Financial output includes `opening`, `reportingECL`, `bank`, `borrowerProjections` and `facilityProjections`. Emissions output includes portfolio/customer/sector/country aggregates, facility attribution, inventory, activity/factor results, operations, bank Scope 3 and `gaps`. ESG output includes borrower decisions, taxonomy, tests, actions, readiness and summary.

For emissions, use `displayFinancedS1S2` / `displayFinancedS3` with their independent coverage fields. Known `financed*` subtotals are not complete totals. Use `bank.completeScope1`, `completeScope2LB`, `completeScope2MB` and their complete combined alternatives for total headlines. A null output means unavailable or not applicable according to its status; display that status rather than a numeric zero.

Native document export and print use `web/lib/platform.ts` and the narrowly scoped `ESGNative.postMessage` bridge. Messages are JSON strings:

| Request | Shape |
|---|---|
| Export | `{id,action:'export',name,mime,text}` |
| Print | `{id,action:'print'}` |
| Reply | `{id,ok,cancelled?,error?}` |

The web caller uses a generated request ID, handles cancellation separately and times out after five minutes. `components/platform-status.tsx` displays the result. Native export permits JSON, CSV or plain text, sanitizes the filename and limits UTF-8 content to 5 MiB. Native document operations are serialized and reject duplicate request IDs. Export success follows completion of the native file write.

The native print flow observes Android `PrintJob` completion, cancellation or failure before replying. If the job remains pending after 290 seconds, it reports that state, releases other document operations and directs the user to the Android print queue; it does not claim success or cancel the job. A second print is blocked while the prior job remains pending. Opening the print dialog alone is not success, and an acknowledgement does not independently inspect the output's content or pagination. The web print helper waits two animation frames before requesting native printing; this is not a device-verified layout guarantee.

HTML file inputs use a single-document SAF picker. Native import accepts a bounded candidate CSV/JSON/text document up to 5 MiB, copies it into private cache and returns the copy through a nonexported FileProvider. Cancellation returns no selection. The web layer still applies its stricter 1,500,000-byte CSV and 1,800,000-byte workspace file limits and semantic validation. Ordinary browser preview has a Blob-download/window-print fallback.
