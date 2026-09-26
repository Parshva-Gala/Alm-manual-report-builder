# Cloud-reference application

`cloud-reference/` preserves the original cloud application source for comparison and future maintenance. It is not loaded by the Android APK and is not a synchronization service. The Android local workspace and the cloud workspace are separate data stores.

## Original architecture

The reference uses React with the Next-compatible Vinext/Vite build, Cloudflare Workers runtime and a D1 binding named `DB`. Its application routes are `app/api/workspace/route.ts`, `app/api/runs/route.ts` and `app/api/review/route.ts`. `lib/server-store.ts` supplies authentication access, bounded JSON parsing, validation, stored-review handling and safe response helpers. `lib/workspace.ts` and the three `.mjs` model modules define the shared analytical behavior.

The cloud authentication boundary obtains the platform-authenticated user identity; every workspace/run/history query is scoped to that owner. This boundary depends on the intended hosting platform. Local development's synthetic sign-in is not a production identity mechanism. The Android edition does not reproduce or claim cloud authentication.

The D1 schema has `workspaces`, `runs` and `events`. Workspace revision checks reject conflicting saves. Scenario captures store complete inputs and a calculated summary. Explicit review recalculates evidence gates before recording a reviewed workspace revision. Review actor names remain demonstration identities even though the workspace owner is authenticated.

## Maintain or redeploy the reference

Use the supplied reference README, package lockfile and hosting configuration for its original build flow. The source contains schema-only migrations under `drizzle/`; build output includes the hosting migration package. Node 22.13 or newer is the declared reference runtime requirement. Tests distinguish pure model checks from API checks requiring the running local development server.

Deployment identifiers in `.openai/hosting.json` describe the original project; they are not credentials, a portable database backup or permission to overwrite an existing deployment. A separate deployment needs its own configured project/database and the platform's authenticated deployment workflow. Do not point a demonstration or test run at another user's production store. Authentication-header trust and binding configuration must be preserved by the intended hosting environment.

The local API tests use synthetic development identity and may reset their test workspace. Run them only against the documented development instance. Export data and back up any real deployment before schema or model changes. No cloud database content is bundled with the source ZIP.

## Sharing improvements between editions

Keep financial, emissions and ESG rules aligned through parity tests. A shared rule change should update defaults/reference data, model version, validation and tests deliberately. Offline-specific changes belong in the local adapter and native bridge; cloud owner scoping and server persistence must not be replaced by Android assumptions.

The editable workspace JSON is the common analytical transfer format. Import loads a draft and applies current validation/review rules; it is not automatic synchronization, authenticated evidence transfer or a migration of captures/history. Cross-version imports require an explicit migration rather than silent field substitution.
