# ESG Banking — Risk & Capital

A private analytical application combining the reviewed ESG banking framework, synthetic portfolio and financial reference model. The original Excel product remains unchanged.

## Workspaces

- Portfolio overview and searchable customer drill-down.
- Four live stress cases, severity controls, management actions and immutable scenario captures.
- Five-year earnings, balance sheet, capital/RWA, internal headroom and liquidity planning schedules.
- Separate probability-weighted reporting ECL.
- Financed emissions, coverage, attribution, activity/factor inputs, bank operations and Scope 3 screening.
- E&S assessments, action evidence, taxonomy criteria and country-specific reporting readiness.
- CSV validation and preview, complete workspace backup/restore, history and printable management reports.

Egypt is the default. Jordan, UAE and Global profiles are available. A country change does not insert unverified capital minima or silently change tax/credit assumptions. Changed taxonomy schemes are comparison views until explicitly reviewed.

## Model boundaries

All initial bank/customer values and review identities are synthetic. Bank calibration and formal model validation are required before decision use. Loans run off without new originations. LCR/NSFR are planning proxies. ESG management scores do not mechanically rewrite PD or RWA. Unsupported methods remain pending; missing emissions remain uncovered. Evidence references are entered records, not authenticated files or external assurance.

## Persistence and controls

Authenticated owner identity scopes every database query. Workspace saves use revision checks to reject conflicting edits. Ordinary saves cannot create review stamps. Explicit review recalculates the evidence gates on the server and records a workspace event. Reviewer names are demonstration identities, not a production maker/checker entitlement system. Scenario captures retain their own complete inputs and calculated summary.

## Run and verify

Requires Node 22.13+ and npm. Install with `npm ci`, then use `npm run dev`. Local development provides the starter's demonstration sign-in. Production uses the private Site authentication boundary and the DB binding defined in `.openai/hosting.json`.

- `node tests/verify-models.mjs` — 38 integration and mutation checks.
- `node tests/verify-api.mjs` — 14 checks against the running local development server; uses only its synthetic sign-in and restores the initial workspace.
- `npx tsc --noEmit` — type checks.
- `npm run build` — deployable Worker and client assets in `dist`.

The pure financial engine was separately reconciled over 11,952 numeric/control comparisons. Emissions passed 5,404 assertions across 34 mutation scenarios; ESG passed 63 parity/mutation checks. These verify implementation against the supplied demonstration methods, not regulatory approval or real-world calibration.

Schema-only migrations are in `drizzle`; the Sites build includes them under `dist/.openai/drizzle`. Source and deployment identifiers remain in `.openai/hosting.json`.
