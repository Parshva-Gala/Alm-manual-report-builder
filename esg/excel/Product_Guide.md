# ESG Banking Product

This edition applies the ESG Banking Framework to an integrated Excel banking model. The framework remains the conceptual and research reference. This guide describes how to demonstrate the model and the boundaries of its calculations.

## Start a demonstration

1. Open the workbook at **Overview**. The six headline measures connect the loan portfolio, accounting ECL, capital, financed emissions, data coverage and activity classification.
2. Open **Settings**. Select Egypt, Jordan, UAE or Global, then choose a stress case and customer. The selected case drives the same five-year financial forecast throughout the workbook.
3. Open **Customer** to show the customer's exposure, E&S disposition, earnings and cash transmission, credit parameters and financed emissions. The E&S disposition is independent of the management score and credit model.
4. Open **Capital**. Compare the capital path before and after approved actions, inspect capital headroom, and review LCR and NSFR. Follow the action link to change the planned equity issue, funding and implementation cost. Proposed or unevidenced actions do not become effective.
5. Review **Disclosures**, including the coverage measures alongside emissions. The country profile links to its requirement register and the readiness records. These indicators do not replace a regulator's prescribed return.

Blue cells are editable. Calculations recalculate when inputs change; missing inputs and unsupported cases display a review state. The workbook uses standard formulas, native charts, tables, filters, validation lists, frozen headings and internal navigation. No macros or external workbook links are required.

## Data and scale

The delivered example contains 24 fictional borrowers, 36 facilities and 48 physical sites across Egypt, Jordan and UAE. Every financial amount, hazard score, factor, evidence reference and calibration in the example is synthetic unless a field explicitly identifies a methodological rule or source.

The prepared ranges support 100 borrowers, 100 facilities, 100 sites, 100 emissions entities and 100 activity allocations; 300 activity records and 300 management actions; and 2,000 taxonomy evidence tests. Enter new records inside the prepared tables, preserve unique IDs and leave calculated columns intact. Capacity beyond these bounds requires extending the dependent ranges and validation tests.

Borrower financials use USD millions. Facilities may use USD, EGP, JOD or AED, with a dated conversion table. The seeded exchange rates are illustrative. Change the rate table before using actual local-currency balances. Emissions use tonnes CO2e, and underlying activity quantities retain their own units. Stable borrower, facility, site, emissions-entity and assessment IDs connect the records. Financed project or asset emissions use their own entity boundary.

## Financial methods

**Accounting ECL.** The reporting-date module applies an IFRS 9-based method with independently editable central, downside and upside probabilities and credit parameters. It distinguishes drawn allowances from undrawn provisions; 12-month from lifetime default horizons; and Stage 3 discounted recovery cash flows. The calculation includes survival, effective-interest discounting and recovery timing. Remaining maturity may extend to 30 years even though the displayed financial plan covers five years. The selected climate stress case does not change the accounting scenario probabilities.

The compact ECL calculation assumes a constant conditional annual PD and LGD within each accounting case and a geometric repayment profile, with final maturity settlement. Entered LGD is the nominal unrecovered fraction; the engine discounts the recoverable portion for the entered recovery lag. This is a transparent model convention requiring calibration. It is not an account-level contractual cash-flow engine. SICR includes relative and absolute PD changes, qualitative flags and past-due backstops. Overrides require evidence and a reviewer; known credit impairment and the default backstop take precedence.

**Climate stress.** Separate annual paths specify carbon cost, energy cost, sector demand, interruption, physical damage, adaptation delivery, transition spending, funding conditions and direct bank losses. Site records aggregate physical effects using revenue and collateral shares. Borrower cash deterioration affects explicitly entered credit sensitivities. E&S scores, green classification and financed-emissions totals do not mechanically become PD or regulatory capital charges.

**Facility forecast.** The loan book runs off through repayments and draws against existing commitments. Default years are explicit conditional stress assumptions, shared across cases unless edited; a scheduled default can therefore also appear in Baseline. Default occurs at year-end before principal collection, and unused commitments are assumed cancelled. The model tracks performing and defaulted balances, recoveries, write-offs, drawn allowances, undrawn provisions and interest. Defaulted gross balances stop accruing by planning convention; the discount on expected recoveries unwinds through interest. This is not the full IFRS gross carrying amount and interest presentation. Future provisions use the active conditional parameters. No new loan originations or automatic cures are assumed.

**Bank and ICAAP.** The model reconciles a complete illustrative loan register to the bank's earnings, cash, balance sheet, eligible capital and RWA. Impairment expense reflects allowance movements and write-offs with the relevant interest adjustment. Retained profits, dividends and approved capital issuance affect capital; changes in RWA affect the denominator. Regulatory CCF, risk weights, default weights and prudential treatment require bank-approved parameters. The stress risk-weight multiplier is an entered planning assumption, not an ESG formula or a prescribed national calibration. Market and operational RWA are separately entered.

Opening book equity is the balance of the entered assets and liabilities after calculated allowances; it is a synthetic starting balance, not an imported historical capital account. AT1 is assumed included in book equity, while Tier 2 is included in term funding. CET1 deducts the assumed AT1 and prudential deductions from equity. Capital eligibility and deductions are held fixed unless the model inputs are changed.

The facility's **approved prudential provision** is a separate USD-million input used for RWA netting, scaled with the forecast exposure. It does not override accounting ECL. The **booked undrawn provision local m (reference)** field is reference information only; calculated ECL determines the opening commitment provision. This distinction needs to be reconciled to a client's approved accounting and prudential policies.

Management actions include capital issuance, term funding and fees, with approval, evidence and the Settings action switch controlling effectiveness. Funding and equity arrive at year-end and affect interest from the following year. The before-action comparison removes the action cash, funding cost, tax and dividend effects consistently, including retained action earnings. Internal CET1, Tier 1, total-capital, LCR and NSFR targets remain editable bank assumptions.

**Liquidity.** The workbook calculates LCR and NSFR planning proxies. LCR approximates the next 30 days at each annual forecast date; NSFR applies entered structural funding factors. Actual annual deposit withdrawals affect cash once. Remaining deposits then drive prospective runoff. Inflows use an annual principal-and-interest collection approximation divided by twelve, a separate recognition factor and an inflow cap. This does not identify exact contractual payment dates. HQLA recognition is also an entered factor; the model does not implement detailed asset eligibility and Level 2 concentration caps. Actual maturity schedules and regulatory classifications are needed for supervisory use. A negative cash balance produces a funding-required state and a visible shortfall. No automatic funding closes the gap.

## ESG, taxonomy and emissions

**E&S decisions.** Sector-weighted management assessments sit alongside inherent risk, legal prohibition, critical issues and current evidence. An attractive score cannot clear a critical issue. Closing an action requires dated evidence and independent verification; a status change alone does not remove the restriction. Readiness measures bank capability and evidence, not capital relief.

**Taxonomy.** The activity register separates facility allocation, technical criteria, DNSH, minimum safeguards and independent review. Individual evidence tests support each claim; unsupported routes remain criteria-required. Allocation controls prevent amounts exceeding the facility's drawn balance. The reviewed Jordan pathways include solar PV, cement, building renovation and road freight, with the applicable conditions and dated thresholds. Internal policy eligibility and a foreign reference benchmark are labelled separately from a domestic taxonomy claim. A taxonomy classification does not constitute credit approval.

**Financed emissions.** The reviewed PCAF 2025 core branches cover corporate lending, project finance, commercial real estate, mortgages and motor vehicles. Attribution uses drawn outstanding and the appropriate entity or origination denominator. Asset-class or entity mismatches, duplicate IDs, unsupported methods and over-allocation are surfaced. Scope-specific coverage and exposure-weighted data-quality scores accompany the amounts. Financed borrower Scope 3 remains separate. Reported emissions and a complete activity-derived inventory are alternative sources for the same scope and are not added together.

**Operations and Scope 3.** Bank operations have their own control boundary. Location-based and market-based Scope 2 are alternatives, not additive totals. Factor geography, year, units and evidence are checked. The 15-category bank Scope 3 register makes unassessed and justified exclusions visible. The seeded inventory intentionally contains coverage gaps so that the demonstration shows how incomplete evidence is handled.

## Jurisdictions and research

Egypt is the initial profile. The country register distinguishes binding requirements, guidance, methodology and internal policy, with version, scope, timing, source link and verification status. It includes CBE sustainable-finance policies and reporting, the supplied ESRMS translation and its verification boundary, Jordan CBJ/ASE distinctions, and the UAE's binding CBUAE C8/2025 climate-risk regulation. Global standards remain methodological references whose local legal application is assessed separately.

Egypt uses EAS and a CBE banking accounting framework. The workbook's IFRS 9-based calculations therefore require reconciliation to the bank's applicable CBE accounting policy. The supplied English ESRMS translation contains a January 2028 implementation deadline; the original official Arabic attachment was not independently authenticated in this review. Internal readiness target dates are not represented as legal deadlines.

The research basis is the existing **ESG_Banking_Framework.md**, including its source hierarchy, profiles of all 45 materials and cross-references, together with the detailed retained source reviews. The original historical ICAAP Jordan workbook informs the separation of risks, capital planning and stress testing. Its historical capital ratios and questionnaire coefficients are not treated as current mandates or calibrated model parameters.

## Readiness for client use

This is a substantial functional demonstration and model foundation. A client implementation still requires real bank data, approved regulatory classifications and parameters, calibrated climate and credit assumptions, evidence workflows, independent model validation, security/access design and agreement on the exact reporting templates. Synthetic values, uncalibrated sensitivities and simplified contractual profiles must not be used as a basis for a live lending, provisioning or supervisory decision.

The accompanying validation note records the tests performed on this delivered version, including any native Excel checks and remaining limits.
