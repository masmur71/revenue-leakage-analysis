# Core SQL analysis — Phase 2

## Reproduction and safety

Run `./sql/03_analysis/03_verify_phase2.ps1` in PowerShell with the existing Docker container running. The runner copies SQL files into `/tmp/leakage-phase2`, checks five supported date cases on `bintel` and `leakage_lab`, compares all diagnostic row fingerprints, and writes the default business report. Run `./sql/02_kpi/03_verify_phase1.ps1` for the underlying KPI regression suite. No source tables, indexes, schemas or data are modified. PostgreSQL forbids even temporary-view DDL in a read-only transaction, so session-local view definitions are created first; all report reads and validation run within `BEGIN READ ONLY`. Validation inserts only into a temporary check table created before that transaction.

For custom dates, pass `-v period_start=YYYY-MM-DD -v period_end=YYYY-MM-DD -v as_of_date=YYYY-MM-DD` to `psql -X -v ON_ERROR_STOP=1 -f /tmp/leakage-phase2/03_analysis/01_business_report.sql`. The Phase 1 date contract applies. For complete detailed datasets, include `00_analysis_context.sql` and select the relevant view in the same psql session; wrap reads in `BEGIN READ ONLY` / `ROLLBACK`. Use `\copy (SELECT * FROM analysis_leakage ORDER BY dimension,member_id) TO 'your-local-path.csv' CSV HEADER` for a local CSV. Excel export/model design belongs to Phase 4.

## Query map and interpretation

All formulas and rounding follow `docs/kpi_dictionary.md`. The context file documents business intent, assumptions, grain, layered CTE logic and validation for each view. `01_business_report.sql` supplies readable outputs; `02_validate_analysis.sql` raises an exception if any check fails. Detailed outputs retain numeric components for weighted aggregation.

| Business question | Diagnostic view | Grain and interpretation |
|---|---|---|
| Where do price, discounts, margins and returns concentrate? | `analysis_leakage` | Dimension/member, separately for branch, SKU and outlet. Reference bridge, billed/event revenue, margin, return rate, ranked price exposure and cumulative share. Never sum across dimensions. |
| Which branches/regions miss target? | `analysis_branch_month` | Dense branch-month panel, with names/region, nominal target, gap and achievement; region results sum components. Partial boundary months retain actuals but have NULL targets. Previous observed month is restricted to the selected period. |
| Does quantity, blended price or discount explain observed change? | `analysis_yoy` | Branch-month versus matching dates one year earlier. Exact sequential quantity/blended-price/discount/return bridge where both sales quantities are nonzero. Price includes product mix; it is not a pure SKU price effect. Missing prior history is NULL, not zero. |
| Which outlets/SKUs return merchandise, and why? | `analysis_return_reasons`, `analysis_leakage` | Branch/outlet/SKU/reason event totals plus member return value rates. Event-period ratios can exceed 100%; small denominators must be inspected before prioritizing. Reasons are recorded descriptions, not established root causes. |
| Who is past credit terms? What is aging/DSO? | `analysis_ar` | Branch/outlet billed balances; invoice-date SCD2 terms, allocated payments at cutoff, positive balances, all five aging buckets, overdue invoice count, oldest outstanding days past due, cash exceptions and overpayments. Credit-only DSO uses trailing 90-day billings and history coverage. |
| Which salesmen are productive or record mismatches? | `analysis_salesman` | Recorded salesman, including zero-activity dimension members. Visit-event and deduplicated-call counts remain separate. Rates use summed components, calls/day uses observed field days. Home branch is a dimension attribute, not transaction attribution. Mismatch is not a proven loss. |
| Which outlets need engagement? | `analysis_rfm` | One onboarded outlet at snapshot, scores and history-completeness flag. Ordered, exclusive rules: inactive (zero orders), champions (R/F/M >=4), at risk (R<=2), recent developing (R>=4), other active. Labels are analytic conventions, not validated interventions. |
| How do onboarding cohorts retain? | `analysis_cohort` | Complete observed cohort-month cells, including zero activity; fixed onboarding denominator. Adjacent active-outlet change is net population change, not individually tracked churn. Month zero is activation. Age summaries compare different eligible cohorts, not a single survival curve. |

## Limits on answering the brief

There are no target quantities, target prices, target discounts, category targets or budgeted mix. Therefore the monetary target gap is measurable, but a volume/price/discount bridge **against target cannot be identified**. The prior-year bridge explains observed revenue change only and must not be presented as the cause of missing target. Sequential decomposition allocates the quantity/price interaction to the blended-price component. Zero-quantity comparisons do not support that decomposition.

Price shortfall is exposure against the supplied reference, not assured collectible revenue. Invalid attached-promotion discounts are a subset of recorded discounts. Return revenue and returned margin are different quantities. These indicators must not be added into a recovery estimate. This phase provides diagnostics; recovery scenarios and business recommendations belong to Phase 5.

AR includes residual cash invoices and does not assume returns created credit notes. DSO is an estimate, not average delinquency. Same-day visit conversion is association. RFM F/M bands are relative and early snapshots have partial history. Pre-2024 onboarding cohorts are excluded from cohort comparison. All data are synthetic and observation ends on 2025-12-31.

## Validation evidence

`docs/analysis-validation-*-source.txt` and matching `*-copy.txt` record hard-failing grain, additive-control, financial-bridge, target-coverage, partial-month, YoY, aging, productivity, RFM and cohort checks. Five cases cover calendar 2025, calendar 2024, mid-2025, partial January 2024 and full observed history. Eight deterministic full-row fingerprints compare all detailed diagnostic outputs, not just displayed top-ten records. The underlying KPI suite separately checks source domains, SCD2 uniqueness, invoice/payment boundaries, returns crossing periods, sparse promotion flags, attribution, RFM/cohort boundaries and invalid parameters.

Read the default numeric evidence in `docs/analysis-report-2025.txt` alongside the findings below. Ratios in raw outputs are fractions; IDR amounts are unrounded until presentation.

## Default 2025 findings

Enterprise event-period revenue is IDR 15,762,528,589.51 and gross profit is IDR 3,851,071,463.51 (24.43% margin). Price shortfall is IDR 3,023,132.22, explicit discounts are IDR 31,877,245.06, and return revenue is IDR 210,279,535.21. The associated return margin reversal is IDR 50,939,314.21; this is already reflected in gross profit. These are separate diagnostic components, not an additive recoverable-opportunity estimate. Above-reference price upside is zero.

Attached-promotion invalid discount and recorded-salesman mismatch are both zero in this period. This limits the supported story: the supplied flags do not substantiate either source of financial leakage. It does not establish that all commercial controls are effective.

At 2025-12-31, billed outstanding AR is IDR 3,396,761,237.52, including IDR 3,239,849,366.28 overdue. Credit-only trailing-90-day DSO is 87.26 days. These balances retain the dictionary's uncertainty about merchandise-return settlements. National sales exceed the nominal target of IDR 15,742,179,409.15, while the branch/region panel identifies local shortfalls hidden by that aggregate.

All 2025 price shortfall is in Medan (B003), IDR 3,023,132.22. Bandung (B002) has the largest annual target shortfall, IDR 32,431,359.96, at 96.67% achievement, and IDR 34,510,872.43 in returns (3.53% of billed sales). Region Barat achieves 98.78% of target; the other two regions exceed target. Bandung's revenue still increases IDR 140,557,354.62 year on year: this illustrates why the observed-growth bridge cannot explain the target gap causally.

The largest recorded return reason by value is Order Error, IDR 53,343,639.52, followed by Damaged, Expired and Quality Complaint at roughly IDR 52 million each. Outlet O0170 has a 100% event-period return ratio on only IDR 33,315 billed sales; a high ratio alone is not a large financial exposure. Detailed SKU and outlet rates are available separately. Jakarta (B001) has the highest branch overdue AR, IDR 290,782,302.49, and a 132.95-day DSO; the trailing-billing denominator is shown in the underlying outlet detail.

Observed salesman productivity totals 3.21 completed visits per active salesman-day. Completion is 95.93%, effective calls are 6.80% of completed visits, and same-day order conversion is 6.98% of deduplicated completed calls. The full 50-salesman table supports comparisons using those distinct denominators; unlogged working days are unknown.

The snapshot contains 765 champions, 33 at-risk outlets, 24 inactive outlets, 2,165 recent-developing outlets and 13 other active outlets. For observed onboarding cohorts, same-month activation is 494/840 (58.81%); month-one activity is 691/805 (85.84%). These denominators differ because later cohorts have no observed month one yet, and later activity may follow delayed activation. They do not establish an improving individual retention curve.

Phase 2 exit criteria are met: every business question has a reproducible diagnostic or an explicit data limitation, all 15 diagnostic checks pass in each of five date scenarios on both databases, and all eight complete diagnostic fingerprints match per scenario. Phase 1 regression also passes 38 checks per scenario, source/copy reference comparisons and four rejected invalid date cases. Numeric evidence is saved alongside this document. Price-exposure ranking uses dense ranks and retains ties, so a displayed top-rank group can contain more than ten members.
