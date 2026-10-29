# Revenue leakage: priorities and recommendations

Decision basis: calendar 2025; receivables snapshot 31 December 2025. Amounts are nominal IDR. Event-period net sales were **15,762,528,589.51**, with gross profit **3,851,071,463.51** (24.43%). Preserve the [KPI contract](kpi_dictionary.md): returns follow return dates, discounts follow sales dates, and AR uses allocated payments at cutoff and invoice-date credit terms.

Observed leakage means recorded revenue reductions or reference-price exposure, not proven misconduct or collectible loss. Returns were **210,279,535.21**, discounts **31,877,245.06**, and below-reference price exposure **3,023,132.22**. These are diagnostics, not an additive recovery total. Recoverable opportunity below is explicitly hypothetical and requires a controlled pilot.

## 1. Prevent avoidable returns first — Sales Operations and Quality

**Finding and magnitude.** Returns are the largest observed revenue deduction; their gross-profit reversal is **50,939,314.21**, already reflected in reported profit. Order Error is the largest recorded reason (**53,343,639.52**). Bandung alone records **34,510,872.43**, or **3.53%** of billed sales. Reason labels suggest investigation routes, not established causes. Evidence: [E1–E2](../sql/03_analysis/04_recommendation_evidence.sql), using `kpi_commercial`, `analysis_return_reasons` and `analysis_leakage`.

**Action.** Within 30 days, review high-value return records, beginning with Bandung and national Order Error cases. Validate order confirmation, picking accuracy, damage handling and expiry rotation with operators. Rank outlets/SKUs by returned value alongside their billed-sales denominator; avoid prioritizing tiny outlets solely by return percentage. Run a 90-day confirmation/handling pilot and track event-period return value and margin reversal against a comparable untreated group.

**Estimated impact.** On an equivalent annual business base, assume 5% of national return value is preventable, with the same reason/product margin mix, unchanged original sales demand and no shifted returns into later periods. Revenue preserved = 210,279,535.21 × 5% = **10,513,976.76**; associated gross profit preserved = 50,939,314.21 × 5% = **2,546,965.71**, before intervention costs. These are two views of one benefit. Bandung and Order Error are investigation cuts inside this national envelope; neither receives an additional impact estimate. A 90-day pilot does not itself demonstrate an annual saving.

## 2. Audit Medan reference-price exceptions — Commercial Finance

**Finding and magnitude.** All **3,023,132.22** of below-reference price exposure is in Medan (B003); above-reference upside is zero. Explicit discounts are a separate bridge component. Invalid attached-promotion discounts and recorded-salesman mismatch sales are both **zero**: their flags support no financial recovery claim. Evidence: [E1 and E3](../sql/03_analysis/04_recommendation_evidence.sql); definitions in [core analysis](core_analysis.md).

**Action.** Within 30 days, reconcile Medan reference prices to dated price agreements and approved exceptions. Pilot approval for unexplained below-reference prices; retain valid negotiated terms. Monitor quantity, billed sales, discount substitution and subsequent returns before expanding the control. Do not classify non-promo rows with `is_promo_valid=false` as invalid promotions.

**Estimated impact.** Assume half the reference shortfall is commercially correctable, unchanged quantity and explicit discounts, and full realization of the price increase without incremental cost or additional returns. The annual revenue ceiling is 3,023,132.22 × 50% = **1,511,566.11**, with the same conditional gross-profit ceiling before costs. Reference prices are not demonstrated willingness to pay; actual recovery could be zero. Keep this scenario separate from return prevention because price changes and later returns can interact. No combined revenue or profit total is claimed.

## 3. Resolve overdue balances — Credit Control

**Finding and magnitude.** Outstanding billed AR is **3,396,761,237.52**, including **3,239,849,366.28** overdue. Jakarta has the highest branch overdue balance (**290,782,302.49**). Evidence: [E4](../sql/03_analysis/04_recommendation_evidence.sql), based on `analysis_ar` and the invoice snapshot; this is a cash exposure, not observed revenue leakage or bad-debt expense.

**Action.** Start a 30-day invoice-level reconciliation and collections queue, ranked by overdue amount and aging. Confirm allocated receipts, disputed invoices and return-related settlements before pursuing balances; reconcile residual cash invoices separately. Assign owners and agreed payment dates, then monitor aging migration and verified cash receipts.

**Estimated impact.** Assuming 5% of snapshot overdue AR is valid, collectible and paid within 90 days specifically because of the intervention, accelerated cash is **161,992,468.31**. This is a one-time timing scenario, not annual revenue, profit or proven avoided write-offs. No financing saving is estimated without a borrowing rate and acceleration duration. Exclude receipts that would have occurred anyway when measuring incremental cash.

## Reproduction, measurement and limits

Run the focused evidence below from the repository root in PowerShell; it creates only session-local views, runs SELECTs in `BEGIN READ ONLY`, and rolls back. It does not refresh Excel or rerun exhaustive Verify.

```powershell
docker cp ./sql/. bintel-pg:/tmp/leakage-phase5
docker exec bintel-pg psql -X -q -U bintel -d bintel -v ON_ERROR_STOP=1 -f /tmp/leakage-phase5/03_analysis/04_recommendation_evidence.sql
```

E5 reproduces return-scenario arithmetic. Use invoice/return/sales-line identifiers to attribute realized benefits once; do not add cash collections to commercial improvements, discount subsets to discounts, or gross-profit benefits to revenue benefits. The reference bridge is reference sales + upside − shortfall − discounts = billed sales; event returns then produce net sales. No blanket discount reduction is recommended without contractual eligibility and demand evidence.

The [verification record](final_qa.md) supplies reconciled workbook evidence; the [optimization report](optimization_report.md) supports tested query alternatives, not financial savings. Keep reference SQL authoritative. All data are synthetic, observation ends in 2025, promotions are sparse, and the model establishes neither causality, collectibility nor credit-note settlement. No production ROI or forecast is validated. Monetary branch targets cannot identify volume/price/discount causes without corresponding target components. Validate proposed interventions and costs on operational data before investment; do not extrapolate visit associations or RFM segments into unsupported revenue claims.
