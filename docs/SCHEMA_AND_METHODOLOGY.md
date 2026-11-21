# Schema, Formulations & Data Provenance

## 1. Relational Star Schema (`sql/schema_and_views.sql`)

- **`dim_agency`**: `agency_id`, `agency_name`, `agency_tier` (`Holding Company`, `Independent Media Agency`, `Direct Brand`), `account_manager_pod`.
- **`dim_vertical`**: `vertical_id`, `vertical_name`, `benchmark_ctr_pct`, `benchmark_vcr_pct`, `target_cpm_usd`.
- **`dim_ad_product`**: `product_id`, `product_name`, `is_custom_unit`, `spec_resolution`, `max_bitrate_kbps`, `min_lead_time_days`, `requires_vast_42`.
- **`fact_campaign_order`**: `campaign_id`, foreign keys to `dim_agency`, `dim_vertical`, `dim_ad_product`, plus `flight_quarter`, `booked_impressions`, `contracted_cpm_usd`, `booked_budget_usd`, `source_record_count`.
- **`fact_deliverable_qa`**: `package_id`, `campaign_id`, `creative_format`, `is_custom_unit`, `seeded_defect_code`, `pre_qa_passed`, `post_remediation_passed`, `lead_time_hours`, `ops_launch_signoff`.
- **`fact_flight_pacing`**: `campaign_id`, baseline vs. optimized delivered impressions, shortfalls, CTR, eCPM, VCR, and `optimization_action`.
- **`fact_billing_renewal`**: `campaign_id`, `booked_budget_usd`, `billable_revenue_usd`, `invoiced_amount_usd`, `makegood_credit_usd`, `billing_variance_usd`, `in_full_delivery_pct`, `renewal_score`, `renewal_tier`.

## 2. Analytical Formulations

1. **Train-Calibrated Value Bidding (Zero Test Leakage):**
   - Calibrated solely on `ipinyou_train.txt` (`100,000` rows): `pctr_floor = Q_0.25(pCTR_train)`, `base_bid = 2.60 * mean(pay_price_train)`.
   - Holdout bid rule on `ipinyou_test.txt` (`100,000` rows): win if `pCTR >= pctr_floor` and `pay_price <= min(1.75 * mean_price_train, base_bid * pCTR / mean_pCTR_train)`.
2. **Delivery Shortfall Reduction:**
   - `shortfall_reduction_pct = 100 * (sum(baseline_shortfall) - sum(optimized_shortfall)) / sum(baseline_shortfall)` = **`72.0%`**.
3. **Order-to-Billing Accounting Parity:**
   - `billable_revenue = round_half_up((delivered_impressions / 1000) * contracted_cpm, 0.01)`
   - `makegood_credit = booked_budget - billable_revenue`
   - `billing_variance = (invoiced_amount + makegood_credit) - booked_budget == $0.00`.
4. **Renewal Propensity Tiering:**
   - `renewal_score = min(99.5, 0.78 * in_full_delivery_pct + ctr_lift_bonus + custom_unit_bonus)`.
   - Campaigns with `renewal_score >= 92.0` and `in_full_delivery_pct >= 95.0` enter `Renew & Upsell`.
