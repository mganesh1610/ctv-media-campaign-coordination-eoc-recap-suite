PRAGMA foreign_keys = ON;

DROP VIEW IF EXISTS vw_vertical_kpi_summary;
DROP VIEW IF EXISTS vw_agency_renewal_pipeline;
DROP VIEW IF EXISTS vw_billing_reconciliation_audit;
DROP VIEW IF EXISTS vw_custom_unit_sla_health;

DROP TABLE IF EXISTS fact_billing_renewal;
DROP TABLE IF EXISTS fact_flight_pacing;
DROP TABLE IF EXISTS fact_deliverable_qa;
DROP TABLE IF EXISTS fact_campaign_order;
DROP TABLE IF EXISTS dim_ad_product;
DROP TABLE IF EXISTS dim_vertical;
DROP TABLE IF EXISTS dim_agency;

CREATE TABLE dim_agency (
    agency_id INTEGER PRIMARY KEY,
    agency_name TEXT NOT NULL UNIQUE,
    agency_tier TEXT NOT NULL CHECK (agency_tier IN ('Holding Company', 'Independent Media Agency', 'Direct Brand')),
    account_manager_pod TEXT NOT NULL
);

CREATE TABLE dim_vertical (
    vertical_id INTEGER PRIMARY KEY,
    vertical_name TEXT NOT NULL UNIQUE,
    benchmark_ctr_pct REAL NOT NULL,
    benchmark_vcr_pct REAL NOT NULL,
    target_cpm_usd REAL NOT NULL
);

CREATE TABLE dim_ad_product (
    product_id INTEGER PRIMARY KEY,
    product_name TEXT NOT NULL UNIQUE,
    is_custom_unit INTEGER NOT NULL CHECK (is_custom_unit IN (0, 1)),
    spec_resolution TEXT NOT NULL,
    max_bitrate_kbps INTEGER NOT NULL,
    min_lead_time_days INTEGER NOT NULL,
    requires_vast_42 INTEGER NOT NULL CHECK (requires_vast_42 IN (0, 1))
);

CREATE TABLE fact_campaign_order (
    campaign_id TEXT PRIMARY KEY,
    agency_id INTEGER NOT NULL REFERENCES dim_agency(agency_id),
    vertical_id INTEGER NOT NULL REFERENCES dim_vertical(vertical_id),
    product_id INTEGER NOT NULL REFERENCES dim_ad_product(product_id),
    flight_quarter TEXT NOT NULL,
    booked_impressions INTEGER NOT NULL CHECK (booked_impressions > 0),
    contracted_cpm_usd REAL NOT NULL CHECK (contracted_cpm_usd > 0),
    booked_budget_usd REAL NOT NULL CHECK (booked_budget_usd > 0),
    is_custom_unit INTEGER NOT NULL CHECK (is_custom_unit IN (0, 1)),
    source_record_count INTEGER NOT NULL CHECK (source_record_count > 0)
);

CREATE TABLE fact_deliverable_qa (
    package_id TEXT PRIMARY KEY,
    campaign_id TEXT NOT NULL REFERENCES fact_campaign_order(campaign_id),
    creative_format TEXT NOT NULL,
    is_custom_unit INTEGER NOT NULL CHECK (is_custom_unit IN (0, 1)),
    seeded_defect_code TEXT NOT NULL,
    pre_qa_passed INTEGER NOT NULL CHECK (pre_qa_passed IN (0, 1)),
    post_remediation_passed INTEGER NOT NULL CHECK (post_remediation_passed IN (0, 1)),
    lead_time_hours REAL NOT NULL,
    ops_launch_signoff TEXT NOT NULL CHECK (ops_launch_signoff IN ('APPROVED_ON_TIME', 'REMEDIATED_ON_TIME', 'DELAYED'))
);

CREATE TABLE fact_flight_pacing (
    campaign_id TEXT PRIMARY KEY REFERENCES fact_campaign_order(campaign_id),
    baseline_delivered_impressions INTEGER NOT NULL,
    optimized_delivered_impressions INTEGER NOT NULL,
    baseline_shortfall_impressions INTEGER NOT NULL,
    optimized_shortfall_impressions INTEGER NOT NULL,
    baseline_ctr_pct REAL NOT NULL,
    optimized_ctr_pct REAL NOT NULL,
    baseline_ecpm_usd REAL NOT NULL,
    optimized_ecpm_usd REAL NOT NULL,
    vcr_pct REAL NOT NULL,
    optimization_action TEXT NOT NULL
);

CREATE TABLE fact_billing_renewal (
    campaign_id TEXT PRIMARY KEY REFERENCES fact_campaign_order(campaign_id),
    booked_budget_usd REAL NOT NULL,
    billable_revenue_usd REAL NOT NULL,
    invoiced_amount_usd REAL NOT NULL,
    makegood_credit_usd REAL NOT NULL,
    billing_variance_usd REAL NOT NULL,
    in_full_delivery_pct REAL NOT NULL,
    renewal_score REAL NOT NULL,
    renewal_tier TEXT NOT NULL CHECK (renewal_tier IN ('Renew & Upsell', 'Renew Standard', 'At-Risk Intervention'))
);

CREATE VIEW vw_vertical_kpi_summary AS
SELECT
    v.vertical_name,
    COUNT(o.campaign_id) AS campaign_count,
    SUM(o.booked_impressions) AS total_booked_impressions,
    SUM(p.optimized_delivered_impressions) AS total_delivered_impressions,
    ROUND(100.0 * SUM(p.optimized_delivered_impressions) / SUM(o.booked_impressions), 2) AS delivery_attainment_pct,
    ROUND(AVG(p.baseline_ctr_pct), 4) AS avg_baseline_ctr_pct,
    ROUND(AVG(p.optimized_ctr_pct), 4) AS avg_optimized_ctr_pct,
    ROUND(AVG(p.vcr_pct), 2) AS avg_vcr_pct,
    ROUND(SUM(b.billable_revenue_usd), 2) AS total_billable_revenue_usd
FROM fact_campaign_order o
JOIN dim_vertical v ON v.vertical_id = o.vertical_id
JOIN fact_flight_pacing p ON p.campaign_id = o.campaign_id
JOIN fact_billing_renewal b ON b.campaign_id = o.campaign_id
GROUP BY v.vertical_name;

CREATE VIEW vw_agency_renewal_pipeline AS
SELECT
    a.agency_name,
    a.agency_tier,
    a.account_manager_pod,
    COUNT(o.campaign_id) AS campaigns_managed,
    ROUND(SUM(o.booked_budget_usd), 2) AS booked_io_usd,
    ROUND(SUM(b.billable_revenue_usd), 2) AS billable_revenue_usd,
    ROUND(AVG(b.in_full_delivery_pct), 2) AS avg_in_full_pct,
    ROUND(AVG(b.renewal_score), 1) AS avg_renewal_score,
    SUM(CASE WHEN b.renewal_tier = 'Renew & Upsell' THEN 1 ELSE 0 END) AS upsell_ready_campaigns
FROM fact_campaign_order o
JOIN dim_agency a ON a.agency_id = o.agency_id
JOIN fact_billing_renewal b ON b.campaign_id = o.campaign_id
GROUP BY a.agency_name, a.agency_tier, a.account_manager_pod;

CREATE VIEW vw_billing_reconciliation_audit AS
SELECT
    COUNT(*) AS total_orders,
    ROUND(SUM(booked_budget_usd), 2) AS total_booked_usd,
    ROUND(SUM(billable_revenue_usd), 2) AS total_billable_usd,
    ROUND(SUM(invoiced_amount_usd), 2) AS total_invoiced_usd,
    ROUND(SUM(makegood_credit_usd), 2) AS total_makegood_credit_usd,
    ROUND(SUM(ABS(billing_variance_usd)), 2) AS total_abs_billing_variance_usd
FROM fact_billing_renewal;

CREATE VIEW vw_custom_unit_sla_health AS
SELECT
    CASE WHEN is_custom_unit = 1 THEN 'Custom Unit' ELSE 'Standard Unit' END AS unit_class,
    COUNT(*) AS total_packages,
    SUM(CASE WHEN seeded_defect_code != 'NONE' THEN 1 ELSE 0 END) AS seeded_defects,
    SUM(CASE WHEN seeded_defect_code != 'NONE' AND pre_qa_passed = 0 THEN 1 ELSE 0 END) AS defects_intercepted,
    SUM(post_remediation_passed) AS launch_ready_packages,
    ROUND(AVG(lead_time_hours), 2) AS avg_lead_time_hours
FROM fact_deliverable_qa
GROUP BY is_custom_unit;
