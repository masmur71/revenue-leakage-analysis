# Source schema ERD

All 13 source tables and 31 declared foreign keys from [schema DDL](../sql/00_setup/analytics_schema.sql). Parent-to-child cardinality is one to zero/many; nullable foreign keys do not require a parent for every child. Distinct role keys are retained. This describes source constraints, not the Excel relationship graph.

```mermaid
erDiagram
    dim_branch {
        text branch_id PK
    }
    dim_date {
        integer date_sk PK
    }
    dim_outlet {
        text outlet_id PK
    }
    dim_outlet_credit_term {
        bigint credit_term_sk PK
    }
    dim_product {
        text sku_id PK
    }
    dim_salesman {
        text salesman_id PK
    }
    dim_salesman_assignment {
        bigint assignment_sk PK
    }
    fact_invoices {
        text invoice_id PK
    }
    fact_payments {
        text payment_id PK
    }
    fact_returns {
        text return_id PK
    }
    fact_sales {
        text sales_line_id PK
    }
    fact_targets {
        text target_id PK
    }
    fact_visits {
        text visit_id PK
    }
    dim_branch ||--o{ dim_outlet : "branch_id"
    dim_outlet ||--o{ dim_outlet_credit_term : "outlet_id"
    dim_branch ||--o{ dim_salesman_assignment : "branch_id"
    dim_outlet ||--o{ dim_salesman_assignment : "outlet_id"
    dim_salesman ||--o{ dim_salesman_assignment : "salesman_id"
    dim_branch ||--o{ dim_salesman : "home_branch_id"
    dim_branch ||--o{ fact_invoices : "branch_id"
    dim_date ||--o{ fact_invoices : "date_sk"
    dim_outlet ||--o{ fact_invoices : "outlet_id"
    dim_salesman ||--o{ fact_invoices : "salesman_id"
    dim_date ||--o{ fact_payments : "date_sk"
    fact_invoices ||--o{ fact_payments : "invoice_id"
    dim_branch ||--o{ fact_returns : "branch_id"
    dim_date ||--o{ fact_returns : "date_sk"
    dim_outlet ||--o{ fact_returns : "outlet_id"
    fact_sales ||--o{ fact_returns : "sales_line_id"
    dim_product ||--o{ fact_returns : "sku_id"
    dim_salesman ||--o{ fact_sales : "assigned_salesman_id"
    dim_salesman_assignment ||--o{ fact_sales : "assignment_sk"
    dim_branch ||--o{ fact_sales : "branch_id"
    dim_date ||--o{ fact_sales : "date_sk"
    fact_invoices ||--o{ fact_sales : "invoice_id"
    dim_outlet ||--o{ fact_sales : "outlet_id"
    dim_salesman ||--o{ fact_sales : "salesman_id"
    dim_product ||--o{ fact_sales : "sku_id"
    dim_branch ||--o{ fact_targets : "branch_id"
    dim_date ||--o{ fact_targets : "date_sk"
    dim_branch ||--o{ fact_visits : "branch_id"
    dim_date ||--o{ fact_visits : "date_sk"
    dim_outlet ||--o{ fact_visits : "outlet_id"
    dim_salesman ||--o{ fact_visits : "salesman_id"
```

Historical joins: invoice date within outlet credit-term bounds; sale date within assignment bounds, both inclusive. SCD2 interval uniqueness is enforced by validation in addition to the declared keys. Aggregate payments by invoice and returns by sales line before parent joins to prevent fanout. See [data dictionary](data_dictionary.md) for all column types and [KPI contract](kpi_dictionary.md) for reporting dates.
