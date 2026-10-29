--
-- PostgreSQL database dump
--

\restrict PwrMNrfWX8YZTADaJ90GgMJsZS95EZGOV2cPeJ0nZce7Jt41YbAeEp5RmflIe56

-- Dumped from database version 16.15 (Debian 16.15-1.pgdg13+2)
-- Dumped by pg_dump version 16.15 (Debian 16.15-1.pgdg13+2)

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
SELECT pg_catalog.set_config('search_path', '', false);
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;

--
-- Name: analytics; Type: SCHEMA; Schema: -; Owner: -
--

CREATE SCHEMA analytics;


SET default_tablespace = '';

SET default_table_access_method = heap;

--
-- Name: dim_branch; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.dim_branch (
    branch_id text NOT NULL,
    branch_name text NOT NULL,
    region_id text NOT NULL,
    region_name text NOT NULL,
    city text,
    province text
);


--
-- Name: dim_date; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.dim_date (
    date_sk integer NOT NULL,
    date date NOT NULL,
    year integer NOT NULL,
    quarter integer NOT NULL,
    month integer NOT NULL,
    month_name text NOT NULL,
    week_of_year integer NOT NULL,
    day_of_month integer NOT NULL,
    day_of_week text NOT NULL,
    is_weekend boolean NOT NULL
);


--
-- Name: dim_outlet; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.dim_outlet (
    outlet_id text NOT NULL,
    outlet_name text NOT NULL,
    branch_id text NOT NULL,
    channel text NOT NULL,
    onboard_date date NOT NULL,
    initial_credit_term_days integer NOT NULL
);


--
-- Name: dim_outlet_credit_term; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.dim_outlet_credit_term (
    credit_term_sk bigint NOT NULL,
    outlet_id text NOT NULL,
    valid_from date NOT NULL,
    valid_to date NOT NULL,
    credit_term_days integer NOT NULL,
    version integer NOT NULL,
    is_current boolean NOT NULL
);


--
-- Name: dim_product; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.dim_product (
    sku_id text NOT NULL,
    sku_name text NOT NULL,
    category text NOT NULL,
    standard_cost numeric(18,2) NOT NULL
);


--
-- Name: dim_salesman; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.dim_salesman (
    salesman_id text NOT NULL,
    salesman_name text NOT NULL,
    home_branch_id text NOT NULL,
    hire_date date NOT NULL,
    status text NOT NULL
);


--
-- Name: dim_salesman_assignment; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.dim_salesman_assignment (
    assignment_sk bigint NOT NULL,
    outlet_id text NOT NULL,
    salesman_id text NOT NULL,
    branch_id text NOT NULL,
    valid_from date NOT NULL,
    valid_to date NOT NULL,
    version integer NOT NULL,
    is_current boolean NOT NULL
);


--
-- Name: fact_invoices; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.fact_invoices (
    invoice_id text NOT NULL,
    sales_order_id text NOT NULL,
    invoice_date date NOT NULL,
    due_date date NOT NULL,
    outlet_id text NOT NULL,
    branch_id text NOT NULL,
    salesman_id text NOT NULL,
    channel text NOT NULL,
    invoice_amount numeric(18,2) NOT NULL,
    credit_term_days integer NOT NULL,
    payment_method text NOT NULL,
    date_sk integer NOT NULL
);


--
-- Name: fact_payments; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.fact_payments (
    payment_id text NOT NULL,
    invoice_id text NOT NULL,
    payment_date date NOT NULL,
    payment_amount numeric(18,2) NOT NULL,
    payment_method text NOT NULL,
    allocation_status text NOT NULL,
    date_sk integer NOT NULL
);


--
-- Name: fact_returns; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.fact_returns (
    return_id text NOT NULL,
    sales_line_id text NOT NULL,
    sales_order_id text NOT NULL,
    return_date date NOT NULL,
    outlet_id text NOT NULL,
    branch_id text NOT NULL,
    sku_id text NOT NULL,
    return_quantity integer NOT NULL,
    return_amount numeric(18,2) NOT NULL,
    returned_cogs numeric(18,2) NOT NULL,
    reason text NOT NULL,
    date_sk integer NOT NULL
);


--
-- Name: fact_sales; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.fact_sales (
    sales_line_id text NOT NULL,
    sales_order_id text NOT NULL,
    invoice_id text NOT NULL,
    sales_date date NOT NULL,
    date_sk integer NOT NULL,
    outlet_id text NOT NULL,
    branch_id text NOT NULL,
    salesman_id text NOT NULL,
    assigned_salesman_id text NOT NULL,
    assignment_sk bigint NOT NULL,
    sku_id text NOT NULL,
    channel text NOT NULL,
    quantity integer NOT NULL,
    unit_price_before_discount numeric(18,2) NOT NULL,
    gross_sales numeric(18,2) NOT NULL,
    discount_amount numeric(18,2) NOT NULL,
    net_sales_before_returns numeric(18,2) NOT NULL,
    return_quantity integer NOT NULL,
    return_amount numeric(18,2) NOT NULL,
    returned_cogs numeric(18,2) NOT NULL,
    net_sales numeric(18,2) NOT NULL,
    unit_cost numeric(18,2) NOT NULL,
    cogs_sold numeric(18,2) NOT NULL,
    net_cogs numeric(18,2) NOT NULL,
    gross_profit numeric(18,2) NOT NULL,
    gross_margin_pct numeric(18,6),
    reference_price numeric(18,2) NOT NULL,
    price_deviation_pct numeric(18,6),
    promotion_id text,
    promo_discount_pct numeric(18,4),
    is_promo_valid boolean NOT NULL,
    is_returned boolean NOT NULL,
    is_salesman_match boolean NOT NULL
);


--
-- Name: fact_targets; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.fact_targets (
    target_id text NOT NULL,
    branch_id text NOT NULL,
    target_month text NOT NULL,
    target_net_sales numeric(18,2) NOT NULL,
    currency text NOT NULL,
    date_sk integer NOT NULL
);


--
-- Name: fact_visits; Type: TABLE; Schema: analytics; Owner: -
--

CREATE TABLE analytics.fact_visits (
    visit_id text NOT NULL,
    visit_date date NOT NULL,
    date_sk integer NOT NULL,
    outlet_id text NOT NULL,
    branch_id text NOT NULL,
    salesman_id text NOT NULL,
    visit_status text NOT NULL,
    is_completed boolean NOT NULL,
    is_effective_call boolean NOT NULL
);


--
-- Name: dim_branch dim_branch_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_branch
    ADD CONSTRAINT dim_branch_pkey PRIMARY KEY (branch_id);


--
-- Name: dim_date dim_date_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_date
    ADD CONSTRAINT dim_date_pkey PRIMARY KEY (date_sk);


--
-- Name: dim_outlet_credit_term dim_outlet_credit_term_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_outlet_credit_term
    ADD CONSTRAINT dim_outlet_credit_term_pkey PRIMARY KEY (credit_term_sk);


--
-- Name: dim_outlet dim_outlet_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_outlet
    ADD CONSTRAINT dim_outlet_pkey PRIMARY KEY (outlet_id);


--
-- Name: dim_product dim_product_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_product
    ADD CONSTRAINT dim_product_pkey PRIMARY KEY (sku_id);


--
-- Name: dim_salesman_assignment dim_salesman_assignment_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_salesman_assignment
    ADD CONSTRAINT dim_salesman_assignment_pkey PRIMARY KEY (assignment_sk);


--
-- Name: dim_salesman dim_salesman_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_salesman
    ADD CONSTRAINT dim_salesman_pkey PRIMARY KEY (salesman_id);


--
-- Name: fact_invoices fact_invoices_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_invoices
    ADD CONSTRAINT fact_invoices_pkey PRIMARY KEY (invoice_id);


--
-- Name: fact_invoices fact_invoices_sales_order_id_key; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_invoices
    ADD CONSTRAINT fact_invoices_sales_order_id_key UNIQUE (sales_order_id);


--
-- Name: fact_payments fact_payments_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_payments
    ADD CONSTRAINT fact_payments_pkey PRIMARY KEY (payment_id);


--
-- Name: fact_returns fact_returns_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_returns
    ADD CONSTRAINT fact_returns_pkey PRIMARY KEY (return_id);


--
-- Name: fact_sales fact_sales_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_sales
    ADD CONSTRAINT fact_sales_pkey PRIMARY KEY (sales_line_id);


--
-- Name: fact_targets fact_targets_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_targets
    ADD CONSTRAINT fact_targets_pkey PRIMARY KEY (target_id);


--
-- Name: fact_visits fact_visits_pkey; Type: CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_visits
    ADD CONSTRAINT fact_visits_pkey PRIMARY KEY (visit_id);


--
-- Name: idx_dim_assignment_outlet_valid; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_dim_assignment_outlet_valid ON analytics.dim_salesman_assignment USING btree (outlet_id, valid_from, valid_to);


--
-- Name: idx_dim_credit_outlet_valid; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_dim_credit_outlet_valid ON analytics.dim_outlet_credit_term USING btree (outlet_id, valid_from, valid_to);


--
-- Name: idx_fact_invoices_branch_id; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_invoices_branch_id ON analytics.fact_invoices USING btree (branch_id);


--
-- Name: idx_fact_invoices_due_date; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_invoices_due_date ON analytics.fact_invoices USING btree (due_date);


--
-- Name: idx_fact_invoices_outlet_id; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_invoices_outlet_id ON analytics.fact_invoices USING btree (outlet_id);


--
-- Name: idx_fact_payments_invoice_id; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_payments_invoice_id ON analytics.fact_payments USING btree (invoice_id);


--
-- Name: idx_fact_payments_payment_date; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_payments_payment_date ON analytics.fact_payments USING btree (payment_date);


--
-- Name: idx_fact_returns_return_date; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_returns_return_date ON analytics.fact_returns USING btree (return_date);


--
-- Name: idx_fact_returns_sales_line_id; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_returns_sales_line_id ON analytics.fact_returns USING btree (sales_line_id);


--
-- Name: idx_fact_sales_assignment_sk; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_sales_assignment_sk ON analytics.fact_sales USING btree (assignment_sk);


--
-- Name: idx_fact_sales_branch_id; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_sales_branch_id ON analytics.fact_sales USING btree (branch_id);


--
-- Name: idx_fact_sales_invoice_id; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_sales_invoice_id ON analytics.fact_sales USING btree (invoice_id);


--
-- Name: idx_fact_sales_outlet_id; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_sales_outlet_id ON analytics.fact_sales USING btree (outlet_id);


--
-- Name: idx_fact_sales_sales_date; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_sales_sales_date ON analytics.fact_sales USING btree (sales_date);


--
-- Name: idx_fact_sales_sku_id; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_sales_sku_id ON analytics.fact_sales USING btree (sku_id);


--
-- Name: idx_fact_targets_branch_month; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_targets_branch_month ON analytics.fact_targets USING btree (branch_id, target_month);


--
-- Name: idx_fact_visits_outlet_id; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_visits_outlet_id ON analytics.fact_visits USING btree (outlet_id);


--
-- Name: idx_fact_visits_visit_date; Type: INDEX; Schema: analytics; Owner: -
--

CREATE INDEX idx_fact_visits_visit_date ON analytics.fact_visits USING btree (visit_date);


--
-- Name: dim_outlet dim_outlet_branch_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_outlet
    ADD CONSTRAINT dim_outlet_branch_id_fkey FOREIGN KEY (branch_id) REFERENCES analytics.dim_branch(branch_id);


--
-- Name: dim_outlet_credit_term dim_outlet_credit_term_outlet_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_outlet_credit_term
    ADD CONSTRAINT dim_outlet_credit_term_outlet_id_fkey FOREIGN KEY (outlet_id) REFERENCES analytics.dim_outlet(outlet_id);


--
-- Name: dim_salesman_assignment dim_salesman_assignment_branch_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_salesman_assignment
    ADD CONSTRAINT dim_salesman_assignment_branch_id_fkey FOREIGN KEY (branch_id) REFERENCES analytics.dim_branch(branch_id);


--
-- Name: dim_salesman_assignment dim_salesman_assignment_outlet_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_salesman_assignment
    ADD CONSTRAINT dim_salesman_assignment_outlet_id_fkey FOREIGN KEY (outlet_id) REFERENCES analytics.dim_outlet(outlet_id);


--
-- Name: dim_salesman_assignment dim_salesman_assignment_salesman_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_salesman_assignment
    ADD CONSTRAINT dim_salesman_assignment_salesman_id_fkey FOREIGN KEY (salesman_id) REFERENCES analytics.dim_salesman(salesman_id);


--
-- Name: dim_salesman dim_salesman_home_branch_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.dim_salesman
    ADD CONSTRAINT dim_salesman_home_branch_id_fkey FOREIGN KEY (home_branch_id) REFERENCES analytics.dim_branch(branch_id);


--
-- Name: fact_invoices fact_invoices_branch_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_invoices
    ADD CONSTRAINT fact_invoices_branch_id_fkey FOREIGN KEY (branch_id) REFERENCES analytics.dim_branch(branch_id);


--
-- Name: fact_invoices fact_invoices_date_sk_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_invoices
    ADD CONSTRAINT fact_invoices_date_sk_fkey FOREIGN KEY (date_sk) REFERENCES analytics.dim_date(date_sk);


--
-- Name: fact_invoices fact_invoices_outlet_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_invoices
    ADD CONSTRAINT fact_invoices_outlet_id_fkey FOREIGN KEY (outlet_id) REFERENCES analytics.dim_outlet(outlet_id);


--
-- Name: fact_invoices fact_invoices_salesman_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_invoices
    ADD CONSTRAINT fact_invoices_salesman_id_fkey FOREIGN KEY (salesman_id) REFERENCES analytics.dim_salesman(salesman_id);


--
-- Name: fact_payments fact_payments_date_sk_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_payments
    ADD CONSTRAINT fact_payments_date_sk_fkey FOREIGN KEY (date_sk) REFERENCES analytics.dim_date(date_sk);


--
-- Name: fact_payments fact_payments_invoice_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_payments
    ADD CONSTRAINT fact_payments_invoice_id_fkey FOREIGN KEY (invoice_id) REFERENCES analytics.fact_invoices(invoice_id);


--
-- Name: fact_returns fact_returns_branch_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_returns
    ADD CONSTRAINT fact_returns_branch_id_fkey FOREIGN KEY (branch_id) REFERENCES analytics.dim_branch(branch_id);


--
-- Name: fact_returns fact_returns_date_sk_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_returns
    ADD CONSTRAINT fact_returns_date_sk_fkey FOREIGN KEY (date_sk) REFERENCES analytics.dim_date(date_sk);


--
-- Name: fact_returns fact_returns_outlet_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_returns
    ADD CONSTRAINT fact_returns_outlet_id_fkey FOREIGN KEY (outlet_id) REFERENCES analytics.dim_outlet(outlet_id);


--
-- Name: fact_returns fact_returns_sales_line_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_returns
    ADD CONSTRAINT fact_returns_sales_line_id_fkey FOREIGN KEY (sales_line_id) REFERENCES analytics.fact_sales(sales_line_id);


--
-- Name: fact_returns fact_returns_sku_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_returns
    ADD CONSTRAINT fact_returns_sku_id_fkey FOREIGN KEY (sku_id) REFERENCES analytics.dim_product(sku_id);


--
-- Name: fact_sales fact_sales_assigned_salesman_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_sales
    ADD CONSTRAINT fact_sales_assigned_salesman_id_fkey FOREIGN KEY (assigned_salesman_id) REFERENCES analytics.dim_salesman(salesman_id);


--
-- Name: fact_sales fact_sales_assignment_sk_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_sales
    ADD CONSTRAINT fact_sales_assignment_sk_fkey FOREIGN KEY (assignment_sk) REFERENCES analytics.dim_salesman_assignment(assignment_sk);


--
-- Name: fact_sales fact_sales_branch_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_sales
    ADD CONSTRAINT fact_sales_branch_id_fkey FOREIGN KEY (branch_id) REFERENCES analytics.dim_branch(branch_id);


--
-- Name: fact_sales fact_sales_date_sk_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_sales
    ADD CONSTRAINT fact_sales_date_sk_fkey FOREIGN KEY (date_sk) REFERENCES analytics.dim_date(date_sk);


--
-- Name: fact_sales fact_sales_invoice_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_sales
    ADD CONSTRAINT fact_sales_invoice_id_fkey FOREIGN KEY (invoice_id) REFERENCES analytics.fact_invoices(invoice_id);


--
-- Name: fact_sales fact_sales_outlet_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_sales
    ADD CONSTRAINT fact_sales_outlet_id_fkey FOREIGN KEY (outlet_id) REFERENCES analytics.dim_outlet(outlet_id);


--
-- Name: fact_sales fact_sales_salesman_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_sales
    ADD CONSTRAINT fact_sales_salesman_id_fkey FOREIGN KEY (salesman_id) REFERENCES analytics.dim_salesman(salesman_id);


--
-- Name: fact_sales fact_sales_sku_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_sales
    ADD CONSTRAINT fact_sales_sku_id_fkey FOREIGN KEY (sku_id) REFERENCES analytics.dim_product(sku_id);


--
-- Name: fact_targets fact_targets_branch_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_targets
    ADD CONSTRAINT fact_targets_branch_id_fkey FOREIGN KEY (branch_id) REFERENCES analytics.dim_branch(branch_id);


--
-- Name: fact_targets fact_targets_date_sk_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_targets
    ADD CONSTRAINT fact_targets_date_sk_fkey FOREIGN KEY (date_sk) REFERENCES analytics.dim_date(date_sk);


--
-- Name: fact_visits fact_visits_branch_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_visits
    ADD CONSTRAINT fact_visits_branch_id_fkey FOREIGN KEY (branch_id) REFERENCES analytics.dim_branch(branch_id);


--
-- Name: fact_visits fact_visits_date_sk_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_visits
    ADD CONSTRAINT fact_visits_date_sk_fkey FOREIGN KEY (date_sk) REFERENCES analytics.dim_date(date_sk);


--
-- Name: fact_visits fact_visits_outlet_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_visits
    ADD CONSTRAINT fact_visits_outlet_id_fkey FOREIGN KEY (outlet_id) REFERENCES analytics.dim_outlet(outlet_id);


--
-- Name: fact_visits fact_visits_salesman_id_fkey; Type: FK CONSTRAINT; Schema: analytics; Owner: -
--

ALTER TABLE ONLY analytics.fact_visits
    ADD CONSTRAINT fact_visits_salesman_id_fkey FOREIGN KEY (salesman_id) REFERENCES analytics.dim_salesman(salesman_id);


--
-- PostgreSQL database dump complete
--

\unrestrict PwrMNrfWX8YZTADaJ90GgMJsZS95EZGOV2cPeJ0nZce7Jt41YbAeEp5RmflIe56

