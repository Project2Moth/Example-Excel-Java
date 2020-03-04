-- liquibase formatted sql
-- changeset nghiemnc:1.1

DROP TABLE IF EXISTS excel_example;

CREATE TABLE excel_example
(
    id                serial,
    time              timestamp,
    display_name      VARCHAR(255),
    account_number    varchar(255) default null,
    store_name        varchar(255) default null,
    turnover_value    DECIMAL      default null,
    fee               decimal      default null,
    amount_to_pay     decimal      default null,
    actually_received decimal      default null,
    PRIMARY KEY (id)
)