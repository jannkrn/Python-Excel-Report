# Python Excel Report Automation

Automated reporting pipeline for warehouse/operations data using Python, SQL, Excel export, historical KPI tracking, and Outlook integration.

## Overview

This project automates the daily creation of an operational Excel report for MOK-related warehouse positions.

The pipeline:

1. queries raw position data from a database via ODBC
2. normalizes and enriches the dataset with business-specific calculations
3. builds KPI tables for reporting
4. updates historical KPI records
5. exports the final report to Excel
6. optionally sends the report via Outlook

The goal is to reduce manual reporting effort and provide a reproducible daily reporting workflow.

## Business Context

The script is designed for operational reporting in a warehouse/logistics environment.

It focuses on:
- daily position data
- volume-related KPIs
- picks per position
- container and stop-point metrics
- hourly KPI aggregation
- comparison with historical performance

## Features

- ODBC database connection
- SQL-based extraction of operational data
- business-specific KPI calculations
- lookup-based enrichment (e.g. route / shelf-meter mapping)
- daily MOK summary generation
- hourly KPI matrix
- historical KPI storage and deduplication
- Excel export with multiple report sheets
- logging to file and console
- Outlook mail integration

## Tech Stack

- Python
- pandas
- pyodbc
- openpyxl
- win32com / Outlook COM
- Excel
- SQL / ODBC

## Project Structure

```text
.
├── main.py
├── README.md
└── output/
