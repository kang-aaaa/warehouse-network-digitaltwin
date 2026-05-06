# Warehouse Network Digital Twin

This repository prepares the Phase 1 warehouse-to-customer network dataset for a road-network-based Digital Twin.

## Current Data Source

- `customer_cleaned.xlsx` is the source workbook.
- `Sheet1` is treated as the operational origin-destination link table.
- Each row in `Sheet1` represents one current warehouse-to-customer assignment row.

## Prepared Outputs

Run the preparation script to create normalized datasets:

```bash
python scripts/prepare_customer_data.py
```

The script generates:

- `data/processed/warehouses.csv`: unique warehouse master data.
- `data/processed/customers.csv`: unique customer master data.
- `data/processed/od_links.csv`: current operational OD rows for matrix routing.
- `data/processed/preparation_summary.json`: machine-readable quality summary.
- `reports/customer_data_preparation_report.md`: human-readable quality report.

## Phase 1 Routing Plan

The generated `od_links.csv` is the input for the next road-network calculation step.
For 10,000+ links, calculate distance/time as a batch matrix first and calculate detailed route geometry only when needed for map display.

Recommended next implementation order:

1. Run this preparation script after every source workbook update.
2. Validate rows flagged in `reports/customer_data_preparation_report.md`.
3. Feed `origin_lon,origin_lat` and `destination_lon,destination_lat` from `od_links.csv` into a self-hosted OSRM, GraphHopper, or Valhalla matrix job.
4. Store road distance/time results with routing engine, profile, graph version, and calculation timestamp.
5. Cache detailed route geometry only for clicked links, outliers, high-volume links, and demo examples.
