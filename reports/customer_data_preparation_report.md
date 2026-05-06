# Customer Data Preparation Report

Generated at: 2026-05-06 06:57:56 UTC

## Input Workbook

- File: `customer_cleaned.xlsx`
- Main sheet used: `Sheet1`

| Sheet | Dimension |
| --- | --- |
| Sheet1 | A1:J17332 |
| Sheet2 | A1:D14 |

## Output Summary

| Metric | Value |
| --- | ---: |
| Raw data rows | 17331 |
| Unique warehouses | 13 |
| Unique customers | 16487 |
| OD link rows | 17331 |
| Unique warehouse-customer pairs | 16494 |
| Duplicate OD pair groups | 582 |
| Duplicate OD extra rows | 837 |
| Rows with coordinate/name validation warnings | 51 |
| Warehouse names with multiple coordinate/address sets | 0 |
| Customer names with multiple coordinate/address sets | 970 |

## Validation Rules

- Warehouse/customer names must be present.
- Warehouse and customer coordinates must parse as finite numbers.
- Coordinates are expected to fall within a broad Korea bounding box: latitude 33.0-39.5, longitude 124.0-132.0.
- Duplicate OD rows are retained in `od_links.csv` because they may represent separate operational rows; the unique pair count is reported for routing de-duplication.

## Invalid Row Samples

| Excel row | Source row no | Warehouse | Customer | Errors |
| ---: | ---: | --- | --- | --- |
| 8210 | 12923 | 광주RDC | 유한회사 서진 | invalid_or_out_of_korea_customer_coordinate |
| 8558 | 14255 | 양지RDC | 금호타이어 대산카독크 | invalid_or_out_of_korea_customer_coordinate |
| 12056 | 20769 | 시흥RDC | (1)민수용(서경환) | invalid_or_out_of_korea_customer_coordinate |
| 12057 | 20770 | 대전RDC | (2)민수용(허석철) | invalid_or_out_of_korea_customer_coordinate |
| 12058 | 20771 | 대전RDC | (3)민수용(김정현) | invalid_or_out_of_korea_customer_coordinate |
| 12060 | 20777 | 대전RDC | (4)민수용(김명진) | invalid_or_out_of_korea_customer_coordinate |
| 12061 | 20778 | 양지RDC | (5)민수용(조창영) | invalid_or_out_of_korea_customer_coordinate |
| 12062 | 20779 | 대구RDC | (6)민수용(김수민) | invalid_or_out_of_korea_customer_coordinate |
| 12063 | 20780 | 양지RDC | (7)민수용(장수한) | invalid_or_out_of_korea_customer_coordinate |
| 12064 | 20781 | 양산RDC | (8)민수용(조용래) | invalid_or_out_of_korea_customer_coordinate |
| 12065 | 20782 | 양지RDC | (9)민수용(김정우) | invalid_or_out_of_korea_customer_coordinate |
| 12066 | 20783 | 양지RDC | (10)민수용(유범석) | invalid_or_out_of_korea_customer_coordinate |
| 12067 | 20784 | 양지RDC | (11)민수용(권도학) | invalid_or_out_of_korea_customer_coordinate |
| 12068 | 20785 | 양지RDC | (12)민수용(심재천) | invalid_or_out_of_korea_customer_coordinate |
| 12069 | 20786 | 시흥RDC | (13)민수용(홍현걸) | invalid_or_out_of_korea_customer_coordinate |
| 12070 | 20787 | 시흥RDC | (14)민수용(김근호) | invalid_or_out_of_korea_customer_coordinate |
| 12071 | 20788 | 강릉RDC | (15)민수용(정재복) | invalid_or_out_of_korea_customer_coordinate |
| 12072 | 20789 | 대전RDC | (16)민수용(박광수) | invalid_or_out_of_korea_customer_coordinate |
| 12073 | 20790 | 대전RDC | (17)민수용(여준구) | invalid_or_out_of_korea_customer_coordinate |
| 12074 | 20791 | 광주RDC | (18)민수용(기용민) | invalid_or_out_of_korea_customer_coordinate |

## Generated Files

- `data/processed/warehouses.csv`
- `data/processed/customers.csv`
- `data/processed/od_links.csv`
- `data/processed/preparation_summary.json`
