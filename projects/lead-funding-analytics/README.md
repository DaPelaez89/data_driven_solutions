# Lead Preparation and Funding Analysis

## Business problem

Lead files and funding reports can use different field names and formats.
Consistent customer information and statuses help teams prepare records for a
CRM, review sales, and reconcile activity for commission reporting.

## My contribution

In the original work, I transformed customer and lead information,
organized statuses, compared records using vehicle identifiers, and prepared
sales and funding summaries. The operational work used Python, pandas,
spreadsheets, and SQL-based data sources.

## Illustrative workflow

1. Map source fields such as `HomePhone`, `State`, `Vin`, and `LeadStatus`
   to a consistent CRM-oriented structure.
2. Normalize selected values, such as state abbreviations and phone numbers,
   and flag records that need review.
3. Link fictional customers, sales, transactions, and representatives using
   their IDs.
4. Summarize sales and refunds by representative. For this demonstration,
   apply the fictional percentage rates in `representatives.csv` to net
   transaction amounts.

## Fictional example

The [sample data](data/) contains five invented customers, two invented
representatives, four sales, and six transactions. It is not derived from
real customer records.

| Demonstration measure | Example result |
| --- | ---: |
| Sale transactions | 3,150.00 |
| Refund transactions | -650.00 |
| Net transaction amount | 2,500.00 |
| Example commission | 135.00 |

Representative `LF-REP-01` has net transactions of 1,000.00 and an
invented rate of 6%, producing an illustrative commission of 60.00.
These numbers and rules are synthetic; they do not represent actual company
performance or payment policies.

## Data files

- [customers.csv](data/customers.csv): fictional leads and customer statuses.
  `CustomerID` and `RepresentativeID` correspond to `customer_id` and
  `representative_id` in the sales file.
- [representatives.csv](data/representatives.csv): fictional representatives
  and example rates.
- [sales.csv](data/sales.csv): customer-to-representative sales.
- [transactions.csv](data/transactions.csv): positive sales and negative
  refunds, linked by `sale_id`.

Phone numbers, email domains, addresses, and vehicle identifiers are
fictional or deliberately invalid placeholders. The sample schema is
designed for illustration and is not an export of an operational system.
