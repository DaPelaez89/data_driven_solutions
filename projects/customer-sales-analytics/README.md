# Customer Status and Sales Reconciliation

## Business problem

Customer records, sales, and refunds can come from different sources. The
analyst needs to organize customer statuses and reconcile transactions by
representative so sales activity can be reviewed for commission payments.

## My contribution

In the original work, I prepared and combined customer and transaction data,
reviewed customer statuses, checked records for duplicates, and created
representative-level sales and funding reports. The operational work used
Python, pandas, spreadsheets, and SQL-based data sources.

## Illustrative workflow

1. Review customer records and separate active, pending, and cancelled
   statuses.
2. Link sales to customers and representatives using IDs.
3. Join sale and refund transactions to each sale and calculate net sales.
4. Group net sales by representative for review. In this example only,
   multiply net sales by the fictional rate in `representatives.csv`.

## Fictional example

The [sample data](data/) contains six invented customers, two invented
representatives, five sales, and seven transactions. It is not derived from
real customer records.

| Demonstration measure | Example result |
| --- | ---: |
| Sale transactions | 3,550.00 |
| Refund transactions | -700.00 |
| Net transaction amount | 2,850.00 |
| Example commission | 130.00 |

For example, representative `CS-REP-01` has net transactions of 1,600.00
and an invented rate of 5%, producing an illustrative commission of 80.00.
These numbers and rules are synthetic; they do not represent actual company
performance or payment policies.

## Data files

- [customers.csv](data/customers.csv): fictional customers and statuses.
- [representatives.csv](data/representatives.csv): fictional representatives
  and example rates.
- [sales.csv](data/sales.csv): customer-to-representative sales.
- [transactions.csv](data/transactions.csv): positive sales and negative
  refunds, linked by `sale_id`.

Phone numbers, email domains, and vehicle identifiers are reserved or
deliberately invalid placeholders. The sample schema is designed for
illustration and is not an export of an operational system.
