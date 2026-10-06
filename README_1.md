# Excel Formulas & Functions Fundamentals

**Task 6 – Excel Formulas & Functions Fundamentals**
Veda Technology | Data Analytics Track | Level 1, Day 6

## About
I practiced the core Excel functions on the Sample Superstore dataset (9,994 rows). Every function is demonstrated on live data in the `Formulas Demo` tab. Next to each result I show the formula (using FORMULATEXT) and a short note on when to use it.

## Workbook structure
| Tab | What it has |
|-----|-------------|
| Raw Data | The Superstore data, stored as an Excel Table named `SalesData` (a named range, so formulas are easy to read) |
| Formulas Demo | 17 examples: what I did, the live result, the formula used, and when to use it |

## Functions covered
| Group | Functions |
|-------|-----------|
| Lookup | VLOOKUP, XLOOKUP, INDEX/MATCH |
| Logic | IF, nested IF |
| Conditional totals | SUMIF, SUMIFS, COUNTIFS, AVERAGEIFS |
| Text | UPPER, LEFT + FIND, TRIM + PROPER, CONCAT, LEN |
| Error handling | IFERROR (lookup not found), text-vs-number mismatch, blank cell check |

## What I learned
- XLOOKUP is easier and safer than VLOOKUP because it does not need a column number
- SUMIFS can use more than one condition, SUMIF only one
- Always test formulas on edge cases such as blank cells and text-vs-number mismatches
- Using a named table (`SalesData`) makes formulas much easier to read

## Files
- `Excel_Formulas_Functions.xlsx` – the workbook
- `README.md`

## Tools used
Microsoft Excel

## Author
Rabia
