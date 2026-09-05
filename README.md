# FundDiff

**Understand every number that changed.**

FundDiff is a private-markets workbook reconciliation tool built for the **Ylookup × Encode AI Hackathon 2026**.

It compares two versions of a fund workbook, identifies material financial movements, translates raw Excel cell changes into human-readable drivers, and reconciles those drivers against the reported NAV movement.

Instead of telling a fund manager:

> `Portfolio!D2 changed by 420000`

FundDiff surfaces:

> **Company A · Fair Value**  
> £12,100,000 → £12,520,000  
> **NAV impact: +£420,000**

Every explanation remains linked to the original workbook cell for traceability.

---

## The problem

Fund operations frequently involve financial data moving between spreadsheets, administrators, managers, and reporting processes.

When a new workbook version arrives, reviewing it is not just a matter of finding changed cells.

The important questions are:

- What changed?
- Which changes are financially material?
- What do those cells represent?
- How did each change affect NAV?
- Does the combination of underlying movements reconcile to the reported NAV change?
- Which source cells should a reviewer inspect?

FundDiff is an MVP for making that review process faster and easier to audit.

---

## What FundDiff does

Upload a **previous** and **updated** Excel workbook.

FundDiff then:

1. Parses both workbooks locally in the browser.
2. Compares worksheets and cells across both versions.
3. Detects numeric movements.
4. Uses row and column context to produce human-readable labels.
5. Separates the reported NAV movement from underlying drivers.
6. Applies deterministic impact-direction rules to supported workbook sections.
7. Ranks drivers by financial impact.
8. Calculates explained and unexplained NAV movement.
9. Links every result back to its original worksheet and cell.

When the identified drivers equal the reported NAV movement, FundDiff marks the movement as reconciled.

---

## Example

The included synthetic demo workbooks produce:

| Driver | NAV impact |
| --- | ---: |
| Company A · Fair Value | +£420,000 |
| Interest income | +£45,200 |
| LP Group A · Distribution | -£128,947 |
| Legal fees | -£18,400 |
| Administration fees | -£4,200 |
| **Explained NAV movement** | **+£313,653** |

Reported NAV:

**£102,418,291 → £102,731,944**

Net movement:

**+£313,653**

Unexplained movement:

**£0**

---

## Demo workbooks

Two synthetic Excel workbooks are included:

```text
demo/NAV_Pack_Previous.xlsx
demo/NAV_Pack_Updated.xlsx
