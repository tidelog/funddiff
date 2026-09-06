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

Upload the two synthetic demo files into FundDiff and click **Compare workbooks**.

The demo files contain synthetic data created solely to demonstrate the application.

## Privacy

Workbook parsing and comparison happen locally in the browser. The current MVP does not send workbook contents to an application server for reconciliation.

## Stack

- Next.js 16
- React
- TypeScript
- Tailwind CSS
- SheetJS
- Lucide React

No backend or database is required for the current MVP.

## Run locally

    npm install && npm run dev

Then open `http://localhost:3000`.

For the demo, upload:

    demo/NAV_Pack_Previous.xlsx
    demo/NAV_Pack_Updated.xlsx

## Production build

    npm run build

## Current scope

FundDiff is a hackathon MVP, not a production accounting system.

The current semantic layer uses deterministic rules for supported workbook sections including Portfolio, Income, Expenses, and Distributions.

For example, an increase in portfolio fair value is treated as a positive NAV driver, while increases in expenses and distributions are treated as negative drivers.

These rules are intentionally deterministic and auditable. A production version would add configurable accounting mappings, workbook-schema detection, formula lineage, validation controls, and support for different administrator templates.

## Design principles

**Deterministic first:** financial arithmetic is calculated from workbook data rather than generated text.

**Traceable:** every identified driver retains its worksheet and source-cell reference.

**Human-readable:** reviewers see business context rather than only Excel coordinates.

**Private by default:** workbook analysis happens locally in the browser.

## Built at

**Rebuild Private Markets: Ylookup × Encode AI Hackathon**  
Encode Hub, Shoreditch  
5-6 September 2026

## Problem evidence

FundDiff was built around a problem described in **Call 1: NAV workflow review with a fund manager** from the hackathon anonymised interview dataset.

The fund manager described receiving NAV work from their administrator that required **six or seven iterations** before it was correct. The problem was not primarily the turnaround time of each revision. It was the repeated review loop.

The interview highlighted three connected problems:

- **Repeated iterations:** NAV and reporting outputs routinely require multiple review-and-correction cycles.
- **Missing quality control:** numbers are not consistently checked to ensure that related figures reconcile.
- **Low trust:** because outputs cannot be assumed to be correct, the fund manager has to review the numbers themselves before reporting to investors.

The fund manager described this missing reconciliation layer as a quality-control gap that creates the review burden. They also explained that the count of review turns, rather than the turnaround time of an individual turn, is what consumes their time.

FundDiff targets that review loop. Instead of manually inspecting another workbook revision, a fund manager can compare the previous and updated versions, see the financial drivers behind the reported NAV movement, identify any unexplained difference, and trace each result back to its source cell.

**Source:** Hackathon dataset, Call 1, *NAV workflow review with a fund manager* (anonymised transcript).
