---
name: bid-leveling
description: Level subcontractor bids for one trade. Use when someone shares two or more subcontractor quotes (PDF, email or text) and asks to compare, level, or check them for scope gaps. Produces an Excel leveling sheet with every scope item marked Included, Excluded, Qualified or Not mentioned, backed by the quote's own words.
---

# Bid leveling

You help a general contractor's estimator compare subcontractor quotes for one trade on an apples-to-apples basis. You read; the estimator decides. Never pick a winner, never invent a price, and never mark anything Included without a quote line that says so.

## Inputs

- Two or more quotes for the same trade: PDFs, pasted email text, or images.
- Optional: the project name, the bid due date, and any scope notes from the estimator (for example "fixtures are owner-furnished").

This free version ships with the **electrical** checklist (`checklists/electrical.csv`). If the quotes are for another trade, say so and offer to level them against a short checklist you draft with the estimator, clearly labeled as unreviewed.

## Steps

1. **Confirm the set.** List each bidder's company name, quote date, and base price as written. Ask once if anything needed is missing (a page, a price). If a quote has no total, record the price as null and flag it.
2. **Read every quote fully**, including fine print, exclusions lists, qualifications, and footers. Exclusions often sit at the end.
3. **Check each checklist item for each bidder** and assign exactly one status:
   - `included`: the quote clearly states it is in scope.
   - `excluded`: the quote clearly states it is not in scope, or "by others".
   - `qualified`: included with a condition or limit (an allowance, "up to", "per plans dated", "one mobilization", "assumes existing service is adequate").
   - `not_mentioned`: the quote says nothing about it. This is the most common status and the main thing estimators need to see.
   For every status other than `not_mentioned`, copy the shortest exact phrase from the quote that supports it into `quote`, plus the page number if known. If you are unsure, use `not_mentioned` and say why in `note`.
4. **Capture extras**: alternates (description + price), allowances (description + amount), unit prices, the quote's validity period, schedule assumptions, and any item in the quote that does not match the checklist (put it in `unmatched`).
5. **Mark items that don't apply.** If the estimator's notes or the quotes make clear an item isn't part of this job (for example site work on an interior remodel), list its id in `not_applicable`. When unsure, leave it applicable; the estimator can switch any row to N in the sheet.
6. **Draft clarification questions** for each bidder: one question per `not_mentioned` item that matters for price, plus any qualification that needs pinning down. Keep them short and specific, for example "Does your price include the utility company's service connection fees?"
7. **Write `bids.json`** in the format below, then run:
   `python3 scripts/build_sheet.py bids.json "<Project> - Electrical leveling.xlsx"`
   The script needs `openpyxl` (`pip install openpyxl`).
8. **Report back in a few lines**: the bidders and their base prices, each bidder's gaps (excluded plus not mentioned, counting only items that apply) with qualified items reported separately, matching the Summary tab, the three gaps most likely to move price, and the file name. Remind the estimator to enter their own cost estimates for gaps in the "Adjustment $" columns so the adjusted totals compare fairly.

## bids.json format

```json
{
  "project": "Riverside Clinic TI",
  "trade": "Electrical",
  "checklist": "electrical",
  "bid_date": "2026-11-04",
  "not_applicable": ["E31", "E32", "E33"],
  "bidders": [
    {
      "name": "Acme Electric",
      "quote_date": "2026-10-28",
      "valid_days": 30,
      "base_price": 412500,
      "items": {
        "E01": {"status": "included", "quote": "Includes all electrical permits", "page": 1},
        "E11": {"status": "excluded", "quote": "Temporary power by GC", "page": 2},
        "E27": {"status": "qualified", "quote": "Fixture allowance of $45,000", "page": 1, "note": "Allowance, not a firm price"}
      },
      "alternates": [{"description": "Add generator and ATS", "price": 68000}],
      "allowances": [{"description": "Light fixtures", "amount": 45000}],
      "unmatched": ["Includes one-year service contract on lighting controls"],
      "questions": ["Does your price include utility company fees for the new service?"]
    }
  ]
}
```

Items left out of `items` are treated as `not_mentioned`. Items in `not_applicable` still show in the sheet but are left out of the gap counts.

## Rules

- Quote the bidder's words; never paraphrase into a status you can't support.
- Do not judge which bidder is best or cheapest overall. Gaps and adjustments decide that, and they belong to the estimator.
- Do not estimate costs for gaps. Leave the adjustment cells empty.
- Keep bidder information private to this conversation; never write it anywhere except the output file.
