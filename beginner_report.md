# Beginner Level Report — Retail Store Sales Analysis

## 1. About the data

One year of transactions (Jan–Dec 2025) for a small retail store selling four
categories — Electronics, Clothing, Groceries and Home & Living — across four
regions (North, South, East, West). 601 order lines in total, ~₹80,700 in
annual sales. I generated this dataset to look like the kind of export a small
business would actually produce — including the usual mess.

## 2. Cleaning steps (what I fixed and why)

| Problem found in raw file | Count | What I did |
|---|---|---|
| Blank `Region` values | 8 | Filled with that month's most common region (a reasonable guess for a store this size) |
| Negative `Unit_Price` (refund entries keyed in wrong) | 5 | Converted to positive — data-entry errors, not actual negative sales |
| `Category` in all caps ("ELECTRONICS") | 4 | Standardised to title case so grouping works |
| Date format inconsistencies | a few | Standardised to YYYY-MM-DD so Sheets treats them as dates |

The raw file is kept untouched in `../data/raw/` so the before/after can be
checked — the audit matches the numbers above.

## 3. Key metrics

- **Total annual sales:** ₹80,688 across 601 orders
- **Average order value:** ₹134
- **Best month:** October (₹9,916)
- **Softest month:** February (₹4,718) — October is 2.1× February
- **Top category by revenue:** Electronics — ₹46,950 (58% of all revenue)
- **Top category by order count:** Groceries (187 orders) and Clothing (170)
- **Regions:** remarkably even — East ₹20,679 / West ₹20,258 / North ₹20,100 / South ₹19,651

## 4. Trends worth noticing

1. **Electronics carries the business.** It takes 58% of revenue from only 25%
   of orders. Groceries is the mirror image — 31% of orders but just 7% of
   revenue. One category drives the money, the other drives the footfall.
2. **October is the spike, February is the dip.** Sales climb through March–April,
   stay healthy mid-year, peak in October, then slide. Oct–Dec together account
   for 28% of the year's sales — meaningful, but the store isn't purely
   festive-dependent either.
3. **Regions are evenly matched — and that's a finding in itself.** No region
   outperforms by more than ~2%. Whatever the store is doing works everywhere,
   which means the growth lever isn't "fix a weak region", it's lifting all of
   them together (or getting more from high-value categories in every region).

## 5. The dashboard

A one-page spreadsheet dashboard (`../dashboard/retail_sales_dashboard.xlsx`,
built in Google Sheets/Excel with native charts and pivot-style summary sheets):

- **Top row — KPI cards:** Total Sales ₹80,688 · Orders 601 · Avg Order Value ₹134 · Best Month: October
- **Monthly sales line chart** (Jan–Dec)
- **Revenue by category bar chart**
- **Revenue by region bar chart**
- **Category × Region revenue matrix**
- Underlying sheets: `Monthly Trend`, `Category Summary`, `Region Summary` feeding the charts

A PNG preview of the layout is in the same folder.

## 6. Business observations / recommendations

- **Protect Electronics availability in September–October.** When 58% of your
  revenue comes from one category, a stockout in peak month is the most
  expensive mistake possible.
- **Use Groceries as a traffic engine, not a profit engine.** Bundle or
  cross-sell Electronics/Clothing to grocery shoppers — they're already in the
  store (or cart) 31% of the time.
- **February is the quiet month (₹4,718).** Run clearance sales or use the
  slack for maintenance and planning rather than expecting organic demand.
- **Because regions are even, test one growth lever across all four** (e.g. a
  loyalty scheme) rather than region-specific fixes — there's no weak region
  to rescue.

## 7. Honest limitations

- One year of data lets me *describe* the October spike but not prove it's an
  annual pattern. With 2–3 years I'd separate seasonality from one-off events.
- The dataset has no cost/margin columns, so "revenue" here ≠ "profit". A
  category like Electronics could earn well on revenue but thin on margin.
- Average order value of ₹134 is dragged around by Groceries; a median view
  (₹~180-ish) tells the typical-basket story better.
