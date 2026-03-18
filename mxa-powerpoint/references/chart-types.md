# MXA Chart Type Catalogue

Charts are the primary vehicle for data-driven arguments in MXA decks. Choose the chart type that most directly supports your lead.

## General Chart Rules

- **One chart per slide** — make it large and central
- **No separate chart title** — the slide lead IS the chart title
- **Annotation mandatory** — highlight the key data point with a callout
- **Source line mandatory** — bottom-left, every chart slide
- **Muted gridlines** — light grey, value axis only. Remove category axis gridlines.
- **Data labels where useful** — on bars/segments for easy reading; remove legend if only one series
- **Use template/brand colours** — never default PowerPoint colours

---

## Waterfall (Bridge) Chart

**When to use:** Show how a value builds from start to end through positive and negative contributions. Classic MXA chart.

**Examples:** Revenue bridge (FY24 → FY25), cost build-up, headcount changes, P&L walk.

**Implementation:** python-pptx doesn't have a native waterfall type. Build with stacked bar:
- Base series (invisible — set fill to no fill or slide background colour)
- Increase series (green/positive colour)
- Decrease series (red/negative colour)
- Total bars (dark/primary colour) — for starting and ending values

**Annotation:** Callout on the largest contributor with text: "Asia-Pacific expansion drove 60% of total growth"

**Data format for chart.py:**
```json
{
  "type": "waterfall",
  "categories": ["FY24 Revenue", "Organic growth", "Acquisitions", "Churn", "FX impact", "FY25 Revenue"],
  "values": [100, 25, 15, -10, -5, 125],
  "totals": [0, 5],
  "colors": {"increase": "2E7D32", "decrease": "C62828", "total": "1565C0"}
}
```

---

## Stacked Bar / 100% Stacked Bar

**When to use:**
- Stacked bar: composition of a total across categories
- 100% stacked: share/mix comparison across categories (market share, revenue mix)

**Examples:** Revenue by segment over time, market share by competitor, cost breakdown.

**Annotation:** Callout on the segment that changed most: "Digital segment grew from 15% to 32% of revenue"

---

## Column / Bar Chart

**When to use:** Compare discrete values across categories. Column (vertical) for comparison, bar (horizontal) when category labels are long.

**Examples:** Revenue by region, performance by business unit, survey responses.

**Formatting:**
- Single series: one colour throughout (primary brand colour)
- Highlight bar: make the key bar an accent colour, others muted grey
- Gap width: 80–120% of bar width

---

## Line Chart

**When to use:** Show trends over time. Use for continuous data with 5+ time periods.

**Examples:** Revenue trend, market growth, customer acquisition over quarters.

**Formatting:**
- Line weight: 2–3pt
- Data markers: small circles at each point
- If multiple series: max 3–4 lines. More becomes unreadable.
- Indexed line chart (base 100): use when comparing growth rates across items with different absolute values

---

## Scatter / Bubble Chart

**When to use:** Show relationship between two (scatter) or three (bubble) variables. The BCG growth-share matrix is the canonical MXA example.

**Examples:** Market attractiveness vs. competitive position, risk vs. return, customer value vs. acquisition cost.

**Formatting:**
- Label each point/bubble directly (no legend)
- Quadrant lines at meaningful thresholds (not just median)
- Quadrant labels in corners (e.g., "Stars", "Cash Cows", "Question Marks", "Dogs")
- Bubble size represents the third variable — include a size legend

---

## Tornado Chart

**When to use:** Sensitivity analysis showing which variables have the most impact on an outcome.

**Examples:** NPV sensitivity to discount rate, revenue, costs. Risk factor impact ranking.

**Implementation:** Horizontal bar chart with bars extending left (downside) and right (upside) from a centre line.

**Formatting:**
- Sort by total bar width (largest impact at top)
- Upside bars: accent colour. Downside bars: muted/grey.
- Label exact values at bar ends

---

## Marimekko (Mekko) Chart

**When to use:** Show both size and composition simultaneously. Each column width represents the category's total size; segments within each column show composition.

**Examples:** Market sizing by region and segment, revenue pool analysis.

**Implementation:** Not natively supported by python-pptx. Build as a series of adjacent stacked rectangles using shapes, or use matplotlib to render as an image.

**Formatting:**
- Column widths proportional to category size
- Label each segment with percentage
- Title each column with category name and total value

---

## Harvey Balls

**When to use:** Qualitative assessment of capability, maturity, or readiness across multiple dimensions in a comparison table.

**Examples:** Vendor evaluation matrix, capability assessment, strategic option scoring.

**Implementation:** Unicode characters or small circle shapes:
- ● Full (100%) — strong / fully capable
- ◕ Three-quarter (75%) — good / mostly capable
- ◑ Half (50%) — moderate / partially capable
- ◔ Quarter (25%) — weak / limited capability
- ○ Empty (0%) — absent / not capable

**Formatting:**
- Place in a table alongside criteria labels and option column headers
- Use consistent size (~14–16pt if using unicode)
- Colour: primary brand colour for filled portions

---

## Pie / Doughnut Chart

**When to use:** Share of a whole. Use sparingly — most MXA consultants prefer stacked bars.

**Only appropriate when:**
- Showing parts of 100%
- 3–5 segments maximum
- One segment is being emphasised

**Formatting:**
- Explode the key segment slightly
- Data labels with percentage and value
- Sort segments largest to smallest (clockwise from 12 o'clock)
- Avoid if segments are similar size (hard to compare)

---

## Choosing the Right Chart

| Question | Chart Type |
|----------|-----------|
| How does a value build from A to B? | Waterfall |
| How do values compare across categories? | Column/bar |
| How has a value changed over time? | Line |
| What's the composition of a total? | Stacked bar / 100% stacked |
| How do two variables relate? | Scatter |
| Which factors have the most impact? | Tornado |
| What's the market size AND composition? | Marimekko |
| How capable/mature is each option? | Harvey balls |
| What share of the total? (simple) | Pie (use sparingly) |
