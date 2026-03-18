# MXA Slide Archetypes

Detailed layout specifications for each slide type. All positions assume 16:9 widescreen (13.333" × 7.5").

## Standard Slide Zones

Every content slide follows the same zone structure:

| Zone | Y Range (inches) | Purpose |
|------|-------------------|---------|
| Title zone | 0.3 – 1.3 | Lead (action title) |
| Subtitle zone | 1.3 – 1.7 | Optional subtitle or framing sentence |
| Content zone | 1.8 – 6.3 | Charts, text, diagrams, tables |
| Footer zone | 6.5 – 7.2 | Source line, footnotes, page number |

Standard margins: 0.5 – 0.75" from left/right edges. Content width: ~11.8".

---

## Title / Cover Slide

**Purpose:** First slide. Sets context, establishes credibility.

**Layout:**
- Deck title — centred, 32–40pt bold, positioned mid-slide (y ~2.5–3.5")
- Subtitle — 18–20pt regular, below title
- Date — 14pt, below subtitle
- Confidentiality notice — 10pt italic, bottom of slide (optional)
- Company logo — top-right or bottom-right corner (if available)

**Content rules:**
- Title should be concise (1–2 lines max)
- Subtitle contains the audience and purpose: "Prepared for [Client] Board of Directors"
- No lead sentence needed — this is the only slide without one

---

## Executive Summary

**Purpose:** Synthesises the entire recommendation in one slide. A busy executive reads only this and gets the full answer.

**Layout:**
- Lead — 22–26pt bold, title zone. States the overarching recommendation.
- 3–5 action-oriented bullets — each summarises a major section of the deck
- Optional: so-what callout box at bottom highlighting the most critical action

**Content rules:**
- Each bullet is a complete sentence stating a finding + implication
- Bullets map 1:1 to deck sections (MECE)
- Specific and quantified: "Consolidating 3 data centres saves $4.2M annually" not "Consolidate data centres for savings"
- Order bullets by impact (highest first) or by logical flow

---

## Divider / Bumper Slide

**Purpose:** Section break that signals structure transition.

**Layout:**
- Section title — 28–36pt bold, centred or left-aligned, mid-slide
- Optional framing sentence — 16pt regular, below title. One line summarising the section's argument.
- Optional section number — large bold numeral (60–80pt) as an accent element
- Background should contrast with content slides (slightly darker or an accent colour fill)

**Content rules:**
- Section title should be an action statement, not a topic label: "Three initiatives drive 15% growth" not "Growth initiatives"
- Keep to 1–2 lines maximum

---

## Content Slide — Chart

**Purpose:** Data-driven argument. The chart IS the slide.

**Layout:**
- Lead — title zone. Action sentence stating what the data shows.
- Chart — fills content zone (y 1.8 – 6.0, width ~10"). Make it large.
- Annotation callout — text box with arrow/line pointing to the key data point
- Source line — 8–10pt italic, bottom-left of footer zone

**Content rules:**
- One chart only. If two charts are needed, use two slides.
- No separate chart title — the lead serves this purpose
- Annotation is mandatory: circle/highlight the key data point + text box explaining its significance
- Minimal surrounding text — at most 1–2 framing bullets beside the chart
- Chart must directly support the lead. If the chart doesn't prove the lead, change the chart or the lead.

**Chart positioning:**
```
Chart area: x=0.75, y=1.8, w=10.0, h=4.2
Annotation box: positioned near the highlighted data point
Source line: x=0.75, y=6.6, w=10.0
```

---

## Content Slide — Text

**Purpose:** Qualitative argument when data isn't available. Use sparingly — MXA decks are chart-heavy.

**Layout:**
- Lead — title zone
- 3–5 bullet points — content zone, 14–16pt regular, left-aligned
- So-what callout box — bottom of content zone or right column, highlighted with accent colour border
- Source line — if referencing external data

**Content rules:**
- Bullets are short — one line each where possible
- Presenter elaborates verbally; slides are visual anchors, not scripts
- Parallel grammatical structure across all bullets
- Bold key terms at the start of each bullet
- If a slide has more than 5 bullets, split into two slides

---

## Comparison Slide

**Purpose:** Side-by-side analysis of 2–3 options, scenarios, or categories.

**Layout:**
- Lead — title zone. Decisive: "Option A outperforms on 4 of 5 criteria" not "Option comparison"
- 2–3 equal-width columns with:
  - Header bar — accent colour fill, white bold text (16–18pt)
  - Body card — light grey fill (#F2F2F2), body text (14–16pt)
- Rotate accent colours for each column header

**Content rules:**
- Headers are descriptive but concise (1–3 words)
- Body content has the same structure across all columns for easy scanning
- Highlight the winning/recommended option with a visual cue (bold border, check icon)
- Include a so-what row or callout box at the bottom stating the recommended choice

**Column positioning (3-column):**
```
Column 1: x=0.75, w=3.6
Column 2: x=4.5, w=3.6
Column 3: x=8.25, w=3.6
Header height: 0.5"
Body starts below header
Gap between columns: ~0.15"
```

---

## Process Flow Slide

**Purpose:** Sequential steps, timelines, or workflows.

**Layout:**
- Lead — title zone
- Horizontal flow of 3–5 steps:
  - Numbered boxes connected by arrows
  - Each box has: step number (large, accent colour), title (bold), 1–2 line description
- Optional timeline bar below the flow

**Content rules:**
- Maximum 5 steps. If more, group into phases on separate slides.
- Steps should be concrete actions, not vague phases
- Arrows indicate sequence/dependency, not decoration

**Flow positioning (4-step):**
```
Step boxes: y=2.5, height=2.5, equal width with gaps
Arrow connectors between boxes
Step 1: x=0.75, w=2.6
Step 2: x=3.6, w=2.6
Step 3: x=6.45, w=2.6
Step 4: x=9.3, w=2.6
```

---

## Next Steps Slide

**Purpose:** Near-end slide with concrete actions.

**Layout:**
- Lead — "Three immediate actions required to capture the $12M opportunity"
- Table or structured list with columns:
  - Action (what)
  - Owner (who)
  - Timeline (when)
  - Success metric (optional)

**Content rules:**
- Every action is specific: "Finance team validates savings model" not "Continue analysis"
- Every action has an owner and deadline
- 3–6 actions maximum
- Ordered by priority or timeline

---

## Appendix Slide

**Purpose:** Detailed backup data referenced from the main deck. Not presented; available for Q&A.

**Layout:**
- Title prefixed with "Appendix:" — "Appendix: Detailed revenue breakdown by region"
- Content fills the slide — can be denser than main slides (smaller fonts acceptable: 10–12pt)
- Appendix slide numbers continue from main deck

**Content rules:**
- Referenced from main deck via footnotes: "See Appendix slide 24 for detailed breakdown"
- Can contain multiple charts, detailed tables, methodology explanations
- Less strict formatting — density is acceptable here
- Still needs a descriptive lead (action title), not just a topic label
