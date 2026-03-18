---
name: mxa-powerpoint
description: Create MXA-quality PowerPoint presentations with structured storylines, action-oriented leads, and consulting-grade formatting. Use when asked to create a strategy deck, consulting presentation, board pack, executive summary, or when user mentions MXA style, pyramid principle, or structured storyline. Handles both template-based editing and from-scratch creation.
metadata: {"openclaw": {"requires": {"bins": ["soffice", "pdftoppm"]}, "emoji": "📊"}}
---

# MXA PowerPoint Skill

Create consulting-grade PowerPoint decks following MXA methodology. Two disciplines govern every deck:

1. **Content strategy** — what goes on the slides (storyline, leads, so-whats)
2. **Formatting discipline** — how the slides look (alignment, font consistency, whitespace)

## Quick Reference

| Task | Command |
|------|---------|
| Read content | `python -m markitdown presentation.pptx` |
| Template thumbnails | `python {baseDir}/scripts/thumbnail.py template.pptx` |
| Rearrange slides | `python {baseDir}/scripts/rearrange.py template.pptx working.pptx 0,5,5,12` |
| Extract inventory | `python {baseDir}/scripts/inventory.py working.pptx inventory.json` |
| Apply replacements | `python {baseDir}/scripts/replace.py working.pptx replacements.json output.pptx` |
| Create from spec (auto-uses MXA template) | `python {baseDir}/scripts/create.py spec.json output.pptx` |
| Create with explicit template | `python {baseDir}/scripts/create.py spec.json output.pptx --template template.pptx` |
| Add chart | `python {baseDir}/scripts/chart.py presentation.pptx chart-spec.json` |
| Unpack for XML edit | `python {baseDir}/scripts/unpack.py input.pptx unpacked/` |
| Repack after XML edit | `python {baseDir}/scripts/pack.py unpacked/ output.pptx` |
| Clean orphans | `python {baseDir}/scripts/clean.py unpacked/` |
| Validate structure | `python {baseDir}/scripts/validate.py unpacked/` |
| Render to images | `soffice --headless --convert-to pdf out.pptx && pdftoppm -jpeg -r 150 out.pdf slide` |

---

## 1. Content Strategy

### 1.1 Ghost Deck — Always Start Here

Before touching any template or writing any code:

1. **Clarify the objective** — What decision must the audience make? What action should they take?
2. **Write leads only** — Draft every slide's lead (action title) in sequence. This is the ghost deck.
3. **Test horizontal logic** — Read leads top-to-bottom. They must tell a coherent, self-contained story. Someone reading only leads gets the full argument.
4. **Group into sections** — Organise leads into 2–4 MECE (Mutually Exclusive, Collectively Exhaustive) buckets. Each section gets a divider slide.
5. **For each slide, decide the message → structure → shape:**
   - **Message**: What single point does this slide communicate?
   - **Structure**: What archetype fits? (exec-summary, comparison, content-text, etc.)
   - **Shape**: What *content shape* best visualises this message? (See §1.6 Content Shapes below.) Use the selection heuristic — bullets is the fallback, not the default.
6. **Assign archetype + shape** — Every content slide gets both an `archetype` (purpose) and a `shape` (visual structure) in the spec JSON.
7. **Check the adjacency rule** — Never assign the same shape to consecutive content slides. If slides 3 and 4 are both `bullets`, rethink — can one become `stat-row`, `split`, or `callout-stack`?
8. **Only then start building.**

### 1.2 Storyline Frameworks

**Situation-Complication-Resolution (SCR)** — for the opening sequence (slides 1–4):
- Situation: context the audience already agrees with
- Complication: the tension or problem driving the need for action
- Resolution: your recommendation

SCR frames the opening. The body uses pyramid principle.

**Pyramid Principle (Minto)** — governing logic for the entire deck:
- State the answer first (executive summary), then support with evidence
- Each level of detail supports the level above
- Arguments at every level are MECE
- Progressive depth: exec summary → section summaries → supporting detail → appendix

### 1.3 Lead-Writing Rules

The slide title ("lead") is the single most important element on every slide.

- **Complete, action-oriented sentence** — "Revenue grew 23% driven by Asia-Pacific expansion", NOT "Revenue Overview"
- **Self-contained** — the lead alone communicates the message without reading the body
- **One lead = one message** — if a lead has "and" connecting two ideas, split into two slides
- **Answer-first** — state the conclusion, not the question
- **Specific, quantified** — "3 of 5 business units exceeded targets" not "Most units performed well"
- **Active voice** — "We recommend consolidating..." not "It is recommended that..."

### 1.4 So-What Requirement

Every slide must pass the "so what?" test:

- **In the lead** — the lead should contain the implication ("...suggesting we should accelerate investment")
- **And/or in a callout box** — a visually distinct box (coloured border or background) stating the takeaway or recommended action
- If removing a slide doesn't break horizontal logic, it shouldn't exist

### 1.5 Required Slide Elements

| Element | Rule |
|---------|------|
| Lead | Every content slide. Complete sentence, action-oriented. |
| Source line | Every data slide. Bottom-left, small italic. `Source: [Author], [Title], [Year]` |
| Footnotes | Caveats, definitions. Bottom, numbered, smaller than body. |
| Chart annotations | Every chart — callout to the key data point supporting the lead. |
| So-what | Every slide — in lead and/or callout box. |
| Page numbers | Every slide, consistent position. |

### 1.6 Content Shapes

Every content slide has two properties in the spec JSON:
- **`archetype`**: the slide's *purpose* (title, exec-summary, content-text, comparison, divider, next-steps, appendix, back-cover). Controls storyline logic.
- **`shape`**: the slide's *visual structure*. Controls which builder renders it. If omitted, the builder infers a default shape from the archetype.

#### Shape Vocabulary

| Shape | What it builds | Best for | Spec fields |
|-------|---------------|----------|-------------|
| `bullets` | Title + subtitle + 3-5 bulleted paragraphs | Qualitative arguments, exec summaries | `subtitle`, `bullets` |
| `stat-row` | Title + 2-4 large numbers with small labels | KPIs, headline metrics, before/after | `stats: [{value, label}]` |
| `n-column` | Title + N equal columns (header bar + body card) | Side-by-side comparison, categories | `columns: [{header, bullets}]` |
| `callout-stack` | Title + 2-3 stacked accent-bordered callout boxes | Key findings, quotes, risk items | `callouts: [{text, accent?}]` |
| `split` | Title + left half (bullets) + right half (callout/evidence box) | Argument + supporting detail | `bullets`, `callout` |
| `process` | Title + 3-5 numbered steps connected by arrows | Workflows, timelines, methodology | `steps: [{title, description}]` |
| `icon-cards` | Title + 3-4 cards with coloured icon circle + header + text | Capabilities, features, principles | `cards: [{icon, header, text}]` |
| `big-quote` | Large quote mark + quote text + attribution | Testimonials, executive statements | `quote`, `attribution` |
| `matrix` | Title + 2×2 grid with axis labels | Prioritisation, risk/impact, positioning | `quadrants: [{label, items}]`, `x_axis`, `y_axis` |

#### Shape Selection Heuristic

When deciding which shape to use for a content slide, work through this decision tree. **`bullets` is the last resort, not the default.**

1. **Mostly numbers or KPIs?** → `stat-row` (2-4 figures) 
2. **Comparing N things side by side?** → `n-column`
3. **Sequential process or methodology?** → `process`
4. **2-3 key quotes or findings to highlight individually?** → `callout-stack` or `big-quote`
5. **One argument supported by evidence or a takeaway box?** → `split`
6. **Capabilities, features, or principles with icons?** → `icon-cards`
7. **Prioritisation or 2×2 positioning?** → `matrix`
8. **Qualitative argument, no richer shape fits?** → `bullets`

#### The Adjacency Rule

**Never assign the same shape to consecutive content slides** (excluding title, divider, and back-cover slides). If two adjacent slides both want `bullets`, rethink — can one become `stat-row`, `callout-stack`, or `split`? Visual variety is structural, not optional.

`create.py` warns (does not block) if consecutive content slides share the same shape.

#### Text-Fit Rules for Shapes

**`stat-row`**: The big number font auto-sizes to prevent wrapping. The builder estimates character width at 0.45"/char at 54pt and scales down if the longest value string would overflow the available card width. Floor is 28pt. Keep stat values short (3-5 characters: "72%", "$4.2M", "100K+") — if a value needs 6+ characters, consider abbreviating or restructuring the data.

**All shapes**: Every text box must contain its text on a single line (for headers/titles) or within its allocated height (for body text). When building the spec, follow these limits:
- **stat-row values**: max ~5 characters per value for 4 cards, ~7 for 3 cards, ~10 for 2 cards
- **Column headers** (`n-column`): 1-3 words
- **Process step titles**: 1-2 words
- **Icon-card headers**: 1-3 words
- **Big-quote text**: 2-4 sentences max

#### Archetype → Default Shape Mapping

When `shape` is omitted from a slide spec, these defaults apply:

| Archetype | Default shape |
|-----------|---------------|
| `title` / `cover` / `back-cover` | *(no shape — dedicated builder)* |
| `divider` / `bumper` | *(no shape — dedicated builder)* |
| `exec-summary` | `bullets` |
| `content-text` / `text` | `bullets` |
| `comparison` | `n-column` |
| `next-steps` | *(dedicated table builder)* |
| `appendix` | `bullets` |

#### Spec JSON Examples

**stat-row:**
```json
{
  "archetype": "content-text",
  "shape": "stat-row",
  "lead": "AI has displaced 100,000+ workers across four sectors since 2024",
  "stats": [
    {"value": "100K+", "label": "Jobs displaced globally"},
    {"value": "72%", "label": "Drop in Indian IT hiring"},
    {"value": "$300B", "label": "Sector at risk"}
  ]
}
```

**callout-stack:**
```json
{
  "archetype": "content-text",
  "shape": "callout-stack",
  "lead": "Critics argue companies are 'AI-washing' layoffs",
  "callouts": [
    {"text": "Klarna reversed its AI-only policy after quality deteriorated"},
    {"text": "India needs 1M AI professionals but <20% of workers are AI-skilled"},
    {"text": "Many 'AI layoffs' coincide with weak demand and margin pressure"}
  ]
}
```

**split:**
```json
{
  "archetype": "content-text",
  "shape": "split",
  "lead": "The Indian IT sector faces structural transformation",
  "subtitle": "Industry Outlook",
  "bullets": [
    "Managed services face 'sharp revenue deflation'",
    "Workforce pyramid being compressed at middle layers",
    "Outcome-based billing replacing headcount-based models"
  ],
  "callout": "Bottom line: 50,000 Indian IT jobs lost in 2025, with fresh graduate absorption falling 75%"
}
```

**process:**
```json
{
  "archetype": "content-text",
  "shape": "process",
  "lead": "AI workforce transition follows a predictable four-stage pattern",
  "steps": [
    {"title": "Announce", "description": "Company cites AI as driver for restructuring"},
    {"title": "Cut", "description": "Roles eliminated, headcount reduced 15-40%"},
    {"title": "Discover", "description": "AI quality issues surface, gaps emerge"},
    {"title": "Re-hire", "description": "Selective re-hiring at higher skill levels"}
  ]
}
```

**big-quote:**
```json
{
  "archetype": "content-text",
  "shape": "big-quote",
  "lead": "Klarna's CEO acknowledged the limits of AI-first workforce strategy",
  "quote": "We went too far in cutting human roles. The AI quality wasn't there yet, and our customers noticed.",
  "attribution": "Sebastian Siemiatkowski, CEO, Klarna — May 2025"
}
```

**icon-cards:**
```json
{
  "archetype": "content-text",
  "shape": "icon-cards",
  "lead": "Three strategic responses are emerging across affected industries",
  "cards": [
    {"icon": "A", "header": "Reskill & Redeploy", "text": "Invest in AI training for displaced workers to fill new AI-adjacent roles"},
    {"icon": "B", "header": "Hybrid Teams", "text": "Pair AI tools with human oversight to maintain quality while reducing headcount"},
    {"icon": "C", "header": "Outcome Pricing", "text": "Shift from headcount billing to outcome-based models that reward efficiency"}
  ]
}
```

### 1.7 Slide Archetypes (Reference)

The archetype table below lists all supported slide purposes. Most use a content shape (see §1.6). Some have dedicated builders.

| Archetype | When to Use | Typical Shape |
|-----------|-------------|---------------|
| Title / cover | First slide | *(dedicated)* |
| Executive summary | Slide 2–3, synthesises entire recommendation | `bullets` (default) or `stat-row` |
| Divider | Section breaks | *(dedicated)* |
| Content (text) | Qualitative argument | Any shape — use the heuristic |
| Content (chart) | Data-driven argument | `bullets` + chart overlay |
| Comparison | N-column side-by-side | `n-column` (default) |
| Next steps | Actions with owners/timelines | *(dedicated table)* |
| Appendix | Backup detail slides | `bullets` (default) |
| Back cover | Closing slide | *(dedicated)* |

Full archetype specifications: `{baseDir}/references/slide-archetypes.md`

### 1.8 Language & Tone

- Active voice throughout
- Parallel grammatical structure at every bullet level
- Quantified claims — "Costs decreased 23% ($4.2M)" not "Costs decreased significantly"
- No orphan content — every element directly supports the lead
- Data over bullets — if a point can be shown with a chart, use a chart

---

## 2. Formatting Discipline

### 2.1 Core Principles

| Principle | Rule |
|-----------|------|
| Font consistency | ALL body text on a slide uses the same font and size. Only exceptions: title, subtitle, table/column headers. |
| Alignment | All objects of the same type must be pixel-aligned. Left edges in same column align. Top edges in same row align. |
| Fill the slide | Make fonts as large as practicable. Minimise dead whitespace in the centre. Edge margins generous and consistent. |
| One font family | One modern sans-serif throughout (Instrument Sans SemiBold preferred, or template font). No creative font pairing. |
| No rounded corners | `rect` geometry only. Angular edges signal rigour. |
| Consistent spacing | Pick one gap size between elements and use it everywhere on the slide. |

### 2.2 Font Size Hierarchy

| Element | Size | Font | Weight |
|---------|------|------|--------|
| Slide title (lead) | 24–28pt | Instrument Sans SemiBold | Bold |
| Subtitle | 20pt | Instrument Sans SemiBold | Bold, dark text (`231F20`) |
| Column/card headers | 18pt+ | Instrument Sans SemiBold | Bold, white on colour |
| Body text / bullets | 18pt | Inherited from placeholder/theme | Regular (not bold, not SemiBold) |
| Source / footnotes | 8–10pt | Instrument Sans SemiBold | Italic, muted grey |

Body text must never be smaller than 14pt. Choose a body size (typically 18pt) that fills the content area well. These slides are for presenting in a room.

**Two font weights**: Use "Instrument Sans SemiBold" for titles, subtitles, and headers. Body text inherits its font from the template placeholder/theme — do not set an explicit font name on body runs.

**Bullet markers**: Wingdings `§` character (small filled square), `buClr` = `schemeClr accent1` (green from theme), `buSzPct="100000"`, hanging indent `marL="252000" indent="-252000"`, `spcBef` = `spcPts val="600"` (6pt).

**Use template placeholders**: For exec-summary and content-text slides (Layout 5), use the built-in placeholders (idx=13 for subtitle, idx=14 for content) — do NOT add separate textboxes. This preserves template positioning and inherited formatting.

### 2.3 Colour Usage

- **From template**: extract palette from `theme1.xml` and use exclusively
- **From scratch**: MXA brand palette by default — dark green primary (#195B44), green accent (#38B34A), dark text (#231F20). Rotate accent colours for multi-column headers: dark green, orange (#F47E4D), blue (#248DC1), purple (#7E3794).
- **Multi-column layouts**: rotate through accent colours (`195B44`, `F47E4D`, `248DC1`, `7E3794`) for header bars
- **Bold key terms** in body text using the accent colour

### 2.4 Chart Formatting

- One chart per slide — the chart IS the slide; surrounding text is minimal framing
- The slide lead serves as chart title — do not add a separate chart title
- Annotation callouts mandatory — highlight the key data point, arrow/line to text box stating the insight
- Source line mandatory at bottom-left for every chart slide
- Native PowerPoint charts via `{baseDir}/scripts/chart.py` — editable in PowerPoint
- MXA chart type catalogue: `{baseDir}/references/chart-types.md`

### 2.5 Text-Fit Verification

Every text box must contain its text without clipping:

1. Estimate line count — usable width = box width − insets. At 14pt, ~55–60 chars/line.
2. Calculate required height — lines × line height + insets.
3. If text overflows — increase box height, reduce text, or split across shapes.
4. **Always visually verify** by rendering to images. Font metrics differ between LibreOffice and PowerPoint.

### 2.6 MXA Template Positioning Reference

These EMU-based constants come from the MXA template and govern all slide geometry. Use them when editing XML directly or verifying shape positions.

**Slide dimensions:** `cx="12192000"` `cy="6858000"` (standard 16:9 widescreen)

**Content X margins:** Content runs from `x="838200"` with width `cx="10515600"`, providing consistent left/right margins.

| Element | Y position (EMU) | Height (cy) | Notes |
|---------|-------------------|-------------|-------|
| Title | `864482` | `480131` | Bottom-anchored (`anchor="b"`). See title box rules below. |
| Subtitle | `1512027` | `409343` | Between title and green line. Bold SemiBold, `231F20`. |
| Green horizontal line | ~`1800000` | — | Built into the template layout. Visual divider. |
| Content area start | `2018377` | — | Below the green line — template default for body placeholders. |
| Custom shapes/cards | `1920000`+ | — | For manually positioned content (columns, cards), start at or below this Y value. |
| Slide number / footer | `6324332` | — | Bottom of slide. |

#### Title Box Rules (Critical)

The title text box uses `anchor="b"` (bottom-anchored) with a deliberately compact height (`cy="480131"`). The bottom edge sits at approximately y + cy = 1,344,613 EMU — well above the green line. This means:

- **Single-line titles**: text sits neatly at the bottom of the box.
- **Multi-line titles**: text overflows **upward** from the anchored bottom. The last line stays at the same position; additional lines extend toward the top of the slide.

This keeps the title's bottom edge at a fixed distance above the green line regardless of line count. **Never increase the title box height to "fit" multi-line titles** — the upward overflow is intentional.

#### Title Colour Rule

Do NOT add explicit `<a:solidFill>` colour overrides to title text. Let titles inherit the template's default colour (green). Adding `231F20` or any explicit colour turns titles black and breaks template styling.

#### Subtitle Placeholder Caution

The MXA template subtitle placeholder renders at ~24pt by default. Long subtitles (>60 characters) will wrap to multiple lines and extend below `y=1512027`. When placing custom shapes below a subtitle:
- Default: start custom content at `y >= 2050000`
- Long subtitle (>60 chars): start at `y >= 2200000`

#### Aspect Ratio Rule

When resizing an image or chart to fit below the green line, always maintain the original aspect ratio. If you reduce the height, reduce the width proportionally and re-centre horizontally. Never stretch or squash images to fill available space.

#### Text-Fit with EMU Calculations

For precise text-fit verification when editing XML:
1. **Usable width** = `cx - 216000` (subtracting default left + right insets of 108000 each)
2. **Characters per line**: at `sz="1400"` (~14pt), assume ~55–60 chars/line
3. **Line heights**: ~220000 EMU at 14pt, ~240000 EMU at 15pt. Add top/bottom insets + spacer paragraphs (`sz="400"` ≈ 60000 EMU each).
4. If text overflows: increase `cy`, reduce `sz`, shorten text, or split across shapes.

### 2.7 Template Slide Philosophy

Template slides in the MXA PowerPoint Template are **inspiration and accelerators**,
not constraints. You can and should build custom layouts when the template doesn't
have a slide that fits the content.

- Template *layouts* (slide masters) provide branding: green accent line, MXA logo,
  corner decoration, footer placeholder, colour theme, fonts.
- Template *slides* show examples of how to use the brand — treat them as a gallery.
- When building a slide, pick the layout that gives you the right branding elements,
  then construct the content area with shapes positioned to MXA standards.
- **Never force content into a template slide that doesn't match.** If you have 3
  comparison columns, build 3 columns — don't use a 4-column template and leave one
  blank or try to cram content into the wrong shape.

---

## 3. Workflows

### 3.1 Template-Based Editing (Primary)

Use when the user provides a .pptx template.

```
1. Analyse template
   python {baseDir}/scripts/thumbnail.py template.pptx
   python -m markitdown template.pptx

2. Ghost deck — write all leads, test horizontal logic, assign archetypes

3. Map leads to template slide layouts

4. Build slide sequence
   python {baseDir}/scripts/rearrange.py template.pptx working.pptx 0,5,5,12,20

5. Extract shape inventory
   python {baseDir}/scripts/inventory.py working.pptx inventory.json

6. Build replacement JSON matching inventory structure (see inventory.py --help)

7. Apply replacements
   python {baseDir}/scripts/replace.py working.pptx replacements.json output.pptx

8. Add charts where needed
   python {baseDir}/scripts/chart.py output.pptx chart-spec.json

9. For complex edits requiring direct XML control:
   python {baseDir}/scripts/unpack.py output.pptx unpacked/
   (edit XML directly)
   python {baseDir}/scripts/clean.py unpacked/
   python {baseDir}/scripts/pack.py unpacked/ final.pptx

10. QA (see Section 4)
```

**Subagent parallelisation**: After step 5, use subagents to build replacement content for multiple slides in parallel. Also use subagents for visual QA (fresh eyes on rendered images).

### 3.2 MXA Template Slide Catalogue

The bundled template contains 29 example slides (indices 0–28). Consulting-content slides (12–26) use Layout 4 with placeholders idx=0 (Title), idx=13 (Subtitle), idx=15 (Section Number), idx=4 (Footer), plus custom shapes. Shape naming follows a consistent pattern (e.g., `Cell 0_0`, `Bar 1`, `Phase Label 2`) — row first, column second, both 0-based.

| Slide | Name | Key shapes | When to use |
|-------|------|------------|-------------|
| 0 | Title Slide | CENTER_TITLE + SUBTITLE | Opening slide |
| 1 | Table of Contents | Table shape | Agenda / contents page |
| 2 | Executive Summary I | idx=13 subtitle, idx=14 content, idx=15 sidebar, idx=16 sidebar title | Full-width exec summary with right sidebar |
| 3 | Section Header | Title + section number + subtitle | Section dividers |
| 4 | 1-Column Content | idx=13 subtitle, idx=14 content, idx=15 section number | Standard single-column text/bullet slides |
| 5 | 2-Column with Picture | Content + Picture placeholder | Text + image side by side |
| 6 | 2-Column with Subtitles | idx=14 + idx=17 content columns, idx=13 + idx=16 subtitles | Two text columns with independent subtitles |
| 7 | 4 Photo Placeholders | 4 picture placeholders (idx 20-24) | Photo gallery |
| 8 | 4 Photo Placeholders (Labeled) | 4 pictures + 4 title labels | Labeled photo gallery |
| 9 | 4 Content (Labeled) | 4 content areas + 4 labels | Quad-panel content |
| 10 | 4-Column Content | 4 paired subtitle/content columns | Multi-column comparison |
| 11 | 4-Column Block Colour | 4 content blocks + icon/picture decorations | Colour-blocked comparison |
| 12 | **SCR Flow** | Pentagon + 2 Chevrons + 3 TextBoxes + vertical connectors | Situation → Complication → Resolution |
| 13 | **From/To Comparison** | Rectangles + connectors + TextBoxes (From/To) + Group | Before/after or current/future state |
| 14 | **Complex Content** | Groups, rectangles, graphics, textboxes | Detailed architecture or process diagrams |
| 15 | **Content with Group** | Group shape + content | Grouped visual content |
| 16 | **Case Study** | Left Header/Body + Right Header/Body (2-col, each w=5.75") | Client case study: Challenge \| Outcomes |
| 17 | **Client Testimonial** | Quote Mark + Quote Body + Attribution + Accent Bar | Direct quote with attribution |
| 18 | **Team Grid** | 2×3 grid: Head/Body (headshot) + Name + Role per person | Team overview (6 people) |
| 19 | **Team Bio** | Head/Body (headshot) + Bio Card | Individual biography |
| 20 | **Approach / Process** | Step circles (1-5) + Phase labels + Descriptions + connector Lines | N-step phased methodology |
| 21 | **Project Timeline** | Week headers + Phase Labels + Row BGs + Bars (Gantt) | Timeline / Gantt chart |
| 22 | **Client Logos** | 3×4 grid of Logo textboxes | Logo wall / client list |
| 23 | **Key Metrics** | Stat + Line + Label (4 groups) | KPI / stat callout blocks |
| 24 | **Risk Assessment** | Hdr 0-3 + Cell row×col grid (4 rows × 4 cols) | Risk/RAID table |
| 25 | **Engagement Governance** | Steering/PM/Tech Lead/Working Team boxes + elbow connectors | Org chart / governance |
| 26 | **Typical Week** | Day headers (Mon-Fri) + 3 rows of Slot shapes | Weekly schedule grid |
| 27 | Logo Divider | (empty) | End-of-section divider with logo |
| 28 | Back Cover | Footer only | Closing slide |

To clone and populate a template slide for the `rearrange.py` → `replace.py` workflow, find shapes by name prefix and index.

### 3.3 From-Spec Creation (Preferred for new decks)

`create.py` uses a **hybrid approach**: it opens the MXA template to access
slide layouts (which inherit the master's green line, logo, corner decoration,
and footer), then builds **custom content using shapes** positioned to MXA
style standards. Template slides are treated as inspiration — the script does
not clone specific template slides. This means any layout (N columns, custom
structures) can be built while still inheriting MXA branding.

The template (`MXA Powerpoint Template.pptx`) is bundled in the skill folder.
It is auto-detected by searching the skill folder, the working directory, and
parent directories — or can be specified via `--template` or the spec's `"template"` key.

**How it works:**
- Title slides use Layout 0 (Title Slide) for CENTER_TITLE + SUBTITLE
- Content slides with subtitle + bullets use Layout 5 (Numbered Title, Subtitle & 1-Col Content) for TITLE + idx=13 subtitle + idx=14 content
- Content slides with custom shapes (comparison, next-steps) use Layout 6 (Numbered Title Only) for TITLE + green line + footer
- Custom content (bullets, columns, cards, tables) is built below the green line
  using textboxes and rectangles positioned to MXA standards
- Falls back to from-scratch generation when no template is found

```
1. Ghost deck — write leads, decide message → structure → layout for each slide

2. Create slide spec JSON (see create.py --help for format)
   - archetype determines what builder to use
   - comparison archetype builds exactly N columns from the spec
   - no need to match template slide indices

3. Generate presentation (template auto-detected)
   python {baseDir}/scripts/create.py spec.json output.pptx

4. Add charts
   python {baseDir}/scripts/chart.py output.pptx chart-spec.json

5. QA (see Section 4)
```

---

## 4. QA Process (Mandatory)

**Assume there are problems. Your job is to find them.**

### Content QA

```bash
python -m markitdown output.pptx
```

1. **Lead read test** — read all leads in sequence. Do they tell a coherent story?
2. **So-what check** — does every slide have a clear so-what?
3. **Source check** — does every data slide cite its source?
4. **Formatting audit** — are body font sizes consistent within each slide?
5. **Layout-content match** — does the visual structure match the number of items? (e.g., 3 comparison items → 3 columns, not 4)
6. **No dead space** — no large blank areas where content should be or unused template regions.

### Visual QA

Convert to images and inspect — **use a subagent for this** (fresh eyes):

```bash
soffice --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

Inspect every slide image for:
- Overlapping elements, text overflow, cut-off text
- Alignment issues — elements that should line up but don't
- Inconsistent spacing or margins
- Low contrast text
- Dead whitespace in the middle of the slide
- Too-small fonts

### Verification Loop

1. Generate → render to images → inspect
2. List issues (if zero found, look harder)
3. Fix issues
4. Re-render and re-inspect affected slides
5. Repeat until a full pass reveals no new issues

**Do not declare success until at least one fix-and-verify cycle completes cleanly.**

### Finalising Checklist

Before delivering any deck, verify every item:

- [ ] Title at `y=864482`, `cy=480131`, `anchor="b"` — compact box, text overflows upward for multi-line
- [ ] All content below the green line (`y >= 1920000` for custom shapes)
- [ ] Colours are from the MXA palette only — no generic reds, blues, or defaults
- [ ] Shapes use `rect` geometry, not `roundRect`
- [ ] Body text is `sz="1400"` or larger, and consistent within each slide
- [ ] Footer matches the standard format (`Title | Month YYYY | slidenum`)
- [ ] One key message per slide
- [ ] Minimal text — would a presenter be comfortable talking to these bullets?
- [ ] All text fits within its containing shape — no clipping, no overflow, verified visually after render
- [ ] Images/charts maintain original aspect ratio

---

## 5. Dependencies

**Python** (pip install):
```
python-pptx defusedxml lxml "markitdown[pptx]"
```

**System**:
- `libreoffice` — `soffice --headless --convert-to pdf`
- `poppler-utils` — `pdftoppm` for PDF to images
