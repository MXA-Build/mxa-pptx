# MXA PowerPoint Skill

A skill for AI coding agents (GitHub Copilot, Claude, OpenClaw) that generates consulting-grade PowerPoint presentations following MXA methodology.

## What It Does

- Builds structured storylines using the Pyramid Principle and SCR framework
- Generates `.pptx` files from a simple JSON spec
- Supports 9 content shapes for visual variety: bullets, stat-row, n-column, callout-stack, split, process, icon-cards, big-quote, matrix
- Includes an adjacency rule to prevent monotonous layouts
- Uses the bundled MXA template for consistent branding (green line, logo, footer)
- Also works from scratch when no template is available

## Installation

### OpenClaw

OpenClaw loads skills from three locations (highest to lowest precedence):

1. **Workspace skills** — `<workspace>/skills/` (per-agent)
2. **Managed skills** — `~/.openclaw/skills/` (shared across all agents)
3. **Bundled skills** — shipped with the install

Clone or copy the `mxa-powerpoint` folder into one of these:

```bash
# Option A: shared across all agents
git clone https://github.com/<owner>/mxa-powerpoint.git ~/.openclaw/skills/mxa-powerpoint

# Option B: per-agent workspace
git clone https://github.com/<owner>/mxa-powerpoint.git <workspace>/skills/mxa-powerpoint
```

OpenClaw will auto-discover the skill on the next session. You can verify it loaded via the Web Control UI or by asking the agent to list its skills.

> **Note:** The skill requires `soffice` and `pdftoppm` on `PATH` for visual QA. These are declared in `metadata.openclaw.requires.bins` in `SKILL.md` — if they're missing, OpenClaw will filter the skill out at load time. Install the system dependencies below first, or remove the `requires.bins` gate if you only need deck generation (no rendering).

### GitHub Copilot

Place in `.github/copilot-skills/` or reference in your Copilot skills configuration.

### Other agents

Point your agent at `SKILL.md` as a system prompt or instruction file. The skill follows the [AgentSkills](https://agentskills.io/) spec.

## Dependencies

### Python packages

```bash
pip install python-pptx defusedxml lxml "markitdown[pptx]"
```

### System tools (for QA / rendering)

| Tool | Purpose | Install |
|------|---------|---------|
| LibreOffice Impress | Convert `.pptx` → PDF | `apt install libreoffice-impress` / `brew install --cask libreoffice` / [libreoffice.org](https://www.libreoffice.org/) |
| poppler-utils | Convert PDF → images for visual QA | `apt install poppler-utils` / `brew install poppler` / [poppler](https://poppler.freedesktop.org/) |

LibreOffice and poppler are only needed for the visual QA step (rendering slides to images for inspection). The core presentation generation works with just the Python packages.

## Quick Start

1. Write a slide spec as JSON:

```json
{
  "title": "Q3 Strategy Update",
  "footer_title": "Strategy Update",
  "footer_date": "March 2026",
  "slides": [
    {
      "archetype": "title",
      "lead": "Q3 Strategy Update",
      "subtitle": "Board Presentation — March 2026"
    },
    {
      "archetype": "exec-summary",
      "lead": "Three priorities will drive growth in Q3",
      "subtitle": "Executive Summary",
      "bullets": [
        "Expand APAC sales team by 40% to capture $2B addressable market",
        "Launch self-serve tier to reduce CAC by 30%",
        "Migrate core platform to event-driven architecture by Q4"
      ]
    },
    {
      "archetype": "content-text",
      "shape": "stat-row",
      "lead": "Q2 results exceeded targets across all key metrics",
      "stats": [
        {"value": "$42M", "label": "Revenue (+23% YoY)"},
        {"value": "91%", "label": "Gross retention"},
        {"value": "1.4x", "label": "LTV/CAC ratio"}
      ]
    }
  ]
}
```

2. Generate the deck:

```bash
python mxa-powerpoint/scripts/create.py spec.json output.pptx
```

The script auto-detects the bundled MXA template. Use `--template path.pptx` to specify a different one.

## Folder Structure

```
mxa-powerpoint/
├── SKILL.md                        # Full skill instructions (the agent reads this)
├── MXA Powerpoint Template.pptx    # Bundled presentation template
├── README.md                       # This file
├── references/
│   ├── chart-types.md              # MXA chart type catalogue
│   ├── ooxml-reference.md          # Office Open XML reference
│   └── slide-archetypes.md         # Slide archetype specifications
└── scripts/
    ├── create.py                   # Generate deck from JSON spec
    ├── chart.py                    # Add native PowerPoint charts
    ├── inventory.py                # Extract shape inventory from .pptx
    ├── replace.py                  # Apply text/content replacements
    ├── rearrange.py                # Reorder/duplicate slides
    ├── thumbnail.py                # Generate slide thumbnails
    ├── unpack.py                   # Unpack .pptx for XML editing
    ├── pack.py                     # Repack XML into .pptx
    ├── clean.py                    # Remove orphaned relationships
    └── validate.py                 # Validate OOXML structure
```

## License

See individual script headers for license information.
