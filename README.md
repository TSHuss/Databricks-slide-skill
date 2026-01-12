# Slide Deck Skill

Generate PowerPoint presentations with Databricks branding. Output files import directly into Google Slides.

## Installation

### Claude.ai (Web)

1. Download this repo as a ZIP
2. Go to [claude.ai](https://claude.ai) → Settings → Capabilities
3. Add as skill → Upload the ZIP

### Claude Code (CLI)

Share this repo with Claude Code and ask it to install the skill. Claude will handle the setup.

**Dependency:** The skill requires `python-pptx`. Claude will install it, or run `pip install python-pptx` manually.

---

## How It Works

Ask Claude to create a presentation. Claude will first determine how to gather content:

### Content Mode

You have existing materials—design docs, meeting notes, outlines, or data. Claude reads and analyzes your content, asks clarifying questions, then transforms it into a structured deck.

### Interview Mode

You're starting from scratch. Claude interviews you to understand your topic, audience, goals, and key points. Once Claude has enough detail, it designs and generates the deck.

In both modes, Claude doesn't just dump content onto slides. It thinks about **presentation design**—choosing slide types that match the shape of your content (timelines for processes, comparisons for trade-offs, callouts for key messages) and varying the visual rhythm to keep the audience engaged.

---

## Visual Design

The generator uses a mix of light and dark backgrounds to create visual rhythm:

- **Light backgrounds**: Content-heavy slides (bullet points, columns, cards, stats, timelines) use light backgrounds for readability
- **Dark backgrounds**: Structural slides (title, section, callout, quote, closing) use Databricks dark templates for visual impact

This mirrors professional presentation design where section breaks and key moments use contrasting backgrounds.

---

## Slide Types

| Type | Description |
|------|-------------|
| `title` | Opening slide with logo and accent triangles |
| `agenda` | Numbered list with hexagon bullets |
| `section` | Section divider to break up topics |
| `content` | Standard bullet points |
| `two-column` | Side-by-side comparison with headers |
| `three-column` | Three columns for options or tiers |
| `comparison` | VS layout with diamond center element |
| `pros-cons` | Two columns with checkmarks and X marks |
| `big-number` | Hero statistic as the focal point |
| `stat-row` | Multiple metrics displayed in a row |
| `callout` | Bold single statement for emphasis |
| `timeline` | Sequential steps or phases |
| `icon-grid` | Grid of features or capabilities |
| `logos` | Customer logos for social proof |
| `checklist` | Items with checkbox status |
| `quote` | Testimonial with attribution |
| `closing` | Thank you slide with contact info |

---

## Customization

### Accent Text

Wrap words in asterisks to highlight them in the accent color:

```json
"title": "Governance of *the entire data estate* is hard"
```

### Speaker Notes

Add presenter notes to any slide:

```json
"notes": "Emphasize the multi-cloud aspect here"
```

### Theme

Colors, typography, and footer settings live in `themes/databricks.json`. Edit this file to adjust the visual styling.

---

## Manual Usage

Run the generator directly from the command line:

```bash
python3 ./scripts/generate-pptx.py \
  --input content.json \
  --output presentation.pptx
```

### Import to Google Slides

1. Go to [slides.google.com](https://slides.google.com)
2. File → Open → Upload
3. Select your .pptx file

---

## File Structure

```
├── SKILL.md              # Claude skill instructions
├── README.md             # This file
├── scripts/
│   └── generate-pptx.py  # Python generator
├── themes/
│   └── databricks.json   # Colors, fonts, footer config
├── assets/
│   └── databricks/       # Template and logo files
└── generated-slides/     # Output directory for presentations
```

## Requirements

- Python 3.8+
- python-pptx

## License

MIT License
