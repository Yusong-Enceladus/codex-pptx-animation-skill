# PPTX Animation Codex Skill

Codex skill for adding click-by-click PowerPoint element animations to existing `.pptx` decks while preserving the static slide layout.

## What It Does

- Adds grouped fade entrance animations through Microsoft PowerPoint's native macOS AppleScript API.
- Supports one-click-per-column, one-click-per-card, and step-by-step flow reveals.
- Supports shape references by name or index, with shape names recommended for production decks.
- Includes helper scripts to list slide shapes and count PowerPoint-recognized animation effects.

## Requirements

- macOS
- Microsoft PowerPoint installed locally
- Python 3
- `python-pptx` for `list_shapes.py`

Install the optional Python dependency:

```bash
python -m pip install -r requirements.txt
```

## Install The Skill

Clone or download this repository, then copy the skill folder into your Codex skills directory:

```bash
mkdir -p ~/.codex/skills
cp -R pptx-animation ~/.codex/skills/
```

Restart Codex if your environment does not hot-reload skills.

## Typical Workflow

1. Generate a shape inventory:

```bash
python ~/.codex/skills/pptx-animation/scripts/list_shapes.py \
  input.pptx --slides 5 6 7 --output shape_inventory.csv
```

2. Create an animation spec using stable PowerPoint shape names:

```json
{
  "duration": 0.45,
  "slides": [
    {
      "slide": 5,
      "groups": [
        ["Column A Title", "Column A Body"],
        ["Column B Title", "Column B Body"]
      ]
    }
  ]
}
```

3. Add animations:

```bash
python ~/.codex/skills/pptx-animation/scripts/add_click_animations.py \
  input.pptx output_animated.pptx spec.json --per-slide
```

4. Validate effect counts:

```bash
python ~/.codex/skills/pptx-animation/scripts/count_effects.py \
  output_animated.pptx --slides 5 6 7
```

## Practical Notes

- Prefer shape names over numeric indices. `python-pptx` shape order can differ from PowerPoint's AppleScript shape order on some slides.
- Use `--per-slide` for large decks. It is slower, but avoids stale PowerPoint timeline objects during long AppleScript runs.
- `--per-slide` closes currently open PowerPoint presentations before and after each slide run.
- If AppleScript reports that `main sequence ... doesn't understand add effect`, first verify the shape reference.
- Always export or render the deck afterward to confirm the final static slide appearance is unchanged.

## License

MIT
