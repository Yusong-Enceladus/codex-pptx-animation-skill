---
name: pptx-animation
description: Use when a PowerPoint .pptx deck needs element-by-element animations, click-by-click reveal sequences, grouped column/card/flow animations, animation QA, or animation XML/API troubleshooting. Especially relevant when adding entrance animations to existing slide elements while preserving layout.
---

# PPTX Animation

## Overview

Use this skill to add and validate PowerPoint element animations in existing `.pptx` files. Prefer Microsoft PowerPoint's native automation API when available; use raw OOXML only as a fallback or for inspection.

## Recommended Workflow

1. Start from a copy of the deck, never the only formal file.
2. Inspect candidate slides with `python-pptx` to list shape indices, names, positions, and text. Use `scripts/list_shapes.py` when a reusable inventory is helpful.
3. Define click groups by presentation logic:
   - Parallel columns/cards: one click per column/card, left-to-right and top-to-bottom.
   - Flow diagrams: one click per process step, then metrics or validation evidence.
   - Comparison pages: reveal baseline/problem first, then proposed method, then result.
   - Section/title/thank-you pages: usually no animation, or only a subtle title fade.
4. Prefer stable shape names in the spec. Numeric indices are acceptable for quick tests, but production decks should use unique PowerPoint shape names because AppleScript shape order can differ from `python-pptx` shape order.
5. Add entrance effects:
   - Use `animation type fade`.
   - First element in each click group uses `trigger on page click`.
   - Remaining elements in the same group use `trigger with previous`.
   - Keep duration around `0.35` to `0.45` seconds.
6. Validate:
   - Use `scripts/count_effects.py` or query PowerPoint directly with `count effects of main sequence of timeline of slide N`.
   - Run the deck QA script if one exists; animated slides should appear as `timing_slides`.
   - Export/render PDF or thumbnails to confirm static final state is visually unchanged.

## Batch Script

Use `scripts/add_click_animations.py` for deterministic batch work.

Input spec format:

```json
{
  "duration": 0.45,
  "slides": [
    {
      "slide": 5,
      "groups": [
        ["Title 1", "Card A", "Card A caption"],
        ["Card B", "Card B caption"]
      ]
    }
  ]
}
```

Each group can contain PowerPoint shape names or PowerPoint shape indices. Prefer shape names for production decks. Use numeric indices only after checking them in PowerPoint, not only in `python-pptx`.

Example:

```bash
python ~/.codex/skills/pptx-animation/scripts/add_click_animations.py \
  input.pptx output_animated.pptx spec.json --per-slide
```

`--per-slide` is slower, but is recommended for large decks because it opens, saves, and closes the deck once per slide. This avoids stale PowerPoint timeline/sequence objects during long AppleScript runs. It will close currently open PowerPoint presentations.

To generate a shape inventory:

```bash
python ~/.codex/skills/pptx-animation/scripts/list_shapes.py \
  input.pptx --slides 5 6 7 --output shape_inventory.csv
```

To count PowerPoint-recognized effects after insertion:

```bash
python ~/.codex/skills/pptx-animation/scripts/count_effects.py \
  output_animated.pptx --slides 5 6 7
```

## AppleScript Pattern

For one group:

```applescript
set seq to main sequence of timeline of s
set e1 to add effect seq for shape 3 of s fx animation type fade trigger on page click
set duration of timing of e1 to 0.45
set e2 to add effect seq for shape 4 of s fx animation type fade trigger with previous
set duration of timing of e2 to 0.45
```

This produces PowerPoint-recognized `<p:timing>` and `<p:animEffect>` entries. It is safer than hand-writing timing XML for broad deck edits.

## Troubleshooting Notes

- If AppleScript reports that `main sequence ... doesn't understand add effect`, first verify the shape reference. In practice this can mean the shape index/name is wrong, not that animation insertion is unsupported.
- If a script works on one slide but fails or hangs across many slides, rerun with `--per-slide`.
- If a slide has grouped or embedded artwork, animate the outer meaningful object unless the internal shapes are independently named and visually useful.
- After adding animations, export a PDF or contact sheet. PowerPoint animations should not change the static final slide state.

## Fallback OOXML Notes

PowerPoint animations live under `ppt/slides/slideN.xml` in `<p:timing>`. A fade entrance typically includes:

- `<p:seq ... nodeType="mainSeq">`
- click-triggered child nodes
- `<p:set>` for `style.visibility`
- `<p:animEffect transition="in" filter="fade">`
- target elements such as `<p:spTgt spid="..."/>`

Use OOXML fallback only when PowerPoint automation is unavailable, and always open the result in PowerPoint afterward.
