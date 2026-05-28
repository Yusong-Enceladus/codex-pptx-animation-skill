#!/usr/bin/env python3
"""Add grouped click-by-click fade animations to a PowerPoint deck.

Requires Microsoft PowerPoint on macOS. Shape references can be PowerPoint shape
names or indices within each slide. Prefer unique shape names for production
decks because PowerPoint's AppleScript shape order can differ from python-pptx.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import time
from pathlib import Path


def quote_applescript_string(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def shape_ref(value: int | str) -> str:
    if isinstance(value, int):
        return f"shape {value}"
    if isinstance(value, str) and value.isdigit():
        return f"shape {int(value)}"
    return f'shape "{quote_applescript_string(str(value))}"'


def build_applescript(pptx_path: Path, spec: dict, open_delay: float) -> str:
    duration = float(spec.get("duration", 0.45))
    lines = [
        f'set pptFile to POSIX file "{quote_applescript_string(str(pptx_path))}"',
        'tell application "Microsoft PowerPoint"',
        "    open pptFile",
        f"    delay {open_delay:.2f}",
        "    set pres to active presentation",
    ]

    for slide_spec in spec.get("slides", []):
        slide_no = int(slide_spec["slide"])
        groups = slide_spec.get("groups", [])
        lines.append(f"    set s to slide {slide_no} of pres")
        lines.append("    set seq to main sequence of timeline of s")
        for group in groups:
            if not group:
                continue
            first = True
            for shape_index in group:
                trigger = "on page click" if first else "with previous"
                lines.append(
                    f"    set fxObj to add effect seq for {shape_ref(shape_index)} of s "
                    f"fx animation type fade trigger {trigger}"
                )
                lines.append(f"    set duration of timing of fxObj to {duration:.2f}")
                first = False

    lines.extend(
        [
            "    save pres",
            "    close pres saving no",
            "end tell",
        ]
    )
    return "\n".join(lines) + "\n"


def close_all_presentations() -> None:
    script = """
tell application "Microsoft PowerPoint"
    repeat while (count presentations) > 0
        close active presentation saving no
    end repeat
end tell
"""
    subprocess.run(["osascript", "-e", script], text=True, capture_output=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_pptx", type=Path)
    parser.add_argument("output_pptx", type=Path)
    parser.add_argument("spec_json", type=Path)
    parser.add_argument("--keep-applescript", type=Path, help="Write generated AppleScript to this path")
    parser.add_argument(
        "--per-slide",
        action="store_true",
        help="Open/save/close once per slide. Slower, but avoids PowerPoint sequence-object instability in large batches.",
    )
    parser.add_argument(
        "--open-delay",
        type=float,
        default=1.0,
        help="Seconds to wait after PowerPoint opens the deck before editing.",
    )
    parser.add_argument(
        "--between-slide-delay",
        type=float,
        default=1.5,
        help="Seconds to wait between per-slide runs after closing PowerPoint presentations.",
    )
    args = parser.parse_args()

    if args.input_pptx.resolve() != args.output_pptx.resolve():
        args.output_pptx.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(args.input_pptx, args.output_pptx)

    spec = json.loads(args.spec_json.read_text(encoding="utf-8"))
    if args.per_slide:
        scripts = []
        for slide_spec in spec.get("slides", []):
            one_slide_spec = {**spec, "slides": [slide_spec]}
            scripts.append(
                (
                    int(slide_spec["slide"]),
                    build_applescript(args.output_pptx.resolve(), one_slide_spec, args.open_delay),
                )
            )
        if args.keep_applescript:
            args.keep_applescript.parent.mkdir(parents=True, exist_ok=True)
            args.keep_applescript.write_text(
                "\n\n".join(f"-- slide {slide_no}\n{script}" for slide_no, script in scripts),
                encoding="utf-8",
            )
        for slide_no, script in scripts:
            close_all_presentations()
            time.sleep(0.5)
            result = subprocess.run(["osascript", "-e", script], text=True, capture_output=True)
            if result.returncode:
                raise SystemExit(
                    result.stderr or result.stdout or f"osascript failed on slide {slide_no} with code {result.returncode}"
                )
            print(f"Animated slide {slide_no}", flush=True)
            close_all_presentations()
            time.sleep(args.between_slide_delay)
    else:
        script = build_applescript(args.output_pptx.resolve(), spec, args.open_delay)
        if args.keep_applescript:
            args.keep_applescript.parent.mkdir(parents=True, exist_ok=True)
            args.keep_applescript.write_text(script, encoding="utf-8")

        result = subprocess.run(["osascript", "-e", script], text=True, capture_output=True)
        if result.returncode:
            raise SystemExit(result.stderr or result.stdout or f"osascript failed with code {result.returncode}")

    print(f"Wrote {args.output_pptx}", flush=True)


if __name__ == "__main__":
    main()
