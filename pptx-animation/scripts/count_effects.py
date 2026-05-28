#!/usr/bin/env python3
"""Count PowerPoint-recognized animation effects per slide on macOS."""

from __future__ import annotations

import argparse
import subprocess
from pathlib import Path


def quote_applescript_string(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def build_script(pptx_path: Path, slides: list[int], open_delay: float) -> str:
    lines = [
        f'set pptFile to POSIX file "{quote_applescript_string(str(pptx_path))}"',
        'tell application "Microsoft PowerPoint"',
        "    open pptFile",
        f"    delay {open_delay:.2f}",
        "    set pres to active presentation",
        "    set outputLines to {\"slide,effects\"}",
    ]

    if slides:
        slide_list = ", ".join(str(slide_no) for slide_no in slides)
        lines.extend(
            [
                f"    repeat with slideNo in {{{slide_list}}}",
                "        set effectCount to count effects of main sequence of timeline of slide slideNo of pres",
                '        set end of outputLines to ((slideNo as text) & "," & (effectCount as text))',
                "    end repeat",
            ]
        )
    else:
        lines.extend(
            [
                "    repeat with slideNo from 1 to (count slides of pres)",
                "        set effectCount to count effects of main sequence of timeline of slide slideNo of pres",
                '        set end of outputLines to ((slideNo as text) & "," & (effectCount as text))',
                "    end repeat",
            ]
        )

    lines.extend(
        [
            "    close pres saving no",
            "end tell",
            'set AppleScript\'s text item delimiters to linefeed',
            "set outputText to outputLines as text",
            'set AppleScript\'s text item delimiters to ""',
            "return outputText",
        ]
    )
    return "\n".join(lines) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--slides", nargs="*", type=int, help="Optional 1-based slide numbers to inspect")
    parser.add_argument("--open-delay", type=float, default=1.0)
    args = parser.parse_args()

    script = build_script(args.pptx.resolve(), args.slides or [], args.open_delay)
    result = subprocess.run(["osascript", "-e", script], text=True, capture_output=True)
    if result.returncode:
        raise SystemExit(result.stderr or result.stdout or f"osascript failed with code {result.returncode}")
    print(result.stdout.strip())


if __name__ == "__main__":
    main()
