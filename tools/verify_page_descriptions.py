"""Utility to validate Lumina layout banner descriptions.

Run with `python tools/verify_page_descriptions.py` from the repository root
to ensure every HTML file that includes the shared `layout` template
resolves to a non-empty description via either inline metadata or the
central lookup table.
"""
from __future__ import annotations

import json
import os
import re
import sys
from typing import Dict, Iterable, List, Tuple

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LAYOUT_PATH = os.path.join(ROOT, "layout.html")


def slugify(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]+", "-", value.strip().lower())
    value = re.sub(r"-+", "-", value).strip("-")
    return value


def read_layout_lookup() -> Dict[str, str]:
    with open(LAYOUT_PATH, "r", encoding="utf-8", errors="ignore") as handle:
        layout_text = handle.read()

    entries_match = re.search(
        r"var __layoutPageDescriptionEntries = \{([\s\S]*?)\};",
        layout_text,
    )
    if not entries_match:
        raise RuntimeError("Could not locate __layoutPageDescriptionEntries in layout.html")

    entries = dict(re.findall(r"'([^']+)':\s*'([^']+)'", entries_match.group(1)))
    lookup: Dict[str, str] = {}
    for key, description in entries.items():
        if not description:
            continue
        lookup[key] = description
        collapsed = key.replace("-", "")
        if collapsed and collapsed != key and collapsed not in entries and collapsed not in lookup:
            lookup[collapsed] = description
    return lookup


def iter_layout_pages(root: str) -> Iterable[Tuple[str, str, bool]]:
    include_pattern = re.compile(r"include\((?:'|\")layout(?:'|\")\s*,\s*\{([\s\S]*?)\}\)")
    for current_root, _, files in os.walk(root):
        for filename in files:
            if not filename.endswith(".html"):
                continue
            path = os.path.join(current_root, filename)
            with open(path, "r", encoding="utf-8", errors="ignore") as handle:
                text = handle.read()
            match = include_pattern.search(text)
            if not match:
                continue
            block = match.group(1)
            has_inline_description = "pageDescription" in block
            props = dict(re.findall(r"(\w+)\s*:\s*([^,\n]+)", block))
            slug_candidate = ""
            for key in ("pageSlug", "currentPage"):
                if key in props:
                    value = props[key]
                    if "||" in value:
                        value = value.split("||")[1]
                    value = re.sub(r"['\"()]", "", value).strip()
                    if value:
                        slug_candidate = value
                        break
            if not slug_candidate:
                slug_candidate = os.path.splitext(filename)[0]
            yield path, slugify(slug_candidate), has_inline_description


def main() -> int:
    lookup = read_layout_lookup()
    missing: List[Tuple[str, str]] = []
    coverage: List[Tuple[str, str]] = []

    for path, slug, has_inline in iter_layout_pages(ROOT):
        if has_inline:
            coverage.append((path, slug))
            continue
        if slug not in lookup:
            missing.append((path, slug))
        else:
            coverage.append((path, slug))

    report = {
        "total_pages": len(coverage),
        "missing_descriptions": missing,
    }

    print(json.dumps(report, indent=2, sort_keys=True))

    if missing:
        for path, slug in missing:
            print(f"Missing description for slug '{slug}' from {path}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
