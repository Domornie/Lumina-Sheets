#!/usr/bin/env python3
"""
Utility to scan the repository for a given token within text files.

Example:
    python tools/find_token.py --token token --max-results 20
"""

import argparse
from pathlib import Path
from typing import Iterable, List, Tuple

DEFAULT_IGNORES = {'.git', '.venv', 'node_modules', '__pycache__'}


def should_skip(path: Path) -> bool:
    return any(part in DEFAULT_IGNORES for part in path.parts)


def is_text_file(path: Path) -> bool:
    try:
        with path.open('rb') as handle:
            chunk = handle.read(1024)
        return b'\0' not in chunk
    except OSError:
        return False


def find_token(root: Path, token: str, *, case_sensitive: bool = False) -> List[Tuple[Path, int, str]]:
    matches: List[Tuple[Path, int, str]] = []
    token_match = token if case_sensitive else token.lower()

    for path in root.rglob('*'):
        if not path.is_file() or should_skip(path):
            continue
        if not is_text_file(path):
            continue

        try:
            content = path.read_text(encoding='utf-8', errors='ignore')
        except OSError:
            continue

        for line_number, line in enumerate(content.splitlines(), start=1):
            haystack = line if case_sensitive else line.lower()
            if token_match in haystack:
                matches.append((path.relative_to(root), line_number, line.strip()))

    return matches


def format_matches(matches: Iterable[Tuple[Path, int, str]]) -> str:
    lines: List[str] = []
    for path, line_number, line in matches:
        snippet = line if len(line) <= 200 else f"{line[:197]}..."
        lines.append(f"{path}:{line_number}: {snippet}")
    return '\n'.join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description='Scan repository files for a token.')
    parser.add_argument('--token', default='token', help='Text to search for (default: token).')
    parser.add_argument('--case-sensitive', action='store_true', help='Match token with case sensitivity.')
    parser.add_argument('--max-results', type=int, default=None, help='Limit the number of reported matches.')

    args = parser.parse_args()
    repo_root = Path(__file__).resolve().parent.parent

    matches = find_token(repo_root, args.token, case_sensitive=args.case_sensitive)
    if args.max_results is not None:
        matches = matches[: args.max_results]

    if not matches:
        print(f"No matches found for '{args.token}'.")
        return

    print(format_matches(matches))


if __name__ == '__main__':
    main()
