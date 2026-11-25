# -*- coding: utf-8 -*-
"""
provenance_check.py - is any of someone else's text living in this skill?

Attribution claims are worthless if they cannot be checked. This compares every
substantial line of this skill against another directory of skills and reports
verbatim matches. Zero matches is the evidence behind NOTICE.md's claim.

A match is not automatically a problem: a standard import line or a row of
dashes will collide by coincidence. The point is that the *count is small and
the matches are trivial*, and anyone can re-run it.

Usage:
    python scripts/provenance_check.py --against ~/.dsh/skills
    python scripts/provenance_check.py --against C:\\path\\to\\skills --min-len 60
"""
import argparse
import pathlib
import sys

EXTS = {".py", ".md", ".json", ".txt", ".yaml", ".yml"}
SKIP_DIRS = {"__pycache__", ".git", ".venv", "node_modules"}


def substantial_lines(root: pathlib.Path, exts, min_len: int) -> dict:
    """Map stripped line -> 'file:line' for every line worth comparing."""
    found = {}
    for path in root.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in exts:
            continue
        if any(part in SKIP_DIRS for part in path.parts):
            continue
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        for n, line in enumerate(text.splitlines(), 1):
            s = line.strip()
            if len(s) >= min_len:
                found.setdefault(s, f"{path.relative_to(root)}:{n}")
    return found


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--against", required=True,
                    help="directory of other skills, or one skill's folder")
    ap.add_argument("--self", dest="own", default=None,
                    help="this skill's root (default: the parent of scripts/)")
    ap.add_argument("--min-len", type=int, default=40)
    ap.add_argument("--show", type=int, default=4, help="examples per folder")
    args = ap.parse_args()

    here = pathlib.Path(args.own).resolve() if args.own else \
        pathlib.Path(__file__).resolve().parent.parent
    against = pathlib.Path(args.against).expanduser().resolve()
    if not against.exists():
        print(f"not found: {against}")
        return 2

    mine = substantial_lines(here, EXTS, args.min_len)
    print(f"this skill: {here}")
    print(f"comparable lines (>= {args.min_len} chars): {len(mine)}\n")

    targets = [against] if (against / "SKILL.md").exists() else \
        sorted(d for d in against.iterdir() if d.is_dir())

    total_hits = 0
    for target in targets:
        if target.resolve() == here:
            continue
        theirs = substantial_lines(target, EXTS, args.min_len)
        shared = sorted(set(mine) & set(theirs))
        total_hits += len(shared)
        label = target.name
        if not shared:
            print(f"{label:22} 0")
            continue
        print(f"{label:22} {len(shared)} shared line(s)")
        for s in shared[:args.show]:
            print(f"    mine  {mine[s]}")
            print(f"    their {theirs[s]}")
            print(f"    text  {s[:88]}")

    print(f"\ntotal shared lines: {total_hits}")
    print("A small number of trivial collisions (imports, dividers) is expected;"
          " anything longer deserves a look.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
