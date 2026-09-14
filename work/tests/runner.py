#!/usr/bin/env python3
"""
Run the plain assert-style tests in a module.

The test files here need no pytest: each test is a function
whose name starts with "test_" and asserts do the checking.
This collects them, runs them, and reports, so each test
file does not carry its own copy of the same twenty lines.

Usage:
    from tests.runner import run
    if __name__ == "__main__":
        run(globals())

Created: 2026-09-12
Last updated: 2026-09-12
"""

import sys


# --------------------------------------------------------------
def collect(namespace):
    """Find every test function in a module namespace."""
    names = sorted(
        n for n in namespace if n.startswith("test_")
    )
    return [(n, namespace[n]) for n in names]


# --------------------------------------------------------------
def run(namespace):
    """Run every test and exit non-zero on any failure."""
    passed = failed = 0
    for name, func in collect(namespace):
        try:
            func()
            passed += 1
            print(f"  ok    {name}")
        except Exception as exc:
            failed += 1
            print(f"  FAIL  {name}: {exc}")

    print("-" * 50)
    print(f"passed: {passed}, failed: {failed}")
    sys.exit(1 if failed else 0)
