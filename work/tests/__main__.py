#!/usr/bin/env python3
"""
Run every test module in this directory, each in its own
process, and exit non-zero if any test failed.

Usage:
    python3 -m tests          from work/

Created: 2026-09-14
Last updated: 2026-09-14
"""

import glob
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


# --------------------------------------------------------------
def modules():
    """Every test module, as a dotted name."""
    names = sorted(glob.glob(os.path.join(HERE, "test_*.py")))
    return ["tests." + os.path.basename(n)[:-3] for n in names]


# --------------------------------------------------------------
def run_one(module):
    """Run one module, print its summary, return failures."""
    done = subprocess.run([sys.executable, "-m", module],
                          cwd=ROOT, capture_output=True,
                          text=True)
    lines = done.stdout.strip().split("\n")
    for line in lines:
        if "FAIL" in line:
            print(f"  {line.strip()}")
    summary = lines[-1] if lines else "no output"
    print(f"{module:24s} {summary}")
    if done.returncode and done.stderr:
        print(done.stderr[-600:])
    return done.returncode != 0


# --------------------------------------------------------------
def main():
    """Run all test modules and report."""
    failed = sum(run_one(m) for m in modules())
    sys.exit(1 if failed else 0)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
