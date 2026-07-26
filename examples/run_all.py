"""Раннер демо: находит examples/*_demo.py, запускает параллельно, тихо."""

import os
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

DEMOS_DIR = Path(__file__).resolve().parent
ROOT = DEMOS_DIR.parent


def run(path: Path):
    env = {**os.environ, "PYTHONPATH": str(ROOT)}
    result = subprocess.run([sys.executable, str(path)], capture_output=True,
                             text=True, cwd=ROOT, env=env)
    return path, result


def main():
    demos = sorted(DEMOS_DIR.glob("*_demo.py"))
    with ThreadPoolExecutor(max_workers=len(demos) or 1) as pool:
        results = list(pool.map(run, demos))

    failed = False
    for path, result in results:
        if result.returncode != 0:
            failed = True
            print(f"=== FAILED: {path.name} ===")
            print(result.stdout)
            print(result.stderr)
    if not failed:
        print(f"run_all: {len(demos)} демо зелёные")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
