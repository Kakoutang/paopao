#!/usr/bin/env python3
"""Lab/advanced Paopao CLI.

The production entrypoint is paopao_run.py, whose help intentionally exposes
only the HTML-source-only factory path. This wrapper exposes optional,
experimental, and diagnostic commands without putting them on the default menu.
"""

from __future__ import annotations

import paopao_run


if __name__ == "__main__":
    raise SystemExit(paopao_run.main(include_lab=True))
