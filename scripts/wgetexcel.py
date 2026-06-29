"""Compatibility wrapper for the row-range wget downloader workflow.

This script preserves the old ``scripts/wgetexcel.py`` entry point. The real
implementation lives in ``modules.gui_wget_range`` and should be edited there.
"""

from pathlib import Path
import sys

# Add the project root to Python's import path so this wrapper works when it is
# launched directly from the ``scripts`` directory.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from modules.gui_wget_range import launch


if __name__ == "__main__":
    # If this compatibility script is executed directly, launch the modular GUI.
    launch()
