"""Compatibility wrapper for the snapshot Excel exporter workflow.

This script preserves the old ``scripts/linkstoexcel.py`` entry point. The real
implementation lives in ``modules.gui_snapshot_exporter`` and should be edited
there.
"""

from pathlib import Path
import sys

# Add the project root to Python's import path so this wrapper works when it is
# launched directly from the ``scripts`` directory.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from modules.gui_snapshot_exporter import launch


if __name__ == "__main__":
    # If this compatibility script is executed directly, launch the modular GUI.
    launch()
