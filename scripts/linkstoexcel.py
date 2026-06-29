from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from modules.gui_snapshot_exporter import launch


if __name__ == "__main__":
    launch()
