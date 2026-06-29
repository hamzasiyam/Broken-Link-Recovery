"""Entry point for the Broken Link Recovery Tool desktop application.

This module intentionally stays small. It imports the launcher-level
``main`` function and delegates all command-line parsing and GUI startup
to ``modules.launcher``.
"""

from modules.launcher import main


if __name__ == "__main__":
    # If this file is executed directly, hand control to the launcher and
    # return its exit code to the operating system.
    raise SystemExit(main())
