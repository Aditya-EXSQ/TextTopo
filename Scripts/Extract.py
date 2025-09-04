import os
import sys

# Ensure project root (parent of Scripts/) is on sys.path so DOCXToText is importable
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
	sys.path.insert(0, PROJECT_ROOT)

from DOCXToText.CLI import main


if __name__ == "__main__":
	sys.exit(main())


