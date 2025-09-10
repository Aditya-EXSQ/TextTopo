#!/usr/bin/env python3
"""
Extract.py - Entry point script for TextTopo DOCX text extraction.

This is a thin wrapper around the CLI module that provides an easy way to run
TextTopo from the Scripts directory.
"""

import sys
import os

# Add the parent directory to the Python path so we can import DOCXToText
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, parent_dir)

# Import and run the CLI
from DOCXToText.CLI import main

if __name__ == "__main__":
    main()
