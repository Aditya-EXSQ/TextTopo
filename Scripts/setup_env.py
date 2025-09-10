#!/usr/bin/env python3
"""
Environment setup script for TextTopo development.
Centralizes Python cache files and sets up development environment.
"""

import os
import shutil
import sys
from pathlib import Path

# Get project root directory
PROJECT_ROOT = Path(__file__).parent.parent
PYCACHE_DIR = PROJECT_ROOT / ".dev" / "pycache"
TEMP_DIR = PROJECT_ROOT / "texttopo_temp"


def setup_pycache_centralization():
    """Set up centralized Python cache directory."""
    print("🔧 Setting up centralized Python cache...")
    
    # Create centralized cache directory
    PYCACHE_DIR.mkdir(parents=True, exist_ok=True)
    
    # Set environment variable for current session
    os.environ['PYTHONPYCACHEPREFIX'] = str(PYCACHE_DIR)
    
    print(f"✅ Python cache will be centralized to: {PYCACHE_DIR}")
    print("💡 To make this permanent, add this to your shell profile:")
    
    if os.name == 'nt':  # Windows
        print(f'   set PYTHONPYCACHEPREFIX={PYCACHE_DIR}')
        print("   Or add to your system environment variables")
    else:  # Unix/Linux/Mac
        print(f'   export PYTHONPYCACHEPREFIX={PYCACHE_DIR}')
        print("   Add this line to your ~/.bashrc or ~/.zshrc")


def clean_old_pycache():
    """Clean up old scattered .pycache directories."""
    print("🧹 Cleaning up old scattered .pycache directories...")
    
    pycache_dirs = list(PROJECT_ROOT.rglob("__pycache__"))
    
    if not pycache_dirs:
        print("✅ No scattered .pycache directories found")
        return
    
    print(f"Found {len(pycache_dirs)} .pycache directories to clean up:")
    for cache_dir in pycache_dirs:
        relative_path = cache_dir.relative_to(PROJECT_ROOT)
        print(f"   Removing: {relative_path}")
        try:
            shutil.rmtree(cache_dir)
        except Exception as e:
            print(f"   ⚠️  Failed to remove {relative_path}: {e}")
    
    print("✅ Cleanup complete!")


def ensure_gitignore():
    """Ensure .gitignore has proper entries."""
    print("📝 Updating .gitignore...")
    
    gitignore_path = PROJECT_ROOT / ".gitignore"
    
    # Entries to add
    entries = [
        "# Python cache files",
        "__pycache__/",
        "*.py[cod]",
        "*$py.class",
        "# Development cache directory", 
        ".dev/",
        "# TextTopo temp directory",
        "texttopo_temp/",
        "# Log files",
        "*.log",
    ]
    
    if gitignore_path.exists():
        with open(gitignore_path, 'r', encoding='utf-8') as f:
            current_content = f.read()
    else:
        current_content = ""
    
    # Add missing entries
    new_entries = []
    for entry in entries:
        if entry not in current_content and not entry.startswith('#'):
            new_entries.append(entry)
    
    if new_entries:
        with open(gitignore_path, 'a', encoding='utf-8') as f:
            f.write('\n\n# Added by TextTopo setup\n')
            f.write('\n'.join(entries))
            f.write('\n')
        print("✅ Updated .gitignore with cache and temp directories")
    else:
        print("✅ .gitignore already up to date")


def create_activation_script():
    """Create activation scripts for easy environment setup."""
    print("📜 Creating environment activation scripts...")
    
    scripts_dir = PROJECT_ROOT / "Scripts"
    
    # Windows batch script
    batch_content = f"""@echo off
echo Setting up TextTopo development environment...
set PYTHONPYCACHEPREFIX={PYCACHE_DIR}
echo [OK] Python cache centralized to: %PYTHONPYCACHEPREFIX%
echo [OK] Environment ready! Run your Python commands now.
cmd /k
"""
    
    batch_file = scripts_dir / "activate_env.bat"
    with open(batch_file, 'w') as f:
        f.write(batch_content)
    
    # Unix shell script  
    shell_content = f"""#!/bin/bash
echo "Setting up TextTopo development environment..."
export PYTHONPYCACHEPREFIX="{PYCACHE_DIR}"
echo "[OK] Python cache centralized to: $PYTHONPYCACHEPREFIX"
echo "[OK] Environment ready! Run your Python commands now."
exec "$SHELL"
"""
    
    shell_file = scripts_dir / "activate_env.sh"
    with open(shell_file, 'w') as f:
        f.write(shell_content)
    
    # Make shell script executable
    if os.name != 'nt':
        os.chmod(shell_file, 0o755)
    
    print("✅ Created activation scripts:")
    print(f"   Windows: {batch_file}")
    print(f"   Unix/Linux/Mac: {shell_file}")


def main():
    """Run all setup tasks."""
    print("🚀 TextTopo Development Environment Setup")
    print("=" * 50)
    
    setup_pycache_centralization()
    print()
    
    clean_old_pycache()
    print()
    
    ensure_gitignore()
    print()
    
    create_activation_script()
    print()
    
    print("🎉 Setup complete!")
    print()
    print("📋 Next steps:")
    print("1. Restart your terminal/IDE to pick up environment changes")
    print("2. Or run the activation script for immediate effect:")
    if os.name == 'nt':
        print("   Scripts\\activate_env.bat")
    else:
        print("   ./Scripts/activate_env.sh")
    print()
    print("💡 From now on, all .pycache files will be centralized!")


if __name__ == "__main__":
    main()
