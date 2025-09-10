#!/usr/bin/env python3
"""
Cleanup utility for TextTopo project.
Removes temporary files, cache files, and build artifacts.
"""

import os
import shutil
import sys
from pathlib import Path

# Get project root directory
PROJECT_ROOT = Path(__file__).parent.parent


def cleanup_pycache():
    """Remove all __pycache__ directories."""
    print("🧹 Cleaning up __pycache__ directories...")
    
    pycache_dirs = list(PROJECT_ROOT.rglob("__pycache__"))
    
    if not pycache_dirs:
        print("✅ No __pycache__ directories found")
        return
    
    for cache_dir in pycache_dirs:
        relative_path = cache_dir.relative_to(PROJECT_ROOT)
        print(f"   Removing: {relative_path}")
        try:
            shutil.rmtree(cache_dir)
        except Exception as e:
            print(f"   ⚠️  Failed to remove {relative_path}: {e}")
    
    print(f"✅ Cleaned up {len(pycache_dirs)} __pycache__ directories")


def cleanup_temp_dirs():
    """Remove temporary directories."""
    print("🗂️  Cleaning up temporary directories...")
    
    temp_patterns = [
        "texttopo_temp",
        ".dev/pycache",
        "**/temp_*",
        "**/tmp_*"
    ]
    
    removed_count = 0
    
    for pattern in temp_patterns:
        for temp_dir in PROJECT_ROOT.glob(pattern):
            if temp_dir.is_dir():
                relative_path = temp_dir.relative_to(PROJECT_ROOT)
                print(f"   Removing: {relative_path}")
                try:
                    shutil.rmtree(temp_dir)
                    removed_count += 1
                except Exception as e:
                    print(f"   ⚠️  Failed to remove {relative_path}: {e}")
    
    if removed_count == 0:
        print("✅ No temporary directories found")
    else:
        print(f"✅ Cleaned up {removed_count} temporary directories")


def cleanup_log_files():
    """Remove log files."""
    print("📄 Cleaning up log files...")
    
    log_files = list(PROJECT_ROOT.rglob("*.log"))
    
    if not log_files:
        print("✅ No log files found")
        return
    
    for log_file in log_files:
        relative_path = log_file.relative_to(PROJECT_ROOT)
        print(f"   Removing: {relative_path}")
        try:
            log_file.unlink()
        except Exception as e:
            print(f"   ⚠️  Failed to remove {relative_path}: {e}")
    
    print(f"✅ Cleaned up {len(log_files)} log files")


def cleanup_build_artifacts():
    """Remove build artifacts and distribution files."""
    print("🔨 Cleaning up build artifacts...")
    
    artifacts = [
        "build",
        "dist", 
        "*.egg-info",
        ".eggs"
    ]
    
    removed_count = 0
    
    for pattern in artifacts:
        for artifact in PROJECT_ROOT.glob(pattern):
            relative_path = artifact.relative_to(PROJECT_ROOT)
            print(f"   Removing: {relative_path}")
            try:
                if artifact.is_dir():
                    shutil.rmtree(artifact)
                else:
                    artifact.unlink()
                removed_count += 1
            except Exception as e:
                print(f"   ⚠️  Failed to remove {relative_path}: {e}")
    
    if removed_count == 0:
        print("✅ No build artifacts found")
    else:
        print(f"✅ Cleaned up {removed_count} build artifacts")


def get_directory_size(path):
    """Get the total size of a directory and its contents."""
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            try:
                total_size += os.path.getsize(file_path)
            except (OSError, FileNotFoundError):
                pass
    return total_size


def main():
    """Run cleanup operations."""
    print("🧹 TextTopo Project Cleanup")
    print("=" * 30)
    
    # Calculate space before cleanup
    print("📊 Calculating current disk usage...")
    initial_size = get_directory_size(PROJECT_ROOT)
    
    cleanup_pycache()
    print()
    
    cleanup_temp_dirs() 
    print()
    
    cleanup_log_files()
    print()
    
    cleanup_build_artifacts()
    print()
    
    # Calculate space after cleanup
    final_size = get_directory_size(PROJECT_ROOT)
    space_saved = initial_size - final_size
    
    if space_saved > 0:
        if space_saved > 1024 * 1024:
            space_saved_mb = space_saved / (1024 * 1024)
            print(f"💾 Space saved: {space_saved_mb:.2f} MB")
        elif space_saved > 1024:
            space_saved_kb = space_saved / 1024
            print(f"💾 Space saved: {space_saved_kb:.2f} KB")
        else:
            print(f"💾 Space saved: {space_saved} bytes")
    else:
        print("💾 No significant space saved")
    
    print()
    print("✨ Cleanup complete!")


if __name__ == "__main__":
    main()
