"""
Cleanup script to remove unnecessary documentation files created during development
"""

import os
from pathlib import Path

# Get the project directory
project_dir = Path(__file__).parent

print("📚 CLEANING UP UNNECESSARY DOCUMENTATION FILES")
print("=" * 50)

# Essential documentation to KEEP
essential_docs = {
    "README.md",  # Main project documentation
    "DEPLOYMENT_GUIDE.md",  # Important for deployment
    "DESKTOP_PERSISTENCE_GUIDE.md",  # User guide
    "STORAGE_DIRECTORY_GUIDE.md",  # User guide
    "VARIABLE_INSERTION_GUIDE.md",  # User guide
    "USER_GUIDE_RICH_TEXT_PERSISTENCE.md"  # User guide
}

# Documentation files to REMOVE (development artifacts)
docs_to_remove = [
    # Development summaries - no longer needed
    "BUILD_SUMMARY.md",
    "CLEANUP_SUMMARY.md", 
    "CLEANUP_TEST_FILES_SUMMARY.md",
    "ENHANCED_EDITOR_CLEANUP_SUMMARY.md",
    "IMPLEMENTATION_STATUS_FINAL.md",
    "CONVERSION_COMPLETE.md",
    "CONVERSION_COMPLETE_FINAL.md",
    "FINAL_DEPLOYMENT_RECOMMENDATION.md",
    
    # Editor evolution docs - now obsolete
    "STREAMLIT_LEXICAL_IMPORT_FIX.md",
    "STREAMLIT_LEXICAL_INTEGRATION.md", 
    "STREAMLIT_QUILL_INTEGRATION.md",
    "ST_QUILL_PARAMETER_FIXES.md",
    "ST_TINY_EDITOR_IMPORT_FIX.md",
    "ST_TINY_EDITOR_INTEGRATION.md",
    "TINYMCE_API_KEY_INTEGRATION.md",
    "RICH_TEXT_EDITOR_OVERHAUL.md",
    "RICH_TEXT_INTEGRATION_SUMMARY.md",
    "RICH_TEXT_AND_PERSISTENCE_INTEGRATION.md",
    "COMPLETE_RICH_TEXT_INTEGRATION.md",
    "WORD_LIKE_EDITOR_COMPLETE.md",
    "WORD_LIKE_QUILL_EDITOR.md",
    "COMPACT_WORD_LIKE_EDITOR.md",
    
    # Technical fix summaries - development artifacts
    "HTML_FORMATTING_FIX_SUMMARY.md",
    "PDF_LETTERHEAD_OVERLAY_FIX.md", 
    "PDF_POSITIONING_AND_FORMATTING_FIX.md",
    "WEASYPRINT_REMOVAL_SUMMARY.md",
    "PYQT6_CLEANUP_SUMMARY.md",
    "NAVIGATION_CONSOLIDATION_SUMMARY.md",
    "OAUTH_NAVIGATION_FIXES.md",
    "STORAGE_SELECTION_SUMMARY.md",
    
    # Feature development docs - now integrated
    "SAMPLE_PDF_GENERATION_FEATURE.md",
    "NO_REGEX_NEEDED_EXPLANATION.md",
    
    # Duplicate/obsolete deployment docs
    "STREAMLIT_CLOUD_DEPLOYMENT.md",  # Covered in main deployment guide
    "STREAMLIT_README.md",  # Duplicate of main README
    
    # This cleanup script
    "cleanup_docs.py"
]

print("📋 ESSENTIAL DOCUMENTATION (keeping):")
for doc in sorted(essential_docs):
    if (project_dir / doc).exists():
        print(f"✅ Keep: {doc}")

print(f"\n🗑️  REMOVING DEVELOPMENT ARTIFACTS:")

removed_count = 0
not_found_count = 0

for filename in docs_to_remove:
    file_path = project_dir / filename
    if file_path.exists():
        try:
            file_path.unlink()
            print(f"✅ Removed: {filename}")
            removed_count += 1
        except Exception as e:
            print(f"❌ Failed to remove {filename}: {e}")
    else:
        print(f"ℹ️  Not found: {filename}")
        not_found_count += 1

print(f"\n📊 DOCUMENTATION CLEANUP SUMMARY:")
print(f"✅ Files removed: {removed_count}")
print(f"ℹ️  Files not found: {not_found_count}")
print(f"📁 Total processed: {len(docs_to_remove)}")
print(f"📚 Essential docs kept: {len(essential_docs)}")

print(f"\n🎯 REMAINING DOCUMENTATION:")
print("✅ README.md - Main project documentation")
print("✅ DEPLOYMENT_GUIDE.md - Deployment instructions")
print("✅ DESKTOP_PERSISTENCE_GUIDE.md - User guide for desktop mode")
print("✅ STORAGE_DIRECTORY_GUIDE.md - Storage configuration guide")
print("✅ VARIABLE_INSERTION_GUIDE.md - Template variable guide")
print("✅ USER_GUIDE_RICH_TEXT_PERSISTENCE.md - Rich text editor guide")

print(f"\n🚀 DOCUMENTATION CLEANUP COMPLETE!")
print("Only essential user and deployment guides remain.")
