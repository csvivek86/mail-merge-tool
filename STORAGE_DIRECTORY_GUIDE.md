# Storage Directory Configuration Guide

## Overview

The NSNA Mail Merge Tool now allows users to customize where their templates, data files, and generated PDFs are stored. By default, the app automatically detects the best location (usually Desktop or Documents), but users can now specify their preferred directory.

## Features

### 1. **User Directory Selection**
- Users can specify a custom directory for storing NSNA Mail Merge files
- Settings are saved across sessions
- Easy reset to default auto-detection

### 2. **Template Priority System**
- **User Templates First**: Saved templates from the user's directory are loaded first
- **App Defaults Second**: If no user template exists, app defaults are used
- **Clear Indicators**: UI shows whether you're using a saved template or default

### 3. **Persistent Configuration**
- Custom storage paths are saved to `~/.nsna_mail_merge_settings.json`
- Settings persist across app sessions
- Easy to reset to defaults

## How to Use

### Configuring Storage Directory

1. **Access Configuration**:
   - In the main app, expand the "📁 Desktop Storage" section
   - Click "⚙️ Configure Storage Directory"

2. **Set Custom Directory**:
   - Enter the full path to your desired directory (e.g., `C:\MyNSNAFiles`)
   - Click "✅ Use This Directory" if the path is valid
   - The app will create a `NSNA_Mail_Merge` folder in your chosen location

3. **Reset to Default**:
   - Click "🔄 Reset to Default" to return to auto-detection
   - The app will find the best location automatically

### Directory Structure

Your chosen directory will contain:
```
Your-Chosen-Directory/
└── NSNA_Mail_Merge/
    ├── templates/          # Saved email and PDF templates
    ├── data/              # Uploaded Excel files
    ├── receipts/          # Generated PDF receipts
    └── exports/           # Other exported files
```

### Template Loading Behavior

#### Email Templates
- **"Saved Template"** option appears if you have a saved template
- Saved templates are loaded by default when available
- Clear indication: "📁 Using your saved template from user directory"

#### PDF Templates
- User's saved PDF template loads automatically when available
- Default template used if no saved template exists
- Clear indication of which template is being used

## Configuration File

Settings are stored in: `~/.nsna_mail_merge_settings.json`

Example:
```json
{
  "storage_path": "C:\\Users\\YourName\\MyNSNAFiles"
}
```

## Default Auto-Detection

If no custom path is configured, the app automatically detects:

### Windows:
1. Desktop folder
2. OneDrive Desktop (if available)
3. Documents folder (fallback)

### macOS/Linux:
1. Desktop folder
2. Documents folder (fallback)

### Docker:
- Uses mounted volume: `/app/desktop_storage`

## Validation

The app validates your custom directory by:
- Checking if the path exists
- Verifying it's a directory (not a file)
- Testing write permissions
- Creating a temporary file to ensure access

## Benefits

1. **Flexibility**: Store files where you prefer
2. **Organization**: Keep NSNA files in your organized folder structure
3. **Backup**: Easier to include in your backup routines
4. **Sharing**: Store on shared drives for team access
5. **Persistence**: Templates and settings survive app updates

## Tips

- Use full paths (e.g., `C:\MyFolder\NSNA` not `MyFolder`)
- Ensure the directory exists before configuring
- Test write permissions in your chosen directory
- Remember the location for backup purposes
- Use descriptive folder names for better organization

## Troubleshooting

### "Directory does not exist"
- Make sure the path is correct
- Create the directory first, then configure the app

### "Cannot write to this directory"
- Check folder permissions
- Try a different location you have write access to
- Run the app as administrator if needed (Windows)

### "Reset not working"
- Delete `~/.nsna_mail_merge_settings.json` manually
- Restart the app

### Templates not loading
- Check that template files exist in `your-path/NSNA_Mail_Merge/templates/`
- Verify the JSON files are not corrupted
- Use "Reset to Default" and reconfigure if needed
