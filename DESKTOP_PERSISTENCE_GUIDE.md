# NSNA Mail Merge Tool - Desktop Persistent Storage Guide

## Overview
This version of the NSNA Mail Merge Tool is designed to store all user data, templates, and files directly on the user's desktop for easy access and persistence across Docker container restarts.

## Storage Structure
When the application runs, it creates a folder structure on your desktop:

```
Desktop/
└── NSNA_Mail_Merge/
    ├── data/           # Uploaded Excel files
    ├── templates/      # Saved email and PDF templates
    ├── receipts/       # Generated PDF receipts
    └── exports/        # Exported data files
```

## Docker Deployment with Desktop Storage

### Prerequisites
- Docker Desktop installed
- Git (to clone the repository)

### Quick Start

1. **Clone the repository:**
   ```bash
   git clone <repository-url>
   cd mail-merge-tool-docker
   ```

2. **Build and run with Docker Compose:**
   ```bash
   # For Windows
   docker-compose up --build

   # For Linux/macOS (edit docker-compose.yml first)
   docker-compose up --build
   ```

3. **Access the application:**
   Open your browser to `http://localhost:8501`

### Platform-Specific Setup

#### Windows
The docker-compose.yml file is pre-configured for Windows:
```yaml
volumes:
  - ${USERPROFILE}/Desktop:/app/desktop_storage
```

#### Linux/macOS
Edit `docker-compose.yml` and replace the volumes section:
```yaml
volumes:
  - ${HOME}/Desktop:/app/desktop_storage
```

### Development Mode
For development with hot reload:
```bash
docker-compose --profile dev up --build
```
This runs on port 8502 and includes source code mounting.

## Features

### Desktop File Persistence
- **Automatic Storage**: All uploaded Excel files are automatically saved to `Desktop/NSNA_Mail_Merge/data/`
- **Template Persistence**: Email and PDF templates are saved to `Desktop/NSNA_Mail_Merge/templates/`
- **Previously Saved Files**: The app shows a dropdown of previously uploaded Excel files
- **Easy Access**: All files are accessible directly from your desktop folder

### Template Management
- **Email Templates**: Saved as JSON files with rich text formatting
- **PDF Templates**: Saved as JSON files with letterhead integration
- **User Settings**: Email settings and preferences are automatically saved

### Generated Receipts
- PDF receipts are saved to `Desktop/NSNA_Mail_Merge/receipts/`
- Files are automatically named with timestamp and recipient information

## Usage

### First Time Setup
1. Start the application using Docker
2. Check your desktop - a new `NSNA_Mail_Merge` folder will appear
3. Upload your first Excel file - it will be saved automatically
4. Create your email and PDF templates - they will be saved automatically

### Subsequent Uses
1. Start the application
2. Your previous files and templates will be automatically loaded
3. Select from previously uploaded Excel files using the dropdown
4. Your saved templates will be pre-loaded in the editors

### File Management
- **Manual Access**: Navigate to `Desktop/NSNA_Mail_Merge/` to access all files
- **Backup**: Copy the entire `NSNA_Mail_Merge` folder to backup all data
- **Migration**: Move the folder to another computer's desktop to transfer data

## Configuration

### Environment Variables
- `DOCKER_DEPLOYMENT=true`: Automatically set in Docker containers
- This tells the app to use the mounted volume instead of trying to access the host desktop directly

### Volume Mounting
The key to desktop persistence is the volume mount in docker-compose.yml:
```yaml
volumes:
  - ${USERPROFILE}/Desktop:/app/desktop_storage  # Windows
  # - ${HOME}/Desktop:/app/desktop_storage        # Linux/macOS
```

## Troubleshooting

### Common Issues

1. **Desktop folder not created**
   - Check Docker Desktop permissions
   - Ensure the volume mount path is correct for your OS

2. **Files not persisting**
   - Verify the volume mount in docker-compose.yml
   - Check that the DOCKER_DEPLOYMENT environment variable is set

3. **Permission issues (Linux/macOS)**
   - Ensure Docker has permission to access the Desktop folder
   - May need to adjust file permissions: `chmod 755 ~/Desktop`

### Verification Steps

1. **Check volume mount:**
   ```bash
   docker exec -it nsna-mail-merge-tool ls -la /app/desktop_storage
   ```

2. **Check environment:**
   ```bash
   docker exec -it nsna-mail-merge-tool env | grep DOCKER_DEPLOYMENT
   ```

3. **Check desktop folder:**
   Look for `NSNA_Mail_Merge` folder on your desktop

## Benefits

### User Experience
- **No data loss**: Files persist across container restarts
- **Easy access**: All files are on the desktop, no need to search containers
- **Familiar location**: Desktop is a familiar location for most users
- **Manual backup**: Users can easily backup their data by copying the folder

### Administrative
- **Portable**: The entire setup can be moved between computers
- **Version control**: Templates and settings are in readable JSON format
- **Debugging**: Easy to inspect saved files and troubleshoot issues
- **Migration**: Simple to migrate data between environments

## Security Considerations

### File Permissions
- The application only has access to the mounted desktop directory
- Files are created with standard user permissions
- No system-wide access is granted

### Data Location
- All sensitive data remains on the user's local machine
- No cloud storage or external services for user data
- Templates and settings are stored locally in JSON format

## Customization

### Changing Storage Location
To use a different storage location, modify the volume mount in docker-compose.yml:
```yaml
volumes:
  - /path/to/your/preferred/location:/app/desktop_storage
```

### Adding Additional Mounts
You can mount additional directories if needed:
```yaml
volumes:
  - ${USERPROFILE}/Desktop:/app/desktop_storage
  - ${USERPROFILE}/Documents/NSNA_Backup:/app/backup_storage
```
