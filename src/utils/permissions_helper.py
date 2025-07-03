import subprocess
import logging
from pathlib import Path

def grant_word_automation_access():
    """Grant automation access to Microsoft Word using AppleScript"""
    try:
        script = '''
        tell application "System Events"
            set wordAccess to UI elements enabled
            if not wordAccess then
                tell application "System Preferences"
                    activate
                    set current pane to pane id "com.apple.preference.security"
                    delay 1
                    tell application "System Events"
                        tell process "System Preferences"
                            click button "Privacy" of toolbar 1 of window 1
                            delay 1
                            select row 5 of table 1 of scroll area 1 of group 1 of window 1
                            delay 1
                            if not (checkbox 1 of row 1 of table 1 of scroll area 2 of group 1 of window 1's value as boolean) then
                                click checkbox 1 of row 1 of table 1 of scroll area 2 of group 1 of window 1
                            end if
                        end tell
                    end tell
                    quit
                end tell
            end if
        end tell
        '''
        
        subprocess.run(['osascript', '-e', script], check=True)
        logging.info("Successfully granted Word automation access")
        return True
    except Exception as e:
        logging.error(f"Failed to grant Word automation access: {str(e)}")
        return False