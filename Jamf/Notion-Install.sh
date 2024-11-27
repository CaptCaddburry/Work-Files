#!/bin/bash

# Location for Waiting Room and Installer
WAITING_ROOM="/Library/Application Support/JAMF/Waiting Room"
INSTALLER_FILE="Notion-3.10.0-universal.dmg"

# Mount Notion Image
hdiutil attach -nobrowse "$WAITING_ROOM/$INSTALLER_FILE"

# Copy over App to Applications folder
cp -r /Volumes/Notion/Notion.app /Applications/

# Unmount and Delete Notion Image
hdiutil detach /Volumes/Notion/
rm "$WAITING_ROOM/$INSTALLER_FILE" "$WAITING_ROOM/$INSTALLER_FILE.cache.xml"
