#!/bin/bash
# Install or uninstall the launchd agent for automatic sync scheduling.
#
# Usage:
#   ./scripts/install_launchd.sh install
#   ./scripts/install_launchd.sh uninstall

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
PLIST_NAME="com.user.outlook-gcal-sync.plist"
PLIST_SRC="$PROJECT_DIR/config/$PLIST_NAME"
PLIST_DST="$HOME/Library/LaunchAgents/$PLIST_NAME"
CONFIG_DIR="$HOME/.config/outlook-gcal-sync"

# Try to find Python in a virtualenv first, then system
if [ -f "$PROJECT_DIR/.venv/bin/python3" ]; then
    PYTHON_PATH="$PROJECT_DIR/.venv/bin/python3"
elif command -v python3 &> /dev/null; then
    PYTHON_PATH="$(which python3)"
else
    echo "Error: python3 not found."
    exit 1
fi

install_agent() {
    echo "Installing launchd agent..."
    echo "  Python:  $PYTHON_PATH"
    echo "  Project: $PROJECT_DIR"
    echo "  Config:  $CONFIG_DIR"

    mkdir -p "$HOME/Library/LaunchAgents"
    mkdir -p "$CONFIG_DIR"

    # Generate plist with actual paths
    sed \
        -e "s|PYTHON_PATH_HERE|$PYTHON_PATH|g" \
        -e "s|PROJECT_PATH_HERE|$PROJECT_DIR|g" \
        -e "s|LOG_PATH_HERE|$CONFIG_DIR|g" \
        "$PLIST_SRC" > "$PLIST_DST"

    # Load the agent
    launchctl load "$PLIST_DST"

    echo ""
    echo "Installed! The sync will run every 15 minutes."
    echo "Plist: $PLIST_DST"
    echo ""
    echo "Useful commands:"
    echo "  launchctl list | grep outlook-gcal   # Check if running"
    echo "  launchctl unload $PLIST_DST           # Stop"
    echo "  launchctl load $PLIST_DST             # Start"
    echo "  tail -f $CONFIG_DIR/sync.log          # Watch logs"
}

uninstall_agent() {
    echo "Uninstalling launchd agent..."
    if [ -f "$PLIST_DST" ]; then
        launchctl unload "$PLIST_DST" 2>/dev/null || true
        rm "$PLIST_DST"
        echo "Removed: $PLIST_DST"
    else
        echo "Agent not installed."
    fi
}

case "${1:-}" in
    install)
        install_agent
        ;;
    uninstall)
        uninstall_agent
        ;;
    *)
        echo "Usage: $0 {install|uninstall}"
        exit 1
        ;;
esac
