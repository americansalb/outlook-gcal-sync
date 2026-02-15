"""Tests for the AppleScript bridge module."""

import subprocess
import pytest
from unittest.mock import patch, MagicMock

from src.outlook.applescript_bridge import (
    run_applescript,
    escape_applescript_string,
    check_outlook_running,
    list_calendars,
    AppleScriptError,
    OutlookNotRunningError,
)


class TestEscapeAppleScriptString:
    def test_plain_string(self):
        assert escape_applescript_string("hello") == "hello"

    def test_double_quotes(self):
        assert escape_applescript_string('say "hi"') == 'say \\"hi\\"'

    def test_backslash(self):
        assert escape_applescript_string("path\\to") == "path\\\\to"

    def test_newline(self):
        assert escape_applescript_string("line1\nline2") == "line1\\nline2"

    def test_combined(self):
        result = escape_applescript_string('He said "hi"\nBye\\end')
        assert result == 'He said \\"hi\\"\\nBye\\\\end'

    def test_empty_string(self):
        assert escape_applescript_string("") == ""


class TestRunAppleScript:
    @patch("src.outlook.applescript_bridge.subprocess.run")
    def test_success(self, mock_run):
        mock_run.return_value = MagicMock(
            returncode=0, stdout="Calendar, Work\n", stderr=""
        )
        result = run_applescript('tell application "Microsoft Outlook" to return "hi"')
        assert result == "Calendar, Work"

    @patch("src.outlook.applescript_bridge.subprocess.run")
    def test_failure(self, mock_run):
        mock_run.return_value = MagicMock(
            returncode=1, stdout="", stderr="execution error: some error"
        )
        with pytest.raises(AppleScriptError, match="AppleScript failed"):
            run_applescript("bad script")

    @patch("src.outlook.applescript_bridge.subprocess.run")
    def test_outlook_not_running(self, mock_run):
        mock_run.return_value = MagicMock(
            returncode=1, stdout="",
            stderr="execution error: Application isn't running. (-600)"
        )
        with pytest.raises(OutlookNotRunningError):
            run_applescript('tell application "Microsoft Outlook" to return name')

    @patch("src.outlook.applescript_bridge.subprocess.run")
    def test_timeout(self, mock_run):
        mock_run.side_effect = subprocess.TimeoutExpired(cmd="osascript", timeout=10)
        with pytest.raises(AppleScriptError, match="timed out"):
            run_applescript("slow script", timeout=10)


class TestCheckOutlookRunning:
    @patch("src.outlook.applescript_bridge.run_applescript")
    def test_running(self, mock_run):
        mock_run.return_value = "true"
        assert check_outlook_running() is True

    @patch("src.outlook.applescript_bridge.run_applescript")
    def test_not_running(self, mock_run):
        mock_run.return_value = "false"
        assert check_outlook_running() is False

    @patch("src.outlook.applescript_bridge.run_applescript")
    def test_error(self, mock_run):
        mock_run.side_effect = AppleScriptError("fail")
        assert check_outlook_running() is False


class TestListCalendars:
    @patch("src.outlook.applescript_bridge.run_applescript")
    def test_multiple_calendars(self, mock_run):
        mock_run.return_value = "Calendar, Work, Personal"
        result = list_calendars()
        assert result == ["Calendar", "Work", "Personal"]

    @patch("src.outlook.applescript_bridge.run_applescript")
    def test_single_calendar(self, mock_run):
        mock_run.return_value = "Calendar"
        result = list_calendars()
        assert result == ["Calendar"]

    @patch("src.outlook.applescript_bridge.run_applescript")
    def test_empty(self, mock_run):
        mock_run.return_value = ""
        result = list_calendars()
        assert result == []
