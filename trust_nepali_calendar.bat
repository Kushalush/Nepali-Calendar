@echo off
set "EXCLUDE_PATH=%APPDATA%\NepaliCalendar"

echo Adding Nepali Calendar to Windows Defender exclusion list...
powershell -Command "Add-MpPreference -ExclusionPath '%EXCLUDE_PATH%'"

echo ✅ Done! Nepali Calendar is now trusted on this user account.
pause
