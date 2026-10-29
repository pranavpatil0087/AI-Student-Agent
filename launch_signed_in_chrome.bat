@echo off
echo ========================================================
echo Starting Your Signed-in Google Chrome with Debugging Port 9222
echo ========================================================
echo Closing any background Chrome instances to unlock your profile...
taskkill /IM chrome.exe /F >nul 2>&1
timeout /t 1 >nul
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --profile-directory="Profile 2" --start-maximized "https://meet.google.com/gda-cfuw-qsk"
echo.
echo Your normal Chrome has been launched with your signed-in Google account (Profile 2)!
echo You can now run: python main.py
echo main.py will automatically attach to this browser window with all your logins intact.
echo ========================================================
