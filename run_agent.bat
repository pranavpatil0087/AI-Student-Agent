@echo off
title AI Student Agent - Web Dashboard
color 0A
echo =================================================================
echo                 AI STUDENT AGENT - WEB DASHBOARD
echo =================================================================
echo  Dashboard running at: http://localhost:5000
echo  Enter your Google Meet link directly on the dashboard screen.
echo  The agent will automate Chrome, mute mic/cam, and join the class!
echo =================================================================
echo.

python main.py

echo.
echo =================================================================
echo  Agent session completed. Notes processed and saved.
echo =================================================================
pause

