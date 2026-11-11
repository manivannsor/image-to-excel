@echo off
title Image to Excel Converter
cd /d "%~dp0"
echo =======================================================
echo    Image to Excel Converter (បម្លែងរូបភាពទៅជា Excel)
echo =======================================================
echo Starting server...
start http://localhost:3000
node server.js
pause

