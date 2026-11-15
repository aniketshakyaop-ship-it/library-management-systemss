@echo off
title ACET Library Management System - Local Web Server
echo ========================================================
echo    ALIGARH COLLEGE OF ENGINEERING AND TECHNOLOGY
echo       DEPARTMENT OF COMPUTER SCIENCE & ENGINEERING
echo           MINI PROJECT: LIBRARY MANAGEMENT SYSTEM
echo ========================================================
echo Student: Gazal Gupta (Roll No: 2401090100037)
echo Guide: Dr. Anand sir (HOD)
echo.
echo Starting local web server at http://localhost:8000 ...
echo Opening your web browser...
start "" "http://localhost:8000"
echo.
echo Server is running! Keep this window open while using the app.
echo Press Ctrl+C to stop the server anytime.
echo ========================================================
python -m http.server 8000
