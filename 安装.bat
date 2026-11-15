@echo off
chcp 65001 >nul
title 日程球 安装
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
