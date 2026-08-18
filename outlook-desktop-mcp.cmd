@echo off
REM Outlook Desktop MCP - Launcher
REM Usage:
REM   outlook-desktop-mcp.cmd mcp    Start the MCP server (stdio)
REM   outlook-desktop-mcp.cmd test   Run COM validation tests

setlocal

set VENV=%~dp0.venv
uv sync -U --dev --directory %~dp0

if not exist "%VENV%" (
   echo ERROR: Virtual environment not found. Run setup first. 1>&2
   exit /b 1
)
set PYTHON=uv run --directory %~dp0

if "%1"=="mcp" (
    %PYTHON% outlook-desktop-mcp
) else if "%1"=="test" (
    %PYTHON% tests\test_email_com.py
) else (
    echo Usage: outlook-desktop-mcp.cmd [mcp^|test] 1>&2
    exit /b 1
)
