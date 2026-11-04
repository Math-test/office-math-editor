@echo off
chcp 65001 >nul
setlocal
rem 數學方程式編輯器：解除安裝
set "DIR=%LOCALAPPDATA%\OfficeMathEditor"
set "MF=%DIR%\manifest.xml"

echo 請先關閉所有 Word、Excel 視窗，再按任意鍵繼續。
pause >nul
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%MF%" /f >nul 2>nul
if exist "%DIR%" rmdir /S /Q "%DIR%"
echo.
echo 已解除安裝。重新打開 Word、Excel 後按鈕就會消失。
pause