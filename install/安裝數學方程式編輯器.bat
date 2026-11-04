@echo off
chcp 65001 >nul
setlocal
rem 數學方程式編輯器：安裝（Word / Excel 桌面版 2016 以上）
rem 在登錄檔登記增益集設定檔（HKCU，不需要系統管理員權限）
set "DIR=%LOCALAPPDATA%\OfficeMathEditor"
set "MF=%DIR%\manifest.xml"
set "URL=https://math-test.github.io/office-math-editor/manifest.xml"

echo ============================================
echo   數學方程式編輯器  安裝
echo ============================================
echo.
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul && goto busy
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul && goto busy
goto go

:busy
echo 請先把所有 Word、Excel 視窗存檔並關閉，再按任意鍵繼續。
pause >nul
goto go

:go
if not exist "%DIR%" mkdir "%DIR%"
echo 正在下載最新的設定檔...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol = 'Tls12'; Invoke-WebRequest -UseBasicParsing '%URL%' -OutFile '%MF%' -TimeoutSec 30; exit 0 } catch { exit 1 }"
if not errorlevel 1 goto reg
rem 下載失敗：用跟這個檔案放在一起的 manifest.xml
if exist "%~dp0manifest.xml" (
  copy /Y "%~dp0manifest.xml" "%MF%" >nul
  goto reg
)
if exist "%~dp0..\manifest.xml" (
  copy /Y "%~dp0..\manifest.xml" "%MF%" >nul
  goto reg
)
echo.
echo [失敗] 無法下載設定檔，請確認網路後再試一次。
pause
exit /b 1

:reg
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%MF%" /t REG_SZ /d "%MF%" /f >nul
if errorlevel 1 (
  echo [失敗] 無法寫入設定，請把這個畫面截圖給管理者。
  pause
  exit /b 1
)
echo.
echo 安裝完成！
echo.
echo  1. 打開 Word，在「常用」索引標籤最右邊按「數學方程式」。
echo  2. Excel 也一樣，在「常用」最右邊。
echo  3. 使用時需要連網。
echo.
pause