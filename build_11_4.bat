@echo off
setlocal ENABLEDELAYEDEXPANSION

:: Читаемые русские сообщения
chcp 65001 >nul

echo ============= КругоЗор — сборка EXE (PyInstaller) =============

:: Определяем Python: сначала локальный, потом системный
set "PYTHON="
if exist "%~dp0python.exe" set "PYTHON=%~dp0python.exe"
if "%PYTHON%"=="" set "PYTHON=py"

:: Имена файлов
set "PRIMARY_SCRIPT=KruGoZor_11_4.py"
set "FALLBACK_SCRIPT=KruGoZor_11_3.py"
set "ICON=icon.ico"
set "APPNAME=KruGoZor"

:: Выбор скрипта
if exist "%PRIMARY_SCRIPT%" (
  set "SCRIPT=%PRIMARY_SCRIPT%"
  echo Использую %SCRIPT%
) else if exist "%FALLBACK_SCRIPT%" (
  set "SCRIPT=%FALLBACK_SCRIPT%"
  echo [ВНИМАНИЕ] %PRIMARY_SCRIPT% не найден — использую %SCRIPT%
) else (
  echo [ОШИБКА] Не найден ни %PRIMARY_SCRIPT%, ни %FALLBACK_SCRIPT%.
  pause
  exit /b 1
)

:: Проверка иконки
if not exist "%ICON%" (
  echo [ОШИБКА] Не найден %ICON%
  pause
  exit /b 1
)

:: Обновление pip и установка зависимостей
echo.
echo === Установка PyInstaller и библиотек ===
"%PYTHON%" -m pip install --upgrade pip
"%PYTHON%" -m pip install --upgrade pyinstaller
"%PYTHON%" -m pip install PyQt5 opencv-python numpy pyvirtualcam psutil keyboard

:: Очистка
rmdir /s /q build  2>nul
rmdir /s /q dist   2>nul
del /q "%APPNAME%.spec" 2>nul

:: Сборка
echo.
echo === Сборка ONEFILE ===
"%PYTHON%" -m PyInstaller ^
 --noconfirm ^
 --clean ^
 --onefile ^
 --windowed ^
 --name "%APPNAME%" ^
 --icon "%ICON%" ^
 --hidden-import PyQt5.sip ^
 --collect-submodules keyboard ^
 --collect-submodules psutil ^
 --collect-submodules pyvirtualcam ^
 --collect-submodules cv2 ^
 "%SCRIPT%"

:: Копируем icon.ico в dist, чтобы трею было откуда брать
if exist dist (
  copy /y "%ICON%" "dist\icon.ico" >nul
)

echo.
echo [OK] Готово! Файл: dist\%APPNAME%.exe
echo.
pause
