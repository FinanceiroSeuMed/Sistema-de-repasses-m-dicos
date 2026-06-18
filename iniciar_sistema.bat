@echo off
title Sistema de Repasses Medicos - SeuMed
cd /d "C:\RepassesmedicosOMIE"
echo ============================================================
echo   Sistema de Repasses Medicos - SeuMed
echo ============================================================
echo.
echo   Iniciando o servidor...
echo   Abra no navegador:  http://127.0.0.1:8000/
echo.
echo   (Deixe esta janela aberta enquanto usa o sistema.)
echo   (Para encerrar: feche esta janela ou tecle Ctrl+C.)
echo ============================================================
echo.
".venv\Scripts\python.exe" manage.py runserver 127.0.0.1:8000 --noreload
echo.
echo Servidor encerrado.
pause
