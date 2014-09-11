@REM -*- mode: batch; coding: utf-8 -*-
@REM (c) Valik mailto:vasnake@gmail.com

@echo off
chcp 1251 > nul
set wd=%~dp0
pushd "%wd%"
set NLS_LANG=AMERICAN_CIS.UTF8
set PYTHONPATH=
@cls

FOR /F "eol=* tokens=1* delims=*" %%i in (dwg.list) do (
	@echo %%i [%%~nxi]
	c:\python25\python -u "dwg.dump.py" "%%i" 1>%%~nxi.log 2>%%~nxi.err
)
popd
exit
