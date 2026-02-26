@echo off
:: Use the full server path because P: does not exist for background tasks
xcopy "C:\DPR\ChicagoReport\*" "\\192.168.0.14\Public\Production\DPR\" /Y /I /E /F