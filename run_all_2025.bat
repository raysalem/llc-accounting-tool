@echo off
set "VENDOR_PATH=\\192.168.1.90\Documents Private\taxes\2025\vendor.xlsx"
set "TARGET_DIR=\\192.168.1.90\Documents Private\taxes\2025\*.lnk"

echo [Batch] Using Vendor File: %VENDOR_PATH%
node batch_run.js "%TARGET_DIR%" --1099=NEC --vendor-file "%VENDOR_PATH%" --all --save
pause
