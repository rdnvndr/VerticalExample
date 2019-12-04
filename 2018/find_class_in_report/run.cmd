@echo off

SET VBS="find_class_in_report.vbs"

set AMD64="AMD64"
if %AMD64% == "%PROCESSOR_ARCHITECTURE%" ( 
    %windir%\SysWow64\wscript.exe %VBS% 
) else ( 
    %windir%\SYSTEM32\wscript.exe %VBS% 
)

