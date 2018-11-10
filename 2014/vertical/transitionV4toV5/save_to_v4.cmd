echo off
set AMD64="AMD64"
SET VBS="save_to_v4.vbs"
if %AMD64% == "%PROCESSOR_ARCHITECTURE%" ( %windir%\SysWow64\wscript.exe %VBS%) 
else ( %windir%\SYSTEM32\wscript.exe %VBS%)

