echo off
set AMD64="AMD64"
SET VBS="replace_stat.vbs"
if %AMD64% == "%PROCESSOR_ARCHITECTURE%" ( %windir%\SysWow64\wscript.exe %VBS%) 
else ( %windir%\SYSTEM32\wscript.exe %VBS%)

