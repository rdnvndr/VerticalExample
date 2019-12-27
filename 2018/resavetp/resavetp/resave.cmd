set REPO=..\Structure\Vertical\Program Files\ASCON\Vertical
set PROG=C:\Program Files (x86)\ASCON\Vertical

copy "%REPO%\Template\assembly.vtp" "%PROG%\Template\assembly.vtp" /Y
copy "%REPO%\Template\detail.vtp"   "%PROG%\Template\detail.vtp" /Y
copy "%REPO%\Template\ttp.ttp"      "%PROG%\Template\ttp.ttp" /Y

copy "%REPO%\Samples\планка.vtp"       "%PROG%\Samples\планка.vtp" /Y
copy "%REPO%\Samples\Вал выходной.vtp" "%PROG%\Samples\Вал выходной.vtp" /Y
copy "%REPO%\Samples\АБВ.001.005 Лист рессоры.ttp"   "%PROG%\Samples\АБВ.001.005 Лист рессоры.ttp" /Y
copy "%REPO%\Samples\Редуктор\078.505.9.0100.00.vtp" "%PROG%\Samples\Редуктор\078.505.9.0100.00.vtp" /Y


resave.exe "%REPO%\RestoreFiles\Model\structure.vtp" ^
           "%PROG%\Template\assembly.vtp" ^
           "%PROG%\Template\detail.vtp" ^
           "%PROG%\Template\ttp.ttp" ^
           "%PROG%\Samples\планка.vtp" ^
           "%PROG%\Samples\Вал выходной.vtp" ^
           "%PROG%\Samples\АБВ.001.005 Лист рессоры.ttp" ^
           "%PROG%\Samples\Редуктор\078.505.9.0100.00.vtp"


copy "%PROG%\Template\assembly.vtp" "%REPO%\Template\assembly.vtp" /Y
copy "%PROG%\Template\detail.vtp"   "%REPO%\Template\detail.vtp" /Y
copy "%PROG%\Template\ttp.ttp"      "%REPO%\Template\ttp.ttp" /Y

copy "%PROG%\Samples\планка.vtp"       "%REPO%\Samples\планка.vtp"  /Y
copy "%PROG%\Samples\Вал выходной.vtp" "%REPO%\Samples\Вал выходной.vtp" /Y
copy "%PROG%\Samples\АБВ.001.005 Лист рессоры.ttp"   "%REPO%\Samples\АБВ.001.005 Лист рессоры.ttp" /Y
copy "%PROG%\Samples\Редуктор\078.505.9.0100.00.vtp" "%REPO%\Samples\Редуктор\078.505.9.0100.00.vtp" /Y
