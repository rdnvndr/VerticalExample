strFile = "structureV5.vtp"

' Авторизация через Полином
set manager = CreateObject("Ascon.Integration.AuthenticationManager")
call manager.Authenticate()

' Получаем модель
set vModel = CreateObject("vkernel.VModel")

' Режимы открытия техпроцесса или модели
UNDEFINED            =    0 ' Режим неизвестен. Не используется
OPEN_FULL_STR_SERVER =    1 ' Прочитать локальный (указанный) файл, применив к 
                            ' нему структуру метаданных (классы и безопасность), 
                            ' хранящуюся на сервере
OPEN_FULL_STR_LOCAL  =    2 ' Прочитать из локального файла и метаданные и данные
OPEN_SERVER_STR_ONLY =    4 ' Прочитать метаданные, хранящиеся на сервере
OPEN_LOCAL_STR_ONLY  =    8 ' Прочитать метаданные, хранящиеся локально
FILES_EXTRACT        =   16 ' При чтении файла техпроцесса вытаскивать присоединенные 
                            ' к нему файлы чертежей и т.п. Применяется совместно 
                            ' с OPEN_FULL_STR_SERVER или OPEN_FULL_STR_LOCAL.
STR_SERVER_CHECKIN   =   32 ' Начиная с V3 не используется
COMPACT_METADATA     =   64 ' При записи файла техпроцесса удалять классы, 
                            ' неиспользуемые объектами данной модели.
WRITE_UNICODE        =  128 ' Если флаг установлен, то при сохранении файла все 
                            ' строки будут записаны в кодировке UNICODE, иначе – 
                            ' в той, которая используется в Вашей системе 
                            ' по-умолчанию (WIN1251).
OPEN_READONLY        =  256 ' Открыть ТП только для чтения
SAVE_MERGED          =  512 ' При сохранении ТП сохранять объекты фрагментов в 
                            ' файл основного ТП. Состояние фрагментов ТП в памяти 
                            ' не изменять
SAVE_UNTOUCHED       = 1024 ' Сохранить ТП не изменяя состояние модели в памяти, 
                            ' т.е. сделать независимую копию ТП 

if not vModel.vrLoadModel(strFile, nothing, OPEN_LOCAL_STR_ONLY) then 
   MsgBox("Невозможно открыть фаил модели")
   quit
end if

' Применяет права доступа
vModel.vrApplySecurity()

' Убираем проверку плагинов
vModel.vrGetClassVector.vrLocate("root").vrnClassValueItem("after_load").vrFunctionCode = ""

' Заменяет функцию обновления   
Set FSO = CreateObject("Scripting.FileSystemObject")
Set File = FSO.GetFile("update.vbs")
Set TextStream = File.OpenAsTextStream(1)
vModel.vrGetClassVector.vrPatchFunction("vrPatch")=TextStream.ReadAll()
TextStream.Close

' Сохраняет модель
call vModel.vrResetOpenMode(OPEN_SERVER_STR_ONLY + WRITE_UNICODE)
call vModel.vrSaveModel("structure.vtp", nothing)
MsgBox "Файл сохранен"

' Получение версии файла
DATAVER     = 0 ' получить версию данных (объектов) локального техпроцесса.
METADATAVER = 1 ' получить версию метаданных (классов и безопасности), примененных 
                ' к локальному техпроцессу.
SAVEVER     = 2 ' получить версию, в которой требуется сохранить техпроцесс.
FULLVER     = 3 ' получить версию файла техпроцесса.
' MsgBox vModel.vrFileVersion(FULLVER)

call manager.Deauthenticate()
