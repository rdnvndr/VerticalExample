function IsConfirmed( obj )
   IsConfirmed = true
   set ObjToConfirm = nothing
   set VDSEs = obj.vrObjectsVector.vrCreateIterator("dse",obj,true)
   if VDSEs.vrFirst then
      'Проверяем утвержден ли ТП
      set VDSE = VDSEs.vrGetObject
      if VDSE.vrAttrByName("tp_confirmed").vrValue then
         'Получаем не утвержденное ИИ
         set VIIs = VDSE.vrObjectsVector.vrCreateIterator("changing",VDSE,true)
         do while IsConfirmed and VIIs.vrNext
            IsConfirmed = VIIs.vrGetObject.vrAttrByName("changing_done").vrValue
         loop   
      end if
   end if
end function

' Выбор файла
Function SelectFile( )
   Dim objExec, strMSHTA, wshShell
   SelectFile = ""

   strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
            & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
            & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""

   Set wshShell = CreateObject( "WScript.Shell" )
   Set objExec = wshShell.Exec( strMSHTA )
   SelectFile = objExec.StdOut.ReadLine( )
   Set objExec = Nothing
   Set wshShell = Nothing
End Function

Function SaveFile()
   ' Авторизуемся в универсальном справочнике
   SET uniRef = CreateObject("UniReference.UniRefer")
   if not uniRef.GlobalVars.Logon.LogonAuto then
      MsgBox("Авторизация не произведена")
      Exit function
   end if

   ' Получаем модель
   set vModel = CreateObject("vkernel.VModel")

   ' UNDEFINED = $00000000;            - Режим неизвестен. Не используется
   ' OPEN_FULL_STR_SERVER = $00000001; - Прочитать локальный (указанный) файл, применив к нему структуру метаданных (классы и безопасность), хранящуюся на сервере
   ' OPEN_FULL_STR_LOCAL = $00000002;  - Прочитать из локального файла и метаданные и данные.
   ' OPEN_SERVER_STR_ONLY = $00000004; - Прочитать метаданные, хранящиеся на сервере.
   ' OPEN_LOCAL_STR_ONLY = $00000008;  - Прочитать метаданные, хранящиеся локально.
   ' FILES_EXTRACT = $00000010;        - При чтении файла техпроцесса вытаскивать присоединенные к нему файлы чертежей и т.п. Применяется совместно с OPEN_FULL_STR_SERVER или OPEN_FULL_STR_LOCAL.
   ' STR_SERVER_CHECKIN = $00000020;   - Начиная с V3 не используется
   ' COMPACT_METADATA = $00000040;     - При записи файла техпроцесса удалять классы, неиспользуемые объектами данной модели. Значительно уменьшает размер файла. Техпроцесс, записанный таким образом, рекомендуется открывать исключительно в режиме синхронизации по серверу (OPEN_FULL_STR_SERVER) или только для чтения. В противном случае могут возникнуть проблемы с созданием новых объектов.
   ' WRITE_UNICODE = $00000080;        - Если флаг установлен, то при сохранении файла все строки будут записаны в кодировке UNICODE, иначе – в той, которая используется в Вашей системе по-умолчанию (WIN1251). Снятие флага приведет к уменьшению размера файла.
   ' OPEN_READONLY = $00000100;        - Открыть ТП только для чтения
   ' SAVE_MERGED = $00000200;          - При сохранении ТП сохранять объекты фрагментов в файл основного ТП. Состояние фрагментов ТП в памяти не изменять
   ' SAVE_UNTOUCHED = $00000400;       - Сохранить ТП не изменяя состояние модели в памяти, т.е. сделать независимую копию ТП
 
   if not vModel.vrLoadModel(strFile,nothing,2) then
      MsgBox("Невозможно открыть фаил")
      Exit function
   end if
   
   if vModel.vrFileVersion(3) > 26 then
      MsgBox("Скрипт работает только с техпроцессами до версии V5!")
      Exit function   
   end if

   ' Применить права доступа, иначе открыт ТП только на чтение
   vModel.vrApplySecurity()

   'Получим плагин для работы с ТТП/ГТП
   ON ERROR RESUME NEXT
   CLASS_TTPDocument = "{C7A531B8-1E9D-4D27-8B99-2C8E089CB827}"
   set VPlugins = vModel.vrGetPlugins
   set TTPDocument = nothing
   set TTPDocument = VPlugins.vrItemByStrGUID(CLASS_TTPDocument)

   ' если обновляемый ТП является ТТП/ГТП
   if not (TTPDocument is nothing) then
      TTPDocument.UpdateTPFromV4SP1
      TTPDocument.RepairLinkTable
   end if
  
   set root =  vModel.vrGetObjVector.vrItem(0)
   if not IsConfirmed(root) then
        MsgBox("В техпроцессе до версии V5 имеются не утверждённые извещения!" _
               & vbCR & "Необходимо их утвердить в старой версии!")
        Exit function  
   end if
   
   ' Контрольный комплект ТД
   set Builder = CreateObject("ReportBuilder.RReportBuilder")
   Builder.rModel = vModel
   Builder.rConfirmingTP = false
   set package = Builder.rBuildStart
   
   ' Получить статистику
   set Essentials = CreateObject("ReportAddons.ReportEssentials")
   if package is nothing then
      for i_doc = 0 to Builder.rBuildDocCount-1
        set p_bdoc = Builder.rBuildDoc(i_doc)
        Essentials.rGatherEssentials p_bdoc.rDstDocument
      next
   else
      Essentials.rGatherEssentials package
   end if

   ' Сохраним предыдущую статистику
   set pack_data = root.vrAttrByName("package_part")
   if not (pack_data is nothing) and root.vrAttrExists("prev_package_part") then
      set prev_pack = root.vrAttrByName("prev_package_part")
      if not (prev_pack is nothing) then
         prev_pack.vrFile.vsDiskFullName = pack_data.vrFile.vsInternalFullName
      end if
   end if   
   ' Записать статистику
   Essentials.rSaveToAttr pack_data

   ' Сохранить техпроцесс
   Set objRegEx = CreateObject("VBScript.RegExp")
   objRegEx.Global = True   
   objRegEx.IgnoreCase = False
   objRegEx.Pattern = "\.vtp$"
   outfile = objRegEx.Replace(strFile, "_out.vtp")
   if outfile = strFile then
      objRegEx.Pattern = "\.ttp$"
      outfile = objRegEx.Replace(strFile, "_out.ttp")
   end if
   call vModel.vrSaveModel(outfile, nothing)
   MsgBox "Файл сохранен:" & vbCR & "    " & outfile
End Function

Dim strFile
strFile = SelectFile( )
If strFile = "" Then 
   MsgBox "Файл не выбран."
else
   SaveFile
End If