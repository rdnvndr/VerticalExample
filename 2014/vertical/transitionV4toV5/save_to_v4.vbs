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
   if not vModel.vrLoadModel(strFile,nothing,2) then
      MsgBox("Невозможно открыть фаил")
      Exit function
   end if

   ' Применить права доступа, иначе открыт ТП только на чтение
   vModel.vrApplySecurity()
   ' Сохранить техпроцесс
   Set objRegEx = CreateObject("VBScript.RegExp")
   objRegEx.Global = True   
   objRegEx.IgnoreCase = False
   objRegEx.Pattern = "\.vtp$"
   outfile = objRegEx.Replace(strFile, "_v4.vtp")
   if outfile = strFile then
      objRegEx.Pattern = "\.ttp$"
      outfile = objRegEx.Replace(strFile, "_v4.ttp")
   end if
   call vModel.vrSaveModelVersion(outfile, nothing, 26)
   MsgBox "Файл сохранен:" & vbCR & "    " & outfile
   ' ver_type
   '   0 – получить версию данных (объектов) локального техпроцесса.
   '   1 – получить версию метаданных (классов и безопасности), примененных к локальному техпроцессу.
   '   2 – получить версию, в которой требуется сохранить техпроцесс.
   '   3 – получить версию файла техпроцесса.
'   MsgBox vModel.vrFileVersion(3)
End Function

Dim strFile
strFile = SelectFile( )
If strFile = "" Then 
   MsgBox "Файл не выбран."
else
   SaveFile
End If