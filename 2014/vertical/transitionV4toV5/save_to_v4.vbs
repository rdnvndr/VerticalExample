' ����� �����
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
   ' ������������ � ������������� �����������
   SET uniRef = CreateObject("UniReference.UniRefer")
   if not uniRef.GlobalVars.Logon.LogonAuto then
      MsgBox("����������� �� �����������")
      Exit function
   end if

   ' �������� ������
   set vModel = CreateObject("vkernel.VModel")
   if not vModel.vrLoadModel(strFile,nothing,2) then
      MsgBox("���������� ������� ����")
      Exit function
   end if

   ' ��������� ����� �������, ����� ������ �� ������ �� ������
   vModel.vrApplySecurity()
   ' ��������� ����������
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
   MsgBox "���� ��������:" & vbCR & "    " & outfile
   ' ver_type
   '   0 � �������� ������ ������ (��������) ���������� �����������.
   '   1 � �������� ������ ���������� (������� � ������������), ����������� � ���������� �����������.
   '   2 � �������� ������, � ������� ��������� ��������� ����������.
   '   3 � �������� ������ ����� �����������.
'   MsgBox vModel.vrFileVersion(3)
End Function

Dim strFile
strFile = SelectFile( )
If strFile = "" Then 
   MsgBox "���� �� ������."
else
   SaveFile
End If