function IsConfirmed( obj )
   IsConfirmed = true
   set ObjToConfirm = nothing
   set VDSEs = obj.vrObjectsVector.vrCreateIterator("dse",obj,true)
   if VDSEs.vrFirst then
      '��������� ��������� �� ��
      set VDSE = VDSEs.vrGetObject
      if VDSE.vrAttrByName("tp_confirmed").vrValue then
         '�������� �� ������������ ��
         set VIIs = VDSE.vrObjectsVector.vrCreateIterator("changing",VDSE,true)
         do while IsConfirmed and VIIs.vrNext
            IsConfirmed = VIIs.vrGetObject.vrAttrByName("changing_done").vrValue
         loop   
      end if
   end if
end function

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

   ' UNDEFINED = $00000000;            - ����� ����������. �� ������������
   ' OPEN_FULL_STR_SERVER = $00000001; - ��������� ��������� (���������) ����, �������� � ���� ��������� ���������� (������ � ������������), ���������� �� �������
   ' OPEN_FULL_STR_LOCAL = $00000002;  - ��������� �� ���������� ����� � ���������� � ������.
   ' OPEN_SERVER_STR_ONLY = $00000004; - ��������� ����������, ���������� �� �������.
   ' OPEN_LOCAL_STR_ONLY = $00000008;  - ��������� ����������, ���������� ��������.
   ' FILES_EXTRACT = $00000010;        - ��� ������ ����� ����������� ����������� �������������� � ���� ����� �������� � �.�. ����������� ��������� � OPEN_FULL_STR_SERVER ��� OPEN_FULL_STR_LOCAL.
   ' STR_SERVER_CHECKIN = $00000020;   - ������� � V3 �� ������������
   ' COMPACT_METADATA = $00000040;     - ��� ������ ����� ����������� ������� ������, �������������� ��������� ������ ������. ����������� ��������� ������ �����. ����������, ���������� ����� �������, ������������� ��������� ������������� � ������ ������������� �� ������� (OPEN_FULL_STR_SERVER) ��� ������ ��� ������. � ��������� ������ ����� ���������� �������� � ��������� ����� ��������.
   ' WRITE_UNICODE = $00000080;        - ���� ���� ����������, �� ��� ���������� ����� ��� ������ ����� �������� � ��������� UNICODE, ����� � � ���, ������� ������������ � ����� ������� ��-��������� (WIN1251). ������ ����� �������� � ���������� ������� �����.
   ' OPEN_READONLY = $00000100;        - ������� �� ������ ��� ������
   ' SAVE_MERGED = $00000200;          - ��� ���������� �� ��������� ������� ���������� � ���� ��������� ��. ��������� ���������� �� � ������ �� ��������
   ' SAVE_UNTOUCHED = $00000400;       - ��������� �� �� ������� ��������� ������ � ������, �.�. ������� ����������� ����� ��
 
   if not vModel.vrLoadModel(strFile,nothing,2) then
      MsgBox("���������� ������� ����")
      Exit function
   end if
   
   if vModel.vrFileVersion(3) > 26 then
      MsgBox("������ �������� ������ � ������������� �� ������ V5!")
      Exit function   
   end if

   ' ��������� ����� �������, ����� ������ �� ������ �� ������
   vModel.vrApplySecurity()

   '������� ������ ��� ������ � ���/���
   ON ERROR RESUME NEXT
   CLASS_TTPDocument = "{C7A531B8-1E9D-4D27-8B99-2C8E089CB827}"
   set VPlugins = vModel.vrGetPlugins
   set TTPDocument = nothing
   set TTPDocument = VPlugins.vrItemByStrGUID(CLASS_TTPDocument)

   ' ���� ����������� �� �������� ���/���
   if not (TTPDocument is nothing) then
      TTPDocument.UpdateTPFromV4SP1
      TTPDocument.RepairLinkTable
   end if
  
   set root =  vModel.vrGetObjVector.vrItem(0)
   if not IsConfirmed(root) then
        MsgBox("� ����������� �� ������ V5 ������� �� ����������� ���������!" _
               & vbCR & "���������� �� ��������� � ������ ������!")
        Exit function  
   end if
   
   ' ����������� �������� ��
   set Builder = CreateObject("ReportBuilder.RReportBuilder")
   Builder.rModel = vModel
   Builder.rConfirmingTP = false
   set package = Builder.rBuildStart
   
   ' �������� ����������
   set Essentials = CreateObject("ReportAddons.ReportEssentials")
   if package is nothing then
      for i_doc = 0 to Builder.rBuildDocCount-1
        set p_bdoc = Builder.rBuildDoc(i_doc)
        Essentials.rGatherEssentials p_bdoc.rDstDocument
      next
   else
      Essentials.rGatherEssentials package
   end if

   ' �������� ���������� ����������
   set pack_data = root.vrAttrByName("package_part")
   if not (pack_data is nothing) and root.vrAttrExists("prev_package_part") then
      set prev_pack = root.vrAttrByName("prev_package_part")
      if not (prev_pack is nothing) then
         prev_pack.vrFile.vsDiskFullName = pack_data.vrFile.vsInternalFullName
      end if
   end if   
   ' �������� ����������
   Essentials.rSaveToAttr pack_data

   ' ��������� ����������
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
   MsgBox "���� ��������:" & vbCR & "    " & outfile
End Function

Dim strFile
strFile = SelectFile( )
If strFile = "" Then 
   MsgBox "���� �� ������."
else
   SaveFile
End If