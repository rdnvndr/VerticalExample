
' ������ ������ � ��� 

' �����������
SET UniRef = CreateObject("UniReference.UniRefer")
if not UniRef.GlobalVars.Logon.LogonAsParams("�������������","111","��������������") then
  MsgBox("����������� �� �����������")
end if

' ��������� ��������� ������
Set ProfRef = UniRef.BOListGroup.GroupByName("������").RootClassByName("KTE")

If Not IsNull(ProfRef) Then
   ' ��������� ������ LIST
   Set ProfListRef = ProfRef.ClassByName("VID")
end if   

' ��������� ������ ��������
Set RootObjects = UniRef.BOListObject()

' �������� �������� ������ ��� KTE.VID 
Set KTEObject = RootObjects.GetRootObject("KTE.VID ")

' �������� ��������� �������� ����� ��� ��������� �������  
if KTEObject.GetChildCollectionObject(true,true,true) then 
   ' �������� ������ ���������� ������ 
   Set GroupObject = KTEObject.ChildObjects(3)  

   Set Attr = GroupObject.Attributes.AttrObjectByIndex(2)  
   MsgBox   Attr.FileExt
 
   Attr.SaveToFile("XXX.PNG")
   MsgBox(Attr.NameAttr + "=" + Attr.DataAttr)

'   Const ForReading = 1, ForWriting = 2, ForAppending = 8
'   Dim fso, f
'   Set fso = CreateObject("Scripting.FileSystemObject")
'   Set f = fso.createTextFile("blob" & ".bmp")
'   f.Write Attr.SaveToStream
'   f.Close
end if        

set Attr = nothing
set ListObject = nothing
set GroupObject = nothing
set ProfObject = nothing
set RootObjects = nothing

set ProfListRef = nothing
set ProfRef = nothing
set UniRef = nothing