
' Пример работы с УТС 

' Авторизация
SET UniRef = CreateObject("UniReference.UniRefer")
if not UniRef.GlobalVars.Logon.LogonAsParams("Администратор","111","Администраторы") then
  MsgBox("Авторизация не произведена")
end if

' Получение корневого класса
Set ProfRef = UniRef.BOListGroup.GroupByName("Другие").RootClassByName("KTE")

If Not IsNull(ProfRef) Then
   ' Получение класса LIST
   Set ProfListRef = ProfRef.ClassByName("VID")
end if   

' Получение списка объектов
Set RootObjects = UniRef.BOListObject()

' Получить корневой объект для KTE.VID 
Set KTEObject = RootObjects.GetRootObject("KTE.VID ")

' Получить коллекцию объектов детей для корневого объекта  
if KTEObject.GetChildCollectionObject(true,true,true) then 
   ' Получить первый подчиненый объект 
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