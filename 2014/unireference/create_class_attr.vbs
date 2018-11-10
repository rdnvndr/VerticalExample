
cROOT="xROOT"
cCLASS="xCLASS"
cATTR1="xATTR1"
cATTR2="xATTR2"

' Авторизация
SET UR = CreateObject("UniReference.UniRefer")
if not UR.GlobalVars.Logon.LogonAsParams("Администратор","111","Администраторы") then
  MsgBox("Авторизация не произведена")
end if

' Получение корневого класса
set iGroup= UR.BOListGroup.GroupByName("Другие")
Set iRoot = iGroup.RootClassByName(cROOT)

If iRoot is nothing Then
   stGUID = UR.GlobalVars.GetGUID22
   set iRoot = iGroup.AddRootClass(stGUID, cROOT, cROOT, cROOT)
   iRoot.SaveRootClass
   MsgBox "Create  RootClass"
end if   

set iClass = iRoot.ClassByName(cCLASS)
If iClass is nothing Then
   stGUID = UR.GlobalVars.GetGUID22
   set iClass = iRoot.AddClass(stGUID, cCLASS, cCLASS, cCLASS, cCLASS, true)
   iClass.SaveClass
   MsgBox "Create Class"
end if

set iAttr1 = iClass.AttrClassByName(cATTR1)
If iAttr1 is nothing Then
   set iAttr1 = iClass.AddAttribute(cATTR1, cATTR1, 0, cATTR1, "", "string", 100, "", true)
   MsgBox "Create Attr1"
end if
iAttr1.MasMetaData(2)=True ' Caption
iAttr1.SaveAttr

set iAttr2 = iClass.AttrClassByName(cATTR2)
If iAttr2 is nothing Then
   set iAttr2 = iClass.AddAttribute(cATTR2, cATTR2, 0, cATTR2, "", "blob", 0, "", true)
   MsgBox "Create Attr2"
end if
iAttr2.MasMetaData(5)=True ' Blob
iAttr2.SaveAttr

MsgBox "All OK"