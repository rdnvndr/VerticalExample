function showObjVrt(startObj)

   set obj_vec = startObj.vrObjectsVector
   set obj_iter = obj_vec.vrCreateIterator("",startObj,true) 
   MsgBox  startObj.vrClass.vrName & "=" & startObj.vrObjStrID 
   ' & cstr(startObj.vrAttrByName("indexoper").vrValue)
   do while obj_iter.vrNext  
      call showObjVrt(obj_iter.vrGetObject)
   loop
       
   set obj_iter = nothing
   set obj_vec = nothing   

end function

' Авторизуемся в универсальном справочнике
SET m_uniref = CreateObject("UniReference.UniRefer")
if not m_uniref.GlobalVars.Logon.LogonAsDialog(0) then
  MsgBox("Авторизация не произведена")
end if

' Получаем модель
set vmodel = CreateObject("vkernel.VModel")

if not vmodel.vrLoadModel("test.vtp",nothing,1) then
  MsgBox("Невозможно открыть фаил")
end if

' Применить права доступа, иначе открыт ТП только на чтение
vmodel.vrApplySecurity()

set startObj = vmodel.vrGetObjVector.vrGetObjByStrID("{A6108314-27EE-4636-A41D-AC01028648CE}")
call showObjVrt(startObj)

set startObj = nothing          
set vmodel = nothing
set m_uniref = nothing