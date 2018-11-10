function copyObject(inObject,outObject)
       
      ' Копирование атрибутов
      if not isNull(inObject) then
         ' Получаем атрибут и выводим на экран
         dim numAttr
         for numAttr=0 to inObject.vrAttrCount-1
             set inAttr = inObject.vrAttrByIndex(numAttr)
             if outObject.vrAttrExists(inAttr.vrName) then
                
                dim inheritedInClassName, inheritedOutClassName, inClassName, outClassName   
                inheritedOutClassName = outObject.vrAttrByName(inAttr.vrName).vrClassValue.vrInheritedFrom.vrClass.vrName 
                inheritedInClassName = inAttr.vrClassValue.vrInheritedFrom.vrClass.vrName
                outClassName = outObject.vrAttrByName(inAttr.vrName).vrClassValue.vrClass.vrName 
                inClassName = inAttr.vrClassValue.vrClass.vrName          
                
                if (inheritedOutClassName = inheritedInClassName) or inheritedOutClassName = inClassName or inheritedOutClassName = inClassName  then
                     outObject.vrAttrByName(inAttr.vrName).vrAssignFrom(inAttr) 
                end if
             
             end if    
         next
      end if
      
      ' Копирование подчиненных объектов
      set outClass = outObject.vrClass
      set iterInObject = inObject.vrObjectsVector.vrCreateIterator("",inObject,true) 
      if iterInObject.vrFirst then 
          do 
            dim numClass
            for numClass=0 to  outClass.vrChildsCount-1
              if outClass.vriChildClassItem(numClass).vrName = iterInObject.vrGetObject.vrClass.vrName then
                 outObject.vrAddChildLink(iterInObject.vrGetObject)
              end if
            next
         loop while iterInObject.vrNext
      end if
End Function


' Авторизуемся в универсальном справочнике
SET uniRef = CreateObject("UniReference.UniRefer")
if not uniRef.GlobalVars.Logon.LogonAsDialog(0) then
  MsgBox("Авторизация не произведена")
end if

' Получаем модель
set vModel = CreateObject("vkernel.VModel")

if not vModel.vrLoadModel("test.vtp",nothing,1) then
  MsgBox("Невозможно открыть фаил")
end if

' Применить права доступа, иначе открыт ТП только на чтение
vModel.vrApplySecurity()

' Получаем список из root для классов DSE
set objRoot = vModel.vrGetObjVector.vrItem(0)
set iterRoot = vModel.vrGetObjVector.vrCreateIterator("dse",objRoot,true)

' Получаем первый объект DSE
iterRoot.vrFirst
set objDSE = iterRoot.vrGetObject

if not isNull(objDSE) then
   ' Получаем атрибут namedse и выводим на экран
   set attrDSE = objDSE.vrAttrByName("namedse")
   if not isNull(attrDSE) then
       MsgBox(attrDSE.vrClassValue.vrClass.vrName)  
   end if
end if

' Создание операции
set objVecDSE = objDSE.vrObjectsVector
set objOper = objVecDSE.vrCreate("public_oper")
objDSE.vrAddChildLink(objOper)

'Получаем список операций
set iterOper = objDSE.vrObjectsVector.vrCreateIterator("operations",objDSE,true)
iterOper.vrFirst
set objOperIter = iterOper.vrGetObject 

' Копирование объекта
call copyObject(objOperIter,objOper)

call vmodel.vrSaveModel("test.vtp",nothing)
  
set objOperIter = nothing
set iterOper = nothing 

set objOper = nothing
set objVecDSE = nothing

set objDSE = nothing
set iterRoot = nothing
set objRoot = nothing

set vModel = nothing
set uniRef = nothing