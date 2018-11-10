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

Function vrPatch(mdl_old, mdl_new)
...
...
...
  ' Переписываем статистику у утверждённого техпроцесса для версии V5 и выше
  If v_old < 27 Then
     set root =  new_objs.vrItem(0)
     if not IsConfirmed(root) then
        MsgBox("В техпроцессе до версии V5 имеются не утверждённые извещения!" _
               & vbCR & "Необходимо их утвердить в старой версии!")
     else
       'Получим плагин для работы с ТТП/ГТП
       CLASS_TTPDocument = "{C7A531B8-1E9D-4D27-8B99-2C8E089CB827}"
       set VPlugins = vModel.vrGetPlugins
       set TTPDocument = nothing
       set TTPDocument = VPlugins.vrItemByStrGUID(CLASS_TTPDocument)

       ' если обновляемый ТП является ТТП/ГТП
       if not (TTPDocument is nothing) then
          TTPDocument.UpdateTPFromV4SP1
          TTPDocument.RepairLinkTable
       end if

        ' Контрольный комплект ТД
        set Builder = CreateObject("ReportBuilder.RReportBuilder")
        Builder.rModel = mdl_new
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
     end if  
  End if
  'AdminMessage "ТП обновлен до версии 2.0.1"

  s_helper.vsAllowChangesByEvents mdl_new, false
  s_helper.vsAllowChangesByEvents mdl_old, false
End Function