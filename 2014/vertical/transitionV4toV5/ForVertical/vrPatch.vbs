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

Function vrPatch(mdl_old, mdl_new)
...
...
...
  ' ������������ ���������� � ������������ ����������� ��� ������ V5 � ����
  If v_old < 27 Then
     set root =  new_objs.vrItem(0)
     if not IsConfirmed(root) then
        MsgBox("� ����������� �� ������ V5 ������� �� ����������� ���������!" _
               & vbCR & "���������� �� ��������� � ������ ������!")
     else
       '������� ������ ��� ������ � ���/���
       CLASS_TTPDocument = "{C7A531B8-1E9D-4D27-8B99-2C8E089CB827}"
       set VPlugins = vModel.vrGetPlugins
       set TTPDocument = nothing
       set TTPDocument = VPlugins.vrItemByStrGUID(CLASS_TTPDocument)

       ' ���� ����������� �� �������� ���/���
       if not (TTPDocument is nothing) then
          TTPDocument.UpdateTPFromV4SP1
          TTPDocument.RepairLinkTable
       end if

        ' ����������� �������� ��
        set Builder = CreateObject("ReportBuilder.RReportBuilder")
        Builder.rModel = mdl_new
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
     end if  
  End if
  'AdminMessage "�� �������� �� ������ 2.0.1"

  s_helper.vsAllowChangesByEvents mdl_new, false
  s_helper.vsAllowChangesByEvents mdl_old, false
End Function