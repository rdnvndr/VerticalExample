dim apkg_id()
dim apkg_obozn()
dim apkg_name()

dim opkg_id()
dim opkg_obozn()
dim opkg_name()                    

dim parent_id()

Function vrPatch(mdl_old, mdl_new)
  ON ERROR RESUME NEXT
  msgbox "vrPatchV4"
  v_old = mdl_old.vrFileVersion(3) 'старая версия
  v_new = mdl_new.vrFileVersion(1) 'новая версия
  'от 1 до 3 - первая версия (V1)
  'от 4 до 12 - вторая версия (V2)
  '13 - вторая версия SP1 (V2SP1)
  'от 14 до 16 - третья версия (V3)
  'от 17 до 24 - четвертая версия (V4)
  'от 25 до 26 - четвертая версия SP1 (V4SP1)
  '27 - пятая версия (V5)

  CLASS_TTPDocument = "{C7A531B8-1E9D-4D27-8B99-2C8E089CB827}"
  
  set old_objs = mdl_old.vrGetObjVector
  set old_FilterVector = mdl_old.vrGetClassVector.vrFilterVector
  set new_objs = mdl_new.vrGetObjVector

  set UniRefer = GetObject(v_zero, "UniReference.UniRefer")

  set new_dse = GetDSEObject(mdl_new)
  must_convert_packages = (not(new_dse is nothing))and(new_dse.vrClass.vrName="assembly")
  set old_dse = GetDSEObject(mdl_old)

  'специальная заглушка для сборки мусора, синхронизирует связи новой модели с новой структурой
  new_dse.vrAddChildLink new_objs.vrItem(0)
  
  'Игонорируем запреты плагинов на изменение моделей
  set s_helper = CreateObject("cfg_tools.VScriptHelper")
  s_helper.vsAllowChangesByEvents mdl_new, true
  s_helper.vsAllowChangesByEvents mdl_old, true
  
  redim apkg_id(0)
  redim apkg_obozn(0)
  redim apkg_name(0)

  For i = 0 To old_objs.vrObjectsCount - 1 'Цикл по объектам
       set old_obj = old_objs.vrItem(i)
       set new_obj = new_objs.vrGetObjByStrID(old_obj.vrObjStrID) 'объект в новой структуре
       If Not new_obj is Nothing Then

           '''''''''''''''' Обновим значения атрибутов если модель младше V3. Begin ''''''''''''''''
             'В третей версии было применено наследование, поэтому имена классов и атрибутов
             'остались теже, но из-за примененного наслелования идентификаторы атрибутов класса
             'поменялись, в результате модель будет считать эти атрибы как новые атрибуты класса
             'и естественно значений старых атриубтов в объектах не окажется. Поэтому
             'просто получаем по имени атрибут старого объекта, атрибут нового объекта
             'и копируем значение из старого атрибута в новый
             'ВНИМАНИЕ: Проводим обновление значений атрибутов модели в самом начале, т.к. ниже
             'в зависимости от версии может происходить корректировка значений атрибутов
           If (v_old < 16) Then  
             set VMaterialClass = mdl_new.vrGetClassVector.vrLocate("material")
             set VStepClass = mdl_new.vrGetClassVector.vrLocate("step")
             'Было введено наследование в результате чего поплыли идентификаторы атрибутов
             'но имена атрибутов остались, поэтому синхронизацию проводим опираясь на имена атрибутов
             'set s_helper = CreateObject("cfg_tools.VScriptHelper")
             For j = 0 To new_obj.vrAttrCount - 1 'Цикл по атрибутам нового объекта

               'Копируем значения атрибутов из старой модели в новую
               set VAttribute = new_obj.vrAttrByIndex(j)
               if VAttribute.vrClassValue.vrType=0 then
                 'Если копируется текст переход то переприсваиваем его через vrValue,
                 'т.к. у этого атрибута поменялся тип данных с String на Text
                 if new_obj.vrClass.vrIsBase(VStepClass)and(VAttribute.vrName="name") then
                   VAttribute.vrValue=old_obj.vrAttrByName("name").vrValue
                 'Обрабатываем материал
                 elseif new_obj.vrClass.vrIsBase(VMaterialClass) then
                   'копируем значение атрибута obozn -> gost
                   'т.к. атрибут obozn был удален из материала
                   if (VAttribute.vrName="gost") then
                     VAttribute.vrAssignFrom(old_obj.vrAttrByName("obozn"))
                   end if
                   'Пробуем получить значение ЕН со старого объекта
                   'если этого сделать не удается то инициализируем его 1
                   if (VAttribute.vrName="en") then
                     if old_obj.vrAttrExists("en") then
                       en = old_obj.vrAttrByName("en").vrValue
                       'Если сконвертировать строковое значение ЕН не удасться(произойдет ошибка)
                       'то в результате того что работает директива ON ERROR RESUME NEXT
                       'мы перейдем на следующую строку.
                       if (en="")or(cint(en)<=0) then
                         VAttribute.vrValue = 1
                       else
                         VAttribute.vrValue = cint(en)
                       end if
                     else
                       VAttribute.vrValue = 1
                     end if
                   end if
                   if old_obj.vrAttrExists(VAttribute.vrName) then
                     VAttribute.vrAssignFrom(old_obj.vrAttrByName(VAttribute.vrName))
                   end if
                 'Для общего случая просто находим атрибут и копируем из него значение
                 'в объект новой модели.
                 elseif old_obj.vrAttrExists(VAttribute.vrName) then
                   'VAttribute.vrValue = old_obj.vrAttrByName(VAttribute.vrName).vrValue
                   VAttribute.vrAssignFrom(old_obj.vrAttrByName(VAttribute.vrName))
                 end if

               end if

             Next
           end if
           '''''''''''''''' Обновим значения атрибутов если модель младше V3. End ''''''''''''''''
           
           '''''''''''''''' Обновление с первой версии. Begin ''''''''''''''''''
              '1. Вытаскиваем из операции оборудование и профессию в отдельные объекты
              '2. Конвертикруем время назначенное на переходы и операции в секунды, т.к. базовая ед. изм. времени сек.
              '3. Заполняем атрибут detail.material и обновляем Location материала
              '4. Обновляем Location для всех объектов, т.к. в V1 хранился только идентификатор объекта
           If v_old < 4 Then
             'операции
             If old_obj.vrClass.vrClassVector.vrFilterVector.vrLocateConstraint("operations", old_obj.vrClass.vrName) Then
               'Обновим Location операции
               ChangeLocation old_obj, new_obj, "operid", "OPER.LIST="
 
               '1. Вытащим станок с профессией из атрибутов старой операции и 
               'добавим их как подчиненные объекты Оборудование->Профессия новой операции
               AddHardwareFromV1 old_obj, new_obj 
 
               'Вытащим из атрибутов операции СОЖ и добавим ее как объект к новой операции
               if old_obj.vrAttrExists("sogid") then
                 If old_obj.vrAttrByName("sogid").vrValue <> "" Then
                   set sog_obj = new_objs.vrCreate("sog")
                   sog_obj.vrAttrByName("name").vrValue = old_obj.vrAttrByName("sog").vrValue
                   sog_obj.vrAttrByName("id").vrValue = "SOG.MARKA=" & old_obj.vrAttrByName("sogid").vrValue
                   new_obj.vrAddChildLink(sog_obj)
                 End If
               end if
                           
               '2. Переводим время в секунды для операции (базовая единица)
               new_obj.vrAttrByName("tosn").vrValue = old_obj.vrAttrByName("tosn").vrValue * 60
               new_obj.vrAttrByName("tvspom").vrValue = old_obj.vrAttrByName("tvspom").vrValue * 60
               new_obj.vrAttrByName("timesht").vrValue = old_obj.vrAttrByName("timesht").vrValue * 60
               new_obj.vrAttrByName("timepz").vrValue = old_obj.vrAttrByName("timepz").vrValue * 60
             ElseIf old_obj.vrClass.vrClassVector.vrFilterVector.vrLocateConstraint("steps", old_obj.vrClass.vrName) Then
               '2. Переводим время в секунды переходов (базовая единица)
               new_obj.vrAttrByName("to").vrValue = old_obj.vrAttrByName("to").vrValue * 60
               new_obj.vrAttrByName("tv").vrValue = old_obj.vrAttrByName("tv").vrValue * 60
             End If

             '''''''''''''''' Некоторые атрибуты были переименованы '''''''''''''''''
             'оборудование
             If old_obj.vrClass.vrName = "stanok" then
               new_obj.vrAttrByName("obozn").vrValue = "MEX_PRSP.TYPESIZE=" & old_obj.vrAttrByName("model").vrValue
 
             '3. Заполняем атрибут detail.material и обновляем Location материала
             ElseIf old_obj.vrClass.vrName = "detail" then
               new_obj.vrAttrByName("material").vrValue = old_obj.vrAttrByName("sortament").vrValue & " " &_
                 old_obj.vrAttrByName("mainsize").vrValue & " " & old_obj.vrAttrByName("gostzagot").vrValue & " / " &_
                 old_obj.vrAttrByName("markamater").vrValue & " " & old_obj.vrAttrByName("gostmater").vrValue
               If old_obj.vrAttrByName("matrext").vrValue = "" Then
                  v_Location  = UniRefer.ConnectionList.ConnectServer.GetOneFieldSQL("SELECT GUID FROM ZAGOTOV_EXEMPLAR WHERE GUID_MATL = '" & old_obj.vrAttrByName("matrid").vrValue & "' AND GUID_SORTAM = '" & old_obj.vrAttrByName("sortamid").vrValue & "' ")
                  new_obj.vrAttrByName("matrid").vrValue = "ZAGOTOV.EXEMPLAR=" & v_Location
               End If
 
             '4. Меняем GUID на Location 
             'для КТЭ
             ElseIf old_obj.vrClass.vrName = "cte" then
               ChangeLocation old_obj, new_obj, "id", "KTE.LIST="
 
             'для переходов
             ElseIf old_obj.vrClass.vrName = "cnt_step" then
               ChangeLocation old_obj, new_obj, "ii_id", "II.TYPESIZE="
 
             ElseIf old_obj.vrClass.vrName = "mex_step" then
               ChangeLocation old_obj, new_obj, "id", "MEX_STEP.LIST="
 
             ElseIf old_obj.vrClass.vrName = "pok_step" then
               ChangeLocation old_obj, new_obj, "id", "MEX_STEP.LIST="
 
             ElseIf old_obj.vrClass.vrName = "mex_step" then
               ChangeLocation old_obj, new_obj, "id", "MEX_STEP.LIST="
 
             ElseIf old_obj.vrClass.vrName = "public_step" then
               ChangeLocation old_obj, new_obj, "id", "MEX_STEP.LIST="
 
             ElseIf old_obj.vrClass.vrName = "sbr_step" then
               ChangeLocation old_obj, new_obj, "id", "MEX_STEP.LIST="
 
             ElseIf old_obj.vrClass.vrName = "sht_step" then
               ChangeLocation old_obj, new_obj, "id", "MEX_STEP.LIST="
 
             ElseIf old_obj.vrClass.vrName = "sub_step" then
               ChangeLocation old_obj, new_obj, "id", "MEX_VSTP.LIST="
 
             ElseIf old_obj.vrClass.vrName = "svr_step" then
               ChangeLocation old_obj, new_obj, "id", "MEX_STEP.LIST="
 
             ElseIf old_obj.vrClass.vrName = "trm_step" then
               ChangeLocation old_obj, new_obj, "id", "MEX_STEP.LIST="
 
             'для всей оснастки
             ElseIf old_obj.vrClass.vrName = "fix_tool" then
               ChangeLocation old_obj, new_obj, "id", "MEX_PRSP.TYPESIZE="
 
             ElseIf old_obj.vrClass.vrName = "hand_tool" then
               ChangeLocation old_obj, new_obj, "id", "SLI.TYPESIZE="
 
             ElseIf old_obj.vrClass.vrName = "metrical_device" then
               ChangeLocation old_obj, new_obj, "id", "IZM_PRB.MARKA="
 
             ElseIf old_obj.vrClass.vrName = "metrical_tool" then
               ChangeLocation old_obj, new_obj, "id", "II.TYPESIZE="
 
             ElseIf old_obj.vrClass.vrName = "pok_tool" then
               ChangeLocation old_obj, new_obj, "id", "POK_OSN.MARKA="
 
             ElseIf old_obj.vrClass.vrName = "protect_tool" then
               ChangeLocation old_obj, new_obj, "id", "SIZ.MARKA="
 
             ElseIf old_obj.vrClass.vrName = "ri" then
               ChangeLocation old_obj, new_obj, "id", "RI.TYPESIZE="
 
             ElseIf old_obj.vrClass.vrName = "ri_blade" then
               ChangeLocation old_obj, new_obj, "id", "RI_BLADE.TYPESIZE="
 
             ElseIf old_obj.vrClass.vrName = "sbr_osnast" then
               ChangeLocation old_obj, new_obj, "id", "SBR_OSNT.MARKA="
 
             ElseIf old_obj.vrClass.vrName = "sbr_tool" then
               ChangeLocation old_obj, new_obj, "id", "SBR_TOOL.MARKA="
 
             ElseIf old_obj.vrClass.vrName = "sht_osnast" then
               ChangeLocation old_obj, new_obj, "id", "SHT_OSNT.MARKA="
 
             ElseIf old_obj.vrClass.vrName = "sht_tool" then
               ChangeLocation old_obj, new_obj, "id", "SHT_TOOL.MARKA="
 
             ElseIf old_obj.vrClass.vrName = "sog" then
               ChangeLocation old_obj, new_obj, "id", "SOG.MARKA="
 
             ElseIf old_obj.vrClass.vrName = "sub_material" then
               ChangeLocation old_obj, new_obj, "id", "VSP_MATR.MARKA="
 
             ElseIf old_obj.vrClass.vrName = "sub_tool" then
               ChangeLocation old_obj, new_obj, "id", "VI.TYPESIZE="
 
     '        ElseIf old_obj.vrClass.vrName = "svr_cable" then
     '		   ChangeLocation old_obj, new_obj, "id", "????="

     '        ElseIf old_obj.vrClass.vrName = "svr_electrod" then
     '          ChangeLocation old_obj, new_obj, "id", "????="

             ElseIf old_obj.vrClass.vrName = "svr_tool" then
               ChangeLocation old_obj, new_obj, "id", "SVR_TOOL.MARKA="

             ElseIf old_obj.vrClass.vrName = "trm_tool" then
               ChangeLocation old_obj, new_obj, "id", "TRM_OSNT.MARKA="

             ElseIf old_obj.vrClass.vrName = "truck_tool" then
               ChangeLocation old_obj, new_obj, "id", "GRZ_PRSP.TYPESIZE="

             End If

           End If
           '''''''''''''''' Обновление с первой версии. End ''''''''''''''''''

           '''''''''''''''' Обновление со второй версии. Begin '''''''''''''''
              'Перемещаем профессии из под операций к оборудованию в ТП версии V2.
              'В V2 профессии лежали только под операциями. В структуре текущией версии
              'профессии лежат под оборудованием, если оно указано.
           If (v_old > 3)and(v_old < 13) Then
              If old_FilterVector.vrLocateConstraint("operations", old_obj.vrClass.vrName) Then
                Set new_operation = new_obj
                Set old_operation = old_obj
                'Получаем оборудование новой операции
                Set VIterator = new_operation.vrObjectsVector.vrCreateIterator("equipment",new_operation,true)
                If VIterator.vrFirst Then 'если оборудование нашли то профессию будем добавлять в него
                  Set new_equipment = VIterator.vrGetObject
                  'Бежим по профессиям старой оперции и перемещаем их в оборудование новой операции
                  Set VIterator = old_operation.vrObjectsVector.vrCreateIterator("workers",old_operation,true)
                  While VIterator.vrNext
                    'Копируем профессию из старой операции в новое оборудование. 
                    'Если профессия может жить под операцией в новой модели, то перемещаем ее
                    'перекидывая связи
                    Set old_worker = VIterator.vrGetObject
                    set new_worker = new_objs.vrGetObjByStrID(old_worker.vrObjStrID)
                    if new_worker is nothing then
                      set new_worker = new_equipment.vrAddChild(old_worker.vrClass.vrName)
                      new_worker.vrAssignFrom(old_worker)
                    else
                      new_equipment.vrAddChildLink new_worker 
                      new_operation.vrDeleteChildLink new_worker
                    end if
                  Wend
                End if

              End If
           End If
           '''''''''''''''' Обновление со второй версии. End '''''''''''''''

           '''''''''''''''' Обновление со второй версии SP3. Begin '''''''''''''''
              '1. Конвертируем комплектование со V2 SP3 в новую версию
              '2. Переносим строку режимов резания в новый общий атрибут для всех режимов
           If (v_old < 16) Then  
             '1. Конвертируем комплектование операции
             if (must_convert_packages)and(old_objs.vrObjFitsFilter(old_obj,"operations")) then

               'Сбрасываем информацию о созданной комплектовки для новой операции
               redim opkg_id(0)
               redim opkg_obozn(0)
               redim opkg_name(0)
                                                 
               if (old_obj.vrClass.vrnChildClassItem("package") is nothing) then
                 set vsteps = old_objs.vrCreateIterator("steps", old_obj, true)
                 while vsteps.vrNext
                   set vstep_old = vsteps.vrGetObject
                   if not(vstep_old.vrClass.vrnChildClassItem("package") is nothing) then
                     if old_FilterVector.vrLocateFilter("packages") then
                       set vpackages = old_objs.vrCreateIterator("packages", vstep_old, true)
                     else
                       set vpackages = old_objs.vrCreateIterator("package", vstep_old, true)
                     end if
                     while vpackages.vrNext
                       AddPackageToOperation vpackages.vrGetObject, new_dse, new_obj
                     wend
                   end if
                 wend
               else
                 if old_FilterVector.vrLocateFilter("packages") then
                   set vpackages = old_objs.vrCreateIterator("packages", old_obj, true)
                 else
                   set vpackages = old_objs.vrCreateIterator("package", old_obj, true)
                 end if
                 while vpackages.vrNext
                   AddPackageToOperation vpackages.vrGetObject, new_dse, new_obj
                 wend
               end if

             end if

             '2. Обновляем режимы. Переносим строку режимов из старого атрибута в общий атрибут режимов
             if old_obj.vrClass.vrName="regrez" then
               new_obj.vrAttrByName("mode_string").vrValue = old_obj.vrAttrByName("rezstring").vrValue
             end if

           End If
           '''''''''''''''' Обновление со второй версии. End '''''''''''''''
           
           '''''''''''''''' Обновление с третей версии SP1. Begin '''''''''''''''
           'Начиная с V4 необходимо учитывать специфику ТТП/ГТП. Идентификаторы объектов должны
           'быть сохранены. Добавлении объектов в новую модель, необходимо обсудить с Поляковым Павлом
              '1. Конвертируем применяемость карт в операциях с предыдущих версий в новый формат хранения
              '2. Конвертируем контрольные операции и контрольные переходы в новую структуру
              '2.1 Преобразуем контрольные операции хранящиеся в public_oper в cnt_oper
              '2.2 Преобразуем контрольные переходы cnt_step хранящиеся в операциях общего назначения
              '    public_oper в переходы общего назначения. Актуально для ТП младше третей версий 
              '3. Скопировать значения атрибутов: Разработал, Проверил, Утвердил, Нормоконтроль из ДСЕ 
              '   во все операции
              '4. Конвертируем автонумерацию из старого формата в новый
           If v_old < 24 Then
             
             '1. Конвертируем применяемость карт для операций с предыдущих версий
             'Ранее к операции были добавлены карты она выводилась, после реализации
             'коллективной работы, механизм был изменен и карты стали хранится в атрибуте операции
             if new_objs.vrObjFitsFilter(new_obj, "operation") then
               ' неприменяемость карт. До V4 неприменяемость хранилась в виде связи operation->report
               ' теперь она хранится в атрибуте operation.unused_reports как IObjectSet
               if not (mdl_old.vrGetClassVector.vrLocate("report") is nothing) then
                 set it = old_objs.vrCreateIterator("report",old_obj,true)
                 set iobjset = new_obj.vrAttrByName("unused_reports").vrValue
                 do while it.vrNext
                   iobjset.vrAddStr it.vrGetObject.vrObjStrID 
                 loop
                end if ' есть в старом ТП класс report для импорта неприменяемости 
             end if             

             '2. Преобразуем контрольные операции из public_oper в cnt_oper
             'До версии V3, контрольные переходы(cnt_step) можно было добавлять в любую операцию
             'В V3, контролируемые параметры(cnt_step) можно было добавлять только в контрольные операции, проверялся код операции
             'Контрольная операция доя V4 хранилась в классе public_oper
             'Задача: Конвертнуть все контрольные операции из public_oper в cnt_oper, все контрольные 
             'переходы созданные до V3 требуется конвертнуть из класса cnt_step в public_step
             if (new_obj.vrClass.vrName = "public_oper") then

               set pub_oper = new_obj
               if pub_oper.vrAttrExists("kodoper") then 
                 kodoper = Left(pub_oper.vrAttrByName("kodoper").vrValue,2)
               else
                 kodoper = ""
               end if
               if pub_oper.vrAttrExists("nameoper") then 
                 nameoper = pub_oper.vrAttrByName("nameoper").vrValue
               else
                 nameoper = ""
               end if
               
               if (kodoper="02")or(kodoper="03") or (instr(1, nameoper, "КОНТРОЛ" , 1)>0) then
                 'копируем атрибуты операции
                 set cnt_oper = new_objs.vrCloneAs(pub_oper,"cnt_oper")
                 'копируем связи на подчиненные объекты
                 set childs = new_objs.vrCreateIterator("",pub_oper,true)
                 while childs.vrNext
                   cnt_oper.vrAddChildLink(childs.vrGetObject)
                 wend
                 
                 'Заменяем контрольную операцию во всех родительских объектах
                 redim preserve parent_id(0)
                 parent_cnt = 0
                 set parents = mdl_new.vrGetObjVector.vrCreateIterator("", pub_oper, false)
                 while parents.vrNext
                   redim preserve parent_id(parent_cnt+1)
                   parent_id(parent_cnt) = parents.vrGetObject.vrObjStrId
                   parent_cnt = parent_cnt + 1
                 wend
                 for k = 0 to parent_cnt - 1
                   set parent_obj = mdl_new.vrGetObjVector.vrGetObjByStrID(parent_id(k))
                   'Заменяем контрольную операцию с сохранение vrObjID
                   parent_obj.vrInsertChildLink cnt_oper,pub_oper 
                   parent_obj.vrDeleteChildLink pub_oper 
                 next
                 'После замены присавиваем ID старой операции Заменяем контрольную операцию с сохранение vrObjID
                 s_helper.vsCopyObjGUID pub_oper, cnt_oper
                 
                 'приводим в порядок контрольные переходы контрольной операции
                 set prev_new_obj = nothing
                 set old_childs = old_objs.vrCreateIterator("", old_obj, true)
                 while old_childs.vrPrev
                   if old_childs.vrGetObject.vrClass.vrName = "cnt_step" then
                     
                     set old_cnt_step = old_childs.vrGetObject
                     'Получаем контрольный переход из новой модели, 
                     'если его нет, значит он был только у операции
                     'и т.к. в новой структуре связи между public_oper и cnt_step нет, то и объекта 
                     'тоже нет. Если же объект есть значит он был добавлен еще к КТЭ и требуется его
                     'добавить к новой контрольной операции на прежнее место
                     set new_cnt_step = new_objs.vrGetObjByStrID(old_cnt_step.vrObjStrID)
                     if new_cnt_step is nothing then
                       set new_cnt_step = new_objs.vrCloneAs(old_cnt_step, "cnt_step")
                       'set new_cnt_step = new_objs.vrCreate("cnt_step")
                       'new_cnt_step.vrAssignFrom(old_cnt_step)
                       'Сохраняем ObjID для перехода
                       s_helper.vsCopyObjGUID old_cnt_step, new_cnt_step

                       if prev_new_obj is nothing then
                         cnt_oper.vrAddChildLink new_cnt_step
                       else
                         cnt_oper.vrInsertChildLink new_cnt_step, prev_new_obj
                       end if
    
                       'Копируем подчиненные объекты
                       set old_cnt_step_childs = old_objs.vrCreateIterator("", old_cnt_step, true)
                       while old_cnt_step_childs.vrNext
                         set old_cnt_step_child = old_cnt_step_childs.vrGetObject
                         set new_cnt_step_child = new_objs.vrCloneAs(old_cnt_step_child, old_cnt_step_child.vrClass.vrName)
                         'set new_cnt_step_child = new_objs.vrCreate(old_cnt_step_child.vrClass.vrName)
                         'new_cnt_step_child.vrAssignFrom(old_cnt_step_child)
                         new_cnt_step.vrAddChildLink new_cnt_step_child
                       wend
                     else
                       if prev_new_obj is nothing then
                         cnt_oper.vrAddChildLink new_cnt_step
                       else
                         cnt_oper.vrInsertChildLink new_cnt_step, prev_new_obj
                       end if
                     end if  
                   
                     'Если ТП младше третей версии, то вытащим ИИ из контрольного перехода
                     'т.к. до третей версии ИИ лежал внутри контрольного перехода
                     If v_old < 16 Then
                       'Из старого контролируемого параметра контрольной операции вытаскиваем данные
                       'по ИИ и добавляем как измерительный инструмент
                       if cstr(old_cnt_step.vrAttrByName("ii_id").vrValue)<>"" then
                         set metrical_tool = new_cnt_step.vrAddChild("metrical_tool")
                         metrical_tool.vrAttrByName("id").vrValue = old_cnt_step.vrAttrByName("ii_id").vrValue
                         metrical_tool.vrAttrByName("name").vrValue = old_cnt_step.vrAttrByName("ii_name").vrValue
                         metrical_tool.vrAttrByName("obozn").vrValue = old_cnt_step.vrAttrByName("ii_obozn").vrValue
                         metrical_tool.vrAttrByName("gost").vrValue = old_cnt_step.vrAttrByName("ii_gost").vrValue
                       end if
                     end if
                   
                     'Пробуем сконвертировать значение старого строкового атрибута v_pk
                     'в новый атрибут factor_period с типом данных float
                     if (old_cnt_step.vrAttrExists("v_pk")) then
                       v_pk_new = ""
                       v_pk_old = trim(old_cnt_step.vrAttrByName("v_pk").vrValue)
                       n = len(v_pk_old)
                       for j = 1 to n
                         nextchar = mid(v_pk_old, j, 1)
                         if instr("0123456789,.", nextchar)>0 then
                           v_pk_new = v_pk_new & nextchar
                         end if
                       next
                       v_pk_new = Replace(v_pk_new, ",", ".")
                       new_cnt_step.vrAttrByName("factor_period").vrValue = cdbl(v_pk_new)
                     end if
                     
                   end if   
                    
                   'если в новой модели не нашли аналогичный объект, то сохраняем тот что получили ранее
                   set test_prev = new_objs.vrGetObjByStrID(old_childs.vrGetObject.vrObjStrID) 
                   if not(test_prev is nothing) then
                     set prev_new_obj = test_prev
                   end if
                   
                 wend
               
                 'Т.к. заменили операцию на контрольную, то обновляем и переменню new_obj
                 set new_obj = cnt_oper
               else
                 'Если ТП младше третей версии, то преобразуем контрольные переходы
                 'не контрольной операции в переходы общего назначения
                 If v_old < 16 Then
                   
                   'Переберем все подчиненные объекты операции общего назначения старой модели
                   'в обратном порядке и обработаем каждый встреченный нами контрольный переход
                   set prev_new_obj = nothing
                   set old_childs = old_objs.vrCreateIterator("", old_obj, true)
                   while old_childs.vrPrev
                     set new_substep = nothing
                     if old_childs.vrGetObject.vrClass.vrName = "cnt_step" then
                       set old_cnt_step = old_childs.vrGetObject

                       'Преобразуем старый контролируемый параметр во вспомогательный переход
                       'с подчиненным измерительным инструментом
                       'Создаем вспомогательный переход и копируем данные из старого контролируемого параметра
                       set new_substep = new_objs.vrCreate("sub_step")
                       new_substep.vrAttrByName("name").vrValue=old_cnt_step.vrAttrByName("name").vrValue
                       new_substep.vrAttrByName("to").vrValue=old_cnt_step.vrAttrByName("to").vrValue
                       new_substep.vrAttrByName("tv").vrValue=old_cnt_step.vrAttrByName("tv").vrValue
                       new_substep.vrAttrByName("comments").vrValue=old_cnt_step.vrAttrByName("comments").vrValue

                       'Добавляем вспомогательный переход к операции
                       if prev_new_obj is nothing then
                         new_obj.vrAddChildLink new_substep 
                       else
                         new_obj.vrInsertChildLink new_substep, prev_new_obj
                       end if

                       'Пробуем найти контрольный переход в новой модели. Он там может 
                       'остаться если он был добавлен еще и к КТЭ
                       set new_cnt_step = new_objs.vrGetObjByStrID(old_cnt_step.vrObjStrID)
                       'если объект котрольного перехода в новой модели был найден, то
                       'заменяем его у всех родителей на новый вспомогательный переход
                       if not(new_cnt_step is nothing) then
                         'Добавляем вспомогательный переход ко всем родителям контролируемого параметра
                         set new_parents = new_objs.vrCreateIterator("",new_cnt_step,false)
                         while new_parents.vrNext
                           new_parents.vrGetObject.vrInsertChildLink new_substep, new_cnt_step
                         wend
                         'Удаляем контролируемый параметр от всех радителей
                         while new_parents.vrFirst
                           new_parents.vrGetObject.vrDeleteChildLink new_cnt_step
                         wend
                       end if
                       
                       'Теперь мы уверены что контрольного перехода в новой модели не осталось и
                       'копируем ObjID со старого контрольного перехода
                       s_helper.vsCopyObjGUID old_cnt_step, new_cnt_step
    
                       'Копируем подчиненные объекты
                       set old_cnt_step_childs = old_objs.vrCreateIterator("", old_cnt_step, true)
                       while old_cnt_step_childs.vrNext
                         set old_cnt_step_child = old_cnt_step_childs.vrGetObject
                         set new_sub_step_child = new_objs.vrCloneAs(old_cnt_step_child, old_cnt_step_child.vrClass.vrName)
                         'set new_sub_step_child = new_objs.vrCreate(old_cnt_step_child.vrClass.vrName)
                         'new_sub_step_child.vrAssignFrom(old_cnt_step_child)
                         new_substep.vrAddChildLink new_sub_step_child
                       wend
                       'Копируем добавляем измерительный инструмент
                       if cstr(old_cnt_step.vrAttrByName("ii_id").vrValue)<>"" then
                         set metrical_tool = new_substep.vrAddChild("metrical_tool")
                         metrical_tool.vrAttrByName("id").vrValue = old_cnt_step.vrAttrByName("ii_id").vrValue
                         metrical_tool.vrAttrByName("name").vrValue = old_cnt_step.vrAttrByName("ii_name").vrValue
                         metrical_tool.vrAttrByName("obozn").vrValue = old_cnt_step.vrAttrByName("ii_obozn").vrValue
                         metrical_tool.vrAttrByName("gost").vrValue = old_cnt_step.vrAttrByName("ii_gost").vrValue
                       end if
                       
                     end if
                     
                     'если в новой модели не нашли аналогичный объект, то сохраняем тот что получили ранее
                     set test_prev = new_objs.vrGetObjByStrID(old_childs.vrGetObject.vrObjStrID) 
                     if not(test_prev is nothing) then
                       set prev_new_obj = test_prev
                     end if
                   wend
                 end if
               end if
             
             end if
             
             '3. Копируем атрибуты Разработал, Проверил, Утвердил, Нормоконтроль с ДСЕ в операции
             if new_objs.vrObjFitsFilter(new_obj, "operation") then
               'Для старых ТП информация о: Разработал, Утвердил, Проверил, Нормоконтроль была в ДСЕ
               'Теперь, в V4, она также хранится и в операции, поэтому копируем ее из ДСЕ в операции
               'чтобы при формировании карт ОК выводилась таже инфа что и раноьше
               new_obj.vrAttrByName("data").vrValue = old_dse.vrAttrByName("data").vrValue
               new_obj.vrAttrByName("audittp").vrValue = old_dse.vrAttrByName("audittp").vrValue
               new_obj.vrAttrByName("data_audittp").vrValue = old_dse.vrAttrByName("data_audittp").vrValue
               new_obj.vrAttrByName("controltp").vrValue = old_dse.vrAttrByName("controltp").vrValue
               new_obj.vrAttrByName("datacontroltp").vrValue = old_dse.vrAttrByName("datacontroltp").vrValue
               new_obj.vrAttrByName("btk").vrValue = old_dse.vrAttrByName("btk").vrValue
               new_obj.vrAttrByName("data_btk").vrValue = old_dse.vrAttrByName("data_btk").vrValue
               new_obj.vrAttrByName("ncontrol").vrValue = old_dse.vrAttrByName("ncontrol").vrValue
               new_obj.vrAttrByName("data_ncontrol").vrValue = old_dse.vrAttrByName("data_ncontrol").vrValue
             end if
           
             '4. Конвертим автонумерацию из старой модели в новую
             CLASS_Autoinc = "{E7223F66-4BCC-4439-817D-41B9454DE261}"
             set VPlugins = mdl_new.vrGetPlugins
             set AutoNums = VPlugins.vrAddByStrGUID(CLASS_Autoinc)
             AutoNums.vsImportFromModel mdl_old
           
           End If
           '''''''''''''''' Обновление со третей версии SP1. End '''''''''''''''

           '''''''''''''''' Обновление с четвертой версии SP1. Begin '''''''''''''''
           'Начиная с V4 необходимо учитывать специфику ТТП/ГТП. Идентификаторы объектов должны
           'быть сохранены. Добавлении объектов в новую модель, необходимо обсудить с Поляковым Павлом
           '   1. Обновляем время в операциях и переходах для ТП созданных в версии V4 и V4SP1
           '   2. Заменяем идентификатор "Кода вида производства"(code_classwork) на двухзначный код
           '   3. Обновляем значения в атрибутах заявки, т.к. появился базовый класс requestunit
           '   4. Обновляем карты нормирвоания в ТП
           '   5. Конвертируем все операции в ЧПУ у которых заполнен атрибут cnc_file. 
           '      Операции ТТП/ГТП игнорируем
           '   6. Обновляем ТТП/ГТП, делаем наследуемыми в ЕТП все объекты для классов которых
           '      была снята галка в фильтре notinheritedclasses
           If v_old < 27 Then
             
             '1. Обновляем времена в ДСЕ, операциях и переходах только для ТП V4. 
             'В версии V4 эти атрибуты были вычисляемыми, в текущей версии эти атрибуты опять стали обычными. 
             if (new_objs.vrObjFitsFilter(new_obj, "dseunit")) and (v_old > 16) and (old_obj.vrAttrExists("timesht")) then
                 set op_iter = mdl_old.vrGetObjVector.vrCreateIterator("operation", old_obj, true) 
                 if op_iter <> null then
                    new_obj.vrAttrByName("timesht").vrValue = old_obj.vrAttrByName("timesht").vrValue
                    'new_obj.vrAttrByName("tshtk").vrValue = old_obj.vrAttrByName("tshtk").vrValue
                 end if
             end if
             'Обновляем время в операциях
             if new_objs.vrObjFitsFilter(new_obj, "operation") and (v_old > 16) then
               new_obj.vrAttrByName("timesht").vrValue = old_obj.vrAttrByName("timesht").vrValue
               new_obj.vrAttrByName("timepz").vrValue = old_obj.vrAttrByName("timepz").vrValue
               new_obj.vrAttrByName("tshtk").vrValue = old_obj.vrAttrByName("tshtk").vrValue
               new_obj.vrAttrByName("tosn").vrValue = old_obj.vrAttrByName("tosn").vrValue
               new_obj.vrAttrByName("tvspom").vrValue = old_obj.vrAttrByName("tvspom").vrValue
             end if
             'Обновляем время в переходах
             if new_objs.vrObjFitsFilter(new_obj, "step") and (v_old > 16) then
               new_obj.vrAttrByName("to").vrValue = old_obj.vrAttrByName("to").vrValue
               new_obj.vrAttrByName("tv").vrValue = old_obj.vrAttrByName("tv").vrValue
             end if
            
             '2. Теперь в атирбут dse.code_classwork пишем не гуид а двухзначный код
             'поэтому заменяем гуид на соответствующий код
             if new_objs.vrObjFitsFilter(new_obj, "dseunit") and new_obj.vrAttrExists("code_classwork")  then
               code_classwork = new_obj.vrAttrByName("code_classwork").vrValue
               if (code_classwork<>"") then
                 set utils = CreateObject("Ascon.Vertical.TransitionPeriodUtils")
                 code_classwork = utils.GetClassWorkCode(code_classwork)
                 if code_classwork <> "" then new_obj.vrAttrByName("code_classwork").vrValue = code_classwork
               end if
             end if
             
             '3. Обновляем атрибуты заявки т.к. начиная с V5 у нее появился базовый класс
             if old_obj.vrClass.vrName = "request" then
               new_obj.vrAttrByName("designation").vrValue = old_obj.vrAttrByName("designation").vrValue
               new_obj.vrAttrByName("status").vrValue = old_obj.vrAttrByName("status").vrValue
             end if
              
             '4. Обновляем карты нормирвоания в ТП
             if not(new_obj.vrClass.vrnChildClassItem("norm_map") is nothing) then
               'Обновляем norm_map
               if not(old_obj.vrClass.vrnChildClassItem("norm_map") is nothing) then
                 set old_norm_maps = old_objs.vrCreateIterator("norm_map", old_obj, true)
                 while old_norm_maps.vrNext 
                   set old_norm_map = old_norm_maps.vrGetObject
                   set new_norm_map = new_objs.vrGetObjByStrID(old_norm_map.vrObjStrID)
                   if not(new_norm_map is nothing) then
                     new_norm_map.vrAttrByName("norm_version").vrValue = "4.0"
                     select case old_norm_map.vrAttrByName("kart_type").vrValue
                       case 1 new_norm_map.vrAttrByName("kart_type").vrValue = "N41hnqAOjO5NJYOl3bmfJa"
                       case 2 new_norm_map.vrAttrByName("kart_type").vrValue = "06DtJtCUWmOME.TrMPfFTb"
                       case 3 new_norm_map.vrAttrByName("kart_type").vrValue = "IYWxssdNquptoIHfcVH4Ya"
                       case 4 new_norm_map.vrAttrByName("kart_type").vrValue = "QIwTjrmsUW.mKfPjyg6jdc"
                     end select 
                     'добавляем параметры norm_attr из norm_map.allparams
                     AddNormAttrs new_norm_map
                   end if
                 wend
               end if

               'Ковертируем step_map в norm_map
               if not(old_obj.vrClass.vrnChildClassItem("step_map") is nothing) then
                 set old_step_maps = old_objs.vrCreateIterator("step_map", old_obj, true)
                 while old_step_maps.vrNext           
                   set old_step_map = old_step_maps.vrGetObject
                   set new_norm_map = new_obj.vrAddChild("norm_map") 
                   new_norm_map.vrAttrByName("kartname").vrValue = old_step_map.vrAttrByName("name").vrValue
                   new_norm_map.vrAttrByName("norm_value").vrValue = old_step_map.vrAttrByName("value").vrValue
                   new_norm_map.vrAttrByName("allparams").vrValue = old_step_map.vrAttrByName("history").vrValue
                   select case old_step_map.vrAttrByName("type").vrValue
                     case 0 new_norm_map.vrAttrByName("kart_type").vrValue = "06DtJtCUWmOME.TrMPfFTb"
                     case 1 new_norm_map.vrAttrByName("kart_type").vrValue = "QIwTjrmsUW.mKfPjyg6jdc"
                     case 2 new_norm_map.vrAttrByName("kart_type").vrValue = "N99V.EvLhDJxsVGFBERNQc"
                     case 3 new_norm_map.vrAttrByName("kart_type").vrValue = "r7PaqcF9m5PzXk6IziaM.d"
                     case 4 new_norm_map.vrAttrByName("kart_type").vrValue = ".vEQXDOrfnM_yvHbd15A.d"
                     case 5 new_norm_map.vrAttrByName("kart_type").vrValue = "ehYp4aNjhdU2UJEFpHAEQc"
                   end select 
                   'Там какие-то гуиды, поэтому ничего не делаем
                   'добавляем параметры norm_attr из norm_map.allparams
                   'AddNormAttrs new_norm_map
                 wend
               end if
             end if
             
             'Получим плагин для работы с ТТП/ГТП
             set VPlugins = mdl_new.vrGetPlugins
             set TTPDocument = nothing
             set TTPDocument = VPlugins.vrItemByStrGUID(CLASS_TTPDocument)

             '5. Конвертируем все операции в ЧПУ у которых заполнен атрибут cnc_file,
             '   если обновляемый ТП не является ТТП/ГТП
             if (TTPDocument is nothing)and(new_objs.vrObjFitsFilter(new_obj, "operation"))and(old_obj.vrAttrExists("cnc_file")) then
               if old_obj.vrAttrByName("cnc_file").vrFile.vsInternalFullName <> "" then
                 'Создадим операцию ЧПУ
                 set new_cnc_oper = new_objs.vrCloneAs(new_obj, "cnc_oper")
                 'set new_cnc_oper = new_objs.vrCreate("cnc_oper")
                 'new_cnc_oper.vrAssignFrom(new_obj)
                 new_cnc_oper_id = new_obj.vrObjStrID

                 'Добавим операцию ко всем родителям
                 set parent_it = new_obj.vrObjectsVector.vrCreateIterator("", new_obj, false)
                 while parent_it.vrNext 
                   set new_parent = parent_it.vrGetObject
                   'Добавим операцию ЧПУ
                   new_parent.vrInsertChildLink new_cnc_oper, new_obj 
                 wend

                 'Добавим программу ЧПУ
                 set new_placement = new_cnc_oper.vrAddChild("placement")
                 new_placement.vrAttrByName("obozn").vrValue = "Установ 1"
                 set program_nc = new_placement.vrAddChild("program_nc")
                 set new_cnc_doc = program_nc.vrAddChild("documents")
                 new_cnc_doc.vrAttrByName("file").vrFile.vsDiskFullName = old_obj.vrAttrByName("cnc_file").vrFile.vsInternalFullName
                 new_cnc_doc.vrAttrByName("caption").vrValue = "Программа ЧПУ"
               
                 'Скопируем содержимое конвертируемой операции в операцию ЧПУ
                 set child_it = new_obj.vrObjectsVector.vrCreateIterator("", new_obj, true)
                 while child_it.vrNext 
                   set child_obj = child_it.vrGetObject
                   'Содержимое копируем либо к операции либо к установу, в зависимости от связей
                   if not(new_cnc_oper.vrClass.vrnChildClassItem(child_obj.vrClass.vrName) is nothing) then
                     new_cnc_oper.vrAddChildLink(child_obj)
                   elseif not(new_placement.vrClass.vrnChildClassItem(child_obj.vrClass.vrName) is nothing) then
                     new_placement.vrAddChildLink(child_obj)
                   end if
                 wend
                 
                 'Удалим старую операцию от всех родителей
                 set parent_it = new_obj.vrObjectsVector.vrCreateIterator("", new_obj, false)
                 while parent_it.vrNext 
                   set new_parent = parent_it.vrGetObject
                   'Удалим операцию
                   new_parent.vrDeleteChildLink new_obj 
                 wend
                 'Сохраняем ObjID старой операции в новую операцию ЧПУ(cnc_oper)
                 s_helper.vsCopyObjGUID new_obj, new_cnc_oper                 
               end if
             end if

             '6. Обновляем ТТП/ГТП, делаем наследуемыми в ЕТП все объекты для классов которых
             'была снята галка в фильтре notinheritedclasses
             if not(TTPDocument is nothing) then
               TTPDocument.UpdateTPFromV4SP1 
             end if
             
           end if
           '''''''''''''''' Обновление с четвертой версии SP1. End '''''''''''''''
       End If
  Next
  'AdminMessage "ТП обновлен до версии 2.0.1"

  s_helper.vsAllowChangesByEvents mdl_new, false
  s_helper.vsAllowChangesByEvents mdl_old, false
End Function

'Получить из операции ТП V1 объект професии
function GetProfFromV1(OperV1, new_objs)
  'вернем объект профессии если в нем есть хотя бы codprof или classjob
  if (OperV1.vrAttrByName("codprof").vrValue<>"")or(OperV1.vrAttrByName("classjob").vrValue<>"") then
    set GetProfFromV1 = new_objs.vrCreate("worker")
    GetProfFromV1.vrAttrByName("name").vrValue = OperV1.vrAttrByName("nameprof").vrValue
    GetProfFromV1.vrAttrByName("code").vrValue = OperV1.vrAttrByName("codprof").vrValue
    GetProfFromV1.vrAttrByName("classjob").vrValue = OperV1.vrAttrByName("classjob").vrValue
    GetProfFromV1.vrAttrByName("cm").vrValue = OperV1.vrAttrByName("cm").vrValue
    GetProfFromV1.vrAttrByName("yt").vrValue = OperV1.vrAttrByName("yt").vrValue
    GetProfFromV1.vrAttrByName("kr").vrValue = OperV1.vrAttrByName("kr").vrValue
  else
    set GetProfFromV1 = nothing
  end if
end function

function IsMustHaveHardware(OperClassName) 
  'Определяем массив операций для которых оборудование не обязательно
  dim oper_without_hardware(3)
  oper_without_hardware(0) = "cnt_oper"
  oper_without_hardware(1) = "sbr_oper"
  oper_without_hardware(2) = "pok_oper"
  oper_without_hardware(3) = "public_oper"

  'Проверям должна ли операция обязательно иметь оборудование
  IsMustHaveHardware = true
  while (i<=ubound(oper_without_hardware,1))and(IsMustHaveHardware) 
    IsMustHaveHardware = oper_without_hardware(i) <> OperClassName
    i = i + 1
  wend
end function

'Получить из операции ТП V1 объект станок
sub AddHardwareFromV1(OperV1, NewOper)
  
  'Получаем имя класса операции
  OperClsName = OperV1.vrClass.vrName

  'Проверям должна ли операция обязательно иметь оборудование
  MustHaveHardware = IsMustHaveHardware(OperV1.vrClass.vrName)
  
  'Получаем профессию из операции
  set Prof = GetProfFromV1(OperV1, NewOper.vrObjectsVector)
  'вернем объект станок если ключевой атрибут equipmentid заполнен
  'или если имеется профессия указанная на операцию у которого обязательно 
  'должно быть оборудование. Тогда путь создаться оборудование пустышка, 
  'чтобы добавить профессию
  if (OperV1.vrAttrByName("stanokid").vrValue<>"") then
    'Определим класс оборудования который надо создать
    If OperV1.vrClass.vrName="mex_oper" Then
      StanokClsName = "stanok"
    ElseIf OperV1.vrClass.vrName = "pok_oper" Then
      StanokClsName = "pok_hardware"
    ElseIf OperV1.vrClass.vrName = "sbr_oper" Then
      StanokClsName = "sbr_hardware"
    ElseIf OperV1.vrClass.vrName = "sht_oper" Then
      StanokClsName = "sht_hardware"
    ElseIf OperV1.vrClass.vrName = "svr_oper" Then
      StanokClsName = "svr_hardware"
    ElseIf OperV1.vrClass.vrName = "trm_oper" Then
      StanokClsName = "trm_hardware"
    else
      StanokClsName = ""
    End if

    'Добавим объект оборудования
    if (StanokClsName <> "") then
      set NewHardware = NewOper.vrObjectsVector.vrCreate(StanokClsName)
      NewHardware.vrAttrByName("Obozn").vrValue = OperV1.vrAttrByName("equipment1").vrValue
      NewHardware.vrAttrByName("id").vrValue = "STANOK.MODEL=" & OperV1.vrAttrByName("stanokid").vrValue
      if not (Prof is nothing) then
        NewHardware.vrAddChildLink(Prof)
      end if
      NewOper.vrAddChildLink(NewHardware)
    end if
  
  'Если оборудования нет и операция позволяет быть ей без оборудования, 
  'то добавляем профессию прям в операцию
  elseif (not MustHaveHardware)and(not(Prof is nothing)) then
     NewOper.vrAddChildLink(Prof)
  end if

end sub

'Преобразовать строку параметров в объекты класса norm_attr
'функция используется для обновления трудового нормирования версии V4 и V4SP1
sub AddNormAttrs(new_norm_map)
  'Формат строки: Имя[Переменная](Идентификатор)=Значение(Строка:Столбец);...
  allparams_str = new_norm_map.vrAttrByName("allparams").vrValue
  if allparams_str <> "" then
    Start = 1
    'Находимся в цикле пока не дойдем до конца в allparams_str
    while Start > 0
      set norm_attr = new_norm_map.vrAddChild("norm_attr")
      'Копируем наименование
      DelimPos = instr(Start, allparams_str, "[", 0)
      norm_attr.vrAttrByName("attr_caption").vrValue = mid(allparams_str, Start, DelimPos-Start)
      'Копируем имя переменной
      Start = DelimPos + 1
      DelimPos = instr(Start, allparams_str, "]", 0)
      norm_attr.vrAttrByName("attr_name").vrValue = mid(allparams_str, Start, DelimPos-Start)
      'Копируем идентификатор
      Start = DelimPos + 2
      DelimPos = instr(Start, allparams_str, ")", 0)
      norm_attr.vrAttrByName("attr_id").vrValue = mid(allparams_str, Start, DelimPos-Start)
      'Копируем значение
      Start = DelimPos + 2
      DelimPos = instr(Start, allparams_str, "(", 0)
      norm_attr.vrAttrByName("attr_value").vrValue = mid(allparams_str, Start, DelimPos-Start)
      'Получаем местоположение разделителя параметров и встаем на первый символ следующего параметра
      Start = instr(Start, allparams_str, ";", 0)
      if Start = Len(allparams_str) then
        Start = 0
      end if
      if Start > 0 then
        Start = Start + 1
      end if
    wend
  end if
end sub

Function GetDSEObject(vmodel)
  '$DEFINE vmodel as IVModel
  set vdse = vmodel.vrGetObjVector.vrCreateIterator("dse",vmodel.vrGetObjVector.vrItem(0),true)
  if vdse.vrFirst then
    set GetDSEObject = vdse.vrGetObject
  else
    set GetDSEObject = nothing
  end if
End Function

Sub ChangeLocation(old_obj, new_obj, name_attr, prefix)
	If (old_obj.vrAttrByName(name_attr).vrValue <>"") Then	  
		new_obj.vrAttrByName(name_attr).vrValue = prefix & old_obj.vrAttrByName(name_attr).vrValue
	End If
End Sub

Function AdminMessage(v_mess)
        Set UniRefer = GetObject(v_zero, "UniReference.UniRefer")
        If UniRefer Is Nothing Then Exit Function
        Set Logon = UniRefer.GlobalVars.Logon
        If Logon is nothing Then Exit Function
        If Not Logon.FlagLogon Then Exit Function
        If Logon.CheckPrivilegeGroup(Logon.IDGoupUser , "administration") Then
          msgbox v_mess
        End If
End Function

Function vrUpdateObj(obj_old, obj_new)
End Function

sub AddPackageToOperation(package_old, dse_new, operation_new)
  '$DEFINE operation_old as IVObject, operation_new as IVObject, dse_new as IVObject

  'Получаем ключевые данные по старой комплектовке
  loodsman_type = Trim(LCase(package_old.vrAttrByName("loodsman_type").vrValue))
  obozn = package_old.vrAttrByName("obozn").vrValue
  name = package_old.vrAttrByName("name").vrValue

  'Ищем комплектовку в комплектование СЕ и если ее нет то создаем
  has_apkg = false
  i=0
  while (not has_apkg)and(i<ubound(apkg_obozn,1))
    has_apkg = (StrComp(apkg_obozn(i),obozn,1)=0)and((StrComp(apkg_name(i),name,1)=0))
    if not has_apkg then
      i = i+1
    end if
  wend
  if (not has_apkg) then
    i = ubound(apkg_id,1)
    redim preserve apkg_id(i+1)
    redim preserve apkg_obozn(i+1)
    redim preserve apkg_name(i+1)

    if (loodsman_type="cборочная единица") then
      set apkg_obj = dse_new.vrAddChild("apkg_assembly")
    elseif (loodsman_type="деталь") then
      set apkg_obj = dse_new.vrAddChild("apkg_detail")
    elseif ((loodsman_type="материал по кд")or(loodsman_type="материал вспомогательный")) then
      set apkg_obj = dse_new.vrAddChild("apkg_material")
      apkg_obj.vrAttrByName("norma").vrMeasureUnit="VBB4733768E3C438997F2A7AC9182FBF0"
    elseif (obozn<>"") then
      set apkg_obj = dse_new.vrAddChild("apkg_detail")
    else
      set apkg_obj = dse_new.vrAddChild("apkg_material")
    end if
    apkg_obj.vrAttrByName("pos").vrValue = package_old.vrAttrByName("pos").vrValue
    apkg_obj.vrAttrByName("obozn").vrValue = obozn
    apkg_obj.vrAttrByName("name").vrValue = name
    apkg_obj.vrAttrByName("id_pdm").vrValue = package_old.vrAttrByName("id_pdm").vrValue
    apkg_obj.vrAttrByName("loodsman_product").vrValue = package_old.vrAttrByName("loodsman_product").vrValue
    apkg_obj.vrAttrByName("loodsman_version").vrValue = package_old.vrAttrByName("loodsman_version").vrValue
    apkg_obj.vrAttrByName("bo_location").vrValue = package_old.vrAttrByName("matrext").vrValue

    'Обнуляем норму и кол-во
    if apkg_obj.vrAttrExists("norma") then
      apkg_obj.vrAttrByName("norma").vrMeasureUnit="VBB4733768E3C438997F2A7AC9182FBF0"
      apkg_obj.vrAttrByName("norma").vrValue = 0
    else
      apkg_obj.vrAttrByName("ki").vrValue = 0
    end if

    apkg_id(i) = apkg_obj.vrObjStrID
    apkg_obozn(i) = obozn
    apkg_name(i) = name
  else
    set apkg_obj = dse_new.vrObjectsVector.vrGetObjByStrID(apkg_id(i))
  end if

  'Ищем комплектовку в комплектование операции и если ее нет то создаем, а если
  'есть то увеличиваем норму или колв-во как у комплектования СЕ так и у комплектования операции
  has_opkg = false
  i=0
  while (not has_opkg)and(i<ubound(opkg_obozn,1))
    has_opkg = (StrComp(opkg_obozn(i),obozn,1)=0)and((StrComp(opkg_name(i),name,1)=0))
    if not has_opkg then
      i = i+1
    end if
  wend
  if (not has_opkg) then
    i = ubound(opkg_id,1)
    redim preserve opkg_id(i+1)
    redim preserve opkg_obozn(i+1)
    redim preserve opkg_name(i+1)

    'Создаем новую комплектовку для операции
    set package_new = operation_new.vrAddChild(replace(apkg_obj.vrClass.vrName,"apkg","pkg"))
    package_new.vrAttrByName("opp").vrValue = package_old.vrAttrByName("opp").vrValue
    package_new.vrAttrByName("singlepod").vrValue = package_old.vrAttrByName("singlepod").vrValue
    package_new.vrAttrByName("generalpod").vrValue = package_old.vrAttrByName("generalpod").vrValue
    package_new.vrAttrByName("steppod").vrValue = package_old.vrAttrByName("steppod").vrValue
    'Копируем кол-во/норму в новую комплектовку операции и увеличиваем значение кол-ва/нормы
    'у комплектовки СЕ
    if (package_new.vrClass.vrName="pkg_material") then
      norma = package_old.vrAttrByName("norma").vrValue
      apkg_obj.vrAttrByName("norma").vrValue = apkg_obj.vrAttrByName("norma").vrValue + norma
      package_new.vrAttrByName("norma").vrMeasureUnit="VBB4733768E3C438997F2A7AC9182FBF0"
      package_new.vrAttrByName("norma").vrValue = norma
    else
      ki = package_old.vrAttrByName("ki").vrValue
      apkg_obj.vrAttrByName("ki").vrValue = apkg_obj.vrAttrByName("ki").vrValue + ki
      package_new.vrAttrByName("ki").vrValue = ki
    end if

    opkg_id(i) = package_new.vrObjStrID
    opkg_obozn(i) = obozn
    opkg_name(i) = name

    apkg_obj.vrAddChildLink package_new
  else
    'Комплектовка уже была создана, поэтому увеличиваем кол-во/норму у существующей
    'комплектовки у операции и у комплектовки СЕ
    set package_new = dse_new.vrObjectsVector.vrGetObjByStrID(opkg_id(i))
    if (package_new.vrClass.vrName="pkg_material") then
      norma = package_old.vrAttrByName("norma").vrValue
      apkg_obj.vrAttrByName("norma").vrValue = apkg_obj.vrAttrByName("norma").vrValue + norma
      package_new.vrAttrByName("norma").vrValue = package_new.vrAttrByName("norma").vrValue + norma
    else
      ki = package_old.vrAttrByName("ki").vrValue
      apkg_obj.vrAttrByName("ki").vrValue = apkg_obj.vrAttrByName("ki").vrValue + ki
      package_new.vrAttrByName("ki").vrValue = package_new.vrAttrByName("ki").vrValue + ki
    end if
  end if

End sub