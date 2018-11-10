Dim MeasureUnit 
MeasureUnit = "VDF643282616547778305DB3EDA1BE6A2" ' Длина, дециметры 
set MeasurementServer = nothing
set MeasurementServer = CreateObject("MeasurementServer.CoMeasurementServer")
if not(MeasurementServer is nothing) then
  set MUnit = MeasurementServer.MUnit(MeasureUnit)
  if not(MUnit is nothing) then
      MsgBox(CStr(MUnit.okeiCode))
  end if
end if