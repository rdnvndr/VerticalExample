' ������������ � ������������� �����������
SET m_uniref = CreateObject("UniReference.UniRefer")
if not m_uniref.GlobalVars.Logon.LogonAsParams("��������","111","���������") then
  MsgBox("����������� �� �����������")
end if

data = InputBox("������� ������ ��� ����������: ")
pass = InputBox("������� ������ ��� ����������: ")
encryptText = m_uniref.GlobalVars.Encrypt(data, pass) 
decryptText = m_uniref.GlobalVars.Decrypt(encryptText, pass)
MsgBox "��������� �����: "  & data & vbCr & vbCr _ 
     & "�������������  �����: "  & encryptText & vbCr _ 
     & "�������������� �����: " & decryptText

set m_uniref = nothing