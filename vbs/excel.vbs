'�������擾
Dim args: args = WScript.Arguments(0)

'�G���[�������ɏ������p��
On Error Resume Next

'�����m�F
MsgBox "������" & args

' Excel�A�v���P�[�V�����̃C���X�^���X����
Dim objXls: Set objXls = CreateObject("Excel.Application")
' If Not objXls Then WScript.Quit 1

MsgBox "�������P�I"

' Excel�̕\��
objXls.Visible = True
' objXls.ScreenUpdating = False

' Workbook��V�K�쐬
Dim objWorkbook
Set objWorkbook = objXls.Workbooks.Add()

'
' ���낢�돈������c
'

' Workbook�����
objWorkbook.Close

' Excel�̏I��
' objXls.ScreenUpdating = True
MsgBox "�������Q�I"

objXls.Quit

' �C���X�^���X�̔j��
Set objWorkbook = Nothing
Set objXls = Nothing

'�߂�l���w�肵�ďI������
WScript.Quit 0
