'引数を取得
Dim args: args = WScript.Arguments(0)

'エラー発生時に処理を継続
On Error Resume Next

'引数確認
MsgBox "引数は" & args

' Excelアプリケーションのインスタンス生成
Dim objXls: Set objXls = CreateObject("Excel.Application")
' If Not objXls Then WScript.Quit 1

MsgBox "処理中１！"

' Excelの表示
objXls.Visible = True
' objXls.ScreenUpdating = False

' Workbookを新規作成
Dim objWorkbook
Set objWorkbook = objXls.Workbooks.Add()

'
' いろいろ処理する…
'

' Workbookを閉じる
objWorkbook.Close

' Excelの終了
' objXls.ScreenUpdating = True
MsgBox "処理中２！"

objXls.Quit

' インスタンスの破棄
Set objWorkbook = Nothing
Set objXls = Nothing

'戻り値を指定して終了する
WScript.Quit 0
