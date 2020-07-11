VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'*****************************************************************************
'[概要] ユーザ定義セル関数のHELPを設定
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub InstHelpMsg()
    Dim strMsg As String
    Dim strInMsg(1 To 3) As String
    
    On Error Resume Next
    strMsg = "以下の例のようにセルの値を区切り文字で連結します" & vbLf
    strMsg = strMsg & "　　'AAA','BBB','CCC'　　や　　[AAA],[BBB],[CCC]" & vbLf
    strMsg = strMsg & "SQLのIN演算子の条件の羅列などに利用すると便利です"
    strInMsg(1) = "：　'　や　[]"
    strInMsg(2) = "：　,　など"
    strInMsg(3) = "：値を連結したいセル範囲を指定します(複数指定可)"
    Call Application.MacroOptions("VALUEJOIN", strMsg, , , , , "EasySQL", , , , Array(strInMsg(1), strInMsg(2), strInMsg(3)))
    
    strMsg = "SQLの実行結果を2次元配列で返します" & vbLf
    strMsg = strMsg & "範囲を指定して配列数式形式(Ctrl+Shift+Enter)で取り出してください"
    strInMsg(1) = "：SQLの入力されたセル"
    strInMsg(2) = "：再計算のトリガーにしたいセルがあれば指定します(いくつでも可)"
    Call Application.MacroOptions("GetSQLRecordset", strMsg, , , , , "EasySQL", , , , Array(strInMsg(1), strInMsg(2)))
    
    strMsg = "セル参照の反映とコメント削除を実行した後の、データベースに渡すSQLを表示します"
    strInMsg(1) = "：SQLの入力されたセル"
    Call Application.MacroOptions("ReplaceCellRef", strMsg, , , , , "EasySQL", , , , Array(strInMsg(1)))
    'この1行がないと、MacroOptionsコマンドはbasファイルのAttributeを書き換えるため
    'このファイルを閉じる時に「変更を保存しますか？」と変更がなくても表示されてしまう。
    ThisWorkbook.Saved = True
End Sub

'*****************************************************************************
'[概要] ユーザ定義セル関数のHELPを削除
'[引数] なし
'[戻値] なし
'*****************************************************************************
'Public Sub UninstHelpMsg()
'    On Error Resume Next
'    Call Application.MacroOptions("VALUEJOIN", "", , , , , 14)
'    Call Application.MacroOptions("GetSQLRecordset", "", , , , , 14)
'    Call Application.MacroOptions("ReplaceCellRef", "", , , , , 14)
'    '14：ユーザ定義(デフォルト値)
'End Sub


