Attribute VB_Name = "CellFunctions"
Option Explicit
'Option Private Module
'*****************************************************************************
'Option Private Module をコメントアウトすることで、外部にセル関数を公開する
'*****************************************************************************

'*****************************************************************************
'[概要] セル関数用のTextJoinもどき
'[引数] 両端文字:例："'"→'テキスト'、""""→"テキスト"、"[]"→[テキスト]
'       区切り文字:区切り文字（","等）、
'       連結セル():連結する Range
'[戻値] 連結後の文字列
'*****************************************************************************
Public Function VALUEJOIN(ByVal 両端文字 As String, ByVal 区切り文字 As String, ParamArray 連結セル())
Attribute VALUEJOIN.VB_Description = "以下の例のようにセルの値を区切り文字で連結します\n　　'AAA','BBB','CCC'　　や　　[AAA],[BBB],[CCC]\nSQLのIN演算子の条件の羅列などに利用すると便利です"
Attribute VALUEJOIN.VB_ProcData.VB_Invoke_Func = " \n18"
    Dim i       As Long
    Dim objCell As Range
    Dim strL    As String '左端に付ける文字
    Dim strR    As String '右端に付ける文字
    
    If Len(両端文字) <= 1 Then
        strL = 両端文字
        strR = 両端文字
    Else
        strL = Left(両端文字, Int(Len(両端文字) / 2))
        strR = Right(両端文字, Int(Len(両端文字) / 2))
    End If
    
    For i = LBound(連結セル) To UBound(連結セル)
        For Each objCell In 連結セル(i)
            If Not IsError(objCell.Value) Then
                If objCell.Value <> "" Then
                    VALUEJOIN = VALUEJOIN & 区切り文字 & strL & objCell.Text & strR
                End If
            End If
        Next
    Next
    '先頭の区切り文字を削除
    VALUEJOIN = Mid(VALUEJOIN, Len(区切り文字) + 1)
End Function

'*****************************************************************************
'[概要] SQLの結果を2次元配列で取得する
'[引数] SQLの入力されたセル、Dummy():再計算のトリガーにしたいセルがあれば設定する
'[戻値] 実行結果(2次元配列)※セル関数で配列数式形式(Ctrl+Shift+Enter)で取り出す
'*****************************************************************************
Public Function GetSQLRecordset(ByRef objSQLCell As Range, ParamArray Dummy()) As Variant
Attribute GetSQLRecordset.VB_Description = "SQLの実行結果を2次元配列で返します\n範囲を指定して配列数式形式(Ctrl+Shift+Enter)で取り出してください"
Attribute GetSQLRecordset.VB_ProcData.VB_Invoke_Func = " \n18"
On Error GoTo ErrHandle
    'SQLを取得し、構文チェックを実施する
    Dim clsDBAccess  As New DBAccess
    clsDBAccess.SQL = ReplaceCellReference(objSQLCell)
    Call clsDBAccess.CheckSQL

    'SELECT文の実行結果の2次元配列を取得
    GetSQLRecordset = clsDBAccess.ExecuteToArray()
    Exit Function
ErrHandle:
    'エラーメッセージを表示
    GetSQLRecordset = Err.Description
End Function

'*****************************************************************************
'[概要] セル参照部分の置換とコメント削除後のSQL(データベースに渡すSQL)を表示
'[引数] SQLの入力させたセル
'[戻値] セルの参照値を置換したSQL
'*****************************************************************************
Public Function ReplaceCellRef(ByRef objSQLCell As Range) As String
Attribute ReplaceCellRef.VB_Description = "セル参照の反映とコメント削除を実行した後の、データベースに渡すSQLを表示します"
Attribute ReplaceCellRef.VB_ProcData.VB_Invoke_Func = " \n18"
    ReplaceCellRef = ReplaceCellReference(objSQLCell)
End Function

