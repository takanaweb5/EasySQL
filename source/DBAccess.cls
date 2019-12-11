VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const C_CONNECTSTR = "Provider={Provider};Data Source=""{FileName}"";Extended Properties=""{Properties}"";"
Private Const C_PROVIDER = "Microsoft.ACE.OLEDB.12.0"
'Private Const C_PROVIDER = "Microsoft.Jet.OLEDB.4.0"
Private Const C_PROPERTIES = "Excel 12.0"
'Private Const C_PROPERTIES = "Excel 8.0"
Private Const adOpenStatic = 3   'adOpenForwardOnly ではレコード件数が返ってこない
Private Const adLockReadOnly = 1
Private Const adCmdText = 1

Private FSQL As String

'*****************************************************************************
'[概要] データベースの接続文字列を取得する
'[引数] なし
'[戻値] データベース接続文字列
'*****************************************************************************
Private Function GetConnection() As String
    GetConnection = C_CONNECTSTR
    GetConnection = Replace(GetConnection, "{Provider}", C_PROVIDER)
    GetConnection = Replace(GetConnection, "{FileName}", ActiveWorkbook.FullName)
    GetConnection = Replace(GetConnection, "{Properties}", C_PROPERTIES)
End Function

'*****************************************************************************
'[概要] ＳＱＬの構文チェックを行う
'       項目名の間違いはパラメータと判定されるのでパラメータの情報を表示する例外を生成する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub CheckSQL()
    Dim Param  As Object
    Dim strErr As String
    With CreateObject("ADODB.Command")
        .ActiveConnection = GetConnection()
        .CommandType = adCmdText
        .CommandText = FSQL
        For Each Param In .Parameters
            strErr = strErr & Param.Name & " ? " & vbLf
        Next
    End With
    If strErr <> "" Then
        Call Err.Raise(513, , strErr)
    End If
End Sub

'*****************************************************************************
'[概要] 更新系またはDDLのSQLを実行する(※Accessファイルに対して実行可能)
'[引数] なし
'[戻値] 更新系SQLの時、更新レコード数
'*****************************************************************************
Public Function Execute() As Long
    Dim vRecCount As Variant
    With CreateObject("ADODB.Command")
        .ActiveConnection = GetConnection()
        .CommandType = adCmdText
        .CommandText = FSQL
        Call .Execute(vRecCount)
    End With
    On Error Resume Next
    Execute = vRecCount
    On Error GoTo 0
End Function

'*****************************************************************************
'[概要] SQLのSELECT文を実行し結果をセルに設定する
'[引数] 結果を設定するセル、SQL
'[戻値] SQL結果のレコード数
'*****************************************************************************
Public Function ExecuteToRange(ByRef objRange As Range, Optional ByVal strSQL As String = "") As Long
    If strSQL = "" Then
        strSQL = FSQL
    End If
    Dim objRecordset  As Object
    Set objRecordset = CreateObject("ADODB.Recordset")
    Call objRecordset.Open(strSQL, GetConnection(), adOpenStatic, adLockReadOnly)
On Error GoTo ErrHandle
    
    '見出しの設定
    Dim i As Long
    Application.ScreenUpdating = False
    For i = 1 To objRecordset.Fields.Count
        objRange.Cells(1, i) = objRecordset.Fields(i - 1).Name
    Next
    Application.ScreenUpdating = True
    DoEvents
    
    '明細の設定
    Application.ScreenUpdating = False
    Call objRange.Cells(2, 1).CopyFromRecordset(objRecordset)
    Application.ScreenUpdating = True
    
    ExecuteToRange = objRecordset.RecordCount
    Call objRecordset.Close
    Exit Function
ErrHandle:
    Call objRecordset.Close
    'エラー再作成
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

'*****************************************************************************
'[概要] SQLのSELECT文を実行し結果を2次元配列の表で取得する
'[引数] SQL
'[戻値] 実行結果の2次元配列
'*****************************************************************************
Public Function ExecuteToArray(Optional ByVal strSQL As String = "") As Variant
    If strSQL = "" Then
        strSQL = FSQL
    End If
    Dim objRecordset  As Object
    Set objRecordset = CreateObject("ADODB.Recordset")
    Call objRecordset.Open(strSQL, GetConnection(), adOpenStatic, adLockReadOnly)
On Error GoTo ErrHandle
    
    ReDim vData(0 To objRecordset.RecordCount, 1 To objRecordset.Fields.Count) '(行,列)
    Dim i As Long
    Dim j As Long
    
    '0行目に見出しを設定する
    For i = 1 To objRecordset.Fields.Count
        vData(0, i) = objRecordset.Fields(i - 1).Name
    Next
    
    '明細の設定
    For j = 1 To objRecordset.RecordCount
        For i = 1 To objRecordset.Fields.Count
            If IsNull(objRecordset.Fields(i - 1).Value) Then
                vData(j, i) = CVErr(xlErrNull)
            Else
                vData(j, i) = objRecordset.Fields(i - 1).Value
            End If
        Next
        objRecordset.MoveNext
    Next
    
    ExecuteToArray = vData()
    Call objRecordset.Close
    Exit Function
ErrHandle:
    Call objRecordset.Close
    'エラー再作成
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

'*****************************************************************************
'[プロパティ] SQL
'[引数] なし
'*****************************************************************************
Public Property Get SQL() As String
    SQL = FSQL
End Property
Public Property Let SQL(ByVal strSQL As String)
    FSQL = strSQL
End Property

'*****************************************************************************
'[概要] SQLのコメントを削除する　※クラスメソッドとして実装したいが
'       VBAにクラスメソッドはないのでインスタンスメソッドとして実装する
'[引数] コメント削除前のSQL
'[戻値] コメント削除後のSQL
'*****************************************************************************
Public Function DeleteComment(ByVal strSQL As String) As String
On Error GoTo ErrHandle
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.MultiLine = True
    objRegExp.Global = True
    
    '例：WHERE A IN('--1', "--2", [A--])  ※←これはコメントにしない
    Const OKPATTERN = "'.*?'|"".*?""|\[.*?\]"
    
    '例：-- これはコメント  ※行末までを削除
    objRegExp.Pattern = Replace("({OKPATTERN})|--.*$", "{OKPATTERN}", OKPATTERN)
    strSQL = objRegExp.Replace(strSQL, "$1")
    
    '例：/* これはコメント */  ※複数行にわたっていても削除
    strSQL = Replace(strSQL, vbLf, Chr(0))  '改行をChr(0)に変換
    objRegExp.Pattern = Replace("({OKPATTERN})|\/\*.+?\*\/", "{OKPATTERN}", OKPATTERN)
    strSQL = objRegExp.Replace(strSQL, "$1")
ErrHandle:
    DeleteComment = Replace(strSQL, Chr(0), vbLf)  'Chr(0)を改行に戻す
End Function

