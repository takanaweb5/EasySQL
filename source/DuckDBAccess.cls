VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DuckDBAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' DuckDB C API 定数（ヘッダ duckdb.h 準拠）
Private Const DuckDBSuccess As Long = 0
Private Const DuckDBTypeBoolean As Long = 1
Private Const DuckDBTypeTinyInt As Long = 2
Private Const DuckDBTypeSmallInt As Long = 3
Private Const DuckDBTypeInteger As Long = 4
Private Const DuckDBTypeBigInt As Long = 5
Private Const DuckDBTypeFloat As Long = 10
Private Const DuckDBTypeDouble As Long = 11
Private Const DuckDBTypeTimestamp As Long = 12
Private Const DuckDBTypeDate As Long = 13
Private Const DuckDBTypeTime As Long = 14
Private Const DuckDBTypeVarchar As Long = 17
Private Const DuckDBTypeSqlNull As Long = 36

' 必要に応じて duckdb.dll のパスを指定（空文字ならワークブックと同階層）
Private Const DLLPATH As String = "C:\MyPC\Downloads\duckdb.dll"
Private Const DBPATH As String = "z:\test.duckdb"

Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub duckdb_free Lib "duckdb.dll" (ByVal ptr As LongPtr)

' DB / 接続 / クエリ関連
Private Declare PtrSafe Function duckdb_open Lib "duckdb.dll" (ByVal path As LongPtr, ByRef out_db As LongPtr) As Long
Private Declare PtrSafe Function duckdb_close Lib "duckdb.dll" (ByRef db As LongPtr) As Long
Private Declare PtrSafe Function duckdb_connect Lib "duckdb.dll" (ByVal db As LongPtr, ByRef out_conn As LongPtr) As Long
Private Declare PtrSafe Function duckdb_disconnect Lib "duckdb.dll" (ByRef conn As LongPtr) As Long
Private Declare PtrSafe Function duckdb_query Lib "duckdb.dll" (ByVal conn As LongPtr, ByVal query As LongPtr, ByVal out_result As LongPtr) As Long
Private Declare PtrSafe Sub duckdb_destroy_result Lib "duckdb.dll" (ByVal result As LongPtr)
Private Declare PtrSafe Function duckdb_result_error Lib "duckdb.dll" (ByVal result As LongPtr) As LongPtr
Private Declare PtrSafe Function duckdb_rows_changed Lib "duckdb.dll" (ByVal result As LongPtr) As LongLong

' 結果メタデータ・チャンク取得関連
Private Declare PtrSafe Function duckdb_column_count Lib "duckdb.dll" (ByVal result As LongPtr) As LongLong
Private Declare PtrSafe Function duckdb_column_name Lib "duckdb.dll" (ByVal result As LongPtr, ByVal col As LongLong) As LongPtr
Private Declare PtrSafe Function duckdb_result_chunk_count Lib "duckdb.dll" (ByVal result As LongPtr) As LongLong
Private Declare PtrSafe Function duckdb_result_get_chunk Lib "duckdb.dll" (ByVal result As LongPtr, ByVal chunk_index As LongLong) As LongPtr
Private Declare PtrSafe Sub duckdb_destroy_data_chunk Lib "duckdb.dll" (ByRef chunk As LongPtr)
Private Declare PtrSafe Function duckdb_data_chunk_get_size Lib "duckdb.dll" (ByVal chunk As LongPtr) As LongLong
Private Declare PtrSafe Function duckdb_data_chunk_get_column_count Lib "duckdb.dll" (ByVal chunk As LongPtr) As LongLong
Private Declare PtrSafe Function duckdb_data_chunk_get_vector Lib "duckdb.dll" (ByVal chunk As LongPtr, ByVal col_idx As LongLong) As LongPtr

' ベクタ（列データ・NULL判定・型）取得関連
Private Declare PtrSafe Function duckdb_vector_get_column_type Lib "duckdb.dll" (ByVal vector As LongPtr) As LongPtr
Private Declare PtrSafe Function duckdb_get_type_id Lib "duckdb.dll" (ByVal logical_type As LongPtr) As Long
Private Declare PtrSafe Sub duckdb_destroy_logical_type Lib "duckdb.dll" (ByRef type_ptr As LongPtr)
Private Declare PtrSafe Function duckdb_vector_get_data Lib "duckdb.dll" (ByVal vector As LongPtr) As LongPtr
Private Declare PtrSafe Function duckdb_vector_get_validity Lib "duckdb.dll" (ByVal vector As LongPtr) As LongPtr
Private Declare PtrSafe Function duckdb_validity_row_is_valid Lib "duckdb.dll" (ByVal validity As LongPtr, ByVal Row As LongLong) As Byte

' Appender関連
Private Declare PtrSafe Function duckdb_appender_create Lib "duckdb.dll" (ByVal conn As LongPtr, ByVal schema As LongPtr, ByVal table As LongPtr, ByRef out_appender As LongPtr) As Long
Private Declare PtrSafe Function duckdb_appender_end_row Lib "duckdb.dll" (ByVal appender As LongPtr) As Long
Private Declare PtrSafe Function duckdb_appender_flush Lib "duckdb.dll" (ByVal appender As LongPtr) As Long
Private Declare PtrSafe Function duckdb_appender_destroy Lib "duckdb.dll" (ByRef appender As LongPtr) As Long
Private Declare PtrSafe Function duckdb_appender_error Lib "duckdb.dll" (ByVal appender As LongPtr) As LongPtr
Private Declare PtrSafe Function duckdb_append_null Lib "duckdb.dll" (ByVal appender As LongPtr) As Long
Private Declare PtrSafe Function duckdb_append_bool Lib "duckdb.dll" (ByVal appender As LongPtr, ByVal val As Byte) As Long
Private Declare PtrSafe Function duckdb_append_int64 Lib "duckdb.dll" (ByVal appender As LongPtr, ByVal val As LongLong) As Long
Private Declare PtrSafe Function duckdb_append_double Lib "duckdb.dll" (ByVal appender As LongPtr, ByVal val As Double) As Long
Private Declare PtrSafe Function duckdb_append_date Lib "duckdb.dll" (ByVal appender As LongPtr, ByVal val As Long) As Long
Private Declare PtrSafe Function duckdb_append_time Lib "duckdb.dll" (ByVal appender As LongPtr, ByVal val As LongLong) As Long
Private Declare PtrSafe Function duckdb_append_timestamp Lib "duckdb.dll" (ByVal appender As LongPtr, ByVal val As LongLong) As Long
Private Declare PtrSafe Function duckdb_append_varchar Lib "duckdb.dll" (ByVal appender As LongPtr, ByVal val As LongPtr) As Long

Private Const TYPE_INFER_ROWS As Long = 1000

Private hDll As LongPtr
Private FSQL As String
Private FHeaders As Boolean
Private FDB As LongPtr
Private FConn As LongPtr

' duckdb_string_t(16バイト) 短い文字列(<=12バイト)はインライン格納
Private Type DuckDBStringShort
    strLen As Long
    inlined(0 To 11) As Byte
End Type

' duckdb_string_t(16バイト) 長い文字列は先頭4バイトのプレフィックス+ポインタ
Private Type DuckDBStringLong
    strLen As Long
    prefix(0 To 3) As Byte
    ptr As LongPtr
End Type

'*****************************************************************************
'[概要] コンストラクタ
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub Class_Initialize()
    If DLLPATH = "" Then
        hDll = LoadLibrary(ThisWorkbook.path & "\duckdb.dll")
    Else
        hDll = LoadLibrary(DLLPATH)
    End If
    If OpenDB(FDB, FConn) <> DuckDBSuccess Then Call Err.Raise(513, , "DuckDBのオープンに失敗しました")
    FHeaders = True
End Sub

'*****************************************************************************
'[概要] デストラクタ
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub Class_Terminate()
    Call CloseDB(FDB, FConn)
    If hDll <> 0 Then Call FreeLibrary(hDll)
End Sub

'*****************************************************************************
'[プロパティ] Headers ※True=最初の行を見出しとして扱う
'[引数] なし
'*****************************************************************************
Public Property Get Headers() As Boolean
    Headers = FHeaders
End Property
Public Property Let Headers(ByVal blnHeaders As Boolean)
    FHeaders = blnHeaders
End Property

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
'[概要] DBAccessクラスのインターフェースにあわせるため
'[引数] SQL
'[戻値] 構文エラーのパラメータ情報
'*****************************************************************************
Public Function CheckSQL(Optional ByVal strSQL As String = "") As String
End Function

'*****************************************************************************
'[概要] 更新系またはDDLのSQLを実行する
'[引数] SQL
'[戻値] 更新系SQLの時、更新レコード数（SELECTやDDLでは0）
'*****************************************************************************
Public Function Execute(Optional ByVal strSQL As String = "") As Long
    Dim res As LongPtr ' duckdb_result 構造体バッファ（環境に合わせて十分なサイズを確保）
    Dim resBuf(0 To 255) As LongPtr
    
    If strSQL = "" Then strSQL = FSQL
    
On Error GoTo ErrHandle
    res = VarPtr(resBuf(0))
    If duckdb_query(FConn, StrPtr(StringToUTF8(strSQL)), res) <> DuckDBSuccess Then
        Call Err.Raise(513, , UTF8ToString(duckdb_result_error(res)))
    End If
    
    Execute = CLng(duckdb_rows_changed(res))
    Call duckdb_destroy_result(res)
    Exit Function
ErrHandle:
    Call duckdb_destroy_result(res)
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

'*****************************************************************************
'[概要] SQLのSELECT文を実行し結果をセルに設定する
'[引数] 結果を設定するセル、SQL
'[戻値] SQL結果のレコード数
'*****************************************************************************
Public Function ExecuteToRange(ByRef objRange As Range, Optional ByVal strSQL As String = "") As Long
    Dim vData As Variant
    vData = ExecuteToArray(strSQL)
    Set objRange = objRange.Resize(UBound(vData, 1) - LBound(vData, 1) + 1, UBound(vData, 2) - LBound(vData, 2) + 1)
    objRange.Value = vData
    ExecuteToRange = UBound(vData, 1)
End Function

'*****************************************************************************
'[概要] SQLのSELECT文を実行し結果を2次元配列の表で取得する
'[引数] SQL
'[戻値] 実行結果の2次元配列
'*****************************************************************************
Public Function ExecuteToArray(Optional ByVal strSQL As String = "") As Variant
    Dim resBuf(0 To 255) As LongPtr
    Dim res As LongPtr
    Dim colCount As Long, chunkCount As Long, totalRows As Long
    Dim cIdx As Long, rBase As Long, rowInChunk As Long, x As Long, y As Long
    Dim chunk As LongPtr, chunkSize As Long, vec As LongPtr, lType As LongPtr, typeId As Long
    Dim dataPtr As LongPtr, valPtr As LongPtr
    Dim vData() As Variant
    
    If strSQL = "" Then strSQL = FSQL
    
On Error GoTo ErrHandle
    res = VarPtr(resBuf(0))
    If duckdb_query(FConn, StrPtr(StringToUTF8(strSQL)), res) <> DuckDBSuccess Then
        Call Err.Raise(513, , UTF8ToString(duckdb_result_error(res)))
    End If
    
    colCount = CLng(duckdb_column_count(res))
    chunkCount = CLng(duckdb_result_chunk_count(res))
    
    ' 総行数を集計
    For cIdx = 0 To chunkCount - 1
        chunk = duckdb_result_get_chunk(res, cIdx)
        totalRows = totalRows + CLng(duckdb_data_chunk_get_size(chunk))
        Call duckdb_destroy_data_chunk(chunk)
    Next
    
    If FHeaders Then
        ReDim vData(0 To totalRows, 1 To colCount)
        For x = 1 To colCount
            vData(0, x) = UTF8ToString(duckdb_column_name(res, x - 1))
        Next
    Else
        ReDim vData(1 To totalRows, 1 To colCount)
    End If
    
    rBase = 1
    For cIdx = 0 To chunkCount - 1
        chunk = duckdb_result_get_chunk(res, cIdx)
        chunkSize = CLng(duckdb_data_chunk_get_size(chunk))
        
        For x = 1 To colCount
            vec = duckdb_data_chunk_get_vector(chunk, x - 1)
            lType = duckdb_vector_get_column_type(vec)
            typeId = duckdb_get_type_id(lType)
            dataPtr = duckdb_vector_get_data(vec)
            valPtr = duckdb_vector_get_validity(vec)
            
            For rowInChunk = 0 To chunkSize - 1
                y = rBase + rowInChunk
                If valPtr = 0 Then
                    ' バリデーションポインタがNULLなら、すべて有効
                    vData(y, x) = ReadVectorValue(dataPtr, typeId, rowInChunk)
                Else
                    ' バリデーションポインタがある場合、NULLチェック
                    If duckdb_validity_row_is_valid(valPtr, rowInChunk) = 0 Then
                        vData(y, x) = Empty
                    Else
                        vData(y, x) = ReadVectorValue(dataPtr, typeId, rowInChunk)
                    End If
                End If
            Next
            Call duckdb_destroy_logical_type(lType)
        Next
        
        rBase = rBase + chunkSize
        Call duckdb_destroy_data_chunk(chunk)
    Next
    
    Call duckdb_destroy_result(res)
    ExecuteToArray = vData
    Exit Function
ErrHandle:
    Call duckdb_destroy_result(res)
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

'*****************************************************************************
'[概要] 範囲のデータをTEMPテーブルに投入し、テーブル名を返す（同名のテーブルが既にあれば投入せずにその名前を返す）
'[引数] テーブル名、投入元の範囲（1行目を列名、2行目以降をデータとする）
'[戻値] 作成（または既存）のTEMPテーブル名
'*****************************************************************************
Public Function CreateTempTable(ByVal strName As String, ByRef objRange As Range) As String
    Dim vData As Variant
    Dim vOne(1 To 1, 1 To 1) As Variant
    Dim strTypes() As String
    Dim strCols As String
    Dim strHead As String
    Dim strErr As String
    Dim strSQL As String
    Dim resBuf(0 To 255) As LongPtr
    Dim res As LongPtr
    Dim app As LongPtr
    Dim lngRow As Long, lngCol As Long

On Error GoTo ErrHandle
    vData = objRange.Value
    If Not IsArray(vData) Then
        vOne(1, 1) = vData
        vData = vOne
    End If

    ' 列数分の型配列を確保
    ReDim strTypes(1 To objRange.Columns.Count)
    For lngCol = 1 To objRange.Columns.Count
        strTypes(lngCol) = InferColumnType(vData, lngCol, Application.Min(objRange.Rows.Count, TYPE_INFER_ROWS))
        If IsError(vData(1, lngCol)) Then
            strHead = ""
        Else
            strHead = CStr(vData(1, lngCol))
        End If
        ' もし列名が空なら、col1, col2, ... とする
        If strHead = "" Then
            strHead = "col" & lngCol
        End If

        ' 例："col1" TEXT, "col2" INTEGER, ...
        If strCols <> "" Then
            strCols = strCols & ", "
        End If
        strCols = strCols & """" & strHead & """ " & strTypes(lngCol)
    Next

    ' テーブル作成SQLを組み立て
    strSQL = "CREATE TEMP TABLE """ & strName & """ (" & strCols & ")"
    res = VarPtr(resBuf(0))
    If duckdb_query(FConn, StrPtr(StringToUTF8(strSQL)), res) <> DuckDBSuccess Then
        strErr = UTF8ToString(duckdb_result_error(res))
        Call duckdb_destroy_result(res)
        If InStr(strErr, strName) = 0 Or InStr(strErr, "already exists") = 0 Then Call Err.Raise(513, , strErr)
        CreateTempTable = strName
        Exit Function
    End If
    Call duckdb_destroy_result(res)

    ' Appenderを使ってデータを挿入
    If duckdb_appender_create(FConn, 0, StrPtr(StringToUTF8(strName)), app) <> DuckDBSuccess Then Call Err.Raise(513, , "Appenderの作成に失敗しました")
    For lngRow = 2 To objRange.Rows.Count
        For lngCol = 1 To objRange.Columns.Count
            If AppendValue(app, vData(lngRow, lngCol), strTypes(lngCol)) <> DuckDBSuccess Then Call Err.Raise(513, , UTF8ToString(duckdb_appender_error(app)))
        Next
        If duckdb_appender_end_row(app) <> DuckDBSuccess Then Call Err.Raise(513, , UTF8ToString(duckdb_appender_error(app)))
    Next
    If duckdb_appender_flush(app) <> DuckDBSuccess Then Call Err.Raise(513, , UTF8ToString(duckdb_appender_error(app)))
    Call duckdb_appender_destroy(app)
    CreateTempTable = strName
    Exit Function
ErrHandle:
    If app <> 0 Then Call duckdb_appender_destroy(app)
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

'*****************************************************************************
'[概要] SQL内の{{範囲}}を、範囲のデータから作成したテンポラリテーブル名（ダブルクオート付き）に置換する
'[引数] 置換前のSQL、SQLの入力させたセル
'[戻値] 置換後のSQL（不正な範囲の{{...}}は残す）
'*****************************************************************************
Public Function ReplaceTempTableRef(ByVal strSQL As String, ByRef objSQLCell As Range) As String
    Dim objRegExp As Object
    Dim objMatch As Object
    Dim objRng As Range
    Dim strRef As String
    
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True
    objRegExp.Pattern = "\{\{(.+?)\}\}"
    
    ReplaceTempTableRef = strSQL
    If objRegExp.Test(strSQL) Then
        For Each objMatch In objRegExp.Execute(strSQL)
            strRef = objMatch.SubMatches(0)
            Set objRng = ExecSQL.ToRange(strRef, objSQLCell)
            If objRng Is Nothing Then
                If Right(strRef, 1) = "!" Then
                    Set objRng = ExecSQL.ToRange(strRef & "A1", objSQLCell)
                Else
                    Set objRng = ExecSQL.ToRange(strRef & "!A1", objSQLCell)
                End If
                If Not objRng Is Nothing Then Set objRng = objRng.Worksheet.Cells
            End If
            If Not objRng Is Nothing Then Set objRng = TrimRange(objRng)
            If Not objRng Is Nothing Then
                ReplaceTempTableRef = Replace(ReplaceTempTableRef, objMatch.Value, """" & CreateTempTable(MakeTempTableName(objSQLCell.Worksheet.Parent.Name & "!" & objSQLCell.Worksheet.Name & "!" & strRef), objRng) & """")
            End If
        Next
    End If
End Function

'*****************************************************************************
'[概要] 全行・全列の範囲を、ヘッダ行の最後の列とA列の最後の行までに絞る
'[引数] 対象の範囲
'[戻値] 絞り込み後の範囲（終端が始端より前になる場合はNothing）
'*****************************************************************************
Private Function TrimRange(ByRef objRange As Range) As Range
    Dim ws As Worksheet
    Dim lngLastRow As Long, lngLastCol As Long
    Dim lngEndRow As Long, lngEndCol As Long

    Set ws = objRange.Worksheet
    If objRange.Rows.Count <> ws.Rows.Count And objRange.Columns.Count <> ws.Columns.Count Then
        Set TrimRange = objRange
        Exit Function
    End If
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Or Application.WorksheetFunction.CountA(ws.Columns(1)) = 0 Then
        Set TrimRange = ws.Cells(objRange.Row, objRange.Column)
        Exit Function
    End If

    lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lngEndRow = Application.Min(objRange.Row + objRange.Rows.Count - 1, lngLastRow)
    lngEndCol = Application.Min(objRange.Column + objRange.Columns.Count - 1, lngLastCol)
    If lngEndRow < objRange.Row Or lngEndCol < objRange.Column Then Exit Function
    Set TrimRange = ws.Range(ws.Cells(objRange.Row, objRange.Column), ws.Cells(lngEndRow, lngEndCol))
End Function

'*****************************************************************************
'[概要] 文字列からTEMPテーブル名を作成する（使えない文字は_に置換）
'[引数] 名前の元になる文字列（ブック!シート!参照）
'[戻値] TEMPテーブル名（先頭はTMP_）
'*****************************************************************************
Private Function MakeTempTableName(ByVal strSource As String) As String
    Dim strSrc As String, strName As String
    Dim i As Long, code As Long
    strSrc = strSource
    For i = 1 To Len(strSrc)
        code = AscW(Mid(strSrc, i, 1))
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or code = 95 Or code < 0 Or code > 127 Then
            strName = strName & Mid(strSrc, i, 1)
        Else
            strName = strName & "_"
        End If
    Next
    MakeTempTableName = "TMP_" & strName
End Function

'*****************************************************************************
'[概要] 列の値からDuckDBの列型を推論する（日付・時刻は DATE / TIME / TIMESTAMP を判定）
'[引数] 値の2次元配列、列番号、判定終了行（見出しの次の行から判定する）
'[戻値] DuckDBの型名
'*****************************************************************************
Private Function InferColumnType(ByRef vData As Variant, ByVal lngCol As Long, ByVal lngLast As Long) As String
    Dim lngRow As Long, v As Variant, strKind As String, strKinds As String
    Dim blnDay As Boolean, blnTime As Boolean
    blnDay = True
    blnTime = True
    ' 明細行（行番号2）から判定終了行までループ
    For lngRow = 2 To lngLast
        v = vData(lngRow, lngCol)
        strKind = GetCellKind(v)
        If strKind <> "" And InStr(strKinds, strKind) = 0 Then
            strKinds = strKinds & strKind
        End If
        If strKind = "T" Then
            If CDbl(v) <> Int(CDbl(v)) Then blnDay = False
            If Int(CDbl(v)) <> 0 Then blnTime = False
        End If
    Next
    Select Case strKinds
    Case "B"
        InferColumnType = "BOOLEAN"
    Case "I"
        InferColumnType = "BIGINT"
    Case "D", "DI", "ID"
        InferColumnType = "DOUBLE"
    Case "T"
        Select Case True
        Case blnDay
            InferColumnType = "DATE"
        Case blnTime
            InferColumnType = "TIME"
        Case Else
            InferColumnType = "TIMESTAMP"
        End Select
    Case Else
        InferColumnType = "VARCHAR"
    End Select
End Function

'*****************************************************************************
'[概要] セルの値の種類を返す（B:真偽値、I:整数、D:小数、T:日付時刻、S:文字列、空文字:NULL扱い）
'[引数] セルの値
'[戻値] 値の種類
'*****************************************************************************
Private Function GetCellKind(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
    Case vbBoolean
        GetCellKind = "B"
    Case vbDate
        GetCellKind = "T"
    Case vbString
        If v <> "" Then GetCellKind = "S"
    Case Else
        If IsNumeric(v) Then
            If CDbl(v) = Int(CDbl(v)) And Abs(CDbl(v)) < 9E+18 Then
                GetCellKind = "I"
            Else
                GetCellKind = "D"
            End If
        Else
            GetCellKind = "S"
        End If
    End Select
End Function

'*****************************************************************************
'[概要] 値を列型に応じて変換し、Appenderに追加する（空の値はNULLとして追加）
'[引数] Appender、値、DuckDBの型名
'[戻値] DuckDBの処理結果（0:成功）
'*****************************************************************************
Private Function AppendValue(ByVal app As LongPtr, ByVal v As Variant, ByVal strType As String) As Long
    If IsError(v) Or IsEmpty(v) Then
        AppendValue = duckdb_append_null(app)
    ElseIf VarType(v) = vbString And v = "" Then
        AppendValue = duckdb_append_null(app)
    Else
        Select Case strType
        Case "BOOLEAN"
            AppendValue = duckdb_append_bool(app, IIf(CBool(v), 1, 0))
        Case "BIGINT"
            If CDbl(v) <> Int(CDbl(v)) Then
                Call Err.Raise(13, , "BIGINT列に小数が含まれています")
            End If
            AppendValue = duckdb_append_int64(app, CLngLng(CDbl(v)))
        Case "DOUBLE"
            AppendValue = duckdb_append_double(app, CDbl(v))
        Case "DATE"
            If CDbl(v) <> Int(CDbl(v)) Then
                Call Err.Raise(13, , "DATE列に時刻が含まれています")
            End If
            AppendValue = duckdb_append_date(app, CLng(Int(CDbl(v))) - CLng(DateSerial(1970, 1, 1)))
        Case "TIME"
            AppendValue = duckdb_append_time(app, CLngLng(Round((CDbl(v) - Int(CDbl(v))) * 86400000000#)))
        Case "TIMESTAMP"
            AppendValue = duckdb_append_timestamp(app, CLngLng(Round((CDbl(v) - CDbl(DateSerial(1970, 1, 1))) * 86400000000#)))
        Case Else
            AppendValue = duckdb_append_varchar(app, StrPtr(StringToUTF8(CStr(v))))
        End Select
    End If
End Function

'*****************************************************************************
'[概要] ベクタのデータ領域から指定行の値を読み取る
'[引数] データ領域ポインタ、型ID、行番号(0始まり)
'[戻値] 読み取った値
'*****************************************************************************
Private Function ReadVectorValue(ByVal dataPtr As LongPtr, ByVal typeId As Long, ByVal idx As Long) As Variant
    Dim i32 As Long, i64 As LongLong, dbl As Double, flt As Single, bln As Boolean
    Dim i16 As Integer, byt As Byte
    Select Case typeId
        Case DuckDBTypeTinyInt
            'int8を読み取る(VBAのByteは符号なしなので128以上は-256で符号付きに補正)
            Call CopyMemory(byt, ByVal (dataPtr + idx * 1), 1)
            ReadVectorValue = CLng(IIf(byt > 127, byt - 256, byt))
        Case DuckDBTypeSmallInt
            'int16はVBAのInteger(符号付き2バイト)にそのままコピーできる
            Call CopyMemory(i16, ByVal (dataPtr + idx * 2), 2)
            ReadVectorValue = CLng(i16)
        Case DuckDBTypeInteger
            Call CopyMemory(i32, ByVal (dataPtr + idx * 4), 4)
            ReadVectorValue = i32
        Case DuckDBTypeBigInt
            Call CopyMemory(i64, ByVal (dataPtr + idx * 8), 8)
            ReadVectorValue = i64
        Case DuckDBTypeDouble
            Call CopyMemory(dbl, ByVal (dataPtr + idx * 8), 8)
            ReadVectorValue = dbl
        Case DuckDBTypeFloat
            Call CopyMemory(flt, ByVal (dataPtr + idx * 4), 4)
            ReadVectorValue = flt
        Case DuckDBTypeBoolean
            Call CopyMemory(bln, ByVal (dataPtr + idx * 1), 1)
            ReadVectorValue = bln
        Case DuckDBTypeDate
            '1970-01-01からの日数(int32)を日付に変換
            Call CopyMemory(i32, ByVal (dataPtr + idx * 4), 4)
            ReadVectorValue = DateSerial(1970, 1, 1) + i32
        Case DuckDBTypeTime
            '00:00:00からのマイクロ秒(int64)を時刻に変換
            Call CopyMemory(i64, ByVal (dataPtr + idx * 8), 8)
            ReadVectorValue = CDate(i64 / 86400000000#)
        Case DuckDBTypeTimestamp
            '1970-01-01 00:00:00(UTC)からのマイクロ秒(int64)を日時に変換
            Call CopyMemory(i64, ByVal (dataPtr + idx * 8), 8)
            ReadVectorValue = #1/1/1970# + i64 / 86400000000#
        Case DuckDBTypeVarchar
            ReadVectorValue = ReadDuckString(dataPtr + idx * 16)
        Case DuckDBTypeSqlNull
            ReadVectorValue = Empty
        Case Else
            ReadVectorValue = Empty
    End Select
End Function

'*****************************************************************************
'[概要] duckdb_string_t構造体から文字列を取得する
'[引数] duckdb_string_t構造体へのポインタ
'[戻値] 取得した文字列
'*****************************************************************************
Private Function ReadDuckString(ByVal strStructPtr As LongPtr) As String
    Dim s As DuckDBStringShort '16バイト全体のビューとしても使う
    Call CopyMemory(s, ByVal strStructPtr, LenB(s))
    If s.strLen <= 0 Then Exit Function
    If s.strLen <= 12 Then
        ReadDuckString = BytesUTF8ToString(VarPtr(s.inlined(0)), s.strLen)
    Else
        Dim sl As DuckDBStringLong
        LSet sl = s '共用体ビューの切り替え(16バイトをそのまま転送)
        ReadDuckString = BytesUTF8ToString(sl.ptr, s.strLen)
    End If
End Function

'*****************************************************************************
'[概要] データベースを開いてコネクションを取得する
'[引数] データベースハンドル、コネクションハンドル
'[戻値] 成功時0(DuckDBSuccess)
'*****************************************************************************
Private Function OpenDB(ByRef db As LongPtr, ByRef conn As LongPtr) As Long
    Dim sPath As String, pPath As LongPtr
    If DBPATH <> "" And DBPATH <> ":memory:" Then
        sPath = StringToUTF8(DBPATH)
        pPath = StrPtr(sPath)
    End If
    OpenDB = duckdb_open(pPath, db)
    If OpenDB = DuckDBSuccess Then
        OpenDB = duckdb_connect(db, conn)
    End If
End Function

'*****************************************************************************
'[概要] コネクションを切断してデータベースを閉じる
'[引数] データベースハンドル、コネクションハンドル
'[戻値] なし
'*****************************************************************************
Private Sub CloseDB(ByRef db As LongPtr, ByRef conn As LongPtr)
    If conn <> 0 Then Call duckdb_disconnect(conn)
    If db <> 0 Then Call duckdb_close(db)
End Sub

'*****************************************************************************
'[概要] メモリ上のUTF-8バイト列を文字列に変換する
'[引数] バイト列ポインタ、バイト数
'[戻値] 変換後文字列
'*****************************************************************************
Private Function BytesUTF8ToString(ByVal ptr As LongPtr, ByVal byteLen As Long) As String
    If ptr = 0 Or byteLen <= 0 Then Exit Function
    'null終端のバッファにコピーしてUTF8ToStringで変換する
    Dim buf() As Byte
    ReDim buf(0 To byteLen) As Byte
    Call CopyMemory(buf(0), ByVal ptr, byteLen)
    BytesUTF8ToString = UTF8ToString(VarPtr(buf(0)))
End Function
