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
Private Declare PtrSafe Function duckdb_validity_row_is_valid Lib "duckdb.dll" (ByVal validity As LongPtr, ByVal row As LongLong) As Byte

Private hDll As LongPtr
Private FSQL As String
Private FHeaders As Boolean

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
        hDll = LoadLibrary(ThisWorkbook.Path & "\duckdb.dll")
    Else
        hDll = LoadLibrary(DLLPATH)
    End If
    FHeaders = True
End Sub

'*****************************************************************************
'[概要] デストラクタ
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub Class_Terminate()
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
'[概要] 更新系またはDDLのSQLを実行する
'[引数] SQL
'[戻値] 更新系SQLの時、更新レコード数（SELECTやDDLでは0）
'*****************************************************************************
Public Function Execute(Optional ByVal strSQL As String = "") As Long
    Dim db As LongPtr, conn As LongPtr
    Dim res As LongPtr ' duckdb_result 構造体バッファ（環境に合わせて十分なサイズを確保）
    Dim resBuf(0 To 255) As LongPtr
    
    If strSQL = "" Then strSQL = FSQL
    If OpenDB(db, conn) <> DuckDBSuccess Then Err.Raise 513, , "DuckDBのオープンに失敗しました"
    
On Error GoTo ErrHandle
    res = VarPtr(resBuf(0))
    If duckdb_query(conn, StrPtr(StringToUTF8(strSQL)), res) <> DuckDBSuccess Then
        Call Err.Raise(513, , UTF8ToString(duckdb_result_error(res)))
    End If
    
    Execute = CLng(duckdb_rows_changed(res))
    Call duckdb_destroy_result(res)
    Call CloseDB(db, conn)
    Exit Function
ErrHandle:
    Call duckdb_destroy_result(res)
    Call CloseDB(db, conn)
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
    Dim db As LongPtr, conn As LongPtr
    Dim resBuf(0 To 255) As LongPtr
    Dim res As LongPtr
    Dim colCount As Long, chunkCount As Long, totalRows As Long
    Dim cIdx As Long, rBase As Long, rowInChunk As Long, x As Long, y As Long
    Dim chunk As LongPtr, chunkSize As Long, vec As LongPtr, lType As LongPtr, typeId As Long
    Dim dataPtr As LongPtr, valPtr As LongPtr
    Dim vData() As Variant
    
    If strSQL = "" Then strSQL = FSQL
    If OpenDB(db, conn) <> DuckDBSuccess Then Err.Raise 513, , "DuckDBのオープンに失敗しました"
    
On Error GoTo ErrHandle
    res = VarPtr(resBuf(0))
    If duckdb_query(conn, StrPtr(StringToUTF8(strSQL)), res) <> DuckDBSuccess Then
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
    Call CloseDB(db, conn)
    ExecuteToArray = vData
    Exit Function
ErrHandle:
    Call duckdb_destroy_result(res)
    Call CloseDB(db, conn)
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
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
