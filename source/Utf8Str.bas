Attribute VB_Name = "Utf8Str"
Option Explicit

Private Const CP_UTF8 As Long = 65001
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
                                                                     ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

'*****************************************************************************
'[概要] UTF-8文字列へのポインタをVB文字列に変換する
'[引数] null終端UTF-8文字列へのポインタ
'[戻値] 変換後の文字列
'*****************************************************************************
Public Function UTF8ToString(ByVal pUtf8 As LongPtr) As String
On Error GoTo ErrHandle
    Dim size As Long
    size = MultiByteToWideChar(CP_UTF8, 0, pUtf8, -1, 0, 0)
    'sizeはnull終端文字を含むため、長さ0の文字列の時は1が返る
    If size <= 1 Then Exit Function

    Dim result As String
    '必要な長さを確保する(null終端を除く)
    result = String(size - 1, vbNullChar)
    If MultiByteToWideChar(CP_UTF8, 0, pUtf8, -1, StrPtr(result), size) <> 0 Then
        UTF8ToString = result
    End If

    Exit Function
ErrHandle:
    UTF8ToString = ""
End Function

'*****************************************************************************
'[概要] VB文字列をUTF-8のバイト列に変換する
'[引数] 変換元の文字列
'[戻値] UTF-8に変換したバイト列(null終端を含む)
'*****************************************************************************
Public Function StringToUTF8(ByVal text As String) As String
On Error GoTo ErrHandle
    If text = "" Then Exit Function
    
    Dim size As Long
    size = WideCharToMultiByte(CP_UTF8, 0, StrPtr(text), -1, 0, 0, 0, 0)
    If size = 0 Then Exit Function
    
    ReDim buf(size) As Byte
    If WideCharToMultiByte(CP_UTF8, 0, StrPtr(text), -1, VarPtr(buf(0)), size, 0, 0) <> 0 Then
        StringToUTF8 = buf
    End If
    
    Exit Function
ErrHandle:
    StringToUTF8 = ""
End Function

