VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SQLITE_ROW = 100

Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_step Lib "sqlite3.dll" (ByVal stmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_open Lib "sqlite3.dll" (ByVal filename As LongPtr, ByRef ppDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_close Lib "sqlite3.dll" (ByVal DB As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_errmsg Lib "sqlite3.dll" (ByVal DB As LongPtr) As LongPtr
Private Declare PtrSafe Function sqlite3_changes Lib "sqlite3.dll" (ByVal DB As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_finalize Lib "sqlite3.dll" (ByVal stmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_column_count Lib "sqlite3.dll" (ByVal stmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_column_name Lib "sqlite3.dll" (ByVal stmt As LongPtr, ByVal N As Long) As LongPtr
Private Declare PtrSafe Function sqlite3_column_type Lib "sqlite3.dll" (ByVal stmt As LongPtr, ByVal col As Long) As Long
Private Declare PtrSafe Function sqlite3_column_int Lib "sqlite3.dll" (ByVal stmt As LongPtr, ByVal col As Long) As Long
Private Declare PtrSafe Function sqlite3_column_text Lib "sqlite3.dll" (ByVal stmt As LongPtr, ByVal col As Long) As LongPtr
Private Declare PtrSafe Function sqlite3_column_double Lib "sqlite3.dll" (ByVal stmt As LongPtr, ByVal col As Long) As Double
Private Declare PtrSafe Function sqlite3_exec Lib "sqlite3.dll" (ByVal DB As LongPtr, ByVal SQL As LongPtr, _
    ByVal callback As LongPtr, _
    ByVal param As LongPtr, _
    ByRef errmsg As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_prepare_v2 Lib "sqlite3.dll" (ByVal DB As LongPtr, ByVal zSql As LongPtr, _
    ByVal nByte As Long, _
    ByRef ppStmt As LongPtr, _
    ByVal pzTail As LongPtr) As Long

'PATH�̒ʂ��Ă��Ȃ�sqlite3.dll���g�p�\�ɂ��邽�߂�LoadLibrary�����s����
Private Const DLLPATH = "C:\Fujitsu\Softwaredisc\Software2\ODD\Software\Application\PowerDirector_AVCHD\bin\powerdirector\Setup\Utility\InteropPalette\x64\Lib\sqlite3.dll"
Private hDll As LongPtr

Private FSQL As String
Private FHeaders As Boolean '�ŏ��̍s�����o���Ƃ��Ĉ������ǂ���

Const DBPATH = "z:\test.db"

'*****************************************************************************
'[�T�v] �R���X�g���N�^
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub Class_Initialize()
    If DLLPATH = "" Then
        hDll = LoadLibrary(ActiveWorkbook.Path & "\sqlite3.dll")
    Else
        hDll = LoadLibrary(DLLPATH)
    End If
    FHeaders = True '�ŏ��̍s�����o���Ƃ��Ĉ���
End Sub

'*****************************************************************************
'[�T�v] �f�X�g���N�^
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub Class_Terminate()
    Call FreeLibrary(hDll)
End Sub

'*****************************************************************************
'[�v���p�e�B] Headers ��True=�ŏ��̍s�����o���Ƃ��Ĉ���
'[����] �Ȃ�
'*****************************************************************************
Public Property Get Headers() As Boolean
    Headers = FHeaders
End Property
Public Property Let Headers(ByVal blnHeaders As Boolean)
    FHeaders = blnHeaders
End Property

'*****************************************************************************
'[�T�v] �r�p�k�̍\���`�F�b�N���s��
'[����] SQL
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub CheckSQL(Optional ByVal strSQL As String = "")
End Sub

'*****************************************************************************
'[�T�v] �X�V�n�܂���DDL��SQL�����s����(��Access�t�@�C���ɑ΂��Ď��s�\)
'[����] SQL
'[�ߒl] �X�V�nSQL�̎��A�X�V���R�[�h��
'*****************************************************************************
Public Function Execute(Optional ByVal strSQL As String = "") As Long
    Dim retval As Long
    Dim errmsg As LongPtr
    Dim DB As LongPtr
    
    '�f�[�^�x�[�X���J��
    retval = sqlite3_open(StrPtr(StringToUTF8(DBPATH)), DB)
    If retval <> 0 Then
        Call Err.Raise(513, , "�f�[�^�x�[�X�t�@�C�����I�[�v���o���܂���")
    End If
    
On Error GoTo ErrHandle
    If strSQL = "" Then
        strSQL = FSQL
    End If
    
    retval = sqlite3_exec(DB, StrPtr(StringToUTF8(strSQL)), 0, 0, errmsg)
    If retval <> 0 Then
        Call Err.Raise(513, , UTF8ToString(errmsg))
    End If
    
    ' �ύX���ꂽ�s�����擾
    Execute = sqlite3_changes(DB)
    Exit Function
ErrHandle:
    Call sqlite3_close(DB)
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

'*****************************************************************************
'[�T�v] SQL��SELECT�������s�����ʂ��Z���ɐݒ肷��
'[����] ���ʂ�ݒ肷��Z���ASQL
'[�ߒl] SQL���ʂ̃��R�[�h��
'*****************************************************************************
Public Function ExecuteToRange(ByRef objRange As Range, Optional ByVal strSQL As String = "") As Long
    Dim vData() As Variant
    vData = ExecuteToArray(strSQL)
    
    Set objRange = objRange.Resize(UBound(vData, 1) - LBound(vData, 1) + 1, UBound(vData, 2) - LBound(vData, 2) + 1)
    objRange.Value = vData

    'NULL�̃Z���̓N���A����
    Call ClearNullCells(objRange)
    
    ExecuteToRange = UBound(vData, 1)
End Function

'*****************************************************************************
'[�T�v] SELECT�̌��ʂ���NULL�l�̃Z�����N���A����
'[����] �ΏۃZ���͈�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ClearNullCells(ByRef objRange As Range)
    Dim objCell As Range
    Dim objErrCells As Range
    Dim objNullRange As Range
    
    '�������̂��߂�objRange����G���[���܂ރZ�����擾
    On Error Resume Next '�G���[���Ȃ��ꍇ�ɃG���[�𖳎�����
    Set objErrCells = objRange.SpecialCells(xlCellTypeConstants, xlErrors)
    On Error GoTo 0
    
    If objErrCells Is Nothing Then
        Exit Sub
    End If
    
    For Each objCell In objErrCells
        If objCell.Value = CVErr(xlErrNull) Then
            If objNullRange Is Nothing Then
                Set objNullRange = objCell
            Else
                Set objNullRange = Union(objNullRange, objCell)
            End If
        End If
    Next

    If Not objNullRange Is Nothing Then
        Call objNullRange.ClearContents
    End If
End Sub

'*****************************************************************************
'[�T�v] SQL��SELECT�������s�����ʂ�2�����z��̕\�Ŏ擾����
'[����] SQL
'[�ߒl] ���s���ʂ�2�����z��
'*****************************************************************************
Public Function ExecuteToArray(Optional ByVal strSQL As String = "") As Variant
    Dim retval As Long
    Dim errmsg As LongPtr
    Dim DB As LongPtr
    Dim stmt As LongPtr
    
    '�f�[�^�x�[�X���J��
    retval = sqlite3_open(StrPtr(StringToUTF8(DBPATH)), DB)
    If retval <> 0 Then
        Call Err.Raise(513, , "�f�[�^�x�[�X�t�@�C�����I�[�v���o���܂���")
    End If
    
On Error GoTo ErrHandle
    If strSQL = "" Then
        strSQL = FSQL
    End If
        
    'SQL������
    retval = sqlite3_prepare_v2(DB, StrPtr(StringToUTF8(strSQL)), -1, stmt, 0)
    If retval <> 0 Then
        Call Err.Raise(513, , UTF8ToString(sqlite3_errmsg(DB)))
    End If
    
    '�J���������擾
    Dim colCount As Long
    colCount = sqlite3_column_count(stmt)
    
    '�s�����擾
    Dim rowCount As Long
    Do While sqlite3_step(stmt) = SQLITE_ROW
        rowCount = rowCount + 1
    Loop
    
    Dim x As Long: Dim y As Long:
    If FHeaders Then
        ReDim vData(0 To rowCount, 1 To colCount) '(�s,��)
    Else
        ReDim vData(1 To rowCount, 1 To colCount) '(�s,��)
    End If
    
    Call sqlite3_finalize(stmt): stmt = 0
    retval = sqlite3_prepare_v2(DB, StrPtr(StringToUTF8(strSQL)), -1, stmt, 0)
    If retval <> 0 Then
        Call Err.Raise(513, , UTF8ToString(sqlite3_errmsg(DB)))
    End If
        
    '0�s�ڂɌ��o����ݒ肷��
    If FHeaders Then
        For x = 1 To colCount
            vData(0, x) = CStr(UTF8ToString(sqlite3_column_name(stmt, x - 1)))
        Next
    End If
    
    '���ׂ̐ݒ�
    For y = 1 To rowCount
        Call sqlite3_step(stmt)
        For x = 1 To colCount
            Select Case sqlite3_column_type(stmt, x - 1)
            Case 1 ' SQLITE_INTEGER
                vData(y, x) = sqlite3_column_int(stmt, x - 1)
            Case 2 ' SQLITE_FLOAT
                vData(y, x) = sqlite3_column_double(stmt, x - 1)
            Case 3 ' SQLITE_TEXT
                vData(y, x) = CStr(UTF8ToString(sqlite3_column_text(stmt, x - 1)))
            Case 5 ' SQLITE_NULL
                vData(y, x) = CVErr(xlErrNull)
            Case Else
                vData(y, x) = CVErr(xlErrValue)
            End Select
        Next
    Next
    
    Call sqlite3_finalize(stmt)
    Call sqlite3_close(DB)
    
    ExecuteToArray = vData()
    Exit Function
ErrHandle:
    If stmt <> 0 Then Call sqlite3_finalize(stmt)
    Call sqlite3_close(DB)
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

'*****************************************************************************
'[�v���p�e�B] SQL
'[����] �Ȃ�
'*****************************************************************************
Public Property Get SQL() As String
    SQL = FSQL
End Property
Public Property Let SQL(ByVal strSQL As String)
    FSQL = strSQL
End Property