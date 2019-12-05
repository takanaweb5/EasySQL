Attribute VB_Name = "CellFunctions"
Option Explicit
'Option Private Module
'*****************************************************************************
'Option Private Module ���R�����g�A�E�g���邱�ƂŁA�O���ɃZ���֐������J����
'*****************************************************************************

'*****************************************************************************
'[�T�v] �Z���֐��p��TextJoin���ǂ�
'[����] ���[����:��F"'"��'�e�L�X�g'�A""""��"�e�L�X�g"�A"[]"��[�e�L�X�g]
'       ��؂蕶��:��؂蕶���i","���j�A
'       �A���Z��():�A������ Range
'[�ߒl] �A����̕�����
'*****************************************************************************
Public Function VALUEJOIN(ByVal ���[���� As String, ByVal ��؂蕶�� As String, ParamArray �A���Z��())
Attribute VALUEJOIN.VB_Description = "�ȉ��̗�̂悤�ɃZ���̒l����؂蕶���ŘA�����܂�\n�@�@'AAA','BBB','CCC'�@�@��@�@[AAA],[BBB],[CCC]\nSQL��IN���Z�q�̏����̗���Ȃǂɗ��p����ƕ֗��ł�"
Attribute VALUEJOIN.VB_ProcData.VB_Invoke_Func = " \n18"
    Dim i       As Long
    Dim objCell As Range
    Dim strL    As String '���[�ɕt���镶��
    Dim strR    As String '�E�[�ɕt���镶��
    
    If Len(���[����) <= 1 Then
        strL = ���[����
        strR = ���[����
    Else
        strL = Left(���[����, Int(Len(���[����) / 2))
        strR = Right(���[����, Int(Len(���[����) / 2))
    End If
    
    For i = LBound(�A���Z��) To UBound(�A���Z��)
        For Each objCell In �A���Z��(i)
            If Not IsError(objCell.Value) Then
                If objCell.Value <> "" Then
                    VALUEJOIN = VALUEJOIN & ��؂蕶�� & strL & objCell.Text & strR
                End If
            End If
        Next
    Next
    '�擪�̋�؂蕶�����폜
    VALUEJOIN = Mid(VALUEJOIN, Len(��؂蕶��) + 1)
End Function

'*****************************************************************************
'[�T�v] SQL�̌��ʂ�2�����z��Ŏ擾����
'[����] SQL�̓��͂��ꂽ�Z���ADummy():�Čv�Z�̃g���K�[�ɂ������Z��������ΐݒ肷��
'[�ߒl] ���s����(2�����z��)���Z���֐��Ŕz�񐔎��`��(Ctrl+Shift+Enter)�Ŏ��o��
'*****************************************************************************
Public Function GetSQLRecordset(ByRef objSQLCell As Range, ParamArray Dummy()) As Variant
Attribute GetSQLRecordset.VB_Description = "SQL�̎��s���ʂ�2�����z��ŕԂ��܂�\n�͈͂��w�肵�Ĕz�񐔎��`��(Ctrl+Shift+Enter)�Ŏ��o���Ă�������"
Attribute GetSQLRecordset.VB_ProcData.VB_Invoke_Func = " \n18"
On Error GoTo ErrHandle
    'SQL���擾���A�\���`�F�b�N�����{����
    Dim clsDBAccess  As New DBAccess
    clsDBAccess.SQL = ReplaceCellReference(objSQLCell)
    Call clsDBAccess.CheckSQL

    'SELECT���̎��s���ʂ�2�����z����擾
    GetSQLRecordset = clsDBAccess.ExecuteToArray()
    Exit Function
ErrHandle:
    '�G���[���b�Z�[�W��\��
    GetSQLRecordset = Err.Description
End Function

'*****************************************************************************
'[�T�v] �Z���Q�ƕ����̒u���ƃR�����g�폜���SQL(�f�[�^�x�[�X�ɓn��SQL)��\��
'[����] SQL�̓��͂������Z��
'[�ߒl] �Z���̎Q�ƒl��u������SQL
'*****************************************************************************
Public Function ReplaceCellRef(ByRef objSQLCell As Range) As String
Attribute ReplaceCellRef.VB_Description = "�Z���Q�Ƃ̔��f�ƃR�����g�폜�����s������́A�f�[�^�x�[�X�ɓn��SQL��\�����܂�"
Attribute ReplaceCellRef.VB_ProcData.VB_Invoke_Func = " \n18"
    ReplaceCellRef = ReplaceCellReference(objSQLCell)
End Function

