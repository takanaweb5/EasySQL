Attribute VB_Name = "ExecSQL"
Option Explicit

'*****************************************************************************
'[�T�v] �I�����ꂽ�Z��������SELECT���̂ЂȌ`���쐬����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub MakeSQL()
    Dim objSelection As Range
    Dim objArea   As Range
    Dim strSQL    As String
    Dim strSELECT As String
    Dim strFROM   As String
    Dim i As Long
    
    If Selection Is Nothing Then
        Exit Sub
    End If
    If TypeOf Selection Is Range Then
        Set objSelection = Selection
    Else
        Exit Sub
    End If
        
    'SELECT��̐ݒ�
    For Each objArea In objSelection.Areas
        For i = 1 To objArea.Columns.Count
            If strSELECT = "" Then
                strSELECT = "SELECT DISTINCT"
                strSELECT = strSELECT & vbCrLf & "       [" & objArea(1, i).Text & "]"
            Else
                strSELECT = strSELECT & vbCrLf & "     , [" & objArea(1, i).Text & "]"
            End If
        Next
    Next

    'FROM��̐ݒ�
    If objSelection.Areas.Count = 1 And objSelection.Rows.Count > 1 Then
        strFROM = Replace("  FROM [{Sheet}${Range}]", "{Sheet}", objSelection.Worksheet.Name)
        strFROM = Replace(strFROM, "{Range}", objSelection.AddressLocal(False, False, xlA1))
    Else
        strFROM = Replace("  FROM [{Sheet}$]", "{Sheet}", objSelection.Worksheet.Name)
    End If
    
    '���̑��̋�̎��ʎq�̂ݐݒ�
    strSQL = strSELECT & vbCrLf & _
               strFROM & vbCrLf & _
               " WHERE " & vbCrLf & _
               " GROUP BY" & vbCrLf & _
               "HAVING " & vbCrLf & _
               " ORDER BY"
    
    Call SetClipbordText(strSQL)
    Dim strMsg As String
    strMsg = "�ȉ���SQL���N���b�v�{�[�h�ɃR�s�[���܂����B" & vbCrLf & strSQL
    Call MsgBox(strMsg)
End Sub

'*****************************************************************************
'[�T�v] �N���b�v�{�[�h�Ƀe�L�X�g��ݒ肷��
'[����] �ݒ肷�镶����
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub SetClipbordText(ByVal strText As String)
On Error GoTo ErrHandle
    Dim objCb As New DataObject
    Call objCb.Clear
    Call objCb.SetText(strText)
    Call objCb.PutInClipboard
ErrHandle:
End Sub

'*****************************************************************************
'[�T�v] SQL�������s����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ExecuteSQL()
    If Dir(ActiveWorkbook.FullName) = "" Then
        Call MsgBox("��x���ۑ�����Ă��Ȃ��t�@�C���ł͎��s�ł��܂���")
        Exit Sub
    End If
    
    If Selection Is Nothing Then
        Exit Sub
    End If
    If Not (TypeOf Selection Is Range) Then
        Call ActiveCell.Select
    End If
    Dim objCurrentSheet As Worksheet
    Set objCurrentSheet = ActiveSheet
    
    Dim objSQLCell As Range
    Set objSQLCell = SelectCell("SQL�̓��͂��ꂽ�Z����I�����Ă�������", Selection)
    If objSQLCell Is Nothing Then
        Exit Sub
    Else
        Call objCurrentSheet.Activate
    End If
    
    Dim strSQL As String
    strSQL = ReplaceCellRef(objSQLCell)
    If IsSelect(strSQL) = True Then
        Call ShowRecord(strSQL)
    Else
        Call Execute(strSQL)
    End If
End Sub

'*****************************************************************************
'[�T�v] SELECT�����ǂ������肷��
'[����] SQL
'[�ߒl] True�FSELECT��
'*****************************************************************************
Private Function IsSelect(ByVal strSQL As String) As Boolean
    strSQL = DeleteEtc(strSQL)
    strSQL = UCase(strSQL)
    strSQL = Replace(strSQL, vbLf, Chr(0))  '���s��Chr(0)�ɕϊ�
    strSQL = Trim(strSQL)
    If Left(strSQL, 6) <> "SELECT" And Left(strSQL, 9) <> "TRANSFORM" Then
        IsSelect = False
        Exit Function
    End If
    
    'SELECT * INTO ���̓f�[�^�x�[�X���X�V���邽�߁AFalse�Ƃ���
    If FindINTO(strSQL) = True Then
        IsSelect = False
    Else
        IsSelect = True
    End If
End Function

'*****************************************************************************
'[�T�v] SQL�̃R�����g�╶���񃊃e�������폜����
'[����] �R�����g�폜�O��SQL
'[�ߒl] �R�����g�폜���SQL
'*****************************************************************************
Private Function DeleteEtc(ByVal strSQL As String) As String
On Error GoTo ErrHandle
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True
    
    ' 'xxx' or "xxx" or [xxx] �Ɋ܂܂�镶�����INTO���܂߂Ă��ׂč폜����
    objRegExp.Pattern = "'.+?'|"".+?""|\[.+?\]"
    strSQL = objRegExp.Replace(strSQL, "")
ErrHandle:
    DeleteEtc = strSQL
End Function

'*****************************************************************************
'[�T�v] INTO�傪���邩�ǂ������肷��
'[����] SQL
'[�ߒl] True�FINTO�傠��
'*****************************************************************************
Private Function FindINTO(ByVal strSQL As String) As Boolean
On Error GoTo ErrHandle
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    
    '�P���INTO������
    objRegExp.Pattern = "\bINTO\b"
    FindINTO = objRegExp.Test(strSQL)
ErrHandle:
End Function

'*****************************************************************************
'[�T�v] DDL�܂���DML��SQL�����s����
'[����] SQL
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub Execute(ByVal strSQL As String)
On Error GoTo ErrHandle
    'SQL�̍\���`�F�b�N�����{����
    Dim clsDBAccess  As New DBAccess
    clsDBAccess.SQL = strSQL
    Call clsDBAccess.CheckSQL
    
    '�R�}���h�����s
    Dim dblTime As Double
    Dim lngRecCount As Long
    dblTime = Timer()
    lngRecCount = clsDBAccess.Execute
    Call MsgBox("�X�V������ " & lngRecCount & " ���ł�" & vbCrLf & vbCrLf & _
                "���s���ԁF" & Int(Timer() - dblTime) & " �b")
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] SELECT���̌��ʂ�\�`���ŃZ���ɓW�J����
'[����] SQL
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ShowRecord(ByVal strSQL As String)
On Error GoTo ErrHandle
    '�Z����I��������
    Dim objTopLeftCell As Range
    Set objTopLeftCell = SelectCell("���ʂ�\������Z����I�����Ă�������", Selection)
    If objTopLeftCell Is Nothing Then
        Exit Sub
    Else
        '�I��̈�̍���̃Z����ݒ�
        Set objTopLeftCell = objTopLeftCell.Cells(1)
    End If
    
    '���ʂ̃V�[�g��\�����āA���ʂ̃Z����I��
    Call objTopLeftCell.Worksheet.Activate
    Call objTopLeftCell.Select
    
    'SQL�̍\���`�F�b�N�����{����
    Dim clsDBAccess  As New DBAccess
    clsDBAccess.SQL = strSQL
    Call clsDBAccess.CheckSQL

    'SELECT���̎��s���ʂ̃��R�[�h�Z�b�g���Z���ɐݒ�
    Dim dblTime As Double
    Dim lngRecCount As Long
    dblTime = Timer()
    lngRecCount = clsDBAccess.ExecuteToRange(objTopLeftCell)
    Call MsgBox("���R�[�h������ " & lngRecCount & " ���ł�" & vbCrLf & vbCrLf & _
                "���s���ԁF" & Int(Timer() - dblTime) & " �b")
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] �t�H�[����\�����ăZ����I��������
'[����] �\�����郁�b�Z�[�W�AobjCurrentCell�F�����I��������Z��
'[�ߒl] �I�����ꂽ�Z���i�L�����Z������Nothing�j
'*****************************************************************************
Public Function SelectCell(ByVal strMsg As String, ByRef objCurrentCell As Range) As Range
    Dim strCell As String
    '�t�H�[����\��
    With frmSelectCell
        .Label.Caption = strMsg
        Call objCurrentCell.Worksheet.Activate
        .RefEdit.Text = objCurrentCell.AddressLocal
        Call .Show
        If .IsOK = True Then
            strCell = .RefEdit
        End If
    End With
    Call Unload(frmSelectCell)
    If strCell <> "" Then
        Set SelectCell = Range(strCell)
        If SelectCell.Address = SelectCell.Cells(1, 1).MergeArea.Address Then
            Set SelectCell = SelectCell.Cells(1, 1)
        End If
    End If
End Function

'*****************************************************************************
'[�T�v] SQL��{A1}�̕�����A1�Z���̒��g�Œu������
'       �������A{MYPATH}�̕����̓J�����g�t�H���_�ɒu������
'               {MYSHEET}�̕�����SQL�̂���V�[�g���ɒu������
'[����] SQL�̓��͂������Z��
'[�ߒl] �Z���̎Q�ƒl��u������SQL
'*****************************************************************************
Public Function ReplaceCellRef(ByRef objSQLCell As Range) As String
Attribute ReplaceCellRef.VB_Description = "�Z���Q�Ƃ̔��f�ƃR�����g�폜�����s������́A�f�[�^�x�[�X�ɓn��SQL��\�����܂�"
Attribute ReplaceCellRef.VB_ProcData.VB_Invoke_Func = " \n18"
On Error GoTo ErrHandle
    Dim objRegExp  As Object
    Dim objMatch   As Object
    Dim strSubSQL  As String
    Dim strReplace As String
    Dim dummy As New DBAccess
    
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True
    objRegExp.Pattern = "\{(.+?)\}"
    
    ReplaceCellRef = dummy.DeleteComment(GetRangeText(objSQLCell))
    If objRegExp.Test(ReplaceCellRef) Then
        For Each objMatch In objRegExp.Execute(ReplaceCellRef)
            strReplace = objMatch.SubMatches(0)
            Select Case StrConv(strReplace, vbUpperCase)
            Case "MYPATH"
                ReplaceCellRef = Replace(ReplaceCellRef, objMatch, ActiveWorkbook.Path)
            Case "MYSHEET"
                ReplaceCellRef = Replace(ReplaceCellRef, objMatch, objSQLCell.Worksheet.Name)
            Case Else
                Select Case IsCellAddress(strReplace, objSQLCell.Worksheet)
                Case 1 '����V�[�g�̃Z���̎�
                    '����V�[�g���̃Z���̓��e�Œu������@����FRange("A1")��
                    strSubSQL = ReplaceCellRef(objSQLCell.Worksheet.Range(strReplace))
                    ReplaceCellRef = Replace(ReplaceCellRef, objMatch, strSubSQL)
                Case 2 '�ʃV�[�g�̃Z���̎�
                    '�ʃV�[�g���̃Z���̓��e�Œu������@����FRange("Sheet1!A1")��
                    strSubSQL = ReplaceCellRef(Range(strReplace))
                    ReplaceCellRef = Replace(ReplaceCellRef, objMatch, strSubSQL)
                End Select
            End Select
        Next
    End If
ErrHandle:
End Function

'*****************************************************************************
'[�T�v] strAddress��Cell���w���A�h���X���ǂ���
'[����] �`�F�b�N�Ώۂ̕�����(�A�h���X �܂��� ���O)�A�J�����g�V�[�g
'[�ߒl] 0:Not�A�h���X�A1:�J�����g�V�[�g�̃A�h���X�A2:�ʃV�[�g�̃A�h���X
'*****************************************************************************
Private Function IsCellAddress(ByVal strAddress As String, ByRef objWorksheet As Worksheet) As Long
    If TypeOf Range(strAddress) Is Range Then
    Else
        IsCellAddress = 0
        Exit Function
    End If
On Error GoTo ErrHandle
    Dim dummy As Range
    Set dummy = objWorksheet.Range(strAddress)  '�G���[�Ȃ�΁A�ʃV�[�g�Ɣ���
    IsCellAddress = 1
    Exit Function
ErrHandle:
    IsCellAddress = 2
    Exit Function
End Function

'*****************************************************************************
'[�T�v] SQL�̑I�����ꂽ�Z���̒l���擾����
'[����] SQL�̓��͂�����Range
'[�ߒl] �Z���̒l�i�����s�̎��F�l�������l�łȂ��Z���̒l�����s�ŘA���j
'               �i������̎��F�l�������l�łȂ��Z���̒l��[]�ł͂���","�ŘA���j
'*****************************************************************************
Private Function GetRangeText(ByRef objRange As Range) As String
    Dim i As Long
    Dim strCellText As String
    
    '�P��Z���̎�
    If objRange.Count = 1 Or objRange.Address = objRange(1, 1).MergeArea.Address Then
        GetRangeText = objRange(1).Text
        Exit Function
    End If
    
    '���o���I����
    If objRange.Rows.Count = 1 Or objRange.Rows.Count = Rows.Count Then
        '�񐔕�LOOP���A�e���ڂ��J�b�R�ł͂��݃R���}�ŘA���@��F[��], [��]
        For i = 1 To objRange.Columns.Count
            strCellText = objRange.Cells(1, i)
            If strCellText <> "" Then
                If GetRangeText = "" Then
                    GetRangeText = "[" & strCellText & "]"
                Else
                    GetRangeText = GetRangeText & ",[" & strCellText & "]"
                End If
            End If
        Next
    Else
        '�s����LOOP���A�e�Z���̒l�����s�ŘA��
        For i = 1 To objRange.Rows.Count
            strCellText = objRange.Cells(i, 1)
            If strCellText <> "" Then
                If GetRangeText = "" Then
                    GetRangeText = strCellText
                Else
                    GetRangeText = GetRangeText & vbLf & strCellText
                End If
            End If
        Next
    End If
End Function

'*****************************************************************************
'[�T�v] SQL�̌��ʂ�2�����z��Ŏ擾����
'[����] SQL�̓��͂��ꂽ�Z���ADummy():�Čv�Z�̃g���K�[�ɂ������Z��������ΐݒ肷��
'[�ߒl] ���s����(2�����z��)���Z���֐��Ŕz�񐔎��`��(Ctrl+Shift+Enter)�Ŏ��o��
'*****************************************************************************
Public Function GetSQLRecordset(ByRef objSQLCell As Range, ParamArray dummy()) As Variant
Attribute GetSQLRecordset.VB_Description = "SQL�̎��s���ʂ�2�����z��ŕԂ��܂�\n�͈͂��w�肵�Ĕz�񐔎��`��(Ctrl+Shift+Enter)�Ŏ��o���Ă�������"
Attribute GetSQLRecordset.VB_ProcData.VB_Invoke_Func = " \n18"
On Error GoTo ErrHandle
    'SQL���擾���A�\���`�F�b�N�����{����
    Dim clsDBAccess  As New DBAccess
    clsDBAccess.SQL = ReplaceCellRef(objSQLCell)
    Call clsDBAccess.CheckSQL

    'SELECT���̎��s���ʂ�2�����z����擾
    GetSQLRecordset = clsDBAccess.ExecuteToArray()
    Exit Function
ErrHandle:
    '�G���[���b�Z�[�W��\��
    GetSQLRecordset = Err.Description
End Function

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
            If objCell.Value <> "" Then
                VALUEJOIN = VALUEJOIN & ��؂蕶�� & strL & objCell.Text & strR
            End If
        Next
    Next
    '�擪�̋�؂蕶�����폜
    VALUEJOIN = Mid(VALUEJOIN, Len(��؂蕶��) + 1)
End Function
