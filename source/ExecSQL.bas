Attribute VB_Name = "ExecSQL"
Option Explicit
Option Private Module

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
                strSELECT = strSELECT & vbCrLf & "  [" & objArea(1, i).TEXT & "]"
            Else
                strSELECT = strSELECT & vbCrLf & ", [" & objArea(1, i).TEXT & "]"
            End If
        Next
    Next

    'FROM��̐ݒ�
    If objSelection.Areas.Count = 1 And objSelection.Rows.Count > 1 Then
        strFROM = Replace("FROM [{Sheet}${Range}]", "{Sheet}", objSelection.Worksheet.Name)
        strFROM = Replace(strFROM, "{Range}", objSelection.AddressLocal(False, False, xlA1))
    Else
        strFROM = Replace("FROM [{Sheet}$]", "{Sheet}", objSelection.Worksheet.Name)
    End If
    
    '���̑��̋�̎��ʎq�̂ݐݒ�
    strSQL = strSELECT & vbCrLf & _
               strFROM & vbCrLf & _
               "WHERE " & vbCrLf & _
               "GROUP BY" & vbCrLf & _
               "HAVING " & vbCrLf & _
               "ORDER BY"
    
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
'[�T�v] SQL�����s����
'[����] 1:Select���p,2:�X�V�n�p
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ExecuteSQL1()
    Call ExecuteSQL(True)
End Sub
Public Sub ExecuteSQL2()
    Call ExecuteSQL(False)
End Sub

'*****************************************************************************
'[�T�v] SQL�������s����
'[����] True:Select���AFalse:�X�V�nSQL
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ExecuteSQL(ByVal IsSelect As Boolean)
    If ActiveWorkbook.Path = "" Then
        Call MsgBox("��x���ۑ�����Ă��Ȃ��t�@�C���̓G���[�ɂȂ邱�Ƃ�����܂�")
    End If
    
    If Selection Is Nothing Then
        Exit Sub
    End If
    If Not (TypeOf Selection Is Range) Then
        Exit Sub
    End If
    
    Dim objSQLCell As Range
    Set objSQLCell = Selection
    If objSQLCell Is Nothing Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = ReplaceCellReference(objSQLCell)
    
    If IsSelect Then
        Call ShowRecord(strSQL)
    Else
        Call Execute(strSQL)
    End If
End Sub

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
    lngRecCount = clsDBAccess.Execute()
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
    'SQL�̍\���`�F�b�N�����{����
    Dim clsDBAccess  As New DBAccess
    clsDBAccess.SQL = strSQL
    Call clsDBAccess.CheckSQL

    '�Z����I��������
    Dim objTopLeftCell As Range
    Set objTopLeftCell = SelectCell("���ʂ�\������Z����I�����Ă�������", Selection)
    If objTopLeftCell Is Nothing Then
        Exit Sub
    Else
        '�I��̈�̍���̃Z����ݒ�
        Set objTopLeftCell = objTopLeftCell.Cells(1)
        
        '���ʂ̃V�[�g��\�����āA���ʂ̃Z����I��
        Call objTopLeftCell.Worksheet.Activate
        Call objTopLeftCell.Select
        DoEvents
    End If
    
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
        .RefEdit.TEXT = objCurrentCell.AddressLocal
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
Public Function ReplaceCellReference(ByRef objSQLCell As Range) As String
On Error GoTo ErrHandle
    Dim objRegExp  As Object
    Dim objMatch   As Object
    Dim strSubSQL  As String
    Dim strReplace As String
    
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True
    objRegExp.Pattern = "\{(.+?)\}"
    
    ReplaceCellReference = DBAccess.DeleteComment(GetRangeText(objSQLCell))
    If objRegExp.Test(ReplaceCellReference) Then
        For Each objMatch In objRegExp.Execute(ReplaceCellReference)
            strReplace = objMatch.SubMatches(0)
            Select Case StrConv(strReplace, vbUpperCase)
            Case "MYPATH"
                ReplaceCellReference = Replace(ReplaceCellReference, objMatch, ActiveWorkbook.Path)
            Case "MYSHEET"
                ReplaceCellReference = Replace(ReplaceCellReference, objMatch, objSQLCell.Worksheet.Name)
            Case Else
                Select Case IsCellAddress(strReplace, objSQLCell.Worksheet)
                Case 1 '����V�[�g�̃Z���̎�
                    '����V�[�g���̃Z���̓��e�Œu������@����FRange("A1")��
                    strSubSQL = ReplaceCellReference(objSQLCell.Worksheet.Range(strReplace))
                    ReplaceCellReference = Replace(ReplaceCellReference, objMatch, strSubSQL)
                Case 2 '�ʃV�[�g�̃Z���̎�
                    '�ʃV�[�g���̃Z���̓��e�Œu������@����FRange("Sheet1!A1")��
                    strSubSQL = ReplaceCellReference(Range(strReplace))
                    ReplaceCellReference = Replace(ReplaceCellReference, objMatch, strSubSQL)
                End Select
            End Select
        Next
    End If
ErrHandle:
End Function

'*****************************************************************************
'[�T�v] strAddress��Cell���w���A�h���X���ǂ���
'[����] �`�F�b�N�Ώۂ̕�����(�A�h���X �܂��� ���O)�A�J�����g�V�[�g
'[�ߒl] 0:�����ȃA�h���X�A1:�J�����g�V�[�g�̃A�h���X�A2:�ʃV�[�g�̃A�h���X
'*****************************************************************************
Private Function IsCellAddress(ByVal strAddress As String, ByRef objWorksheet As Worksheet) As Long
    Dim Dummy As Range
    On Error Resume Next
    Set Dummy = Range(strAddress)
    If Err.Number <> 0 Then
        IsCellAddress = 0 '�G���[�̎��͖����ȃA�h���X
    Else
        Set Dummy = objWorksheet.Range(strAddress)
        If Err.Number <> 0 Then
            IsCellAddress = 2 '�G���[�̎��͕ʃV�[�g�̃A�h���X
        Else
            IsCellAddress = 1 '�G���[�łȂ���΃J�����g�V�[�g�̃A�h���X
        End If
    End If
    On Error GoTo 0
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
        GetRangeText = objRange(1).TEXT
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
'[�T�v] �f�[�^�x�[�X�ڑ��q���쐬����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub MakeConnectStr()
    Const EXCEL = "[EXCEL 12.0;DATABASE={File}]"
    Const ACCESS1 = "[MS ACCESS;DATABASE={File}]"
    Const ACCESS2 = "[MS ACCESS;DATABASE={File};PWD={Password}]"
    Const TEXT = "SELECT * " & vbCrLf & "  FROM [TEXT;DATABASE={Folder}].[{File}]"

On Error GoTo ErrHandle
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Excel,*.xl*,Access,*.md?;*.accdb,�e�L�X�g,*.txt;*.csv,���ׂ�,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim strExt As String
    Dim strFolder As String
    Dim strFile As String
    With CreateObject("Scripting.FileSystemObject")
        strExt = LCase(.GetExtensionName(vDBName))
        strFolder = .GetParentFolderName(vDBName)
        strFile = .GetFileName(vDBName)
    End With
        
    Dim strConnect  As String
    Select Case True
    Case Left(strExt, 2) = "xl"
        strConnect = Replace(EXCEL, "{File}", vDBName)
    Case Left(strExt, 2) = "md" Or strExt = "accdb"
        Dim strPass As String
        strPass = GetPassword(vDBName)
        If strPass = "" Then
            strConnect = Replace(ACCESS1, "{File}", vDBName)
        Else
            strConnect = Replace(ACCESS2, "{File}", vDBName)
            strConnect = Replace(strConnect, "{Password}", strPass)
        End If
    Case Else
        strConnect = Replace(TEXT, "{Folder}", strFolder)
        strConnect = Replace(strConnect, "{File}", strFile)
    End Select
    
    Call SetClipbordText(strConnect)
    Call MsgBox("�ȉ��̃f�[�^�x�[�X�ڑ��q���N���b�v�{�[�h�ɃR�s�[���܂����B" & vbCrLf & strConnect)
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub


