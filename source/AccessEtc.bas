Attribute VB_Name = "AccessEtc"
Option Explicit

Private Const C_CONNECTSTR = "Provider={Provider};Data Source=""{FileName}"";Jet OLEDB:Database Password={Password};"
'Private Const C_PROVIDER = "Microsoft.Jet.OLEDB.4.0"  'Access2003�ȑO�̌`����mdb�t�@�C�����쐬���鎞�͂�����ɂ���
Private Const C_PROVIDER = "Microsoft.ACE.OLEDB.12.0"
Private Const C_WARNING = "/* [...]�������e�[�u�����ɕύX���Ă���SQL�����s���Ă������� */"

'*****************************************************************************
'[�T�v] �f�[�^�x�[�X�t�@�C�����쐬����iAccess�t�@�C���̂݉j
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub CreateDB()
On Error GoTo ErrHandle
    Dim strDBName As String
    strDBName = InputBox("�쐬����Access�t�@�C�������t���p�X�œ��͂��Ă�������")
    If strDBName <> "" Then
        Call CreateMDBFile(strDBName)
    End If
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] Access�t�@�C���̃e�[�u������\������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ShowTables()
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Access�t�@�C��,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim objCatalog As Object
    Dim objTable As Object
    Set objCatalog = CreateObject("ADOX.Catalog")
        
    On Error Resume Next
    objCatalog.ActiveConnection = GetConnection(vDBName)
    Dim strErr As String
    strErr = Err.Description
    On Error GoTo ErrHandle
    
    If strErr <> "" Then
        If InStr(1, strErr, "�p�X���[�h") > 0 Then
            objCatalog.ActiveConnection = GetConnection(vDBName, InputBox("�p�X���[�h����͂��Ă�������"))
        Else
            '�G���[�̍č쐬
            objCatalog.ActiveConnection = GetConnection(vDBName)
        End If
    End If
    
    Dim objTopLeftCell As Range
    Set objTopLeftCell = SelectCell("���ʂ�\������Z����I�����Ă�������", Selection)
    If objTopLeftCell Is Nothing Then
        Exit Sub
    End If
    
    '���o���ݒ�
    objTopLeftCell.Cells(1, 1) = "�e�[�u����"
    objTopLeftCell.Cells(1, 2) = "�^�C�v"
    
    '���ׂ̐ݒ�
    Dim i As Long
    i = 1
    For Each objTable In objCatalog.Tables
        If objTable.Type <> "SYSTEM TABLE" And objTable.Type <> "ACCESS TABLE" Then
            i = i + 1
            objTopLeftCell.Cells(i, 1) = objTable.Name
            objTopLeftCell.Cells(i, 2) = objTable.Type
        End If
    Next
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] �f�[�^�x�[�X�̐ڑ���������擾����
'[����] MDB�t�@�C�����A�p�X���[�h
'[�ߒl] �f�[�^�x�[�X�ڑ�������
'*****************************************************************************
Private Function GetConnection(ByVal strFileName As String, Optional ByVal strPassword As String = "") As String
    GetConnection = C_CONNECTSTR
    GetConnection = Replace(GetConnection, "{Provider}", C_PROVIDER)
    GetConnection = Replace(GetConnection, "{FileName}", strFileName)
    GetConnection = Replace(GetConnection, "{Password}", strPassword)
End Function

'*****************************************************************************
'[�T�v] MDB�t�@�C�����쐬����
'[����] MDB�t�@�C�����A�p�X���[�h
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub CreateMDBFile(ByVal strFileName As String, Optional ByVal strPassword As String = "")
    With CreateObject("ADOX.Catalog")
        Call .Create(GetConnection(strFileName, strPassword))
    End With
End Sub

'*****************************************************************************
'[�T�v] �e�[�u���C���|�[�g�p��SQL���쐬����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub MakeImportSQL()
On Error GoTo ErrHandle
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Access�t�@�C��,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim objTable As Range
    Set objTable = SelectCell("�C���|�[�g����f�[�^�̈��I�����Ă�������", Selection)
    If objTable Is Nothing Then
        Exit Sub
    End If

    Dim lngSelect As Long
    Dim strMsg As String
    strMsg = "�����ꂩ��I�����ĉ�����" & vbCrLf
    strMsg = strMsg & "�@�u �͂� �v���� �����̃e�[�u���ɒǉ�����" & vbCrLf
    strMsg = strMsg & "�@�u�������v���� �V�����e�[�u�����쐬����"
    lngSelect = MsgBox(strMsg, vbYesNoCancel + vbDefaultButton2)
    If lngSelect = vbCancel Then
        Exit Sub
    End If
    
    Dim strDB As String
    strDB = "[MS ACCESS;DATABASE={FileName}].[...]"
    strDB = Replace(strDB, "{FileName}", vDBName)
    
    Dim strFROM As String
    strFROM = "[{Sheet}${Range}]"
    strFROM = Replace(strFROM, "{Sheet}", objTable.Worksheet.Name)
    strFROM = Replace(strFROM, "{Range}", objTable.AddressLocal(False, False, xlA1))
    
    Dim strSQL As String
    If lngSelect = vbYes Then
        strSQL = "INSERT INTO " & strDB & vbCrLf
        strSQL = strSQL & "SELECT *" & vbCrLf
        strSQL = strSQL & "  FROM " & strFROM
    Else
        strSQL = "SELECT *" & vbCrLf
        strSQL = strSQL & "  INTO " & strDB & vbCrLf
        strSQL = strSQL & "  FROM " & strFROM
    End If

    Call MsgBox(GetMessage(strSQL))
    strSQL = C_WARNING & vbCrLf & strSQL
    Call SetClipbordText(strSQL)
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] �e�[�u���폜�p��SQL���쐬����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub MakeDeleteTableSQL()
On Error GoTo ErrHandle
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Access�t�@�C��,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim lngSelect As Long
    Dim strMsg As String
    strMsg = "�����ꂩ��I�����ĉ�����" & vbCrLf
    strMsg = strMsg & "�@�u �͂� �v���� �e�[�u���̃f�[�^�����ׂč폜����" & vbCrLf
    strMsg = strMsg & "�@�u�������v���� �e�[�u�����̂��폜����"
    lngSelect = MsgBox(strMsg, vbYesNoCancel + vbDefaultButton2)
    If lngSelect = vbCancel Then
        Exit Sub
    End If
    
    Dim strDB As String
    strDB = "[MS ACCESS;DATABASE={FileName}].[...]"
    strDB = Replace(strDB, "{FileName}", vDBName)
    
    Dim strSQL As String
    If lngSelect = vbYes Then
        strSQL = "DELETE FROM " & strDB
    Else
        strSQL = "DROP TABLE " & strDB
    End If
    Call MsgBox(GetMessage(strSQL))
    strSQL = C_WARNING & vbCrLf & strSQL
    Call SetClipbordText(strSQL)
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] �_�C�A���O�ɏo�͂��郁�b�Z�[�W��ҏW����
'[����] SQL
'[�ߒl] �_�C�A���O�ɏo�͂��郁�b�Z�[�W
'*****************************************************************************
Private Function GetMessage(ByVal strSQL As String) As String
    GetMessage = "�ȉ���SQL���N���b�v�{�[�h�ɃR�s�[���܂����B" & vbCrLf
    GetMessage = GetMessage & " [...]�������e�[�u�����ɕύX����SQL�����s���Ă��������B" & vbCrLf
    GetMessage = GetMessage & strSQL
End Function

