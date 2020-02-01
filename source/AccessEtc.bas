Attribute VB_Name = "AccessEtc"
Option Explicit
Option Private Module

Private Const C_CONNECTSTR = "Provider={Provider};Data Source=""{FileName}"";Jet OLEDB:Database Password={Password};"
'Private Const C_CONNECTSTR = "Provider={Provider};Data Source=""{FileName}"";Jet OLEDB:Database Password={Password};Jet OLEDB:Engine Type=5" 'Access2003�ȑO�̌`��
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
        Call CreateMDBFile(strDBName, InputBox("�p�X���[�h�ݒ肷��ꍇ�̂݃p�X���[�h����͂��Ă�������"))
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
On Error GoTo ErrHandle
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Access�t�@�C��,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim objCatalog As Object
    Dim objTable As Object
    Set objCatalog = CreateObject("ADOX.Catalog")
    objCatalog.ActiveConnection = GetConnection(vDBName)
        
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
'[�T�v] �f�[�^�x�[�X�ڑ��I�u�W�F�N�g���擾����
'[����] MDB�t�@�C�����A�p�X���[�h
'[�ߒl] �f�[�^�x�[�X�ڑ��I�u�W�F�N�g
'*****************************************************************************
Private Function GetConnection(ByVal strFileName As String) As Object
    Set GetConnection = CreateObject("ADODB.Connection")
    On Error Resume Next
    Call GetConnection.Open(GetConStr(strFileName))
    If Err.Number = 0 Then
        Exit Function
    End If
    
    Dim strErr As String
    strErr = Err.Description
    On Error GoTo 0
    
    If InStr(1, strErr, "�p�X���[�h") > 0 Then
        Call GetConnection.Open(GetConStr(strFileName, InputBox("�p�X���[�h����͂��Ă�������")))
    Else
        '�G���[�̍č쐬
        Call GetConnection.Open(GetConStr(strFileName))
    End If
End Function

'*****************************************************************************
'[�T�v] �f�[�^�x�[�X�̐ڑ���������擾����
'[����] MDB�t�@�C�����A�p�X���[�h
'[�ߒl] �f�[�^�x�[�X�ڑ�������
'*****************************************************************************
Private Function GetConStr(ByVal strFileName As String, Optional ByVal strPassword As String = "") As String
    GetConStr = C_CONNECTSTR
    GetConStr = Replace(GetConStr, "{Provider}", C_PROVIDER)
    GetConStr = Replace(GetConStr, "{FileName}", strFileName)
    GetConStr = Replace(GetConStr, "{Password}", strPassword)
End Function

'*****************************************************************************
'[�T�v] MDB�t�@�C�����쐬����
'[����] MDB�t�@�C�����A�p�X���[�h
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub CreateMDBFile(ByVal strFileName As String, Optional ByVal strPassword As String = "")
    With CreateObject("ADOX.Catalog")
        Call .Create(GetConStr(strFileName, strPassword))
    End With
End Sub

'*****************************************************************************
'[�T�v] SELECT���̂ЂȌ^���쐬����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub MakeSelectSQL()
On Error GoTo ErrHandle
    Dim strDB As String
    strDB = GetDatabaseStr()
    If strDB = "" Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "SELECT *" & vbCrLf
    strSQL = strSQL & "  FROM " & strDB
    
    Call MsgBox(GetMessage())
    strSQL = C_WARNING & vbCrLf & strSQL
    Call SetClipbordText(strSQL)
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] �e�[�u���C���|�[�g�p��SQL���쐬����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub MakeImportSQL()
On Error GoTo ErrHandle
    Dim strDB As String
    strDB = GetDatabaseStr()
    If strDB = "" Then
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

    Call MsgBox(GetMessage())
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
    Dim strDB As String
    strDB = GetDatabaseStr()
    If strDB = "" Then
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
    
    Dim strSQL As String
    If lngSelect = vbYes Then
        strSQL = "DELETE FROM " & strDB
    Else
        strSQL = "DROP TABLE " & strDB
    End If
    Call MsgBox(GetMessage())
    strSQL = C_WARNING & vbCrLf & strSQL
    Call SetClipbordText(strSQL)
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] �N�G���쐬�p��SQL���쐬����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub MakeQuerySQL()
On Error GoTo ErrHandle
    Dim strDB As String
    strDB = GetDatabaseStr()
    If strDB = "" Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "CREATE VIEW " & strDB & " AS" & vbCrLf
    strSQL = strSQL & "select_statement"
    
    Call MsgBox(Replace(GetMessage(), "�e�[�u����", "�N�G����"))
    strSQL = "/* [...]�������쐬����N�G�����ɕύX���Aselect_statement�̕�����SELECT������͂���SQL�����s���Ă������� */" & vbCrLf & strSQL
    Call SetClipbordText(strSQL)
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] �N�G���폜�p��SQL���쐬����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub MakeDeleteQuerySQL()
On Error GoTo ErrHandle
    Dim strDB As String
    strDB = GetDatabaseStr()
    If strDB = "" Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "DROP VIEW " & strDB
    
    Call MsgBox(Replace(GetMessage(), "�e�[�u����", "�N�G����"))
    strSQL = Replace(C_WARNING, "�e�[�u����", "�N�G����") & vbCrLf & strSQL
    Call SetClipbordText(strSQL)
    Exit Sub
ErrHandle:
    '�G���[���b�Z�[�W��\��
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[�T�v] �_�C�A���O�ɏo�͂��郁�b�Z�[�W��ҏW����
'[����] �Ȃ�
'[�ߒl] �_�C�A���O�ɏo�͂��郁�b�Z�[�W
'*****************************************************************************
Private Function GetMessage() As String
    GetMessage = "SQL���N���b�v�{�[�h�ɃR�s�[���܂����B" & vbCrLf
    GetMessage = GetMessage & " [...]�������e�[�u�����ɕύX����SQL�����s���Ă��������B"
End Function

'*****************************************************************************
'[�T�v] �f�[�^�x�[�X�ڑ����ʎq���擾����
'[����] �Ȃ�
'[�ߒl] ��F[MS ACCESS;DATABASE=C:\TMP\sample.accdb;PWD=1234].[...]
'*****************************************************************************
Private Function GetDatabaseStr() As String
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Access�t�@�C��,*.*")
    If vDBName = False Then
        Exit Function
    End If
    
    Dim strPassword As String
    strPassword = GetPassword(vDBName)
    
    Dim strDB As String
    If strPassword = "" Then
        strDB = "[MS ACCESS;DATABASE={FileName}].[...]"
    Else
        strDB = "[MS ACCESS;DATABASE={FileName};PWD={Password}].[...]"
        strDB = Replace(strDB, "{Password}", strPassword)
    End If
    GetDatabaseStr = Replace(strDB, "{FileName}", vDBName)
End Function

'*****************************************************************************
'[�T�v] �f�[�^�x�[�X�̃p�X���[�h���擾����(�p�X���[�h�̑Ó����͖��`�F�b�N)
'[����] MDB�t�@�C����
'[�ߒl] �p�X���[�h(�p�X���[�h���ݒ�̃t�@�C���̎��͋�̕�����)
'*****************************************************************************
Private Function GetPassword(ByVal strFileName As String) As String
    Dim objConnection As Object
    Set objConnection = CreateObject("ADODB.Connection")
    On Error Resume Next
    Call objConnection.Open(GetConStr(strFileName))
    If Err.Number = 0 Then
        '�p�X���[�h���ݒ�̎�
        Call objConnection.Close
        Exit Function
    End If
    
    Dim strErr As String
    strErr = Err.Description
    On Error GoTo 0
    
    If InStr(1, strErr, "�p�X���[�h") > 0 Then
        GetPassword = InputBox("�p�X���[�h����͂��Ă�������")
    End If
End Function

