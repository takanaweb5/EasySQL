Attribute VB_Name = "AccessEtc"
Option Explicit
Option Private Module

Private Const C_CONNECTSTR = "Provider={Provider};Data Source=""{FileName}"";Jet OLEDB:Database Password={Password};"
'Private Const C_CONNECTSTR = "Provider={Provider};Data Source=""{FileName}"";Jet OLEDB:Database Password={Password};Jet OLEDB:Engine Type=5" 'Access2003以前の形式
'Private Const C_PROVIDER = "Microsoft.Jet.OLEDB.4.0"  'Access2003以前の形式のmdbファイルを作成する時はこちらにする
Private Const C_PROVIDER = "Microsoft.ACE.OLEDB.12.0"
Private Const C_WARNING = "/* [...]部分をテーブル名に変更してからSQLを実行してください */"

'*****************************************************************************
'[概要] データベースファイルを作成する（Accessファイルのみ可）
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub CreateDB()
On Error GoTo ErrHandle
    Dim strDBName As String
    strDBName = InputBox("作成するAccessファイル名をフルパスで入力してください")
    If strDBName <> "" Then
        Call CreateMDBFile(strDBName, InputBox("パスワード設定する場合のみパスワードを入力してください"))
    End If
    Exit Sub
ErrHandle:
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[概要] Accessファイルのテーブル情報を表示する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub ShowTables()
On Error GoTo ErrHandle
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Accessファイル,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim objCatalog As Object
    Dim objTable As Object
    Set objCatalog = CreateObject("ADOX.Catalog")
    objCatalog.ActiveConnection = GetConnection(vDBName)
        
    Dim objTopLeftCell As Range
    Set objTopLeftCell = SelectCell("結果を表示するセルを選択してください", Selection)
    If objTopLeftCell Is Nothing Then
        Exit Sub
    End If
    
    '見出し設定
    objTopLeftCell.Cells(1, 1) = "テーブル名"
    objTopLeftCell.Cells(1, 2) = "タイプ"
    
    '明細の設定
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
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[概要] データベース接続オブジェクトを取得する
'[引数] MDBファイル名、パスワード
'[戻値] データベース接続オブジェクト
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
    
    If InStr(1, strErr, "パスワード") > 0 Then
        Call GetConnection.Open(GetConStr(strFileName, InputBox("パスワードを入力してください")))
    Else
        'エラーの再作成
        Call GetConnection.Open(GetConStr(strFileName))
    End If
End Function

'*****************************************************************************
'[概要] データベースの接続文字列を取得する
'[引数] MDBファイル名、パスワード
'[戻値] データベース接続文字列
'*****************************************************************************
Private Function GetConStr(ByVal strFileName As String, Optional ByVal strPassword As String = "") As String
    GetConStr = C_CONNECTSTR
    GetConStr = Replace(GetConStr, "{Provider}", C_PROVIDER)
    GetConStr = Replace(GetConStr, "{FileName}", strFileName)
    GetConStr = Replace(GetConStr, "{Password}", strPassword)
End Function

'*****************************************************************************
'[概要] MDBファイルを作成する
'[引数] MDBファイル名、パスワード
'[戻値] なし
'*****************************************************************************
Private Sub CreateMDBFile(ByVal strFileName As String, Optional ByVal strPassword As String = "")
    With CreateObject("ADOX.Catalog")
        Call .Create(GetConStr(strFileName, strPassword))
    End With
End Sub

'*****************************************************************************
'[概要] SELECT文のひな型を作成する
'[引数] なし
'[戻値] なし
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
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[概要] テーブルインポート用のSQLを作成する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub MakeImportSQL()
On Error GoTo ErrHandle
    Dim strDB As String
    strDB = GetDatabaseStr()
    If strDB = "" Then
        Exit Sub
    End If
    
    Dim objTable As Range
    Set objTable = SelectCell("インポートするデータ領域を選択してください", Selection)
    If objTable Is Nothing Then
        Exit Sub
    End If

    Dim lngSelect As Long
    Dim strMsg As String
    strMsg = "いずれかを選択して下さい" & vbCrLf
    strMsg = strMsg & "　「 はい 」････ 既存のテーブルに追加する" & vbCrLf
    strMsg = strMsg & "　「いいえ」････ 新しくテーブルを作成する"
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
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[概要] テーブル削除用のSQLを作成する
'[引数] なし
'[戻値] なし
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
    strMsg = "いずれかを選択して下さい" & vbCrLf
    strMsg = strMsg & "　「 はい 」････ テーブルのデータをすべて削除する" & vbCrLf
    strMsg = strMsg & "　「いいえ」････ テーブル自体を削除する"
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
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[概要] クエリ作成用のSQLを作成する
'[引数] なし
'[戻値] なし
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
    
    Call MsgBox(Replace(GetMessage(), "テーブル名", "クエリ名"))
    strSQL = "/* [...]部分を作成するクエリ名に変更し、select_statementの部分にSELECT文を入力してSQLを実行してください */" & vbCrLf & strSQL
    Call SetClipbordText(strSQL)
    Exit Sub
ErrHandle:
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[概要] クエリ削除用のSQLを作成する
'[引数] なし
'[戻値] なし
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
    
    Call MsgBox(Replace(GetMessage(), "テーブル名", "クエリ名"))
    strSQL = Replace(C_WARNING, "テーブル名", "クエリ名") & vbCrLf & strSQL
    Call SetClipbordText(strSQL)
    Exit Sub
ErrHandle:
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[概要] ダイアログに出力するメッセージを編集する
'[引数] なし
'[戻値] ダイアログに出力するメッセージ
'*****************************************************************************
Private Function GetMessage() As String
    GetMessage = "SQLをクリップボードにコピーしました。" & vbCrLf
    GetMessage = GetMessage & " [...]部分をテーブル名に変更してSQLを実行してください。"
End Function

'*****************************************************************************
'[概要] データベース接続識別子を取得する
'[引数] なし
'[戻値] 例：[MS ACCESS;DATABASE=C:\TMP\sample.accdb;PWD=1234].[...]
'*****************************************************************************
Private Function GetDatabaseStr() As String
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Accessファイル,*.*")
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
'[概要] データベースのパスワードを取得する(パスワードの妥当性は未チェック)
'[引数] MDBファイル名
'[戻値] パスワード(パスワード未設定のファイルの時は空の文字列)
'*****************************************************************************
Private Function GetPassword(ByVal strFileName As String) As String
    Dim objConnection As Object
    Set objConnection = CreateObject("ADODB.Connection")
    On Error Resume Next
    Call objConnection.Open(GetConStr(strFileName))
    If Err.Number = 0 Then
        'パスワード未設定の時
        Call objConnection.Close
        Exit Function
    End If
    
    Dim strErr As String
    strErr = Err.Description
    On Error GoTo 0
    
    If InStr(1, strErr, "パスワード") > 0 Then
        GetPassword = InputBox("パスワードを入力してください")
    End If
End Function

