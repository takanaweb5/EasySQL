Attribute VB_Name = "AccessEtc"
Option Explicit

Private Const C_CONNECTSTR = "Provider={Provider};Data Source=""{FileName}"";Jet OLEDB:Database Password={Password};"
'Private Const C_PROVIDER = "Microsoft.Jet.OLEDB.4.0"  'Access2003以前の形式のmdbファイルを作成する時はこちらにする
Private Const C_PROVIDER = "Microsoft.ACE.OLEDB.12.0"
Private Const C_WARNING = "/* 必要に応じて[テーブル名]を変更してからSQLを実行してください */"

'*****************************************************************************
'[ 概  要 ]　データベースファイルを作成する（Accessファイルのみ可）
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub CreateDB()
    Dim strDBName As String
    strDBName = InputBox("作成するAccessファイル名をフルパスで入力してください")
    If strDBName <> "" Then
        Call CreateMDBFile(strDBName)
    End If
End Sub

'*****************************************************************************
'[ 概  要 ]　Accessファイルのテーブル情報を表示する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub ShowTables()
    Dim vDBName As Variant
    Dim objTopLeftCell As Range
    @@@@@@@@@@Dim vTables As Variant
        
    vDBName = Application.GetOpenFilename("Accessファイル,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    Set objTopLeftCell = SelectCell("結果を表示するセルを選択してください", Selection)
    If objTopLeftCell Is Nothing Then
        Exit Sub
    End If
    vTables = GetTableNames(vDBName)
    objTopLeftCell(1).Resize(UBound(vTables, 1) + 1, 2) = vTables
End Sub

'*****************************************************************************
'[ 概  要 ]　データベースの接続文字列を取得する
'[ 引  数 ]　@@@@@@@@@@@@@@@@@@
'[ 戻り値 ]　データベース接続文字列
'*****************************************************************************
Private Function GetConnection(ByVal strFileName As String, ByVal strPassword As String) As String
    GetConnection = C_CONNECTSTR
    GetConnection = Replace(GetConnection, "{Provider}", C_PROVIDER)
    GetConnection = Replace(GetConnection, "{FileName}", strFileName)
    GetConnection = Replace(GetConnection, "{Password}", strPassword)
End Function

'*****************************************************************************
'[ 概  要 ]　MDBファイルを作成する
'[ 引  数 ]　MDBファイル名、パスワード
'[ 戻り値 ]　なし
'*****************************************************************************
Private Sub CreateMDBFile(ByVal strFileName As String, Optional ByVal strPassword As String = "")
    With CreateObject("ADOX.Catalog")
        Call .Create(GetConnection(strFileName, strPassword))
    End With
End Sub

'*****************************************************************************
'[ 概  要 ]　MDBファイルのテーブルの一覧を取得する
'[ 引  数 ]　MDBファイル名、パスワード
'[ 戻り値 ]　テーブル情報の２次元配列
'*****************************************************************************
Private Function GetTableNames(ByVal strFileName As String, Optional ByVal strPassword As String = "") As Variant
    Dim objCatalog As Object
    Dim objTable As Object
    Set objCatalog = CreateObject("ADOX.Catalog")
    objCatalog.ActiveConnection = GetConnection(strFileName, strPassword)
    
    ReDim Result(0 To objCatalog.Tables.Count, 1 To 2)
    
    '見出し設定
    Result(0, 1) = "テーブル名"
    Result(0, 2) = "タイプ"
    
    '明細の設定
    Dim i As Long
    For Each objTable In objCatalog.Tables
        If objTable.Type <> "SYSTEM TABLE" And objTable.Type <> "ACCESS TABLE" Then
            i = i + 1
            Result(i, 1) = objTable.Name
            Result(i, 2) = objTable.Type
        End If
    Next
    GetTableNames = Result
End Function

'*****************************************************************************
'[ 概  要 ]　テーブルインポート用のSQLを作成する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub MakeImportSQL()
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Accessファイル,*.*")
    If vDBName = False Then
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
    
    Dim strDB As String
    strDB = "[MS ACCESS;DATABASE={FileName}].[テーブル名]"
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
End Sub

'*****************************************************************************
'[ 概  要 ]　テーブル削除用のSQLを作成する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Public Sub MakeDeleteTableSQL()
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Accessファイル,*.*")
    If vDBName = False Then
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
    
    Dim strDB As String
    strDB = "[MS ACCESS;DATABASE={FileName}].[テーブル名]"
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
End Sub

'*****************************************************************************
'[ 概  要 ]　ダイアログに出力するメッセージを編集する
'[ 引  数 ]　なし
'[ 戻り値 ]　なし
'*****************************************************************************
Private Function GetMessage(ByVal strSQL As String) As String
    GetMessage = "以下のSQLをクリップボードにコピーしました。" & vbCrLf
    GetMessage = GetMessage & "必要に応じてテーブル名を変更して適用なセルに貼りつけて「SQL実行」コマンドを実行してください。" & vbCrLf
    GetMessage = GetMessage & strSQL
End Function

