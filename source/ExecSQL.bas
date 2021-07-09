Attribute VB_Name = "ExecSQL"
Option Explicit
Option Private Module

'*****************************************************************************
'[概要] 選択されたセルを元にSELECT文のひな形を作成する
'[引数] なし
'[戻値] なし
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
        
    'SELECT句の設定
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

    'FROM句の設定
    If objSelection.Areas.Count = 1 And objSelection.Rows.Count > 1 Then
        strFROM = Replace("FROM [{Sheet}${Range}]", "{Sheet}", objSelection.Worksheet.Name)
        strFROM = Replace(strFROM, "{Range}", objSelection.AddressLocal(False, False, xlA1))
    Else
        strFROM = Replace("FROM [{Sheet}$]", "{Sheet}", objSelection.Worksheet.Name)
    End If
    
    'その他の句の識別子のみ設定
    strSQL = strSELECT & vbCrLf & _
               strFROM & vbCrLf & _
               "WHERE " & vbCrLf & _
               "GROUP BY" & vbCrLf & _
               "HAVING " & vbCrLf & _
               "ORDER BY"
    
    Call SetClipbordText(strSQL)
    Dim strMsg As String
    strMsg = "以下のSQLをクリップボードにコピーしました。" & vbCrLf & strSQL
    Call MsgBox(strMsg)
End Sub

'*****************************************************************************
'[概要] クリップボードにテキストを設定する
'[引数] 設定する文字列
'[戻値] なし
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
'[概要] SQLを実行する
'[引数] 1:Select文用,2:更新系用
'[戻値] なし
'*****************************************************************************
Public Sub ExecuteSQL1()
    Call ExecuteSQL(True)
End Sub
Public Sub ExecuteSQL2()
    Call ExecuteSQL(False)
End Sub

'*****************************************************************************
'[概要] SQL文を実行する
'[引数] True:Select文、False:更新系SQL
'[戻値] なし
'*****************************************************************************
Private Sub ExecuteSQL(ByVal IsSelect As Boolean)
    If ActiveWorkbook.Path = "" Then
        Call MsgBox("一度も保存されていないファイルはエラーになることがあります")
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
'[概要] DDLまたはDMLのSQLを実行する
'[引数] SQL
'[戻値] なし
'*****************************************************************************
Private Sub Execute(ByVal strSQL As String)
On Error GoTo ErrHandle
    'SQLの構文チェックを実施する
    Dim clsDBAccess  As New DBAccess
    clsDBAccess.SQL = strSQL
    Call clsDBAccess.CheckSQL
    
    'コマンドを実行
    Dim dblTime As Double
    Dim lngRecCount As Long
    dblTime = Timer()
    lngRecCount = clsDBAccess.Execute()
    Call MsgBox("更新件数は " & lngRecCount & " 件です" & vbCrLf & vbCrLf & _
                "実行時間：" & Int(Timer() - dblTime) & " 秒")
    Exit Sub
ErrHandle:
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[概要] SELECT文の結果を表形式でセルに展開する
'[引数] SQL
'[戻値] なし
'*****************************************************************************
Private Sub ShowRecord(ByVal strSQL As String)
On Error GoTo ErrHandle
    'SQLの構文チェックを実施する
    Dim clsDBAccess  As New DBAccess
    clsDBAccess.SQL = strSQL
    Call clsDBAccess.CheckSQL

    'セルを選択させる
    Dim objTopLeftCell As Range
    Set objTopLeftCell = SelectCell("結果を表示するセルを選択してください", Selection)
    If objTopLeftCell Is Nothing Then
        Exit Sub
    Else
        '選択領域の左上のセルを設定
        Set objTopLeftCell = objTopLeftCell.Cells(1)
        
        '結果のシートを表示して、結果のセルを選択
        Call objTopLeftCell.Worksheet.Activate
        Call objTopLeftCell.Select
        DoEvents
    End If
    
    'SELECT文の実行結果のレコードセットをセルに設定
    Dim dblTime As Double
    Dim lngRecCount As Long
    dblTime = Timer()
    lngRecCount = clsDBAccess.ExecuteToRange(objTopLeftCell)
    Call MsgBox("レコード件数は " & lngRecCount & " 件です" & vbCrLf & vbCrLf & _
                "実行時間：" & Int(Timer() - dblTime) & " 秒")
    Exit Sub
ErrHandle:
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub

'*****************************************************************************
'[概要] フォームを表示してセルを選択させる
'[引数] 表示するメッセージ、objCurrentCell：初期選択させるセル
'[戻値] 選択されたセル（キャンセル時はNothing）
'*****************************************************************************
Public Function SelectCell(ByVal strMsg As String, ByRef objCurrentCell As Range) As Range
    Dim strCell As String
    'フォームを表示
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
'[概要] SQLの{A1}の部分をA1セルの中身で置換する
'       ただし、{MYPATH}の部分はカレントフォルダに置換える
'               {MYSHEET}の部分はSQLのあるシート名に置換える
'[引数] SQLの入力させたセル
'[戻値] セルの参照値を置換したSQL
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
                Case 1 '同一シートのセルの時
                    '同一シート内のセルの内容で置換える　※例：Range("A1")等
                    strSubSQL = ReplaceCellReference(objSQLCell.Worksheet.Range(strReplace))
                    ReplaceCellReference = Replace(ReplaceCellReference, objMatch, strSubSQL)
                Case 2 '別シートのセルの時
                    '別シート内のセルの内容で置換える　※例：Range("Sheet1!A1")等
                    strSubSQL = ReplaceCellReference(Range(strReplace))
                    ReplaceCellReference = Replace(ReplaceCellReference, objMatch, strSubSQL)
                End Select
            End Select
        Next
    End If
ErrHandle:
End Function

'*****************************************************************************
'[概要] strAddressがCellを指すアドレスかどうか
'[引数] チェック対象の文字列(アドレス または 名前)、カレントシート
'[戻値] 0:無効なアドレス、1:カレントシートのアドレス、2:別シートのアドレス
'*****************************************************************************
Private Function IsCellAddress(ByVal strAddress As String, ByRef objWorksheet As Worksheet) As Long
    Dim Dummy As Range
    On Error Resume Next
    Set Dummy = Range(strAddress)
    If Err.Number <> 0 Then
        IsCellAddress = 0 'エラーの時は無効なアドレス
    Else
        Set Dummy = objWorksheet.Range(strAddress)
        If Err.Number <> 0 Then
            IsCellAddress = 2 'エラーの時は別シートのアドレス
        Else
            IsCellAddress = 1 'エラーでなければカレントシートのアドレス
        End If
    End If
    On Error GoTo 0
End Function

'*****************************************************************************
'[概要] SQLの選択されたセルの値を取得する
'[引数] SQLの入力させたRange
'[戻値] セルの値（複数行の時：値が初期値でないセルの値を改行で連結）
'               （複数列の時：値が初期値でないセルの値を[]ではさみ","で連結）
'*****************************************************************************
Private Function GetRangeText(ByRef objRange As Range) As String
    Dim i As Long
    Dim strCellText As String
    
    '単一セルの時
    If objRange.Count = 1 Or objRange.Address = objRange(1, 1).MergeArea.Address Then
        GetRangeText = objRange(1).TEXT
        Exit Function
    End If
    
    '見出し選択時
    If objRange.Rows.Count = 1 Or objRange.Rows.Count = Rows.Count Then
        '列数分LOOPし、各項目をカッコではさみコンマで連結　例：[姓], [名]
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
        '行数分LOOPし、各セルの値を改行で連結
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
'[概要] データベース接続子を作成する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub MakeConnectStr()
    Const EXCEL = "[EXCEL 12.0;DATABASE={File}]"
    Const ACCESS1 = "[MS ACCESS;DATABASE={File}]"
    Const ACCESS2 = "[MS ACCESS;DATABASE={File};PWD={Password}]"
    Const TEXT = "SELECT * " & vbCrLf & "  FROM [TEXT;DATABASE={Folder}].[{File}]"

On Error GoTo ErrHandle
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("Excel,*.xl*,Access,*.md?;*.accdb,テキスト,*.txt;*.csv,すべて,*.*")
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
    Call MsgBox("以下のデータベース接続子をクリップボードにコピーしました。" & vbCrLf & strConnect)
    Exit Sub
ErrHandle:
    'エラーメッセージを表示
    Call MsgBox(Err.Description)
End Sub


