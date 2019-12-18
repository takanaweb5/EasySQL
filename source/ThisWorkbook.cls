VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'*****************************************************************************
'[イベント] AddinInstall
'[概要] アドインのインストール時に実行される
'*****************************************************************************
Private Sub Workbook_AddinInstall()
    Call Workbook_Open
End Sub

'*****************************************************************************
'[イベント] AddinUninstall
'[概要] アドインのアンインストール時に実行される
'*****************************************************************************
Private Sub Workbook_AddinUninstall()
    'ツールバーを削除し、セル関数のHELPも削除
    On Error Resume Next
    Call Application.CommandBars("EasySQL").Delete
    Call UninstHelpMsg
End Sub

'*****************************************************************************
'[イベント] WorkbookOpen
'[概要] ファイルを開く時に実行される
'*****************************************************************************
Private Sub Workbook_Open()
    On Error Resume Next
    Application.CommandBars("EasySQL").Visible = True
    If Err.Number = 0 Then
        'アドインでToolbarの存在が確認された時は何もしない
        'EXCEL2013以降だとアドインがインストールされていてもTOOLBARが
        '認識されない不具合があるのでその時はErr.Number <> 0となる
        Exit Sub
    End If
    
    Call InstToolBar
    Dim i As Long
    For i = 1 To AddIns.Count
        If UCase(AddIns(i).Name) = UCase(ThisWorkbook.Name) Then
            If AddIns(i).Installed = True Then
                'アドインがインストールされていればHELPの再設定は不要
                Exit Sub
            End If
        End If
    Next i
    Call InstHelpMsg
End Sub

'*****************************************************************************
'[イベント] Workbook_BeforeClose
'[概要] ファイルを閉じる時に実行される
'*****************************************************************************
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim i As Long
    'アドインなら、何にもしない
    For i = 1 To AddIns.Count
        If UCase(AddIns(i).Name) = UCase(ThisWorkbook.Name) Then
            If AddIns(i).Installed = True Then
                Exit Sub
            End If
        End If
    Next i
    
    'ツールバーを削除し、セル関数のHELPも削除
    On Error Resume Next
    Call Application.CommandBars("EasySQL").Delete
    Call UninstHelpMsg
End Sub

'*****************************************************************************
'[概要] ツールバーのインストール
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub InstToolBar()
    On Error Resume Next
    Call Application.CommandBars("EasySQL").Delete
    
    Dim objCmdBar   As CommandBar
    Set objCmdBar = Application.CommandBars.Add("EasySQL")
    
    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "SQL実行"
        .OnAction = "ExecuteSQL"
        .TooltipText = "SQLを実行する"
        .Style = msoButtonCaption
    End With

    With objCmdBar.Controls.Add(msoControlButton)
        .Caption = "SQL作成"
        .OnAction = "MakeSQL"
        .TooltipText = "SQLのひな形を作成する"
        .Style = msoButtonCaption
    End With
    
    With objCmdBar.Controls.Add(msoControlPopup)
        .Caption = "Access"
        With .Controls.Add(msoControlButton)
            .Caption = "データベース作成"
            .OnAction = "CreateDB"
            .TooltipText = "Accessのデータベースファイルを作成する"
            .Style = msoButtonCaption
        End With

        With .Controls.Add(msoControlButton)
            .Caption = "テーブル一覧表示"
            .OnAction = "ShowTables"
            .TooltipText = "Accessデータベースのテーブルの一覧を表示する"
            .Style = msoButtonCaption
        End With
        
        With .Controls.Add(msoControlButton)
            .Caption = "インポート用SQL作成"
            .OnAction = "MakeImportSQL"
            .TooltipText = "テーブルインポート用のSQLを作成する"
            .Style = msoButtonCaption
        End With

        With .Controls.Add(msoControlButton)
            .Caption = "削除用SQL作成"
            .OnAction = "MakeDeleteTableSQL"
            .TooltipText = "テーブル削除用のSQLを作成する"
            .Style = msoButtonCaption
        End With
    End With
End Sub

'*****************************************************************************
'[概要] ユーザ定義セル関数のHELPを設定
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub InstHelpMsg()
    Dim strMsg As String
    Dim strInMsg(1 To 3) As String
    On Error Resume Next
    
    strMsg = "以下の例のようにセルの値を区切り文字で連結します" & vbLf
    strMsg = strMsg & "　　'AAA','BBB','CCC'　　や　　[AAA],[BBB],[CCC]" & vbLf
    strMsg = strMsg & "SQLのIN演算子の条件の羅列などに利用すると便利です"
    strInMsg(1) = "：　'　や　[]"
    strInMsg(2) = "：　,　など"
    strInMsg(3) = "：値を連結したいセル範囲を指定します(複数指定可)"
    Call Application.MacroOptions("VALUEJOIN", strMsg, , , , , "EasySQL", , , , Array(strInMsg(1), strInMsg(2), strInMsg(3)))
    
    strMsg = "SQLの実行結果を2次元配列で返します" & vbLf
    strMsg = strMsg & "範囲を指定して配列数式形式(Ctrl+Shift+Enter)で取り出してください"
    strInMsg(1) = "：SQLの入力されたセル"
    strInMsg(2) = "：再計算のトリガーにしたいセルがあれば指定します(いくつでも可)"
    Call Application.MacroOptions("GetSQLRecordset", strMsg, , , , , "EasySQL", , , , Array(strInMsg(1), strInMsg(2)))
    
    strMsg = "セル参照の反映とコメント削除を実行した後の、データベースに渡すSQLを表示します"
    strInMsg(1) = "：SQLの入力されたセル"
    Call Application.MacroOptions("ReplaceCellRef", strMsg, , , , , "EasySQL", , , , Array(strInMsg(1)))
End Sub

'*****************************************************************************
'[概要] ユーザ定義セル関数のHELPを削除
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub UninstHelpMsg()
    On Error Resume Next
    Call Application.MacroOptions("VALUEJOIN", "", , , , , 14, , , , Array(""))
    Call Application.MacroOptions("GetSQLRecordset", "", , , , , 14, , , , Array(""))
    Call Application.MacroOptions("ReplaceCellRef", "", , , , , 14, , , , Array(""))
    '14：ユーザ定義(デフォルト値)
End Sub


