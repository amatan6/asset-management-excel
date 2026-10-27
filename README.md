Private Sub Worksheet_Change(ByVal Target As Range)
    ' スキャン用セルの位置（F2セルをスキャン入力欄に設定）
    Dim scanCell As Range
    Set scanCell = Me.Range("F2")
    
    ' F2セルが変更されたときのみ実行
    If Not Intersect(Target, scanCell) Is Nothing Then
        If scanCell.Value = "" Then Exit Sub
        
        Dim searchVal As String
        searchVal = CStr(scanCell.Value)
        
        ' 連続動作の制御と画面チラつき防止
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        Dim foundCell As Range
        ' A列（資産番号）から一致するものを検索
        Set foundCell = Me.Columns("A").Find(What:=searchVal, LookAt:=xlWhole, LookIn:=xlValues)
        
        If Not foundCell Is Nothing Then
            Dim statusCell As Range
            Set statusCell = Me.Cells(foundCell.Row, "C") ' C列がステータス
            
            ' ステータスの交互切り替え
            If statusCell.Value = "在庫有り" Then
                statusCell.Value = "貸し出し中"
            Else
                statusCell.Value = "在庫有り"
            End If
            
            ' 最終更新日時を記録（D列）
            Me.Cells(foundCell.Row, "D").Value = Now
            
        Else
            MsgBox "資産番号「 " & searchVal & " 」が見つかりませんでした。", vbExclamation, "エラー"
        End If
        
        ' スキャン用セルをクリアして再選択
        scanCell.Value = ""
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        scanCell.Select
    End If
End Sub
