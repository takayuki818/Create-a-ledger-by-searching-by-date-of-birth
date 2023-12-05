Attribute VB_Name = "検索登録"
Option Explicit
Sub 保護切替()
    With ActiveSheet
        Select Case .ProtectContents
            Case True
                .Unprotect
                MsgBox "シート保護を解除しました"
            Case False
                .Protect
                MsgBox "シートを保護しました"
        End Select
    End With
End Sub
Sub 生年月日検索()
    Dim 終行 As Long, 行 As Long, 件数 As Long, 添字 As Long
    Dim 検索日 As Date
    With Sheets("検索登録フォーム")
        検索日 = .Range("検索日")
    End With
    With Sheets("宛名一覧")
        終行 = .Cells(Rows.Count, 2).End(xlUp).Row
        For 行 = 2 To 終行
            If .Cells(行, 1) = 検索日 Then 件数 = 件数 + 1
        Next
        If 件数 > 0 Then
            ReDim 配列(1 To 件数, 1 To 6)
            For 行 = 2 To 終行
                If .Cells(行, 1) = 検索日 Then
                    添字 = 添字 + 1
                    配列(添字, 2) = .Cells(行, 2)
                    Select Case .Cells(行, 4)
                        Case "": 配列(添字, 3) = .Cells(行, 3)
                        Case Else: 配列(添字, 3) = .Cells(行, 3) & "（" & .Cells(行, 4) & "）"
                    End Select
                    配列(添字, 4) = Format(.Cells(行, 5), "000-0000")
                    配列(添字, 5) = .Cells(行, 6)
                    Select Case .Cells(行, 7)
                        Case 1, "1": 配列(添字, 6) = "1.死亡"
                        Case 2, "2": 配列(添字, 6) = "2.転出"
                        Case 3, "3": 配列(添字, 6) = "3.職権消除"
                    End Select
                    If .Cells(行, 8) <> "" Then
                        配列(添字, 6) = 配列(添字, 6) & "(" & Format(.Cells(行, 8), "ge.mm.dd") & ")"
                    End If
                End If
            Next
        End If
    End With
    With Sheets("登録台帳")
        終行 = .Cells(Rows.Count, 3).End(xlUp).Row
        For 添字 = 1 To 件数
            For 行 = 2 To 終行
                If 配列(添字, 2) = .Cells(行, 3) Then
                    配列(添字, 1) = "登録済"
                    Exit For
                End If
            Next
            If 配列(添字, 1) = "" Then 配列(添字, 1) = "未登録"
        Next
    End With
    With Sheets("検索登録フォーム")
        .Cells(9, 1).Resize(Rows.Count - 8, 6).ClearContents
        If 件数 > 0 Then
            .Cells(9, 1).Resize(件数, 6) = 配列
            Else: MsgBox "検索該当なし"
        End If
    End With
End Sub
Sub 生年月日検索クリア()
    With Sheets("検索登録フォーム")
        .Unprotect
        Application.EnableEvents = False
        .Range("検索日コード,検索日").ClearContents
        .Range("検索日コード").Activate
        .Cells(9, 1).Resize(Rows.Count - 8, 6).ClearContents
        Application.EnableEvents = True
        .Protect
    End With
End Sub
Sub 台帳登録()
    Dim 配列(1 To 1, 1 To 8)
    Dim 選択行 As Long, 列 As Long, 終行 As Long
    With Sheets("検索登録フォーム")
        選択行 = .Range("選択行")
        If .Cells(選択行, 1) = "登録済" Then
            If MsgBox("【注意】" & vbCrLf & "既に台帳登録済の宛名番号です" & vbCrLf & vbCrLf & "台帳登録してよろしいですか？", vbYesNo) = vbNo Then Exit Sub
        End If
        配列(1, 1) = .Range("例月区分")
        配列(1, 2) = .Range("管理区分")
        配列(1, 3) = .Cells(選択行, 2)
        If InStr(.Cells(選択行, 3), "（") > 0 Then
            配列(1, 4) = Left(.Cells(選択行, 3), InStr(.Cells(選択行, 3), "（") - 1)
            Else: 配列(1, 4) = .Cells(選択行, 3)
        End If
        配列(1, 5) = .Cells(選択行, 4)
        配列(1, 6) = .Cells(選択行, 5)
        配列(1, 7) = .Range("検索日")
        配列(1, 8) = .Cells(選択行, 6)
    End With
    With Sheets("登録台帳")
        For 列 = 1 To .Cells(1, Columns.Count).End(xlToLeft).Column
            If 終行 < .Cells(Rows.Count, 列).End(xlUp).Row Then 終行 = .Cells(Rows.Count, 列).End(xlUp).Row
        Next
        Range(.Cells(終行 + 1, 1), .Cells(終行 + 1, 8)) = 配列
'        MsgBox "台帳登録完了"
        Call 生年月日検索
    End With
End Sub
