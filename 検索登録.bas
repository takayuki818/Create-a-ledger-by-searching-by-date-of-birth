Attribute VB_Name = "�����o�^"
Option Explicit
Sub �ی�ؑ�()
    With ActiveSheet
        Select Case .ProtectContents
            Case True
                .Unprotect
                MsgBox "�V�[�g�ی���������܂���"
            Case False
                .Protect
                MsgBox "�V�[�g��ی삵�܂���"
        End Select
    End With
End Sub
Sub ���N��������()
    Dim �I�s As Long, �s As Long, ���� As Long, �Y�� As Long
    Dim ������ As Date
    With Sheets("�����o�^�t�H�[��")
        ������ = .Range("������")
    End With
    With Sheets("�����ꗗ")
        �I�s = .Cells(Rows.Count, 2).End(xlUp).Row
        For �s = 2 To �I�s
            If .Cells(�s, 1) = ������ Then ���� = ���� + 1
        Next
        If ���� > 0 Then
            ReDim �z��(1 To ����, 1 To 6)
            For �s = 2 To �I�s
                If .Cells(�s, 1) = ������ Then
                    �Y�� = �Y�� + 1
                    �z��(�Y��, 2) = .Cells(�s, 2)
                    Select Case .Cells(�s, 4)
                        Case "": �z��(�Y��, 3) = .Cells(�s, 3)
                        Case Else: �z��(�Y��, 3) = .Cells(�s, 3) & "�i" & .Cells(�s, 4) & "�j"
                    End Select
                    �z��(�Y��, 4) = Format(.Cells(�s, 5), "000-0000")
                    �z��(�Y��, 5) = .Cells(�s, 6)
                    Select Case .Cells(�s, 7)
                        Case 1, "1": �z��(�Y��, 6) = "1.���S"
                        Case 2, "2": �z��(�Y��, 6) = "2.�]�o"
                        Case 3, "3": �z��(�Y��, 6) = "3.�E������"
                    End Select
                    If .Cells(�s, 8) <> "" Then
                        �z��(�Y��, 6) = �z��(�Y��, 6) & "(" & Format(.Cells(�s, 8), "ge.mm.dd") & ")"
                    End If
                End If
            Next
        End If
    End With
    With Sheets("�o�^�䒠")
        �I�s = .Cells(Rows.Count, 3).End(xlUp).Row
        For �Y�� = 1 To ����
            For �s = 2 To �I�s
                If �z��(�Y��, 2) = .Cells(�s, 3) Then
                    �z��(�Y��, 1) = "�o�^��"
                    Exit For
                End If
            Next
            If �z��(�Y��, 1) = "" Then �z��(�Y��, 1) = "���o�^"
        Next
    End With
    With Sheets("�����o�^�t�H�[��")
        .Cells(9, 1).Resize(Rows.Count - 8, 6).ClearContents
        If ���� > 0 Then
            .Cells(9, 1).Resize(����, 6) = �z��
            Else: MsgBox "�����Y���Ȃ�"
        End If
    End With
End Sub
Sub ���N���������N���A()
    With Sheets("�����o�^�t�H�[��")
        .Unprotect
        Application.EnableEvents = False
        .Range("�������R�[�h,������").ClearContents
        .Range("�������R�[�h").Activate
        .Cells(9, 1).Resize(Rows.Count - 8, 6).ClearContents
        Application.EnableEvents = True
        .Protect
    End With
End Sub
Sub �䒠�o�^()
    Dim �z��(1 To 1, 1 To 8)
    Dim �I���s As Long, �� As Long, �I�s As Long
    With Sheets("�����o�^�t�H�[��")
        �I���s = .Range("�I���s")
        If .Cells(�I���s, 1) = "�o�^��" Then
            If MsgBox("�y���Ӂz" & vbCrLf & "���ɑ䒠�o�^�ς̈����ԍ��ł�" & vbCrLf & vbCrLf & "�䒠�o�^���Ă�낵���ł����H", vbYesNo) = vbNo Then Exit Sub
        End If
        �z��(1, 1) = .Range("�ጎ�敪")
        �z��(1, 2) = .Range("�Ǘ��敪")
        �z��(1, 3) = .Cells(�I���s, 2)
        If InStr(.Cells(�I���s, 3), "�i") > 0 Then
            �z��(1, 4) = Left(.Cells(�I���s, 3), InStr(.Cells(�I���s, 3), "�i") - 1)
            Else: �z��(1, 4) = .Cells(�I���s, 3)
        End If
        �z��(1, 5) = .Cells(�I���s, 4)
        �z��(1, 6) = .Cells(�I���s, 5)
        �z��(1, 7) = .Range("������")
        �z��(1, 8) = .Cells(�I���s, 6)
    End With
    With Sheets("�o�^�䒠")
        For �� = 1 To .Cells(1, Columns.Count).End(xlToLeft).Column
            If �I�s < .Cells(Rows.Count, ��).End(xlUp).Row Then �I�s = .Cells(Rows.Count, ��).End(xlUp).Row
        Next
        Range(.Cells(�I�s + 1, 1), .Cells(�I�s + 1, 8)) = �z��
'        MsgBox "�䒠�o�^����"
        Call ���N��������
    End With
End Sub
