# excel-vba-automation
Excel VBA




Rem delete unnecessary rows

Sub del_unnecessary()
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    For x = last_row To 1 Step -1
        If Cells(x, 1) = "Buy Now" Or Cells(x, 1) = "" Or Cells(x, 1) = "img" Then
        Rows(x).Delete
        End If
    Next x


    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    For x = last_row To 1 Step -1
    If x > 1 Then
     If Left(Cells(x, 1), 2) = "‡ß" And Left(Cells(x - 1, 1), 2) = "‡ß" Then Rows(x - 1).Delete
    End If
    Next x
End Sub


Sub row_to_col()
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    For x = last_row To 1 Step -1
    If x > 1 Then
     If Left(Cells(x, 1), 2) = "‡ß" Then
        Cells(x - 1, 2) = Cells(x, 1)
        Rows(x).Delete
     End If
    End If
    Next
End Sub
