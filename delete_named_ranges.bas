Attribute VB_Name = "delete_named_ranges"
Public Sub Ranges_Named_Delete(wb As Workbook)

    Dim i As Long

    ' On Error Resume Next ' нужно избегать. Чем быстрее обнаружится ошибка, тем она быстрее испрвится

    With wb

        For i = .Names.Count To 1 Step -1

            .Names(i).Delete

        Next
    End With
End Sub
