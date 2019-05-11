Attribute VB_Name = "delete_connections"
Sub delete_connections()

' Removes data-connections i.e. references to inserted txt files (reduces file size and improves performance)

Dim i As Long

    On Error Resume Next
    
    For i = ActiveWorkbook.Connections.Count To 1 Step -1
    
        ActiveWorkbook.Connections.Item(i).Delete
        
    Next

End Sub
