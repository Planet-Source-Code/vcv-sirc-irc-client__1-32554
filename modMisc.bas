Attribute VB_Name = "modMisc"
Public Sub lvSort(lstView As ListView, whichCol As Integer)
    If whichCol > lstView.ColumnHeaders.Count Then Exit Sub
    
    If lstView.SortOrder = lvwAscending Then
        lstView.SortOrder = lvwDescending
    Else
        lstView.SortOrder = lvwAscending
    End If
    
    lstView.SortKey = whichCol - 1
    lstView.Sorted = True
End Sub


