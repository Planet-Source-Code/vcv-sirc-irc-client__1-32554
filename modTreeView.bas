Attribute VB_Name = "modTreeView"
Public serverNodes()    As Node
Public serverCount      As Integer

Public Function treeview_AddChannel(tvList As TreeView, strChannel As String, serverID As Integer)
    If treeview_GetChannelIndex(tvList, strChannel, serverID) <> -1 Then
        'Add something another time..
        Exit Function
    End If
    
    Dim i As Integer, curServer As Integer, newChanGroup As Node, newChannel As Node, nParent As Integer
    curServer = -1
    For i = 1 To tvList.Nodes.Count
        If tvList.Nodes.item(i).parent Is Nothing Then
            If curServer <> -1 Then
                '* No channel list found, add the fucker, and add the channel
                Set newChanGroup = tvList.Nodes.Add(curServer, tvwChild, , "Channels", 4)
                newChanGroup.EnsureVisible
                newChanGroup.Expanded = True
                newChanGroup.Sorted = True
                Set newChannel = tvList.Nodes.Add(newChanGroup.Index, tvwChild, , strChannel, 2)
                Exit Function
            Else
                If LeftOf(tvList.Nodes.item(i).Text, ":") = serverID Then
                    curServer = i
                End If
            End If
        Else
            If tvList.Nodes.item(i).Text = "Channels" And curServer <> -1 Then
                Set newChannel = tvList.Nodes.Add(i, tvwChild, , strChannel, 2)
                Exit Function
            End If
        End If
    Next i
    
    If curServer <> -1 Then
        Set newChanGroup = tvList.Nodes.Add(curServer, tvwChild, , "Channels", 4)
        newChanGroup.EnsureVisible
        newChanGroup.Expanded = True
        newChanGroup.Sorted = True
        Set newChannel = tvList.Nodes.Add(newChanGroup.Index, tvwChild, , strChannel, 2)
        Exit Function
    End If
End Function


Public Sub treeview_NewServer(tvList As TreeView, strHostAdd As String, serverID As Integer)

    Dim i As Integer
    For i = 1 To serverCount
        If serverNodes(i) Is Nothing Then
            Set serverNodes(i) = tvList.Nodes.Add(, , , serverID & ": " & strHostAdd, 6)
            serverNodes(i).BOLD = True
            serverNodes(i).EnsureVisible
            serverNodes(i).Expanded = True
            Exit Sub
        End If
    Next i

    serverCount = serverCount + 1
    ReDim serverNodes(serverCount) As Node
    Set serverNodes(serverCount) = tvList.Nodes.Add(, , , serverID & ": " & strHostAdd, 6)
    serverNodes(serverCount).BOLD = True
    serverNodes(serverCount).EnsureVisible
    serverNodes(serverCount).Expanded = True
    
End Sub


Public Function treeview_AddQuery(tvList As TreeView, strQuery As String, serverID As Integer)
    If treeview_GetQueryIndex(tvList, strQuery, serverID) <> -1 Then
        'Add something another time..
        'this means the channel is already there!
        Exit Function
    End If
    
    Dim i As Integer, curServer As Integer, newQueryGroup As Node, newQuery As Node, nParent As Integer
    curServer = -1
    For i = 1 To tvList.Nodes.Count
        If tvList.Nodes.item(i).parent Is Nothing Then
            If curServer <> -1 Then
                '* No channel list found, add the fucker, and add the channel
                Set newQueryGroup = tvList.Nodes.Add(curServer, tvwChild, , "Queries", 5)
                newQueryGroup.EnsureVisible
                newQueryGroup.Expanded = True
                newQueryGroup.Sorted = True
                Set newQuery = tvList.Nodes.Add(newQueryGroup.Index, tvwChild, , strQuery, 3)
                Exit Function
            Else
                If LeftOf(tvList.Nodes.item(i).Text, ":") = serverID Then
                    curServer = i
                End If
            End If
        Else
            If tvList.Nodes.item(i).Text = "Queries" And curServer <> -1 Then
                Set newQuery = tvList.Nodes.Add(i, tvwChild, , strQuery, 3)
                Exit Function
            End If
        End If
    Next i
    
    If curServer <> -1 Then
        Set newQueryGroup = tvList.Nodes.Add(curServer, tvwChild, , "Queries", 5)
        newQueryGroup.EnsureVisible
        newQueryGroup.Expanded = True
        newQueryGroup.Sorted = True
        Set newQuery = tvList.Nodes.Add(newQueryGroup.Index, tvwChild, , strQuery, 3)
        Exit Function
    End If
End Function


Public Function treeview_GetQueryIndex(tvList As TreeView, strQuery As String, serverID As Integer) As Integer
    Dim i As Integer
    
    For i = 1 To tvList.Nodes.Count
        If tvList.Nodes.item(i).parent Is Nothing Then
        ElseIf tvList.Nodes.item(i).parent.Text = "Queries" Then
            If tvList.Nodes.item(i).parent.parent Is Nothing Then
            ElseIf LeftOf(tvList.Nodes.item(i).parent.parent.Text, ":") = serverID Then
                If tvList.Nodes.item(i).Text = strQuery Then
                    treeview_GetQueryIndex = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    treeview_GetQueryIndex = -1
End Function

Public Function treeview_GetChannelIndex(tvList As TreeView, strChannel As String, serverID As Integer) As Integer
    Dim i As Integer
    
    For i = 1 To tvList.Nodes.Count
        If tvList.Nodes.item(i).parent Is Nothing Then
        ElseIf tvList.Nodes.item(i).parent.Text = "Channels" Then
            If tvList.Nodes.item(i).parent.parent Is Nothing Then
            ElseIf LeftOf(tvList.Nodes.item(i).parent.parent.Text, ":") = serverID Then
                If tvList.Nodes.item(i).Text = strChannel Then
                    treeview_GetChannelIndex = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    treeview_GetChannelIndex = -1
End Function
Public Function treeview_GetStatusIndex(tvList As TreeView, serverID As Integer) As Integer
    Dim i As Integer

    For i = 1 To tvList.Nodes.Count
        If tvList.Nodes.item(i).parent Is Nothing Then
            If CInt(LeftOf(tvList.Nodes.item(i).Text, ":")) = serverID Then
                treeview_GetStatusIndex = i
                Exit Function
            End If
        End If
    Next i
    
    treeview_GetStatusIndex = -1
End Function

Public Sub treeview_RemoveQuery(tvList As TreeView, strQuery As String, serverID As Integer)
    Dim nIndex As Integer
    nIndex = treeview_GetQueryIndex(tvList, strQuery, serverID)
    'MsgBox nIndex & "~" & strQuery & "~"
        
    If nIndex <> -1 Then
        If tvList.Nodes.item(nIndex).FirstSibling = tvList.Nodes.item(nIndex).LastSibling Then
            tvList.Nodes.Remove nIndex - 1
        Else
            tvList.Nodes.Remove nIndex
        End If
    End If
End Sub

Public Sub treeview_RemoveChannel(tvList As TreeView, strChannel As String, serverID As Integer)
    Dim nIndex As Integer
    nIndex = treeview_GetChannelIndex(tvList, strChannel, serverID)
        
    If nIndex <> -1 Then
        If tvList.Nodes.item(nIndex).FirstSibling = tvList.Nodes.item(nIndex).LastSibling Then
            tvList.Nodes.Remove nIndex - 1
        Else
            tvList.Nodes.Remove nIndex
        End If
    End If
    
End Sub
Public Sub treeview_RemoveServer(tvList As TreeView, serverID As Integer)
    Dim i As Integer
    For i = 1 To tvList.Nodes.Count
        If tvList.Nodes.item(i).parent Is Nothing Then
            If LeftOf(tvList.Nodes.item(i).Text, ":") = serverID Then
                tvList.Nodes.Remove i
                Exit Sub
            End If
        End If
    Next i
End Sub


Public Sub treeview_SetActive(tvList As TreeView, strWinName As String, serverID As Integer)
    Dim i As Integer, curServer As Integer
    For i = 1 To tvList.Nodes.Count
        If tvList.Nodes.item(i).parent Is Nothing Then
            curServer = curServer + 1
            If curServer = serverID And strWinName = "Status" Then
                Set tvList.selectedItem = tvList.Nodes.item(i)
                Exit Sub
            End If
        Else
            If tvList.Nodes.item(i).Text = strWinName And serverID = curServer Then
                Set tvList.selectedItem = tvList.Nodes.item(i)
                Exit Sub
            End If
        End If
    Next i
    
End Sub


