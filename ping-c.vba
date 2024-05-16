Sub Ping()
    Dim StartRow, HostIpColumn, StatusColumn, LastReachableColumn As Integer
    Dim TotalRows As Integer
    Dim hostIp As String
    
    StartRow = 2
    HostIpColumn = 2
    StatusColumn = 3
    LastReachableColumn = 4
    
    TotalRows = Cells(Rows.Count, HostIpColumn).End(xlUp).Row
    Range(Cells(StartRow, StatusColumn), Cells(TotalRows, StatusColumn)).ClearContents
    Range(Cells(StartRow, LastReachableColumn), Cells(TotalRows, LastReachableColumn)).ClearContents
    
    For i = StartRow To TotalRows
        hostIp = ActiveSheet.Cells(i, HostIpColumn).Value
        
        If Not hostIp = "" And Not hostIp = "host not reachable" Then
            Dim objShell, returnCode
            
            Set objShell = CreateObject("wscript.shell")
            returnCode = objShell.Run("ping -n 1 -w 1000 " & hostIp, 0, True)
            
            If returnCode = 0 Then
                ActiveSheet.Cells(i, StatusColumn).Value = "Online"
                ActiveSheet.Cells(i, StatusColumn).Font.Color = vbGreen
                ActiveSheet.Cells(i, LastReachableColumn).Value = Now
            Else
                ActiveSheet.Cells(i, StatusColumn).Value = "Offline"
                ActiveSheet.Cells(i, StatusColumn).Font.Color = vbRed
            End If
        End If
    Next
End Sub
