Sub Ping()

    Dim StartRow, HostIpColumn, ResultColumn As Integer
    Dim TotalRows As Integer
    
    StartRow = 2
    HostIpColumn = 2
    ResultColumn = 3
    
    TotalRows = Cells(Rows.Count, HostIpColumn).End(xlUp).Row
    Range(Cells(StartRow, ResultColumn), Cells(TotalRows, ResultColumn).End(xlDown)).ClearContents
    
    Dim hostIp As String
    
    For i = StartRow To TotalRows
    
        hostIp = ActiveSheet.Cells(i, HostIpColumn).Value
        
        If Not hostIp = "" And Not hostIp = "host not reachable" Then
        
            Dim objShell, returnCode
            
            Set objShell = CreateObject("wscript.shell")
            returnCode = objShell.Run("ping -n 1 -w 1000 " & hostIp, 0, True)
            
            If returnCode = 0 Then
                ActiveSheet.Cells(i, ResultColumn).Value = "Online"
                ActiveSheet.Cells(i, ResultColumn).Font.Color = vbGreen
                ActiveSheet.Cells(i, ResultColumn + 1).Value = Now
            Else
                ActiveSheet.Cells(i, ResultColumn).Value = "Offline"
                ActiveSheet.Cells(i, ResultColumn).Font.Color = vbRed
            End If
        End If
        
    Next

End Sub