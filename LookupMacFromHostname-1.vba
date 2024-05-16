Sub LookupMacFromHostname()
    Dim StartRow, HostNameColumn, IpAddressColumn, MacAddressColumn As Integer
    Dim TotalRows As Integer
    Dim hostName As String
    Dim ipAddress As String
    Dim macAddress As String
    
    StartRow = 2
    HostNameColumn = 1
    IpAddressColumn = 2
    MacAddressColumn = 5
    
    TotalRows = Cells(Rows.Count, HostNameColumn).End(xlUp).Row
    
    Range(Cells(StartRow, IpAddressColumn), Cells(TotalRows, IpAddressColumn)).ClearContents
    Range(Cells(StartRow, MacAddressColumn), Cells(TotalRows, MacAddressColumn)).ClearContents

    For i = StartRow To TotalRows
        hostName = ActiveSheet.Cells(i, HostNameColumn).Value
        
        If hostName <> "" Then
            Dim objShell As Object
            Dim returnCode As String
            
            Set objShell = CreateObject("wscript.shell")
            returnCode = objShell.exec("nslookup " & hostName).stdout.ReadAll
         
            Dim ip As String
            ip = FindIP(returnCode)
            
            If ip <> "" Then
                ipAddress = ip
                ActiveSheet.Cells(i, IpAddressColumn).Value = ipAddress
                ActiveSheet.Cells(i, IpAddressColumn).Font.Color = vbBlack
                
                returnCode = objShell.exec("nmap -sP " & ipAddress).stdout.ReadAll
                
                macAddress = FindMAC(returnCode, ipAddress)
                
                If macAddress <> "" Then
                    ActiveSheet.Cells(i, MacAddressColumn).Value = macAddress
                    ActiveSheet.Cells(i, MacAddressColumn).Font.Color = vbBlack
                Else
                    ActiveSheet.Cells(i, MacAddressColumn).Value = "MAC not found"
                    ActiveSheet.Cells(i, MacAddressColumn).Font.Color = vbRed
                End If
                
             Else
                ActiveSheet.Cells(i, IpAddressColumn).Value = "host not reachable"
                ActiveSheet.Cells(i, IpAddressColumn).Font.Color = vbRed
            End If
        Else
            ActiveSheet.Cells(i, IpAddressColumn).Value = ""
        End If
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

' FindMAC functionを入れましょう