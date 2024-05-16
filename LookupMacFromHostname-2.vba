Sub LookupMacFromHostname()
    Dim StartRow, IpAddressColumn, ResultColumn As Integer
    Dim TotalRows As Integer
    Dim ipAddress As String
    
    StartRow = 2
    IpAddressColumn = 2
    ResultColumn = 5
    
    TotalRows = Cells(Rows.Count, IpAddressColumn).End(xlUp).Row

    Range(Cells(StartRow, ResultColumn), Cells(TotalRows, ResultColumn).End(xlDown)).ClearContents

    For i = StartRow To TotalRows
        ipAddress = ActiveSheet.Cells(i, IpAddressColumn).Value
        
        If ipAddress <> "" Then

            Dim objShell As Object
            Dim returnCode As String
                    
            Set objShell = CreateObject("wscript.shell")
            
            returnCode = objShell.exec("arp -a -v " & ipAddress).stdout.ReadAll
         
            
            Dim mac As String
            mac = FindMAC(returnCode, ipAddress)
            
            If mac <> "" Then
                ActiveSheet.Cells(i, ResultColumn).Value = mac
                ActiveSheet.Cells(i, ResultColumn).Font.Color = vbBlack
             Else
                ActiveSheet.Cells(i, ResultColumn).Value = "err: arp entry missing"
                ActiveSheet.Cells(i, ResultColumn).Font.Color = vbRed
            End If
        Else
            ActiveSheet.Cells(i, ResultColumn).Value = ""
        End If
        Application.Wait Now + TimeValue("00:00:01")
    Next

End Sub

Function FindMAC(strTest As String, ip As String) As String
    Dim substring As String
    Dim i, start, ipLength, totalLength As Integer
    
    If InStr(strTest, ip) <> 0 Then
       start = InStr(strTest, ip)
       ipLength = Len(ip)
       totalLength = Len(strTest)
       
       substring = Mid(strTest, start + ipLength, totalLength - start - ipLength)
       
       For Each test In Split(Trim(substring), " ")
            If test <> "" And Len(test) = 17 Then
                FindMAC = UCase(test)
                Exit For
            End If
       Next
       
    Else
    
    End If
   
End Function

