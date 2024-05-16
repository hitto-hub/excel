' かすプログラム

Sub LookupMacFromIp()
    Dim StartRow, IpAddressColumn, ResultColumn As Integer
    Dim TotalRows As Integer
    Dim ipAddress As String
    
    StartRow = 2
    IpAddressColumn = 2
    ResultColumn = 5
    
    TotalRows = Cells(Rows.Count, IpAddressColumn).End(xlUp).Row
    
    Range(Cells(StartRow, ResultColumn), Cells(TotalRows, ResultColumn)).ClearContents

    For i = StartRow To TotalRows
        ipAddress = ActiveSheet.Cells(i, IpAddressColumn).Value
        
        If ipAddress <> "" Then
            Dim objShell As Object
            Dim returnCode As String
            
            Set objShell = CreateObject("wscript.shell")
            returnCode = objShell.exec("nmap -sP " & ipAddress).stdout.ReadAll
            
            Dim mac As String
            mac = FindMAC(returnCode, ipAddress)
            
            If mac <> "" Then
                ActiveSheet.Cells(i, ResultColumn).Value = mac
                ActiveSheet.Cells(i, ResultColumn).Font.Color = vbBlack
            Else
                ActiveSheet.Cells(i, ResultColumn).Value = "MAC not found"
                ActiveSheet.Cells(i, ResultColumn).Font.Color = vbRed
            End If
        Else
            ActiveSheet.Cells(i, ResultColumn).Value = ""
        End If
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

Function FindMAC(strTest As String, ip As String) As String
    Dim macPattern As String
    macPattern = "(\w{2}-\w{2}-\w{2}-\w{2}-\w{2}-\w{2})"
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = ip & "\s*" & macPattern
    
    Dim Matches As Object
    Set Matches = re.Execute(strTest)
    
    If Matches.Count > 0 Then
        FindMAC = Matches(0).SubMatches(0)
    Else
        FindMAC = ""
    End If
End Function
