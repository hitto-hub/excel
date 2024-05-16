' うまくいかない
' うまく切り取れていない、、、 多分

Sub LookupIpFromHostname()
    Dim StartRow, HostNameColumn, ResultColumn As Integer
    Dim TotalRows As Integer
    Dim hostName As String
    
    StartRow = 2
    HostNameColumn = 1
    ResultColumn = 2
    
    TotalRows = Cells(Rows.Count, HostNameColumn).End(xlUp).Row
    
    Range(Cells(StartRow, ResultColumn), Cells(TotalRows, ResultColumn)).ClearContents

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
                ActiveSheet.Cells(i, ResultColumn).Value = ip
                ActiveSheet.Cells(i, ResultColumn).Font.Color = vbBlack
            Else
                ActiveSheet.Cells(i, ResultColumn).Value = "host not reachable"
                ActiveSheet.Cells(i, ResultColumn).Font.Color = vbRed
            End If
        Else
            ActiveSheet.Cells(i, ResultColumn).Value = ""
        End If
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

Function FindIP(strTest As String) As String
    Dim ipPattern As String
    ipPattern = "(\d{1,3}\.){3}\d{1,3}"
    
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = ipPattern
    
    Dim Matches As Object
    Set Matches = re.Execute(strTest)
    
    If Matches.Count > 0 Then
        FindIP = Matches(0).Value
    Else
        FindIP = ""
    End If
End Function
