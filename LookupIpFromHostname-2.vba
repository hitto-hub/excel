
Sub LookupIpFromHostname()
    
    Dim StartRow, HostNameColumn, ResultColumn As Integer
    Dim TotalRows As Integer
    Dim hostName As String
    
    StartRow = 2
    HostNameColumn = 1
    ResultColumn = 2
    
    TotalRows = Cells(Rows.Count, HostNameColumn).End(xlUp).Row
    
    Range(Cells(StartRow, ResultColumn), Cells(TotalRows, ResultColumn).End(xlDown)).ClearContents

    
    For i = StartRow To TotalRows
        hostName = ActiveSheet.Cells(i, HostNameColumn).Value
        
        If hostName <> "" Then

            Dim objShell As Object
            Dim returnCode As String
                    
            Set objShell = CreateObject("wscript.shell")
            
            returnCode = objShell.exec("ping -n 1 -w 500 -4 " & hostName).stdout.ReadAll
         
            Dim ip As String
            ip = FindIP(returnCode)
            
            If ip <> "" Then
                ActiveSheet.Cells(i, ResultColumn).Value = Mid(ip, 2, Len(ip) - 2)
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
    Dim RegEx As Object
    Dim valid As Boolean
    Dim Matches As Object
    Dim i As Integer
    
    Set RegEx = CreateObject("VBScript.RegExp")
   
    RegEx.Pattern = "\[\b(?:\d{1,3}\.){3}\d{1,3}\b\]"
    
    valid = RegEx.test(strTest)
    
    If valid Then
        Set Matches = RegEx.Execute(strTest)
        FindIP = Matches(0)
    Else
        FindIP = ""
    End If
End Function
