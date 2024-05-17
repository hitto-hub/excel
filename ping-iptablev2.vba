Sub PingAndSetStatus()
    Dim StartRow As Integer
    Dim EndRow As Integer
    Dim HostIpColumn As Integer
    Dim StatusColumn As Integer
    Dim UserColumn As Integer
    Dim hostIp As String
    Dim i As Integer
    Dim totalRows As Integer
    Dim progress As Single
    Dim progressUpdateFrequency As Integer
    
    StartRow = 3 ' Starting from the 3rd row
    HostIpColumn = 1 ' IP addresses are in column A
    StatusColumn = 2 ' Status will be set in column B
    UserColumn = 3 ' User information is in column C
    
    EndRow = Cells(Rows.Count, HostIpColumn).End(xlUp).row - 1 ' Find the last row with data in column A, exclude the last row
    totalRows = EndRow - StartRow + 1 ' Total number of rows to process
    progressUpdateFrequency = 10 ' Update progress every 10 rows
    
    For i = StartRow To EndRow
        hostIp = Cells(i, HostIpColumn).value
        
        If Not hostIp = "" Then
            Dim objShell As Object
            Dim returnCode As Long
            
            Set objShell = CreateObject("wscript.shell")
            returnCode = objShell.Run("ping -n 1 -w 50 " & hostIp, 0, True)
            
            If returnCode = 0 Then ' If ping is successful
                Cells(i, StatusColumn).value = 3
            Else ' If ping fails
                If Cells(i, UserColumn).value <> "" Then
                    Cells(i, StatusColumn).value = 2
                Else
                    Cells(i, StatusColumn).value = 1
                End If
            End If
        End If
        
        ' Update progress in status bar every 10 rows
        If (i - StartRow + 1) Mod progressUpdateFrequency = 0 Or i = EndRow Then
            progress = (i - StartRow + 1) / totalRows * 100
            Application.StatusBar = "Progress: " & Format(progress, "0.00") & "% (" & (i - StartRow + 1) & " of " & totalRows & ")"
        End If
    Next i
    
    ' Reset status bar
    Application.StatusBar = False
End Sub
