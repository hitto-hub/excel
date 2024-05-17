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
    
    StartRow = 3 ' Starting from the 3rd row
    HostIpColumn = 1 ' IP addresses are in column A
    StatusColumn = 2 ' Status will be set in column B
    UserColumn = 3 ' User information is in column C
    
    EndRow = Cells(Rows.Count, HostIpColumn).End(xlUp).Row - 1 ' Find the last row with data in column A, exclude the last row
    totalRows = EndRow - StartRow + 1 ' Total number of rows to process
    
    For i = StartRow To EndRow
        hostIp = Cells(i, HostIpColumn).Value
        
        If Not hostIp = "" Then
            Dim objShell As Object
            Dim returnCode As Long
            
            Set objShell = CreateObject("wscript.shell")
            returnCode = objShell.Run("ping -n 1 -w 50 " & hostIp, 0, True)
            
            If returnCode = 0 Then ' If ping is successful
                Cells(i, StatusColumn).Value = 3
            Else ' If ping fails
                If Cells(i, UserColumn).Value <> "" Then
                    Cells(i, StatusColumn).Value = 2
                Else
                    Cells(i, StatusColumn).Value = 1
                End If
            End If
        End If
        
        ' Update progress in status bar
        progress = (i - StartRow + 1) / totalRows * 100
        Application.StatusBar = "Progress: " & Format(progress, "0.00") & "% (" & (i - StartRow + 1) & " of " & totalRows & ")"
    Next i
    
    ' Reset status bar
    Application.StatusBar = False
End Sub

