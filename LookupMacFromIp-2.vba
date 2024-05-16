Sub LookupMacFromIp()
    Dim StartRow, IpAddressColumn, ResultColumn As Integer
    Dim TotalRows As Integer
    Dim ipAddress As String
    
    ' 初期値の設定
    StartRow = 2
    IpAddressColumn = 2
    ResultColumn = 5
    
    ' 最後の行番号を取得
    TotalRows = Cells(Rows.Count, IpAddressColumn).End(xlUp).Row
    
    ' 以前の内容をクリア
    Range(Cells(StartRow, ResultColumn), Cells(TotalRows, ResultColumn)).ClearContents

    ' IPアドレスのループ
    For i = StartRow To TotalRows
        ipAddress = ActiveSheet.Cells(i, IpAddressColumn).Value
        
        ' IPアドレスが空でないか評価
        If ipAddress <> "" Then

            ' Ping コマンドを使用してIPアドレスにアクセス
            Dim objShell As Object
            Dim returnCode As Integer
            
            Set objShell = CreateObject("wscript.shell")
            returnCode = objShell.Run("ping -n 1 -w 1000 " & ipAddress, 0, True)
            
            If returnCode = 0 Then
                ' ARP コマンドを使用してMACアドレスを取得
                Dim arpResult As String
                arpResult = objShell.exec("arp -a " & ipAddress).stdout.ReadAll
                
                ' MACアドレスを結果から見つける
                Dim mac As String
                mac = FindMAC(arpResult, ipAddress)
                
                ' MACアドレスが存在する場合
                If mac <> "" Then
                    ActiveSheet.Cells(i, ResultColumn).Value = mac
                    ActiveSheet.Cells(i, ResultColumn).Font.Color = vbBlack
                Else
                    ' MACアドレスが見つからない場合
                    ActiveSheet.Cells(i, ResultColumn).Value = "MAC not found"
                    ActiveSheet.Cells(i, ResultColumn).Font.Color = vbRed
                End If
            Else
                ' Pingが失敗した場合
                ActiveSheet.Cells(i, ResultColumn).Value = "Host not reachable"
                ActiveSheet.Cells(i, ResultColumn).Font.Color = vbRed
            End If
        Else
            ActiveSheet.Cells(i, ResultColumn).Value = ""
        End If
        
        ' 次のコマンドの前に1秒待つ
        Application.Wait Now + TimeValue("00:00:01")
    Next
End Sub

Function FindMAC(strTest As String, ip As String) As String
    Dim substring As String
    Dim start, ipLength, totalLength As Integer
    
    ' 結果にIPアドレスが存在するか確認
    If InStr(strTest, ip) <> 0 Then
        ' IPアドレスの開始位置を取得
        start = InStr(strTest, ip)
        ipLength = Len(ip)
        totalLength = Len(strTest)
        
        ' サブストリングを取得
        substring = Mid(strTest, start + ipLength, totalLength - start - ipLength)
        
        ' サブストリングからMACアドレスを抽出
        Dim test As Variant
        For Each test In Split(Trim(substring), " ")
            ' MACアドレスのパターンに一致するか確認
            If test <> "" And Len(test) = 17 Then
                ' MACアドレスを返す
                FindMAC = UCase(test)
                Exit Function
            End If
        Next
    End If
    
    ' MACアドレスが見つからなかった場合
    FindMAC = ""
End Function
