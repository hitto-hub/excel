' 作戦
' 並列処理: 同時に複数のpingを実行することで、処理時間を短縮できる。
' VBAには並列処理機能ないらしい、
' Windows API。VBAから他のスクリプト言語を呼び出せる機能を利用して並列処理を実現できそう。

' ubuntu bashでの実行例
echo 192.168.0.{1..254} | fmt -1 | xargs -P 100 -I@ bash -c "ping -c 1 @ | awk '{print strftime(\"%F %T \")\"(@)\"\"\t\" \$0}{fflush() }'" | grep -e'ttl=' -e'timeout' -e'Unreachable'