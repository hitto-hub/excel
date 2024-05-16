# excel

## Ping 〇

**Done**

## LookupIpFromMac

### 案1 ×

ExcelシートのMACアドレス列からデータを読み込みます。
Nmapを使用してネットワーク全体をスキャンし、結果を取得します。
スキャン結果から、指定されたMACアドレスに対応するIPアドレスを抽出します。
IPアドレスをExcelシートの対応するセルに書き込みます。

ネットワーク全体をスキャンするのは時間がかかるためだめ。
うまくやれば耐えるかもしれない。

### 案2 △

Excelシートの指定された列（MACColumn）からMACアドレスを取得します。
arpコマンドを使用してMACアドレスからIPアドレスを取得します。
取得したIPアドレスをExcelシートに表示します。

arpコマンドで取得できるのは、直近で通信した相手のみ。

## LookupIpFromHostname

### 案1 ×

**Done**

Excelシートのホスト名列からデータを読み込みます。
nslookupコマンドを使用して、ホスト名からIPアドレスを取得します。
取得したIPアドレスをExcelシートの対応するセルに書き込みます。

### 案2 〇

**Done**

Excelシートの指定された列（HostNameColumn）からホスト名を取得します。
pingコマンドを使用してホスト名からIPアドレスを取得します。
取得したIPアドレスをExcelシートに表示します。

## LookupMacFromIp

### 案1 ×

** Done **

ExcelシートのIPアドレス列からデータを読み込みます。
Nmapを使用して指定のIPアドレスをスキャンし、結果を取得します。
スキャン結果から、指定されたIPアドレスに対応するMACアドレスを抽出します。
MACアドレスをExcelシートの対応するセルに書き込みます。

ネットワーク全体をスキャンするのは時間がかかるためだめ。

### 案2 〇

** Done **

Excelシートの指定された列（IPColumn）からIPアドレスを取得します。
IPアドレスに対してPingコマンドを使用して、MACアドレスを取得します。
arpコマンドを使用してIPアドレスからMACアドレスを取得します。
取得したMACアドレスをExcelシートに表示します。

## LookupMacFromHostname

### 案1 ×

** Done **

Excelシートのホスト名列からデータを読み込みます。
nslookupコマンドを使用して、ホスト名からIPアドレスを取得します。
取得したIPアドレスを用いて、Nmapを使用して指定のIPアドレスをスキャンし、結果を取得します。
スキャン結果から、指定されたIPアドレスに対応するMACアドレスを抽出します。
MACアドレスをExcelシートの対応するセルに書き込みます。

### 案2 〇

** Done **

Excelシートの指定された列（HostNameColumn）からホスト名を取得します。
pingコマンドを使用してホスト名からIPアドレスを取得します。
IPアドレスに対してarpコマンドを使用してMACアドレスを取得します。
取得したMACアドレスをExcelシートに表示します。

実質、LookupMacFromIpと同じ

## 注意

arpコマンドを使用しているものは、同じネットワーク内でないといけないかも

excel
A    B	C	D	E
Hostname	IP-Address	Status	Last reachable at	MAC Address
