### Office 64 Bit Migration Check Tool 
Winnie 11/18

在Migrate 32Bit到64Bit時，大多數VBA 程式碼不需要修改，除非您有使用到Declare Statement去呼叫Windows的API(使用32Bit的Data Type如Long，用在Pointers及Handles等)。若是此情形，將 PtrSafe 新增到 Declare，並以 LongPtr 取代 long，就可讓 Declare 陳述式同時與 32 位元和 64 位元相容。

透過偵測工具檢測是否Macro有使用到Declare 陳述式

1. 從此網址安裝Office Code Compatibility Inspector Add-in：https://www.microsoft.com/en-us/download/confirmation.aspx?id=15001

2. 預設解壓縮路徑

![](https://i.imgur.com/MsQiPKj.png)

3. 安裝完前往，檔案>選項>增益集，確認下列增益集是否成功Enable

![](https://i.imgur.com/WDnRf6j.png)

4. 若確認增益集已設定好，則到　檔案＞選項＞快速存取工具列，選擇[開發人員]索引標籤

![](https://i.imgur.com/z2p22tI.png)

5. 選擇VBA Inspector

![](https://i.imgur.com/7xDLLo1.png)

6. 你會在Excel上方看到Inspect VBA Code(此功能僅限使用在含有Macro的Excel文件上)

![](https://i.imgur.com/CnDOMCT.png)

7. 檢測Declare Statements及VBA Project Reference (64 Bit 相容性)

![](https://i.imgur.com/FdaCb1C.png)

8. 需要關注的是Declare Statements的數量，再針對檢查出的Declare Statements進行修改

![](https://i.imgur.com/LATrOBW.png)


### 修改範例：（Reference from: [網址](https://www.saka-en.com/office/vba-32bit-64bit-declare-statement-branch/))


#If VBA7 And Win64 Then
  ' 64Bit 版
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
  ' 32Bit 版
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
#End If

