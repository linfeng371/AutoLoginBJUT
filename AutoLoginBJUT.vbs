' ***********************************************************
' *作者：凛风
' *E-Mail：Linfeng371@outlook.com
' *有任何问题请通过电子邮件联系我（请先确认账号密码的正确性）
' ***********************************************************

UserName = "替换为你的网关账号(注意保留引号)"
PassWord = "替换为你的网关密码(注意保留引号)"
' 例： 	UserName = "123456"
'      	PassWord = "789123"
dim v6ip
'v6ip = GetIpv6Address
' 如果使用ipv6，请删掉上一行的第一个符号“ ' ”。如果不知道ipv6或不使用ipv6请务必保留！


Function Login(UserName,PassWord,v6ip)
If IsNull(v6ip) Then
	req = "DDDDD=" + UserName + "&upass=" + PassWord +"&0MKKey=Login"
Else 
	req = "DDDDD=" + UserName + "&upass=" + PassWord +"&0MKKey=Login&v6ip=" + v6ip
End If 
set Http=createobject("MSXML2.XMLHTTP")
Http.Open "POST", "https://lgn.bjut.edu.cn", False
http.setRequestHeader "CONTENT-TYPE","application/x-www-form-urlencoded"
http.Send req

Do
	WScript.Sleep 20
Loop Until Http.readystate=4

ret = Http.responseText

If Instr(ret,"successfully") = 0 Then
	Login = false
Else
	Login = true
End If

End Function

Function GetIpv6Address()
Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
Set Items = wmiService.ExecQuery("Select * From Win32_NetworkAdapter")
s=""
For Each objItem in Items
    If LCase(Left(objItem.PNPDeviceID, 4))="pci\" Then
        Set Its = wmiService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration where Caption='"& objItem.Caption &"'")
        For Each It in Its
            If not IsNull(It.IPAddress) Then
                If not IsNull(It.IPAddress(2)) Then
                    address = It.IPAddress(2)
                End If
            End If
        Next
    End If
Next
GetIpv6Address = address

End Function


Login UserName,PassWord,v6ip
