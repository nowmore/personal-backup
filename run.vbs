
Dim WMI, Objs, Process
Set WMI=GetObject("WinMgmts:")
Set Objs=WMI.InstancesOf("Win32_Process")
For Each Obj In Objs
Process = Obj.Description
if Process = "aria2c.exe" then
    WScript.Quit
end if
Next

CreateObject("WScript.Shell").Run "C:\Users\when-\aria2\aria2c.exe --conf-path=C:\Users\when-\aria2\aria2.conf",0
