Dim strHost

strHost = Trim(Inputbox("Please enter the host name to ping", "Enter the hostname/IP address"))

PingForever strHost

Sub PingForever(strHost)
    Dim Shell, strCommand, ReturnCode

    Set Shell = CreateObject("wscript.shell")
    strCommand = "ping -n 1 " & strHost
    Do While(True)
        ReturnCode = Shell.Run(strCommand, 0, True)     
        If ReturnCode = 0 Then
            msgbox "Server is Up", 48, "Get your data, bitch!"
			exit do
        End If
        Wscript.Sleep 180000
    Loop
End Sub
