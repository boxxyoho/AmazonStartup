Dim objWMIService, colAdapters, objAdapter, objShell
Set objShell = CreateObject("Wscript.Shell")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}//./root/cimv2")

Set colAdapters = objWMIService.ExecQuery _
  ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

IP = ""

For Each objAdapter in colAdapters
   For Each strIPAddress in objAdapter.IPAddress
     If Not IsNull(strIPAddress) Then
	If Left(strIPAddress, 6) = "10.41." Then
		If IP = "" Then
			IP = strIPAddress
			Set objadp = objAdapter
		Else	
			pos = instr(strreverse(strIPAddress), ".")
			If CInt(Right(strIPAddress, pos - 1)) < CInt(Right(IP, pos - 1)) Then
				IP = strIPAddress
				Set objadp = objAdapter
			End If
		End If
	End If
      End If
  Next
Next

strIPAddress = IP
Set objAdapter = objadp

pos = instr(strreverse(strIPAddress), ".")
lsize = Len(strIPAddress) - pos + 1
gw = objAdapter.DefaultIPGateway
strIPAddress2 = Left(strIPAddress,lsize) & CInt(Right(strIPAddress, pos - 1)) + 1
If Right(Left(strIPAddress, 8), 1) = "." Then
		objAdapter.SetDNSServerSearchOrder(Array(Left(strIPAddress,lsize) & "10"))
Else	
	objAdapter.SetDNSServerSearchOrder(Array("10.41." & Right(Left(strIPAddress, 8), 1) & ".10"))
End If

res = objAdapter.EnableStatic(Array(strIPAddress, strIPAddress2), Array("255.255.255.0","255.255.255.0"))
res = objAdapter.SetGateways(gw)


Do
	Set objExec = objShell.Exec("C:\Windows\System32\inetsrv\appcmd set site /site.name:Perspectica /-bindings.[protocol='https']")				

	If Waitexe(objExec, 20000) = 0 Then
		Exit Do
	End If
	connect = connect + 1
Loop Until connect < 3
Do

	Set objExec = objShell.Exec("C:\Windows\System32\inetsrv\appcmd set site /site.name:Perspectica /+bindings.[protocol='https',bindingInformation='" & strIPAddress & ":443:']")		
	If Waitexe(objExec, 20000) = 0 Then
		Exit Do
	End If
	connect = connect + 1
Loop Until connect < 3
Do
	Set objExec = objShell.Exec("C:\Windows\System32\inetsrv\appcmd set site /site.name:Webservice /-bindings.[protocol='https']")				

	If Waitexe(objExec, 20000) = 0 Then
		Exit Do
	End If
	connect = connect + 1
Loop Until connect < 3
Do

	Set objExec = objShell.Exec("C:\Windows\System32\inetsrv\appcmd set site /site.name:Webservice /+bindings.[protocol='https',bindingInformation='" & strIPAddress2 & ":443:']")		
	If Waitexe(objExec, 20000) = 0 Then
		Exit Do
	End If
	connect = connect + 1
Loop Until connect < 3
Do
	Set objExec = objShell.Exec("netsh http add sslcert ipport=" & strIPAddress & ":443 certhash=ec26a862fc8c7cf75fb9004e57fe3b13255272eb appid={4dc3e181-e14b-4a21-b022-59fc669b0914}")
	If Waitexe(objExec, 20000) = 0 Then
		Exit Do
	End If
	connect = connect + 1
Loop Until connect < 3
Do
	Set objExec = objShell.Exec("netsh http add sslcert ipport=" & strIPAddress2 & ":443 certhash=ec26a862fc8c7cf75fb9004e57fe3b13255272eb appid={4dc3e181-e14b-4a21-b022-59fc669b0914}")
	If Waitexe(objExec, 20000) = 0 Then
		Exit Do
	End If
	connect = connect + 1
Loop Until connect < 3	

Set objExec = objShell.Exec("ipconfig /registerdns")

Set colAdapters = Nothing
Set objWMIService = Nothing

Function Waitexe(objExec, MaxWaitTime)
    CurrentWaitTime = 0
    Do While objExec.Status = 0
	     WScript.Sleep 100
	     CurrentWaitTime = CurrentWaitTime + 100
	           If CurrentWaitTime >= MaxWaitTime Then
	                  objExec.StdIn.Close()
	                  Waitexe = objExec.ExitCode
	                  Exit Function
	           End if
	Loop
	Waitexe = objExec.ExitCode
End Function
