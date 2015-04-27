'Grabs the network interfaces and resets the IP address to the IP address that Amazon has set it to. When you launch a new instance, the DNS settings break due
'to the DNS change. This essentially changes DNS settings and associates the necessary IP addresses to the IIS websites.

'sets variables used for VBScript.

Dim objWMIService, colAdapters, objAdapter, objShell

'Calls the wscript shell.

Set objShell = CreateObject("Wscript.Shell")

'objWMIService is used to get impersonate level details on system objects.

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}//./root/cimv2")

'grabs the Win32_network adapters and finds out which ones are set to "True"

Set colAdapters = objWMIService.ExecQuery _
  ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

'sets the IP address variable
  
IP = ""

'for each adapter that objWMIService.ExecQuery pulls, it runs through a loop to extract the in use IP Addresses.

For Each objAdapter in colAdapters

   'another nested loop that goes through the ip addresses.
   For Each strIPAddress in objAdapter.IPAddress
   
    'if the IP address is not null then it checks if the left 6 characters match 10.41. This means that you are using the correct interface.
    'if the IP varable is blank (this is basically a command that takes place first then the else takes place, then it sets the IP as the active IP from  the current strIPAddress that is being looped.
    'it then sets the objadp variable as the adapter that is currently being used in the loop.
   
     If Not IsNull(strIPAddress) Then
	    If Left(strIPAddress, 6) = "10.41." Then
		  If IP = "" Then
			IP = strIPAddress
			Set objadp = objAdapter
		  
		  'if none of the above is true (IP is already filled out) then it sets a new varable called pos (position) to the result position of reverse string IP address and the ".".
          'this should result in either 2,3 or 4 (based on last octet). We then compare if strIPAddress of the current loop is less than the IP address set in the variable before. 
          'We do this in case a server has multiple IP addresses. The lowest IP address is used for the website and base IP of the server.
          'If true then we reset the IP and adapter to the updated lowest IP.
		  
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

'sets the IP variable to the strIPAddress to use below. Same with objAdapter and objadp. Below is a declaration of variables and setting objects.

strIPAddress = IP
Set objAdapter = objadp

'gets the position of the last octet (10.41.4.<- finds that value) by reversing the IP address.Sets it to varable pos.

pos = instr(strreverse(strIPAddress), ".")

'gets the length of the IP address (10.41.4.10 for example would be 10) then subtracks pos (which could be 3 (for the ".") and adds 1 (to go beyond the "."). Assigns this to lsize variable.

lsize = Len(strIPAddress) - pos + 1

'sets the default gateway ip address that is stored in objAdapter and sets it to variable gw.

gw = objAdapter.DefaultIPGateway

'sets a secondary ipaddress variable. Takes strIPAddress and increments it by 1.

strIPAddress2 = Left(strIPAddress,lsize) & CInt(Right(strIPAddress, pos - 1)) + 1

'Check to see if its in the DMZ/UI subnets. These subnets are usually 9 characters long so we do a check for 8 characters.
'Grabs the first eight characters of the IP (if its 10.41.4.) then the first on the right (".") and checks if it equals ".".
'if true, we set the DNS search order to 10.41.4."10".
'if the ip is in the DMZ/UI network, this will prove false. So we grab one extra space and append "10".
'I believe this if statement can be omitted, and just a statement to change the DNSServerSearchOrder takes place as the first 8/9 octets don't matter since we will
'be changing them anyway. The lsize variable is set to the correct position, so the first if code should work for the else statement.

If Right(Left(strIPAddress, 8), 1) = "." Then
		objAdapter.SetDNSServerSearchOrder(Array(Left(strIPAddress,lsize) & "10"))
Else	
	objAdapter.SetDNSServerSearchOrder(Array("10.41." & Right(Left(strIPAddress, 8), 1) & ".10"))
End If

'Sets a variable array to store the IP address and gateways. Sets the commands to the variable as they will be called shortly.

res = objAdapter.EnableStatic(Array(strIPAddress, strIPAddress2), Array("255.255.255.0","255.255.255.0"))
res = objAdapter.SetGateways(gw)

'The following code sets the iis servers. The first one sets a perspectica site. The second sets a binding to the https site.
'The third sets up a webservice website and the forth associates a binding. The fifth and sixth loops assocate the necessary certificates to the websites.
'Finally a register dns command takes place.

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

'function that is used for waiting. It waits about a total of 2 minutes if required.

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
