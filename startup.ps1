<#
Grabs the network interfaces and resets the IP address to the IP address that Amazon has set it to. When you launch a new instance, the DNS settings break due
to the DNS change. This essentially changes DNS settings and associates the necessary IP addresses to the IIS websites.
#>

Import-Module WebAdministration

#sets the necessary variables
$IP = ""
$strIPAddress = ""
$pos = ""
$defaultGateWay = ""
$ipAddresses = Get-NetIPConfiguration

#for each adapter that objWMIService.ExecQuery pulls, it runs through a loop to extract the in use IP Addresses.

ForEach ($ipaddress in $ipAddresses.IPv4Address){
       
    #if the IP address is not null then it checks if the left 6 characters match 10.41. This means that you are using the correct interface.
    #if the IP varable is blank (this is basically a command that takes place first then the else takes place, then it sets the IP as the active IP from  the current strIPAddress that is being looped.
    #it then sets the objadp variable as the adapter that is currently being used in the loop.
    
	if ($ipaddress.substring(0, 6) -eq "10.41.") {
		if ($IP -eq "") {
		    $IP = $ipaddress
		  
         #if none of the above is true (IP is already filled out) then it sets a new varable called pos (position) to the result position of reverse string IP address and the ".".
         #this should result in either 2,3 or 4 (based on last octet). We then compare if strIPAddress of the current loop is less than the IP address set in the variable before. 
         #We do this in case a server has multiple IP addresses. The lowest IP address is used for the website and base IP of the server.
         #If true then we reset the IP and adapter to the updated lowest IP.
          
        }Else {	
			$pos = reverseString($ipaddress).indexOf(".")
			if (($ipaddress.substring($ipaddress.length - $pos, $ipaddress.length - $pos + 1)) -lt ($IP.substring($IP.length - $pos, $IP.length - $pos +1))){
				$IP = $ipaddress
            }
        }
	}
}

$defaultGateWay = $ipaddress.IPv4DefaultGateway

#gets the position of the last octet (10.41.4.<- finds that value) by reversing the IP address.Sets it to varable pos.

$pos = reverse($IP).indexOf(".")

#gets the length of the IP address (10.41.4.10 for example would be 10) then subtracks pos (which could be 3 (for the ".") and adds 1 (to go beyond the "."). Assigns this to lsize variable.

$lsize = $IP.length - $pos + 1

#sets a secondary ipaddress variable. Takes strIPAddress and increments it by 1. 

$IP2 = $IP.Substring(0,lsize) + $IP.Substring($lsize, IP.length) + 1

<#Check to see if its in the DMZ/UI subnets. These subnets are usually 9 characters long so we do a check for 8 characters.
Grabs the first eight characters of the IP (if its 10.41.4.) then the first on the right (".") and checks if it equals ".".
if true, we set the DNS search order to 10.41.4."10".
if the ip is in the DMZ/UI network, this will prove false. So we grab one extra space and append "10".
I believe this if statement can be omitted, and just a statement to change the DNSServerSearchOrder takes place as the first 8/9 octets don't matter since we will
be changing them anyway. The lsize variable is set to the correct position, so the first if code should work for the else statement.#>

#Sets a variable array to store the IP address and gateways. Sets the commands to the variable as they will be called shortly.

&{$adapter = Get-NetAdapter -Name Ethernet;New-NetIPAddress -InterfaceAlias $adapter.Name -AddressFamily IPv4 -IPAddress ($IP, $IP2) -PrefixLength 24 -DefaultGateway $defaultGateWay; Set-DnsClientServerAddress -InterfaceAlias $adapter.Name -ServerAddresses (($IP.substring($IP.length - $pos, $IP.length - $pos +1) +"10"))}

#set DNS address
#$dnsInterface = Get-DnsClientServerAddress
#ForEach ($dnsServer in $dnsInterface.ServerAddresses){
#    if ($dnsServer.contains("10.41.")){
#        Set-DnsClientServerAddress -InterfaceIndex $dnsServer.InterfaceIndex -ServerAddresses (($IP.substring($IP.length - $pos, $IP.length - $pos +1) +"10"))
#    }
#}

<#
The following code sets the iis servers. The first one sets a perspectica site. The second sets a binding to the https site.
The third sets up a webservice website and the forth associates a binding. The fifth and sixth loops assocate the necessary certificates to the websites.
Finally a register dns command takes place.
#>

Remove-Website -name Perspectica

New-Website -Name Perspectica -ApplicationPool Perspectica -IPAddress $IP -Port 443 -Ssl

Remove-Website -name Perspectica

New-Website -Name Webservice -ApplicationPool Webservice -IPAddress $IP2 -Port 443 -Ssl

netsh http add sslcert ipport=" & $IP & ":443 certhash=‎8cdc9274f1d0a266b762fddac9f3b228694d841f appid='{4dc3e181-e14b-4a21-b022-59fc669b0914}'

netsh http add sslcert ipport=" & $IP2 & ":443 certhash=8cdc9274f1d0a266b762fddac9f3b228694d841f appid='{4dc3e181-e14b-4a21-b022-59fc669b0914}'

ipconfig /registerdns
ipconfig /flushdns

<#Function that is used to reverse the IP address#>

Function reverseString($ipaddress){
    $text = $ipaddress.ToCharArray()
    [Array];;Reverse($text)
    -join $text
    return $text
}
