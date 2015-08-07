<#
Grabs the network interfaces and resets the IP address to the IP address that Amazon has set it to. When you launch a new instance, the DNS settings break due
to the DNS change. This essentially changes DNS settings and associates the necessary IP addresses to the IIS websites.
#>

<#Function that is used to reverse the IP address#>

Function reverseString($ipaddress){
    $text = $ipaddress -split ""
    [Array]::Reverse($text)
    -join $text
#   return $text
}

Import-Module WebAdministration

#sets the necessary variables
$IP = ""
$pos = ""
$defaultGateWay = ""
#grab the current NetIPConfiguration values and store them.
$ipAddresses = Get-NetIPConfiguration

#for each ipaddress that is configured, it runs through a loop to extract them and we start doing comparisons.

ForEach ($ipaddress in $ipAddresses.IPv4Address){
       
    #if the IP address is not null then it checks if the left 6 characters match 10.41. This means that you are using the correct interface.
    #if the IP varable is something else, we skip this alltogether and the script does nothing. This will error out if there are no "10.41" IP's.
	if ($ipaddress.ipaddress.substring(0, 6) -eq "10.41.") {
		if ($IP -eq "") {
		    $IP = $ipaddress.ipaddress
		  
         #if none of the above is true (IP is already filled out) then it sets a new varable called pos (position) to the result position of reverse string IP address and the ".".
         #this should result in either 2,3 or 4 (based on last octet). We then compare if strIPAddress of the current loop is less than the IP address set in the variable before. 
         #We do this in case a server has multiple IP addresses. The lowest IP address is used for the website and base IP of the server.
         #If true then we reset the IP and adapter to the updated lowest IP.
          
        }Else {	
            $pos = reverseString($ipaddress.ipaddress)
            $pos = $pos.indexOf(".")
			if ($ipaddress.ipaddress.substring($ipaddress.ipaddress.length - $pos) -lt ($IP.substring($IP.length - $pos))){
				$IP2 = $IP
                $IP = $ipaddress.ipaddress
            }
            else {
                $IP2 = $ipaddress.ipaddress
            }
        }
	}
}
#Here we grab the default gateway that will be set in a command later.
$defaultGateWay = $ipAddresses.IPv4DefaultGateway | select-object -Property NextHop -ExpandProperty NextHop

#gets the position of the last octet (10.41.4.<- finds that value) by reversing the IP address.Sets it to varable pos.

$pos = reverseString($IP).indexOf(".")

#gets the length of the IP address (10.41.4.10 for example would be 10) then subtracks pos (which could be 3 (for the ".") and adds 1 (to go beyond the "."). Assigns this to lsize variable. This is no longer used.

#$lsize = $IP.length - $pos + 1

<#Check to see if its in the DMZ/UI subnets. These subnets are usually 9 characters long so we do a check for 8 characters.
Grabs the first eight characters of the IP (if its 10.41.4.) then the first on the right (".") and checks if it equals ".".
if true, we set the DNS search order to 10.41.4."10".
if the ip is in the DMZ/UI network, this will prove false. So we grab one extra space and append "10".
I believe this if statement can be omitted, and just a statement to change the DNSServerSearchOrder takes place as the first 8/9 octets don't matter since we will
be changing them anyway. The lsize variable is set to the correct position, so the first if code should work for the else statement.#>

#Sets a variable array to store the IP address and gateways. Sets the commands to the variable as they will be called shortly.
#write-host ($IP.substring(0, $IP.length - $pos -1) + "10")
&{$adapter = Get-NetAdapter -Name Ethernet; Set-DnsClientServerAddress -InterfaceAlias $adapter.Name -ServerAddresses (($IP.substring(0, $IP.length - $pos -1) +"10"))}

#Use the following if you want to re-configure the IP address
#&{$adapter = Get-NetAdapter -Name Ethernet;New-NetIPAddress -InterfaceAlias $adapter.Name -IPAddress ($IP) -PrefixLength 24 -DefaultGateway $defaultGateWay; Set-DnsClientServerAddress -InterfaceAlias $adapter.Name -ServerAddresses (($IP.substring(0, $IP.length - $pos) +"10"))}

#set DNS address - alternative way
#$dnsInterface = Get-DnsClientServerAddress
#ForEach ($dnsServer in $dnsInterface.ServerAddresses){
#    if ($dnsServer.contains("10.41.")){
#        Set-DnsClientServerAddress -InterfaceIndex $dnsServer.InterfaceIndex -ServerAddresses (($IP.substring($IP.length - $pos, $IP.length - $pos +1) +"10"))
#    }
#}

<#
The following code sets the iis servers. The first one sets a perspectica site and assigns it to run on port 443.
The third sets up a webservice website and the forth associates a binding. The fifth and sixth loops assocate the necessary certificates to the websites.
Finally a register dns command takes place.
#>

Remove-Website -name ListWebsiteNameHere

New-Website -Name ListWebsiteNameHere -ApplicationPool ListAppPoolNameHere -IPAddress $IP -Port 443 -Ssl -Force

Remove-Website -name ListWebsiteNameHere

New-Website -Name ListWebsiteNameHere -ApplicationPool ListAppPoolNameHere -IPAddress $IP2 -Port 443 -Ssl -Force

Get-ChildItem Cert:\LocalMachine\My | select -First 1 | New-Item IIS:\SslBindings\"$IP"!443

Get-ChildItem Cert:\LocalMachine\My | select -First 1 | New-Item IIS:\SslBindings\"$IP2"!443

#Register the DNS entry and flush the DNS
ipconfig /registerdns
ipconfig /flushdns
