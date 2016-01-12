PARAM(
[STRING]$BusinessName=$(Read-Host -prompt "What is the customers Full Business Name? e.g. ExamplePtyLtd"),
[STRING]$ShortName=$(Read-Host -prompt "What is the customers Short Business Name? e.g. ExamplePtyLtd to exapl"),
[STRING]$EmailDomainName=$(Read-Host -prompt "What is the customers Full Domain Name? e.g. example.com.au"),
[STRING]$DomainController,
[STRING]$ExchangeServer
)

# Import modules and snapins
if ( (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) -eq $null ){Add-PsSnapin Microsoft.Exchange.Management.PowerShell.E2010}
if ( (Get-module -Name ActiveDirectory -ErrorAction SilentlyContinue) -eq $null ){Import-module ActiveDirectory}

# Functions
function logMessage($type, $string){
    $colour=""
    if($type -eq "warning"){$colour = "yellow"}
    elseif($type -eq "error"){$colour = "red"}
    elseif($type -eq "info"){$colour = "white"}
    write-host $string -foregroundcolor $colour
    ($(get-date -format "yyyy-mm-dd_ss") + ":" + $type.ToUpper() + ": " + $string) | out-file -Filepath $LogFile -append
}
function logMessageSilently($type, $string){
    $colour=""
    if($type -eq "warning"){$colour = "yellow"}
    elseif($type -eq "error"){$colour = "red"}
    elseif($type -eq "info"){$colour = "white"}
    ($(get-date -format "yyyy-mm-dd_ss") + ":" + $type.ToUpper() + ": " + $string) | out-file -Filepath $logfile -append
}
function ConvertFrom-DN{
param([string]$DN=(Throw '$DN is required!'))
    foreach ( $item in ($DN.replace('\,','~').split(",")))
    {
        switch -regex ($item.TrimStart().Substring(0,3))
        {
            "CN=" {$CN = '/' + $item.replace("CN=","");continue}
            "OU=" {$ou += ,$item.replace("OU=","");$ou += '/';continue}
            "DC=" {$DC += $item.replace("DC=","");$DC += '.';continue}
        }
    }
    $canoincal = $dc.Substring(0,$dc.length - 1)
    for ($i = $ou.count;$i -ge 0;$i -- ){$canoincal += $ou[$i]}
    $canoincal += $cn.ToString().replace('~',',')
    return $canoincal
}
function validateCustomerName($newCustName){
   	if((Get-ADOrganizationalUnit -Filter 'Name -like $newCustName' -SearchBase $hostedExchange_baseDN) -ne $null){
   		logMessage error ("The customer " + $newCustName + " already exists in this environment! Please try again.")
   		echo $false
   	}else{
   		if($newCustName -match '[^a-z]'){
   			logMessage error ("The customer " + $newCustName + " should contain letters ONLY! e.g A to Z. Please try again.")
   			echo $false
   		}else{
   			echo $true
   	    }
    }
}
function validateCustomerShortname($newCustShortName){
	if((get-mailbox | Where-Object{$_.alias -like "$newCustShortName*"}) -ne $null){
		logMessage error ("The short name " + $newCustShortName + " is already used in this environment! Please try again.")
        echo $false
	}else{
		if($newCustShortName.length -gt 6){
			logMessage error ("The shortname is longer than 6 characters, that doesn't seem like a short name. Please try again.")
            echo $false
		}else{
			if($newCustShortName -match '[^a-z]'){
				logMessage error ("The shortname " + $newCustShortName + " should contain letters ONLY! e.g A to Z. Please try again.")
                echo $false
			}else{
				echo $True
			}
		}
	}
}
function validateCustomerDomain($newCustDomain){
    $failADCount="0"
    $failTLDCount="0"
    $tlds=".com",".org",".net",".edu",".asn",".co.uk",".co.nz",".info",".biz",".com.au",".net.au",".org.au",".edu.au",".asn.au","id.au"
    $tlds | foreach-object{
        if($newCustDomain -like "*$_"){
            $failTLDCount="1"
            Get-AcceptedDomain | ForEach-Object{
                if($newCustDomain -eq $_.DomainName){logMessage error ("Domain " + $newCustDomain + " is already an accepted domain"); $failADCount="1"}
            }
            if($failADCount -eq "0"){echo $true}elseif($failADCount -eq "1"){echo $false}else{logMessage error "Something went wrong! The failADCount variable in the validateCustomerDomain function should be 0 or 1"}
        }
    }
    if($failTLDCount -eq "0"){logMessage error ("The domain " + $newCustDomain + " doesn't match any of the valid TLD's in this script."); echo $false}
}
function Select-MyDomainController{
# Shamelessly stolen from http://blog.vertigion.com/post/16926803996/powershell-pick-a-domain-controller
    param(
        [Parameter(
            HelpMessage = "If set, will return the result intead of setting globally."
        )]
        [switch] $Return
    )
     
    $DCList = $null
    while ($DCList -eq $null) {
        $DCList = Get-ADDomainController -Filter * | Select HostName
    }
    if ($Return.isPresent) {
        $DC = ''
        while ($DC -eq $myDC) {
            $DC = $DCList[$(Get-Random -Minimum 0 -Maximum ($DCList | Measure-Object).Count)].HostName
        }
        return $DC
    } else {
        $global:myDC = $DCList[$(Get-Random -Minimum 0 -Maximum ($DCList | Measure-Object).Count)].HostName
        echo $myDC
    }
}
function Select-MyExchangeHost{
# Shamelessly stolen (and modified) from http://blog.vertigion.com/post/16926803996/powershell-pick-a-domain-controller
    param(
        [Parameter(
            HelpMessage = "If set, will return the result intead of setting globally."
        )]
        [switch] $Return
    )
     
    $EXList = $null
    while ($EXList -eq $null) {
        $EXList = @(Get-ExchangeServer | Where-Object {$_.serverrole -like "*Mailbox*"})
    }
    if ($Return.isPresent) {
        $EX = ''
        while ($EX -eq $myEX) {
            $EX = $EXList[$(Get-Random -Minimum 0 -Maximum ($EXList | Measure-Object).Count)].fqdn
        }
        return $EX
    } else {
        $global:myEX = $EXList[$(Get-Random -Minimum 0 -Maximum ($EXList | Measure-Object).Count)].fqdn
        echo $myEx
    }
}
function createDelay($seconds){Start-Sleep -s $seconds}
function createCustomerUPN($newCustDomain){Set-ADForest -Identity $hostedExchange_domain -UPNSuffixes @{Add=$newCustDomain} | out-file -Filepath $SessionTranscript -append}
function createCustomerOUs($newCustName){
	New-ADOrganizationalUnit -Name $newCustName -Path $hostedExchange_baseDN -Server $hostedExchange_dchost
	$newCustOUs = "Users","Contacts","Groups","ResourceMailboxes","SharedMailboxes"
	$newCustOUs | ForEach-Object{New-ADOrganizationalUnit -Name $_ -Path "OU=$newCustName,$hostedExchange_baseDN" -Server $hostedExchange_dchost}
}
function createAcceptedDomain($newCustDomain){new-AcceptedDomain -Name $newCustDomain -DomainName $newCustDomain -DomainType 'Authoritative' | out-file -Filepath $SessionTranscript -append}
function createEmailAddressPolicies($newCustName, $newCustDomain){
	new-EmailAddressPolicy -Name $newCustName-Users -RecipientContainer $hostedExchange_baseOU/$newCustName/Users -IncludedRecipients 'AllRecipients' -Priority 'Lowest' -EnabledEmailAddressTemplates SMTP:%g.%s@$newCustDomain -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append
    new-EmailAddressPolicy -Name $newCustName-SharedMailboxes -RecipientContainer $hostedExchange_baseOU/$newCustName/SharedMailboxes -IncludedRecipients 'AllRecipients' -Priority 'Lowest' -EnabledEmailAddressTemplates SMTP:%s@$newCustDomain -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append
    new-EmailAddressPolicy -Name $newCustName-ResourceMailboxes -RecipientContainer $hostedExchange_baseOU/$newCustName/ResourceMailboxes -IncludedRecipients 'AllRecipients' -Priority 'Lowest' -EnabledEmailAddressTemplates SMTP:%s@$newCustDomain -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append
	new-EmailAddressPolicy -Name $newCustName-Groups -RecipientContainer $hostedExchange_baseOU/$newCustName/Groups -IncludedRecipients 'AllRecipients' -Priority 'Lowest' -EnabledEmailAddressTemplates SMTP:%m@$newCustDomain -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append
}
function createAddressLists($newCustName){
	new-AddressList -Name "$newCustName - All Users" -RecipientContainer $hostedExchange_baseOU/$newCustName -IncludedRecipients 'MailboxUsers' -Container '\' -DisplayName "$newCustName - All Users"  -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append
	new-AddressList -Name "$newCustName - All Groups" -RecipientContainer $hostedExchange_baseOU/$newCustName/Groups -IncludedRecipients 'MailGroups' -Container '\' -DisplayName "$newCustName - All Groups" -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append
	new-AddressList -Name "$newCustName - All Contacts" -RecipientContainer $hostedExchange_baseOU/$newCustName/Contacts -IncludedRecipients 'MailContacts' -Container '\' -DisplayName "$newCustName - All Contacts" -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append
	new-AddressList -Name "$newCustName - All Rooms" -RecipientContainer $hostedExchange_baseOU/$newCustName/ResourceMailboxes -IncludedRecipients 'Resources' -Container '\' -DisplayName "$newCustName - All Rooms" -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append
}
function createCustomerGAL($newCustName){New-GlobalAddressList –Name "$newCustName – Global Address List” -IncludedRecipients 'AllRecipients' –RecipientContainer "$hostedExchange_baseOU/$newCustName"  -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append}
function createCustomerOAB($newCustName){New-OfflineAddressBook -Name $newCustName -AddressLists "\$newCustName - All Groups","\$newCustName - All Contacts","\$newCustName - All Rooms","\$newCustName - All Users" -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append}
function createCustomerABP($newCustName){New-AddressBookPolicy -Name $newCustName -GlobalAddressList "\$newCustName – Global Address List” -OfflineAddressBook "\$newCustName" -RoomList "\$newCustName - All Rooms" -AddressLists "\$newCustName - All Contacts","\$newCustName - All Groups","\$newCustName - All Rooms","\$newCustName - All Users" -DomainController $hostedExchange_dchost | out-file -Filepath $SessionTranscript -append}
function createCustomerAttributes($newCustName, $newCustShortName, $newCustDomain){
    get-ADOrganizationalUnit -Identity "OU=$newCustName,$hostedExchange_baseDN" | Set-ADOrganizationalUnit -Description $newCustShortName
    $custAcceptedDomains = $(get-ADOrganizationalUnit -Identity "OU=$newCustName,$hostedExchange_baseDN" -Properties proxyAddresses)
    $custAcceptedDomains.proxyAddresses = ("PRIMARY:" + $newCustDomain)
    Set-ADOrganizationalUnit -Instance $custAcceptedDomains
}
function updateAddressLists($newCustName){
	update-AddressList -Identity "\$newCustName - All Rooms"
    update-AddressList -Identity "\$newCustName - All Users"
    update-AddressList -Identity "\$newCustName - All Contacts"	
	update-AddressList -Identity "\$newCustName - All Groups"
}
function updateEmailAddressPolicies($newCustName){
	update-EmailAddressPolicy -Identity $newCustName-Groups
	update-EmailAddressPolicy -Identity $newCustName-SharedMailboxes
	update-EmailAddressPolicy -Identity $newCustName-ResourceMailboxes
    update-EmailAddressPolicy -Identity $newCustName-Users
}
function verifyEmailAddressPolicies($newCustName){
    $verifyFail="0"
	if(-not (Get-EmailAddressPolicy $newCustName-Users -ErrorAction SilentlyContinue)){logMessageSilently error ("EmailAddressPolicy " + $newCustName + "-Users doesn't exist!"); $verifyFail="1"}
    if(-not (Get-EmailAddressPolicy $newCustName-SharedMailboxes -ErrorAction SilentlyContinue)){logMessageSilently error ("EmailAddressPolicy " + $newCustName + "-SharedMailboxes doesn't exist!"); $verifyFail="1"}
    if(-not (Get-EmailAddressPolicy $newCustName-ResourceMailboxes -ErrorAction SilentlyContinue)){logMessageSilently error ("EmailAddressPolicy " + $newCustName + "-ResourceMailboxes doesn't exist!"); $verifyFail="1"}
    if(-not (Get-EmailAddressPolicy $newCustName-Groups -ErrorAction SilentlyContinue)){logMessageSilently error ("EmailAddressPolicy " + $newCustName + "-Groups doesn't exist!"); $verifyFail="1"}
    if($verifyFail -eq "0"){echo $true}else{echo $false}
}
function readyToStart($BusinessName, $ShortName, $EmailDomainName){
    if($(validateCustomerName $BusinessName) -eq $false){
        logMessage error ("BusinessName error!"); echo $false
    }elseif($(validateCustomerShortname $ShortName) -eq $false){
        logMessage error ("ShortBusinessName error!"); echo $false
    }elseif($(validateCustomerDomain $EmailDomainName) -eq $false){
        logMessage error ("EmailDomainName error!"); echo $false
    }else{
        echo $true
    }
}


# Static Variables
$hostedExchange_ClientOU="HostedClients"                                            # relative OU to root of domain
$LogFile="C:\manage\scripts\new-hecustomer.log"                                     # log all messages to this location
$SessionTranscript="C:\manage\scripts\new-hecustomer.SessionReport." + $BusinessName + ".log"         # log all messages to this location
#
# Dynamic Global Variables
#
# --> Set base domain info (DN,Canonical,DNSRoot)
$ad = Get-ADDomain
$hostedExchange_domain = $ad.DNSroot
$hostedExchange_baseDN = "OU=" + $hostedExchange_ClientOU + "," + $ad.DistinguishedName
$hostedExchange_baseOU =  $hostedExchange_domain + "/" + $hostedExchange_ClientOU
#
# --> Select Domain Controller
$hostedExchange_dchost=$(Select-MyDomainController)
#
# --> Select Exchange Server
$hostedExchange_exhost=$(Select-MyExchangeHost)

# Start of show
if($(readyToStart $BusinessName $ShortName $EmailDomainName) -eq $true){
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$hostedExchange_exhost/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session -AllowClobber -WarningAction SilentlyContinue
        # -- UPN
	    createCustomerUPN $EmailDomainName
        # -- Organisations Units
        createCustomerOUs $BusinessName
        # -- Pause processing temporarily to allow for replication (fix this later)
        createDelay 20
        # -- -- Custom HostedExchange Attributes (Root OU Description + proxyAddresses)
        createCustomerAttributes $BusinessName $ShortName $EmailDomainName
        # -- Accepted Domain
	    createAcceptedDomain $EmailDomainName
        createDelay 20
        # -- Email Address Policies
        createEmailAddressPolicies $BusinessName $EmailDomainName
        createDelay 20
        # -- Address Lists
        createAddressLists $BusinessName
        createDelay 20
        # -- Global Address List
        createCustomerGAL $BusinessName
        createDelay 20
        # -- OAB
        createCustomerOAB $BusinessName
        createDelay 20
        # -- Address Book Policy
        createCustomerABP $BusinessName
        createDelay 20
        # -- Update Email Address Policies
        createDelay 20
        updateEmailAddressPolicies $BusinessName
        # -- Update Address Lists
        createDelay 20
        updateAddressLists $BusinessName
    Remove-PSSession $Session
}

