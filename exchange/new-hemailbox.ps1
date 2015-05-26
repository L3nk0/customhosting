PARAM(
[STRING]$BusinessName=$(Read-Host -prompt "What is the customers Full Business Name? e.g. ExamplePtyLtd"),
[STRING]$DisplayName=$(Read-Host -prompt "What is the DisplayName? e.g. 'Billy Bob' or 'Accounts @ Some Business'"),
[STRING]$GivenName=$(Read-Host -prompt "What is the users name? e.g. 'Billy', 'Accounts' or 'Info'"),
[STRING]$Surname,
[STRING]$Password,
[STRING]$DomainController,
[STRING]$ExchangeServer
)
#
# Import modules and snapins
if ( (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) -eq $null ){Add-PsSnapin Microsoft.Exchange.Management.PowerShell.E2010}
if ( (Get-module -Name ActiveDirectory -ErrorAction SilentlyContinue) -eq $null ){Import-module ActiveDirectory}
#
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
    $tlds=".com",".org",".net",".edu",".asn",".co.uk",".co.nz",".info",".biz",".com.au",".net.au",".org.au",".edu.au",".asn.au"
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
function Create-Delay($seconds){Start-Sleep -s $seconds}
function Generate-Password{
    $words=(import-csv $wordlist | get-random -count 2)
    $a=(($words[0].words).substring(0,1).toupper()+($words[0].words).substring(1).tolower())
    $b=(($words[1].words).substring(0,1).toupper()+($words[1].words).substring(1).tolower())
    $num=(Get-Random -Minimum 9 -Maximum 100)
    echo ($a + $b + $num)
}
function Generate-Username($ShortName, $GivenName, $Surname){
    $user=""
    $finaluser=""
    if($Surname -ne $null){
        $user=($ShortName + "." + $GivenName + "." + $Surname)
    }elseif($Surname -eq $null){
        $user=($shortName + "." + $GivenName)
    }
    if($user.length -ge "20"){
        $free="false"
        $count=0
        while($free -eq "false"){
            $test=$(try{get-aduser ($user.substring(0,19) + $count)}catch{echo $false})
            if($test -eq $false){
                $finaluser=($user.substring(0,19) + $count)
                $free="true" 
            }elseif($test -ne $false){
                $count++
            }
        }
    }
    if($finaluser -ne ""){echo $finaluser}elseif($finaluser -eq ""){echo $false}
}
function Generate-UPN($Domain, $GivenName, $Surname){
    if($Surname -ne $null){
        echo ($GivenName + "." + $Surname + "@" + $Domain)
    }elseif($Surname -eq $null){
        echo ($GivenName + "@" + $Domain)
    }
}
function Check-CustomerExists($custName){
    if((Get-ADOrganizationalUnit -Filter 'Name -like $custName' -SearchBase $hostedExchange_baseDN) -ne $null){
   		echo $true
   	}else{
   		if((Get-ADOrganizationalUnit -Filter 'Name -like $custName' -SearchBase $hostedExchange_baseDN) -eq $null){
   			echo $false
   		}else{
   			echo $false
   	    }
    }
}
function Get-CustomerShortname($custName){
    $test=$(Try {(get-ADOrganizationalUnit -Identity "OU=$custName,$hostedExchange_baseDN" -Properties Description).description}Catch{echo $false})
    if($test -ne $false){
        $(get-ADOrganizationalUnit -Identity "OU=$custName,$hostedExchange_baseDN" -Properties Description).description
    }elseif($test -eq $false){
        echo $false
    }else{
        # Something fucked up
        echo $false
    }
}
function Get-CustomerUPN($custName){
    $test=$(Try {((get-ADOrganizationalUnit -Identity "OU=$custName,$hostedExchange_baseDN" -Properties proxyAddresses).proxyAddresses)}Catch{echo $false})
    if($test -ne $false){
        $(get-ADOrganizationalUnit -Identity "OU=$custName,$hostedExchange_baseDN" -Properties proxyAddresses).proxyaddresses | foreach-object{
            if($_ -like "PRIMARY:*"){
                $var = $_.split(':')
                if($var[0] -eq "PRIMARY"){
                    echo $var[1]
                }
            }
        }
    }elseif($test -eq $false){
        echo $false
    }else{
        #something fucked up
        echo $false
    }
}
function Create-HEUserAccount($UserName, $UPN, $DisplayName, $Password, $GivenName, $Surname){
    if($Surname -ne $null){
        New-ADUser -SamAccountName $UserName -UserPrincipalName $UPN -Name $UserName -GivenName $GivenName -Surname $Surname -DisplayName $DisplayName -Path $userOU -AccountPassword $(convertto-securestring $Password -asplaintext -force) -PasswordNeverExpires $True -Enabled $True -Server $hostedExchange_dchost  | out-file -Filepath $SessionTranscript -append
    }elseif($Surname -eq $null){
        New-ADUser -SamAccountName $UserName -UserPrincipalName $UPN -Name $UserName -GivenName $GivenName -DisplayName $DisplayName -Path $userOU -AccountPassword $(convertto-securestring $Password -asplaintext -force) -PasswordNeverExpires $True -Enabled $True -Server $hostedExchange_dchost  | out-file -Filepath $SessionTranscript -append
    }
}
function Create-HEUserMailbox($UserName, $BusinessName){
    Enable-mailbox $UserName -AddressBookPolicy $BusinessName -Database MailboxStore1 -Alias $UserName -DomainController $hostedExchange_dchost  | out-file -Filepath $SessionTranscript -append
}
function readyToStart($BusinessName){
    if((Check-CustomerExists $BusinessName) -eq "True"){
        if((Get-CustomerShortname $BusinessName) -ne "False"){
            if((Get-CustomerUPN $BusinessName) -ne "False"){
                echo $true
            }else{
                logMessage error ("There was an error finding the UPN for " + $BusinessName)
                echo $false
            }
        }else{
            logMessage error ("There was an error checking the shortname for " + $BusinessName)
            echo $false
        }
    }else{
        logMessage error ("BusinessName " + $BusinessName + " does not exist!")
        echo $false    
    }
}
#
# Static Variables
$wordlist="C:\Manage\Scripts\words.csv"                # Path to a csv full of words. See ftp://10.20.30.211/clientfiles/words.csv
$hostedExchange_ClientOU="HostedClients"               # relative OU to root of domain
$LogFile="C:\manage\scripts\new-hemailbox.log"         # log all messages to this location
$allReports= @()                                       # Somewhere to store reports
#
# Dynamic Global Variables
#
# --> Set base domain info (DN,Canonical,DNSRoot)
$ad = Get-ADDomain
$hostedExchange_domain = $ad.DNSroot
$hostedExchange_baseDN = "OU=" + $hostedExchange_ClientOU + "," + $ad.DistinguishedName
$hostedExchange_baseOU =  $hostedExchange_domain + "/" + $hostedExchange_ClientOU
$userOU = ("OU=Users,OU=" + $BusinessName + "," + $hostedExchange_baseDN)
#
# --> Select Domain Controller
$hostedExchange_dchost=$(Select-MyDomainController)
#
# --> Select Exchange Server
$hostedExchange_exhost=$(Select-MyExchangeHost)
#
# Start of show
if($(readyToStart $BusinessName) -eq $true){
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$hostedExchange_exhost/PowerShell/ -Authentication Kerberos
    Import-PSSession $Session -AllowClobber
        
        if($surname -eq ""){
        
        # --> Get the Customers Shortname
        $Shortname=(Get-CustomerShortname $BusinessName)
        # --> Get the Customers Primary Domain
        $Domain=(Get-CustomerUPN $BusinessName)
        # --> Generate the Users Password
        $Password=(Generate-Password)
        # --> Generate the Users Username
        $UserName=(Generate-Username $Shortname $GivenName)
        # --> Generate the Users UPN
        $UPN=(Generate-UPN $Domain $GivenName)
        # --> Set the Session Report location
        $SessionTranscript="C:\manage\scripts\new-hemailbox.SessionReport." + $Shortname + "." + $UserName + ".log"

        # --> Create the AD User Account
        Create-HEUserAccount $UserName $UPN $DisplayName $Password $GivenName

        # --> Create the Users Mailbox
        Create-HEUserMailbox $UserName $BusinessName

        # --> Generate a report
        $thisReport = "" | select BusinessName,DisplayName,UserName,UPN,Password
        $thisReport.BusinessName = $BusinessName
        $thisReport.DisplayName = $DisplayName
        $thisReport.UserName = $UserName
        $thisReport.UPN = $UPN
        $thisReport.Password = $Password
        $allReports += $thisReport

        }else{
        
        # --> Get the Customers Shortname
        $Shortname=(Get-CustomerShortname $BusinessName)
        # --> Get the Customers Primary Domain
        $Domain=(Get-CustomerUPN $BusinessName)
        # --> Generate the Users Password
        $Password=(Generate-Password)
        # --> Generate the Users Username
        $UserName=(Generate-Username $Shortname $GivenName $Surname)
        # --> Generate the Users UPN
        $UPN=(Generate-UPN $Domain $GivenName $Surname)
        # --> Set the Session Report location
        $SessionTranscript="C:\manage\scripts\new-hemailbox.SessionReport." + $Shortname + "." + $UserName + ".log"

        # --> Create the AD User Account
        Create-HEUserAccount $UserName $UPN $DisplayName $Password $GivenName $Surname

        # --> Create the Users Mailbox
        Create-HEUserMailbox $UserName $BusinessName

        # --> Generate a report
        $thisReport = "" | select BusinessName,DisplayName,UserName,UPN,Password
        $thisReport.BusinessName = $BusinessName
        $thisReport.DisplayName = $DisplayName
        $thisReport.UserName = $UserName
        $thisReport.UPN = $UPN
        $thisReport.Password = $Password
        $allReports += $thisReport    
        
        }
                
    Remove-PSSession $Session
}
$allReports | ft
