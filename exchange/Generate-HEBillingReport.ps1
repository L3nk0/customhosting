# Import modules and snapins
if ( (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) -eq $null ){Add-PsSnapin Microsoft.Exchange.Management.PowerShell.E2010}else{Write-Host -BackgroundColor Black -ForegroundColor Red "The MSExchange 2010 Management tools do not appear to be installed! Exiting."; exit}
if ( (Get-module -Name ActiveDirectory -ErrorAction SilentlyContinue) -eq $null ){Get-module ActiveDirectory}else{Write-Host -BackgroundColor Black -ForegroundColor Red "The MSActiveDirectory module to not appear to be installed! Exiting."; exit}


# Static Variables
$hostedExchange_ClientOU="HostedClients"  # relative OU to root of domain
$emailTo = "accounts@somedomain.com.au"
$emailFrom = "billingreport@somedomain.com.au"
$smtpserver = "somehost.somedomain.com.au"
$subject = "Hosted Exchange customer/mailboxes/quota billing report"
$body = "See attachment for details"

# Functions
function ConvertFrom-DN
{
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

Function Get-Increment([float] $value, [int] $increment){    
    if($value -gt 1)
    {
      [Math]::Ceiling($value / $increment) * $increment;
    }
    else
    {
      [math]::Ceiling($value)    
    }    
}
function Select-MyDomainController{
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
#
# --> misc
$date = Get-Date -UFormat "%Y-%m-%d"
$csv = "c:\temp\HE_customer_details-$date.csv"
#
# Create an array where all the report's objects are stored 
$myCollection = @()
#
# Processing
$subous = "Contacts","Groups","ResourceMailboxes","SharedMailboxes","Users"
$customers = Get-ADOrganizationalUnit -Searchbase $hostedexchange_baseDN -SearchScope 1 -Filter *
foreach($customer in $customers){
    $custname = $($customer).Name
    $custou = $hostedExchange_baseOU + "/" + $custname
    $count = 0
    $quota = 0
    foreach($subou in $subous){
        $currentOU = $custou + "/" + $subou
        $mailboxes = Get-Mailbox -OrganizationalUnit $currentOU
        foreach($mailbox in $mailboxes){
            $count = $count + 1
            $quota = $quota + $(Get-MailboxStatistics $mailbox | select @{label="quota";expression={$_.TotalItemSize.Value.ToMB()}}).quota
        }
        
    }
    $allocatedQuota = ""
    if($($count * 2048 -le $(Get-Increment -value $quota -increment 2048))){$allocatedQuota = $(Get-Increment -value $quota -increment 2048)}else{$allocatedQuota = $count * 2048}
    $diffQuota = $allocatedQuota - $($count *2048)

    # Define the attributes of the object
    $myobj = "" | select customer,mailboxCount,usedQuota,allocatedQuota,additionalQuota
    
    # Attach values to the object attributes
    $myobj.customer = $custname
    $myobj.mailboxCount = $count
    $myobj.usedQuota = $quota
    $myobj.allocatedQuota = $allocatedQuota
    $myobj.additionalQuota = $diffQuota

    # Add the object to the array
    $myCollection += $myobj
}

$myCollection | ft