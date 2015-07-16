# This script with generate a report based on pooled quota.
# e.g.
# Each mailbox sold comes with 2GB of storage. Customer A has 10 mailboxes which gives them 20GB's of 
# pooled quota. If User 1 at Customer A uses more than 2GB of storage, as long as the organisation is
# below the pooled quota limit, no additional charges will be added.
#
PARAM(
[STRING]$EmailReport
)
# Import modules and snapins
if ( (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) -eq $null ){Add-PsSnapin Microsoft.Exchange.Management.PowerShell.E2010}
if ( (Get-module -Name ActiveDirectory -ErrorAction SilentlyContinue) -eq $null ){Import-module ActiveDirectory}

# Static Variables
$hostedExchange_ClientOU="HostedClients"  # relative OU to root of domain
$emailTo = "accounts@somedomain.com.au"
$emailFrom = "billingreport@somedomain.com.au"
$smtpserver = "somehost.somedomain.com.au"
$subject = "Hosted Exchange customer/mailboxes/quota billing report"
$body = "See attachment for details"
$date = (Get-Date -UFormat "%Y-%m-%d")
$csv = "c:\temp\HE_customer_details-$date.csv"
# Costings
$mbcost="20" # ExGST cost per mailbox in dollars
$aqcost="5"  # ExGST cost per additional 1GB of quota in dollars

# Functions
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
function Output-Report{
    if($EmailReport -eq $true){
        $myCollection | Export-Csv $csv -NoTypeInformation
        Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -SmtpServer $smtpserver -Body $body -Attachments $csv
    }else{
    echo ("All Quota's are shown in MegaBytes"); $myCollection | ft -autosize; echo ("Total Cost for all services exGST: $" + $myTotalCost); echo ("Total Number of Mailboxes: " + $myTotalMBCount)
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
$myTotal = 0
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
    
        if($($count * 2048 -le $(Get-Increment -value $quota -increment 2048))){
            $allocatedQuota = $(Get-Increment -value $quota -increment 2048)
        }else{
            $allocatedQuota = $count * 2048
        }
    
        $poolquota = ($count * 2048)
        $diffQuota = $allocatedQuota - $poolquota

        # Generate the costings
        $additionalQuotaCost=(($diffQuota/1024)*$aqcost)
        $baseCost=(($count)*$mbcost)
        $totalCost=($additionalQuotaCost + $baseCost)
        $myTotalCost=($totalCost + $myTotalCost)

        # Total MailboxCount
        $myTotalMBCount=($count + $myTotalMBCount)
    
        # Define the attributes of the object
        $myobj = "" | select Customer,Mailboxes,PoolQuota,QuotaUsed,ExtraQuota,BaseCost,ExtraQuotaCost,TotalCost
    
        # Attach values to the object attributes
        $myobj.Customer = $custname
        $myobj.Mailboxes = $count
        $myobj.PoolQuota = $poolquota
        $myobj.QuotaUsed = $quota
        $myobj.ExtraQuota = $diffQuota
        $myobj.BaseCost = ("$" + $baseCost)
        $myobj.ExtraQuotaCost = ("$" + $additionalQuotaCost)
        $myobj.TotalCost = ("$" + $totalCost)

        # Add the object to the array
        $myCollection += $myobj
}

Output-Report