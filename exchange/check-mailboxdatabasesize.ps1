# To execute from within NSClient++
#
#[NRPE Handlers]
#check_exmbdbsize=cmd /c echo C:\Scripts\Nagios\Check-MailboxDatabaseSize.ps1 | PowerShell.exe -Command -
#
#
# On the check_nrpe command include the -t 30, since it takes some time to load the Exchange cmdlet's.

Add-PsSnapin Microsoft.Exchange.Management.PowerShell.E2010

$NagiosStatus = "0"
$NagiosDescription = ""

Get-mailboxdatabase | foreach-object{
    $whitespace=((Get-MailboxDatabase -Status $_.name).AvailableNewMailboxSpace).ToMB()
    $totalspace=((Get-MailboxDatabase -Status $_.name).DatabaseSize).ToMB()

    # Look for large databases - critical if over 160GB
    if($totalsize -ge "160000"){
        # Format the output for Nagios
	    if ($NagiosDescription -ne "") 	{$NagiosDescription = $NagiosDescription + ", "}
			$NagiosDescription = $NagiosDescription + "Mailbox Database " + $_.name + " has reached " + $totalspace + "MB in size."
			
		# Set the status to failed.
		$NagiosStatus = "2"

    # Look for large databases with spare whitespace - Warning if between 140GB and 160GB with less than 2GB whitespace
    }elseif(($totalsize -ge "140000") -and ($totalsize -le "160000") -and ($whitespace -le "2048")){
        # Format the output for Nagios
	    if ($NagiosDescription -ne "") 	{$NagiosDescription = $NagiosDescription + ", "}
			$NagiosDescription = $NagiosDescription + "Mailbox Database " + $_.name + " has reached " + $totalspace + "MB in size and has less than 2048MB of whitespace."
			
	    # Don't lower the status level if we already have a critical event
	    if ($NagiosStatus -ne "2") {
			$NagiosStatus = "1" }
	}

}

# Output, what level should we tell our caller?
if ($NagiosStatus -eq "2") {
	Write-Host "CRITICAL: " $NagiosDescription
} elseif ($NagiosStatus -eq "1") {
	Write-Host "WARNING: " $NagiosDescription
} else {
	Write-Host "OK: All mail queues within limits."
}

exit $NagiosStatus
