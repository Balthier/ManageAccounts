
# Handles the logging to file
Function toggleLogging($toggle) {
	## Function Variables
	# Current date
	$Date = Get-Date
	# Current time: HH:MM:SS
	$Time = "{0}:{1:d2}:{2:d2}" -f $date.hour, $date.minute, $date.second
	# Get the base path of the main/calling script
	$scriptPath = $MyInvocation.PSCommandPath
	# Set the log directory
	$logPath = Split-Path -Parent $scriptPath
	$logDir = "$logPath\Logs"
	# Set the log file
	$logFile = "{0}-{1:d2}-{2:d2}.log" -f $date.year, $date.month, $date.day
	
	# Turn on the logs
	if ($toggle -eq "On") {
		# Create the log directory, if it doesn't exist
		New-Item $LogDir -Type Directory -Force
		try { 
			# Make each log distinguishable
			Add-Content "$logDir\$logFile" "`n"
			Add-Content "$logDir\$logFile" "`nCurrent Date/Time - $Date"
			Add-Content "$logDir\$logFile" "`n"
			# Attempt to start logging
			Start-Transcript -Path "$logDir\$logFile" -Append
		}
		# On fail, stop the previous logging attempt 
		catch { 
			Stop-Transcript
			Add-Content "$logDir\$logFile" "`n"
			Add-Content "$logDir\$logFile" "`nCurrent Date/Time - $Date"
			Add-Content "$logDir\$logFile" "`n"
			Start-Transcript -Path "$logDir\$logFile" -Append
		}
	}
	# Turn off the logs
	if ($toggle -eq "Off") {
		Stop-Transcript
	}
}

Function connectAD($ADServer,$Creds,$eLevel,$wLevel) {
	Write-Host "`n"
	Write-Host "Opening connection to $ADServer"
	$session = new-pssession -computer $ADServer -Credential $Creds -ErrorAction $eLevel -WarningAction $wLevel
	Invoke-Command -Session $session -ScriptBlock {Import-Module "C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1" -DisableNameChecking} -ErrorAction $eLevel -WarningAction $wLevel
	Import-PSSession -session $session -module ActiveDirectory -AllowClobber -DisableNameChecking -ErrorAction $eLevel -WarningAction $wLevel
	return $session
}

Function disconnectAD($session) {
	Write-Host "`n"
	Write-Host "Closing connection to $ADServer..."
	Remove-PSSession $session
}

Function connectO365($Creds,$eLevel,$wLevel) {
	Write-Host "`n"
	Write-Host "Opening connection to Office 365"
	$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Creds -Authentication Basic -AllowRedirection -ErrorAction $eLevel -WarningAction $wLevel
	Import-PSSession -session $O365Session -AllowClobber -DisableNameChecking -ErrorAction $eLevel -WarningAction $wLevel
	Connect-MsolService -Credential $Creds -ErrorAction $eLevel -WarningAction $wLevel
	return $O365Session
}

Function disconnectO365($O365Session) {
	Write-Host "`n"
	Write-Host "Closing connection to Office 365..."
	Remove-PSSession $O365Session
}

Function getDefaults() {
	$wLevel = "SilentlyContinue"
	$eLevel = "Stop"
	$EmailDomain = "COMPANY.COM"
	$Domain = "COMPANY.DOMAIN.LOCAL"
	$CurrentUser = $env:UserName
	$Creds = Get-Credential -Credential "$CurrentUser@$EmailDomain"
	$ADServer = "Active-Directory-Server-01"
	return $wLevel,$eLevel,$EmailDomain,$Domain,$CurrentUser,$Creds,$ADServer
}

Function setColourScheme() {
	$Shell = $Host.UI.RawUI
	$Shell.BackgroundColor = "Black"
	$Shell.ForegroundColor = "Green"
	Clear
}