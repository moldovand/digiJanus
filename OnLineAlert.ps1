<#===========================================================================================================================
 Script Name: OnlineAlert.ps1
 Description: Pings computer, if no response waits X minutes tries again. Sends email if successful or max attempts reached.
      Inputs: Remote ComputerName, recipient email address, minutes to wait between retries, and how many attempts.
     Outputs: Ping status, email when ping is successful or when reaching the maximum number of attempts entered.
       Notes: 
     Example: .\OnlineAlert.ps1
      Author: Richard Wright
Date Created: 10/26/2017
     Credits: 
Last Revised: 10/27/2017
   ChangeLog: Date	   Who	Description of changes
              10/26/2017   RMW  Added color for clarity, variable for the sleep time.
              10/27/2017   RMW  Added the maximum number of attempts to try. Email if successful or not.
=============================================================================================================================
Instructions
------------
1. Enter the computer name of the system you want to be notified once it is online.
2. Enter an email address in which to send a status report once the system is detected as online,
   or once your number of attempts has been reached.
   NOTE: In the $SMTPsettings section, be sure the following are correct:
      - From = "NoReply@domain"
	where "NoReply@domain" should be your NoReply email address.
      - SmtpServer = "Your_SMTP_Server"
	where "Your_SMTP_Server" should be your SMTP server.
3. Enter the number of minutes to wait between attempting to contact the system.
4. Enter the number of attempts you want to make to detect if the system is online.

============================================================================================================================#>
Clear

$Computer = Read-Host "Remote computer name"
$EmailTo = Read-Host "Email address to notify"
[int]$SleepTimer = Read-Host "Minutes between attempts"
[int]$SleepSeconds = $SleepTimer * 60
[int]$Attempts = Read-Host "Number of attempts"
[int]$AttemptsCounter = 0
$StartDate = Get-Date -Format D
$StartTime = Get-Date -Format T
Write-Host "Testing to see if $Computer is online..."

Do 
{
   $AttemptsCounter++
   $RemainingAttempts = ([int]$Attempts - [int]$AttemptsCounter)
   $Online = Test-Connection -ComputerName $Computer -Quiet
   If ($Online -NE "True") 
   {
       Write-Host "Computer $Computer is " -NoNewLine
       Write-Host "Offline" -BackgroundColor Red -ForegroundColor Black -NoNewline
       If ($AttemptsCounter -eq $Attempts) {
          Write-Host "."
       }
       Else {
          Write-Host ". Pausing for $SleepSeconds seconds. Remaining attempts: $RemainingAttempts"
       }
   }

   #Check the number of attempts, break out if reached.
   If ($AttemptsCounter -eq $Attempts) {break}

   #Delay
   Start-Sleep -s ($SleepTimer * 60)
}
While ($Online -NE "True")

$EndDate = Get-Date -Format D
$EndTime = Get-Date -Format T
If ($Online -NE "True") {
   Write-Host "Maximum number of attempts reached. Sending email then exiting..."
   }
Else {
   Write-Host
   Write-Host "Computer $Computer is " -NoNewline
   Write-Host "ONLINE" -BackgroundColor Green -ForegroundColor White
   Write-Host "Search began on $StartDate at $StartTime."
   Write-Host "Found online on $EndDate at $EndTime."
   Write-Host "Sending email."
}

<#=============================
Email Settings
===============================#>
#Subject / Body
If ($Online -NE "True") {
   $Subject = "$EndDate at $EndTime $Computer NOT Online"
   $Body = "Search began on $StartDate at $StartTime and on $EndDate at $EndTime after $Attempts attempts $Computer was found to be OFFLINE."
}
Else {
   $Subject = "$EndDate at $EndTime $Computer ONLINE"
   $Body = "Search began on $StartDate at $StartTime and on $EndDate at $EndTime the computer $Computer was found to be ONLINE."
}

$SMTPsettings = @{
	To =  "$EmailTo"
	From = "NoReply@domain"
	Subject = $Subject
	Body = $Body
	SmtpServer = "Your_SMTP_Server"
	}

#Send Email With Retries
$StopEmailLoop=$false
[int]$RetryCount=0

Do {
	Try {
			Send-MailMessage @SMTPsettings -ErrorAction Stop;
			$StopEmailLoop = $true
		}
		Catch {
				If ($RetryCount -gt 5){
					Write-Host "Cannot send email. The script will exit."
					$StopEmailLoop = $true
					}
				Else {
					Write-Host "Cannot send email. Trying again in 15 seconds."
					Start-Sleep -Seconds 15
					$RetryCount ++
					}
		}
	}
While ($StopEmailLoop -eq $false)