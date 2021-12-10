$Outlook = New-Object -ComObject Outlook.Application
#Username variable
$userID = Read-Host "Please enter the Username: "
#Request number variable
$userSubject = read-host "Please enter the request number: "
#User email address is stored in this variable
$emailAddress = Read-Host "Please enter the user's email address: "
#User PIN is stored in this variable
$userPIN = Read-Host "Please enter the user's PIN number: "
#User initals stored in this variable
$userPW = Read-Host "Please enter the user's initials, all lowercase: "
$optionFlag = 0

While ($optionFlag -eq 0)
{

#Confirms that the user entered the correct information
write-host("You have entered user name: '" + $userID + "' and request # " + $userSubject + " and email address '" + $emailAddress + "' and PIN: '" + $userPIN + "' and user initials '" + $userPW + "' is this all correct?`nWARNING IT IS VERY IMPORTANT THIS IS CORRECT!")
$userResponse = read-host "`n`nIs this correct? Enter Y(Yes)/N(No): "

#If user selects yes, both emails are sent using all data provided by user.
if ($userResponse -eq "Y" -and $userResponse -ne "N")
{
#User account details email.
$Mail = $Outlook.CreateItem(0)
$Mail.Recipients.Add($emailAddress)
$Mail.Subject = "Request # " + $userSubject
$Mail.Body = "Hello,`n`nYour user ID is: " + $userID + "`n`nPlease make sure to keep this in a safe place.`n`nFrom,`n`nWhoever"
$Mail.Send()

#User PIN/Password email.
$Mail = $Outlook.CreateItem(0)
$Mail.Recipients.Add($emailAddress)
$Mail.Subject = "fyi"
$Mail.Body = "Here is your PIN number and password, please keep these in a safe place.`n`nPIN: " + $userPIN + "`nPassword: " + "#" + $userPW + $userPIN + "`n`nFrom,`n`nWhoever"
$Mail.Send()
Write-Host("Account Email sent`nPassword Email Sent")
$optionFlag = 1
$emailCheckFlag = 1
}

#If user selects no, no emails are sent and program is terminated.
if ($userResponse -eq "N" -and $userResponse -ne "Y")
{
$optionFlag = 1
Write-Host("Program Exited, no emails sent")
}

#If statement that recognizes the user mistyped and allows them the opportunity to check their input again.
if ($userResponse -ne "Y" -and $userResponse -ne "N")
{
Write-Host("Looks like you did not select Y or N.`nPress any key to try again...")
Read-Host
$optionFlag = 0
}


}



