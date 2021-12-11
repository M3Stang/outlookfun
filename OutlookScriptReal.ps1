#Marco Testoni
#12/09/2021
#Double Email Project

Do{

$appFlag = 0
$Outlook = New-Object -ComObject Outlook.Application

#User email address is stored in this variable
$emailAddressFirst = Read-Host "Please enter the user's email address"

#Variable used to confirm the email address the user types.
$confirmedEmail = Read-Host "Please confirm the email address once more"

#Verifies the emails match before continuing with the rest of the questions.
if ($confirmedEmail -eq $emailAddressFirst)
{

#Username variable
$userID = Read-Host "Please enter the Username"

#Request number variable
$userSubject = Read-Host "Please enter the request number"

#User PIN is stored in this variable
$userPIN = Read-Host "Please enter the user's PIN number"

#User initals stored in this variable
$userPW = Read-Host "Please enter the user's initials, all lowercase"

$optionFlag = 0

While ($optionFlag -eq 0)
{

#Confirms that the user entered the correct information
Write-Host("`n`nCONFIRMATION`n`nEmail Address: '" + $confirmedEmail + "'`nUsername: '" + $userID + "'`nRequest #: '" + $userSubject + "'`nPIN: '" + $userPin + "'`nInitials: '" + $userPW + "'`n`nIs this all correct?`nWARNING IT IS VERY IMPORTANT THAT THIS IS CORRECT!")
$userResponse = Read-Host "`n`nIs this correct? Type 'yes' to confirm or 'no' to re-enter the information: "

#If user selects yes, both emails are sent using all data provided by user.
if ($userResponse -eq "yes" -and $userResponse -ne "no")
{
#User account details email.
$Mail = $Outlook.CreateItem(0)
$Mail.Recipients.Add($confirmedEmail)
$Mail.Subject = "Request # " + $userSubject
$Mail.Body = "Hello,`n`nYour user ID is: " + $userID + "`n`nPlease make sure to keep this in a safe place.`n`nFrom,`n`nWhoever"
$Mail.Send()

#User PIN/Password email.
$Mail = $Outlook.CreateItem(0)
$Mail.Recipients.Add($confirmedEmail)
$Mail.Subject = "fyi"
$Mail.Body = "Here is your PIN number and password, please keep these in a safe place.`n`nPIN: " + $userPIN + "`nPassword: " + "#" + $userPW + $userPIN + "`n`nFrom,`n`nWhoever"
$Mail.Send()
Write-Host("Account Email sent`nPassword Email Sent")
$optionFlag = 1
$appFlag = 1
}

#If user selects no, no emails are sent and program restarts
if ($userResponse -eq "no" -and $userResponse -ne "yes")
{

Write-Host("You specified no. You will now return to the beginning of the program to retype the information. Press any key to continue...")
Read-Host
Clear-Host
$optionFlag = 1

}

#If statement that recognizes the user mistyped and allows them the opportunity to check their input again.
if ($userResponse -ne "yes" -and $userResponse -ne "no")
{

Write-Host("You did not type 'yes' or 'no'.`nPress any key to try again...")
Read-Host
$optionFlag = 0

}

}

} 

else
{
Write-Host("The email addresses entered do not match. Press any key to try again...")
Read-Host
Clear-Host
}
} While ($appFlag -eq 0)




