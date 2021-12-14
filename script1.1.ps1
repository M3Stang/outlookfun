#Marco Testoni
#12/09/2021
#Double Email Project
#Ver 1.1
#Revised 12/14/2021
#Added additional comments and allowed the user to keep running the program without having to relaunch it every time it sends emails.

##Do-while loop that allows the program to keep running/restart as necessary
Do{

#Flag that is used in the Do-While loop that allows the program to keep running.
$appFlag = 0
#Connects to an active Outlook Instance.
$Outlook = New-Object -ComObject Outlook.Application

#User email address is stored in this variable
$emailAddressFirst = Read-Host "Please enter the user's email address"

#Variable used to confirm the email address the user types.
$confirmedEmail = Read-Host "Please confirm the email address once more"

##Verifies the emails match before continuing with the rest of the questions.
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

#optionFlag that is a condition in running the below while loop.
$optionFlag = 0

##While loop that allows the user to confirm their selection and allows the user to make sure they are choosing either yes or no before terminating the application or sending the emails.
While ($optionFlag -eq 0)
{

#Confirms that the user entered the correct information
Write-Host("`n`nCONFIRMATION`n`nEmail Address: '" + $confirmedEmail + "'`nUsername: '" + $userID + "'`nRequest #: '" + $userSubject + "'`nPIN: '" + $userPin + "'`nInitials: '" + $userPW + "'`n`nIs this all correct?`nWARNING IT IS VERY IMPORTANT THAT THIS IS CORRECT!")
$userResponse = Read-Host "`n`nIs this correct? Type 'yes' to confirm or 'no' to re-enter the information"

#If user selects yes, both emails are sent using all data provided by user.
if ($userResponse -eq "yes" -and $userResponse -ne "no")
{
##User account details email.

#Outlook creates an email
$Mail = $Outlook.CreateItem(0)
#Sets the "To" field to the confirmed email variable
$Mail.Recipients.Add($confirmedEmail)
#Sets the subject to the request number.
$Mail.Subject = "Request # " + $userSubject
#Sets the body of the email to variables the user has entered.
$Mail.Body = "Hello,`n`nYour user ID is: " + $userID + "`n`nPlease make sure to keep this in a safe place.`n`nFrom,`n`nWhoever"
#Sends the email.
$Mail.Send()

##User PIN/Password email.

#Outlook creates an email.
$Mail = $Outlook.CreateItem(0)
#Sets the "To" field to the confirmed email variable
$Mail.Recipients.Add($confirmedEmail)
#Sets subject
$Mail.Subject = "fyi"
#Sets the body of the email to variables the user has entered.
$Mail.Body = "Here is your PIN number and password, please keep these in a safe place.`n`nPIN: " + $userPIN + "`nPassword: " + "#" + $userPW + $userPIN + "`n`nFrom,`n`nWhoever"
#Sends the email.
$Mail.Send()
#Tells the user that the emails sent successfully, waits for a key press and restarts the program by setting the optionFlag to 1, thus exiting the while loop.
Write-Host("Account Email sent`nPassword Email Sent`nPress any key to run the script again for a new user, or you can close the application.")
Read-Host
Clear-Host
$optionFlag = 1
}

#If user selects no, no emails are sent and program restarts
if ($userResponse -eq "no" -and $userResponse -ne "yes")
{
#Informs the user that they selected no, and that the program will restart, waits for a keypress, clears the screen, and sets the optionFlag to 1 to exit the while loop.
Write-Host("You specified no. You will now return to the beginning of the program to retype the information. Press any key to continue...")
Read-Host
Clear-Host
$optionFlag = 1

}

#If statement that recognizes the user mistyped and allows them the opportunity to check their input again.
if ($userResponse -ne "yes" -and $userResponse -ne "no")
{
#Informs the user that they mistyped their input while allowing them to press any key, keeping the optionFlag at 0 so the user can correctly type their response without having to restart the program.
Write-Host("You did not type 'yes' or 'no'.`nPress any key to try again...")
Read-Host
$optionFlag = 0

}

}

} 

##Else statement for the If statement that checks that both emails answered are correct, if they do not equal one another, the below code block executes.
else
{
#Informs the user that the emails do not match and allows the user to press any key and clears the screen to restart the program.
Write-Host("The email addresses entered do not match. Press any key to try again...")
Read-Host
Clear-Host
}
##End of Do-While loop. This value will never change allowing the program to run forever after opening it once, until the user manually exits it.
} While ($appFlag -eq 0)




