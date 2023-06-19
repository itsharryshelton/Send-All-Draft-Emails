# Send-All-Draft-Emails
.exe that will allow a user to send all draft emails within Outlook.

Tested working within Windows 11 Enterprise Multi Session with Outlook 365.

Why was this created: 
SAGE importing client emails into the Draft folder of Outlook, requires user to send them all, this was originally done with VSCode whgich was within Outlook itself. Later versions of 365 caused this to stop working for no reason.
As a result, I created this code instead that uses Powershell that hopefully doesn't stop working in the future because Microsoft changed something.

Run the .exe, enter the email address of the draft folder location (It needs to know which email, even if you only have one, but that does mean it allows for multi email outlook setups), click Ok, then it will run through.

![image](https://github.com/itsharryshelton/Send-All-Draft-Emails/assets/136495601/7d8a704f-285a-4e67-be3b-465413105a50)

I did find that if you have more than 250 emails, it may struggle to pick them all up, the only way at the moment to resolve this is to just rerun the application and hope it picks it up again.

I've provided the source code within the Source folder, called "SendAllDraft.ps1", feel free to adjust or improve the code. The icon for the exe is also in the folder for use too.
