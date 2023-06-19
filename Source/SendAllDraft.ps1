# Prompt the user to enter their email address
$emailAddress = Read-Host "Enter your email address"

# Create Outlook Application object
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

# Set drafts folder
$myDraftsFolder = $namespace.Folders.Item($emailAddress).Folders.Item('Drafts')

# Get draft items
$draftItems = $myDraftsFolder.Items | Where-Object { $_.Class -eq 43 }

# Create an array to store the items that need to be sent
$itemsToSend = @()

# Loop through all Draft Items
foreach ($item in $draftItems) {
    if ($item.Subject -like 'RE: *') {
        Write-Verbose "Skipping inline response item: $($item.Subject)" -Verbose
        continue
    }

    if (![string]::IsNullOrEmpty($item.To.Trim())) {
        Write-Verbose "Adding item to send: $($item.Subject)" -Verbose
        $itemsToSend += $item
    }
}

# Loop through the items to send and send each one
foreach ($item in $itemsToSend) {
    $mailItem = $namespace.GetItemFromID($item.EntryID)
    Write-Verbose "Sending: $($mailItem.Subject)" -Verbose
    $mailItem.Send()
}

# Send any remaining unsent items - disabled, weird com issues result in the following code, haven't had time to figure out - if any get missed, best thing is just to rerun script, usually will run through it properly again
# $Outlook.Send()

# Prompt the user to press Enter to close the script - Remove hashtags for two lines below to turn it on
# Write-Host "Press Enter to close the script."
# $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
