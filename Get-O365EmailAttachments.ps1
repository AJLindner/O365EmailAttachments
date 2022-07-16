<#
    .NOTES
        AJ Lindner (AJ@AJLindner.info)
        July 16, 2022

    .SYNOPSIS
        Downloads attachments from emails in an Office 365 mailbox

    .DESCRIPTION
        This is a sample script for automating the process of downloading attachments from messages in an Office 365 mailbox.
        In this specific example, the script processes all unread messages in a specified folder in the user's mailbox.
        For each message, it will download all attachments to a specified location and mark the message as read.

        This sample script is hosted on Github:
        https://github.com/AJLindner/O365EmailAttachments

        There is an article associated with this script:
        <linkedin article here>
#>


# Connect to Azure with a self-signed certificate in the local user store
$CertificateThumbprint = "DB67338C234E5B72531BE069CFA1EF289267EC1D"
$Certificate = Get-ChildItem "cert:\LocalMachine\My\$CertificateThumbprint"

$AzureApp = @{
    ClientID = "b660fa90-bd54-4d3f-aa8c-60d8268a922b"   # The Application (client) ID from your Azure app
    TenantID = "129a584a-2aa5-4eb7-9aab-9cf2c240fc72"   # The Directory (tenant) ID for your Azure tenant
    Certificate = $Certificate 
}

Connect-MgGraph @AzureApp

# Download attachments from unread messages in the specified folder

$FilePath = "\\$($ENV:computername)\NetworkShare"   # The location to save the attachment(s) to
$MailboxUser = "aj@ajlindner.info"                  # The user mailbox to search
$FolderName = "Auto-Generated Report"               # The name of the folder that contains the messages to process

$MailFolder = Get-MgUserMailFolder -UserID $MailboxUser -Filter "displayname eq '$FolderName'"
$UnreadMessages = Get-MgUserMailFolderMessage -UserID $MailboxUser -MailFolderId $MailFolder.ID -Filter "isRead ne true"

ForEach ($Message in $UnreadMessages) {
    
    $Attachments = Get-MgUserMessageAttachment -UserID $MailboxUser -MessageId $Message.Id
    
    ForEach ($Attachment in $Attachments) {
        $FileName = $Attachment.Name    # The name (with extension) of the file to save. Defaults to the attachment name.
        $File = "$Filepath\$FileName"
        $Bytes = [Convert]::FromBase64String($Attachment.AdditionalProperties.ContentBytes)
        [IO.File]::WriteAllBytes($File, $Bytes)
    }
    
    Update-MgUserMessage -UserID $MailboxUser -MessageId $Message.ID -IsRead:$true
}