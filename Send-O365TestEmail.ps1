<#
    .NOTES
        AJ Lindner (AJ@AJLindner.info)
        July 16, 2022

    .SYNOPSIS
        Sends a test email from an Office 365 mailbox with an auto-generated attachment

    .DESCRIPTION
        This is a sample script for sending an email from an Office 365 mailbox with an automatically generated attachment
        for development/testing purposes. This script will output a .txt file with a random guid to the specified location
        and send that file as an attachment in an email.

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

# Generate a new test file
$Guid = (New-Guid).guid
$File = "$($ENV:USERPROFILE)\Attachment_$guid.txt"
"$Guid : This is an auto-generated attachment" | out-file $File

# Send the email with the test file attached
$Recipient = "aj@ajlindner.info"    # The recipient for the email
$MailboxUser = "aj@ajlindner.info"  # The mailbox to send the email from

$MGEmail = @{
    ToRecipients = @(
        @{
            emailAddress = @{
                address = $Recipient
            }
        }
    )
    Attachments = @(
        @{
            "@odata.type"= "#microsoft.graph.fileAttachment"
            Name = ((Get-Item -Path $File).name)
            ContentBytes = ( [Convert]::ToBase64String([IO.File]::ReadAllBytes($File)) )
        }
    )
    Subject = "Auto-Generated Report"
    Body = @{
        contentType = "html"
        content = "Please see the attached Report."
    }
}

Send-MgUserMail -UserID $MailboxUser -Message $MGEmail