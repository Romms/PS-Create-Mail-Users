#Requires -Version 2.0

param (
    [switch]$slientmode = $false
)

# Configs
$EMAIL_FROM = 'import@mail.server.com'
$EMAIL_TO = 'admin@server.com'
$EMAIL_SUBJECT = 'Import completed'
$EMAIL_SMTP = '10.10.1.1'
$USER_DOMAIN = 'server.com'

$csvImportFilePath_default = 'people_import.csv'
$csvImportFilePathSuccess_default = 'people_import_success.csv'
$csvImportFilePathFailed_default = 'people_import_failed.csv'

. ".\scripts\Export-CSV -Append.ps1"
. ".\scripts\functions.ps1"

. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
Connect-ExchangeServer -auto

Add-Type -AssemblyName System.Web

if ($slientmode) {
    $csvImportFilePath = $csvImportFilePath_default
    $csvImportFilePathSuccess = $csvImportFilePathSuccess_default
    $csvImportFilePathFailed = $csvImportFilePathFailed_default
    $cleanImportFile = $true
} else {
    $csvImportFilePath = Read-Host "Type csv import file destination [$csvImportFilePath_default]"
    $csvImportFilePathSuccess = Read-Host "Type csv success result file destination [$csvImportFilePathSuccess_default]"
    $csvImportFilePathFailed = Read-Host "Type csv failed result file destination [$csvImportFilePathFailed_default]"

    if (StringIsNullOrWhitespace($csvImportFilePath)) {
        $csvImportFilePath = $csvImportFilePath_default
    }
    if (StringIsNullOrWhitespace($csvImportFilePathSuccess)) {
        $csvImportFilePathSuccess = $csvImportFilePathSuccess_default
    }
    if (StringIsNullOrWhitespace($csvImportFilePathFailed)) {
        $csvImportFilePathFailed = $csvImportFilePathFailed_default
    }

    $cleanImportFile_input = Read-Host "Clear import file $csvImportFilePath after process end [Y]"
    $cleanImportFile = if ($cleanImportFile_input -eq 'N' -or $cleanImportFile_input -eq 'n') {
        $false
    } else {
        $true
    }
}

$people = Import-PeopleFromCsv -Path $csvImportFilePath
$people_success = New-Object System.Collections.Generic.List[System.Object]
$people_failed = New-Object System.Collections.Generic.List[System.Object]

Write-Host "Processing..."

if ($true) {
    $people | ForEach-Object {
        $User = $_
        $User.FirstName = Normalize-Name($User.FirstName).trim()
        $User.LastName = Normalize-Name($User.LastName).trim()

        if (StringIsNullOrWhitespace($User.Alias)) {
            $User.Alias = $($User.FirstName + '.' + $User.LastName).ToLower()
        }
        $User.Alias = Remove-NonAlphabeticCharacters($User.Alias).trim()
        
        if (StringIsNullOrWhitespace($User.Password)) {
            $User.Password = [System.Web.Security.Membership]::GeneratePassword(12,0)
        }
        $User.Password = $User.Password.trim();

        Write-Host "$($User.FirstName) $($User.LastName) - " -NoNewLine
        Try {
            Create-Mailbox -FirstName $($User.FirstName) -LastName $($User.LastName) -Year $($User.Year) -Alias $($User.Alias) -Password $($User.Password)
            Write-Host "Added"
            $people_success.add($User)
        }
        Catch {
            Write-Host "ERROR!" -foregroundcolor "red"
            Write-Host $($_.Exception.Message) -foregroundcolor "red"
            $people_failed.add($User)
        }

    }
}

Export-PeopleToCsv -inputValue $people_success -Path $csvImportFilePathSuccess
Export-PeopleToCsv -inputValue $people_failed -Path $csvImportFilePathFailed

if ($cleanImportFile -eq $true) {
    Export-PeopleToCsv -inputValue $(@()) -Path $csvImportFilePath
}

$people_success_emails = '';
$people_success | ForEach-Object {
    $people_success_emails += $_.Alias + "@" + $USER_DOMAIN + "`n";
}

Write-Host "Imported users:`n"
Write-Host $people_success_emails

Send-MailMessage `
    -from $EMAIL_FROM `
    -to $EMAIL_TO `
    -subject $EMAIL_SUBJECT `
    -body $("See attachments`n`n" + "Added users:`n`n" + $people_success_emails) `
    -Attachments $csvImportFilePathSuccess, $csvImportFilePathFailed `
    -smtpServer $EMAIL_SMTP

Write-Host "Email notification sent to $EMAIL_TO"
Write-Host -NoNewLine 'Press Enter to exit... '
Read-Host
