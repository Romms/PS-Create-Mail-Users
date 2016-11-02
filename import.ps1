#Requires -Version 2.0
. ".\scripts\Export-CSV -Append.ps1"
. ".\scripts\functions.ps1"

Add-Type -AssemblyName System.Web

$csvImportFilePath_default = 'people_import.csv'
$csvImportFilePathSuccess_default = 'people_import_success.csv'
$csvImportFilePathFailed_default = 'people_import_failed.csv'

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

$people = Import-PeopleFromCsv -Path $csvImportFilePath
$people_success = New-Object System.Collections.Generic.List[System.Object]
$people_failed = New-Object System.Collections.Generic.List[System.Object]

Write-Host "Processing..."

if ($true) {
    foreach( $user in $people) {
        $User.FirstName = Normalize-Name($User.FirstName).trim()
        $User.LastName = Normalize-Name($User.LastName).trim()

        if (StringIsNullOrWhitespace($User.Alias)) {
            $User.Alias = $User.FirstName + $User.LastName
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

Write-Host -NoNewLine 'Press Enter to exit... '
Read-Host