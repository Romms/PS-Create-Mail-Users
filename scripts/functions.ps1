function Translit-ToLat
{
    #������� ������������� ��������������� ���������� ������� ������
	# ������ �� 27 ���� 2010 �. N 55 (� ������).
    param(
        [string]$inputString
    )
    $translitUsual = @{
        [char]'�' = "a"
        [char]'�' = "A"
        [char]'�' = "b"
        [char]'�' = "B"
        [char]'�' = "v"
        [char]'�' = "V"
        [char]'�' = "h"
        [char]'�' = "H"
        [char]'�' = "g"
        [char]'�' = "G"
        [char]'�' = "d"
        [char]'�' = "D"
        [char]'�' = "e"
        [char]'�' = "E"
        [char]'�' = "ie"
        [char]'�' = "Ie"
        [char]'�' = "zh"
        [char]'�' = "Zh"
        [char]'�' = "z"
        [char]'�' = "Z"
        [char]'�' = "y"
        [char]'�' = "Y"
        [char]'�' = "i"
        [char]'�' = "I"
        [char]'�' = "i"
        [char]'�' = "I"
        [char]'�' = "i"
        [char]'�' = "I"
        [char]'�' = "k"
        [char]'�' = "K"
        [char]'�' = "l"
        [char]'�' = "L"
        [char]'�' = "m"
        [char]'�' = "M"
        [char]'�' = "n"
        [char]'�' = "N"
        [char]'�' = "o"
        [char]'�' = "O"
        [char]'�' = "p"
        [char]'�' = "P"
        [char]'�' = "r"
        [char]'�' = "R"
        [char]'�' = "s"
        [char]'�' = "S"
        [char]'�' = "t"
        [char]'�' = "T"
        [char]'�' = "u"
        [char]'�' = "U"
        [char]'�' = "f"
        [char]'�' = "F"
        [char]'�' = "kh"
        [char]'�' = "Kh"
        [char]'�' = "ts"
        [char]'�' = "Ts"
        [char]'�' = "ch"
        [char]'�' = "Ch"
        [char]'�' = "sh"
        [char]'�' = "Sh"
        [char]'�' = "shch"
        [char]'�' = "Shch"
        [char]'�' = "iu"
        [char]'�' = "Iu"
        [char]'�' = "ia"
        [char]'�' = "Ia"
    }

    $translitAtTheBegin = @{
        [char]'�' = "ye"
        [char]'�' = "Ye"
        [char]'�' = "yi"
        [char]'�' = "Yi"
        [char]'�' = "y"
        [char]'�' = "Y"
        [char]'�' = "yu"
        [char]'�' = "Yu"
        [char]'�' = "ya"
        [char]'�' = "Ya"

    }

    $translitRussian = @{
        [char]'�' = "e" 
        [char]'�' = "E"
        [char]'�' = "" # "``"
        [char]'�' = "" # "``"
        [char]'�' = "y" # "y`"
        [char]'�' = "Y" # "Y`"
        [char]'�' = "" # "`"
        [char]'�' = "" # "`"
        [char]'�' = "e" # "e`"
        [char]'�' = "e" # "E`"
    }

    $outChars = ""
    $begin = $true
    foreach ($c in $inChars = $inputString.ToCharArray())
    {

        $ch = $Null
        if ($begin) {
            if ($translitAtTheBegin[$c] -ne $Null ){
                $ch = $translitAtTheBegin[$c]
            } 
        }

        if ($ch -eq $Null) {
            if($translitUsual[$c] -ne $Null ) {
                $ch = $translitUsual[$c]
            } elseif ($translitRussian[$c] -ne $Null ) {
                $ch = $translitRussian[$c]
            } else {
                $ch = $c
            }
        }

        $outChars += $ch
        $begin = $false;
    }

    $outChars
}


function Convert-ToLatinCharacters {
    param(
        [string]$inputString
    )
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($inputString))
}

function Remove-NonAlphabeticCharacters {
    param(
        [string]$inputString
    )
    $inputString -replace '[^a-zA-Z0-9]', ''
}

function Normalize-Name {
    param(
        [string]$inputString
    )
    $inputString = Translit-ToLat($inputString)
    $inputString = Remove-NonAlphabeticCharacters($inputString)

    $inputString
}

function Create-Mailbox {
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$Year,
        [string]$Alias,
        [string]$Password
    )
    $PasswordEncrypted = ConvertTo-SecureString $Password -asplaintext -force

    $result = New-Mailbox `
        -Name $($FirstName+' '+$LastName) `
        -Alias $Alias `
        -OrganizationalUnit $('unicyb.kiev.ua/Students/01.09.' + $Year) `
        -UserPrincipalName $($Alias + '@unicyb.kiev.ua') `
        -SamAccountName $Alias `
        -FirstName $FirstName `
        -Initials '' `
        -LastName $LastName `
        -Password $PasswordEncrypted `
        -ResetPasswordOnNextLogon $false `
        -ErrorAction Stop
}

function Import-PeopleFromCsv {
    param (
        [string]$Path
    )
    Import-Csv -Path $Path -Delimiter ';'
}

function Export-PeopleToCsv {
    param (
        [parameter(ValueFromPipeline = $true)]
        $inputValue,

        [string]$Path
    )
    process {
        if ($firstRun -eq $Null) {
            $firstRun = $true
        } else {
            $firstRun = $false
        }

        if ($firstRun) {
            #Print headers
            ('' | select 'FirstName', 'LastName', 'Year', 'Alias', 'Password' | ConvertTo-Csv -NoType -Delimiter ';')[0] `
                | Out-File $Path -Encoding 'UTF8'
        }

        $inputValue `
            | select 'FirstName', 'LastName', 'Year', 'Alias', 'Password' `
            | Export-Csv -Append -Path $Path -Delimiter ';' -Encoding 'UTF8' -NoTypeInformation
    }
}

function StringIsNullOrWhitespace([string] $string)
{
    if ($string -ne $null) { $string = $string.Trim() }
    return [string]::IsNullOrEmpty($string)
}