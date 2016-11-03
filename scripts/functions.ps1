function Translit-ToLat
{
    #Правила транслітерації регламентуються Постановою Кабінету міністрів
    # України від 27 січня 2010 р. N 55 (зі змінами).
    param(
        [string]$inputString
    )
    $translitUsual = @{
        [char]'а' = "a"
        [char]'А' = "A"
        [char]'б' = "b"
        [char]'Б' = "B"
        [char]'в' = "v"
        [char]'В' = "V"
        [char]'г' = "h"
        [char]'Г' = "H"
        [char]'ґ' = "g"
        [char]'Ґ' = "G"
        [char]'д' = "d"
        [char]'Д' = "D"
        [char]'е' = "e"
        [char]'Е' = "E"
        [char]'є' = "ie"
        [char]'Є' = "Ie"
        [char]'ж' = "zh"
        [char]'Ж' = "Zh"
        [char]'з' = "z"
        [char]'З' = "Z"
        [char]'и' = "y"
        [char]'И' = "Y"
        [char]'і' = "i"
        [char]'І' = "I"
        [char]'ї' = "i"
        [char]'Ї' = "I"
        [char]'й' = "i"
        [char]'Й' = "I"
        [char]'к' = "k"
        [char]'К' = "K"
        [char]'л' = "l"
        [char]'Л' = "L"
        [char]'м' = "m"
        [char]'М' = "M"
        [char]'н' = "n"
        [char]'Н' = "N"
        [char]'о' = "o"
        [char]'О' = "O"
        [char]'п' = "p"
        [char]'П' = "P"
        [char]'р' = "r"
        [char]'Р' = "R"
        [char]'с' = "s"
        [char]'С' = "S"
        [char]'т' = "t"
        [char]'Т' = "T"
        [char]'у' = "u"
        [char]'У' = "U"
        [char]'ф' = "f"
        [char]'Ф' = "F"
        [char]'х' = "kh"
        [char]'Х' = "Kh"
        [char]'ц' = "ts"
        [char]'Ц' = "Ts"
        [char]'ч' = "ch"
        [char]'Ч' = "Ch"
        [char]'ш' = "sh"
        [char]'Ш' = "Sh"
        [char]'щ' = "shch"
        [char]'Щ' = "Shch"
        [char]'ю' = "iu"
        [char]'Ю' = "Iu"
        [char]'я' = "ia"
        [char]'Я' = "Ia"
    }

    $translitAtTheBegin = @{
        [char]'є' = "ye"
        [char]'Є' = "Ye"
        [char]'ї' = "yi"
        [char]'Ї' = "Yi"
        [char]'й' = "y"
        [char]'Й' = "Y"
        [char]'ю' = "yu"
        [char]'Ю' = "Yu"
        [char]'я' = "ya"
        [char]'Я' = "Ya"

    }

    $translitRussian = @{
        [char]'ё' = "e" 
        [char]'Ё' = "E"
        [char]'ъ' = "" # "``"
        [char]'Ъ' = "" # "``"
        [char]'ы' = "y" # "y`"
        [char]'Ы' = "Y" # "Y`"
        [char]'ь' = "" # "`"
        [char]'Ь' = "" # "`"
        [char]'э' = "e" # "e`"
        [char]'Э' = "e" # "E`"
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
    $inputString -replace '[^a-zA-Z0-9\.\-]', ''
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
    Import-Csv -Path $Path -Delimiter ','
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
            ('' | select 'FirstName', 'LastName', 'Year', 'Alias', 'Password' | ConvertTo-Csv -NoType -Delimiter ',')[0] `
                | Out-File $Path -Encoding 'UTF8'
        }

        $inputValue `
            | select 'FirstName', 'LastName', 'Year', 'Alias', 'Password' `
            | Export-Csv -Append -Path $Path -Delimiter ',' -Encoding 'UTF8' -NoTypeInformation
    }
}

function StringIsNullOrWhitespace([string] $string)
{
    if ($string -ne $null) { $string = $string.Trim() }
    return [string]::IsNullOrEmpty($string)
}