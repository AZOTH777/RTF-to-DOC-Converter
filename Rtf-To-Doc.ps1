 param (
    [Parameter(Mandatory=$true)][string]$Path
 )

Function Convert-Dir($Path){

    $Files=Get-ChildItem $Path -Filter *.rtf -Recurse | % { $_.FullName }
    $Word=New-Object –ComObject WORD.APPLICATION

    foreach ($File in $Files) {
        if (($File -ne $null) -and ($Word -ne $null)){
            $Doc=$Word.Documents.Open($File)
            $Name=($Doc.Fullname).replace(".rtf",".doc")

            if (Test-Path $Name){

            } else {
                Write-Host $Name
                $Doc.saveas([ref] $Name, [ref] 0)  
                $Doc.close()
            }
        }
    }
}

Convert-Dir $Path;