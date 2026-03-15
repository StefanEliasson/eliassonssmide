# --- 1. Hitta index.htm ---
$indexFile = Join-Path -Path "." -ChildPath "index.htm"

if (Test-Path $indexFile) {
    Write-Host "Fixar index.htm..."

    # Läs filen i ANSI (Excel-format)
    $content = Get-Content $indexFile -Encoding Default

    # Ta bort gamla charset-rader
    $content = $content -replace 'charset=windows-1252', 'charset=UTF-8'
    $content = $content -replace 'charset=iso-8859-1', 'charset=UTF-8'

    # Om ingen charset finns, lägg till en ren rad direkt efter <head>
    if ($content -notmatch 'charset=UTF-8') {
        $content = $content -replace '<head>', '<head>' + "`r`n" + '<meta charset="UTF-8">'
    }

    # Spara som UTF-8
    Set-Content -Path $indexFile -Value $content -Encoding UTF8

    # Döp om index.htm → index.html
    Rename-Item -Path $indexFile -NewName "index.html" -Force
    Write-Host "index.htm döptes om till index.html"
}

# --- 2. Hitta alla .htm i index-filer ---
$sheetFolder = Join-Path -Path "." -ChildPath "index_files"

if (Test-Path $sheetFolder) {
    $sheetFiles = Get-ChildItem -Path $sheetFolder -Filter *.htm

    foreach ($file in $sheetFiles) {
        Write-Host "Fixar $($file.Name)..."

        # Läs i ANSI
        $content = Get-Content $file.FullName -Encoding Default

        # Byt charset
        $content = $content -replace 'charset=windows-1252', 'charset=UTF-8'
        $content = $content -replace 'charset=iso-8859-1', 'charset=UTF-8'

        # Lägg till charset om det saknas
        if ($content -notmatch 'charset=UTF-8') {
            $content = $content -replace '<head>', '<head>' + "`r`n" + '<meta charset="UTF-8">'
        }

        # Spara som UTF-8
        Set-Content -Path $file.FullName -Value $content -Encoding UTF8
    }
}

Write-Host "Klart! Alla filer är nu UTF-8 och index.htm är omdöpt."
