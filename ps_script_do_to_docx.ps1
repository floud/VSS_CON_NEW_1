$inputDir = "C:\path\to\doc\files"
$outputDir = "C:\path\to\docx\files"

# Создание выходной директории, если она не существует
if (-Not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Force -Path $outputDir
}

# Рекурсивный поиск и конвертация файлов
Get-ChildItem -Path $inputDir -Recurse -Filter *.doc | ForEach-Object {
    $docFile = $_.FullName
    $relativePath = $_.FullName.Substring($inputDir.Length) -replace '^\\', ''
    $outputFile = Join-Path $outputDir ($relativePath -replace '\.doc$', '.docx')

    # Создание директорий для выходных файлов, если они не существуют
    $outputDirPath = [System.IO.Path]::GetDirectoryName($outputFile)
    if (-Not (Test-Path -Path $outputDirPath)) {
        New-Item -ItemType Directory -Force -Path $outputDirPath
    }

    # Конвертация с помощью LibreOffice
    Start-Process -FilePath "soffice.exe" -ArgumentList "--headless", "--convert-to", "docx", "--outdir", $outputDirPath, $docFile -NoNewWindow -Wait

    Write-Host "Converted $relativePath to ${relativePath -replace '\.doc$', '.docx'}"
}