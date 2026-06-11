param(
    [string[]]$Files
)

$Path = Join-Path (Get-Location) "source"
$sourceEncoding = New-Object System.Text.UTF8Encoding($false, $true)
$targetEncoding = [System.Text.Encoding]::GetEncoding(1251)
$extensions = @('.txt', '.csv', '.md', '.bas', '.cls', '.frm', '.doccls', '.vbs', '.ps1')

function Test-IsUtf8Encoded {
    param(
        [byte[]]$Bytes
    )

    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
        $text = $utf8.GetString($Bytes)
        $roundTripBytes = $utf8.GetBytes($text)
        return [System.Linq.Enumerable]::SequenceEqual($Bytes, $roundTripBytes)
    }
    catch {
        return $false
    }
}

$getChildItemParams = @{
    Path = $Path
    File = $true
}

$filesToProcess = @()

if ($Files -and $Files.Count -gt 0) {
    foreach ($file in $Files) {
        $candidatePath = if ([System.IO.Path]::IsPathRooted($file)) { $file } else { Join-Path $Path $file }

        if (-not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
            Write-Warning "Skipped, file not found: $candidatePath"
            continue
        }

        $item = Get-Item -LiteralPath $candidatePath
        if ($item.Extension.ToLowerInvariant() -notin $extensions) {
            Write-Warning "Skipped, unsupported extension: $($item.FullName)"
            continue
        }

        $filesToProcess += $item
    }
}
else {
    $filesToProcess = Get-ChildItem @getChildItemParams | Where-Object {
        $_.Extension.ToLowerInvariant() -in $extensions
    }
}

$filesToProcess | ForEach-Object {
    $filePath = $_.FullName
    $bytes = [System.IO.File]::ReadAllBytes($filePath)

    if (-not (Test-IsUtf8Encoded -Bytes $bytes)) {
        Write-Host "Skipped, already CP1251 or not valid UTF-8: $filePath"
        return
    }

    $content = [System.IO.File]::ReadAllText($filePath, $sourceEncoding)
    [System.IO.File]::WriteAllText($filePath, $content, $targetEncoding)
    Write-Host "Converted from UTF-8: $filePath"
}