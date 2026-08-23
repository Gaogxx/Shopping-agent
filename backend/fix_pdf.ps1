Write-Host "=== Finding remaining xiaolvjing traces ==="
$p = 'xiaolvjing|greenwhale|GreenWhale|xlj|XiaoLv'

# 1. Find ProgIDs
Write-Host "`n--- ProgIDs ---"
$roots = @("HKLM:\SOFTWARE\Classes", "HKCU:\Software\Classes")
foreach ($root in $roots) {
    Get-ChildItem $root -EA 0 | ForEach-Object {
        $n = $_.PSChildName
        if ($n -match $p -and $n -notmatch '^\.') {
            Write-Host "PROGID: $n"
            $cmd = "$($_.PSPath)\shell\open\command"
            if (Test-Path $cmd) {
                $v = (Get-ItemProperty $cmd -EA 0).'(default)'
                Write-Host "  cmd: $v"
            }
        }
    }
}

# 2. ALL OpenWithProgids entries
Write-Host "`n--- Checking all OpenWithProgids ---"
foreach ($root in $roots) {
    Get-ChildItem $root -EA 0 | ForEach-Object {
        $ow = "$($_.PSPath)\OpenWithProgids"
        if (Test-Path $ow) {
            $props = Get-ItemProperty $ow -EA 0
            foreach ($prop in $props.PSObject.Properties) {
                $n = $prop.Name
                $v = [string]$prop.Value
                if ($n -match $p) {
                    $extName = $_.PSChildName
                    Write-Host "FOUND: ${extName} -> $n"
                    Remove-ItemProperty -Path $ow -Name $n -Force -EA 0
                    Write-Host "  Deleted"
                }
            }
        }
    }
}

# 3. HKCR .pdf default
Write-Host "`n--- HKCR .pdf default ---"
$h = "HKLM:\SOFTWARE\Classes\.pdf"
if (Test-Path $h) {
    $def = (Get-ItemProperty $h -EA 0).'(default)'
    Write-Host "default = $def"
    if ($def -match $p) {
        Write-Host "FOUND in default!"
    }
}
$h2 = "HKCU:\Software\Classes\.pdf"
if (Test-Path $h2) {
    $def2 = (Get-ItemProperty $h2 -EA 0).'(default)'
    Write-Host "HKCU default = $def2"
}

# 4. Show current .pdf OpenWithList
Write-Host "`n--- Current .pdf OpenWithList ---"
$ol = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\OpenWithList"
if (Test-Path $ol) {
    $list = Get-ItemProperty $ol -EA 0
    foreach ($prop in $list.PSObject.Properties) {
        $n = $prop.Name
        if ($n -match '^[a-z]$') {
            Write-Host "  $n = $($prop.Value)"
        }
    }
}

Write-Host "`nDone."
