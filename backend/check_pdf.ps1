# Check PDF file association
Write-Host "=== UserChoice ==="
Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice" -EA SilentlyContinue | fl ProgId

Write-Host "=== HKCR .pdf ==="
Get-ItemProperty "HKLM:\SOFTWARE\Classes\.pdf" -EA SilentlyContinue | fl "(default)"

Write-Host "=== OpenWithList ==="
$key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\OpenWithList"
if (Test-Path $key) {
    $props = Get-ItemProperty $key
    $props.PSObject.Properties | Where-Object { $_.Name -match '^[a-z]$' } | ForEach-Object { Write-Host "$($_.Name) = $($_.Value)" }
}

Write-Host "=== OpenWithProgids ==="
$key2 = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\OpenWithProgids"
if (Test-Path $key2) {
    $props2 = Get-ItemProperty $key2
    $props2.PSObject.Properties | Where-Object { $_.Name -ne 'PSPath' -and $_.Name -ne 'PSParentPath' -and $_.Name -ne 'PSChildName' -and $_.Name -ne 'PSDrive' -and $_.Name -ne 'PSProvider' } | ForEach-Object { Write-Host "$($_.Name) = $($_.Value)" }
}

Write-Host "=== Searching for xiaolvjing in registry ==="
$results = Get-ChildItem "HKCU:\Software\Classes" -Recurse -EA SilentlyContinue | Where-Object { $_.PSPath -match 'xiaolvjing|小绿鲸|greenwhale|GreenWhale|xlsreader' } | Select-Object -First 10
if ($results) { $results | ForEach-Object { Write-Host $_.PSPath } } else { Write-Host "Not found in HKCU Classes" }
