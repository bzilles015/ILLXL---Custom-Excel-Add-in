# ILLXL installer
#   pwsh -File install.ps1                    install / update the Nexcess build
#   pwsh -File install.ps1 -Edition Public    install / update the public build
#   pwsh -File install.ps1 -Uninstall         stop loading ILLXL (files are kept)
#
# Copies the built add-in to Excel's own add-ins folder
#   %APPDATA%\Microsoft\AddIns\
# and tells Excel to load it from there at startup. Excel then never runs the
# copy in build\, so rebuilding can't touch the add-in you're working with.
# The previous installed copy is kept as <name>.xlam.bak.
param(
    [ValidateSet('Nexcess', 'Public')] [string]$Edition = 'Nexcess',
    [switch]$Uninstall
)
$ErrorActionPreference = 'Stop'
$name = @{ Nexcess = 'ILLXL - Nexcess.xlam'; Public = 'ILLXL.xlam' }[$Edition]
$src = Join-Path $PSScriptRoot "build\$name"
$dir = Join-Path $env:APPDATA 'Microsoft\AddIns'
$dst = Join-Path $dir $name
$reg = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "Close Excel first - it rewrites its add-in list when it exits, which would undo this."
}

# Excel's startup add-ins are the values OPEN, OPEN1, OPEN2 ... under $reg
$opens = (Get-Item $reg).Property | Where-Object { $_ -match '^OPEN\d*$' }
$illxl = $opens | Where-Object { (Get-ItemPropertyValue $reg $_) -match 'ILLXL[^\\]*\.xlam' }

if ($Uninstall) {
    foreach ($o in $illxl) { $was = Get-ItemPropertyValue $reg $o; Remove-ItemProperty $reg $o; "Excel will no longer load: $was" }
    "Done. The add-in files were left in place."
    return
}

if (-not (Test-Path $src)) { throw "Build it first:  pwsh -File build.ps1 -Edition $Edition" }
New-Item -ItemType Directory -Force $dir | Out-Null
if (Test-Path $dst) { Copy-Item $dst "$dst.bak" -Force; "Backed up the previous copy to $dst.bak" }
Copy-Item $src $dst -Force
Unblock-File $dst                                   # no "downloaded from the internet" block

# Point Excel at the installed copy: reuse an existing ILLXL entry, else the next free OPENn
$value = '"' + $dst + '"'
if ($illxl) {
    Set-ItemProperty $reg ($illxl | Select-Object -First 1) $value
    foreach ($extra in ($illxl | Select-Object -Skip 1)) { Remove-ItemProperty $reg $extra }   # never load two copies
} else {
    $n = 0; while ($opens -contains ($(if ($n) { "OPEN$n" } else { 'OPEN' }))) { $n++ }
    New-ItemProperty $reg ($(if ($n) { "OPEN$n" } else { 'OPEN' })) -Value $value -PropertyType String | Out-Null
}
"Installed: $dst"
"Excel will load it the next time it starts."
