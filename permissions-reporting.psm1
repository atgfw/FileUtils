$Script:toolsPath = Join-Path -Path $PSScriptRoot -ChildPath "dependencies"

if (Test-Path -PathType Container $toolsPath) {
    Write-Verbose "Dependencies folder ($toolsPath) already exists"
}
elseif (Test-Path -PathType Leaf $toolsPath) {
    throw "Dependencies folder ($toolsPath) is a file!"
}
else {
    Write-Verbose "Creating dependencies folder at $toolsPath"
    New-Item -ItemType Directory -Path $toolsPath -ErrorAction Stop
}

Invoke-WebRequest -Uri "https://live.sysinternals.com/accesschk64.exe" -OutFile $toolsPath
Invoke-WebRequest -Uri "https://live.sysinternals.com/AccessEnum.exe" -OutFile $toolsPath

Get-ChildItem -Path $PSScriptRoot\src\*.ps1 | ForEach-Object {
    . $_.FullName
    Export-ModuleMember -Function ([System.IO.Path]::GetFileNameWithoutExtension($_.Name))
}