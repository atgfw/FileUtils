function Get-InheritanceBreaks {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    Add-Type -AssemblyName System.Core

    $i = 0
    $j = 0
    Write-Host "Scanning folders under: $Path"

    $toplevel = [System.IO.Directory]::EnumerateDirectories($Path)
    $jmax = ($toplevel | Measure-Object).count
    foreach ($top in $toplevel) {
        $files = [System.IO.Directory]::EnumerateDirectories($top,"*","AllDirectories")
        $acl = Get-ACL -Path $top
        if ($acl.AreAccessRulesProtected) {
            Write-Output $top
        }
        $j++
        foreach ($file in $files) {
            $i++
            $acl = Get-ACL -Path $file
            if ($acl.AreAccessRulesProtected) {
                Write-Output $file
            }
            if ($i % 300 -eq 0) {
                Write-Progress -Activity "Scanning ACLs" -Status $file -PercentComplete (($j/$jmax) * 100)
            }
        }
    }
}