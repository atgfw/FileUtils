function Get-ADUserAccessChk {
    param (
        [Parameter(Mandatory)]
        [string[]]$Directories,

        [Parameter(Mandatory)]
        [Microsoft.ActiveDirectory.Management.ADUser[]]$Users
    )

    # Validate directories
    $ValidDirs = $Directories | Where-Object { Test-Path $_ }
    if ($ValidDirs.Count -eq 0) {
        throw "No valid directories provided."
    }

    $Results = @()
    $total = $Users.Count
    $i = 0

    foreach ($user in $Users) {
        $i++
        Write-Progress -Activity "Scanning Permissions" -Status "Scanning access for $($user.Name)" -PercentComplete (($i / $total) * 100)

        $result = [ordered]@{ Name = $user.Name }

        foreach ($dir in $ValidDirs) {
            try {
                $output = $null
                if (Test-Path $dir -PathType Leaf) {
                    $output = accesschk64.exe $user.SamAccountName $dir -nobanner 2>&1
                }
                else {
                    $output = accesschk64.exe $user.SamAccountName $dir -nobanner -d 2>&1
                }

                if ($output -match "No matching objects found.") {
                    $result[$dir] = "Error"
                    continue
                }

                $read  = ($output[0] -eq "R")
                $write = ($output[1] -eq "W")
                $result[$dir] = if ($read -or $write) { $true } else { $false }
            }
            catch {
                $result[$dir] = "Error"
            }
        }

        $Results += New-Object PSObject -Property $result
    }
    Write-Progress -Activity "Scanning Permissions" -Completed
    
    return $Results
}