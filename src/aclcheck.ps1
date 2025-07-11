function Get-AclPrincipalReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    $wellKnownSIDs = @(
        'S-1-5-11',        # Authenticated Users
        'S-1-5-18',        # SYSTEM
        'S-1-5-32-544',    # Administrators group
        'S-1-1-0'          # Everyone
    )

    $entries = Get-Acl -Path $Path |
        Select-Object -ExpandProperty Access |
        ForEach-Object {
            $access = $_
            $sid = $null
            $ntAccount = $null
            $displayName = $null
            $enabled = 'Unknown'

            try {
                $sid = $access.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier])
                $ntAccount = $sid.Translate([System.Security.Principal.NTAccount]).Value
                $displayName = $ntAccount
            } catch {
                $displayName = $access.IdentityReference.Value
            }

            if ($sid -and $wellKnownSIDs -contains $sid.Value) {
                $enabled = 'WellKnown'
            } else {
                try {
                    $user = Get-ADUser -Identity $sid.Value -Properties Enabled -ErrorAction Stop
                    if ($user) {
                        $enabled = if ($user.Enabled) { 'Enabled' } else { 'Disabled' }
                    }
                } catch {
                    try {
                        $group = Get-ADGroup -Identity $sid.Value -ErrorAction Stop
                        if ($group) { $enabled = 'Group' }
                    } catch {
                        $enabled = 'Unknown'
                    }
                }
            }

            [PSCustomObject]@{
                FileSystemRights = $access.FileSystemRights
                DisplayName      = $displayName
                Enabled          = $enabled
            }
        }

    return $entries
}
