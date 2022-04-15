[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $PermissionsFile,

    [Parameter(Mandatory = $true)]
    [string]
    $MappingUsersFile,

    [Parameter(Mandatory = $true)]
    [string]
    $MappingSharedMailboxFile,

    [Parameter(Mandatory = $true)]
    [string]
    $OutputFile
)
BEGIN{
    $ErrorActionPreference = "Stop"
    $Mapping = @()
    $Mapping += Import-csv $MappingUsersFile
    $Mapping += Import-Csv $MappingSharedMailboxFile
    $MappingHash = $Mapping | Group-Object OldUPN -AsHashTable

    $CSV = Import-Csv $PermissionsFile
}
PROCESS{
    $results = @()
    foreach ($Line in $CSV){
        $OldEmail = $Line.Email
        $OldUserGrantedPermission = $Line.OldUserGrantedPermission
        $Permission = $line.Permission

        $NewEmail = $MappingHash[$OldEmail].NewUPN
        if(-not $NewEmail){
            Write-Host "Email '$OldEmail' not in mapping..." -ForegroundColor Cyan
            continue
        }

        $NewUserGrantedPermission = $MappingHash[$OldUserGrantedPermission].NewUPN
        if(-not $NewUserGrantedPermission){
            Write-Host "User Granted Permission '$OldUserGrantedPermission' not in mapping..." -ForegroundColor Cyan
            continue
        }
        $results += [PSCustomObject]@{
            Email = $OldEmail
            NewEmail = $NewEmail
            UserGrantedPermission = $OldUserGrantedPermission
            NewUserGrantedPermission = $NewUserGrantedPermission
            Permission = $Permission
        }
    }
}
END{
    if($results){
        $results | Export-Csv -NoTypeInformation -LiteralPath $OutputFile
        Write-Host "Results saved to: '$OutputFile'" -ForegroundColor Green
    }
}
