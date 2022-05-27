[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [String]
    $SharePointAdminUrl,

    [Parameter(Mandatory = $false)]
    [switch]
    $IncludeMailboxStats,

    [Parameter(Mandatory = $false)]
    [string]
    $Path = "M365_Assessment_$(Get-Date -Format yyyy.MM.dd_HH.mm).xlsx"
)
BEGIN {
    $ErrorActionPreference = "Stop"
    Connect-ExchangeOnline -ShowBanner:$false | Out-Null
    Connect-SPOService -Url $SharePointAdminUrl | Out-Null
    Connect-MicrosoftTeams | Out-Null
}
Process {
    # Email
    Write-Host "Getting mailboxes..." -NoNewline
    $Mailboxes = Get-Mailbox -ResultSize Unlimited
    Write-Host "DONE" -ForegroundColor Green
    Write-Host "Saving mailbox info to Excel file..." -NoNewline
    $Mailboxes | Where-Object RecipientTypeDetails -eq UserMailbox | `
        Select-Object DisplayName, PrimarySmtpAddress, AccountDisabled | `
        Sort-Object DisplayName | `
        Export-Excel -WorksheetName "UserMailboxes" -Path $Path -AutoSize -AutoFilter -FreezeTopRow
    $Mailboxes | Where-Object RecipientTypeDetails -eq SharedMailbox | `
        Select-Object DisplayName, PrimarySmtpAddress | `
        Sort-Object DisplayName | `
        Export-Excel -WorksheetName "SharedMailboxes" -Path $Path -AutoSize -AutoFilter -FreezeTopRow
    $Mailboxes | Where-Object RecipientTypeDetails -eq RoomMailbox | `
        Select-Object DisplayName, PrimarySmtpAddress | `
        Sort-Object DisplayName | `
        Export-Excel -WorksheetName "RoomMailboxes" -Path $Path -AutoSize -AutoFilter -FreezeTopRow
    $Mailboxes | Where-Object RecipientTypeDetails -eq EquipmentMailbox | `
        Select-Object DisplayName, PrimarySmtpAddress | `
        Sort-Object DisplayName | `
        Export-Excel -WorksheetName "EquipmentMailboxes" -Path $Path -AutoSize -AutoFilter -FreezeTopRow
    $Mailboxes | Where-Object RecipientTypeDetails -eq SchedulingMailbox | `
        Select-Object DisplayName, PrimarySmtpAddress | `
        Sort-Object DisplayName | `
        Export-Excel -WorksheetName "SchedulingMailboxes" -Path $Path -AutoSize -AutoFilter -FreezeTopRow
    Write-Host "DONE" -ForegroundColor Green
    if ($IncludeMailboxStats) {
        Write-Host "Getting mailbox statistics..." -NoNewline
        $Mailbox_Stats = $Mailboxes | ForEach-Object {
            Get-EXOMailboxStatistics $_.PrimarySmtpAddress -PropertySets All
        }
        Write-Host "DONE" -ForegroundColor Green
        Write-Host "Saving mailbox statistics to Excel file..." -NoNewline
        $Mailbox_Stats | Select-Object DisplayName, ItemCount, @{n = "SizeGB"; e = { [math]::Round( ((($_.TotalItemSize.Value -split " ")[2].TrimStart("(") -replace ",") / 1GB), 2 ) } }, LastLogonTime | `
            Sort-Object DisplayName | `
            Export-Excel -WorksheetName "MailboxStats" -Path $Path -FreezeTopRow -AutoSize -AutoFilter
        Write-Host "DONE" -ForegroundColor Green
    }

    # SharePoint
    Write-Host "Getting SharePoint sites..." -NoNewline
    $Sites = Get-SPOSite -IncludePersonalSite $True -Limit all
    $OneDrive_Sites = $Sites | Where-Object { $_.Url -like '*-my.sharepoint.com/personal/*' }
    $SharePoint_Sites = $Sites | Where-Object { $_.Url -notlike '*-my.sharepoint.com/personal/*' }
    Write-Host "DONE" -ForegroundColor Green
    Write-Host "Saving SharePoint site data to Excel file..." -NoNewline
    $SharePoint_Sites | Where-Object { $_.Title } | `
        Select-Object Title, @{n = 'SizeMb'; e = { $_.StorageUsageCurrent } }, Template, @{n = 'MemberCount'; e = { if ($_.GroupId.guid -eq '00000000-0000-0000-0000-000000000000') { "n/a" }else { Get-UnifiedGroupLinks -Identity $_.GroupId.guid -LinkType member | Measure-Object | Select-Object -ExpandProperty count } } }, @{n = 'Owner'; e = { if ($_.GroupId.guid -eq '00000000-0000-0000-0000-000000000000') { $_.Owner }else { (Get-UnifiedGroupLinks -Identity $_.GroupId.guid -LinkType Owner).Name -join " | " } } }, IsTeamsConnected, LockState | `
        Sort-Object Title | `
        Export-Excel -WorksheetName "SharePoint" -Path $Path -FreezeTopRow -AutoFilter -AutoSize
    $OneDrive_Sites | Select-Object Title, @{n = 'SizeMb'; e = { $_.StorageUsageCurrent } }, Owner, LockState | `
        Sort-Object Title | `
        Export-Excel -WorksheetName "OneDrive" -Path $Path -FreezeTopRow -AutoFilter -AutoSize
    Write-Host "DONE" -ForegroundColor Green

    # Teams
    Write-Host "Getting Teams sites..." -NoNewline
    $Teams = Get-Team
    $Teams_Stats = $Teams | Select-Object DisplayName, @{n = "Owner"; e = { (Get-TeamUser -GroupId $_.GroupId -Role Owner).Name -join " | " } }, @{n = 'MemberCount'; e = { Get-TeamUser -GroupId $_.GroupId -Role Member | Measure-Object | Select-Object -ExpandProperty count } }
    Write-Host "DONE" -ForegroundColor Green
    Write-Host "Saving Teams data to Excel file..." -NoNewline
    $Teams_stats | Select-Object DisplayName, MemberCount, Owner | `
        Export-Excel -WorksheetName "Teams" -Path $Path -AutoSize -FreezeTopRow -AutoFilter
    Write-Host "DONE" -ForegroundColor Green
}
END {
    Write-Host
    Write-Host "Excel report saved to: " -NoNewline
    Write-Host $Path -ForegroundColor Green

    $Summary = @"
Mailboxes           = $($Mailboxes | Where-Object {$_.RecipientTypeDetails -match 'UserMailbox|SharedMailbox|RoomMailbox|EquipmentMailbox'} | Measure-Object | Select-Object -ExpandProperty Count)
SharePoint Sites    = $($SharePoint_Sites | Measure-Object | Select-Object -ExpandProperty Count)
OneDrive Sites      = $($OneDrive_Sites | Measure-Object | Select-Object -ExpandProperty Count)
Teams Sites         = $($Teams | Measure-Object | Select-Object -ExpandProperty Count)
"@
    Write-Host
    Write-Host "Results Summary"
    Write-Host "---------------"
    Write-Host $Summary
    Write-Host

    Disconnect-ExchangeOnline -Confirm:$false *>&1 | Out-Null
    Disconnect-SPOService | Out-Null
    Disconnect-MicrosoftTeams -Confirm:$false | Out-Null
}
