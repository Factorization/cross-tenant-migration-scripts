[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [String]
    $ExportPath,

    [Parameter(Mandatory = $true)]
    [string]
    $MappingFilePath
)
BEGIN {
    $Global:ErrorActionPreference = "Stop"

    try {
        Get-Command "Get-Mailbox" | Out-Null
    }
    catch {
        Write-Error "Not connected to Exchange Online. You must connect to Exchange Online first."
        exit
    }


    $File = Join-Path $exportPath "CSV_Exports\Export.csv"
    $CSV = Import-Csv $File
    $Mapping = Import-Csv $MappingFilePath

    $AllTargetMailboxes = Get-Mailbox -ResultSize Unlimited
}
PROCESS {

    foreach ($Line in $CSV) {
        $SourceEmail = $Line.PrimarySmtpAddress
        $TargetEmail = $Mapping | Where-Object { $_.Source_Email -eq $SourceEmail } | Select-Object -ExpandProperty Target_Email

        If (-not $TargetEmail) {
            Write-Host "Source email $SourceEmail not in Mapping file." -ForegroundColor Red
            Continue
        }

        $TargetMB = $AllTargetMailboxes | Where-Object { $_.PrimarySmtpAddress -eq $TargetEmail }
        if (-not $TargetMB) {
            Write-Host "Target email $TargetEmail not in Target EXO." -ForegroundColor Red
            Continue
        }

        $SourceProxyAddresses = @()
        $SourceProxyAddresses += "x500:$($Line.LegacyExchangeDN)"
        $SourceProxyAddresses += ($Line.EmailAddresses -split ";") | Where-Object { $_ -like "x500:*" } | ForEach-Object { $_ -creplace "^X500:", "x500:" }
        $SourceProxyAddresses = $SourceProxyAddresses | Sort-Object -Unique

        $TargetProxyAddresses = $TargetMB.EmailAddresses

        $MissingProxy = $false
        foreach ($Proxy in $SourceProxyAddresses) {
            if ($Proxy -notin $TargetProxyAddresses) {
                $MissingProxy = $true
                Write-Host
                Write-Host "Missing proxy:" -ForegroundColor Red
                write-host "`tSource email: $SourceEmail" -ForegroundColor Red
                Write-Host "`tTarget email: $TargetEmail" -ForegroundColor Red
                Write-Host "`tProxy: $Proxy" -ForegroundColor Red
            }
        }
        if (-not $MissingProxy) {
            Write-Host "Validated $SourceEmail | $TargetEmail" -ForegroundColor Green
        }
    }
}
