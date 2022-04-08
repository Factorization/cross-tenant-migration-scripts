[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $InputFile
)
BEGIN {
    $global:ErrorActionPreference = "Stop"
    If (-not (Test-Path $InputFile)) {
        Write-Host "File $InputFile does not exist. Exiting." -ForegroundColor Red
        exit
    }
    If (-not $InputFile.EndsWith('.csv')) {
        Write-Host "File $InputFile does not end with '.csv'. Exiting." -ForegroundColor Red
        exit
    }

    $Mailboxes = Import-CSV $InputFile
    $DATE = Get-Date -Format yyyy-MM-dd_HH.mm
    $ErrorFile = "Email_Forwarding_Address_Errors_$DATE.csv"

    function SetForwardingAddresses($Mailboxes) {
        $Total = $Mailboxes | Measure-Object | Select-Object -ExpandProperty Count
        $i = 0
        $ErrorCount = 0
        $ErrorList = @()
        foreach ($User in $Mailboxes) {
            Write-Progress -Id 1 -Activity "Setting forwarding addresses..." -Status "Mailboxes: [$i/$Total] | Errors: $ErrorCount" -PercentComplete ($i / $Total * 100)
            $i++

            $Source_Email = $User.SourceEmail
            $Target_Email = $User.TargetEmail

            Try {
                $Mailbox = Get-Mailbox $Source_Email
                if (($Mailbox | Measure-Object | Select-Object -ExpandProperty Count) -ne 1) {
                    Throw "Invalid mailbox count."
                }
            }
            Catch {
                $err = $_
                $User | Add-Member -MemberType NoteProperty -Name Error -Value $err
                Write-Host "Error getting mailbox $Source_Email. Error: $err" -ForegroundColor Red
                $ErrorCount += 1
                $ErrorList += $User
                Continue
            }

            Try {
                # Set-Mailbox $Source_Email -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $Target_Email
                Write-Verbose "Set forwarding address for $Source_Email to $Target_Email"
            }
            Catch {
                $err = $_
                $User | Add-Member -MemberType NoteProperty -Name Error -Value $err
                Write-Host "Error getting mailbox $Source_Email. Error: $err" -ForegroundColor Red
                $ErrorCount += 1
                $ErrorList += $User
                Continue
            }
        }
        if ($ErrorList) {
            $ErrorList | Export-Csv -NoTypeInformation -Append -LiteralPath $ErrorFile
        }
    }

}
PROCESS {

    Write-Verbose "Working on mailboxes..."
    if ($Mailboxes) {
        SetForwardingAddresses -Mailboxes $Mailboxes
    }
    else {
        Write-Host "No mailboxes in CSV file." -ForegroundColor Cyan
    }

}
END {}
