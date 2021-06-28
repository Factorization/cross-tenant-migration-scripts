[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $InputFile
)
BEGIN {
    If (-not (Test-Path $InputFile)) {
        Write-Host "File $InputFile does not exist. Exiting." -ForegroundColor Red
        exit
    }
    If (-not $InputFile.EndsWith('.xlsx')) {
        Write-Host "File $InputFile does not end with '.xlsx'. Exiting." -ForegroundColor Red
        exit
    }

    $User_Mailboxes = Import-Excel $InputFile -WorksheetName "User Mailboxes" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    $Shared_Mailboxes = Import-Excel $InputFile -WorksheetName "Shared Mailboxes" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
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

            $Source_Email = $User."Source Mailbox"
            $Target_Email = $User."Target Mailbox"

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
                Start-Sleep -Seconds 1
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

    Write-Verbose "Working on user mailboxes..."
    if ($User_Mailboxes) {
        SetForwardingAddresses -Mailboxes $User_Mailboxes
    }
    else {
        Write-Host "No user mailboxes in Excel file." -ForegroundColor Cyan
    }

    Write-Verbose "Working on shared mailboxes..."
    if ($Shared_Mailboxes) {
        SetForwardingAddresses -Mailboxes $Shared_Mailboxes
    }
    else {
        Write-Host "No shared mailboxes in Excel file." -ForegroundColor Cyan
    }

}
END {}
