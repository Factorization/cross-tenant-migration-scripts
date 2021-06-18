[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]
    $OldFile,

    [Parameter(Mandatory=$true)]
    [string]
    $NewFile
)
BEGIN{}
PROCESS{
    $Mailbox_Types = @(
        "User Mailboxes",
        "Shared Mailboxes",
        "Room Mailboxes",
        "Equipment Mailboxes"
    )

    $Old = @()
    $New = @()
    foreach ($T in $Mailbox_Types){
        Try{
            $temp_old = (Import-Excel $OldFile -WorksheetName $T -erroraction Stop).'Source Mailbox'
        }
        catch{
            $temp_old = @()
        }
        Try{
            $temp_new = (Import-Excel $NewFile -WorksheetName $T -erroraction Stop).'Source Mailbox'
        }
        catch{
            $temp_new = @()
        }
        $old += $temp_old
        $New += $temp_new
    }
    return $temp_new
}
