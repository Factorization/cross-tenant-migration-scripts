<#
 .Synopsis
	Exports object information for cross tenant migrations

 .Description
	This script exports object information from AD, EXO and AzureAD. It can be used to recreate users in a target domain for a cross tenant migration.

 .Parameter GroupName
	The name of the group that contains all the objects that you want to export.

 .Parameter DomainName
	The AD domain name. This is used for the Get-ADUser command. This can also be a specific Domain controller.

 .Example
	.\CrossTenantObjectExport.ps1 -GroupName "Migration Users" -DomainName ad.domain

	This runs the basic user export for the users that are a member of the group "Migration Users" .

 .Example
	.\CrossTenantObjectExport.ps1 -GroupName "Migration Users" -DomainName ad.domain -IncludeInboxRules

	This runs the basic export and includes the inbox rules for each user (This requires Exchange Admin)

 .Example
	.\CrossTenantObjectExport.ps1 -GroupName "Migration Users" -DomainName ad.domain -IncludeMailboxStats

	This runs the basic export and includes the mailbox and folder statistics for each user (This requires Exchange Admin)

 .Example
	.\CrossTenantObjectExport.ps1 -GroupName "Migration Users" -DomainName ad.domain -IncludeMailboxPermissions

	This runs the basic export and includes the mailbox permissions for each user (This requires Exchange Admin)

 .Example
	.\CrossTenantObjectExport.ps1 -GroupName "Migration Users" -DomainName ad.domain -IncludeInboxRules -IncludeMailboxStats -IncludeMailboxPermissions

	This runs the complete export (This requires Exchange Admin)

 .Notes

	#######################################################
	#  .                                               .  #
	#  .                Written By:                    .  #
	#.....................................................#
	#  .              Jeffrey Kraemer                  .  #
	#  .                  ENS, Inc.                    .  #
	#  .            jkraemer@ens-inc.com               .  #
	#.....................................................#
	#  .                                               .  #
	#######################################################
#>

#Requires -Modules ExchangeOnlineManagement, AzureAD, ActiveDirectory, MSOnline

[CmdletBinding()]
Param(
	# Cross tenant group name
	[Parameter(Mandatory = $True)]
	[string]
	$GroupName,

	# AD domain
	[Parameter(Mandatory = $True)]
	[String]
	$DomainName,

	# Output root folder
	[Parameter(Mandatory = $false)]
	[String]
	$Path = $PSScriptRoot,

	# Include mailbox stats
	[Parameter(Mandatory = $False)]
	[switch]
	$IncludeMailboxStats,

	# Include inbox rules
	[Parameter(Mandatory = $false)]
	[switch]
	$IncludeInboxRules,

	# Include mailbox permissions
	[Parameter(Mandatory = $false)]
	[switch]
	$IncludeMailboxPermissions
)
BEGIN {
	$Global:ErrorActionPreference = 'Stop'

	#############
	# Functions #
	#############
	Function CreateFolder($Path) {
		New-Item -Path $Path -ItemType Directory -Force | Out-Null
	}
	Function ExportXML($Object, $Path, $Email) {
		$Email = $Email + ".xml"
		$Path = Join-Path $Path $Email
		$Object | Export-Clixml -Path $Path | Out-Null
	}
	Function TestRequiredModule() {
		try {
			Get-Command Get-Mailbox | Out-Null
		}
		Catch {
			Write-Host "Not connected to Exchange Online. Please run Connect-ExchangeOnline before running this script." -ForegroundColor Red
			Exit
		}
		Try {
			Get-AzureADTenantDetail | Out-Null
		}
		catch {
			Write-Host "Not connected to Azure AD. Please run Connect-AzureAD before running this script." -ForegroundColor Red
			Exit
		}
		Try {
			Get-MsolAccountSku | Out-Null
		}
		Catch {
			Write-Host "Not connected to MSOnline Service. Please run Connect-MsolService before running this script." -ForegroundColor Red
		}
		Try {
			Get-Command Get-AdUser | Out-Null
		}
		catch {
			Write-Host "Missing AD PowerShell module. Install AD PowerShell module from RSAT before running this script." -ForegroundColor Red
			Exit
		}
	}

	Function ExportUserInfo($Email) {
		$SourceEmailAddress = $Email

		Write-Verbose "Getting user info..."
		$User = Get-User $SourceEmailAddress
		Write-Verbose "Exporting user info to XML.."
		ExportXML -Object $User -Path $GetUserOutput -Email $SourceEmailAddress

		Write-Verbose "Getting recipient info..."
		$Recipient = Get-EXORecipient $SourceEmailAddress -PropertySets All
		Write-Verbose "Exporting recipient info to XML..."
		ExportXML -Object $Recipient -Path $GetRecipientOutput -Email $SourceEmailAddress

		Write-Verbose "Getting mailbox info..."
		$Mailbox = Get-EXOMailbox $SourceEmailAddress -PropertySets All
		Write-Verbose "Exporting mailbox info to XML..."
		ExportXML -Object $Mailbox -Path $GetMailboxOutput -Email $SourceEmailAddress

		if ($IncludeMailboxStats) {
			Write-Verbose "Getting mailbox statistics..."
			$Statistics = Get-EXOMailboxStatistics $SourceEmailAddress -PropertySets All
			Write-Verbose "Exporting mailbox statistics to XML..."
			ExportXML -Object $Statistics -Path $GetMailboxStatisticsOutput -Email $SourceEmailAddress
		}

		Write-Verbose "Getting Azure AD user info..."
		$AzureADUser = Get-AzureADUser -ObjectId $Mailbox.UserPrincipalName
		Write-Verbose "Exporting Azure AD user info to XML..."
		ExportXML -Object $AzureADUser -Path $GetAzureADUserOutput -Email $SourceEmailAddress

		Write-Verbose "Getting MSOL user info..."
		$MSOLUser = Get-MsolUser -UserPrincipalName $Mailbox.UserPrincipalName
		Write-Verbose "Exporting Msol user info to XML..."
		ExportXML -Object $MSOLUser -Path $GetMsolUserOutput -Email $SourceEmailAddress

		if ($mailbox.IsDirSynced) {
			Write-Verbose "Getting AD user info..."
			$ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$($Mailbox.UserPrincipalName)'" -Server $DomainName -Properties *
			Write-Verbose "Exporting AD user info to XML..."
			ExportXML -Object $ADUser -Path $GetADUserOutput -Email $SourceEmailAddress
		}
		else {
			Write-Verbose "User is not synced from AD. Skipping AD user export..."
			$ADUser = $null
		}

		if ($IncludeMailboxPermissions) {
			Write-Verbose "Getting recipient permissions..."
			$RecipientPermission = Get-EXORecipientPermission $SourceEmailAddress
			Write-Verbose "Exporting recipient permissions to XML..."
			ExportXML -Object $RecipientPermission -Path $GetRecipientPermissionOutput -Email $SourceEmailAddress
		}

		if ($IncludeMailboxPermissions) {
			Write-Verbose "Getting mailbox folder permissions..."
			$MailboxFolderPermission = Get-EXOMailboxFolderPermission $SourceEmailAddress
			Write-Verbose "Exporting mailbox folder permissions to XML..."
			ExportXML -Object $MailboxFolderPermission -Path $GetMailboxFolderPermissionOutput -Email $SourceEmailAddress
		}

		if ($IncludeMailboxStats) {
			Write-Verbose "Getting mailbox folder statistics..."
			$MailboxFolderStatistics = Get-EXOMailboxFolderStatistics $SourceEmailAddress
			Write-Verbose "Exporting mailbox folder statistics to XML..."
			ExportXML -Object $MailboxFolderStatistics -Path $GetMailboxFolderStatistics -Email $SourceEmailAddress
		}

		if ($IncludeMailboxStats) {
			Write-Verbose "Getting mailbox sub folder permissions..."
			$Folders = $MailboxFolderStatistics.folderpath | ForEach-Object { $_.replace("/", "\") }
			$Permissions = $Folders | ForEach-Object { Get-EXOMailboxFolderPermission "$($SourceEmailAddress):$_" -ErrorAction silentlycontinue }
			Write-Verbose "Exporting mailbox sub folder permissions to XML..."
			ExportXML -Object $Permissions -Path $GetMailboxSubFolderPermissionOutput -Email $SourceEmailAddress
		}

		if ($IncludeInboxRules) {
			Write-Verbose "Getting inbox rules..."
			$Rules = Get-InboxRule -Mailbox $SourceEmailAddress
			Write-Verbose "Exporting inbox rules to XML..."
			ExportXML -Object $Rules -Path $GetInboxRulesOutput -Email $SourceEmailAddress
		}

		# CSV results
		Return [PSCustomObject]@{
			Alias                                  = $Mailbox.Alias
			FirstName                              = $User.FirstName
			LastName                               = $User.LastName
			DisplayName                            = $User.DisplayName
			Name                                   = $User.name
			SamAccountName                         = $ADUser.SamAccountName
			RecipientType                          = $User.RecipientType
			RecipientTypeDetails                   = $User.RecipientTypeDetails
			UserPrincipalName                      = $Mailbox.UserPrincipalName
			PrimarySmtpAddress                     = $mailbox.PrimarySmtpAddress
			ExchangeGuid                           = $Mailbox.ExchangeGuid
			ArchiveGuid                            = $Mailbox.ArchiveGuid
			LegacyExchangeDN                       = $Mailbox.LegacyExchangeDN
			EmailAddresses                         = $Mailbox.EmailAddresses -join ";"
			Manager                                = $User.Manager
			Title                                  = $User.Title
			HomePhone                              = $User.HomePhone
			MobilePhone                            = $User.MobilePhone
			OtherHomePhone                         = $User.OtherHomePhone -join ";"
			OtherTelephone                         = $User.OtherTelephone -join ";"
			Phone                                  = $User.Phone
			Fax                                    = $User.Fax
			OtherFax                               = $User.OtherFax -join ";"
			City                                   = $User.City
			Company                                = $User.Company
			Department                             = $User.Department
			Description                            = $ADUser.Description
			Division                               = $ADUser.Division
			EmployeeID                             = $ADUser.EmployeeID
			EmployeeNumber                         = $ADUser.EmployeeNumber
			extensionAttribute1                    = $ADUser.extensionAttribute1
			extensionAttribute2                    = $ADUser.extensionAttribute2
			extensionAttribute3                    = $ADUser.extensionAttribute3
			extensionAttribute4                    = $ADUser.extensionAttribute4
			extensionAttribute5                    = $ADUser.extensionAttribute5
			extensionAttribute6                    = $ADUser.extensionAttribute6
			extensionAttribute7                    = $ADUser.extensionAttribute7
			extensionAttribute8                    = $ADUser.extensionAttribute8
			extensionAttribute9                    = $ADUser.extensionAttribute9
			extensionAttribute10                   = $ADUser.extensionAttribute10
			extensionAttribute11                   = $ADUser.extensionAttribute11
			extensionAttribute12                   = $ADUser.extensionAttribute12
			extensionAttribute13                   = $ADUser.extensionAttribute13
			extensionAttribute14                   = $ADUser.extensionAttribute14
			extensionAttribute15                   = $ADUser.extensionAttribute15
			Office                                 = $user.Office
			Organization                           = $ADUser.Organization
			PostalCode                             = $User.PostalCode
			State                                  = $User.StateOrProvince
			StreetAddress                          = $User.StreetAddress
			LitigationHoldEnabled                  = $Mailbox.LitigationHoldEnabled
			LitigationHoldDate                     = $Mailbox.LitigationHoldDate
			LitigationHoldOwner                    = $Mailbox.LitigationHoldOwner
			LitigationHoldDuration                 = $Mailbox.LitigationHoldDuration
			AcceptMessagesOnlyFrom                 = $Mailbox.AcceptMessagesOnlyFrom -join ";"
			AcceptMessagesOnlyFromDLMembers        = $Mailbox.AcceptMessagesOnlyFromDLMembers -join ";"
			AcceptMessagesOnlyFromSendersOrMembers = $Mailbox.AcceptMessagesOnlyFromSendersOrMembers -join ";"
			HiddenFromAddressListsEnabled          = $Mailbox.HiddenFromAddressListsEnabled
			MailboxSize                            = $Statistics.TotalItemSize
			MailboxItemCount                       = $Statistics.ItemCount
		}
	}
	Function ExportGroupInfo($Email) {
		$SourceEmailAddress = $Email

		Write-Verbose "Getting distribution group..."
		$DL = Get-DistributionGroup $SourceEmailAddress
		Write-Verbose "Exporting distribution group to XML..."
		ExportXML -Object $DL -Path $GetDistributionGroupOutput -Email $SourceEmailAddress

		Write-Verbose "Getting distribution group members..."
		$DLMembers = Get-DistributionGroupMember $SourceEmailAddress
		Write-Verbose "Exporting distribution group members to XML..."
		ExportXML -Object $DLMembers -Path $GetDistributionGroupMemberOutput -Email $SourceEmailAddress
	}
	Function ExportUnifiedGroupInfo($Email) {
		$SourceEmailAddress = $Email

		Write-Verbose "Getting unified group..."
		$UG = Get-UnifiedGroup $SourceEmailAddress
		Write-Verbose "Exporting unified group to XML..."
		ExportXML -Object $UG -Path $GetUnifiedGroupOutput -Email $SourceEmailAddress

		Write-Verbose "Getting unified group members..."
		$UGMembers = Get-UnifiedGroupLinks $SourceEmailAddress -LinkType Members
		Write-Verbose "Exporting unified group members to XML..."
		ExportXML -Object $UGMembers -Path $GetUnifiedGroupMemberOutput -Email $SourceEmailAddress
	}
	##########
	# Setup #
	##########

	# Test required commands
	Write-Verbose "Testing required commands exist..."
	TestRequiredModule

	# Get Group Info
	Try {
		Write-Verbose "Getting distribution group $GroupName..."
		$Group = Get-DistributionGroup -Identity $GroupName
	}
	Catch {
		Write-Host "Group `"$GroupName`" does not exist. Please verify the group name and try again." -ForegroundColor Red
		Break
	}
	Write-Verbose "Getting distribution group members for group $GroupName..."
	$Members = Get-DistributionGroupMember -Identity $Group.Identity
	if (-not $Members) {
		Write-Host "Group `"$GroupName`" has no members. Please verify the group and try again."
		Break
	}

	# Create Output Folders
	$Date = Get-Date -Format yyyy-MM-dd_HH.mm
	$Root = Join-Path $Path $Date
	$GetUserOutput = Join-Path $Root "User_Output_XMLs"
	$GetRecipientOutput = Join-Path $Root "Recipient_Output_XMLs"
	$GetMailboxOutput = Join-Path $Root "Mailbox_Output_XMLs"
	$GetMailboxStatisticsOutput = Join-Path $Root "Mailbox_Statistics_XMLs"
	$GetAzureADUserOutput = Join-Path $Root "Azure_AD_User_Output_XMLs"
	$GetMsolUserOutput = Join-Path $Root "MSOL_User_Output_XMls"
	$GetADUserOutput = Join-Path $Root "AD_User_Output_XMLs"
	$GetRecipientPermissionOutput = Join-Path $Root "Recipient_Permission_XMLs"
	$GetMailboxFolderPermissionOutput = Join-Path $Root "Mailbox_Folder_Permission_XMLs"
	$GetMailboxFolderStatistics = Join-Path $Root "Mailbox_Folder_Statistics_XMLs"
	$GetMailboxSubFolderPermissionOutput = Join-Path $Root "Mailbox_Sub_Folder_Permission_XMLs"
	$GetInboxRulesOutput = Join-Path $Root "Inbox_Rules_XMLs"
	$GetDistributionGroupOutput = Join-Path $Root "Distribution_Group_XMLs"
	$GetDistributionGroupMemberOutput = Join-Path $Root "Distribution_Group_Member_XMLs"
	$GetUnifiedGroupOutput = Join-Path $Root "Unified_Group_XMLs"
	$GetUnifiedGroupMemberOutput = Join-Path $Root "Unified_Group_Member_XMLs"
	$CsvExport = Join-Path $Root "CSV_Exports"

	$Folders = @(
		$GetUserOutput,
		$GetRecipientOutput,
		$GetMailboxOutput,
		$GetMailboxStatisticsOutput,
		$GetAzureADUserOutput,
		$GetMsolUserOutput,
		$GetADUserOutput,
		$GetRecipientPermissionOutput,
		$GetMailboxFolderPermissionOutput,
		$GetMailboxFolderStatistics,
		$GetMailboxSubFolderPermissionOutput,
		$GetInboxRulesOutput,
		$GetDistributionGroupOutput,
		$GetDistributionGroupMemberOutput,
		$GetUnifiedGroupOutput,
		$GetUnifiedGroupMemberOutput,
		$CsvExport
	)
	foreach ($Path in $Folders) {
		Write-Verbose "Creating output folder $Path..."
		CreateFolder -Path $Path
	}
}
PROCESS {
	$Total = $Members | Measure-Object | Select-Object -ExpandProperty Count
	$i = 0
	$CSV_Results = @()
	$ErrorList = @()
	foreach ($Member in $Members) {
		Write-Progress -Id 1 -Activity "Exporting info..." -Status "Objects: [$i/$Total] | Errors: $($ErrorList.Length)" -PercentComplete ($i / $Total * 100) -CurrentOperation "Working on object $($Member.PrimarySmtpAddress)"
		$i++
		$SourceEmailAddress = $Member.PrimarySmtpAddress

		# If user object
		if ($Member.RecipientType -eq "UserMailbox") {
			# Get user info
			Write-Verbose "Working on user $SourceEmailAddress..."
			try {
				$CSV_Results += ExportUserInfo -Email $SourceEmailAddress
			}
			Catch {
				Write-Verbose "Failed to export user $SourceEmailAddress"
				$ErrObject = [PSCustomObject]@{
					"Email"                = $SourceEmailAddress
					"RecipientType"        = $Member.RecipientType
					"RecipientTypeDetails" = $Member.RecipientTypeDetails
					"Error"                = $err
				}
				$ErrorList += $ErrObject
			}
		}

		# If group object
		elseif ($Member.RecipientTypeDetails -eq "MailUniversalDistributionGroup" -or $Member.RecipientTypeDetails -eq "MailUniversalSecurityGroup") {
			# Get group info
			Write-Verbose "Working on group $SourceEmailAddress..."
			try {
				ExportGroupInfo -Email $SourceEmailAddress
			}
			Catch {
				$err = $_
				Write-Verbose "Failed to export group $SourceEmailAddress"
				$ErrObject = [PSCustomObject]@{
					"Email"                = $SourceEmailAddress
					"RecipientType"        = $Member.RecipientType
					"RecipientTypeDetails" = $Member.RecipientTypeDetails
					"Error"                = $err
				}
				$ErrorList += $ErrObject
			}
		}

		# unified groups (O365 Groups Teams)
		elseif ($Member.RecipientTypeDetails -eq "GroupMailbox") {
			Write-Verbose "Working on unified group $SourceEmailAddress..."
			try {
				ExportUnifiedGroupInfo -Email $SourceEmailAddress
			}
			Catch {
				$err = $_
				Write-Verbose "Failed to export unified group $SourceEmailAddress"
				$ErrObject = [PSCustomObject]@{
					"Email"                = $SourceEmailAddress
					"RecipientType"        = $Member.RecipientType
					"RecipientTypeDetails" = $Member.RecipientTypeDetails
					"Error"                = $err
				}
				$ErrorList += $ErrObject
			}
		}

		# unknown object
		else {
			$ErrObject = @{
				"Email"                = $SourceEmailAddress
				"RecipientType"        = $Member.RecipientType
				"RecipientTypeDetails" = $Member.RecipientTypeDetails
				"Error"                = "Not parsed by script"
			}
			$ErrorList += $ErrObject
		}
	}

	#Export CSV results
	if ($CSV_Results) {
		$CSV_Results | Export-Csv -NoTypeInformation (Join-Path $CsvExport "Export.csv")
	}
	if ($ErrorList) {
		$ErrorList | Export-Csv -NoTypeInformation (Join-Path $CsvExport "Error_Export.csv")
	}

}
END {
	Write-Host "Results saved to: $Root"
}
