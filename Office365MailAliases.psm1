<#
.SYNOPSIS
  This module contains functions to create mail aliases in Office 365

.DESCRIPTION
  These mail aliases are created per domain name or organization. This is to make sure
  that organizations get unique email addresses.

.INPUTS
  None

.OUTPUTS
  None

.NOTES
  Author: Jean-Paul van Ravensberg, Cloudenius.com

.EXAMPLE
  Select-MailAlias -DomainName Google.com -ExportAliasesToMailDraft -Verbose

  Create a mail alias for Google.com and provide Verbose output. After selecting the mail alias,
  create a draft mail in the mailbox of the user that contains all the used mail aliases.

.EXAMPLE
  New-MailAlias -NumberOfAliases 9 -Verbose

  Warm up aliases for later use and provide Verbose output.
#>

Function New-MailAlias {
    param(
        [parameter(Mandatory = $false, HelpMessage = "Specify the amount of aliases required")]
        [ValidateNotNullOrEmpty()]
        [int]$NumberOfAliases,

        [parameter(Mandatory = $false, HelpMessage = "Specify the amount of aliases required")]
        [ValidateNotNullOrEmpty()]
        [string]$EmailDomain,

        [parameter(Mandatory = $false, HelpMessage = "Specify the amount of aliases required")]
        [ValidateNotNullOrEmpty()]
        [string]$Owner,

        [parameter(Mandatory = $false, HelpMessage = "Specify the amount of aliases required")]
        [ValidateNotNullOrEmpty()]
        [string]$GroupNamePrefix
    )

    ## Login to Office 365
    If (!(Get-PSSession | Where-Object {$_.ComputerName -eq "outlook.office365.com" -and $_.State -eq "Opened"})) {
        Connect-EXOPSSession
    }

    Write-Verbose "Creating $NumberOfAliases aliases"

    Foreach ($i in 1..$NumberOfAliases) {
        $Random = Get-Random -Minimum 10000 -Maximum 99999
        $GroupName = $GroupNamePrefix + $Random
        $GroupEmail = ($GroupName + "@" + $EmailDomain)

        Write-Verbose "Creating alias $i with name $GroupName"

        # Create the new Distribution Group
        Try {
            New-DistributionGroup -Name $GroupName -Type "Security" `
                -ManagedBy $Owner -PrimarySmtpAddress $GroupEmail
        }

        Catch [Exception] {
            Write-Error "Distribution Group already exists or another error occurred"
            Break
        }

        # Allow external senders to mail to the address & set _CLAIMABLE suffix
        Set-DistributionGroup -Identity $GroupName -RequireSenderAuthenticationEnabled:$false -DisplayName $($GroupName + "_CLAIMABLE")

        # Modify the new Distribution Group with SendOnBehalf permissions
        Add-RecipientPermission -Identity $GroupName -AccessRights SendAs -Trustee $Owner -Confirm:$false

        # Add the owner to the Distribution Group
        Add-DistributionGroupMember -Identity $GroupName -Member $Owner

        Write-Verbose "Created group called $GroupName with owner $Owner"
    }
}

Function Select-MailAlias {
    param(
        [parameter(Mandatory = $true, HelpMessage = "Specify the domain name of the website")]
        [ValidateNotNullOrEmpty()]
        [string]$DomainName,

        [parameter(Mandatory = $true, HelpMessage = "Create a draft mail in the mailbox of the user that contains all the used mail aliases")]
        [switch]$ExportAliasesToMailDraft
    )

    ## Login to Office 365
    If (!(Get-PSSession | Where-Object {$_.ComputerName -eq "outlook.office365.com" -and $_.State -eq "Opened"})) {
        Connect-EXOPSSession
    }

    Write-Verbose "Claiming an alias for $DomainName"

    # Check if domain name already exists in Distribution Group
    $ExistingDistributionGroup = Get-DistributionGroup | Where-Object {$_.DisplayName -like "*$DomainName*"}

    If ($ExistingDistributionGroup) {
        Write-Verbose "Distribution Group already exists. Returning the name of the Distribution Group already in use"

        $EmailDomain = $DistributionGroup.PrimarySmtpAddress.Split('@')[1]
        $DisplayName = $DomainName + " - " + $EmailDomain
    }

    Else {
        # Search for unused alias and return the oldest one
        $ClaimableDistributionGroups = Get-DistributionGroup | Where-Object {$_.DisplayName -Like "*_CLAIMABLE"} | Sort-Object WhenCreatedUtc

        If (!($ClaimableDistributionGroups)) {
            Write-Error "No claimable Mail Aliases found. Please run New-Alias first."
            Break
        }

        While (!($ClaimableDistributionGroups = Get-DistributionGroup | Where-Object {$_.DisplayName -Like "*_CLAIMABLE"})) {
            Write-Output "Waiting for a new claimable Distribution Group. Pause 5 seconds..."
            Start-Sleep -Seconds 5
        }

        Write-Verbose "Found $($ClaimableDistributionGroups.GetType().Count) claimable Distribution Group(s)"

        # Rename unused alias & change description
        $DistributionGroup = $ClaimableDistributionGroups[0]

        Write-Verbose "Picking $DistributionGroup for the rename"

        If ($DistributionGroup.WhenCreated.AddHours(1) -gt (Get-Date)) {
            Write-Warning "Be aware that this alias is <60 minutes old and might not be active yet"
        }

        # Change the Display Name for the Distribution Group
        $EmailDomain = $DistributionGroup.PrimarySmtpAddress.Split('@')[1]
        $DisplayName = $DomainName + " - " + $EmailDomain

        Set-DistributionGroup -Identity $DistributionGroup.Name -DisplayName $DisplayName
    }

    # Create the draft mail in the mailbox of the user that contains all the used mail aliases
    If ($ExportAliasesToMailDraft) {
        New-MailMessage -Body (Get-UsedMailAlias | Select-Object Name, DisplayName | Out-String) -Subject "Used Mailbox Aliases"
    }

    # Return the new name of the alias
    return New-Object PSObject -Property ([ordered]@{"Name" = $DistributionGroup.Name; "DisplayName" = $DisplayName; "E-mail" = $DistributionGroup.PrimarySmtpAddress})
}

Function Get-UsedMailAlias {
    param(
        [parameter(Mandatory = $false, HelpMessage = "Name prefix that is used to identify the Mail Aliases")]
        [ValidateNotNullOrEmpty()]
        [string]$GroupNamePrefix,

        [parameter(Mandatory = $true, HelpMessage = "Create a draft mail in the mailbox of the user that contains all the used mail aliases")]
        [switch]$ExportAliasesToMailDraft
    )

    ## Login to Office 365
    If (!(Get-PSSession | Where-Object {$_.ComputerName -eq "outlook.office365.com" -and $_.State -eq "Opened"})) {
        Connect-EXOPSSession
    }

    # Check if domain name already exists in Distribution Group
    $ExistingDistributionGroup = Get-DistributionGroup | Where-Object `
        {$_.Name -like "$GroupNamePrefix*" -and $_.DisplayName -notlike "*_CLAIMABLE"}

    # Create the draft mail in the mailbox of the user that contains all the used mail aliases
    If ($ExistingDistributionGroup -and $ExportAliasesToMailDraft) {
        New-MailMessage -Body ($ExistingDistributionGroup | Select-Object Name, DisplayName | Out-String) -Subject "Used Mailbox Aliases"
    }

    # Return the new name of the alias(es)
    If ($ExistingDistributionGroup) {
        return $ExistingDistributionGroup | Select-Object Name, DisplayName, PrimarySmtpAddress
    }
    Else {
        return
    }
}

Function Get-UnusedMailAlias {
    param(
        [parameter(Mandatory = $false, HelpMessage = "Name prefix that is used to identify the Mail Aliases")]
        [ValidateNotNullOrEmpty()]
        [string]$GroupNamePrefix
    )

    ## Login to Office 365
    If (!(Get-PSSession | Where-Object {$_.ComputerName -eq "outlook.office365.com" -and $_.State -eq "Opened"})) {
        Connect-EXOPSSession
    }

    # Check if domain name already exists in Distribution Group
    $ExistingDistributionGroup = Get-DistributionGroup | Where-Object `
        {$_.Name -like "$GroupNamePrefix*" -and $_.DisplayName -like "*_CLAIMABLE"}

    # Return the names of the unused alias(es)
    If ($ExistingDistributionGroup) {
        return $ExistingDistributionGroup | Select-Object Name, DisplayName, PrimarySmtpAddress
    }
    Else {
        return
    }
}