Param (
    [Parameter(Mandatory=$true)]
    [String]
    $CallQueueName,

    [Parameter(Mandatory=$true)]
    [ValidateRange(0,2700)]
    [Int]
    $TimeoutThreshold,

    [Parameter(Mandatory=$true)]
    [ValidateRange(0,200)]
    [Int]
    $OverflowThreshold,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Attendant","Serial","RoundRobin","LongestIdle")]
    [String]
    $RoutingMethod,

    [Parameter(Mandatory=$true)]
    [String]
    $PresenceBasedRouting,

    [Parameter(Mandatory=$true)]
    [String[]]
    $Users,

    [Parameter(Mandatory=$true)]
    [String]
    $DefaultMusicOnHold,

    [Parameter(Mandatory=$true)]
    [String]
    $Owner,

    [Parameter(Mandatory=$true)]
    [String]
    $AllowOptOut,

    [Parameter(Mandatory=$true)]
    [string]
    $ConferenceMode,

    [Parameter(Mandatory=$true)]
    [ValidateRange(15,180)]
    [Int]
    $AgentAlertTime,

    [Parameter(Mandatory=$true)]
    [String]
    $WelcomeTextToSpeech

)

$creds = Get-AutomationPSCredential -Name 'ServiceAccount'

$AppClientID = Get-AutomationVariable -Name 'ClientID'

$TenTenantID = Get-AutomationVariable -Name 'TenantID'

$Certificate = Get-AutomationVariable -Name 'Cert'

Connect-MicrosoftTeams -Credential $creds

Connect-AzureAD -Credential $creds

Connect-MGGraph -ClientID $AppClientID -TenantID $TenTenantID -CertificateThumbprint $Certificate


$CallQueueNameNS = $CallQueueName -replace '\s',''

$UPN = "CQ-" + $CallQueueNameNS + "@mwdev20.co.uk"

$ConvertedConfMode = Switch ($ConferenceMode) {'yes'{$True}'No'{$False}}

$ConvertedPresRouting = Switch ($PresenceBasedRouting) {'yes'{$True}'No'{$False}}

$ConvertedAllowOptOut = Switch ($AllowOptOut) {'yes'{$True}'No'{$False}}

$ConvertedDefaultMoH = Switch ($DefaultMusicOnHold) {'yes'{$True}'No'{$False}}

Write-Host "Creating new M365 group to store members for Call Queue" -ForegroundColor Green

#$CQUsers = $users.Split(" ")

$GroupMembers = ForEach ($user in $users) {

"https://graph.microsoft.com/v1.0/users/" + $user

}

$Owners = "https://graph.microsoft.com/v1.0/users/" + $owner

$params = @{
	description = $CallQueueName
	displayName = $CallQueueName
	groupTypes = @(
		"Unified"
	)
	mailEnabled = $true
	mailNickname = "$CallQueueNameNS"
	securityEnabled = $false
    Visibility = "Private"
    "Members@odata.bind" = @($Groupmembers)
    "owners@odata.bind" = @($Owners)
}

New-MgGroup -BodyParameter $params

#New-UnifiedGroup -AccessType Private -Alias $CallQueueNameNS -AlwaysSubscribeMembersToCalendarEvents:$False -AutoSubscribeNewMembers $false -Owner $owner -DisplayName $CallQueueName -EmailAddresses $UPN -Members $users

Write-Host "Wait 10 seconds for replication" -ForegroundColor Green

Start-Sleep -Seconds 10

Write-Host "Creating new resource account for Call Queue" -ForegroundColor Green

New-CsOnlineApplicationInstance -UserPrincipalName $UPN -ApplicationId "11cd3e2e-fccb-42ad-ad00-878b93575e07" -DisplayName $CallQueueName

Write-Host "Created Resource account, waiting 30 seconds for replication" -ForeGroundColor Green

$GroupID = Get-MgGroup -Filter "DisplayName eq '$CallQueueName'"

Start-Sleep -Seconds 30

Update-MGUser -UserID $UPN -UsageLocation "GB"

#Set-AzureADUser -ObjectId $UPN -UsageLocation "GB"

Write-Host "Creating Call Queue" -ForegroundColor Green

New-CsCallQueue -Name $CallQueueName -OverflowThreshold $OverflowThreshold -TimeoutThreshold $TimeoutThreshold -RoutingMethod $RoutingMethod -PresenceBasedRouting $ConvertedPresRouting -UseDefaultMusicOnHold $ConvertedDefaultMoH -DistributionLists $GroupID.Id -AgentAlertTime $AgentAlertTime -AllowOptOut $ConvertedAllowOptOut -WelcomeTextToSpeechPrompt $WelcomeTextToSpeech -ConferenceMode $ConvertedConfMode

Write-Host "Adding random unassigned phone number to Call Queue. Script will pause for 2 minutes before executing assignment to allow resource account licence allocation" -ForegroundColor Green

Start-Sleep -Seconds 120

$SpareNumbers = Get-CsPhoneNumberAssignment | Where {$_.Capability -eq 'VoiceApplicationAssignment' -and $_.PstnAssignmentStatus -eq 'Unassigned'}

$CQNumber = $SpareNumbers | Get-Random

Set-CsPhoneNumberAssignment -Identity $UPN -PhoneNumber $CQNumber.TelephoneNumber -PhoneNumberType CallingPlan

Write-Host "Phone number assigned. Moving on to Associating resource account to Call Queue" -ForegroundColor Green

$CallQueueID = (Get-CSCallQueue -NameFilter $CallQueueName).Identity

$ResourceID = (Get-CsOnlineApplicationInstance -Identity $UPN).objectId

New-CsOnlineApplicationInstanceAssociation -ConfigurationId $CallQueueID -Identities $ResourceID -ConfigurationType CallQueue


