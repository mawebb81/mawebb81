<#
.SYNOPSIS
Function to create a Call Queue in Teams along with a group for members and phone number

.DESCRIPTION
The function will create a Call Queue in Teams with the settings as per the defined parameters along with a Resource Account.
It will create an M365 group and add the specified users to it and then assign that group as the agent group for the queue. 
Function then moves on to finding a random unallocated phone number in the tenant, assigning it to the call queue and finally associates the CQ to the Resource Account

.PARAMETER CallQueueName
Define the name of the Call Queue, this will be the display name

.PARAMETER TimeoutThreshold
Set the required timeout threshold for the CQ. Defines the time (in seconds) that a call can be in the queue before that call times out.
The TimeoutThreshold can be any integer value between 0 and 2700 seconds (inclusive), and is rounded to the nearest 15th interval.

.PARAMETER OverflowThreshold
Set the required overflow threshold for the CQ. Defines the number of calls that can be in the queue at any one time before the overflow action is triggered
The OverflowThreshold can be any integer value between 0 and 200, inclusive. A value of 0 causes calls not to reach agents and the overflow action to be taken immediately

.PARAMETER RoutingMethod
Set the routing method for the CQ. This can be one of the following values: Attendant, Serial, RoundRobin or LongestIdle

.PARAMETER PresenceBasedRouting
Set whether Presence Based Routing should be enabled for the CQ. Boolean value of $True or $False

.PARAMETER Users
Enter the users (agents) for the Call Queue. The list here should be entered individually within speech marks and separated with commas i.e. "user1@domain.com", "user2@domain.com"

.PARAMETER DefaultMusicOnHold
Specify whether the default music on hold should be used, must be set if no custom music is defined. Boolean value of $True or $False

.PARAMETER Owner
Define who the Owner of the CQ is. This should be a single username in the format "user1@domain.com". This user will also be set as the Owner of the M365 group

.PARAMETER AllowOptOut
Specify whether users/agents are allowed to opt in or out of taking calls for the Call Queue. Boolean value of $True or $False

.PARAMETER ConferenceMode
Specify whether Conference mode should be enabled or not. 
Conference mode significantly reduces the amount of time it takes for a caller to be connected to an agent, after the agent accepts the call. In most cases this should
be set to $True. Refer to Microsoft documentation for further info. Boolean value of $True or $False

.PARAMETER AgentAlertTime
Specify the Agent Alert Time for the Call Queue. Represents the time (in seconds) that a call can remain unanswered before it is automatically routed to the next agent. 
Integer value which must be between 15 and 180. If not specified will be set to 30.

.PARAMETER WelcomeTextToSpeech
Provide any required welcome text to speech words that will be played once the caller connects to the Call Queue. String value.

.EXAMPLE
New-MWCallQueue -CallQueueName "Test Call Queue" -TimeoutThreshold 30 -OverflowThreshold 40 -RoutingMethod LongestIdle -PresenceBasedRouting $true -Users "mark.webb@mwdev20.co.uk","adelev@mwdev20.co.uk" -DefaultMusicOnHold $true -Owner "adelev@mwdev20.co.uk" -AllowOptOut $true -ConferenceMode $true -AgentAlertTime 45 -WelcomeTextToSpeech "hello"

.NOTES
Initial Release
#>
Function New-MWCallQueue {

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
        [boolean]
        $PresenceBasedRouting,
    
        [Parameter(Mandatory=$true)]
        [String[]]
        $Users,
    
        [Parameter(Mandatory=$true)]
        [boolean]
        $DefaultMusicOnHold,
    
        [Parameter(Mandatory=$true)]
        [String]
        $Owner,
    
        [Parameter(Mandatory=$true)]
        [boolean]
        $AllowOptOut,
    
        [Parameter(Mandatory=$true)]
        [boolean]
        $ConferenceMode=$True,
    
        [Parameter(Mandatory=$true)]
        [ValidateRange(15,180)]
        [Int]
        $AgentAlertTime,
    
        [Parameter(Mandatory=$true)]
        [String]
        $WelcomeTextToSpeech
    
    )
    
    $CallQueueNameNS = $CallQueueName -replace '\s',''
    
    $UPN = "CQ-" + $CallQueueNameNS + "@mwdev20.co.uk"
    
    Write-Host "Connecting to Microsoft Teams and Graph, enter your credentials when prompted"
    
    Connect-MicrosoftTeams
    
    Connect-MgGraph -Scopes Directory.ReadWrite.All, Group.ReadWrite.All, User.Read.All
    
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
        "Owners@odata.bind" = @($Owners)
    }
    
    
    
    Try {
    
    Write-Host "Creating new M365 group to store members for Call Queue" -ForegroundColor Green
    
    New-MgGroup -BodyParameter $params
    
    Write-Host "Wait 10 seconds for replication" -ForegroundColor Green
    
    Start-Sleep -Seconds 10
    
    Write-Host "Creating new resource account for Call Queue" -ForegroundColor Green
    
    New-CsOnlineApplicationInstance -UserPrincipalName $UPN -ApplicationId 11cd3e2e-fccb-42ad-ad00-878b93575e07 -DisplayName $CallQueueName
    
    Write-Host "Created Resource account, waiting 30 seconds for replication" -ForeGroundColor Green
    
    $GroupID = Get-MgGroup -Filter "DisplayName eq '$CallQueueName'"
    
    #$GroupID = Get-AzureADGroup -SearchString $CallQueueNameNS
    
    Start-Sleep -Seconds 30
    
    Update-MGUser -UserID $UPN -UsageLocation "GB"
    
    Write-Host "Creating Call Queue" -ForegroundColor Green
    
    New-CsCallQueue -Name $CallQueueName -OverflowThreshold $OverflowThreshold -TimeoutThreshold $TimeoutThreshold -RoutingMethod $RoutingMethod -PresenceBasedRouting $PresenceBasedRouting -UseDefaultMusicOnHold $DefaultMusicOnHold -DistributionLists $GroupID.ID -AgentAlertTime $AgentAlertTime -AllowOptOut $AllowOptOut -WelcomeTextToSpeechPrompt $WelcomeTextToSpeech -ConferenceMode $ConferenceMode
    
    Write-Host "Adding random unassigned phone number to Call Queue. Script will pause for 2 minutes before executing assignment to allow resource account licence allocation" -ForegroundColor Green
    
    Start-Sleep -Seconds 120
    
    $SpareNumbers = Get-CsPhoneNumberAssignment | Where {$_.Capability -eq 'VoiceApplicationAssignment' -and $_.PstnAssignmentStatus -eq 'Unassigned'}
    
    $CQNumber = $SpareNumbers | Get-Random
    
    Set-CsPhoneNumberAssignment -Identity $UPN -PhoneNumber $CQNumber.TelephoneNumber -PhoneNumberType CallingPlan
    
    Write-Host "Phone number assigned. Moving on to Associating resource account to Call Queue" -ForegroundColor Green
    
    $CallQueueID = (Get-CSCallQueue -NameFilter $CallQueueName).Identity
    
    $ResourceID = (Get-CsOnlineApplicationInstance -Identity $UPN).objectId
    
    New-CsOnlineApplicationInstanceAssociation -ConfigurationId $CallQueueID -Identities $ResourceID -ConfigurationType CallQueue
    
    }
    
    catch [System.UnauthorizedAccessException]
    
    {
    Write-Host "You're not connected to Teams/MS Graph, run Connect-MicrosoftTeams or Connect-MgGraph and try again" -ForegroundColor Red
    }
    
    }
    
