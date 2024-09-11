##Get All Call Queues##

$CallQueues = Get-CsCallQueue | Select Name, RoutingMethod, Agents, Users, AllowOptOut, ConferenceMode, AgentAlertTime, OverflowThreshold, TimeoutThreshold, WelcomeTextToSpeechPrompt

##Get Resource Accounts for phone numbers and populate each queue with number##

ForEach ($Queue in $CallQueues) {

$QueueDetails = Get-CsOnlineApplicationInstance | Where {$Queue.Name -eq $_.DisplayName}

$queuePhonenumber = $QueueDetails.PhoneNumber.Replace("tel:+","")

$Queue | Add-Member -NotePropertyName "PhoneNumber" -NotePropertyValue $queuePhonenumber

}

##Convert Routing Method to named value as default writes to SharePoint as a integer for some reason. If ever work out why can change this.##

ForEach ($Queue in $CallQueues) {

If ($queue.RoutingMethod -eq 0) {
    $queue.RoutingMethod = "Attendant"
}

elseIf ($queue.RoutingMethod -eq 1) {
    $queue.RoutingMethod = "Serial"
}

elseIf ($queue.RoutingMethod -eq 2) {
    $queue.RoutingMethod = "Round Robin"
}

elseIf ($Queue.RoutingMethod -eq 3) {
    $queue.RoutingMethod = "Longest Idle"
}
}

##Convert AAD ObjectID's contained in Agents property to UPN's##

ForEach ($Queue in $CallQueues) {

$Agents = $Queue.Agents

ForEach ($agent in $agents) {

$AgentID = (Get-AzureADUser -ObjectId $agent.ObjectId)

$agent.objectID = $agentID.UserPrincipalName
}
}


##Write details to SharePoint List. First, need to convert the agents property to a string as it's currently an object and won't write to list in that state when multiple users in agents property##

ForEach ($Queue in $CallQueues) {

$AgentString = $queue.agents.objectID | Out-String

Add-PnPListItem -List "CallQueues" -Value @{"PresenceBasedRouting" = $Queue.Name; "WelcomGreeting" = $Queue.WelcomeTextToSpeech; "RoutingMethod" = $Queue.RoutingMethod; "Agents" = $AgentString; "AllowOptOut" = $Queue.AllowOptOut; "ConferenceMode" = $Queue.ConferenceMode; "AgentAlertTime" = $Queue.AgentAlertTime; "OverflowThreshold" = $Queue.OverflowThreshold; "TimeoutThreshold" = $Queue.TimeoutThreshold; "WelcomeTexttoSpeech" = $Queue.WelcomeTextToSpeechPrompt; "PhoneNumber" = $Queue.PhoneNumber}

}
