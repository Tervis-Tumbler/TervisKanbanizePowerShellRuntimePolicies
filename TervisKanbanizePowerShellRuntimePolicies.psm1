function Install-TervisKanbanizePowerShellRuntimePolicies {
    param (
        $PathToScriptForScheduledTask = $PSScriptRoot
    )
    Install-PasswordStatePowerShell

    $ScheduledTasksCredential = Get-PasswordstateCredential -PasswordID 259

    Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask `
        -Credential $ScheduledTasksCredential `
        -ScheduledTaskFunctionName "Invoke-TervisKanbanizePowerShellRuntimePolicies" `
        -RepetitionInterval OnceAWeekMondayMorning

    $KanbanizeCredential = Get-PasswordstateCredential -PasswordID 2998

    Install-TervisKanbanize -Email $KanbanizeCredential.UserName -Pass $KanbanizeCredential.GetNetworkCredential().password
}

function Uninstall-TervisKanbnaizePowerShellRuntimePolicies {
    param (
        $PathToScriptForScheduledTask = $PSScriptRoot
    )
    Uninstall-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask -ScheduledTaskFunctionName "Invoke-TervisKanbanizePowerShellRuntimePolicies"
}

Function Invoke-TervisKanbanizePowerShellRuntimePolicies {
    $Cards = Get-TervisKanbnaizeAllTasksFromAllBoards
    $OpenTrackITWorkOrders = Get-TervisTrackITUnOfficialWorkOrder

    Move-CompletedCardsThatHaveAllInformationToArchive -Cards $Cards -WorkOrders $OpenTrackITWorkOrders
    Import-UnassignedTrackItsToKanbanize -Cards $Cards
    Move-CardsInWaitingForScheduledDateThatDontHaveScheduledDateSet -Cards $Cards
    Move-CardsInWaitingForScheduledDateThatHaveReachedTheirDate -Cards $Cards
    Move-CardsInWaitingForScheduledDateThatHaveCommentAfterMovement -Cards $Cards
}

function Move-CompletedCardsThatHaveAllInformationToArchive {
    param (
        $Cards,
        $WorkOrders
    )
    
    $CardsThatCanBeArchived = $Cards | 
    where columnpath -In "Done","Archive" |
    where type -ne "None" |
    where assignee -NE "None" |
    where TrackITID -NotIn $($OpenTrackITWorkOrders.woid)

    foreach ($Card in $CardsThatCanBeArchived) {
        Move-KanbanizeTaskToArchive -CardID $Card.TaskID
    }
}

function Move-CardsInWaitingForScheduledDateThatDontHaveScheduledDateSet {
    param (
        $Cards
    )
    $CardsInScheduledDateThatDontHaveScheduledDateSet = $Cards | 
    where columnpath -Match "Waiting for Scheduled date" | 
    where {$_.scheduleddate -eq $null -or $_.scheduleddate -eq "" }
    
    foreach ($Card in $CardsInScheduledDateThatDontHaveScheduledDateSet) {
        Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.TaskID -Column "In Progress.Waiting to be worked on"
    }
}

function Move-CardsInWaitingForScheduledDateThatHaveReachedTheirDate {
    param (
        $Cards
    )
    $CardsInScheduledDateThatHaveReachedTheirDate = $Cards | 
    where columnpath -Match "Waiting for Scheduled date" | 
    where {(Get-Date $_.scheduleddate) -le (Get-Date) }
    
    foreach ($Card in $CardsInScheduledDateThatHaveReachedTheirDate) {
        Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.TaskID -Column "In Progress.Waiting to be worked on"
    }
}

Function Move-CardsInWaitingForScheduledDateThatHaveCommentAfterMovement {
    param (
        $Cards
    )
    $CardsInWaitingForScheduledDate = $Cards | 
    where columnpath -Match "Waiting for Scheduled date"

    foreach ($Card in $CardsInWaitingForScheduledDate) {
        $CardDetails = Get-KanbanizeTaskDetails -BoardID $Card.boardid -TaskID $Card.taskid -History yes | 
        Add-TervisKanbanizeCardDetailsProperties -PassThru
    
        $LastCommentDate = $CardDetails.HistoryDetails |
        Where-Object historyevent -eq "Comment added" |
        Sort-Object EntryDateTime -Descending | 
        Select-Object -First 1 |
        Select-Object -ExpandProperty EntryDateTime
    
        $DateMovedToScheduledDateColumn = $CardDetails.HistoryDetails |
        Where-Object historyevent -eq "Task moved" |
        Where-Object TransitionToColumn -EQ "In Progress.Waiting for scheduled date" |
        Sort-Object EntryDateTime -Descending | 
        Select-Object -First 1 |
        Select-Object -ExpandProperty EntryDateTime

        if ($LastCommentDate -gt $DateMovedToScheduledDateColumn) {
            Move-KanbanizeTask -BoardID $CardDetails.boardid -TaskID $CardDetails.taskid -Column "In Progress.Waiting to be worked on"
        }
    }
}

#function Import-UnassignedTrackItsToKanbanize {
#    param (
#        $Cards
#    )
#
#    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess -HelpDeskTriageProcess
#
#    $TriageProcessBoardID = 29
#    $TriageProcessStartingColumn = "Requested"
#
#    $WorkOrdersWithoutKanbanizeID = Get-TervisTrackITUnOfficialWorkOrder | where { -not $_.KanbanizeID }
#
#    foreach ($UnassignedWorkOrder in $UnassignedWorkOrders ) {
#        try {
#            if($UnassignedWorkOrder.Wo_Num -in $($Cards.TrackITID)) {throw "There is already a card for this Track IT"}
#
#            New-KanbanizeCardFromTrackITWorkOrder -WorkOrder $UnassignedWorkOrder -DestinationBoardID $TriageProcessBoardID -DestinationColumn $TriageProcessStartingColumn
#            Edit-TrackITWorkOrder -WorkOrderNumber $WorkOrder.Wo_Num -AssignedTechnician "Backlog" | Out-Null
#        } catch {            
#            $ErrorMessage = "Error running Import-UnassignedTrackItsToKanbanize: " + $UnassignedWorkOrder.Wo_Num + " -  " + $UnassignedWorkOrder.Task
#            Send-MailMessage -From HelpDeskBot@tervis.com -to HelpDeskDispatch@tervis.com -subject $ErrorMessage -SmtpServer cudaspam.tervis.com -Body $_.Exception|format-list -force
#        }
#    }
#}

Function Get-BusinesssServicesCardAnalysis {
    $WorkOrdersWithoutKanbanizeID = Get-TervisTrackITUnOfficialWorkOrder | where { -not $_.KanbanizeID }
    $WorkOrdersWithoutKanbanizeID| group type
    Groups = $WorkOrdersWithoutKanbanizeID| group type -AsHashTable -AsString
    $Groups.'Business Services'|group respons | sort count -Descending
}

function Import-TrackItsToKanbanize {
    param (
        $Cards
    )

    $TriageProcessStartingColumn = "Requested"
    
    $TypeToTriageBoardIDMapping = [PSCustomObject][Ordered]@{
        WorkOrderType = "Technical Services" 
        TriageBoardID = 29
    }
    #,
    #[PSCustomObject][Ordered]@{
    #    WorkOrderType = "Business Services"
    #    TriageBoardID = 71
    #}

    $WorkOrdersToImport = Get-TervisTrackITUnOfficialWorkOrder | 
    where Type -In $TypeToTriageBoardIDMapping.WorkOrderType | 
    where { -not $_.KanbanizeID }

    foreach ($WorkOrderToImport in $WorkOrdersToImport ) {
        try {
            if($WorkOrderToImport.Wo_Num -in $($Cards.TrackITID)) {throw "There is already a card for this Track IT"}

            $DestinationBoardID = $TypeToTriageBoardIDMapping | 
            where WorkOrderType -EQ $WorkOrderToImport.Type |
            select -ExpandProperty TriageBoardID

            New-KanbanizeCardFromTrackITWorkOrder -WorkOrder $WorkOrderToImport -DestinationBoardID $DestinationBoardID -DestinationColumn $TriageProcessStartingColumn
        } catch {            
            $ErrorMessage = "Error running Import-UnassignedTrackItsToKanbanize: " + $UnassignedWorkOrder.Wo_Num + " -  " + $UnassignedWorkOrder.Task
            Send-MailMessage -From HelpDeskBot@tervis.com -to HelpDeskDispatch@tervis.com -subject $ErrorMessage -SmtpServer cudaspam.tervis.com -Body $_.Exception|format-list -force
        }
    }
}


Function Close-WorkOrdersForEmployeesWhoDontWorkAtTervis {
    $WorkOrders = Get-TervisTrackITUnOfficialWorkOrder
    $ADUser = get-aduser -Filter *
    $WorkOrdersWithoutRequestorInAD = $WorkOrders | where REQUEST_EMAIL -NotIn $ADUser.UserPrincipalName
    $WorkOrdersWithoutRequestorInAD | group request_email
}
