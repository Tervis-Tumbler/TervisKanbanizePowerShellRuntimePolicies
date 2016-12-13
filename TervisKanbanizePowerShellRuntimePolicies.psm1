function Install-TervisKanbanizeWorkflow {
        
    Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask `
        -ScheduledTaskUserPassword $ScheduledTaskUserPassword `
        -ScheduledTaskFunctionName "Send-EmailRequestingPaylocityReportBeRun" `
        -RepetitionInterval OnceAWeekMondayMorning

    Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask `
        -ScheduledTaskUserPassword $ScheduledTaskUserPassword `
        -ScheduledTaskFunctionName "Invoke-PaylocityToActiveDirectory" `
        -RepetitionInterval OnceAWeekTuesdayMorning

    Install-TervisPaylocity -PathToPaylocityDataExport $PathToPaylocityDataExport -PaylocityDepartmentsWithNiceNamesJsonPath $PaylocityDepartmentsWithNiceNamesJsonPath

    $Credential = Get-PasswordstateCredential -PasswordID 1234
    Install-TervisKanbanize -Email
}

function Move-CompletedCardsThatHaveAllInformationToArchive {
    $OpenTrackITWorkOrders = Get-TervisTrackITUnOfficialWorkOrders
    
    $CardsThatCanBeArchived = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess | 
    where columnpath -In "Done","Archive" |
    where type -ne "None" |
    where assignee -NE "None" |
    where color -in ("#cc1a33","#f37325","#77569b","#067db7") |
    where TrackITID -NotIn $($OpenTrackITWorkOrders.woid)

    foreach ($Card in $CardsThatCanBeArchived) {
        Move-KanbanizeTaskToArchive -CardID $Card.TaskID
    }
}

function Move-CardsInScheduledDateThatDontHaveScheduledDateSet {
    $CardsInScheduledDateThatDontHaveScheduledDateSet = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess -HelpDeskTriageProcess | 
    where columnpath -Match "Waiting for Scheduled date" | 
    where {$_.scheduleddate -eq $null -or $_.scheduleddate -eq "" }
    
    foreach ($Card in $CardsInScheduledDateThatDontHaveScheduledDateSet) {
        Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.TaskID -Column "In Progress.Waiting to be worked on"
    }
}

#Unfinished
function Move-CardsInDoneListThatHaveStillHaveSomethingIncomplete {
    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess
    
    $CardsInDoneList = $Cards |
    where columnpath -Match "Done"
    
    $OpenTrackITWorkOrders = get-TervisTrackITWorkOrders

    $CardsThatAreOpenInTrackITButDoneInKanbanize = Compare-Object -ReferenceObject $OpenTrackITWorkOrders.woid -DifferenceObject $Cardsindonelist.trackitid -PassThru -IncludeEqual |
    where sideindicator -EQ "=="

    foreach ($Card in $CardsThatCanBeArchived){
        Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.TaskID -Column "Archive"
    }
}

function Import-UnassignedTrackItsToKanbanize {

    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess -HelpDeskTriageProcess

    $TriageProcessBoardID = 29
    $TriageProcessStartingColumn = "Requested"

    $UnassignedWorkOrders = Get-UnassignedTrackITs

    foreach ($UnassignedWorkOrder in $UnassignedWorkOrders ) {
        try {
            if($UnassignedWorkOrder.Wo_Num -in $($Cards.TrackITID)) {throw "There is already a card for this Track IT"}

            New-KanbanizeCardFromTrackITWorkOrder -WorkOrder $UnassignedWorkOrder -DestinationBoardID $TriageProcessBoardID -DestinationColumn $TriageProcessStartingColumn
            Edit-TrackITWorkOrder -WorkOrderNumber $WorkOrder.Wo_Num -AssignedTechnician "Backlog" | Out-Null
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
