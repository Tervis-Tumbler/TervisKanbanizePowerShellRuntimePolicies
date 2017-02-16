function Install-TervisKanbanizePowerShellRuntimePolicies {
    param (
        $PathToScriptForScheduledTask = $PSScriptRoot
    )
    Install-PasswordStatePowerShell

    $ScheduledTasksCredential = Get-PasswordstateCredential -PasswordID 259

    Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask `
        -Credential $ScheduledTasksCredential `
        -ScheduledTaskFunctionName "Invoke-TervisKanbanizePowerShellRuntimePolicies" `
        -RepetitionInterval EverWorkdayDuringTheDayEvery15Minutes

    Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask `
        -Credential $ScheduledTasksCredential `
        -ScheduledTaskFunctionName "New-EachMondayRecurringCards" `
        -RepetitionInterval OnceAWeekMondayMorning

    Install-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask `
        -Credential $ScheduledTasksCredential `
        -ScheduledTaskFunctionName "New-EachWorkDayRecurringCards" `
        -RepetitionInterval EverWorkdayOnceAtTheStartOfTheDay

    $KanbanizeCredential = Get-PasswordstateCredential -PasswordID 2998

    Install-TervisKanbanize -Email $KanbanizeCredential.UserName -Pass $KanbanizeCredential.GetNetworkCredential().password
}

function New-EachWorkDayRecurringCards {
    $Cards = Get-TervisKanbnaizeAllTasksFromAllBoards
    if ( -not ($Cards | where BoardID -eq 29 | where Title -eq "Gather Kanban cards" )) {
        New-KanbanizeTask -BoardID 29 -Title "Gather Kanban cards" -Type "Kanban cards gather" -Column Requested
    } else {
        Send-MailMessage -From HelpDeskBot@tervis.com -to HelpDeskDispatch@tervis.com -subject "The previous Gather Kanban cards card has not been completed yet" -SmtpServer cudaspam.tervis.com -Body "The previous Gather Kanban cards card has not been completed yet"
    }


    if ( -not ($Cards | where BoardID -eq 29 | where Title -eq "Review requested kanban cards" )) {
        New-KanbanizeTask -BoardID 29 -Title "Review requested kanban cards" -Type "Kanban cards requested review" -Column Requested
    } else {
        Send-MailMessage -From HelpDeskBot@tervis.com -to HelpDeskDispatch@tervis.com -subject "The previous Review requested kanban cards card has not been completed yet" -SmtpServer cudaspam.tervis.com -Body "The previous Review requested kanban cards card has not been completed yet"
    }
}

function New-EachMondayRecurringCards {
    if ( -not ($Cards | where BoardID -eq 29 | where Title -eq "Review requested kanban cards" )) {
        New-KanbanizeTask -BoardID 29 -Title "Review ordered kanban cards" -Type "Kanban cards ordered review" -Column Requested
    } else {
        Send-MailMessage -From HelpDeskBot@tervis.com -to HelpDeskDispatch@tervis.com -subject "The previous Review ordered kanban cards card has not been completed yet" -SmtpServer cudaspam.tervis.com -Body "The previous Review ordered kanban cards card has not been completed yet"
    }
}

function Uninstall-TervisKanbnaizePowerShellRuntimePolicies {
    param (
        $PathToScriptForScheduledTask = $PSScriptRoot
    )
    Uninstall-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask -ScheduledTaskFunctionName "Invoke-TervisKanbanizePowerShellRuntimePolicies"
    Uninstall-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask -ScheduledTaskFunctionName "New-EachMondayRecurringCards"
    Uninstall-PowerShellApplicationScheduledTask -PathToScriptForScheduledTask $PathToScriptForScheduledTask -ScheduledTaskFunctionName "New-EachWorkDayRecurringCards"
}

function Invoke-TervisKanbanizePowerShellRuntimePolicies {
    $Cards = Get-TervisKanbnaizeAllTasksFromAllBoards
    $WorkOrders = Get-TervisTrackITUnOfficialWorkOrder

    Import-TrackItsToKanbanize -Cards $Cards -WorkOrders $WorkOrders
    Move-CompletedCardsThatHaveAllInformationToArchive -Cards $Cards -WorkOrders $WorkOrders
    Move-CardsInWaitingForScheduledDateThatDontHaveScheduledDateSet -Cards $Cards
    Move-CardsInWaitingForScheduledDateThatHaveReachedTheirDate -Cards $Cards
    Move-CardsInWaitingForScheduledDateThatHaveCommentAfterMovement -Cards $Cards
    Add-WorkInstructionLinkToCards -Cards $Cards
    Remove-HelpDeskCardsOlderThan30Days -Cards $Cards
}

function Remove-HelpDeskCardsOlderThan30Days {
    param(
        $Cards
    )
    
    $CardsToBeDeleted = $Cards | 
    where boardid -eq 32 |
    where lanename -NE "Unplanned Work" |
    where columnpath -Match requested |
    where CreatedAtDateTime -LT (Get-Date).AddDays(-30) 

    foreach($Card in $CardsToBeDeleted) {
        if ($Card.TrackITID) {
            try {
                Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot
                $Result = $null
                $Result = Close-TrackITWorkOrder -WorkOrderNumber $Card.TrackITID -Resolution @"
The Tervis Help Desk team always strives to resolve all requests in a timely, professional manner.  

However, as is the case at most companies, the volume of requests has always outweighed the resources available to resolve issues. 

Therefore, in the spirit of managing this backlog, effective immediately, all requests that have remained unresolved for 30 days will be closed.  If the original requestor feels the issue still requires attention from the Help Desk, they will need to resubmit a request.

A process exists to ‘expedite’ resolution, but this will require manager approval and the expectation of management is that the request is truly urgent and essential.

Thank you in advance for your cooperation and understanding.
"@
                $Result
            } catch { continue }            
        }
        if (-not $Result -or $Result.Success -eq "true" -or $Result.data.Code -eq "Business.HelpDesk.046" ) {
            Remove-KanbanizeTask -BoardID 32 -TaskID $Card.taskid
        }
    }
}

function Add-WorkInstructionLinkToCards {
    param(
        $Cards
    )
    
    $TervisKanbanizePowerShellTypeMetaData = Get-TervisKanbanizePowerShellTypeMetaData -Cards $Cards

    $TypesWithWorkInsructions = $TervisKanbanizePowerShellTypeMetaData | 
    where WorkInstruction |
    Select -ExpandProperty Type

    $CardsThatNeedWorkInstructionAdded = $Cards |
    where Type -ne "None" |
    where Type -In $TypesWithWorkInsructions |
    where {-Not $_.WorkInstruction}
    
    foreach ($Card in $CardsThatNeedWorkInstructionAdded) {
        $WorkInstructionURI = Get-WorkInstructionURI -Type $Card.Type -Cards $Cards
        if ($WorkInstructionURI) {
            Edit-KanbanizeTask -TaskID $Card.taskid -BoardID $Card.BoardID -CustomFields @{"Work Instruction"="$WorkInstructionURI"}
        }
    }
}

function Invoke-FixBrokenTypesOnCardsWithWorkInstructions {
    param (
        $Cards
    )
    $CardsThatNeedFixing = $Cards | where boardid -NE 32 | where boardid -ne 33 | where WorkInstruction
    $TervisKanbanizePowerShellTypeMetaData = Get-TervisKanbanizePowerShellTypeMetaData -Cards $Cards
    
    foreach ($Card in $CardsThatNeedFixing) {
        $Type = $TervisKanbanizePowerShellTypeMetaData | where WorkInstruction -EQ $Card.WorkInstruction | select -ExpandProperty Type
        Edit-KanbanizeTask -TaskID $Card.taskid -BoardID $Card.BoardID -Type $Type
    }
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

function Get-BusinesssServicesCardAnalysis {
    $WorkOrdersWithoutKanbanizeID = Get-TervisTrackITUnOfficialWorkOrder | where { -not $_.KanbanizeID }
    $WorkOrdersWithoutKanbanizeID| group type
    Groups = $WorkOrdersWithoutKanbanizeID| group type -AsHashTable -AsString
    $Groups.'Business Services'|group respons | sort count -Descending
}

function Import-TrackItsToKanbanize {
    param (
        $Cards,
        $WorkOrders
    )

    $TriageProcessStartingColumn = "Requested"
    
    $TypeToTriageBoardIDMapping = [PSCustomObject][Ordered]@{
        WorkOrderType = "Technical Services" 
        TriageBoardID = 29
    }<#,
    [PSCustomObject][Ordered]@{
        WorkOrderType = "Business Services"
        TriageBoardID = 71
    }#>

    $WorkOrdersToImport = $WorkOrders | 
    where Type -In $TypeToTriageBoardIDMapping.WorkOrderType | 
    where WOTYPE2 -NE "EBS" |
    where { -not $_.KanbanizeID }

    foreach ($WorkOrderToImport in $WorkOrdersToImport ) {
        try {
            $CardThatAlreadyExistsForWorkOrderBeingImported = $Cards | 
            where TrackITID -EQ $WorkOrderToImport.Wo_Num

            if($CardThatAlreadyExistsForWorkOrderBeingImported) {
                Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot
                $Response = Edit-TervisTrackITWorkOrder -WorkOrderNumber $WorkOrder.Wo_Num -KanbanizeCardID $CardThatAlreadyExistsForWorkOrderBeingImported.taskid | Out-Null
                
                if (($Response.success | ConvertTo-Boolean) -eq $false) {
                    throw "There is already a card for this Track IT"
                }
            } else {          
                $DestinationBoardID = $TypeToTriageBoardIDMapping | 
                where WorkOrderType -EQ $WorkOrderToImport.Type |
                select -ExpandProperty TriageBoardID

                New-KanbanizeCardFromTrackITWorkOrder -WorkOrder $WorkOrderToImport -DestinationBoardID $DestinationBoardID -DestinationColumn $TriageProcessStartingColumn
            }
        } catch {            
            $ErrorMessage = "Error running Import-UnassignedTrackItsToKanbanize: " + $WorkOrderToImport.Wo_Num + " -  " + $WorkOrderToImport.Task
            Send-MailMessage -From HelpDeskBot@tervis.com -to HelpDeskDispatch@tervis.com -subject $ErrorMessage -SmtpServer cudaspam.tervis.com -Body $_.Exception|format-list -force
        }
    }
}


function Close-WorkOrdersForEmployeesWhoDontWorkAtTervis {
    $WorkOrders = Get-TervisTrackITUnOfficialWorkOrder
    $ADUser = get-aduser -Filter *
    $WorkOrdersWithoutRequestorInAD = $WorkOrders | where REQUEST_EMAIL -NotIn $ADUser.UserPrincipalName
    $WorkOrdersWithoutRequestorInAD | group request_email
}
