' Copyright (c) 2014 Trevor Lowing

'This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Data

Module ModuleReport


    ''' <summary>
    ''' Generates the report.
    ''' </summary>
    ''' <param name="strReportType">Type of the report (Stats, Baseline or Enterprise).</param>
    Public Sub ProjectReport(ByVal strReportType As String)

        Dim objTask As MSProject.Task
        Dim colReport As New SortedDictionary(Of String, String)
        Dim strReport, strColor, strCritical, strStartsIn As String
        Dim dblSlack As Double
        Dim dtmNow As Date
        Dim strComment, strFooter As String
        Dim blnBaselineReport As Boolean = False
        strStartsIn = ""
        strCritical = ""
        strColor = ""
        strComment = ""



        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        If ActiveProject.Tasks.Count < 3 Then
            MsgBox("Very few tasks, this appears to be a blank schedule.")
            Exit Sub
        End If

        Dim intOneDay, intStatusCutoff, intTaskCountLength As Integer
        intOneDay = (60 * ActiveProject.HoursPerDay)
        intStatusCutoff = My.Settings.StatusWeeks * 7
        dtmNow = FutureDate(My.Settings.StatusDay)
        If My.Settings.AutoSetStatusDate Then 'option is enabled in settings to set status date to next selected date
            If dtmNow = Date.Today() Then
                ActiveProject.StatusDate = Nothing
            Else
                ActiveProject.StatusDate = dtmNow
            End If

        End If
        intTaskCountLength = (DateDiff("d", ActiveProject.ProjectStart, ActiveProject.ProjectFinish) + 7).ToString.Length 'use during reporting to make day count constant for sorting in array keys


        For Each objTask In ActiveProject.Tasks
            strComment = ""
            If Not objTask Is Nothing Then
                ' see if the task is in the status window or a full report was asked for
                'ID, Name, Start Finish, Float, Percent, Owner, Owner Email, baselinestart
                If objTask.Summary = False And objTask.Active = True Then
                    dblSlack = objTask.TotalSlack / intOneDay

                    strColor = "black"
                    strCritical = ""
                    If objTask.Critical Then
                        strColor = "red"
                        strCritical = "*"
                    End If
                    strComment = ""
                    strStartsIn = startsIn(objTask.Start, objTask.Finish, dtmNow, intTaskCountLength, objTask.PercentComplete)
                    Select Case strReportType
                        Case "Baseline"
                            strComment = BaselineCheck(objTask)
                            blnBaselineReport = True
                            strStartsIn = "Baseline Report by Task ID"
                        Case "Status"
                            If (DateDiff("d", dtmNow, objTask.Start) <= intStatusCutoff And objTask.Active And objTask.PercentComplete < 100 And objTask.Summary = False) Or (objTask.ActualStart.ToString <> "NA") Then
                                strComment = StatusCheck(objTask, dtmNow)
                            End If
                        Case "Enterprise"
                            strComment = CheckEnterpriseTaskFields(objTask)
                            strStartsIn = "Baseline Report by Task ID"


                    End Select
                End If

                strReport = "<tr><td style='color:" & strColor & "'>" & objTask.ID & strCritical & "</td><td style='text-align: left;'>" & objTask.Name & "</td><td>" & objTask.ResourceNames & "&nbsp;</td><td>" & objTask.PercentComplete & "%</td><td>" & objTask.Duration / intOneDay & "D</td><td>" & FormatDateTime(CDate(objTask.Start), 2) & "</td><td>" & FormatDateTime(CDate(objTask.Finish), 2) & "</td><td>" & Math.Round(dblSlack) & "D</td><td style='text-align: left;'>" & strComment & "</td></tr>" & vbCrLf
                'objProjApp.Application.DateDifference(objTask.BaselineStart-Now(),objProjApp.ActiveProject.Calendar)
                If strComment <> "" Then    'objTask.Start <= dtmNow And objTask.PercentComplete < 100
                    If Len(strStartsIn) > 0 Then

                        If Not colReport.ContainsKey(strStartsIn) Then
                            colReport.Add(strStartsIn, strReport)    'add contact
                        Else
                            colReport(strStartsIn) = colReport(strStartsIn) & strReport    'update existing contact
                        End If
                    Else
                        colReport(strStartsIn) = colReport(strStartsIn) & strReport    ' default contact

                    End If
                End If
            End If

        Next
        strFooter = ""
        If strReportType = "Enterprise" Then
            strFooter = ProjectCheckEnterprise()
        ElseIf strReportType = "Status" Then
            strFooter = ProjectCheck()
        ElseIf strReportType = "Baseline" Then
            If DateDiff("d", ActiveProject.ProjectStart, Now()) > 30 Then
                strFooter = "<ul><li><strong>Notice</strong> - Project Start Date more than 30 days in the past.  Consider setting to current date or date project will baseline.<br/></ul>" & vbCrLf
            End If
        End If

        Call ExportReport(strReportType, colReport, strFooter)



    End Sub

 

    ''' <summary>
    ''' Performes project level checks to default settings based on user's preferences.
    ''' </summary>
    ''' <returns></returns>
    Function ProjectCheck() As String
        'correct calendar
        'enterprise fields
        'no split tasks
        'no autologic
        On Error GoTo Err_Init
        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        Dim strComment As String
        strComment = "<ul>"
        If ActiveProject.Calendar.Name <> My.Settings.DefaultCalendar Then
            strComment = strComment & "<li><strong>Notice</strong> - Project Calendar is not '" & My.Settings.DefaultCalendar & "'.</li>" & vbCrLf
        End If
        If ActiveProject.AutoLinkTasks = True Then
            strComment = strComment & "<li><strong>Notice</strong> - Current project has 'Auto Linked Tasks' Enabled.  Consider disabling.</li>" & vbCrLf
        End If
        If ActiveProject.AutoSplitTasks = True Then
            strComment = strComment & "<li><strong>Notice</strong> - Current project has 'Auto Split Tasks' Enabled.  Consider disabling.</li>" & vbCrLf
        End If
        If ActiveProject.NewTasksEstimated = True Then
            strComment = strComment & "<li><strong>Notice</strong> - Current project has 'New Tasks Estimated' Enabled.  Consider disabling.</li>" & vbCrLf
        End If
        If ActiveProject.ScheduleFromStart = False Then
            strComment = strComment & "<li><strong>Notice</strong> - Current project has 'Schedule From Start' set to False.  Consider setting to True.</li>" & vbCrLf
        End If
        If ActiveProject.AutoTrack = False Then
            strComment = strComment & "<li><strong>Notice</strong> - Current project has 'Auto Track' set to False.  Consider setting to True.</li>" & vbCrLf
        End If
        If ActiveProject.AutoAddResources = True Then
            strComment = strComment & "<li><strong>Notice</strong> - Current project has 'Auto Add Resources' set to True.  Consider setting to False for remider to build tem from Enterprise Resource Pool.</li>" & vbCrLf
        End If
        If ActiveProject.DefaultEffortDriven = True Then
            strComment = strComment & "<li><strong>Notice</strong> - Current project has 'Default Effort Driven' set to True.  Consider setting to False unless Project is Resource constrained.</li>" & vbCrLf
        End If
        If ActiveProject.HonorConstraints = False Then
            strComment = strComment & "<li><strong>Notice</strong> - Current project has 'Honor Constraints' set to False.  Consider setting to True so logic drives schedule.</li>" & vbCrLf
        End If

        If DateDiff("d", ActiveProject.CurrentDate, Now()) > 3 Then
            strComment = strComment & "<li><strong>Notice</strong> - Project 'Current Date' in Project Tab-->Project Information is different from today.  Consider setting to current date or status may not be accurate.</li>" & vbCrLf
        End If

        ProjectCheck = strComment & "</ul>" & vbCrLf
        Exit Function

Err_Init:
        MsgBox("ProjectCheck" & Err.Number & " - " & Err.Description)

    End Function

    ''' <summary>
    ''' Performs checks on Enterprise custom Fields according to user's preferences.
    ''' </summary>
    ''' <returns></returns>
    Function ProjectCheckEnterprise() As String

        Dim strSpecialReport
        Dim strValue, strReturn As Object
        strReturn = ""
        strSpecialReport = ""

        If Not My.Settings.EnterpriseProjectFields Is Nothing Then

            strReturn = "<ul>"
            For Each strItem As String In My.Settings.EnterpriseProjectFields
                strValue = getEnterpriseProjectCFValue(strItem)
                If Len(Trim(strValue)) = 0 Or InStr(strValue, "xxx") > 0 Then
                    strReturn = strReturn & "<li><strong>" & strItem & "</strong> is blank or default. Please enter value.</li>"
                Else
                    strReturn = strReturn & "<li><strong>" & strItem & "</strong>=" & strValue & ".</li>"
                End If
            Next

            For Each strItem As String In My.Settings.CustomReports
                strSpecialReport = strSpecialReport & SpecialReport("Field " & strItem, strItem)
            Next

            ProjectCheckEnterprise = strReturn & "</ul>" & strSpecialReport
        Else
            ProjectCheckEnterprise = ""
            MsgBox("No Project Level Enterprise fields Identified")
        End If

    End Function






    ''' <summary>
    ''' Generates a Special report from the field chosen. Any task with a non-default value for the noted field is placed in the report in dstart date order.
    ''' </summary>
    ''' <param name="strTitle">The string title.</param>
    ''' <param name="strField">The string field.</param>
    ''' <returns></returns>
    Function SpecialReport(ByVal strTitle As String, ByVal strField As String) As String
        On Error GoTo Err_Init
        'stock report contains id, name, start, finish and requested field.  Sort by date.
        Dim objTask As MSProject.Task
        Dim colReport As New SortedDictionary(Of String, String)
        Dim strReport, strValue, strReturn As String
        Dim strKey As Object
        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        If ActiveProject.Tasks.Count < 3 Then
            MsgBox("Very few tasks, this appears to be a blank schedule.")
            SpecialReport = ""
            Exit Function
        End If
        strReturn = ""


        For Each objTask In ActiveProject.Tasks
            If Not objTask Is Nothing Then
                strValue = getEnterpriseTaskCFValue(objTask.ID, strField)
                If strValue <> "No" And Trim(strValue) <> "" Then
                    strReport = "<tr><td>" & objTask.ID & "</td><td style='text-align: left;'>" & objTask.Name & "</td><td>" & FormatDateTime(CDate(objTask.Start), 2) & "</td><td>" & FormatDateTime(CDate(objTask.Finish), 2) & "</td><td style='text-align: left;'>" & strValue & "</td></tr>" & vbCrLf

                    colReport.Add(Format(CDate(objTask.Start), "yyyy/MM/dd") & objTask.ID, strReport)
                End If
            End If
        Next

        If colReport.Count > 0 Then

            strReturn = "<table><tr class='selected'><th colspan='5'>" & strTitle & "</th></tr>"
            strReturn = strReturn & "<tr><th>ID</th><th>Name</th><th>Start</th><th>Finish</th><th>" & strField & "</th></tr>"
            For Each kvp As KeyValuePair(Of String, String) In colReport
                If Len(kvp.Key) > 0 Then
                    strReturn = strReturn & kvp.Value
                End If
            Next kvp
            strReturn = strReturn & "</table>" & vbCrLf
        End If
        SpecialReport = strReturn
        Exit Function

Err_Init:
        MsgBox("SpecialReport" & Err.Number & " - " & Err.Description)

    End Function

    ''' <summary>
    ''' Perfomas normal Project status checks.
    ''' Weekly Status Type Checks
    '''      Late Start
    '''      Late Finish
    '''      Future Actual
    '''      constraint dates
    '''      estimated durations
    '''      fixed work not fixed duration
    '''      start slips
    '''      finish slips
    ''' </summary>
    ''' <param name="objTask">The object task.</param>
    ''' <param name="dtmStatusDate">The DTM status date.</param>
    ''' <returns></returns>
    Function StatusCheck(ByRef objTask As MSProject.Task, ByVal dtmStatusDate As Date) As String

        On Error GoTo Err_Init
        Dim strComment As String
        Dim intDelta As Integer
        strComment = ""

        If objTask.Start <= dtmStatusDate Then
            If objTask.Start <= dtmStatusDate And objTask.PercentComplete = 0 Then
                strComment = strComment & "<strong>Status</strong>- Delayed Start: Start Task or push out start date. <br/>"
            ElseIf objTask.Finish <= dtmStatusDate And objTask.PercentComplete < 100 Then
                strComment = strComment & "<strong>Status</strong>- Missed Tasks: Set % Complete = 100% or increase duration. Tasks that are supposed to have been completed (prior to the status date) with actual or forecast finishes after the baseline date, OR have a finish variance > zero. Helps identify how well (or poorly) the schedule is meeting the baseline plan. Ratio of # of missed tasks to the ""Baseline Count"" should not exceed 5%. (Project Quality Best Practice)<br/>"
            ElseIf objTask.PercentComplete < 100 And objTask.Duration > 0 Then

                intDelta = Globals.ThisAddIn.Application.DateDifference(dtmStatusDate, objTask.Finish, Globals.ThisAddIn.Application.ActiveProject.Calendar)  'TODO can use task calendar?
                If intDelta <= 0 Then ' task should have even finished before the status date
                    intDelta = 100
                Else
                    intDelta = 100 * ((objTask.Duration - (intDelta)) / objTask.Duration)
                End If
                If Math.Abs(intDelta - objTask.PercentComplete) > 5 Then
                    strComment = strComment & "<strong>Status</strong>- Ongoing Task: Please review percent complete (Actual=" & objTask.PercentComplete & "%, Expected=" & intDelta & "%).<br/>"
                Else
                    strComment = strComment & "<strong>Status</strong>- Ongoing Task: Please review percent complete.<br/>"
                End If
            End If
        End If
        If (objTask.ActualStart.ToString <> "NA") Then
            If (objTask.ActualStart > dtmStatusDate) Then
                strComment = strComment & "<strong>Status</strong>- Future Actual Date: Actual Start date is in the future ( " & FormatDateTime(CDate(objTask.ActualStart), 2) & " ). Set start date before or on Status Date. (Project Quality Best Practice)<br/>"
            End If
        End If
        If (objTask.ActualFinish.ToString <> "NA") Then
            If (objTask.ActualFinish > dtmStatusDate) Then
                strComment = strComment & "<strong>Status</strong>- Future Actual Date: Actual Finish date is in the future ( " & FormatDateTime(CDate(objTask.ActualFinish), 2) & " ). Set finish date before or on Status Date.(Project Quality Best Practice)<br/>"
            End If
        End If
        If objTask.ConstraintDate.ToString <> "NA" And objTask.PercentComplete < 100 Then
            strComment = strComment & "<strong>Health</strong>- Hard Constraint Date: Consider removing date constraint since they prevent tasks from being logic-driven. The Ratio of ""Hard Constraint"" tasks to total incomplete tasks should not exceed 5% (Project Quality Best Practice)<br/>"
        End If
        If objTask.Estimated Then
            strComment = strComment & "<strong>Health</strong>- Estimated Duration: Please enter a specific duration for the task.<br/>"
        End If

        ' Check for leads and lags
        strComment = strComment & GetLags(objTask)

        If objTask.BaselineStart.ToString <> "NA" Then    'Application.DateValue("NA")
            If objTask.Start > objTask.BaselineStart And objTask.PercentComplete < 100 Then
                'intDelta = Application.DateDifference(dtmStart, dtmBaselineStart, ActiveProject.Calendar) / (60 * ActiveProject.HoursPerDay)
                intDelta = DateDiff("d", objTask.BaselineStart, objTask.Start)
                strComment = strComment & "<strong>Health</strong>- Slips: Start date is " & intDelta & " days late<br/>"
            End If
        End If
        If objTask.BaselineFinish.ToString <> "NA" Then
            If objTask.Finish > objTask.BaselineFinish And objTask.PercentComplete < 100 Then
                'intDelta = Application.DateDifference(dtmFinish, dtmBaselineFinish, ActiveProject.Calendar) / (60 * ActiveProject.HoursPerDay)
                intDelta = DateDiff("d", objTask.BaselineFinish, objTask.Finish)
                strComment = strComment & "<strong>Health</strong>- Slips: Finish date is " & intDelta & " days late<br/>"
            End If
        End If
        StatusCheck = strComment

        Exit Function

Err_Init:
        MsgBox("StatusCheck" & Err.Number & " - " & Err.Description)
        Resume Next
    End Function



    Function GetLags(ByRef objTask) As String
        On Error GoTo Err_Init

        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject
        Dim strReturn As String
        Dim objTaskDependencies As MSProject.TaskDependencies
        Dim objTaskDependency As MSProject.TaskDependency
        Dim intOneDay As Integer
        intOneDay = (60 * ActiveProject.HoursPerDay)

        strReturn = ""

        If Not objTask Is Nothing Then
            objTaskDependencies = objTask.TaskDependencies
            For Each objTaskDependency In objTaskDependencies
                If objTaskDependency.To.ID = objTask.ID Then
                    If objTaskDependency.Lag <> 0 Then

                        strReturn = strReturn & ": Lag from task " & objTaskDependency.From.ID & " to task " & objTaskDependency.To.ID & "=" & objTaskDependency.Lag / intOneDay & "days."
                    End If
                End If
            Next
        End If

        If strReturn.Length > 0 Then
            strReturn = "<strong>Health</strong>- Dependency Lags or Leads: Please consider removing any lags on linked tasks. Consider adding a task to represent the reason for the lag for better clarity and transparency. Any incomplete task that has a lag in its predecessor can distort float and may cause resource conflicts and distort the critical path. < or = 5% (Project Quality Best Practice)" & strReturn & "<br/>"
        End If
        GetLags = strReturn

        Exit Function

Err_Init:
        MsgBox("GetLags" & Err.Number & " - " & Err.Description & " - " & Err.Source)
        Resume Next
    End Function




    ''' <summary>
    ''' Performs task level checks that should be addressed before conducting baselines.
    ''' Pre- Baseline Type Checks
    '''      Use of FF,SS or SF Constraints
    '''      Logical Starts
    '''      Logical Finishes
    '''      Long Duration
    '''      High float
    '''      No Resources
    '''      Resource on non-work task (summary or milestone)
    '''      Links to Summary Tasks
    '''      Correct Calendar (blank or matches project calendar).
    '''      Rollup turned on for work task
    '''      high float
    '''      effort driven
    '''    TODO: task names have noun and verb
    ''' </summary>
    ''' <param name="objTask">The object task.</param>
    ''' <returns></returns>
    Function BaselineCheck(ByRef objTask As MSProject.Task) As String
        On Error GoTo Err_Init

        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        Dim intOneDay As Integer
        intOneDay = (60 * ActiveProject.HoursPerDay)
        Dim strComment As String
        strComment = ""
        If InStr(objTask.Predecessors & objTask.Successors, "ss") > 0 Or _
           InStr(objTask.Predecessors & objTask.Successors, "ff") > 0 Or _
           InStr(objTask.Predecessors & objTask.Successors, "sf") > 0 Then
            strComment = strComment & "<strong>Health</strong>- SS,FF,SF Relationships: Non-typical (something other Finish-to-Start, or FS) relationships tasks are strongly discouraged since they are difficult to analyze and should be limited to <10 % of total tasks. (Project Quality Best Practice)<br/>"
        End If
        If Len(Trim(objTask.Predecessors)) = 0 Or Len(Trim(objTask.Successors)) = 0 Then
            strComment = strComment & "<strong>Health</strong>- Logic (logic missing):	Any task that is missing a predecessor, or successor, or both Shows how well (or poorly) the schedule is linked together	Every task has a predecessor and a successor (no dangling activities). Logical Start or Finish. (Project Quality Best Practice)<br/>"
        End If
        If (Len(Trim(objTask.Predecessors)) <> 0 Or Len(Trim(objTask.Successors)) <> 0) And objTask.Summary = True Then
            strComment = strComment & "<strong>Health</strong>- Summary Task with Links: Summary tasks should not have links.<br/>"
        End If
        If objTask.Duration / intOneDay > 44 And objTask.PercentComplete < 100 Then
            strComment = strComment & "<strong>Health</strong>- High Duration Task: Any incomplete task that has a duration greater than 44 working days (2 months). Consider breaking task into smaller tasks to make progressing more accurate. Helps make tasks more manageable and provides better insight. Ratio of # of tasks to total incomplete tasks should be < or = 5% (Project Quality Best Practice)<br/>"
        End If
        If objTask.TotalSlack / intOneDay > 44 And objTask.PercentComplete < 100 Then
            strComment = strComment & "<strong>Health</strong>- High Float Task: An incomplete task with float greater than 44 working days (2 months). Please check for a missing predecessor or successor; to check the stability and logic of the network	Ratio of # of tasks with high float to total incomplete tasks; should be < or = 5%. (Project Quality Best Practice)<br/>"
        End If
        If objTask.TotalSlack / intOneDay < 0 And objTask.PercentComplete < 100 Then
            strComment = strComment & "<strong>Health</strong>- Negative Float Task: An incomplete task with float less than 0 working days. The task may be delaying completion of one or more milestones. (Project Quality Best Practice)<br/>"
        End If

        If Len(Trim(objTask.ResourceNames)) = 0 And objTask.PercentComplete < 100 And Not objTask.Summary And Not objTask.Milestone Then
            strComment = strComment & "<strong>Health</strong>- No Resources: Consider adding resource names. Verifify that all tasks with durations of 1 or more days have $ or resources assigned. (Project Quality Best Practice)<br/>"
        End If
        If Len(Trim(objTask.ResourceNames)) <> 0 And (objTask.Summary Or objTask.Milestone) Then
            strComment = strComment & "<strong>Health</strong>- Non Work Summary or Milestone Task with Resource: Please remove resource names.<br/>"
        End If
        If objTask.Rollup = True And Not objTask.Summary = True And Not objTask.Milestone = True Then
            strComment = strComment & "<strong>Notice</strong>- Rollup Turned On for regular task.<br/>"
        End If

        ' Check for leads and lags
        strComment = strComment & GetLags(objTask)


        'MSProject.PjTaskFixedType.pjFixedDuration.ToString()


        If objTask.Type <> ActiveProject.DefaultTaskType And objTask.PercentComplete < 100 And Not (objTask.Summary) Then
            strComment = strComment & "<strong>Health</strong>- Task Type (" & [Enum].GetName(GetType(MSProject.PjTaskFixedType), objTask.Type) & ") does not match project default task type (" & ActiveProject.DefaultTaskType.ToString & "): Consider changing Task Type to match.<br/>"
        End If
        If objTask.EffortDriven And ActiveProject.DefaultTaskType = MSProject.PjTaskFixedType.pjFixedDuration And objTask.PercentComplete < 100 = False Then
            strComment = strComment & "<strong>Health</strong>- Fixed Duration Task is marked Effort Driven. This may cause problems if you do not intend to have a resource driven plan.<br/>"
        End If
        If objTask.BaselineStart.ToString = "NA" Or objTask.BaselineFinish.ToString = "NA" Then
            strComment = strComment & "<strong>Notice</strong>- Task Baseline Dates Missing: Baseline Start or Finish Date Missing.<br/>"
        End If
        If objTask.Calendar <> "" And objTask.Calendar <> "None" And objTask.Calendar <> ActiveProject.Calendar.Name Then
            strComment = strComment & "<strong>Notice</strong>- Task Calendar differs from Project Calendar (Task=" & objTask.Calendar & ", Project=" & ActiveProject.Calendar.Name & "). Ensure calendar override is intentional.<br/>"
        End If

        If getEnterpriseTaskCFValue(objTask.ID, "PTO_KeyActivityFlag") = "Yes" And getEnterpriseTaskCFValue(objTask.ID, "PTO_ExecutiveReportingDescription") = "" Then
            strComment = strComment & "<strong>Notice</strong>- Executive Desc Missing for item marked Key Activity (PTO Reporting View)." & getEnterpriseTaskCFValue(objTask.ID, "PTO_KeyActivityFlag") & "<br/>"
        End If

        BaselineCheck = strComment
        Exit Function

Err_Init:
        MsgBox("BaselineCheck" & Err.Number & " - " & Err.Description)
        Resume Next
    End Function

    ''' <summary>
    ''' Generates a ProjectQuality report .
    ''' </summary>
    Sub GetProjectQualityReport()

        'stock report contains id, name, start, finish and requested field.  Sort by date.
        Dim objTask As MSProject.Task
        Dim colReport As New SortedDictionary(Of String, String)
        Dim strTitle, strReturn, strDetails, strNote, strPath, strDTM, strCSS, strType As String
        Dim intTaskCount, intTaskCountSelected, intTaskCountSelectedTemp, intDelta As Integer
        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project

        Dim dtmStatusDate As Date
        dtmStatusDate = FutureDate(My.Settings.StatusDay)
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        If ActiveProject.Tasks.Count < 3 Then
            MsgBox("Very few tasks, this appears to be a blank schedule.")
            Exit Sub
        End If

        strReturn = ""
        strNote = ""

        strTitle = "Project Quality Report for " & ActiveProject.Name & " as of " & FormatDateTime(FutureDate(My.Settings.StatusDay), vbShortDate).ToString

        strReturn = "<table><tr class='selected'><th colspan='5'>" & strTitle & "</th></tr>"
        strReturn = strReturn & "<tr><th>Test</th><th>Result</th><th>Limit</th><th>Reason</th><th>Details</th></tr>"

        ' Test, Results, Limit, Reason, Details
        'task count
        'link count
        ' Logic (logic missing)
        ' Lags < or = 5%
        ' Leads
        ' No-FS links Non-typical tasks limited to <10 % of total tasks
        'Hard Constraints -Ratio of "Hard Constraint" tasks to total incomplete tasks should not exceed 5%
        'High Float
        'Ratio of # of tasks with high float to total incomplete tasks; should be < or = 5%
        ' Negative Float
        'High Duration Ratio of # of tasks to total incomplete tasks should be < or = 5% 
        'Invalid Dates
        'Missing Resources
        'Missed Tasks Ratio of # of missed tasks to the "Baseline Count" should not exceed 5%
        'Critical Path Test: Project completion date (or other task/milestone) shows a very large negative total float ||  Revised EF for project complete
        'Critical Path Length Index (CPLI):Ratio of critical path length + total float to the critical path length should = 1 (>1 favorable; <1 unfav.)
        'Baseline Execution Index (BEI) 
        'Ratio measure of # of tasks completed to tasks that should have been completed (by the status date)


        'Logic (logic missing)
        intTaskCount = 0
        intTaskCountSelected = 0
        strDetails = ""

        For Each objTask In ActiveProject.Tasks

            If Not objTask Is Nothing Then
                strNote = ""

                If objTask.Summary = False And objTask.Active = True Then
                    intTaskCount = intTaskCount + 1
                    If Len(Trim(objTask.Predecessors)) = 0 Or Len(Trim(objTask.Successors)) = 0 Then
                        intTaskCountSelected = intTaskCountSelected + 1

                        If Len(Trim(objTask.Predecessors)) = 0 Then
                            strNote = strNote & "<td>X</td>"
                        Else
                            strNote = strNote & "<td>&nbsp;</td>"
                        End If

                        If Len(Trim(objTask.Successors)) = 0 Then
                            strNote = strNote & "<td>X</td>"
                        Else
                            strNote = strNote & "<td>&nbsp;</td>"
                        End If

                        strDetails = strDetails & "<tr><td>" & objTask.ID & "</td><td>" & objTask.Name & "</td>" & strNote & "</tr>"
                    End If
                End If

            End If
        Next

        strDetails = "<a href=""#"" onclick=""toggle_visibility('table1');"">Show Details</a><table><tbody id=""table1"" style=""display:none""><tr><th>ID</th><th>Name</th><th>Missing Pred</th><th>Missing Succ</th>" & strDetails & "</tbody></table>"


        strReturn = strReturn & "<tr><td>Missing Logic</td><td>" & intTaskCountSelected & " Tasks missing logic</td><td>2 (first and last task)</td><td>Any task that is missing a predecessor, or successor, or both. Shows how well (or poorly) the schedule is linked together. Every task except project or program start and finish has a predecessor and a successor (no dangling activities)</td><td>" & strDetails & "</td></tr>"

        'Lags	< or = 5%
        intTaskCount = 0
        intTaskCountSelected = 0
        strDetails = ""
        Dim objTaskDependencies As MSProject.TaskDependencies
        Dim objTaskDependency As MSProject.TaskDependency
        Dim intOneDay As Integer
        intOneDay = (60 * ActiveProject.HoursPerDay)

        For Each objTask In ActiveProject.Tasks

            If Not objTask Is Nothing Then
                strNote = ""
                If objTask.Summary = False And objTask.Active = True And objTask.PercentComplete < 100 Then
                    intTaskCount = intTaskCount + 1
                    objTaskDependencies = objTask.TaskDependencies
                    For Each objTaskDependency In objTaskDependencies
                        If objTaskDependency.To.ID = objTask.ID Then
                            If objTaskDependency.Lag <> 0 Then
                                intTaskCountSelected = intTaskCountSelected + 1
                                strNote = strNote & "Lag from task " & objTaskDependency.From.ID & " to task " & objTaskDependency.To.ID & "=" & objTaskDependency.Lag / intOneDay & "days.</br>"
                            End If
                        End If
                    Next
                    If strNote.Length > 0 Then
                        strDetails = strDetails & "<tr><td>" & objTask.ID & "</td><td>" & objTask.Name & "</td><td>" & strNote & "</td></tr>"
                    End If
                End If
            End If
        Next

        strDetails = "<a href=""#"" onclick=""toggle_visibility('table2');"">Show Details</a><table><tbody id=""table2"" style=""display:none""><tr><th>ID</th><th>Name</th><th>Lags</th>" & strDetails & "</tbody></table>"
        strReturn = strReturn & "<tr><td>Lags</td><td>" & Math.Round(((intTaskCountSelected / intTaskCount) * 100)) & "% (" & intTaskCountSelected & "/" & intTaskCount & ") tasks with lags</td><td>&lt; or = 5%</td><td>Any incomplete task that has a lag in its predecessor. To prevent adverse effects on critical path and subsequent analysis.</td><td>" & strDetails & "</td></tr>"

        'non FS Tasks
        intTaskCount = 0
        intTaskCountSelected = 0
        strDetails = ""
        strType = ""

        For Each objTask In ActiveProject.Tasks

            If Not objTask Is Nothing Then
                strNote = ""
                If objTask.Summary = False And objTask.Active = True And objTask.PercentComplete < 100 Then
                    intTaskCount = intTaskCount + 1
                    objTaskDependencies = objTask.TaskDependencies
                    For Each objTaskDependency In objTaskDependencies
                        If objTaskDependency.To.ID = objTask.ID Then

                            If objTaskDependency.Type <> MSProject.PjTaskLinkType.pjFinishToStart Then
                                intTaskCountSelected = intTaskCountSelected + 1
                                Select Case objTaskDependency.Type
                                    Case MSProject.PjTaskLinkType.pjStartToFinish
                                        strType = "SF"
                                    Case MSProject.PjTaskLinkType.pjStartToStart
                                        strType = "SS"
                                    Case MSProject.PjTaskLinkType.pjFinishToFinish
                                        strType = "FF"
                                End Select
                                strNote = strNote & "Link from task " & objTaskDependency.From.ID & " to task " & objTaskDependency.To.ID & "=" & strType & "</br>"
                            End If
                        End If
                    Next
                    If strNote.Length > 0 Then
                        strDetails = strDetails & "<tr><td>" & objTask.ID & "</td><td>" & objTask.Name & "</td><td>" & strNote & "</td></tr>"
                    End If
                End If
            End If
        Next

        strDetails = "<a href=""#"" onclick=""toggle_visibility('table3');"">Show Details</a><table><tbody id=""table3"" style=""display:none""><tr><th>ID</th><th>Name</th><th>Non-FS Links</th>" & strDetails & "</tbody></table>"
        strReturn = strReturn & "<tr><td>Non-FS Links</td><td>" & Math.Round(((intTaskCountSelected / intTaskCount) * 100)) & "% (" & intTaskCountSelected & "/" & intTaskCount & ") tasks with lags</td><td>&lt; 10%</td><td>All incomplete tasks that have predecessor(s) should state their relationship (eg., Finish-to-Start, or FS) to the predecessor. Identify non-typical (something other than FS) relationships. Non-typical tasks limited to <10 % of total tasks</td><td>" & strDetails & "</td></tr>"




        'Hard Constraints
        intTaskCount = 0
        intTaskCountSelected = 0
        strDetails = ""
        For Each objTask In ActiveProject.Tasks

            If Not objTask Is Nothing Then
                strNote = ""

                If objTask.Summary = False And objTask.Active = True And objTask.PercentComplete < 100 Then
                    intTaskCount = intTaskCount + 1
                    If (objTask.ConstraintDate.ToString <> "NA") Then
                        intTaskCountSelected = intTaskCountSelected + 1
                        Select Case objTask.ConstraintType
                            Case 0
                                strType = "As soon as possible (ASAP)"
                            Case 1
                                strType = "As late as possible (ALAP)"
                            Case 2
                                strType = "Must start on (MSO)"
                            Case 3
                                strType = "Must finish on (MFO)"
                            Case 4
                                strType = "Start no earlier than (SNET)"
                            Case 5
                                strType = "Start no later than (SNLT)"
                            Case 6
                                strType = "Finish no earlier than (FNET)"
                            Case 7
                                strType = "Finish no later than (FNLT)"

                        End Select

                        strDetails = strDetails & "<tr><td>" & objTask.ID & "</td><td>" & objTask.Name & "</td><td>" & FormatDateTime(CDate(objTask.ConstraintDate), 2) & "</td><td>" & strType & "</td></tr>"
                    End If
                End If

            End If
        Next

        strDetails = "<a href=""#"" onclick=""toggle_visibility('table4');"">Show Details</a><table><tbody id=""table4"" style=""display:none""><tr><th>ID</th><th>Name</th><th>Constraint Date</th><th>Constraint Type</th>" & strDetails & "</tbody></table>"


        strReturn = strReturn & "<tr><td>Hard Constraint Date</td><td>" & Math.Round(((intTaskCountSelected / intTaskCount) * 100)) & "% (" & intTaskCountSelected & "/" & intTaskCount & ") tasks with constraint dates</td><td>5%</td><td>An incomplete task that has any type of constraint	To identify constraints that prevent tasks from being logic-driven	Ratio of ""Hard Constraint"" tasks to total incomplete tasks should not exceed 5%</td><td>" & strDetails & "</td></tr>"



        'High FLoat
        intTaskCount = 0
        intTaskCountSelected = 0
        strDetails = ""
        For Each objTask In ActiveProject.Tasks

            If Not objTask Is Nothing Then
                strNote = ""

                If objTask.Summary = False And objTask.Active = True And objTask.PercentComplete < 100 Then

                    intTaskCount = intTaskCount + 1
                    If (objTask.TotalSlack / intOneDay > 44) Then
                        intTaskCountSelected = intTaskCountSelected + 1

                        strDetails = strDetails & "<tr><td>" & objTask.ID & "</td><td>" & objTask.Name & "</td><td>" & objTask.TotalSlack / intOneDay & " days</td></tr>"
                    End If
                End If

            End If
        Next

        strDetails = "<a href=""#"" onclick=""toggle_visibility('table5');"">Show Details</a><table><tbody id=""table5"" style=""display:none""><tr><th>ID</th><th>Name</th><th>Total Float/Slack</th>" & strDetails & "</tbody></table>"


        strReturn = strReturn & "<tr><td>High Float</td><td>" & Math.Round(((intTaskCountSelected / intTaskCount) * 100)) & "% (" & intTaskCountSelected & "/" & intTaskCount & ") tasks with high float</td><td>< or = 5%</td><td>An incomplete task with float greater than 44 working days (2 months) Another check for a missing predecessor or successor; to check the stability and logic of the network	Ratio of # of tasks with high float to total incomplete tasks; should be < or = 5%</td><td>" & strDetails & "</td></tr>"

        'Negative FLoat
        intTaskCount = 0
        intTaskCountSelected = 0
        strDetails = ""
        For Each objTask In ActiveProject.Tasks

            If Not objTask Is Nothing Then
                strNote = ""

                If objTask.Summary = False And objTask.Active = True And objTask.PercentComplete < 100 Then
                    intTaskCount = intTaskCount + 1
                    If (objTask.TotalSlack / intOneDay < 0) Then
                        intTaskCountSelected = intTaskCountSelected + 1

                        strDetails = strDetails & "<tr><td>" & objTask.ID & "</td><td>" & objTask.Name & "</td><td>" & objTask.TotalSlack / intOneDay & " days</td></tr>"
                    End If
                End If

            End If
        Next

        strDetails = "<a href=""#"" onclick=""toggle_visibility('table6');"">Show Details</a><table><tbody id=""table6"" style=""display:none""><tr><th>ID</th><th>Name</th><th>Total Float/Slack</th>" & strDetails & "</tbody></table>"


        strReturn = strReturn & "<tr><td>Negative Float</td><td>" & intTaskCountSelected & " tasks with negative float</td><td>None</td><td>An incomplete task with float less than 0 working days	Identify tasks that or delaying completion of one or more milestones</td><td>" & strDetails & "</td></tr>"


        'Invalid Dates - late date nd future actuals
        intTaskCount = 0
        intTaskCountSelected = 0
        intTaskCountSelectedTemp = 0
        strDetails = ""
        For Each objTask In ActiveProject.Tasks

            If Not objTask Is Nothing Then
                strNote = ""

                If objTask.Summary = False And objTask.Active = True And objTask.PercentComplete < 100 Then
                    intTaskCount = intTaskCount + 1
                    If (objTask.ActualStart.ToString <> "NA" Or objTask.ActualFinish.ToString <> "NA" Or objTask.Start <= dtmStatusDate) Then

                        If objTask.Start <= dtmStatusDate Then
                            If objTask.Start <= dtmStatusDate And objTask.PercentComplete = 0 Then
                                strNote = strNote & "<strong>Status</strong>- Delayed Start: Start Task or push out start date. <br/>"
                                intTaskCountSelectedTemp = 1
                            ElseIf objTask.Finish <= dtmStatusDate And objTask.PercentComplete < 100 Then
                                strNote = strNote & "<strong>Status</strong>- Missed Tasks: Set % Complete = 100% or increase duration. <br/>"
                                intTaskCountSelectedTemp = 1
                            End If
                        End If



                        If (objTask.ActualStart.ToString <> "NA") Then
                            If (objTask.ActualStart > dtmStatusDate) Then
                                strNote = strNote & "<strong>Status</strong>- Future Actual Date: Actual Start date is in the future ( " & FormatDateTime(CDate(objTask.ActualStart), 2) & " ). Set to date before or on Status Date. (Project Quality Best Practice)<br/>"
                                intTaskCountSelectedTemp = 1
                            End If
                        End If
                        If (objTask.ActualFinish.ToString <> "NA") Then
                            If (objTask.ActualFinish > dtmStatusDate) Then
                                strNote = strNote & "<strong>Status</strong>- Future Actual Date: Actual Finish date is in the future ( " & FormatDateTime(CDate(objTask.ActualFinish), 2) & " ). Set to date before or on Status Date.(Project Quality Best Practice)<br/>"
                                intTaskCountSelectedTemp = 1
                            End If
                        End If

                        If intTaskCountSelectedTemp > 0 Then ' one of the Invalid Date conditions exists
                            intTaskCountSelected = intTaskCountSelected + 1
                            intTaskCountSelectedTemp = 0
                        End If

                        strDetails = strDetails & "<tr><td>" & objTask.ID & "</td><td>" & objTask.Name & "</td><td>" & CheckDate(objTask.Start) & "</td><td>" & CheckDate(objTask.Finish) & "</td><td>" & CheckDate(objTask.ActualStart) & "</td><td>" & CheckDate(objTask.ActualFinish) & "</td><td>" & strNote & "</td></tr>"
                    End If
                End If

            End If
        Next

        strDetails = "<a href=""#"" onclick=""toggle_visibility('table6');"">Show Details</a><table><tbody id=""table6"" style=""display:none""><tr><th>ID</th><th>Name</th><th>Start</th><th>Finish</th><th>Actual Start</th><th>Actual Finish</th><th>Note</th>" & strDetails & "</tbody></table>"


        strReturn = strReturn & "<tr><td>Invalid Date</td><td>" & Math.Round(((intTaskCountSelected / intTaskCount) * 100)) & "% (" & intTaskCountSelected & "/" & intTaskCount & ") tasks with invalid dates</td><td>5%</td><td>Tasks that are supposed to have been completed (prior to the status date) with actual or forecast finishes after the baseline date, OR have a finish variance > zero. Helps identify how well (or poorly) the schedule is meeting the baseline plan. Ratio of # of missed tasks to the ""Baseline Count"" should not exceed 5%. (Project Quality Best Practice)</td><td>" & strDetails & "</td></tr>"


        'Missed Dates Dates -baseline  slips
        intTaskCount = 0
        intTaskCountSelected = 0
        intTaskCountSelectedTemp = 0
        strDetails = ""
        For Each objTask In ActiveProject.Tasks

            If Not objTask Is Nothing Then
                strNote = ""

                If objTask.Summary = False And objTask.Active = True And objTask.PercentComplete < 100 Then
                    intTaskCount = intTaskCount + 1
                    If (objTask.ActualStart.ToString <> "NA" Or objTask.ActualFinish.ToString <> "NA" Or objTask.Start <= dtmStatusDate) Then

                        If objTask.BaselineStart.ToString <> "NA" Then    'Application.DateValue("NA")
                            If objTask.Start > objTask.BaselineStart And objTask.PercentComplete < 100 Then
                                'intDelta = Application.DateDifference(dtmStart, dtmBaselineStart, ActiveProject.Calendar) / (60 * ActiveProject.HoursPerDay)
                                intDelta = DateDiff("d", objTask.BaselineStart, objTask.Start)
                                strNote = strNote & "<strong>Health</strong>- Slips: Start date is " & intDelta & " days late<br/>"
                            End If
                        End If
                        If objTask.BaselineFinish.ToString <> "NA" Then
                            If objTask.Finish > objTask.BaselineFinish And objTask.PercentComplete < 100 Then
                                'intDelta = Application.DateDifference(dtmFinish, dtmBaselineFinish, ActiveProject.Calendar) / (60 * ActiveProject.HoursPerDay)
                                intDelta = DateDiff("d", objTask.BaselineFinish, objTask.Finish)
                                strNote = strNote & "<strong>Health</strong>- Slips: Finish date is " & intDelta & " days late<br/>"
                            End If
                        End If




                        If intTaskCountSelectedTemp > 0 Then ' one of the Invalid Date conditions exists
                            intTaskCountSelected = intTaskCountSelected + 1
                            intTaskCountSelectedTemp = 0
                        End If

                        strDetails = strDetails & "<tr><td>" & objTask.ID & "</td><td>" & objTask.Name & "</td><td>" & CheckDate(objTask.Start) & "</td><td>" & CheckDate(objTask.Finish) & "</td><td>" & CheckDate(objTask.ActualStart) & "</td><td>" & CheckDate(objTask.ActualFinish) & "</td><td>" & strNote & "</td></tr>"
                    End If
                End If

            End If
        Next

        strDetails = "<a href=""#"" onclick=""toggle_visibility('table7');"">Show Details</a><table><tbody id=""table7"" style=""display:none""><tr><th>ID</th><th>Name</th><th>Start</th><th>Finish</th><th>Actual Start</th><th>Actual Finish</th><th>Note</th>" & strDetails & "</tbody></table>"


        strReturn = strReturn & "<tr><td>Invalid Date</td><td>" & Math.Round(((intTaskCountSelected / intTaskCount) * 100)) & "% tasks with invalid dates</td><td>5%</td><td>Tasks that are supposed to have been completed (prior to the status date) with actual or forecast finishes after the baseline date, OR have a finish variance > zero. Helps identify how well (or poorly) the schedule is meeting the baseline plan. Ratio of # of missed tasks to the ""Baseline Count"" should not exceed 5%. (Project Quality Best Practice)</td><td>" & strDetails & "</td></tr>"


        'Missing Resources
        intTaskCount = 0
        intTaskCountSelected = 0
        strDetails = ""
        For Each objTask In ActiveProject.Tasks

            If Not objTask Is Nothing Then
                strNote = ""

                If objTask.Summary = False And objTask.Active = True And objTask.PercentComplete < 100 Then
                    intTaskCount = intTaskCount + 1
                    If objTask.Resources.Count = 0 And objTask.Duration > 0 Then
                        intTaskCountSelected = intTaskCountSelected + 1
                        strNote = strNote & "<td>X</td>"
                        strDetails = strDetails & "<tr><td>" & objTask.ID & "</td><td>" & objTask.Name & "</td>" & strNote & "</tr>"
                    End If
                End If

            End If
        Next

        strDetails = "<a href=""#"" onclick=""toggle_visibility('table8');"">Show Details</a><table><tbody id=""table8"" style=""display:none""><tr><th>ID</th><th>Name</th><th>Missing Resources</th>" & strDetails & "</tbody></table>"


        strReturn = strReturn & "<tr><td>Missing Resources</td><td>" & intTaskCountSelected & " Tasks missing resources</td><td>All tasus with a duration of one or more days should have resources.</td><td>Any incomplete task that has resources (hours/$) assigned.	Verifies that all tasks with durations of 1 or more days have $ or resources assigned.</td><td>" & strDetails & "</td></tr>"

        strReturn = strReturn & "</table>" & vbCrLf


        strPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments

        strDTM = FormatDateTime(Now(), vbShortDate) & "-" & FormatDateTime(Now(), vbShortTime)
        strDTM = Replace(strDTM, "/", "-")
        strDTM = Replace(strDTM, ":", "")
        strPath = strPath & "\ProjectQualityReport-" & strDTM & ".htm"
        If My.Settings.CheckboxExcelValue = True Then
            strPath = strPath & ".xls"
            strCSS = "td, th { border: 1px solid }"
        Else
            strCSS = My.Resources.TableCSS
        End If

        strReturn = "<html><head>" & My.Resources.TableJS & "<title>" & strTitle & "</title><style>" & strCSS & "</style></head><body>" & strReturn & "</body>"


        Try
            My.Computer.FileSystem.WriteAllText(strPath, strReturn, False)
            Process.Start("file:\\" & strPath)
        Catch fileException As Exception
            Throw fileException
        End Try



    End Sub


    ''' <summary>
    ''' Checks the enterprise task fields.
    ''' </summary>
    ''' <param name="objTask">The object task.</param>
    ''' <returns></returns>
    Function CheckEnterpriseTaskFields(ByRef objTask As MSProject.Task) As String

        Dim strValue, strReturn As Object
        strReturn = ""
        If Not My.Settings.EnterpriseTaskFields Is Nothing Then
            For Each strItem As String In My.Settings.EnterpriseTaskFields
                If Len(Trim(strItem)) > 0 Then
                    strValue = getEnterpriseTaskCFValue(objTask.ID, strItem)
                    If Len(Trim(strValue)) = 0 Or InStr(strValue, "xxx") > 0 Then
                        strReturn = strReturn & "<strong>PMO</strong>- Field " & strItem & " is blank or default. Please enter value.<br/>"
                    End If
                End If
            Next
        End If

        CheckEnterpriseTaskFields = strReturn

    End Function


    ''' <summary>
    ''' Exports the report.
    ''' </summary>
    ''' <param name="colReport">The col report.</param>
    ''' <param name="strFooter">The string footer.</param>
    Sub ExportReport(ByRef strReportType As String, ByRef colReport As System.Collections.Generic.SortedDictionary(Of String, String), ByVal strFooter As String)
        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject
        Dim strReturn, strDisclaimer, strClass, strPath, strCSS, strDTM As String
        Dim objString As New StringBuilder
        strReturn = ""
        strDisclaimer = ""
        strClass = ""


        strPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments

        strDTM = FormatDateTime(Now(), vbShortDate) & "-" & FormatDateTime(Now(), vbShortTime)
        strDTM = Replace(strDTM, "/", "-")
        strDTM = Replace(strDTM, ":", "")
        strPath = strPath & "\" & strReportType & "Report-" & strDTM & ".htm"
        If My.Settings.CheckboxExcelValue = True Then
            strPath = strPath & ".xls"
            strCSS = "td, th { border: 1px solid }"
        Else
            strCSS = My.Resources.TableCSS
        End If

        objString.Append("<html><head><title>" & strReportType & " Report for " & ActiveProject.Name & " Report as of " & FormatDateTime(FutureDate(My.Settings.StatusDay), vbShortDate).ToString & "</title><style>" & strCSS & "</style></head><body>")

        If colReport.Count > 0 Then

            objString.Append("<table>")

            objString.Append("<tr><th colspan='9'>" & strReportType & " Report for " & ActiveProject.Name & " Report as of " & FormatDateTime(FutureDate(My.Settings.StatusDay), vbShortDate).ToString & "</th></tr>")

            For Each kvp As KeyValuePair(Of String, String) In colReport
                If Len(kvp.Key) > 0 Then

                    objString.Append("<tr class='selected'><td colspan='9'>" & kvp.Key & "</td></tr><tr><th>ID</th><th>Task</th><th>Responsible</th><th>% Comp</th><th>Duration</th><th>Start</th><th>Finish</th><th>Float</th><th>Note</th></tr>" & vbCrLf)
                    'objString.Append("<tr " & strClass & ">")
                    objString.Append(kvp.Value & vbCrLf)
                    ' objString.Append("</tr>")
                End If
            Next kvp
        End If
        objString.Append("<tr><th colspan='9'>&nbsp;</th></tr>" & vbCrLf)
        objString.Append("</tbody></table>" & vbCrLf & strFooter & "</body>")

        Try
            My.Computer.FileSystem.WriteAllText(strPath, objString.ToString, False)
            Process.Start("file:\\" & strPath)
        Catch fileException As Exception
            Throw fileException
        End Try

    End Sub



    ''' <summary>
    ''' Returs a future friday date.
    ''' </summary>
    ''' <param name="DayCount">The day count.</param>
    ''' <returns></returns>
    Function FutureFriday(ByVal DayCount As Integer) As Date
        Dim dtmNew As Date
        dtmNew = DateAdd("d", DayCount, Now())
        FutureFriday = DateAdd("d", (6 - DatePart("w", dtmNew)), dtmNew)
    End Function

    ''' <summary>
    ''' Returs a future friday date.
    '''  TODO: enable negative valuesfor previous days
    ''' </summary>
    ''' <param name="intDay">The int day.</param>
    ''' <returns></returns>
    Function FutureDate(ByVal intDay As Integer) As Date

        Dim dtmNew As Date = Date.Today()
        If intDay > 0 And intDay < 8 Then
            Do While Weekday(dtmNew) <> intDay
                dtmNew = DateAdd("d", 1, dtmNew)
            Loop
        Else
            dtmNew = DateAdd("d", 1, Date.Today())
        End If
        FutureDate = dtmNew
    End Function


    ''' <summary>
    ''' Returns a string for the number of days the task starts in rounded to the neerest 7 days or is task is ongoing or past due.
    ''' </summary>
    ''' <param name="dtmStartDate">The Task Start date.</param>
    ''' <param name="dtmFinishDate">The Task Finish date.</param>
    ''' <param name="dtmStatusDate">The status date.</param>
    ''' <param name="intStringLength">Length of the desired string.</param>
    ''' <param name="intProgress">The int task progress percent complete.</param>
    ''' <returns></returns>
    Function startsIn(ByVal dtmStartDate As Date, ByVal dtmFinishDate As Date, ByVal dtmStatusDate As Date, ByVal intStringLength As Integer, ByVal intProgress As Integer) As String
        Dim iDiffStart, iDiffFinish, iDiffStartWeek, iDiffFinishWeek As Integer
        Dim sDiff

        iDiffStart = Math.Ceiling(DateDiff("d", dtmStatusDate, dtmStartDate))
        iDiffStartWeek = Math.Round(iDiffStart / 7) * 7
        iDiffFinish = Math.Ceiling(DateDiff("d", dtmStatusDate, dtmFinishDate))
        iDiffFinishWeek = Math.Round(iDiffFinish / 7) * 7

        If (intProgress = 0 And iDiffStart < 0) Or (iDiffFinish < 0) Then
            startsIn = ">>>Past Due"
        ElseIf iDiffStart > 0 Then
            sDiff = PadZero(iDiffStartWeek, intStringLength)
            startsIn = "Starting in " & sDiff & " Days"
        Else
            startsIn = "Ongoing"
        End If

    End Function




    ''' <summary>
    ''' Pads the string with zeros.
    ''' </summary>
    ''' <param name="strInput">The string input.</param>
    ''' <param name="iLength">Length of the i.</param>
    ''' <returns></returns>
    Function PadZero(ByVal strInput As String, iLength As Integer) As String

        While Len(strInput) < iLength
            strInput = "0" & strInput
        End While

        PadZero = strInput
    End Function



    ''' <summary>
    ''' Gets the Enterprise field value.
    ''' </summary>
    ''' <param name="strEnterpriseFieldName">Name of the string enterprise field.</param>
    ''' <param name="varType">Type of the variable.</param>
    ''' <returns></returns>
    Function GetField(ByVal strEnterpriseFieldName As Object, ByVal varType As Object) As Object
        Dim fieldValue As Object
        Dim fldField As String

        fieldValue = vbNull
        If varType = Microsoft.Office.Interop.MSProject.PjFieldType.pjProject Then
            fldField = Globals.ThisAddIn.Application.ProjectSummaryTask.GetField(Globals.ThisAddIn.Application.FieldNameToFieldConstant(strEnterpriseFieldName, Microsoft.Office.Interop.MSProject.PjFieldType.pjProject))
            fieldValue = Globals.ThisAddIn.Application.ProjectSummaryTask.GetField(fldField)
        End If

        GetField = fieldValue
        Exit Function

Err_Init:
        MsgBox("GetField" & Err.Number & " - " & Err.Description)
    End Function

    ''' <summary>
    ''' Gets the Enterprise Custom Project Field value.
    ''' TODO:  Fix error handling
    ''' </summary>
    ''' <param name="strEnterpriseFieldName">Name of the string enterprise field.</param>
    ''' <returns></returns>
    Function getEnterpriseProjectCFValue(ByVal strEnterpriseFieldName As Object) As String
        On Error GoTo Err_Init
        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        Dim projectField As Long

        projectField = Globals.ThisAddIn.Application.FieldNameToFieldConstant(strEnterpriseFieldName, Microsoft.Office.Interop.MSProject.PjFieldType.pjProject)
        getEnterpriseProjectCFValue = ActiveProject.ProjectSummaryTask.GetField(projectField)
        Exit Function

Err_Init:
        Err.Clear()
    End Function

    ''' <summary>
    ''' Gets the Enterprise Custom Task Field value.
    ''' TODO:  Fix error handling
    ''' </summary>
    ''' <param name="intTaskID">The int task identifier.</param>
    ''' <param name="strEnterpriseFieldName">Name of the string enterprise field.</param>
    ''' <returns></returns>
    Function getEnterpriseTaskCFValue(ByVal intTaskID As Long, ByVal strEnterpriseFieldName As String) As String
        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject
        getEnterpriseTaskCFValue = ""


        Try
            getEnterpriseTaskCFValue = ActiveProject.Tasks(intTaskID).GetField(Globals.ThisAddIn.Application.FieldNameToFieldConstant(strEnterpriseFieldName, Microsoft.Office.Interop.MSProject.PjFieldType.pjTask))
        Catch ex As Exception
            getEnterpriseTaskCFValue = ""
        End Try

        Return getEnterpriseTaskCFValue

        Exit Function

    End Function

    ''' <summary>
    ''' Gets the Enterprise Custom Project Resource Field value.
    ''' </summary>
    ''' <param name="strEnterpriseFieldName">Name of the string enterprise field.</param>
    ''' <param name="strResourceName">Name of the string resource.</param>
    ''' <returns></returns>
    Function getEnterpriseResourceCFValue(ByVal strEnterpriseFieldName As Object, ByVal strResourceName As Object) As String
        On Error GoTo Err_Init

        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject
        getEnterpriseResourceCFValue = ActiveProject.Resources(strResourceName).GetField(Globals.ThisAddIn.Application.ActiveProject.FieldNameToFieldConstant(strEnterpriseFieldName, Microsoft.Office.Interop.MSProject.PjFieldType.pjResource))
        Exit Function

Err_Init:
        MsgBox("getEnterpriseResourceCFValue" & Err.Number & " - " & Err.Description)
    End Function

    ''' <summary>
    ''' Finds the enterprise fields. 
    ''' TODO: Finish and present to user.
    ''' </summary>
    Private Sub FindEnterpriseFields()
        On Error GoTo Err_Init
        Dim i, x, y As Long
        Dim strName As String

        x = 190873500    '190873600 Enterprise field Constant Range
        y = 190939135

        Dim c As Long
        For i = x To y
            If Not Globals.ThisAddIn.Application.FieldConstantToFieldName(i) = vbNull Then
                If Globals.ThisAddIn.Application.FieldConstantToFieldName(i) <> "" And Globals.ThisAddIn.Application.FieldConstantToFieldName(i) <> "<Unavailable>" Then
                    strName = Globals.ThisAddIn.Application.FieldConstantToFieldName(i)
                    'Debug.Print strName & "(" & i & ")" & "=" & getEnterpriseProjectCFValue(strName)
                End If
            End If
            'c = FieldNameToFieldConstant("Text" & i, pjTask) ' get constant of custom field by name
            'Debug.Print i & ". Rename title of Text" & i
            'Debug.Print "   Name of Text" & i; " is '" & Globals.ThisAddIn.Application.FieldConstantToFieldName(c) & "'"
            'CustomFieldRename FieldID:=c, newName:="Titel of Text " & i  'Rename/set custom field title
            'Debug.Print "   Title of Text" & i; " is '" & CustomFieldGetName(c) & "'" ' get title of custom field
        Next
        Exit Sub

Err_Init:
        MsgBox("EnterpriseFields" & Err.Number & " - " & Err.Description)
        'Resume Next
    End Sub

    Function StripTags(ByVal html As String) As String
        ' Remove HTML tags.
        Return Regex.Replace(html, "<.*?>", "")
    End Function


    ''' <summary>
    ''' Sets the Text field that has been identified to be used for recursive search.
    ''' </summary>
    ''' <param name="objTask">The object task.</param>
    ''' <param name="strValue">if set to <c>true</c> marks the field true.</param>
    Private Sub SetText(ByRef objTask As MSProject.Task, ByRef strValue As String)
        Dim strFieldName As String = My.Settings.StatusTextField
        Select Case strFieldName
            Case "Text1"
                objTask.Text1 = strValue
            Case "Text2"
                objTask.Text2 = strValue
            Case "Text3"
                objTask.Text3 = strValue
            Case "Text4"
                objTask.Text4 = strValue
            Case "Text5"
                objTask.Text5 = strValue
            Case "Text6"
                objTask.Text6 = strValue
            Case "Text7"
                objTask.Text7 = strValue
            Case "Text8"
                objTask.Text8 = strValue
            Case "Text9"
                objTask.Text9 = strValue
            Case "Text10"
                objTask.Text10 = strValue
            Case "Text11"
                objTask.Text11 = strValue
            Case "Text12"
                objTask.Text12 = strValue
            Case "Text13"
                objTask.Text13 = strValue
            Case "Text14"
                objTask.Text14 = strValue
            Case "Text15"
                objTask.Text15 = strValue
            Case "Text16"
                objTask.Text16 = strValue
            Case "Text17"
                objTask.Text17 = strValue
            Case "Text18"
                objTask.Text18 = strValue
            Case "Text19"
                objTask.Text19 = strValue
            Case "Text20"
                objTask.Text20 = strValue
            Case "Text21"
                objTask.Text21 = strValue
            Case "Text22"
                objTask.Text22 = strValue
            Case "Text23"
                objTask.Text23 = strValue
            Case "Text24"
                objTask.Text24 = strValue
            Case "Text25"
                objTask.Text25 = strValue
            Case "Text26"
                objTask.Text26 = strValue
            Case "Text27"
                objTask.Text27 = strValue
            Case "Text28"
                objTask.Text28 = strValue
            Case "Text29"
                objTask.Text29 = strValue
            Case "Text30"
                objTask.Text30 = strValue
            Case Else
        End Select


    End Sub

    ''' <summary>
    ''' Gets the Text field that has been identified to be used for recursive search.
    ''' </summary>
    ''' <param name="objTask">The object task.</param>
    ''' <returns></returns>
    Private Function GetText(ByRef objTask As MSProject.Task) As String
        Dim strFieldName As String = My.Settings.StatusTextField
        Dim strReturn As String
        Select Case strFieldName
            Case "Text1"
                strReturn = objTask.Text1
            Case "Text2"
                strReturn = objTask.Text2
            Case "Text3"
                strReturn = objTask.Text3
            Case "Text4"
                strReturn = objTask.Text4
            Case "Text5"
                strReturn = objTask.Text5
            Case "Text6"
                strReturn = objTask.Text6
            Case "Text7"
                strReturn = objTask.Text7
            Case "Text8"
                strReturn = objTask.Text8
            Case "Text9"
                strReturn = objTask.Text9
            Case "Text10"
                strReturn = objTask.Text10
            Case "Text11"
                strReturn = objTask.Text11
            Case "Text12"
                strReturn = objTask.Text12
            Case "Text13"
                strReturn = objTask.Text13
            Case "Text14"
                strReturn = objTask.Text14
            Case "Text15"
                strReturn = objTask.Text15
            Case "Text16"
                strReturn = objTask.Text16
            Case "Text17"
                strReturn = objTask.Text17
            Case "Text18"
                strReturn = objTask.Text18
            Case "Text19"
                strReturn = objTask.Text19
            Case "Text20"
                strReturn = objTask.Text20
            Case "Text21"
                strReturn = objTask.Text21
            Case "Text22"
                strReturn = objTask.Text22
            Case "Text23"
                strReturn = objTask.Text23
            Case "Text24"
                strReturn = objTask.Text24
            Case "Text25"
                strReturn = objTask.Text25
            Case "Text26"
                strReturn = objTask.Text26
            Case "Text27"
                strReturn = objTask.Text27
            Case "Text28"
                strReturn = objTask.Text28
            Case "Text29"
                strReturn = objTask.Text29
            Case "Text30"
                strReturn = objTask.Text30
            Case Else
                strReturn = ""
        End Select
        GetText = strReturn

    End Function

    ''' <summary>
    ''' Adds the status view.
    ''' </summary>
    Public Sub AddStatusView()
        On Error GoTo Err_Init
        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        ActiveProject.Application.ViewApply(Name:="Gantt Chart", SinglePane:=True, Toggle:=True)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=6, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=0, WrapText:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="Name", Title:="Task Name", Width:=32, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=0, WrapText:=True)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="PercentComplete", Title:="", Width:=12, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, WrapText:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="Duration", Title:="", Width:=12, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, WrapText:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="Start", Title:="", Width:=12, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=0, WrapText:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="Finish", Title:="", Width:=12, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=0, WrapText:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="Predecessors", Title:="", Width:=12, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=0, WrapText:=False, ShowAddNewColumn:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="Successors", Title:="", Width:=12, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=0, WrapText:=False, ShowAddNewColumn:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="Work", Title:="", Width:=12, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, WrapText:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="Resource Names", Title:="", Width:=12, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, WrapText:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:="TotalSlack", Title:="", Width:=12, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, WrapText:=False)
        ActiveProject.Application.TableEditEx(Name:="PM_Status", TaskTable:=True, NewFieldName:=My.Settings.StatusTextField, Title:="Status", Width:=38, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, WrapText:=True)
        ActiveProject.Application.SetSplitBar(12)
        ActiveProject.Application.TimescaleEdit(MajorUnits:=2, MajorLabel:=9, Separator:=True, MajorUseFY:=True, TierCount:=1)
        ActiveProject.Application.ViewCopy("PM_Status")

        ActiveProject.Application.TableApply(Name:="PM_Status")
        'ActiveProject.CustomFieldRename(FieldID:=ActiveProject.pjCustomTaskDuration1, NewName:="Optimistic Duration")
Err_Init:
        Err.Clear() 'TODO:  Fix Debugging
    End Sub



    ''' <summary>
    ''' Filter Active Project 
    ''' </summary>
    Private Sub FilterStatus()
        Globals.ThisAddIn.Application.OutlineShowAllTasks()
        Globals.ThisAddIn.Application.FilterEdit(Name:="__Status", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:=My.Settings.StatusTextField, Test:="does not equal", Value:="", ShowInMenu:=False, ShowSummaryTasks:=False)
        Globals.ThisAddIn.Application.FilterApply(Name:="__Status")
    End Sub


End Module
