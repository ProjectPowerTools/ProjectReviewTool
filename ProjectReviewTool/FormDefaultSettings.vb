' Copyright (c) 2014 Trevor Lowing

'This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Imports System.Collections
Imports System.Windows.Forms
Imports System.Data

Public Class FormDefaultSettings

    ''' <summary>
    ''' Handles the Load event of the FormDefaultSettings control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub FormDefaultSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        ' Create the ToolTip and associate with the Form container.
        Dim toolTip1 As ToolTip = New ToolTip()

        ' Set up the delays for the ToolTip.
        toolTip1.AutoPopDelay = 5000
        toolTip1.InitialDelay = 1000
        toolTip1.ReshowDelay = 500
        ' toolTip1.IsBalloon = True
        ' Force the ToolTip text to be displayed whether or not the form is active.
        toolTip1.ShowAlways = True

        ' Set up the ToolTip text for the Button and Checkbox.
        toolTip1.SetToolTip(CheckBoxProjectAutoLinkTasks, WrapTooltipText("Automatically links sequential tasks when you cut, WrapText(move, WrapText(or insert tasks. Can cause problems by unintentionally linking tasks."))
        toolTip1.SetToolTip(CheckBoxProjectAutoSplitTasks, WrapTooltipText("Automatically splits tasks into parts for work complete and work remaining. The preferred method of showing a stop in work is to stop the first task, WrapText(add a task represnting the reason for the stoppage and then adding a new task to represent the remaining work. The fundamental concept is that the plan should reflect what it happening."))
        toolTip1.SetToolTip(CheckBoxProjectAutoTrack, WrapTooltipText("Automatically updates the work and costs of resources assigned to a task when the percent complete changes. This is a recommended option."))
        toolTip1.SetToolTip(CheckBoxProjectAutoAddResources, WrapTooltipText("New resources are automatically created as they are assigned. The recommended value is false when using Enterprise Resources so that only Enterprise Resources are used and the user is prompted if trying to add a non-Enterprise resource."))
        toolTip1.SetToolTip(CheckBoxProjectDefaultED, WrapTooltipText(" The Effort Driven field indicates whether the scheduling for the task is effort driven scheduling. When a task is effort driven, WrapText(Project keeps the total task work at its current value, WrapText(regardless of how many resources are assigned to the task. When new resources are assigned, WrapText(remaining work is distributed to them. Unless you are tracking resoures to the hour it is not recommended to use this option. "))
        toolTip1.SetToolTip(CheckBoxProjectHonorCons, WrapTooltipText("Indicates whether scheduling constraints take precedence over dependencies. The recommended value is True since constraints can be used to manage slack and represent deadlines."))
        toolTip1.SetToolTip(CheckBoxProjectNewTasksEst, WrapTooltipText("New tasks in open projects have estimated durations. Since estimated durations appear with a question mark next to durations it may be more appropriate for planners to intentianall mark a task as estimated to avoid unintentionally having estimated tasks."))
        toolTip1.SetToolTip(CheckBoxProjectSchedFromStart, WrapTooltipText("Project calculates the project schedule forward from the start date. False if the schedule is calculated backward from the finish date. This should only be unchecked when you are drafting a plan with a defined deadline date and/or you wish to baseline tothe project late dates."))


        DataGridViewTaskType.Rows.Add(2)
        DataGridViewTaskType.Rows(0).Cells(0).Value = "Fixed-units task"
        DataGridViewTaskType.Rows(0).Cells(1).Value = "Duration is recalculated."
        DataGridViewTaskType.Rows(0).Cells(2).Value = "Work is recalculated."
        DataGridViewTaskType.Rows(0).Cells(3).Value = "Duration is recalculated."
        DataGridViewTaskType.Rows(1).Cells(0).Value = "Fixed-work task"
        DataGridViewTaskType.Rows(1).Cells(1).Value = "Duration is recalculated."
        DataGridViewTaskType.Rows(1).Cells(2).Value = "Units are recalculated."
        DataGridViewTaskType.Rows(1).Cells(3).Value = "Duration is recalculated."
        DataGridViewTaskType.Rows(2).Cells(0).Value = "Fixed-duration task"
        DataGridViewTaskType.Rows(2).Cells(1).Value = "Work is recalculated."
        DataGridViewTaskType.Rows(2).Cells(2).Value = "Work is recalculated."
        DataGridViewTaskType.Rows(2).Cells(3).Value = "Units are recalculated."


        ComboBoxTaskType.SelectedIndex = ComboBoxTaskType.FindStringExact(My.Settings.DefaultTaskType)
        CheckBoxProjectAutoLinkTasks.Checked = My.Settings.Default_AutoLinkTasks
        CheckBoxProjectAutoSplitTasks.Checked = My.Settings.Default_AutoSplitTasks
        CheckBoxProjectAutoTrack.Checked = My.Settings.Default_AutoTrack
        CheckBoxProjectAutoAddResources.Checked = My.Settings.Default_AutoAddResources
        CheckBoxProjectDefaultED.Checked = My.Settings.Default_DefaultEffortDriven
        CheckBoxProjectHonorCons.Checked = My.Settings.Default_HonorConstraints
        CheckBoxProjectNewTasksEst.Checked = My.Settings.Default_NewTasksEstimated
        CheckBoxProjectSchedFromStart.Checked = My.Settings.Default_ScheduleFromStart

        CheckBoxProjectAutoLinkTasksA.Checked = ActiveProject.AutoLinkTasks
        CheckBoxProjectAutoSplitTasksA.Checked = ActiveProject.AutoSplitTasks
        CheckBoxProjectAutoTrackA.Checked = ActiveProject.AutoTrack
        CheckBoxProjectAutoAddResourcesA.Checked = ActiveProject.AutoAddResources
        CheckBoxProjectDefaultEDA.Checked = ActiveProject.DefaultEffortDriven
        CheckBoxProjectHonorConsA.Checked = ActiveProject.HonorConstraints
        CheckBoxProjectNewTasksEstA.Checked = ActiveProject.NewTasksEstimated
        CheckBoxProjectSchedFromStartA.Checked = ActiveProject.ScheduleFromStart

        If My.Settings.DefaultCalendar IsNot Nothing Then
            ComboBoxCalendar.Items.Clear()
            For Each objCalendar In ActiveProject.BaseCalendars
                ComboBoxCalendar.Items.Add(objCalendar.Name)
            Next objCalendar

            ComboBoxCalendar.SelectedIndex = ComboBoxCalendar.FindStringExact(My.Settings.DefaultCalendar)
        End If


        CheckSettings()

    End Sub

    ''' <summary>
    ''' Checks the settings.
    ''' </summary>
    Private Sub CheckSettings()
        My.Settings.Save()
        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        If ActiveProject.Calendar.Name <> My.Settings.DefaultCalendar Then
            ComboBoxCalendar.BackColor = Drawing.Color.Salmon
        Else
            ComboBoxCalendar.BackColor = Nothing
        End If
        Dim varTaskType As VariantType
        Select Case My.Settings.DefaultTaskType
            Case "Fixed Duration"
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedDuration
            Case "Fixed Duration Not Effort Driven"
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedDuration
            Case "Fixed Work"
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedWork
            Case "Fixed Units"
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedUnits
            Case Else
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedDuration
        End Select

        If ActiveProject.DefaultTaskType <> varTaskType Then
            ComboBoxTaskType.BackColor = Drawing.Color.Salmon
        Else
            ComboBoxTaskType.BackColor = Nothing
        End If

        If My.Settings.Default_AutoLinkTasks <> ActiveProject.AutoLinkTasks Then
            CheckBoxProjectAutoLinkTasks.BackColor = Drawing.Color.Salmon
        Else
            CheckBoxProjectAutoLinkTasks.BackColor = Nothing
        End If
        If My.Settings.Default_AutoSplitTasks <> ActiveProject.AutoSplitTasks Then
            CheckBoxProjectAutoSplitTasks.BackColor = Drawing.Color.Salmon
        Else
            CheckBoxProjectAutoSplitTasks.BackColor = Nothing
        End If
        If My.Settings.Default_AutoTrack <> ActiveProject.AutoTrack Then
            CheckBoxProjectAutoTrack.BackColor = Drawing.Color.Salmon
        Else
            CheckBoxProjectAutoTrack.BackColor = Nothing
        End If
        If My.Settings.Default_AutoAddResources <> ActiveProject.AutoAddResources Then
            CheckBoxProjectAutoAddResources.BackColor = Drawing.Color.Salmon
        Else
            CheckBoxProjectAutoAddResources.BackColor = Nothing
        End If
        If My.Settings.Default_DefaultEffortDriven <> ActiveProject.DefaultEffortDriven Then
            CheckBoxProjectDefaultED.BackColor = Drawing.Color.Salmon
        Else
            CheckBoxProjectDefaultED.BackColor = Nothing
        End If
        If My.Settings.Default_HonorConstraints <> ActiveProject.HonorConstraints Then
            CheckBoxProjectHonorCons.BackColor = Drawing.Color.Salmon
        Else
            CheckBoxProjectHonorCons.BackColor = Nothing
        End If
        If My.Settings.Default_NewTasksEstimated <> ActiveProject.NewTasksEstimated Then
            CheckBoxProjectNewTasksEst.BackColor = Drawing.Color.Salmon
        Else
            CheckBoxProjectNewTasksEst.BackColor = Nothing
        End If
        If My.Settings.Default_ScheduleFromStart <> ActiveProject.ScheduleFromStart Then
            CheckBoxProjectSchedFromStart.BackColor = Drawing.Color.Salmon
        Else
            CheckBoxProjectSchedFromStart.BackColor = Nothing
        End If


    End Sub



    ''' <summary>
    ''' Handles the FormClosing event of the FormDefaultSettings control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="FormClosingEventArgs"/> instance containing the event data.</param>
    Private Sub FormDefaultSettings_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBoxProjectAutoLinkTasks control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBoxProjectAutoLinkTasks_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxProjectAutoLinkTasks.CheckedChanged
        My.Settings.Default_AutoLinkTasks = CheckBoxProjectAutoLinkTasks.Checked
        CheckSettings()
    End Sub

    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBoxProjectAutoSplitTasks control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBoxProjectAutoSplitTasks_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxProjectAutoSplitTasks.CheckedChanged
        My.Settings.Default_AutoSplitTasks = CheckBoxProjectAutoSplitTasks.Checked
        CheckSettings()
    End Sub

    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBoxProjectAutoTrack control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBoxProjectAutoTrack_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxProjectAutoTrack.CheckedChanged
        My.Settings.Default_AutoTrack = CheckBoxProjectAutoTrack.Checked
        CheckSettings()
    End Sub

    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBoxProjectAutoAddResources control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBoxProjectAutoAddResources_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxProjectAutoAddResources.CheckedChanged
        My.Settings.Default_AutoAddResources = CheckBoxProjectAutoAddResources.Checked
        CheckSettings()
    End Sub

    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBoxProjectDefaultED control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBoxProjectDefaultED_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxProjectDefaultED.CheckedChanged
        My.Settings.Default_DefaultEffortDriven = CheckBoxProjectDefaultED.Checked
        CheckSettings()
    End Sub

    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBoxProjectHonorCons control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBoxProjectHonorCons_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxProjectHonorCons.CheckedChanged
        My.Settings.Default_HonorConstraints = CheckBoxProjectHonorCons.Checked
        CheckSettings()
    End Sub

    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBoxProjectNewTasksEst control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBoxProjectNewTasksEst_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxProjectNewTasksEst.CheckedChanged
        My.Settings.Default_NewTasksEstimated = CheckBoxProjectNewTasksEst.Checked
        CheckSettings()
    End Sub

    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBoxProjectSchedFromStart control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBoxProjectSchedFromStart_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxProjectSchedFromStart.CheckedChanged
        My.Settings.Default_ScheduleFromStart = CheckBoxProjectSchedFromStart.Checked
        CheckSettings()
    End Sub


    ''' <summary>
    ''' Fixes the settings.
    ''' </summary>
    Sub FixSettings()
        My.Settings.Save()
        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject

        Dim varTaskType As VariantType
        Select Case My.Settings.DefaultTaskType
            Case "Fixed Duration"
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedDuration
            Case "Fixed Duration Not Effort Driven"
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedDuration
            Case "Fixed Work"
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedWork
            Case "Fixed Units"
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedUnits
            Case Else
                varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedWork
        End Select

        ActiveProject.DefaultTaskType = varTaskType
        Dim objTask As MSProject.Task
        'must do this first since you cannot set the default to effort driven on the project level if there are effort driven tasks
        If CheckBoxApplyTaskTypeDefault.Checked Then
            For Each objTask In ActiveProject.Tasks
                If Not objTask Is Nothing Then
                    ' see if the task is in the status window or a full report was asked for
                    'ID, Name, Start Finish, Float, Percent, Owner, Owner Email, baselinestart
                    If objTask.Summary = False And objTask.PercentComplete < 100 Then
                        objTask.Type = varTaskType
                        If varTaskType = Microsoft.Office.Interop.MSProject.PjTaskFixedType.pjFixedDuration And My.Settings.DefaultTaskType = "Fixed Duration Not Effort Driven" Then
                            objTask.EffortDriven = False
                        End If
                    End If
                End If
            Next
        End If


        ActiveProject.AutoLinkTasks = My.Settings.Default_AutoLinkTasks
        ActiveProject.AutoSplitTasks = My.Settings.Default_AutoSplitTasks
        ActiveProject.AutoTrack = My.Settings.Default_AutoTrack
        ActiveProject.AutoAddResources = My.Settings.Default_AutoAddResources
        ActiveProject.DefaultEffortDriven = My.Settings.Default_DefaultEffortDriven
        ActiveProject.HonorConstraints = My.Settings.Default_HonorConstraints
        ActiveProject.NewTasksEstimated = My.Settings.Default_NewTasksEstimated
        ActiveProject.ScheduleFromStart = My.Settings.Default_ScheduleFromStart
        If ComboBoxCalendar.FindStringExact(My.Settings.DefaultCalendar) Then

            ActiveProject.Application.ProjectSummaryInfo(
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing,
                        My.Settings.DefaultCalendar,
                       Type.Missing,
                       Type.Missing,
                       Type.Missing)

        End If

        CheckBoxProjectAutoLinkTasksA.Checked = ActiveProject.AutoLinkTasks
        CheckBoxProjectAutoSplitTasksA.Checked = ActiveProject.AutoSplitTasks
        CheckBoxProjectAutoTrackA.Checked = ActiveProject.AutoTrack
        CheckBoxProjectAutoAddResourcesA.Checked = ActiveProject.AutoAddResources
        CheckBoxProjectDefaultEDA.Checked = ActiveProject.DefaultEffortDriven
        CheckBoxProjectHonorConsA.Checked = ActiveProject.HonorConstraints
        CheckBoxProjectNewTasksEstA.Checked = ActiveProject.NewTasksEstimated
        CheckBoxProjectSchedFromStartA.Checked = ActiveProject.ScheduleFromStart

        MsgBox("Done, Defaults have been Applied. Changes are not permanant until you click 'Save'.")
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonFixDefaultSettings control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ButtonFixDefaultSettings_Click(sender As Object, e As EventArgs) Handles ButtonFixDefaultSettings.Click
        FixSettings()
        CheckSettings()
    End Sub

    ''' <summary>
    ''' Handles the SelectedIndexChanged event of the ComboBoxCalendar control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ComboBoxCalendar_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxCalendar.SelectedIndexChanged
        My.Settings.DefaultCalendar = ComboBoxCalendar.SelectedItem.ToString
        CheckSettings()
    End Sub

    ''' <summary>
    ''' Handles the SelectedIndexChanged event of the ComboBoxTaskType control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ComboBoxTaskType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxTaskType.SelectedIndexChanged
        My.Settings.DefaultTaskType = ComboBoxTaskType.SelectedItem.ToString
        CheckSettings()
    End Sub


    ''' <summary>
    ''' Handles the LinkClicked event of the LinkLabelType control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="LinkLabelLinkClickedEventArgs"/> instance containing the event data.</param>
    Private Sub LinkLabelType_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabelType.LinkClicked
        System.Diagnostics.Process.Start("https://office.microsoft.com/en-us/project-help/change-the-task-type-project-uses-to-calculate-task-duration-HP010092039.aspx")
    End Sub



End Class