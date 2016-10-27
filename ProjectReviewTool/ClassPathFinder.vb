Imports System.Data
Imports System.Diagnostics

' Copyright (c) 2014 Trevor Lowing

'This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

''' <summary>
''' Utility Class for performing Critical Path Analysis
''' </summary>
Public Class ClassPathFinder
    Private ActiveProject As Microsoft.Office.Interop.MSProject.Project
    Private blnForward As Boolean
    Private intSelectedTaskID As Integer ' task ID of current task
    Private IsSummary As Boolean
    Private IsCritical As Boolean
    Private IsDriver As Boolean
    Private objTask As MSProject.Task
    Private colReport As New DataTable

    'http://www.dotnetperls.com/datatable-vbnet

    ''' <summary>
    ''' Flags the critical path with optionsfor showing path to and from.
    ''' </summary>
    ''' <param name="blnPathTo">if set to <c>true</c> shows path to.</param>
    ''' <param name="blnPathFrom">if set to <c>true</c> shows path from.</param>
    ''' <param name="strFilter">Options to filder the path for dviers and critical tasks.</param>
    Public Sub PathFinder(ByVal blnPathTo As Boolean, ByVal blnPathFrom As Boolean, ByVal strFilter As String)
        'set reference to current project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject
        'clear old filter
        ClearFlags()
        IsCritical = False

        If CheckSelection() Then 'task or milestone selected
            'check if only display drivers, citical or all
            If InStr(strFilter, "driv") > 0 Then
                IsDriver = True
            ElseIf InStr(strFilter, "crit") > 0 Then
                IsCritical = True
            End If
            If blnPathTo And blnPathFrom Then
                PathAll()
            ElseIf blnPathTo Then
                PathTo()
            ElseIf blnPathFrom Then
                PathFrom()
            End If
            CreateReport("Path")
            'FilterProject(blnShowSummary)
        End If
    End Sub


    ''' <summary>
    ''' Checks the currently selected task.
    ''' </summary>
    ''' <returns></returns>
    Private Function CheckSelection() As Boolean
        If Not Globals.ThisAddIn.Application.ActiveSelection.Tasks Is Nothing Then
            If Globals.ThisAddIn.Application.ActiveSelection.Tasks.Count <> 1 Then
                MsgBox("Please select ONE Task to begin.")
                CheckSelection = False
                Exit Function
            End If
            Dim objTask As MSProject.Task
            For Each objTask In Globals.ThisAddIn.Application.ActiveSelection.Tasks
                If objTask.Summary = True Then
                    MsgBox("You chose a summary task. Select a regular task or milestone")
                    CheckSelection = False
                    Exit Function
                ElseIf objTask.Active = False Then
                    MsgBox("You chose a inactive task. Select a active task or milestone")
                    CheckSelection = False
                    Exit Function
                End If
            Next
        End If
        CheckSelection = True

    End Function

    ''' <summary>
    ''' Traces the Path Only For Successor Tasks - blnForward set to true
    ''' </summary>
    Private Sub PathFrom()

        intSelectedTaskID = 0
        blnForward = True
        GetParent()
    End Sub

    ''' <summary>
    ''' Traces the Path Only  for Predecessor Tasks - blnForward set to false
    ''' </summary>
    Private Sub PathTo()
        intSelectedTaskID = 0
        blnForward = False
        GetParent()
    End Sub

    ''' <summary>
    '''  Traces the Path For All Logicallt Linked Tasks - performs one pass for all successors and then one for all predecessors
    ''' </summary>
    Private Sub PathAll()
        intSelectedTaskID = 0
        blnForward = True    ' mark successors
        GetParent()
        blnForward = False    ' mark predecessors
        GetParent()
    End Sub

    ''' <summary>
    ''' Clears the flag fields on all tasks by settign them to False.
    ''' </summary>
    Private Sub ClearFlags()
        Dim objTask As MSProject.Task
        For Each objTask In ActiveProject.Tasks
            If Not (objTask Is Nothing) Then
                If GetFlag(objTask) = True Then SetFlag(objTask, False)
            End If
        Next objTask
    End Sub

    ''' <summary>
    ''' Gets the parent/selectged task to start the recursive loop.
    ''' </summary>
    Private Sub GetParent()
        Dim objTask As MSProject.Task
        For Each objTask In Globals.ThisAddIn.Application.ActiveSelection.Tasks
            If Not (objTask Is Nothing) Then
                intSelectedTaskID = objTask.ID
                GetChildren(objTask)
            End If
        Next objTask
    End Sub

    ''' <summary>
    '''  Traces through all task predecessors or successors and sets the identified Flag Field to True.
    ''' </summary>
    ''' <param name="objTask">The object task.</param>
    Private Sub GetChildren(objTask As MSProject.Task)
        Dim objChildTask As MSProject.Task
        SetFlag(objTask, True)
        If blnForward Then
            For Each objChildTask In objTask.SuccessorTasks
                If GetFlag(objChildTask) <> True And objChildTask.Active Then
                    If IsCritical And Not IsDriver Then
                        If objChildTask.Critical = True Then
                            GetChildren(objChildTask)
                        End If

                    ElseIf IsDriver = True Then
                        If objChildTask.FreeSlack < 100 Then
                            GetChildren(objChildTask)
                        End If
                    Else
                        GetChildren(objChildTask)
                    End If
                End If
            Next objChildTask
        Else
            For Each objChildTask In objTask.PredecessorTasks
                If GetFlag(objChildTask) <> True And objChildTask.Active Then
                    If IsCritical And Not IsDriver Then
                        If objChildTask.Critical = True Then
                            GetChildren(objChildTask)
                        End If

                    ElseIf IsDriver = True Then
                        If objChildTask.FreeSlack < 100 Then
                            GetChildren(objChildTask)
                        End If
                    Else
                        GetChildren(objChildTask)
                    End If
                End If
            Next objChildTask
        End If
    End Sub

    ''' <summary>
    ''' Sets the flag field that has been identified to be used for recursive search.
    ''' </summary>
    ''' <param name="objTask">The object task.</param>
    ''' <param name="blnValue">if set to <c>true</c> marks the field true.</param>
    Private Sub SetFlag(ByRef objTask As MSProject.Task, ByRef blnValue As Boolean)
        Dim strFieldName As String = My.Settings.PathFlagField
        Select Case strFieldName
            Case "Flag1"
                objTask.Flag1 = blnValue
            Case "Flag2"
                objTask.Flag2 = blnValue
            Case "Flag3"
                objTask.Flag3 = blnValue
            Case "Flag4"
                objTask.Flag4 = blnValue
            Case "Flag5"
                objTask.Flag5 = blnValue
            Case "Flag6"
                objTask.Flag6 = blnValue
            Case "Flag7"
                objTask.Flag7 = blnValue
            Case "Flag8"
                objTask.Flag8 = blnValue
            Case "Flag9"
                objTask.Flag9 = blnValue
            Case "Flag10"
                objTask.Flag10 = blnValue
            Case "Flag11"
                objTask.Flag11 = blnValue
            Case "Flag12"
                objTask.Flag12 = blnValue
            Case "Flag13"
                objTask.Flag13 = blnValue
            Case "Flag14"
                objTask.Flag14 = blnValue
            Case "Flag15"
                objTask.Flag15 = blnValue
            Case "Flag16"
                objTask.Flag16 = blnValue
            Case "Flag17"
                objTask.Flag17 = blnValue
            Case "Flag18"
                objTask.Flag18 = blnValue
            Case "Flag19"
                objTask.Flag19 = blnValue
            Case "Flag20"
                objTask.Flag20 = blnValue
            Case Else
        End Select


    End Sub

    ''' <summary>
    ''' Gets the flag field that has been identified to be used for recursive search.
    ''' </summary>
    ''' <param name="objTask">The object task.</param>
    ''' <returns></returns>
    Private Function GetFlag(ByRef objTask As MSProject.Task) As Boolean
        Dim strFieldName As String = My.Settings.PathFlagField
        Dim blnReturn As Boolean
        Select Case strFieldName
            Case "Flag1"
                blnReturn = objTask.Flag1
            Case "Flag2"
                blnReturn = objTask.Flag2
            Case "Flag3"
                blnReturn = objTask.Flag3
            Case "Flag4"
                blnReturn = objTask.Flag4
            Case "Flag5"
                blnReturn = objTask.Flag5
            Case "Flag6"
                blnReturn = objTask.Flag6
            Case "Flag7"
                blnReturn = objTask.Flag7
            Case "Flag8"
                blnReturn = objTask.Flag8
            Case "Flag9"
                blnReturn = objTask.Flag9
            Case "Flag10"
                blnReturn = objTask.Flag10
            Case "Flag11"
                blnReturn = objTask.Flag11
            Case "Flag12"
                blnReturn = objTask.Flag12
            Case "Flag13"
                blnReturn = objTask.Flag13
            Case "Flag14"
                blnReturn = objTask.Flag14
            Case "Flag15"
                blnReturn = objTask.Flag15
            Case "Flag16"
                blnReturn = objTask.Flag16
            Case "Flag17"
                blnReturn = objTask.Flag17
            Case "Flag18"
                blnReturn = objTask.Flag18
            Case "Flag19"
                blnReturn = objTask.Flag19
            Case "Flag20"
                blnReturn = objTask.Flag20
            Case Else
                blnReturn = False
        End Select
        GetFlag = blnReturn

    End Function



    ''' <summary>
    ''' Creates the report.
    ''' </summary>
    ''' <param name="strReportType">The report name.</param>
    Public Sub CreateReport(strReportType As String)
        Dim objTask As MSProject.Task
        Dim strDtm, strPath, strColor, strCritical, strClass, strCSS As String
        Dim dblSlack As Double
        Dim intOneDay, intProjectDuration, intPercentofProjectDur, intDur As Integer


        Dim ActiveProject As Microsoft.Office.Interop.MSProject.Project
        ActiveProject = Globals.ThisAddIn.Application.ActiveProject
        intOneDay = (60 * ActiveProject.HoursPerDay)


        'ActiveProject.StatusDate

        intProjectDuration = DateDiff("d", ActiveProject.ProjectStart, ActiveProject.ProjectFinish)

        strPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments

        strDtm = FormatDateTime(Now(), vbShortDate) & "-" & FormatDateTime(Now(), vbShortTime)
        strDtm = Replace(strDtm, "/", "-")
        strDtm = Replace(strDtm, ":", "")
        strPath = strPath & "\" & strReportType & "Report-" & strDtm & ".htm"
        If My.Settings.CheckboxExcelValue = True Then
            strPath = strPath & ".xls"
            strCSS = "td, th { border: 1px solid }"
        Else
            strCSS = My.Resources.TableCSS
        End If
        'Dim content As String = My.Resources.mytextfile
        Dim objString As New StringBuilder

        objString.Append("<html><head><title>" & ActiveProject.Name & " Report as of " & FormatDateTime(FutureDate(My.Settings.StatusDay), vbShortDate).ToString & "</title><style>" & strCSS & "</style></head><body>")

        objString.Append("<table>")

        ' append headers
        objString.Append("<thead><tr>")
        objString.Append("</th></tr><tr><th>ID</th><th style=""width:10%"">Task</th><th  style=""width:10%"">Predecessors</th><th  style=""width:10%"">Successors</th><th style=""width:10%"">Responsible</th><th>% Comp</th><th>Duration</th><th>Start</th><th>Finish</th><th>Float</th><th>Timelime</th>")
        objString.Append("</tr></thead>")

        objString.Append("<tbody>")
        ' append rows

        For Each objTask In ActiveProject.Tasks
            If Not objTask Is Nothing Then
                ' see if the task is in the status window or a full report was asked for
                'ID, Name, Start Finish, Float, Percent, Owner, Owner Email, baselinestart
                If objTask.Summary = False And objTask.Active = True Then

                    If GetFlag(objTask) <> False Then ' task was flagged for report
                        dblSlack = objTask.TotalSlack / intOneDay

                        intPercentofProjectDur = (DateDiff("d", ActiveProject.ProjectStart, objTask.Start) / intProjectDuration) * 100
                        intDur = (DateDiff("d", objTask.Start, objTask.Finish) / intProjectDuration) * 100
                        If intDur < 3 Then 'duratiom is too short so make it 3% for visability
                            intDur = 3
                        End If
                        strColor = "black"
                        strCritical = ""
                        If objTask.Critical Then
                            strColor = "red"
                            strCritical = "*"
                        End If
                        strClass = ""

                        If objTask.ID = intSelectedTaskID Then
                            strClass = " class='selected'"
                        End If
                        objString.Append("<tr " & strClass & "><td style='color:" & strColor & "'>" & _
                                         objTask.ID & strCritical & "</td><td style='text-align: left;'>" & objTask.Name & "</td><td>" & objTask.Predecessors & _
                                         "&nbsp;</td><td>" & objTask.Successors & _
                                         "&nbsp;</td><td>" & objTask.ResourceNames & _
                                         "&nbsp;</td><td>" & objTask.PercentComplete & "%</td><td>" & objTask.Duration / intOneDay & "D</td><td>" & FormatDateTime(CDate(objTask.Start), 2) & "</td><td>" & _
                                         FormatDateTime(CDate(objTask.Finish), 2) & "</td><td>" & Math.Round(dblSlack) & "D</td><td style='text-align: left;white-space: nowrap;'><div style='display: inline-block;height:4px;width:" & intPercentofProjectDur & "%;'></div><div style='display: inline-block;height:8px;background-color:black;width:" & intDur & "%;'></div></td></tr>" & vbCrLf)

                    End If
                End If
            End If

        Next

        objString.Append("</tbody>")
        objString.Append("</table></body>")

        Try
            My.Computer.FileSystem.WriteAllText(strPath, objString.ToString, False)
            Process.Start("file:\\" & strPath)
        Catch fileException As Exception
            Throw fileException
        End Try
    End Sub

End Class
