' Copyright (c) 2014 Trevor Lowing

'This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Imports System.Net.Mail
Imports System.Net
Imports System.Collections
Imports System.IO
Imports ProjectReviewTool.VBFastTag.PosTag

Module ModuleHelper

    ''' <summary>
    ''' Posts the form.
    ''' </summary>
    ''' <param name="strURL">The string URL.</param>
    ''' <param name="colFields">The collection of fields.</param>
    ''' <returns></returns>
    Function PostForm(ByVal strURL As String, ByRef colFields As Dictionary(Of String, String), ByVal strFilePath As String) As String

        Using client As New Net.WebClient
            Dim reqparm As New Specialized.NameValueCollection
            For Each kvp As KeyValuePair(Of String, String) In colFields
                If Len(kvp.Key) > 0 Then
                    reqparm.Add(kvp.Key, kvp.Value)
                End If
            Next kvp

            Dim responsebytes = client.UploadValues(strURL, "POST", reqparm)
            Dim responsebody = (New Text.UTF8Encoding).GetString(responsebytes)

            If strFilePath.Length > 0 Then
                'https://msdn.microsoft.com/en-us/library/system.web.ui.webcontrols.fileupload.filebytes%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396
                client.UploadFile(strURL, strFilePath)

                'string response = Encoding.UTF8.GetString(responseBinary);


            End If
        End Using


        PostForm = "Thank you for your feedback."
    End Function



    ''' <summary>
    ''' Wraps the tooltip text.
    ''' </summary>
    ''' <param name="strTooltip">The string tooltip.</param>
    ''' <returns></returns>
    Function WrapTooltipText(ByVal strTooltip As String) As String
        Dim MaxStringLength As Int16 = 48

        If strTooltip.Length > MaxStringLength Then
            Dim indexOfSpace = strTooltip.IndexOf(" ", MaxStringLength - 1)
            If indexOfSpace <> -1 AndAlso indexOfSpace <> strTooltip.Length - 1 Then
                Dim firstString As String = strTooltip.Substring(0, indexOfSpace)
                Dim secondString As String = strTooltip.Substring(indexOfSpace)

                Return firstString & Chr(10) & WrapTooltipText(secondString)
            Else
                Return strTooltip
            End If
        Else
            Return strTooltip
        End If
    End Function

    Sub ProgressToNow()
        Dim intDelta As Integer
        Dim dtmStatusDate As Date
        dtmStatusDate = FutureDate(My.Settings.StatusDay)
        If Not Globals.ThisAddIn.Application.ActiveSelection.Tasks Is Nothing Then
            'If Globals.ThisAddIn.Application.ActiveSelection.Tasks.Count <> 1 Then
            'MsgBox("Please select ONE Task to begin.")
            'CheckSelection = False
            'Exit Function
            'End If
            Dim objTask As MSProject.Task
            For Each objTask In Globals.ThisAddIn.Application.ActiveSelection.Tasks
                If objTask.Summary = False Then
                    'if dtmStatusDate > objTask.Finish ask them to push, extend or complete

                    If dtmStatusDate > objTask.Finish Then
                    Else
                        intDelta = Globals.ThisAddIn.Application.DateDifference(dtmStatusDate, objTask.Finish, Globals.ThisAddIn.Application.ActiveProject.Calendar)
                        intDelta = 100 * ((objTask.Duration - (intDelta)) / objTask.Duration)
                        objTask.PercentComplete = intDelta
                    End If



                End If
            Next
        End If

    End Sub

    Sub CheckGrammar(ByVal strSentence As String)

        ' read the english lexicon data
        Dim lexicon = File.ReadAllText("Grammar\lexicon.txt")
        ' run the sample loop
        Dim ft = New FastTag(lexicon)
        Dim tagResult = ft.Tag(strSentence)
        For Each varTagResult As Object In tagResult
            Dim message = String.Format("[{0} {1}]", varTagResult.Word, varTagResult.PosTag)
            Console.WriteLine(message)
        Next

    End Sub


    ''' <summary>
    ''' Clears the current Filter.
    ''' </summary>
    Public Sub FilterRemoveAll()
        Globals.ThisAddIn.Application.OutlineShowAllTasks()
        Globals.ThisAddIn.Application.FilterApply(Name:="All Tasks")
    End Sub

    Public Function CheckDate(ByRef objDate As Object) As String
        Dim strReturn As String


        If (objDate.ToString <> "NA") Then
            strReturn = FormatDateTime(CDate(objDate), 2)
        Else
            strReturn = ""
        End If


        CheckDate = strReturn

    End Function

End Module
