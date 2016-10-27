' Copyright (c) 2014 Trevor Lowing

'This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Imports Microsoft.Office.Tools.Ribbon

Public Class ReportRibbon

    ''' <summary>
    ''' Handles the Load event of the ReportRibbon control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonUIEventArgs"/> instance containing the event data.</param>
    Private Sub ReportRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        'change tab name to match project versions
        Dim intVersion As Integer

        intVersion = CInt(Globals.ThisAddIn.Application.Application.Version.ToString())

        If intVersion >= 15 Then
            TabReview.Label = "REVIEW"
            'MsgBox(Globals.ThisAddIn.Application.Application.Version.ToString())
        Else
            TabReview.Label = "Review"
        End If

        '' load from values saved in settings file
        ToggleButtonExcel.Checked = My.Settings.CheckboxExcelValue
        If ToggleButtonExcel.Checked Then
            ToggleButtonExcel.Label = "Export Excel"
            ToggleButtonExcel.OfficeImageId = "ExportExcel"
        Else
            ToggleButtonExcel.Label = "Export HTML"
            ToggleButtonExcel.OfficeImageId = "ExportHtmlDocument"
        End If

    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonStatusReport control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonStatusReport_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonStatusReport.Click
        ProjectReport("Status")
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonBaselineReport control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonBaselineReport_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonBaselineReport.Click
        ProjectReport("Baseline")
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonEnterpriseReport control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonEnterpriseReport_Click(sender As Object, e As RibbonControlEventArgs)
        ProjectReport("Enterprise")
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonRunPath control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonRunPath_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonRunPath.Click
        Dim MyPathFinder As New ClassPathFinder
        MyPathFinder.PathFinder(ToggleButtonPathTo.Checked, ToggleButtonPathFrom.Checked, DropDownPathFilter.SelectedItem.Label.ToString())
    End Sub


    ''' <summary>
    ''' Handles the Click event of the ToggleButtonExcel control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ToggleButtonExcel_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButtonExcel.Click
        '' save current values to settings file
        If ToggleButtonExcel.Checked Then
            ToggleButtonExcel.Label = "Export Excel"
            ToggleButtonExcel.OfficeImageId = "ExportExcel"
        Else
            ToggleButtonExcel.Label = "Export HTML"
            ToggleButtonExcel.OfficeImageId = "ExportHtmlDocument"
            'ExportHtmlDocument
        End If
        My.Settings.CheckboxExcelValue = ToggleButtonExcel.Checked
        '' save other values in similar way

        '' finally call Save to save the values
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonEditSettings control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonEditSettings_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonEditSettings.Click
        Dim objForm As New FormSettings
        objForm.ShowDialog()
    End Sub


    ''' <summary>
    ''' Handles the 1 event of the ButtonEnterpriseReport_Click control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonEnterpriseReport_Click_1(sender As Object, e As RibbonControlEventArgs) Handles ButtonEnterpriseReport.Click
        ProjectReport("Enterprise")
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonEditDefaults control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonEditDefaults_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonEditDefaults.Click
        Dim objForm As New FormDefaultSettings
        objForm.ShowDialog()
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonContact control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonContact_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonContact.Click
        Dim objForm As New FormContact
        objForm.ShowDialog()
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonFeatureRequest control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonFeatureRequest_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonFeatureRequest.Click
        System.Diagnostics.Process.Start("http://www.ProjectReviewTool.com/help.php")
    End Sub


    ''' <summary>
    ''' Handles the Click event of the ButtonProgressToNow control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
    Private Sub ButtonProgressToNow_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonProgressToNow.Click
        ProgressToNow()
    End Sub





    Private Sub ButtonProjectQualityReport_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonProjectQualityReport.Click
        GetProjectQualityReport()
    End Sub
End Class
