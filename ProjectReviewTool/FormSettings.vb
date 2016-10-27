' Copyright (c) 2014 Trevor Lowing

'This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Imports System.Windows.Forms
Imports System.Data
Imports System.Collections.Specialized

Public Class FormSettings

    ''' <summary>
    ''' Handles the Load event of the FormSettings control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub FormSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBoxPathFlag.SelectedIndex = ComboBoxPathFlag.FindStringExact(My.Settings.PathFlagField)
        ComboBoxStatusTextField.SelectedIndex = ComboBoxStatusTextField.FindStringExact(My.Settings.StatusTextField)
        ComboBoxStatusDay.SelectedIndex = My.Settings.StatusDay
        ComboBoxStatusWeeks.SelectedIndex = ComboBoxStatusWeeks.FindStringExact(My.Settings.StatusWeeks)
        CheckBoxSetStatusDate.Checked = My.Settings.AutoSetStatusDate
        'repopulate lists
        If My.Settings.EnterpriseProjectFields IsNot Nothing Then
            ComboBoxEnterpriseProjectFields.Items.Clear()
            For Each strItem As String In My.Settings.EnterpriseProjectFields
                'If Not ComboBoxEnterpriseProjectFields.FindString(Item) Then
                ComboBoxEnterpriseProjectFields.Items.Add(strItem)
                'End If
            Next
        Else
            My.Settings.EnterpriseProjectFields = New StringCollection
        End If
        If My.Settings.EnterpriseTaskFields IsNot Nothing Then
            ComboBoxEnterpriseTaskFields.Items.Clear()
            For Each strItem As String In My.Settings.EnterpriseTaskFields
                ' If Not ComboBoxEnterpriseTaskFields.FindString(Item) Then
                ComboBoxEnterpriseTaskFields.Items.Add(strItem)
                'End If
            Next
        Else
            My.Settings.EnterpriseTaskFields = New StringCollection
        End If
        If My.Settings.CustomReports IsNot Nothing Then
            ComboBoxCustomReports.Items.Clear()
            For Each strItem As String In My.Settings.CustomReports
                'If Not ComboBoxCustomReports.FindString(Item) Then
                ComboBoxCustomReports.Items.Add(strItem)
                'End If
            Next
        Else
            My.Settings.CustomReports = New StringCollection
        End If

    End Sub


    ''' <summary>
    ''' Handles the SelectedIndexChanged event of the ComboBoxPathFlag control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ComboBoxPathFlag_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxPathFlag.SelectedIndexChanged
        My.Settings.PathFlagField = ComboBoxPathFlag.SelectedItem.ToString
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles the SelectedIndexChanged event of the ComboBoxStatusDay control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ComboBoxStatusDay_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxStatusDay.SelectedIndexChanged
        My.Settings.StatusDay = ComboBoxStatusDay.SelectedIndex
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles the SelectedIndexChanged event of the ComboBoxStatusWeeks control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ComboBoxStatusWeeks_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxStatusWeeks.SelectedIndexChanged
        My.Settings.StatusWeeks = ComboBoxStatusWeeks.SelectedItem.ToString
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonSave control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ButtonSave_Click(sender As Object, e As EventArgs)

        SaveComboLists()
        Me.Dispose()
    End Sub

    Public Sub SaveComboLists()

        If ComboBoxEnterpriseProjectFields.Items.Count > 0 Then
            My.Settings.EnterpriseProjectFields.Clear()
            For Each Item2 As String In ComboBoxEnterpriseProjectFields.Items
                My.Settings.EnterpriseProjectFields.Add(Item2)
            Next

        End If
        If ComboBoxEnterpriseTaskFields.Items.Count > 0 Then
            My.Settings.EnterpriseTaskFields.Clear()
            For Each Item2 As String In ComboBoxEnterpriseTaskFields.Items
                My.Settings.EnterpriseTaskFields.Add(Item2)
            Next
        End If
        If ComboBoxCustomReports.Items.Count > 0 Then
            My.Settings.CustomReports.Clear()
            For Each Item2 As String In ComboBoxCustomReports.Items
                My.Settings.CustomReports.Add(Item2)
            Next
        End If
        My.Settings.Save()
    End Sub


    ''' <summary>
    ''' Handles the Click event of the ButtonAddProjectField control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ButtonAddProjectField_Click(sender As Object, e As EventArgs) Handles ButtonAddProjectField.Click
        Dim itm As String
        itm = InputBox("Enter new item", "New Item")
        If itm.Trim <> "" Then AddElement(ComboBoxEnterpriseProjectFields, itm)
        SaveComboLists()
    End Sub

    ''' <summary>
    ''' Adds the element.
    ''' </summary>
    ''' <param name="ComboBox1">The combo box1.</param>
    ''' <param name="newItem">The new item.</param>
    Sub AddElement(ByRef ComboBox1 As ComboBox, ByVal newItem As String)
        Dim idx As Integer
        If ComboBox1.FindString(newItem) > 0 Then
            idx = ComboBox1.FindString(newItem)
        Else
            idx = ComboBox1.Items.Add(newItem)
        End If
        ComboBox1.SelectedIndex = idx
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonAddTaskField control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ButtonAddTaskField_Click(sender As Object, e As EventArgs) Handles ButtonAddTaskField.Click
        Dim itm As String
        itm = InputBox("Enter new item", "New Item")
        If itm.Trim <> "" Then AddElement(ComboBoxEnterpriseTaskFields, itm)
        SaveComboLists()
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonAddCustomReport control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ButtonAddCustomReport_Click(sender As Object, e As EventArgs) Handles ButtonAddCustomReport.Click
        Dim itm As String
        itm = InputBox("Enter new item", "New Item")
        If itm.Trim <> "" Then AddElement(ComboBoxCustomReports, itm)
        SaveComboLists()
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonProjectFieldDelete control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ButtonProjectFieldDelete_Click(sender As Object, e As EventArgs) Handles ButtonProjectFieldDelete.Click
        ComboBoxEnterpriseProjectFields.Items.Remove(ComboBoxEnterpriseProjectFields.SelectedItem)
        SaveComboLists()
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonTaskFieldDelete control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ButtonTaskFieldDelete_Click(sender As Object, e As EventArgs) Handles ButtonTaskFieldDelete.Click
        ComboBoxEnterpriseTaskFields.Items.Remove(ComboBoxEnterpriseTaskFields.SelectedItem.ToString)
        SaveComboLists()
    End Sub

    ''' <summary>
    ''' Handles the Click event of the ButtonCustomReportDelete control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ButtonCustomReportDelete_Click(sender As Object, e As EventArgs) Handles ButtonCustomReportDelete.Click
        ComboBoxCustomReports.Items.Remove(ComboBoxCustomReports.SelectedItem)
        SaveComboLists()
    End Sub


    ''' <summary>
    ''' Removes the dups.
    ''' </summary>
    ''' <param name="ComboBox1">The combo box.</param>
    Sub RemoveDups(ByRef ComboBox1 As ComboBox)

        For i As Int16 = 0 To ComboBox1.Items.Count - 2
            For j As Int16 = ComboBox1.Items.Count - 1 To i + 1 Step -1
                If ComboBox1.Items(i).ToString = ComboBox1.Items(j).ToString Then
                    ComboBox1.Items.RemoveAt(j)
                End If
            Next
        Next

    End Sub

    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBoxSetStatusDate control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBoxSetStatusDate_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxSetStatusDate.CheckedChanged
        My.Settings.AutoSetStatusDate = CheckBoxSetStatusDate.Checked
        My.Settings.Save()
    End Sub


    Private Sub ComboBoxStatusTextField_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxStatusTextField.SelectedIndexChanged
        My.Settings.StatusTextField = ComboBoxStatusTextField.SelectedItem.ToString
        My.Settings.Save()
    End Sub

    Private Sub FormSettings_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SaveComboLists()
    End Sub
End Class