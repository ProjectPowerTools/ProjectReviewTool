' Copyright (c) 2014 Trevor Lowing

'This Source Code Form is subject to the terms of the Mozilla Public
' License, v. 2.0. If a copy of the MPL was not distributed with this
' file, You can obtain one at http://mozilla.org/MPL/2.0/.

Imports System.Net.Mail
Imports System.Text.RegularExpressions

Public Class FormContact
    ''' <summary>
    ''' Checks the form.
    ''' </summary>
    Private Sub CheckForm()
        Dim blnReady As Boolean = True
        Dim strMessage As String = ""

        If Trim(TextBoxName.Text) <> "" Then
            TextBoxName.BackColor = Nothing
        Else
            TextBoxName.BackColor = Drawing.Color.Salmon
            blnReady = False
            strMessage = "Please Enter Your Name" & vbCrLf
        End If
        If Trim(TextBoxEmail.Text) <> "" And IsValidEmailFormat(TextBoxEmail.Text) Then
            TextBoxEmail.BackColor = Nothing
        Else
            TextBoxEmail.BackColor = Drawing.Color.Salmon
            blnReady = False
            strMessage = strMessage & "Please Enter A Valid Email Address" & vbCrLf
        End If
        If Trim(TextBoxMessage.Text) <> "" Then
            TextBoxMessage.BackColor = Nothing
        Else
            TextBoxMessage.BackColor = Drawing.Color.Salmon
            blnReady = False
            strMessage = strMessage & "Please Enter A Message" & vbCrLf
        End If

        If ComboBoxReason.SelectedIndex <> -1 Then
            ComboBoxReason.BackColor = Nothing
        Else
            ComboBoxReason.BackColor = Drawing.Color.Salmon
            blnReady = False
            strMessage = strMessage & "Please Enter A Reason" & vbCrLf
        End If

        If blnReady Then
            'Dim colFields As New Dictionary(Of String, String)
            'colFields.Add("Name", TextBoxName.Text)
            ' colFields.Add("Email", TextBoxEmail.Text)
            ' colFields.Add("Reason", ComboBoxReason.SelectedItem.ToString)
            ' colFields.Add("Message", TextBoxMessage.Text)
            'strMessage = PostForm("http://www.ProjectReviewTool.com/contact.php", colFields)
            MsgBox(strMessage)
        Else
            MsgBox(strMessage)
        End If

    End Sub

    ''' <summary>
    ''' Determines whether is valid email format.
    ''' </summary>
    ''' <param name="strInput">The strInput.</param>
    ''' <returns></returns>
    Function IsValidEmailFormat(ByVal strInput As String) As Boolean
        Return Regex.IsMatch(strInput, "^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$")
    End Function

    ''' <summary>
    ''' Handles the Click event of the ButtonSend control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub ButtonSend_Click(sender As Object, e As EventArgs) Handles ButtonSend.Click
        CheckForm()
    End Sub

    ''' <summary>
    ''' Handles the LinkClicked event of the LinkLabelType control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="Windows.Forms.LinkLabelLinkClickedEventArgs"/> instance containing the event data.</param>
    Private Sub LinkLabelType_LinkClicked(sender As Object, e As Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabelType.LinkClicked
        System.Diagnostics.Process.Start("mailto:Trevor.Lowing@uspto.gov")
    End Sub
End Class
