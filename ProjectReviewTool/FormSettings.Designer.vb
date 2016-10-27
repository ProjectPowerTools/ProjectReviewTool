<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormSettings
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ComboBoxPathFlag = New System.Windows.Forms.ComboBox()
        Me.LabelPathFlag = New System.Windows.Forms.Label()
        Me.LabelStatusDay = New System.Windows.Forms.Label()
        Me.ComboBoxStatusDay = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBoxStatusWeeks = New System.Windows.Forms.ComboBox()
        Me.ComboBoxEnterpriseProjectFields = New System.Windows.Forms.ComboBox()
        Me.LabelEnterpriseProjectFields = New System.Windows.Forms.Label()
        Me.LabelEnterpriseTaskFields = New System.Windows.Forms.Label()
        Me.ComboBoxEnterpriseTaskFields = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ComboBoxCustomReports = New System.Windows.Forms.ComboBox()
        Me.ButtonAddProjectField = New System.Windows.Forms.Button()
        Me.ButtonAddTaskField = New System.Windows.Forms.Button()
        Me.ButtonAddCustomReport = New System.Windows.Forms.Button()
        Me.ButtonProjectFieldDelete = New System.Windows.Forms.Button()
        Me.ButtonTaskFieldDelete = New System.Windows.Forms.Button()
        Me.ButtonCustomReportDelete = New System.Windows.Forms.Button()
        Me.CheckBoxSetStatusDate = New System.Windows.Forms.CheckBox()
        Me.ComboBoxStatusTextField = New System.Windows.Forms.ComboBox()
        Me.LabelStatusText = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ComboBoxPathFlag
        '
        Me.ComboBoxPathFlag.FormattingEnabled = True
        Me.ComboBoxPathFlag.Items.AddRange(New Object() {"Flag1", "Flag2", "Flag3", "Flag4", "Flag5", "Flag6", "Flag7", "Flag8", "Flag9", "Flag10", "Flag11", "Flag12", "Flag13", "Flag14", "Flag15", "Flag16", "Flag17", "Flag18", "Flag19", "Flag20"})
        Me.ComboBoxPathFlag.Location = New System.Drawing.Point(132, 117)
        Me.ComboBoxPathFlag.Name = "ComboBoxPathFlag"
        Me.ComboBoxPathFlag.Size = New System.Drawing.Size(395, 24)
        Me.ComboBoxPathFlag.TabIndex = 10
        '
        'LabelPathFlag
        '
        Me.LabelPathFlag.AutoSize = True
        Me.LabelPathFlag.Location = New System.Drawing.Point(12, 120)
        Me.LabelPathFlag.Name = "LabelPathFlag"
        Me.LabelPathFlag.Size = New System.Drawing.Size(102, 17)
        Me.LabelPathFlag.TabIndex = 9
        Me.LabelPathFlag.Text = "Path Flag Field"
        '
        'LabelStatusDay
        '
        Me.LabelStatusDay.AutoSize = True
        Me.LabelStatusDay.Location = New System.Drawing.Point(15, 74)
        Me.LabelStatusDay.Name = "LabelStatusDay"
        Me.LabelStatusDay.Size = New System.Drawing.Size(82, 17)
        Me.LabelStatusDay.TabIndex = 12
        Me.LabelStatusDay.Text = "Status Date"
        '
        'ComboBoxStatusDay
        '
        Me.ComboBoxStatusDay.FormattingEnabled = True
        Me.ComboBoxStatusDay.Items.AddRange(New Object() {"Current Day", "Next Sunday", "Next Monday", "Next Tuesday", "Next Wednesday", "Next Thursday", "Next Friday", "Next Saturday", "Next Day"})
        Me.ComboBoxStatusDay.Location = New System.Drawing.Point(132, 69)
        Me.ComboBoxStatusDay.Name = "ComboBoxStatusDay"
        Me.ComboBoxStatusDay.Size = New System.Drawing.Size(270, 24)
        Me.ComboBoxStatusDay.TabIndex = 13
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(251, 17)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Status Weeks in Future to Take Status"
        '
        'ComboBoxStatusWeeks
        '
        Me.ComboBoxStatusWeeks.FormattingEnabled = True
        Me.ComboBoxStatusWeeks.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15"})
        Me.ComboBoxStatusWeeks.Location = New System.Drawing.Point(300, 26)
        Me.ComboBoxStatusWeeks.Name = "ComboBoxStatusWeeks"
        Me.ComboBoxStatusWeeks.Size = New System.Drawing.Size(227, 24)
        Me.ComboBoxStatusWeeks.TabIndex = 16
        '
        'ComboBoxEnterpriseProjectFields
        '
        Me.ComboBoxEnterpriseProjectFields.FormattingEnabled = True
        Me.ComboBoxEnterpriseProjectFields.Location = New System.Drawing.Point(177, 196)
        Me.ComboBoxEnterpriseProjectFields.Name = "ComboBoxEnterpriseProjectFields"
        Me.ComboBoxEnterpriseProjectFields.Size = New System.Drawing.Size(304, 24)
        Me.ComboBoxEnterpriseProjectFields.Sorted = True
        Me.ComboBoxEnterpriseProjectFields.TabIndex = 18
        '
        'LabelEnterpriseProjectFields
        '
        Me.LabelEnterpriseProjectFields.AutoSize = True
        Me.LabelEnterpriseProjectFields.Location = New System.Drawing.Point(9, 203)
        Me.LabelEnterpriseProjectFields.Name = "LabelEnterpriseProjectFields"
        Me.LabelEnterpriseProjectFields.Size = New System.Drawing.Size(162, 17)
        Me.LabelEnterpriseProjectFields.TabIndex = 19
        Me.LabelEnterpriseProjectFields.Text = "Enterprise Project Fields"
        '
        'LabelEnterpriseTaskFields
        '
        Me.LabelEnterpriseTaskFields.AutoSize = True
        Me.LabelEnterpriseTaskFields.Location = New System.Drawing.Point(10, 256)
        Me.LabelEnterpriseTaskFields.Name = "LabelEnterpriseTaskFields"
        Me.LabelEnterpriseTaskFields.Size = New System.Drawing.Size(149, 17)
        Me.LabelEnterpriseTaskFields.TabIndex = 21
        Me.LabelEnterpriseTaskFields.Text = "Enterprise Task Fields"
        '
        'ComboBoxEnterpriseTaskFields
        '
        Me.ComboBoxEnterpriseTaskFields.FormattingEnabled = True
        Me.ComboBoxEnterpriseTaskFields.Location = New System.Drawing.Point(174, 249)
        Me.ComboBoxEnterpriseTaskFields.Name = "ComboBoxEnterpriseTaskFields"
        Me.ComboBoxEnterpriseTaskFields.Size = New System.Drawing.Size(307, 24)
        Me.ComboBoxEnterpriseTaskFields.Sorted = True
        Me.ComboBoxEnterpriseTaskFields.TabIndex = 20
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 305)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(109, 17)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Custom Reports"
        '
        'ComboBoxCustomReports
        '
        Me.ComboBoxCustomReports.FormattingEnabled = True
        Me.ComboBoxCustomReports.Location = New System.Drawing.Point(174, 298)
        Me.ComboBoxCustomReports.Name = "ComboBoxCustomReports"
        Me.ComboBoxCustomReports.Size = New System.Drawing.Size(307, 24)
        Me.ComboBoxCustomReports.Sorted = True
        Me.ComboBoxCustomReports.TabIndex = 22
        '
        'ButtonAddProjectField
        '
        Me.ButtonAddProjectField.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonAddProjectField.Location = New System.Drawing.Point(487, 197)
        Me.ButtonAddProjectField.Name = "ButtonAddProjectField"
        Me.ButtonAddProjectField.Size = New System.Drawing.Size(40, 23)
        Me.ButtonAddProjectField.TabIndex = 24
        Me.ButtonAddProjectField.Text = "+"
        Me.ButtonAddProjectField.UseVisualStyleBackColor = True
        '
        'ButtonAddTaskField
        '
        Me.ButtonAddTaskField.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonAddTaskField.Location = New System.Drawing.Point(487, 249)
        Me.ButtonAddTaskField.Name = "ButtonAddTaskField"
        Me.ButtonAddTaskField.Size = New System.Drawing.Size(40, 23)
        Me.ButtonAddTaskField.TabIndex = 25
        Me.ButtonAddTaskField.Text = "+"
        Me.ButtonAddTaskField.UseVisualStyleBackColor = True
        '
        'ButtonAddCustomReport
        '
        Me.ButtonAddCustomReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonAddCustomReport.Location = New System.Drawing.Point(487, 298)
        Me.ButtonAddCustomReport.Name = "ButtonAddCustomReport"
        Me.ButtonAddCustomReport.Size = New System.Drawing.Size(40, 23)
        Me.ButtonAddCustomReport.TabIndex = 26
        Me.ButtonAddCustomReport.Text = "+"
        Me.ButtonAddCustomReport.UseVisualStyleBackColor = True
        '
        'ButtonProjectFieldDelete
        '
        Me.ButtonProjectFieldDelete.Location = New System.Drawing.Point(534, 196)
        Me.ButtonProjectFieldDelete.Name = "ButtonProjectFieldDelete"
        Me.ButtonProjectFieldDelete.Size = New System.Drawing.Size(46, 23)
        Me.ButtonProjectFieldDelete.TabIndex = 27
        Me.ButtonProjectFieldDelete.Text = "X"
        Me.ButtonProjectFieldDelete.UseVisualStyleBackColor = True
        '
        'ButtonTaskFieldDelete
        '
        Me.ButtonTaskFieldDelete.Location = New System.Drawing.Point(533, 249)
        Me.ButtonTaskFieldDelete.Name = "ButtonTaskFieldDelete"
        Me.ButtonTaskFieldDelete.Size = New System.Drawing.Size(46, 23)
        Me.ButtonTaskFieldDelete.TabIndex = 28
        Me.ButtonTaskFieldDelete.Text = "X"
        Me.ButtonTaskFieldDelete.UseVisualStyleBackColor = True
        '
        'ButtonCustomReportDelete
        '
        Me.ButtonCustomReportDelete.Location = New System.Drawing.Point(533, 298)
        Me.ButtonCustomReportDelete.Name = "ButtonCustomReportDelete"
        Me.ButtonCustomReportDelete.Size = New System.Drawing.Size(46, 23)
        Me.ButtonCustomReportDelete.TabIndex = 29
        Me.ButtonCustomReportDelete.Text = "X"
        Me.ButtonCustomReportDelete.UseVisualStyleBackColor = True
        '
        'CheckBoxSetStatusDate
        '
        Me.CheckBoxSetStatusDate.AutoSize = True
        Me.CheckBoxSetStatusDate.Location = New System.Drawing.Point(443, 69)
        Me.CheckBoxSetStatusDate.Name = "CheckBoxSetStatusDate"
        Me.CheckBoxSetStatusDate.Size = New System.Drawing.Size(84, 21)
        Me.CheckBoxSetStatusDate.TabIndex = 32
        Me.CheckBoxSetStatusDate.Text = "Auto Set"
        Me.CheckBoxSetStatusDate.UseVisualStyleBackColor = True
        '
        'ComboBoxStatusTextField
        '
        Me.ComboBoxStatusTextField.FormattingEnabled = True
        Me.ComboBoxStatusTextField.Items.AddRange(New Object() {"Text1", "Text2", "Text3", "Text4", "Text5", "Text6", "Text7", "Text8", "Text9", "Text10", "Text11", "Text12", "Text13", "Text14", "Text15", "Text16", "Text17", "Text18", "Text19", "Text20", "Text21", "Text22", "Text23", "Text24", "Text25", "Text26", "Text27", "Text28", "Text29", "Text30", "", "Text15", "Text16", "Text17", "Text18", "Text19", "Text20"})
        Me.ComboBoxStatusTextField.Location = New System.Drawing.Point(132, 156)
        Me.ComboBoxStatusTextField.Name = "ComboBoxStatusTextField"
        Me.ComboBoxStatusTextField.Size = New System.Drawing.Size(395, 24)
        Me.ComboBoxStatusTextField.TabIndex = 34
        '
        'LabelStatusText
        '
        Me.LabelStatusText.AutoSize = True
        Me.LabelStatusText.Location = New System.Drawing.Point(12, 159)
        Me.LabelStatusText.Name = "LabelStatusText"
        Me.LabelStatusText.Size = New System.Drawing.Size(113, 17)
        Me.LabelStatusText.TabIndex = 33
        Me.LabelStatusText.Text = "Status Text Field"
        '
        'FormSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(599, 337)
        Me.Controls.Add(Me.ComboBoxStatusTextField)
        Me.Controls.Add(Me.LabelStatusText)
        Me.Controls.Add(Me.CheckBoxSetStatusDate)
        Me.Controls.Add(Me.ButtonCustomReportDelete)
        Me.Controls.Add(Me.ButtonTaskFieldDelete)
        Me.Controls.Add(Me.ButtonProjectFieldDelete)
        Me.Controls.Add(Me.ButtonAddCustomReport)
        Me.Controls.Add(Me.ButtonAddTaskField)
        Me.Controls.Add(Me.ButtonAddProjectField)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ComboBoxCustomReports)
        Me.Controls.Add(Me.LabelEnterpriseTaskFields)
        Me.Controls.Add(Me.ComboBoxEnterpriseTaskFields)
        Me.Controls.Add(Me.LabelEnterpriseProjectFields)
        Me.Controls.Add(Me.ComboBoxEnterpriseProjectFields)
        Me.Controls.Add(Me.ComboBoxStatusWeeks)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBoxStatusDay)
        Me.Controls.Add(Me.LabelStatusDay)
        Me.Controls.Add(Me.ComboBoxPathFlag)
        Me.Controls.Add(Me.LabelPathFlag)
        Me.Name = "FormSettings"
        Me.ShowIcon = False
        Me.Text = "My Report Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ComboBoxPathFlag As System.Windows.Forms.ComboBox
    Friend WithEvents LabelPathFlag As System.Windows.Forms.Label
    Friend WithEvents LabelStatusDay As System.Windows.Forms.Label
    Friend WithEvents ComboBoxStatusDay As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxStatusWeeks As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxEnterpriseProjectFields As System.Windows.Forms.ComboBox
    Friend WithEvents LabelEnterpriseProjectFields As System.Windows.Forms.Label
    Friend WithEvents LabelEnterpriseTaskFields As System.Windows.Forms.Label
    Friend WithEvents ComboBoxEnterpriseTaskFields As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxCustomReports As System.Windows.Forms.ComboBox
    Friend WithEvents ButtonAddProjectField As System.Windows.Forms.Button
    Friend WithEvents ButtonAddTaskField As System.Windows.Forms.Button
    Friend WithEvents ButtonAddCustomReport As System.Windows.Forms.Button
    Friend WithEvents ButtonProjectFieldDelete As System.Windows.Forms.Button
    Friend WithEvents ButtonTaskFieldDelete As System.Windows.Forms.Button
    Friend WithEvents ButtonCustomReportDelete As System.Windows.Forms.Button
    Friend WithEvents CheckBoxSetStatusDate As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBoxStatusTextField As System.Windows.Forms.ComboBox
    Friend WithEvents LabelStatusText As System.Windows.Forms.Label
End Class
