<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormContact
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
        Me.TextBoxEmail = New System.Windows.Forms.TextBox()
        Me.LabelEmail = New System.Windows.Forms.Label()
        Me.LabelName = New System.Windows.Forms.Label()
        Me.TextBoxName = New System.Windows.Forms.TextBox()
        Me.LabelMessage = New System.Windows.Forms.Label()
        Me.TextBoxMessage = New System.Windows.Forms.TextBox()
        Me.ComboBoxReason = New System.Windows.Forms.ComboBox()
        Me.LabelReason = New System.Windows.Forms.Label()
        Me.LabelMess = New System.Windows.Forms.Label()
        Me.ButtonSend = New System.Windows.Forms.Button()
        Me.LinkLabelType = New System.Windows.Forms.LinkLabel()
        Me.SuspendLayout()
        '
        'TextBoxEmail
        '
        Me.TextBoxEmail.Location = New System.Drawing.Point(119, 74)
        Me.TextBoxEmail.Name = "TextBoxEmail"
        Me.TextBoxEmail.Size = New System.Drawing.Size(395, 22)
        Me.TextBoxEmail.TabIndex = 1
        '
        'LabelEmail
        '
        Me.LabelEmail.AutoSize = True
        Me.LabelEmail.Location = New System.Drawing.Point(25, 79)
        Me.LabelEmail.Name = "LabelEmail"
        Me.LabelEmail.Size = New System.Drawing.Size(42, 17)
        Me.LabelEmail.TabIndex = 1
        Me.LabelEmail.Text = "Email"
        '
        'LabelName
        '
        Me.LabelName.AutoSize = True
        Me.LabelName.Location = New System.Drawing.Point(25, 40)
        Me.LabelName.Name = "LabelName"
        Me.LabelName.Size = New System.Drawing.Size(45, 17)
        Me.LabelName.TabIndex = 3
        Me.LabelName.Text = "Name"
        '
        'TextBoxName
        '
        Me.TextBoxName.Location = New System.Drawing.Point(119, 35)
        Me.TextBoxName.Name = "TextBoxName"
        Me.TextBoxName.Size = New System.Drawing.Size(395, 22)
        Me.TextBoxName.TabIndex = 0
        '
        'LabelMessage
        '
        Me.LabelMessage.AutoSize = True
        Me.LabelMessage.Location = New System.Drawing.Point(25, 480)
        Me.LabelMessage.Name = "LabelMessage"
        Me.LabelMessage.Size = New System.Drawing.Size(65, 17)
        Me.LabelMessage.TabIndex = 5
        Me.LabelMessage.Text = "Message"
        '
        'TextBoxMessage
        '
        Me.TextBoxMessage.Location = New System.Drawing.Point(119, 154)
        Me.TextBoxMessage.Multiline = True
        Me.TextBoxMessage.Name = "TextBoxMessage"
        Me.TextBoxMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxMessage.Size = New System.Drawing.Size(395, 246)
        Me.TextBoxMessage.TabIndex = 3
        '
        'ComboBoxReason
        '
        Me.ComboBoxReason.FormattingEnabled = True
        Me.ComboBoxReason.Items.AddRange(New Object() {"Bug Report", "Feature Request", "Help", "Other", "Training Request"})
        Me.ComboBoxReason.Location = New System.Drawing.Point(119, 113)
        Me.ComboBoxReason.Name = "ComboBoxReason"
        Me.ComboBoxReason.Size = New System.Drawing.Size(395, 24)
        Me.ComboBoxReason.Sorted = True
        Me.ComboBoxReason.TabIndex = 2
        '
        'LabelReason
        '
        Me.LabelReason.AutoSize = True
        Me.LabelReason.Location = New System.Drawing.Point(25, 116)
        Me.LabelReason.Name = "LabelReason"
        Me.LabelReason.Size = New System.Drawing.Size(57, 17)
        Me.LabelReason.TabIndex = 39
        Me.LabelReason.Text = "Reason"
        '
        'LabelMess
        '
        Me.LabelMess.AutoSize = True
        Me.LabelMess.Location = New System.Drawing.Point(25, 157)
        Me.LabelMess.Name = "LabelMess"
        Me.LabelMess.Size = New System.Drawing.Size(65, 17)
        Me.LabelMess.TabIndex = 40
        Me.LabelMess.Text = "Message"
        '
        'ButtonSend
        '
        Me.ButtonSend.Location = New System.Drawing.Point(439, 420)
        Me.ButtonSend.Name = "ButtonSend"
        Me.ButtonSend.Size = New System.Drawing.Size(75, 23)
        Me.ButtonSend.TabIndex = 41
        Me.ButtonSend.Text = "Send"
        Me.ButtonSend.UseVisualStyleBackColor = True
        '
        'LinkLabelType
        '
        Me.LinkLabelType.AutoSize = True
        Me.LinkLabelType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabelType.Location = New System.Drawing.Point(116, 426)
        Me.LinkLabelType.Name = "LinkLabelType"
        Me.LinkLabelType.Size = New System.Drawing.Size(227, 18)
        Me.LinkLabelType.TabIndex = 42
        Me.LinkLabelType.TabStop = True
        Me.LinkLabelType.Text = "Email: Trevor.Lowing@uspto.gov"
        '
        'FormContact
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(541, 470)
        Me.Controls.Add(Me.LinkLabelType)
        Me.Controls.Add(Me.ButtonSend)
        Me.Controls.Add(Me.LabelMess)
        Me.Controls.Add(Me.LabelReason)
        Me.Controls.Add(Me.ComboBoxReason)
        Me.Controls.Add(Me.LabelMessage)
        Me.Controls.Add(Me.TextBoxMessage)
        Me.Controls.Add(Me.LabelName)
        Me.Controls.Add(Me.TextBoxName)
        Me.Controls.Add(Me.LabelEmail)
        Me.Controls.Add(Me.TextBoxEmail)
        Me.Name = "FormContact"
        Me.ShowIcon = False
        Me.Text = "Contact Form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBoxEmail As System.Windows.Forms.TextBox
    Friend WithEvents LabelEmail As System.Windows.Forms.Label
    Friend WithEvents LabelName As System.Windows.Forms.Label
    Friend WithEvents TextBoxName As System.Windows.Forms.TextBox
    Friend WithEvents LabelMessage As System.Windows.Forms.Label
    Friend WithEvents TextBoxMessage As System.Windows.Forms.TextBox
    Friend WithEvents ComboBoxReason As System.Windows.Forms.ComboBox
    Friend WithEvents LabelReason As System.Windows.Forms.Label
    Friend WithEvents LabelMess As System.Windows.Forms.Label
    Friend WithEvents ButtonSend As System.Windows.Forms.Button
    Friend WithEvents LinkLabelType As System.Windows.Forms.LinkLabel
End Class
