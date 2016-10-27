Partial Class ReportRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.TabReview = Me.Factory.CreateRibbonTab
        Me.GroupStatus = Me.Factory.CreateRibbonGroup
        Me.ButtonProgressToNow = Me.Factory.CreateRibbonButton
        Me.Reports = Me.Factory.CreateRibbonGroup
        Me.ButtonStatusReport = Me.Factory.CreateRibbonButton
        Me.ButtonBaselineReport = Me.Factory.CreateRibbonButton
        Me.ButtonEnterpriseReport = Me.Factory.CreateRibbonButton
        Me.ButtonProjectQualityReport = Me.Factory.CreateRibbonButton
        Me.GroupPath = Me.Factory.CreateRibbonGroup
        Me.ButtonRunPath = Me.Factory.CreateRibbonButton
        Me.ToggleButtonPathTo = Me.Factory.CreateRibbonToggleButton
        Me.ToggleButtonPathFrom = Me.Factory.CreateRibbonToggleButton
        Me.DropDownPathFilter = Me.Factory.CreateRibbonDropDown
        Me.GroupSettings = Me.Factory.CreateRibbonGroup
        Me.ButtonEditSettings = Me.Factory.CreateRibbonButton
        Me.ButtonEditDefaults = Me.Factory.CreateRibbonButton
        Me.ToggleButtonExcel = Me.Factory.CreateRibbonToggleButton
        Me.GroupHelp = Me.Factory.CreateRibbonGroup
        Me.ButtonFeatureRequest = Me.Factory.CreateRibbonButton
        Me.ButtonContact = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.TabReview.SuspendLayout()
        Me.GroupStatus.SuspendLayout()
        Me.Reports.SuspendLayout()
        Me.GroupPath.SuspendLayout()
        Me.GroupSettings.SuspendLayout()
        Me.GroupHelp.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'TabReview
        '
        Me.TabReview.Groups.Add(Me.GroupStatus)
        Me.TabReview.Groups.Add(Me.Reports)
        Me.TabReview.Groups.Add(Me.GroupPath)
        Me.TabReview.Groups.Add(Me.GroupSettings)
        Me.TabReview.Groups.Add(Me.GroupHelp)
        Me.TabReview.Label = "REVIEW"
        Me.TabReview.Name = "TabReview"
        '
        'GroupStatus
        '
        Me.GroupStatus.Items.Add(Me.ButtonProgressToNow)
        Me.GroupStatus.Label = "Status"
        Me.GroupStatus.Name = "GroupStatus"
        Me.GroupStatus.Visible = False
        '
        'ButtonProgressToNow
        '
        Me.ButtonProgressToNow.Label = "Progress to Now"
        Me.ButtonProgressToNow.Name = "ButtonProgressToNow"
        Me.ButtonProgressToNow.OfficeImageId = "GroupSiteSummaryEdit"
        Me.ButtonProgressToNow.ShowImage = True
        Me.ButtonProgressToNow.SuperTip = "Update selected Task(s) Percent Complete to match current task schedule."
        '
        'Reports
        '
        Me.Reports.Items.Add(Me.ButtonStatusReport)
        Me.Reports.Items.Add(Me.ButtonBaselineReport)
        Me.Reports.Items.Add(Me.ButtonEnterpriseReport)
        Me.Reports.Items.Add(Me.ButtonProjectQualityReport)
        Me.Reports.Label = "Reports"
        Me.Reports.Name = "Reports"
        '
        'ButtonStatusReport
        '
        Me.ButtonStatusReport.Label = "Status Report"
        Me.ButtonStatusReport.Name = "ButtonStatusReport"
        Me.ButtonStatusReport.OfficeImageId = "GroupCreateReports"
        Me.ButtonStatusReport.ShowImage = True
        Me.ButtonStatusReport.SuperTip = "Exports a report of project task status items."
        '
        'ButtonBaselineReport
        '
        Me.ButtonBaselineReport.Label = "Baseline Report"
        Me.ButtonBaselineReport.Name = "ButtonBaselineReport"
        Me.ButtonBaselineReport.OfficeImageId = "Baseline"
        Me.ButtonBaselineReport.ShowImage = True
        Me.ButtonBaselineReport.SuperTip = "Exports a detailed report to run before baseline or stage gate reviews."
        '
        'ButtonEnterpriseReport
        '
        Me.ButtonEnterpriseReport.Label = "Enterprise Report"
        Me.ButtonEnterpriseReport.Name = "ButtonEnterpriseReport"
        Me.ButtonEnterpriseReport.OfficeImageId = "InviteMembersLB"
        Me.ButtonEnterpriseReport.ShowImage = True
        Me.ButtonEnterpriseReport.SuperTip = "Exports a report on special Enterprise Project fields and generates special repor" & _
    "ts."
        '
        'ButtonProjectQualityReport
        '
        Me.ButtonProjectQualityReport.Label = "Project Quality Report"
        Me.ButtonProjectQualityReport.Name = "ButtonProjectQualityReport"
        Me.ButtonProjectQualityReport.OfficeImageId = "CustomGallery1"
        Me.ButtonProjectQualityReport.ShowImage = True
        '
        'GroupPath
        '
        RibbonDialogLauncherImpl1.Enabled = False
        Me.GroupPath.DialogLauncher = RibbonDialogLauncherImpl1
        Me.GroupPath.Items.Add(Me.ButtonRunPath)
        Me.GroupPath.Items.Add(Me.ToggleButtonPathTo)
        Me.GroupPath.Items.Add(Me.ToggleButtonPathFrom)
        Me.GroupPath.Items.Add(Me.DropDownPathFilter)
        Me.GroupPath.Label = "Path Analysis"
        Me.GroupPath.Name = "GroupPath"
        '
        'ButtonRunPath
        '
        Me.ButtonRunPath.Label = "Run Path"
        Me.ButtonRunPath.Name = "ButtonRunPath"
        Me.ButtonRunPath.OfficeImageId = "GroupQuickSteps"
        Me.ButtonRunPath.ShowImage = True
        Me.ButtonRunPath.SuperTip = "Filters the schedule to show tasks that are driving the currently selected task. " & _
    "Can also show tasks that the currently selected task is driving."
        '
        'ToggleButtonPathTo
        '
        Me.ToggleButtonPathTo.Checked = True
        Me.ToggleButtonPathTo.Label = "Show Path To"
        Me.ToggleButtonPathTo.Name = "ToggleButtonPathTo"
        Me.ToggleButtonPathTo.OfficeImageId = "SmartArtReorderUp"
        Me.ToggleButtonPathTo.ShowImage = True
        '
        'ToggleButtonPathFrom
        '
        Me.ToggleButtonPathFrom.Label = "Show Path From"
        Me.ToggleButtonPathFrom.Name = "ToggleButtonPathFrom"
        Me.ToggleButtonPathFrom.OfficeImageId = "SmartArtReorderDown"
        Me.ToggleButtonPathFrom.ShowImage = True
        '
        'DropDownPathFilter
        '
        RibbonDropDownItemImpl1.Label = "Drivers Only"
        RibbonDropDownItemImpl2.Label = "Critical Only"
        RibbonDropDownItemImpl3.Label = "All Items"
        Me.DropDownPathFilter.Items.Add(RibbonDropDownItemImpl1)
        Me.DropDownPathFilter.Items.Add(RibbonDropDownItemImpl2)
        Me.DropDownPathFilter.Items.Add(RibbonDropDownItemImpl3)
        Me.DropDownPathFilter.Label = ":"
        Me.DropDownPathFilter.Name = "DropDownPathFilter"
        Me.DropDownPathFilter.OfficeImageId = "AutoFilterClassic"
        Me.DropDownPathFilter.ShowImage = True
        '
        'GroupSettings
        '
        Me.GroupSettings.Items.Add(Me.ButtonEditSettings)
        Me.GroupSettings.Items.Add(Me.ButtonEditDefaults)
        Me.GroupSettings.Items.Add(Me.ToggleButtonExcel)
        Me.GroupSettings.Label = "My Settings"
        Me.GroupSettings.Name = "GroupSettings"
        '
        'ButtonEditSettings
        '
        Me.ButtonEditSettings.Label = "Edit Settings"
        Me.ButtonEditSettings.Name = "ButtonEditSettings"
        Me.ButtonEditSettings.OfficeImageId = "ControlsGalleryClassic"
        Me.ButtonEditSettings.ShowImage = True
        Me.ButtonEditSettings.SuperTip = "Edit User Preferences"
        '
        'ButtonEditDefaults
        '
        Me.ButtonEditDefaults.Label = "Edit Defaults"
        Me.ButtonEditDefaults.Name = "ButtonEditDefaults"
        Me.ButtonEditDefaults.OfficeImageId = "ControlsGalleryClassic"
        Me.ButtonEditDefaults.ShowImage = True
        Me.ButtonEditDefaults.SuperTip = "Edit project review defaults."
        '
        'ToggleButtonExcel
        '
        Me.ToggleButtonExcel.Label = "Export Excel"
        Me.ToggleButtonExcel.Name = "ToggleButtonExcel"
        Me.ToggleButtonExcel.OfficeImageId = "ExportHtmlDocument"
        Me.ToggleButtonExcel.ShowImage = True
        Me.ToggleButtonExcel.SuperTip = "Toggle to export reports to HTML or Excel."
        '
        'GroupHelp
        '
        Me.GroupHelp.Items.Add(Me.ButtonFeatureRequest)
        Me.GroupHelp.Items.Add(Me.ButtonContact)
        Me.GroupHelp.Label = "Help"
        Me.GroupHelp.Name = "GroupHelp"
        '
        'ButtonFeatureRequest
        '
        Me.ButtonFeatureRequest.Label = "Documentation"
        Me.ButtonFeatureRequest.Name = "ButtonFeatureRequest"
        Me.ButtonFeatureRequest.OfficeImageId = "Help"
        Me.ButtonFeatureRequest.ShowImage = True
        Me.ButtonFeatureRequest.SuperTip = "Online resources for using Project."
        '
        'ButtonContact
        '
        Me.ButtonContact.Label = "Contact"
        Me.ButtonContact.Name = "ButtonContact"
        Me.ButtonContact.OfficeImageId = "EnvelopesAndLabels"
        Me.ButtonContact.ShowImage = True
        Me.ButtonContact.SuperTip = "Contact tool developer to provide feedback or report a defect."
        '
        'ReportRibbon
        '
        Me.Name = "ReportRibbon"
        Me.RibbonType = "Microsoft.Project.Project"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.TabReview)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.TabReview.ResumeLayout(False)
        Me.TabReview.PerformLayout()
        Me.GroupStatus.ResumeLayout(False)
        Me.GroupStatus.PerformLayout()
        Me.Reports.ResumeLayout(False)
        Me.Reports.PerformLayout()
        Me.GroupPath.ResumeLayout(False)
        Me.GroupPath.PerformLayout()
        Me.GroupSettings.ResumeLayout(False)
        Me.GroupSettings.PerformLayout()
        Me.GroupHelp.ResumeLayout(False)
        Me.GroupHelp.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents TabReview As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Reports As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonStatusReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonBaselineReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GroupSettings As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonEditSettings As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GroupPath As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonRunPath As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DropDownPathFilter As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents ToggleButtonPathTo As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents ToggleButtonPathFrom As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents ToggleButtonExcel As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents GroupStatus As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonProgressToNow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonEnterpriseReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonEditDefaults As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents GroupHelp As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonContact As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonFeatureRequest As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonProjectQualityReport As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As ReportRibbon
        Get
            Return Me.GetRibbon(Of ReportRibbon)()
        End Get
    End Property
End Class
