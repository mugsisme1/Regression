<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form6
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form6))
        Me.krbSPIACurrent = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton()
        Me.krbSPIAEffective = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton()
        Me.kdtpSPIARateDate = New ComponentFactory.Krypton.Toolkit.KryptonDateTimePicker()
        Me.kbOK = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.klNoBench = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.SuspendLayout()
        '
        'krbSPIACurrent
        '
        Me.krbSPIACurrent.Checked = True
        Me.krbSPIACurrent.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.krbSPIACurrent.Location = New System.Drawing.Point(34, 10)
        Me.krbSPIACurrent.Margin = New System.Windows.Forms.Padding(2)
        Me.krbSPIACurrent.Name = "krbSPIACurrent"
        Me.krbSPIACurrent.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.krbSPIACurrent.Size = New System.Drawing.Size(111, 19)
        Me.krbSPIACurrent.TabIndex = 0
        Me.krbSPIACurrent.Text = "Use Current Rates"
        Me.krbSPIACurrent.Values.ExtraText = ""
        Me.krbSPIACurrent.Values.Image = Nothing
        Me.krbSPIACurrent.Values.Text = "Use Current Rates"
        '
        'krbSPIAEffective
        '
        Me.krbSPIAEffective.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.krbSPIAEffective.Location = New System.Drawing.Point(34, 46)
        Me.krbSPIAEffective.Margin = New System.Windows.Forms.Padding(2)
        Me.krbSPIAEffective.Name = "krbSPIAEffective"
        Me.krbSPIAEffective.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.krbSPIAEffective.Size = New System.Drawing.Size(112, 19)
        Me.krbSPIAEffective.TabIndex = 1
        Me.krbSPIAEffective.Text = "Use Effective Date"
        Me.krbSPIAEffective.Values.ExtraText = ""
        Me.krbSPIAEffective.Values.Image = Nothing
        Me.krbSPIAEffective.Values.Text = "Use Effective Date"
        '
        'kdtpSPIARateDate
        '
        Me.kdtpSPIARateDate.CalendarDayOfWeekStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.CalendarDay
        Me.kdtpSPIARateDate.CalendarDayStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.CalendarDay
        Me.kdtpSPIARateDate.CalendarHeaderStyle = ComponentFactory.Krypton.Toolkit.HeaderStyle.Calendar
        Me.kdtpSPIARateDate.CalendarTodayDate = New Date(2011, 12, 19, 0, 0, 0, 0)
        Me.kdtpSPIARateDate.DropButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.InputControl
        Me.kdtpSPIARateDate.InputControlStyle = ComponentFactory.Krypton.Toolkit.InputControlStyle.Standalone
        Me.kdtpSPIARateDate.Location = New System.Drawing.Point(196, 45)
        Me.kdtpSPIARateDate.Margin = New System.Windows.Forms.Padding(2)
        Me.kdtpSPIARateDate.MaxDate = New Date(2020, 12, 31, 0, 0, 0, 0)
        Me.kdtpSPIARateDate.MinDate = New Date(2009, 1, 1, 0, 0, 0, 0)
        Me.kdtpSPIARateDate.Name = "kdtpSPIARateDate"
        Me.kdtpSPIARateDate.Size = New System.Drawing.Size(238, 20)
        Me.kdtpSPIARateDate.TabIndex = 2
        Me.kdtpSPIARateDate.UpDownButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.InputControl
        Me.kdtpSPIARateDate.Visible = False
        '
        'kbOK
        '
        Me.kbOK.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Standalone
        Me.kbOK.Location = New System.Drawing.Point(437, 96)
        Me.kbOK.Margin = New System.Windows.Forms.Padding(2)
        Me.kbOK.Name = "kbOK"
        Me.kbOK.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbOK.Size = New System.Drawing.Size(55, 45)
        Me.kbOK.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbOK.StateCommon.Border.Rounding = 6
        Me.kbOK.StateCommon.Border.Width = 3
        Me.kbOK.TabIndex = 4
        Me.kbOK.Text = "OK"
        Me.kbOK.Values.ExtraText = ""
        Me.kbOK.Values.Image = Nothing
        Me.kbOK.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbOK.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbOK.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbOK.Values.Text = "OK"
        '
        'klNoBench
        '
        Me.klNoBench.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.klNoBench.Location = New System.Drawing.Point(34, 79)
        Me.klNoBench.Margin = New System.Windows.Forms.Padding(2)
        Me.klNoBench.Name = "klNoBench"
        Me.klNoBench.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.klNoBench.Size = New System.Drawing.Size(365, 33)
        Me.klNoBench.StateNormal.LongText.MultiLine = ComponentFactory.Krypton.Toolkit.InheritBool.[True]
        Me.klNoBench.TabIndex = 4
        Me.klNoBench.TabStop = False
        Me.klNoBench.Text = "Note:  When run with an Effective Date, no benchmarks can be created.  " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Mismatch" & _
    "ed values will appear in RED."
        Me.klNoBench.Values.ExtraText = ""
        Me.klNoBench.Values.Image = Nothing
        Me.klNoBench.Values.Text = "Note:  When run with an Effective Date, no benchmarks can be created.  " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Mismatch" & _
    "ed values will appear in RED."
        '
        'Form6
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(503, 152)
        Me.Controls.Add(Me.klNoBench)
        Me.Controls.Add(Me.kbOK)
        Me.Controls.Add(Me.kdtpSPIARateDate)
        Me.Controls.Add(Me.krbSPIAEffective)
        Me.Controls.Add(Me.krbSPIACurrent)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form6"
        Me.Text = "The Regressionator:  Select SPIA Rate Date"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents kdtpSPIARateDate As ComponentFactory.Krypton.Toolkit.KryptonDateTimePicker
    Friend WithEvents kbOK As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents klNoBench As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Public WithEvents krbSPIACurrent As ComponentFactory.Krypton.Toolkit.KryptonRadioButton
    Public WithEvents krbSPIAEffective As ComponentFactory.Krypton.Toolkit.KryptonRadioButton
End Class
