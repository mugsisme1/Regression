<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form9
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form9))
        Me.klNoBench = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.kbOK = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.kdtpSPDARateDate = New ComponentFactory.Krypton.Toolkit.KryptonDateTimePicker()
        Me.krbSPDAEffective = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton()
        Me.krbSPDACurrent = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton()
        Me.SuspendLayout()
        '
        'klNoBench
        '
        Me.klNoBench.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.klNoBench.Location = New System.Drawing.Point(33, 79)
        Me.klNoBench.Margin = New System.Windows.Forms.Padding(2)
        Me.klNoBench.Name = "klNoBench"
        Me.klNoBench.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.klNoBench.Size = New System.Drawing.Size(365, 33)
        Me.klNoBench.StateNormal.LongText.MultiLine = ComponentFactory.Krypton.Toolkit.InheritBool.[True]
        Me.klNoBench.TabIndex = 10
        Me.klNoBench.TabStop = False
        Me.klNoBench.Text = "Note:  When run with an Effective Date, no benchmarks can be created.  " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Mismatch" & _
    "ed values will appear in RED."
        Me.klNoBench.Values.ExtraText = ""
        Me.klNoBench.Values.Image = Nothing
        Me.klNoBench.Values.Text = "Note:  When run with an Effective Date, no benchmarks can be created.  " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Mismatch" & _
    "ed values will appear in RED."
        '
        'kbOK
        '
        Me.kbOK.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Standalone
        Me.kbOK.Location = New System.Drawing.Point(424, 100)
        Me.kbOK.Margin = New System.Windows.Forms.Padding(2)
        Me.kbOK.Name = "kbOK"
        Me.kbOK.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbOK.Size = New System.Drawing.Size(55, 43)
        Me.kbOK.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbOK.StateCommon.Border.Rounding = 6
        Me.kbOK.StateCommon.Border.Width = 3
        Me.kbOK.TabIndex = 9
        Me.kbOK.Text = "OK"
        Me.kbOK.Values.ExtraText = ""
        Me.kbOK.Values.Image = Nothing
        Me.kbOK.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbOK.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbOK.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbOK.Values.Text = "OK"
        '
        'kdtpSPDARateDate
        '
        Me.kdtpSPDARateDate.CalendarDayOfWeekStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.CalendarDay
        Me.kdtpSPDARateDate.CalendarDayStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.CalendarDay
        Me.kdtpSPDARateDate.CalendarHeaderStyle = ComponentFactory.Krypton.Toolkit.HeaderStyle.Calendar
        Me.kdtpSPDARateDate.CalendarTodayDate = New Date(2011, 12, 19, 0, 0, 0, 0)
        Me.kdtpSPDARateDate.DropButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.InputControl
        Me.kdtpSPDARateDate.InputControlStyle = ComponentFactory.Krypton.Toolkit.InputControlStyle.Standalone
        Me.kdtpSPDARateDate.Location = New System.Drawing.Point(195, 45)
        Me.kdtpSPDARateDate.Margin = New System.Windows.Forms.Padding(2)
        Me.kdtpSPDARateDate.MaxDate = New Date(2020, 12, 31, 0, 0, 0, 0)
        Me.kdtpSPDARateDate.MinDate = New Date(2009, 1, 1, 0, 0, 0, 0)
        Me.kdtpSPDARateDate.Name = "kdtpSPDARateDate"
        Me.kdtpSPDARateDate.Size = New System.Drawing.Size(238, 20)
        Me.kdtpSPDARateDate.TabIndex = 7
        Me.kdtpSPDARateDate.UpDownButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.InputControl
        Me.kdtpSPDARateDate.Visible = False
        '
        'krbSPDAEffective
        '
        Me.krbSPDAEffective.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.krbSPDAEffective.Location = New System.Drawing.Point(33, 46)
        Me.krbSPDAEffective.Margin = New System.Windows.Forms.Padding(2)
        Me.krbSPDAEffective.Name = "krbSPDAEffective"
        Me.krbSPDAEffective.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.krbSPDAEffective.Size = New System.Drawing.Size(112, 19)
        Me.krbSPDAEffective.TabIndex = 6
        Me.krbSPDAEffective.Text = "Use Effective Date"
        Me.krbSPDAEffective.Values.ExtraText = ""
        Me.krbSPDAEffective.Values.Image = Nothing
        Me.krbSPDAEffective.Values.Text = "Use Effective Date"
        '
        'krbSPDACurrent
        '
        Me.krbSPDACurrent.Checked = True
        Me.krbSPDACurrent.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.krbSPDACurrent.Location = New System.Drawing.Point(33, 10)
        Me.krbSPDACurrent.Margin = New System.Windows.Forms.Padding(2)
        Me.krbSPDACurrent.Name = "krbSPDACurrent"
        Me.krbSPDACurrent.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.krbSPDACurrent.Size = New System.Drawing.Size(111, 19)
        Me.krbSPDACurrent.TabIndex = 5
        Me.krbSPDACurrent.Text = "Use Current Rates"
        Me.krbSPDACurrent.Values.ExtraText = ""
        Me.krbSPDACurrent.Values.Image = Nothing
        Me.krbSPDACurrent.Values.Text = "Use Current Rates"
        '
        'Form9
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(503, 154)
        Me.Controls.Add(Me.klNoBench)
        Me.Controls.Add(Me.kbOK)
        Me.Controls.Add(Me.kdtpSPDARateDate)
        Me.Controls.Add(Me.krbSPDAEffective)
        Me.Controls.Add(Me.krbSPDACurrent)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form9"
        Me.Text = "The Regressionator:  Select SPDA Rate Date"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents klNoBench As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents kdtpSPDARateDate As ComponentFactory.Krypton.Toolkit.KryptonDateTimePicker
    Public WithEvents krbSPDAEffective As ComponentFactory.Krypton.Toolkit.KryptonRadioButton
    Public WithEvents krbSPDACurrent As ComponentFactory.Krypton.Toolkit.KryptonRadioButton
    Public WithEvents kbOK As ComponentFactory.Krypton.Toolkit.KryptonButton
End Class
