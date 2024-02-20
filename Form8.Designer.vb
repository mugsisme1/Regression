<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form8
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form8))
        Me.klNoBench = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.kbOK = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.kdtpVAHistoricalDate = New ComponentFactory.Krypton.Toolkit.KryptonDateTimePicker()
        Me.krbVAHistoricalEffective = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton()
        Me.krbVAHistoricalCurrent = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton()
        Me.SuspendLayout()
        '
        'klNoBench
        '
        Me.klNoBench.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.klNoBench.Location = New System.Drawing.Point(46, 79)
        Me.klNoBench.Margin = New System.Windows.Forms.Padding(2)
        Me.klNoBench.Name = "klNoBench"
        Me.klNoBench.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.klNoBench.Size = New System.Drawing.Size(371, 19)
        Me.klNoBench.StateNormal.LongText.MultiLine = ComponentFactory.Krypton.Toolkit.InheritBool.[True]
        Me.klNoBench.TabIndex = 10
        Me.klNoBench.TabStop = False
        Me.klNoBench.Text = "Note:  When run with Previous numbers, no benchmarks can be created.  "
        Me.klNoBench.Values.ExtraText = ""
        Me.klNoBench.Values.Image = Nothing
        Me.klNoBench.Values.Text = "Note:  When run with Previous numbers, no benchmarks can be created.  "
        '
        'kbOK
        '
        Me.kbOK.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Standalone
        Me.kbOK.Location = New System.Drawing.Point(459, 96)
        Me.kbOK.Margin = New System.Windows.Forms.Padding(2)
        Me.kbOK.Name = "kbOK"
        Me.kbOK.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbOK.Size = New System.Drawing.Size(55, 47)
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
        'kdtpVAHistoricalDate
        '
        Me.kdtpVAHistoricalDate.CalendarDayOfWeekStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.CalendarDay
        Me.kdtpVAHistoricalDate.CalendarDayStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.CalendarDay
        Me.kdtpVAHistoricalDate.CalendarHeaderStyle = ComponentFactory.Krypton.Toolkit.HeaderStyle.Calendar
        Me.kdtpVAHistoricalDate.CalendarTodayDate = New Date(2011, 12, 19, 0, 0, 0, 0)
        Me.kdtpVAHistoricalDate.DropButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.InputControl
        Me.kdtpVAHistoricalDate.InputControlStyle = ComponentFactory.Krypton.Toolkit.InputControlStyle.Standalone
        Me.kdtpVAHistoricalDate.Location = New System.Drawing.Point(208, 45)
        Me.kdtpVAHistoricalDate.Margin = New System.Windows.Forms.Padding(2)
        Me.kdtpVAHistoricalDate.MaxDate = New Date(2020, 12, 31, 0, 0, 0, 0)
        Me.kdtpVAHistoricalDate.MinDate = New Date(2009, 1, 1, 0, 0, 0, 0)
        Me.kdtpVAHistoricalDate.Name = "kdtpVAHistoricalDate"
        Me.kdtpVAHistoricalDate.Size = New System.Drawing.Size(254, 20)
        Me.kdtpVAHistoricalDate.TabIndex = 7
        Me.kdtpVAHistoricalDate.UpDownButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.InputControl
        Me.kdtpVAHistoricalDate.Visible = False
        '
        'krbVAHistoricalEffective
        '
        Me.krbVAHistoricalEffective.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.krbVAHistoricalEffective.Location = New System.Drawing.Point(24, 45)
        Me.krbVAHistoricalEffective.Margin = New System.Windows.Forms.Padding(2)
        Me.krbVAHistoricalEffective.Name = "krbVAHistoricalEffective"
        Me.krbVAHistoricalEffective.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.krbVAHistoricalEffective.Size = New System.Drawing.Size(160, 19)
        Me.krbVAHistoricalEffective.TabIndex = 6
        Me.krbVAHistoricalEffective.Text = "Use Previous Hist. Numbers"
        Me.krbVAHistoricalEffective.Values.ExtraText = ""
        Me.krbVAHistoricalEffective.Values.Image = Nothing
        Me.krbVAHistoricalEffective.Values.Text = "Use Previous Hist. Numbers"
        '
        'krbVAHistoricalCurrent
        '
        Me.krbVAHistoricalCurrent.Checked = True
        Me.krbVAHistoricalCurrent.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.krbVAHistoricalCurrent.Location = New System.Drawing.Point(24, 10)
        Me.krbVAHistoricalCurrent.Margin = New System.Windows.Forms.Padding(2)
        Me.krbVAHistoricalCurrent.Name = "krbVAHistoricalCurrent"
        Me.krbVAHistoricalCurrent.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.krbVAHistoricalCurrent.Size = New System.Drawing.Size(155, 19)
        Me.krbVAHistoricalCurrent.TabIndex = 5
        Me.krbVAHistoricalCurrent.Text = "Use Current Hist. Numbers"
        Me.krbVAHistoricalCurrent.Values.ExtraText = ""
        Me.krbVAHistoricalCurrent.Values.Image = Nothing
        Me.krbVAHistoricalCurrent.Values.Text = "Use Current Hist. Numbers"
        '
        'Form8
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(525, 154)
        Me.Controls.Add(Me.klNoBench)
        Me.Controls.Add(Me.kbOK)
        Me.Controls.Add(Me.kdtpVAHistoricalDate)
        Me.Controls.Add(Me.krbVAHistoricalEffective)
        Me.Controls.Add(Me.krbVAHistoricalCurrent)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form8"
        Me.Text = "The Regressionator:  Previous Hist. Numbers"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents klNoBench As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents kbOK As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents kdtpVAHistoricalDate As ComponentFactory.Krypton.Toolkit.KryptonDateTimePicker
    Public WithEvents krbVAHistoricalEffective As ComponentFactory.Krypton.Toolkit.KryptonRadioButton
    Public WithEvents krbVAHistoricalCurrent As ComponentFactory.Krypton.Toolkit.KryptonRadioButton
End Class
