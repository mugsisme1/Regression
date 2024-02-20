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
        Me.krbSPIACurrent = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton
        Me.krbSPIAEffective = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton
        Me.kdtpSPIARateDate = New ComponentFactory.Krypton.Toolkit.KryptonDateTimePicker
        Me.kbOK = New ComponentFactory.Krypton.Toolkit.KryptonButton
        Me.klNoBench = New ComponentFactory.Krypton.Toolkit.KryptonLabel
        Me.kbCancel = New ComponentFactory.Krypton.Toolkit.KryptonButton
        Me.SuspendLayout()
        '
        'krbSPIACurrent
        '
        Me.krbSPIACurrent.Checked = True
        Me.krbSPIACurrent.Location = New System.Drawing.Point(46, 12)
        Me.krbSPIACurrent.Name = "krbSPIACurrent"
        Me.krbSPIACurrent.Size = New System.Drawing.Size(160, 23)
        Me.krbSPIACurrent.TabIndex = 0
        Me.krbSPIACurrent.Values.Text = "Use Current Rates"
        '
        'krbSPIAEffective
        '
        Me.krbSPIAEffective.Location = New System.Drawing.Point(46, 57)
        Me.krbSPIAEffective.Name = "krbSPIAEffective"
        Me.krbSPIAEffective.Size = New System.Drawing.Size(160, 23)
        Me.krbSPIAEffective.TabIndex = 1
        Me.krbSPIAEffective.Values.Text = "Use Effective Date"
        '
        'kdtpSPIARateDate
        '
        Me.kdtpSPIARateDate.Location = New System.Drawing.Point(262, 55)
        Me.kdtpSPIARateDate.MaxDate = New Date(2020, 12, 31, 0, 0, 0, 0)
        Me.kdtpSPIARateDate.MinDate = New Date(2009, 1, 1, 0, 0, 0, 0)
        Me.kdtpSPIARateDate.Name = "kdtpSPIARateDate"
        Me.kdtpSPIARateDate.Size = New System.Drawing.Size(318, 25)
        Me.kdtpSPIARateDate.TabIndex = 2
        Me.kdtpSPIARateDate.Visible = False
        '
        'kbOK
        '
        Me.kbOK.Location = New System.Drawing.Point(513, 137)
        Me.kbOK.Name = "kbOK"
        Me.kbOK.Size = New System.Drawing.Size(73, 39)
        Me.kbOK.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbOK.StateCommon.Border.Rounding = 6
        Me.kbOK.StateCommon.Border.Width = 3
        Me.kbOK.TabIndex = 4
        Me.kbOK.Values.Text = "OK"
        '
        'klNoBench
        '
        Me.klNoBench.Location = New System.Drawing.Point(46, 97)
        Me.klNoBench.Name = "klNoBench"
        Me.klNoBench.Size = New System.Drawing.Size(534, 41)
        Me.klNoBench.StateNormal.LongText.MultiLine = ComponentFactory.Krypton.Toolkit.InheritBool.[True]
        Me.klNoBench.TabIndex = 4
        Me.klNoBench.TabStop = False
        Me.klNoBench.Values.Text = "Note:  When run with an Effective Date, no benchmarks can be created.  " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Mismatch" & _
            "ed values will appear in RED."
        '
        'kbCancel
        '
        Me.kbCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.kbCancel.Location = New System.Drawing.Point(408, 137)
        Me.kbCancel.Name = "kbCancel"
        Me.kbCancel.Size = New System.Drawing.Size(73, 39)
        Me.kbCancel.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbCancel.StateCommon.Border.Rounding = 6
        Me.kbCancel.StateCommon.Border.Width = 3
        Me.kbCancel.TabIndex = 3
        Me.kbCancel.Values.Text = "Cancel"
        '
        'Form6
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.kbCancel
        Me.ClientSize = New System.Drawing.Size(626, 187)
        Me.Controls.Add(Me.kbCancel)
        Me.Controls.Add(Me.klNoBench)
        Me.Controls.Add(Me.kbOK)
        Me.Controls.Add(Me.kdtpSPIARateDate)
        Me.Controls.Add(Me.krbSPIAEffective)
        Me.Controls.Add(Me.krbSPIACurrent)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
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
    Friend WithEvents kbCancel As ComponentFactory.Krypton.Toolkit.KryptonButton
End Class
