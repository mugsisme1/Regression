<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form7
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form7))
        Me.krbFIASaveAgeYes = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton()
        Me.krbFIASaveAgeNo = New ComponentFactory.Krypton.Toolkit.KryptonRadioButton()
        Me.klVASaveAge = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.klVASaveAgeNote = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.kbOK = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.SuspendLayout()
        '
        'krbFIASaveAgeYes
        '
        Me.krbFIASaveAgeYes.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.krbFIASaveAgeYes.Location = New System.Drawing.Point(246, 10)
        Me.krbFIASaveAgeYes.Margin = New System.Windows.Forms.Padding(2)
        Me.krbFIASaveAgeYes.Name = "krbFIASaveAgeYes"
        Me.krbFIASaveAgeYes.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.krbFIASaveAgeYes.Size = New System.Drawing.Size(39, 19)
        Me.krbFIASaveAgeYes.TabIndex = 0
        Me.krbFIASaveAgeYes.Text = "Yes"
        Me.krbFIASaveAgeYes.Values.ExtraText = ""
        Me.krbFIASaveAgeYes.Values.Image = Nothing
        Me.krbFIASaveAgeYes.Values.Text = "Yes"
        '
        'krbFIASaveAgeNo
        '
        Me.krbFIASaveAgeNo.Checked = True
        Me.krbFIASaveAgeNo.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.krbFIASaveAgeNo.Location = New System.Drawing.Point(340, 10)
        Me.krbFIASaveAgeNo.Margin = New System.Windows.Forms.Padding(2)
        Me.krbFIASaveAgeNo.Name = "krbFIASaveAgeNo"
        Me.krbFIASaveAgeNo.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.krbFIASaveAgeNo.Size = New System.Drawing.Size(37, 19)
        Me.krbFIASaveAgeNo.TabIndex = 1
        Me.krbFIASaveAgeNo.Text = "No"
        Me.krbFIASaveAgeNo.Values.ExtraText = ""
        Me.krbFIASaveAgeNo.Values.Image = Nothing
        Me.krbFIASaveAgeNo.Values.Text = "No"
        '
        'klVASaveAge
        '
        Me.klVASaveAge.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.klVASaveAge.Location = New System.Drawing.Point(32, 10)
        Me.klVASaveAge.Margin = New System.Windows.Forms.Padding(2)
        Me.klVASaveAge.Name = "klVASaveAge"
        Me.klVASaveAge.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.klVASaveAge.Size = New System.Drawing.Size(140, 19)
        Me.klVASaveAge.TabIndex = 2
        Me.klVASaveAge.TabStop = False
        Me.klVASaveAge.Text = "Change DOB to save age?"
        Me.klVASaveAge.Values.ExtraText = ""
        Me.klVASaveAge.Values.Image = Nothing
        Me.klVASaveAge.Values.Text = "Change DOB to save age?"
        '
        'klVASaveAgeNote
        '
        Me.klVASaveAgeNote.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.klVASaveAgeNote.Location = New System.Drawing.Point(32, 37)
        Me.klVASaveAgeNote.Margin = New System.Windows.Forms.Padding(2)
        Me.klVASaveAgeNote.Name = "klVASaveAgeNote"
        Me.klVASaveAgeNote.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.klVASaveAgeNote.Size = New System.Drawing.Size(312, 33)
        Me.klVASaveAgeNote.TabIndex = 3
        Me.klVASaveAgeNote.TabStop = False
        Me.klVASaveAgeNote.Text = "Selecting ""Yes"" will run the cases with the benchmarked age.  " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Mismatched values" & _
    " will appear in green."
        Me.klVASaveAgeNote.Values.ExtraText = ""
        Me.klVASaveAgeNote.Values.Image = Nothing
        Me.klVASaveAgeNote.Values.Text = "Selecting ""Yes"" will run the cases with the benchmarked age.  " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Mismatched values" & _
    " will appear in green."
        '
        'kbOK
        '
        Me.kbOK.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Standalone
        Me.kbOK.Location = New System.Drawing.Point(372, 46)
        Me.kbOK.Margin = New System.Windows.Forms.Padding(2)
        Me.kbOK.Name = "kbOK"
        Me.kbOK.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbOK.Size = New System.Drawing.Size(59, 46)
        Me.kbOK.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbOK.StateCommon.Border.Rounding = 6
        Me.kbOK.StateCommon.Border.Width = 3
        Me.kbOK.TabIndex = 3
        Me.kbOK.Text = "OK"
        Me.kbOK.Values.ExtraText = ""
        Me.kbOK.Values.Image = Nothing
        Me.kbOK.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbOK.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbOK.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbOK.Values.Text = "OK"
        '
        'Form7
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(442, 94)
        Me.Controls.Add(Me.kbOK)
        Me.Controls.Add(Me.klVASaveAgeNote)
        Me.Controls.Add(Me.klVASaveAge)
        Me.Controls.Add(Me.krbFIASaveAgeNo)
        Me.Controls.Add(Me.krbFIASaveAgeYes)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form7"
        Me.Text = "The Regressionator:  Save FIA Age?"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents krbFIASaveAgeYes As ComponentFactory.Krypton.Toolkit.KryptonRadioButton
    Friend WithEvents krbFIASaveAgeNo As ComponentFactory.Krypton.Toolkit.KryptonRadioButton
    Friend WithEvents klVASaveAge As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents klVASaveAgeNote As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents kbOK As ComponentFactory.Krypton.Toolkit.KryptonButton
End Class
