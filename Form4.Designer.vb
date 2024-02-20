<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form4
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form4))
        Me.kclbNewBench = New ComponentFactory.Krypton.Toolkit.KryptonCheckedListBox
        Me.KryptonButton1 = New ComponentFactory.Krypton.Toolkit.KryptonButton
        Me.KryptonButton2 = New ComponentFactory.Krypton.Toolkit.KryptonButton
        Me.KryptonLabel6 = New ComponentFactory.Krypton.Toolkit.KryptonLabel
        Me.KryptonButton3 = New ComponentFactory.Krypton.Toolkit.KryptonButton
        Me.kbExitView = New ComponentFactory.Krypton.Toolkit.KryptonButton
        Me.SuspendLayout()
        '
        'kclbNewBench
        '
        Me.kclbNewBench.BackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.ButtonForm
        Me.kclbNewBench.BorderStyle = ComponentFactory.Krypton.Toolkit.PaletteBorderStyle.ButtonCalendarDay
        Me.kclbNewBench.CausesValidation = False
        Me.kclbNewBench.CheckOnClick = True
        Me.kclbNewBench.ItemStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.ListItem
        resources.ApplyResources(Me.kclbNewBench, "kclbNewBench")
        Me.kclbNewBench.Name = "kclbNewBench"
        Me.kclbNewBench.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kclbNewBench.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kclbNewBench.StateCommon.Border.Rounding = 6
        Me.kclbNewBench.StateCommon.Border.Width = 3
        Me.kclbNewBench.StateNormal.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kclbNewBench.StateNormal.Border.Rounding = 6
        Me.kclbNewBench.StateNormal.Border.Width = 3
        '
        'KryptonButton1
        '
        Me.KryptonButton1.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        resources.ApplyResources(Me.KryptonButton1, "KryptonButton1")
        Me.KryptonButton1.Name = "KryptonButton1"
        Me.KryptonButton1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonButton1.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.KryptonButton1.StateCommon.Border.Rounding = 6
        Me.KryptonButton1.StateCommon.Border.Width = 3
        Me.KryptonButton1.Values.ExtraText = resources.GetString("KryptonButton1.Values.ExtraText")
        Me.KryptonButton1.Values.Image = Nothing
        Me.KryptonButton1.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.KryptonButton1.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.KryptonButton1.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.KryptonButton1.Values.Text = resources.GetString("KryptonButton1.Values.Text")
        '
        'KryptonButton2
        '
        Me.KryptonButton2.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        resources.ApplyResources(Me.KryptonButton2, "KryptonButton2")
        Me.KryptonButton2.Name = "KryptonButton2"
        Me.KryptonButton2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonButton2.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.KryptonButton2.StateCommon.Border.Rounding = 6
        Me.KryptonButton2.StateCommon.Border.Width = 3
        Me.KryptonButton2.Values.ExtraText = resources.GetString("KryptonButton2.Values.ExtraText")
        Me.KryptonButton2.Values.Image = Nothing
        Me.KryptonButton2.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.KryptonButton2.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.KryptonButton2.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.KryptonButton2.Values.Text = resources.GetString("KryptonButton2.Values.Text")
        '
        'KryptonLabel6
        '
        Me.KryptonLabel6.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        resources.ApplyResources(Me.KryptonLabel6, "KryptonLabel6")
        Me.KryptonLabel6.Name = "KryptonLabel6"
        Me.KryptonLabel6.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel6.StateNormal.ShortText.TextH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center
        Me.KryptonLabel6.Values.ExtraText = resources.GetString("KryptonLabel6.Values.ExtraText")
        Me.KryptonLabel6.Values.Image = Nothing
        Me.KryptonLabel6.Values.Text = resources.GetString("KryptonLabel6.Values.Text")
        '
        'KryptonButton3
        '
        Me.KryptonButton3.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        resources.ApplyResources(Me.KryptonButton3, "KryptonButton3")
        Me.KryptonButton3.Name = "KryptonButton3"
        Me.KryptonButton3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonButton3.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.KryptonButton3.StateCommon.Border.Rounding = 6
        Me.KryptonButton3.StateCommon.Border.Width = 3
        Me.KryptonButton3.Values.ExtraText = resources.GetString("KryptonButton3.Values.ExtraText")
        Me.KryptonButton3.Values.Image = Nothing
        Me.KryptonButton3.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.KryptonButton3.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.KryptonButton3.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.KryptonButton3.Values.Text = resources.GetString("KryptonButton3.Values.Text")
        '
        'kbExitView
        '
        Me.kbExitView.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        Me.kbExitView.DialogResult = System.Windows.Forms.DialogResult.Cancel
        resources.ApplyResources(Me.kbExitView, "kbExitView")
        Me.kbExitView.Name = "kbExitView"
        Me.kbExitView.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbExitView.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbExitView.StateCommon.Border.Rounding = 6
        Me.kbExitView.StateCommon.Border.Width = 3
        Me.kbExitView.StateDisabled.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbExitView.StateDisabled.Border.Rounding = 6
        Me.kbExitView.StateDisabled.Border.Width = 3
        Me.kbExitView.StateNormal.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbExitView.StateNormal.Border.Rounding = 6
        Me.kbExitView.StateNormal.Border.Width = 3
        Me.kbExitView.StatePressed.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbExitView.StatePressed.Border.Rounding = 6
        Me.kbExitView.StatePressed.Border.Width = 3
        Me.kbExitView.StateTracking.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
                    Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbExitView.StateTracking.Border.Rounding = 6
        Me.kbExitView.StateTracking.Border.Width = 3
        Me.kbExitView.Values.ExtraText = resources.GetString("kbExitView.Values.ExtraText")
        Me.kbExitView.Values.Image = Nothing
        Me.kbExitView.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbExitView.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbExitView.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbExitView.Values.Text = resources.GetString("kbExitView.Values.Text")
        '
        'Form4
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.CancelButton = Me.kbExitView
        Me.Controls.Add(Me.kbExitView)
        Me.Controls.Add(Me.KryptonButton3)
        Me.Controls.Add(Me.KryptonLabel6)
        Me.Controls.Add(Me.KryptonButton2)
        Me.Controls.Add(Me.KryptonButton1)
        Me.Controls.Add(Me.kclbNewBench)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Form4"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub Form4_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim cCtrl As Control
        For Each cCtrl In Regression.RegressionMain.Controls
            cCtrl.Enabled = True
        Next
        Regression.RegressionMain.Show()

    End Sub
    Friend WithEvents kclbNewBench As ComponentFactory.Krypton.Toolkit.KryptonCheckedListBox
    Friend WithEvents KryptonButton1 As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents KryptonButton2 As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents KryptonLabel6 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents KryptonButton3 As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents kbExitView As ComponentFactory.Krypton.Toolkit.KryptonButton
End Class
