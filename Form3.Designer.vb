<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.klbPages = New ComponentFactory.Krypton.Toolkit.KryptonListBox()
        Me.KryptonLabel1 = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.kbViewPage = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.PictureBoxEx1 = New Tpsc.Controls.PictureBoxEx()
        Me.KryptonLabel4 = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.kbExitView = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.PictureBoxEx2 = New Tpsc.Controls.PictureBoxEx()
        Me.KryptonLabel5 = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.KryptonLabel2 = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.KryptonLabel3 = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.KryptonLabel6 = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.PictureBoxEx1.SuspendLayout()
        Me.PictureBoxEx2.SuspendLayout()
        Me.SuspendLayout()
        '
        'klbPages
        '
        Me.klbPages.BackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.ButtonForm
        Me.klbPages.BorderStyle = ComponentFactory.Krypton.Toolkit.PaletteBorderStyle.ButtonCalendarDay
        Me.klbPages.ItemStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.ListItem
        Me.klbPages.Location = New System.Drawing.Point(561, 14)
        Me.klbPages.Margin = New System.Windows.Forms.Padding(2)
        Me.klbPages.Name = "klbPages"
        Me.klbPages.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.klbPages.Size = New System.Drawing.Size(87, 189)
        Me.klbPages.Sorted = True
        Me.klbPages.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.klbPages.StateCommon.Border.Rounding = 6
        Me.klbPages.StateCommon.Border.Width = 3
        Me.klbPages.TabIndex = 34
        '
        'KryptonLabel1
        '
        Me.KryptonLabel1.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.KryptonLabel1.Location = New System.Drawing.Point(139, 41)
        Me.KryptonLabel1.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonLabel1.Name = "KryptonLabel1"
        Me.KryptonLabel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel1.Size = New System.Drawing.Size(6, 2)
        Me.KryptonLabel1.TabIndex = 35
        Me.KryptonLabel1.Values.ExtraText = ""
        Me.KryptonLabel1.Values.Image = Nothing
        Me.KryptonLabel1.Values.Text = ""
        '
        'kbViewPage
        '
        Me.kbViewPage.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        Me.kbViewPage.Location = New System.Drawing.Point(561, 265)
        Me.kbViewPage.Margin = New System.Windows.Forms.Padding(2)
        Me.kbViewPage.Name = "kbViewPage"
        Me.kbViewPage.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbViewPage.Size = New System.Drawing.Size(87, 33)
        Me.kbViewPage.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbViewPage.StateCommon.Border.Rounding = 6
        Me.kbViewPage.StateCommon.Border.Width = 3
        Me.kbViewPage.StateDisabled.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbViewPage.StateDisabled.Border.Rounding = 6
        Me.kbViewPage.StateDisabled.Border.Width = 3
        Me.kbViewPage.StateNormal.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbViewPage.StateNormal.Border.Rounding = 6
        Me.kbViewPage.StateNormal.Border.Width = 3
        Me.kbViewPage.StatePressed.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbViewPage.StatePressed.Border.Rounding = 6
        Me.kbViewPage.StatePressed.Border.Width = 3
        Me.kbViewPage.StateTracking.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbViewPage.StateTracking.Border.Rounding = 6
        Me.kbViewPage.StateTracking.Border.Width = 3
        Me.kbViewPage.TabIndex = 36
        Me.kbViewPage.Text = "View Page"
        Me.kbViewPage.Values.ExtraText = ""
        Me.kbViewPage.Values.Image = Nothing
        Me.kbViewPage.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbViewPage.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbViewPage.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbViewPage.Values.Text = "View Page"
        Me.kbViewPage.Visible = False
        '
        'PictureBoxEx1
        '
        Me.PictureBoxEx1.AutoScroll = True
        Me.PictureBoxEx1.BackColor = System.Drawing.Color.White
        Me.PictureBoxEx1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBoxEx1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBoxEx1.Controls.Add(Me.KryptonLabel4)
        Me.PictureBoxEx1.Controls.Add(Me.KryptonLabel1)
        Me.PictureBoxEx1.CurrentZoom = 0.1!
        Me.PictureBoxEx1.DefaultZoom = 0.1!
        Me.PictureBoxEx1.DrawMode = System.Drawing.Drawing2D.InterpolationMode.High
        Me.PictureBoxEx1.Location = New System.Drawing.Point(6, 14)
        Me.PictureBoxEx1.Margin = New System.Windows.Forms.Padding(2)
        Me.PictureBoxEx1.MaximumZoom = 1.0!
        Me.PictureBoxEx1.Name = "PictureBoxEx1"
        Me.PictureBoxEx1.Size = New System.Drawing.Size(550, 694)
        Me.PictureBoxEx1.TabIndex = 41
        Me.ToolTip1.SetToolTip(Me.PictureBoxEx1, "Click and drag to zoom; double-click to restore.")
        Me.PictureBoxEx1.Visible = False
        '
        'KryptonLabel4
        '
        Me.KryptonLabel4.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.KryptonLabel4.Location = New System.Drawing.Point(87, 216)
        Me.KryptonLabel4.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonLabel4.Name = "KryptonLabel4"
        Me.KryptonLabel4.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel4.Size = New System.Drawing.Size(6, 2)
        Me.KryptonLabel4.TabIndex = 36
        Me.KryptonLabel4.Values.ExtraText = ""
        Me.KryptonLabel4.Values.Image = Nothing
        Me.KryptonLabel4.Values.Text = ""
        Me.KryptonLabel4.Visible = False
        '
        'kbExitView
        '
        Me.kbExitView.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        Me.kbExitView.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.kbExitView.Location = New System.Drawing.Point(561, 315)
        Me.kbExitView.Margin = New System.Windows.Forms.Padding(2)
        Me.kbExitView.Name = "kbExitView"
        Me.kbExitView.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbExitView.Size = New System.Drawing.Size(87, 33)
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
        Me.kbExitView.TabIndex = 43
        Me.kbExitView.Text = "Close"
        Me.kbExitView.Values.ExtraText = ""
        Me.kbExitView.Values.Image = Nothing
        Me.kbExitView.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbExitView.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbExitView.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbExitView.Values.Text = "Close"
        '
        'PictureBoxEx2
        '
        Me.PictureBoxEx2.AutoScroll = True
        Me.PictureBoxEx2.BackColor = System.Drawing.Color.White
        Me.PictureBoxEx2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBoxEx2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBoxEx2.Controls.Add(Me.KryptonLabel5)
        Me.PictureBoxEx2.CurrentZoom = 0.1!
        Me.PictureBoxEx2.DefaultZoom = 0.1!
        Me.PictureBoxEx2.DrawMode = System.Drawing.Drawing2D.InterpolationMode.High
        Me.PictureBoxEx2.Location = New System.Drawing.Point(653, 14)
        Me.PictureBoxEx2.Margin = New System.Windows.Forms.Padding(2)
        Me.PictureBoxEx2.MaximumZoom = 1.0!
        Me.PictureBoxEx2.Name = "PictureBoxEx2"
        Me.PictureBoxEx2.Size = New System.Drawing.Size(550, 694)
        Me.PictureBoxEx2.TabIndex = 44
        Me.ToolTip1.SetToolTip(Me.PictureBoxEx2, "Click and drag to zoom; double-click to restore.")
        Me.PictureBoxEx2.Visible = False
        '
        'KryptonLabel5
        '
        Me.KryptonLabel5.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.KryptonLabel5.Location = New System.Drawing.Point(146, 216)
        Me.KryptonLabel5.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonLabel5.Name = "KryptonLabel5"
        Me.KryptonLabel5.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel5.Size = New System.Drawing.Size(6, 2)
        Me.KryptonLabel5.TabIndex = 36
        Me.KryptonLabel5.Values.ExtraText = ""
        Me.KryptonLabel5.Values.Image = Nothing
        Me.KryptonLabel5.Values.Text = ""
        Me.KryptonLabel5.Visible = False
        '
        'KryptonLabel2
        '
        Me.KryptonLabel2.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.KryptonLabel2.Location = New System.Drawing.Point(257, 725)
        Me.KryptonLabel2.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonLabel2.Name = "KryptonLabel2"
        Me.KryptonLabel2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel2.Size = New System.Drawing.Size(34, 19)
        Me.KryptonLabel2.TabIndex = 45
        Me.KryptonLabel2.Text = "TEST"
        Me.KryptonLabel2.Values.ExtraText = ""
        Me.KryptonLabel2.Values.Image = Nothing
        Me.KryptonLabel2.Values.Text = "TEST"
        '
        'KryptonLabel3
        '
        Me.KryptonLabel3.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.KryptonLabel3.Location = New System.Drawing.Point(902, 725)
        Me.KryptonLabel3.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonLabel3.Name = "KryptonLabel3"
        Me.KryptonLabel3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel3.Size = New System.Drawing.Size(46, 19)
        Me.KryptonLabel3.TabIndex = 46
        Me.KryptonLabel3.Text = "BENCH"
        Me.KryptonLabel3.Values.ExtraText = ""
        Me.KryptonLabel3.Values.Image = Nothing
        Me.KryptonLabel3.Values.Text = "BENCH"
        '
        'KryptonLabel6
        '
        Me.KryptonLabel6.AutoSize = False
        Me.KryptonLabel6.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.KryptonLabel6.Location = New System.Drawing.Point(342, 714)
        Me.KryptonLabel6.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonLabel6.Name = "KryptonLabel6"
        Me.KryptonLabel6.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel6.Size = New System.Drawing.Size(517, 30)
        Me.KryptonLabel6.StateNormal.LongText.TextH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center
        Me.KryptonLabel6.StateNormal.ShortText.TextH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center
        Me.KryptonLabel6.TabIndex = 47
        Me.KryptonLabel6.Values.ExtraText = ""
        Me.KryptonLabel6.Values.Image = Nothing
        Me.KryptonLabel6.Values.Text = ""
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 3000
        Me.ToolTip1.InitialDelay = 1000
        Me.ToolTip1.ReshowDelay = 100
        Me.ToolTip1.ToolTipTitle = "ZOOM"
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.CancelButton = Me.kbExitView
        Me.ClientSize = New System.Drawing.Size(1212, 753)
        Me.Controls.Add(Me.PictureBoxEx2)
        Me.Controls.Add(Me.KryptonLabel3)
        Me.Controls.Add(Me.kbExitView)
        Me.Controls.Add(Me.KryptonLabel6)
        Me.Controls.Add(Me.PictureBoxEx1)
        Me.Controls.Add(Me.kbViewPage)
        Me.Controls.Add(Me.KryptonLabel2)
        Me.Controls.Add(Me.klbPages)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.Name = "Form3"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "The Regressionator:  View Illustration"
        Me.PictureBoxEx1.ResumeLayout(False)
        Me.PictureBoxEx1.PerformLayout()
        Me.PictureBoxEx2.ResumeLayout(False)
        Me.PictureBoxEx2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents klbPages As ComponentFactory.Krypton.Toolkit.KryptonListBox
    Friend WithEvents KryptonLabel1 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents kbViewPage As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents PictureBoxEx1 As Tpsc.Controls.PictureBoxEx
    'Friend WithEvents KryptonLabel2 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents kbExitView As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents PictureBoxEx2 As Tpsc.Controls.PictureBoxEx
    'Friend WithEvents KryptonLabel3 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents KryptonLabel4 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents KryptonLabel5 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents KryptonLabel2 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents KryptonLabel3 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents KryptonLabel6 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
