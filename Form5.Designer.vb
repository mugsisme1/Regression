<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form5
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form5))
        Me.kbCompare = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.kbClose = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.KryptonLabel6 = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.kbBench = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.kdgvPalomaStatus = New ComponentFactory.Krypton.Toolkit.KryptonDataGridView()
        Me.Column7 = New ComponentFactory.Krypton.Toolkit.KryptonDataGridViewCheckBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.kbView = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.KryptonBorderEdge1 = New ComponentFactory.Krypton.Toolkit.KryptonBorderEdge()
        Me.KryptonBorderEdge2 = New ComponentFactory.Krypton.Toolkit.KryptonBorderEdge()
        Me.KryptonBorderEdge3 = New ComponentFactory.Krypton.Toolkit.KryptonBorderEdge()
        Me.kbAll = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.kbNone = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.KryptonLabel1 = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.KryptonBorderEdge4 = New ComponentFactory.Krypton.Toolkit.KryptonBorderEdge()
        CType(Me.kdgvPalomaStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'kbCompare
        '
        Me.kbCompare.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        Me.kbCompare.Location = New System.Drawing.Point(214, 224)
        Me.kbCompare.Name = "kbCompare"
        Me.kbCompare.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbCompare.Size = New System.Drawing.Size(106, 32)
        Me.kbCompare.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbCompare.StateCommon.Border.Rounding = 6
        Me.kbCompare.StateCommon.Border.Width = 3
        Me.kbCompare.TabIndex = 2
        Me.kbCompare.Text = "Compare"
        Me.kbCompare.Values.ExtraText = ""
        Me.kbCompare.Values.Image = Nothing
        Me.kbCompare.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbCompare.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbCompare.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbCompare.Values.Text = "Compare"
        '
        'kbClose
        '
        Me.kbClose.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        Me.kbClose.Location = New System.Drawing.Point(994, 224)
        Me.kbClose.Name = "kbClose"
        Me.kbClose.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbClose.Size = New System.Drawing.Size(147, 32)
        Me.kbClose.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbClose.StateCommon.Border.Rounding = 6
        Me.kbClose.StateCommon.Border.Width = 3
        Me.kbClose.TabIndex = 5
        Me.kbClose.Text = "Close"
        Me.kbClose.Values.ExtraText = ""
        Me.kbClose.Values.Image = Nothing
        Me.kbClose.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbClose.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbClose.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbClose.Values.Text = "Close"
        '
        'KryptonLabel6
        '
        Me.KryptonLabel6.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.SuperTip
        Me.KryptonLabel6.Location = New System.Drawing.Point(20, 18)
        Me.KryptonLabel6.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonLabel6.Name = "KryptonLabel6"
        Me.KryptonLabel6.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel6.Size = New System.Drawing.Size(8, 8)
        Me.KryptonLabel6.StateNormal.ShortText.TextH = ComponentFactory.Krypton.Toolkit.PaletteRelativeAlign.Center
        Me.KryptonLabel6.TabIndex = 49
        Me.KryptonLabel6.TabStop = False
        Me.KryptonLabel6.Values.ExtraText = ""
        Me.KryptonLabel6.Values.Image = Nothing
        Me.KryptonLabel6.Values.Text = ""
        '
        'kbBench
        '
        Me.kbBench.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Standalone
        Me.kbBench.Location = New System.Drawing.Point(716, 224)
        Me.kbBench.Name = "kbBench"
        Me.kbBench.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbBench.Size = New System.Drawing.Size(128, 32)
        Me.kbBench.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbBench.StateCommon.Border.Rounding = 6
        Me.kbBench.StateCommon.Border.Width = 3
        Me.kbBench.TabIndex = 3
        Me.kbBench.Text = "Create Benchmark"
        Me.kbBench.Values.ExtraText = ""
        Me.kbBench.Values.Image = Nothing
        Me.kbBench.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbBench.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbBench.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbBench.Values.Text = "Create Benchmark"
        '
        'kdgvPalomaStatus
        '
        Me.kdgvPalomaStatus.AllowUserToAddRows = False
        Me.kdgvPalomaStatus.AllowUserToDeleteRows = False
        Me.kdgvPalomaStatus.AllowUserToOrderColumns = True
        Me.kdgvPalomaStatus.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column7, Me.Column1, Me.Column2, Me.Column4, Me.Column3, Me.Column5, Me.Column6})
        Me.kdgvPalomaStatus.GridStyles.Style = ComponentFactory.Krypton.Toolkit.DataGridViewStyle.Mixed
        Me.kdgvPalomaStatus.GridStyles.StyleBackground = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.GridBackgroundList
        Me.kdgvPalomaStatus.GridStyles.StyleColumn = ComponentFactory.Krypton.Toolkit.GridStyle.Sheet
        Me.kdgvPalomaStatus.GridStyles.StyleDataCells = ComponentFactory.Krypton.Toolkit.GridStyle.List
        Me.kdgvPalomaStatus.GridStyles.StyleRow = ComponentFactory.Krypton.Toolkit.GridStyle.List
        Me.kdgvPalomaStatus.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.kdgvPalomaStatus.Location = New System.Drawing.Point(20, 49)
        Me.kdgvPalomaStatus.Margin = New System.Windows.Forms.Padding(2)
        Me.kdgvPalomaStatus.MultiSelect = False
        Me.kdgvPalomaStatus.Name = "kdgvPalomaStatus"
        Me.kdgvPalomaStatus.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kdgvPalomaStatus.RowHeadersVisible = False
        Me.kdgvPalomaStatus.RowTemplate.Height = 24
        Me.kdgvPalomaStatus.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.kdgvPalomaStatus.Size = New System.Drawing.Size(1121, 123)
        Me.kdgvPalomaStatus.StateCommon.BackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.GridBackgroundList
        Me.kdgvPalomaStatus.StateSelected.DataCell.Back.Color1 = System.Drawing.Color.White
        Me.kdgvPalomaStatus.TabIndex = 1
        '
        'Column7
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.NullValue = False
        Me.Column7.DefaultCellStyle = DataGridViewCellStyle1
        Me.Column7.FalseValue = Nothing
        Me.Column7.HeaderText = "Select"
        Me.Column7.IndeterminateValue = Nothing
        Me.Column7.Name = "Column7"
        Me.Column7.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column7.TrueValue = Nothing
        Me.Column7.Width = 75
        '
        'Column1
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        Me.Column1.DefaultCellStyle = DataGridViewCellStyle2
        Me.Column1.FillWeight = 29.46428!
        Me.Column1.HeaderText = "Client"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 55
        '
        'Column2
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        Me.Column2.DefaultCellStyle = DataGridViewCellStyle3
        Me.Column2.FillWeight = 73.03701!
        Me.Column2.HeaderText = "Status"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 190
        '
        'Column4
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        Me.Column4.DefaultCellStyle = DataGridViewCellStyle4
        Me.Column4.FillWeight = 139.5838!
        Me.Column4.HeaderText = "Last Paloma Compare"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        Me.Column4.ToolTipText = "* means that date will be updated next run"
        Me.Column4.Width = 200
        '
        'Column3
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        Me.Column3.DefaultCellStyle = DataGridViewCellStyle5
        Me.Column3.FillWeight = 127.8951!
        Me.Column3.HeaderText = "Last Paloma Benchmark"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.ToolTipText = "* means that date will be updated next run"
        Me.Column3.Width = 200
        '
        'Column5
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        Me.Column5.DefaultCellStyle = DataGridViewCellStyle6
        Me.Column5.FillWeight = 118.659!
        Me.Column5.HeaderText = "Last Reg. Compare"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        Me.Column5.ToolTipText = "Last Regressionator Compare"
        Me.Column5.Width = 200
        '
        'Column6
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        Me.Column6.DefaultCellStyle = DataGridViewCellStyle7
        Me.Column6.FillWeight = 111.3607!
        Me.Column6.HeaderText = "Last Reg. Benchmark"
        Me.Column6.Name = "Column6"
        Me.Column6.ReadOnly = True
        Me.Column6.ToolTipText = "Last Regressionator Benchmark made"
        Me.Column6.Width = 200
        '
        'kbView
        '
        Me.kbView.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        Me.kbView.Location = New System.Drawing.Point(460, 224)
        Me.kbView.Name = "kbView"
        Me.kbView.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbView.Size = New System.Drawing.Size(106, 32)
        Me.kbView.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbView.StateCommon.Border.Rounding = 6
        Me.kbView.StateCommon.Border.Width = 3
        Me.kbView.TabIndex = 4
        Me.kbView.Text = "View"
        Me.kbView.Values.ExtraText = ""
        Me.kbView.Values.Image = Nothing
        Me.kbView.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbView.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbView.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbView.Values.Text = "View"
        '
        'KryptonBorderEdge1
        '
        Me.KryptonBorderEdge1.BorderStyle = ComponentFactory.Krypton.Toolkit.PaletteBorderStyle.ControlClient
        Me.KryptonBorderEdge1.Location = New System.Drawing.Point(151, 206)
        Me.KryptonBorderEdge1.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonBorderEdge1.Name = "KryptonBorderEdge1"
        Me.KryptonBorderEdge1.Orientation = System.Windows.Forms.Orientation.Vertical
        Me.KryptonBorderEdge1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonBorderEdge1.Size = New System.Drawing.Size(1, 50)
        Me.KryptonBorderEdge1.TabIndex = 66
        Me.KryptonBorderEdge1.Text = "KryptonBorderEdge1"
        '
        'KryptonBorderEdge2
        '
        Me.KryptonBorderEdge2.BorderStyle = ComponentFactory.Krypton.Toolkit.PaletteBorderStyle.ControlClient
        Me.KryptonBorderEdge2.Location = New System.Drawing.Point(406, 206)
        Me.KryptonBorderEdge2.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonBorderEdge2.Name = "KryptonBorderEdge2"
        Me.KryptonBorderEdge2.Orientation = System.Windows.Forms.Orientation.Vertical
        Me.KryptonBorderEdge2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonBorderEdge2.Size = New System.Drawing.Size(1, 50)
        Me.KryptonBorderEdge2.TabIndex = 65
        Me.KryptonBorderEdge2.Text = "KryptonBorderEdge2"
        '
        'KryptonBorderEdge3
        '
        Me.KryptonBorderEdge3.BorderStyle = ComponentFactory.Krypton.Toolkit.PaletteBorderStyle.ControlClient
        Me.KryptonBorderEdge3.Location = New System.Drawing.Point(623, 206)
        Me.KryptonBorderEdge3.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonBorderEdge3.Name = "KryptonBorderEdge3"
        Me.KryptonBorderEdge3.Orientation = System.Windows.Forms.Orientation.Vertical
        Me.KryptonBorderEdge3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonBorderEdge3.Size = New System.Drawing.Size(1, 50)
        Me.KryptonBorderEdge3.TabIndex = 64
        Me.KryptonBorderEdge3.Text = "KryptonBorderEdge3"
        '
        'kbAll
        '
        Me.kbAll.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        Me.kbAll.Location = New System.Drawing.Point(21, 206)
        Me.kbAll.Margin = New System.Windows.Forms.Padding(2)
        Me.kbAll.Name = "kbAll"
        Me.kbAll.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbAll.Size = New System.Drawing.Size(69, 28)
        Me.kbAll.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbAll.StateCommon.Border.Rounding = 6
        Me.kbAll.StateCommon.Border.Width = 3
        Me.kbAll.TabIndex = 0
        Me.kbAll.Text = "All"
        Me.kbAll.Values.ExtraText = ""
        Me.kbAll.Values.Image = Nothing
        Me.kbAll.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbAll.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbAll.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbAll.Values.Text = "All"
        '
        'kbNone
        '
        Me.kbNone.ButtonStyle = ComponentFactory.Krypton.Toolkit.ButtonStyle.Gallery
        Me.kbNone.Location = New System.Drawing.Point(20, 245)
        Me.kbNone.Margin = New System.Windows.Forms.Padding(2)
        Me.kbNone.Name = "kbNone"
        Me.kbNone.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.kbNone.Size = New System.Drawing.Size(69, 28)
        Me.kbNone.StateCommon.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.kbNone.StateCommon.Border.Rounding = 6
        Me.kbNone.StateCommon.Border.Width = 3
        Me.kbNone.TabIndex = 1
        Me.kbNone.Text = "None"
        Me.kbNone.Values.ExtraText = ""
        Me.kbNone.Values.Image = Nothing
        Me.kbNone.Values.ImageStates.ImageCheckedNormal = Nothing
        Me.kbNone.Values.ImageStates.ImageCheckedPressed = Nothing
        Me.kbNone.Values.ImageStates.ImageCheckedTracking = Nothing
        Me.kbNone.Values.Text = "None"
        '
        'KryptonLabel1
        '
        Me.KryptonLabel1.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.KryptonLabel1.Location = New System.Drawing.Point(21, 183)
        Me.KryptonLabel1.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonLabel1.Name = "KryptonLabel1"
        Me.KryptonLabel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel1.Size = New System.Drawing.Size(42, 19)
        Me.KryptonLabel1.TabIndex = 63
        Me.KryptonLabel1.TabStop = False
        Me.KryptonLabel1.Text = "Select:"
        Me.KryptonLabel1.Values.ExtraText = ""
        Me.KryptonLabel1.Values.Image = Nothing
        Me.KryptonLabel1.Values.Text = "Select:"
        '
        'KryptonBorderEdge4
        '
        Me.KryptonBorderEdge4.BorderStyle = ComponentFactory.Krypton.Toolkit.PaletteBorderStyle.ControlClient
        Me.KryptonBorderEdge4.Location = New System.Drawing.Point(935, 206)
        Me.KryptonBorderEdge4.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonBorderEdge4.Name = "KryptonBorderEdge4"
        Me.KryptonBorderEdge4.Orientation = System.Windows.Forms.Orientation.Vertical
        Me.KryptonBorderEdge4.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonBorderEdge4.Size = New System.Drawing.Size(1, 50)
        Me.KryptonBorderEdge4.TabIndex = 0
        Me.KryptonBorderEdge4.Text = "KryptonBorderEdge4"
        '
        'Form5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1161, 284)
        Me.Controls.Add(Me.KryptonBorderEdge4)
        Me.Controls.Add(Me.KryptonLabel1)
        Me.Controls.Add(Me.kbNone)
        Me.Controls.Add(Me.kbAll)
        Me.Controls.Add(Me.KryptonBorderEdge3)
        Me.Controls.Add(Me.KryptonBorderEdge2)
        Me.Controls.Add(Me.KryptonBorderEdge1)
        Me.Controls.Add(Me.kbView)
        Me.Controls.Add(Me.kdgvPalomaStatus)
        Me.Controls.Add(Me.kbBench)
        Me.Controls.Add(Me.KryptonLabel6)
        Me.Controls.Add(Me.kbClose)
        Me.Controls.Add(Me.kbCompare)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form5"
        Me.Text = "  The Regressionator:  Paloma Compare"
        CType(Me.kdgvPalomaStatus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents kbCompare As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents kbClose As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents KryptonLabel6 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents kbBench As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents kdgvPalomaStatus As ComponentFactory.Krypton.Toolkit.KryptonDataGridView
    Friend WithEvents kbView As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents KryptonBorderEdge1 As ComponentFactory.Krypton.Toolkit.KryptonBorderEdge
    Friend WithEvents KryptonBorderEdge2 As ComponentFactory.Krypton.Toolkit.KryptonBorderEdge
    Friend WithEvents KryptonBorderEdge3 As ComponentFactory.Krypton.Toolkit.KryptonBorderEdge
    Friend WithEvents kbAll As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents kbNone As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents KryptonLabel1 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents KryptonBorderEdge4 As ComponentFactory.Krypton.Toolkit.KryptonBorderEdge
    Friend WithEvents Column7 As ComponentFactory.Krypton.Toolkit.KryptonDataGridViewCheckBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
