<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MismatchNotes
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MismatchNotes))
        Me.KryptonLabel3 = New ComponentFactory.Krypton.Toolkit.KryptonLabel()
        Me.KryptonRichTextBox1 = New ComponentFactory.Krypton.Toolkit.KryptonRichTextBox()
        Me.SuspendLayout()
        '
        'KryptonLabel3
        '
        Me.KryptonLabel3.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl
        Me.KryptonLabel3.Location = New System.Drawing.Point(37, 24)
        Me.KryptonLabel3.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonLabel3.Name = "KryptonLabel3"
        Me.KryptonLabel3.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonLabel3.Size = New System.Drawing.Size(207, 19)
        Me.KryptonLabel3.TabIndex = 38
        Me.KryptonLabel3.Text = "Enter notes for case #...mismatch status"
        Me.KryptonLabel3.Values.ExtraText = ""
        Me.KryptonLabel3.Values.Image = Nothing
        Me.KryptonLabel3.Values.Text = "Enter notes for case #...mismatch status"
        '
        'KryptonRichTextBox1
        '
        Me.KryptonRichTextBox1.InputControlStyle = ComponentFactory.Krypton.Toolkit.InputControlStyle.Standalone
        Me.KryptonRichTextBox1.Location = New System.Drawing.Point(37, 60)
        Me.KryptonRichTextBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.KryptonRichTextBox1.Name = "KryptonRichTextBox1"
        Me.KryptonRichTextBox1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonRichTextBox1.Size = New System.Drawing.Size(393, 138)
        Me.KryptonRichTextBox1.TabIndex = 39
        Me.KryptonRichTextBox1.Text = "KryptonRichTextBox1"
        '
        'MismatchNotes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(581, 207)
        Me.Controls.Add(Me.KryptonRichTextBox1)
        Me.Controls.Add(Me.KryptonLabel3)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "MismatchNotes"
        Me.Text = "The Regressionator:  Mismatch Notes"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents KryptonLabel3 As ComponentFactory.Krypton.Toolkit.KryptonLabel
    Friend WithEvents KryptonRichTextBox1 As ComponentFactory.Krypton.Toolkit.KryptonRichTextBox
End Class
