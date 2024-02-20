<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class progress
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(progress))
        Me.SmoothProgressBar1 = New SmoothProgressBar.SmoothProgressBar
        Me.SuspendLayout()
        '
        'SmoothProgressBar1
        '
        Me.SmoothProgressBar1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.SmoothProgressBar1.BackColor = System.Drawing.Color.Transparent
        Me.SmoothProgressBar1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.SmoothProgressBar1.Location = New System.Drawing.Point(18, 16)
        Me.SmoothProgressBar1.Maximum = 100
        Me.SmoothProgressBar1.Minimum = 0
        Me.SmoothProgressBar1.Name = "SmoothProgressBar1"
        Me.SmoothProgressBar1.ProgressBarColor = System.Drawing.Color.SteelBlue
        Me.SmoothProgressBar1.Size = New System.Drawing.Size(442, 25)
        Me.SmoothProgressBar1.TabIndex = 36
        Me.SmoothProgressBar1.TabStop = False
        Me.SmoothProgressBar1.Value = 0
        '
        'progress
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.BackColor = System.Drawing.Color.Gainsboro
        Me.ClientSize = New System.Drawing.Size(479, 57)
        Me.Controls.Add(Me.SmoothProgressBar1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "progress"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "The Regressionator:  Run Progress"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SmoothProgressBar1 As SmoothProgressBar.SmoothProgressBar
End Class
