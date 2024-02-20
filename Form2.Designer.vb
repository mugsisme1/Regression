<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Me.KryptonWorkspace1 = New ComponentFactory.Krypton.Workspace.KryptonWorkspace
        CType(Me.KryptonWorkspace1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'KryptonWorkspace1
        '
        Me.KryptonWorkspace1.Location = New System.Drawing.Point(17, 15)
        Me.KryptonWorkspace1.Name = "KryptonWorkspace1"
        Me.KryptonWorkspace1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.[Global]
        Me.KryptonWorkspace1.Size = New System.Drawing.Size(258, 202)
        Me.KryptonWorkspace1.TabIndex = 0
        Me.KryptonWorkspace1.TabStop = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 255)
        Me.Controls.Add(Me.KryptonWorkspace1)
        Me.Name = "Form2"
        Me.Text = "Form2"
        CType(Me.KryptonWorkspace1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents KryptonWorkspace1 As ComponentFactory.Krypton.Workspace.KryptonWorkspace
End Class
