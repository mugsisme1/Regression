Public Class Form7

    Private Sub Form7_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToParent()
    End Sub
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim CS_NOCLOSE As Integer = Int32.Parse("200", Globalization.NumberStyles.HexNumber)
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = CS_NOCLOSE
            Return cp
        End Get
    End Property

    Private Sub kbOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbOK.Click

        Me.Close()
        Me.Dispose()
        'show the main form
        System.Windows.Forms.Application.DoEvents()
        Regression.RegressionMain.Show()

    End Sub

    Private Sub kbCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Regression.RegressionMain.gbVASaveAgeCancel = True
        Me.Close()
    End Sub
End Class