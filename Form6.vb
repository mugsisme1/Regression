Public Class Form6

    Private Sub klSPIAEffective_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles krbSPIAEffective.CheckedChanged
        kdtpSPIARateDate.Visible = True
    End Sub

    Public Sub kbOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbOK.Click

        Me.Close()


        'show the main form
        System.Windows.Forms.Application.DoEvents()
        Regression.RegressionMain.Show()

    End Sub

    Public Sub Form6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    Private Sub klSPIACurrent_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles krbSPIACurrent.CheckedChanged
        kdtpSPIARateDate.Visible = False
    End Sub

    Private Sub klNoBench_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles klNoBench.Paint

    End Sub

    Public Sub kbCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


        Regression.RegressionMain.gbRateCancel = True

        Me.Close()


    End Sub
End Class