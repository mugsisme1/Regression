Imports System.Windows.Forms.Form
Imports System.Runtime.InteropServices

Public Class progress
    Private Sub SmoothProgressBar1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SmoothProgressBar1.Load

    End Sub

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim CS_NOCLOSE As Integer = Int32.Parse("200", Globalization.NumberStyles.HexNumber)
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = CS_NOCLOSE
            Return cp
        End Get
    End Property

    Protected Overrides Sub OnMove(ByVal e As System.EventArgs)
        'Do Nothing
    End Sub

    Private Sub progress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub
End Class