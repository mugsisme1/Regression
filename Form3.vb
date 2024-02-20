Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
'Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Nini.Config
Imports Scripting
Imports System.Drawing.Imaging
Imports System.IO.Path
Imports System.Drawing
Imports System
Public Class Form3
    Public Shared s As String
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim CS_NOCLOSE As Integer = Int32.Parse("200", Globalization.NumberStyles.HexNumber)
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = CS_NOCLOSE
            Return cp
        End Get
    End Property
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Form3_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        'clear out the picture boxes
        PictureBoxEx1.Image = Nothing
        PictureBoxEx1.Image = Nothing

        'reenable the controls on the main form
        Dim cCtrl As Control
        For Each cCtrl In Regression.RegressionMain.Controls
            cCtrl.Enabled = True
        Next

        'show the main form
        Regression.RegressionMain.Show()

    End Sub

    Public Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToParent()
        'on load, hide the unneeded controls
        KryptonLabel4.Visible = False
        KryptonLabel4.Text = ""

        KryptonLabel5.Visible = False
        KryptonLabel5.Text = ""

    End Sub
    Public Sub ViewPage()
        If klbPages.SelectedIndex = -1 Then
            MsgBox("Please select a page to view.")
            Return
        End If

        'set the directory for the illustration pages
        Dim di As DirectoryInfo = New DirectoryInfo(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\new" & Regression.RegressionMain.klbMismatchedClients.SelectedItem)

        'get the .emf pages
        Dim emf As FileInfo() = di.GetFiles("*.emf")

        'Dim PrintPreviewControl1 As New PrintPreviewControl
        Dim i As Integer = Regression.RegressionMain.klbMismatchedClients.SelectedItem

        s = klbPages.SelectedItem

        'if pages are 1-9, remove leading 0 so they will sort correctly
        If s.Substring(4, 1) = "0" Then
            s = s.Remove(4, 1)
        End If


        'add extension back to page #
        s = s & ".emf"

        'reset the picture boxes
        PictureBoxEx1.Visible = False
        PictureBoxEx2.Visible = False
        PictureBoxEx1.Image = Nothing
        PictureBoxEx2.Image = Nothing


        'Fill the picture box with the illustration page if it exists, else put up message
        If FileIO.FileSystem.FileExists(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\new" & i & "\" & s) Then
            PictureBoxEx1.Visible = True
            PictureBoxEx1.Image = Image.FromFile(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\new" & i & "\" & s)
            KryptonLabel4.Visible = False
        Else
            PictureBoxEx1.Visible = True
            KryptonLabel4.Visible = True
            KryptonLabel4.Text = "There is no illustration page"
        End If

        If FileIO.FileSystem.FileExists(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\" & i & "\" & s) Then
            KryptonLabel5.Visible = False
            PictureBoxEx2.Visible = True
            PictureBoxEx2.Image = Image.FromFile(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\" & i & "\" & s)
        Else
            PictureBoxEx2.Visible = True
            KryptonLabel5.Visible = True
            KryptonLabel5.Text = "There is no illustration page"

        End If

        'set the labels
        KryptonLabel6.Text = Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & " Client # " & Regression.RegressionMain.klbMismatchedClients.SelectedItem & " " & klbPages.SelectedItem

    End Sub
    Public Sub kbviewpage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbViewPage.Click


        If klbPages.SelectedIndex = -1 Then
            MsgBox("Please select a page to view.")
            Return
        End If

        'set the directory for the illustration pages
        Dim di As DirectoryInfo = New DirectoryInfo(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\new" & Regression.RegressionMain.klbMismatchedClients.SelectedItem)

        'get the .emf pages
        Dim emf As FileInfo() = di.GetFiles("*.emf")

        'Dim PrintPreviewControl1 As New PrintPreviewControl
        Dim i As Integer = Regression.RegressionMain.klbMismatchedClients.SelectedItem

        s = klbPages.SelectedItem

        'if pages are 1-9, remove leading 0 so they will sort correctly
        If s.Substring(4, 1) = "0" Then
            s = s.Remove(4, 1)
        End If


        'add extension back to page #
        s = s & ".emf"

        'reset the picture boxes
        PictureBoxEx1.Visible = False
        PictureBoxEx2.Visible = False
        PictureBoxEx1.Image = Nothing
        PictureBoxEx2.Image = Nothing


        'Fill the picture box with the illustration page if it exists, else put up message
        If FileIO.FileSystem.FileExists(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\new" & i & "\" & s) Then
            PictureBoxEx1.Visible = True
            PictureBoxEx1.Image = Image.FromFile(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\new" & i & "\" & s)
            KryptonLabel4.Visible = False
        Else
            PictureBoxEx1.Visible = True
            KryptonLabel4.Visible = True
            KryptonLabel4.Text = "There is no illustration page"
        End If

        If FileIO.FileSystem.FileExists(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\" & i & "\" & s) Then
            KryptonLabel5.Visible = False
            PictureBoxEx2.Visible = True
            PictureBoxEx2.Image = Image.FromFile(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & "\" & i & "\" & s)
        Else
            PictureBoxEx2.Visible = True
            KryptonLabel5.Visible = True
            KryptonLabel5.Text = "There is no illustration page"

        End If

        'set the labels
        KryptonLabel6.Text = Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue) & " Client # " & Regression.RegressionMain.klbMismatchedClients.SelectedItem & " " & klbPages.SelectedItem

    End Sub
    Private Sub klbpages_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles klbPages.SelectedIndexChanged
        ViewPage()
    End Sub
    Private Sub kbexitview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbExitView.Click

        'on exit, reset the picture boxes
        PictureBoxEx1.Image = Nothing
        PictureBoxEx2.Image = Nothing

        'close the form
        Me.Close()
        Me.Dispose()

        'reenable the controls on the main form
        Dim cCtrl As Control
        For Each cCtrl In Regression.RegressionMain.Controls
            If Regression.RegressionMain.gstrpathProduct = "FIA" Then
                If cCtrl.Name = Regression.RegressionMain.kbPaloma.Name Or cCtrl.Name = Regression.RegressionMain.kbPalomaOnly.Name Then
                    cCtrl.Enabled = False
                Else


                    cCtrl.Enabled = True
                End If
            End If
        Next

        'show the main form
        System.Windows.Forms.Application.DoEvents()
        Regression.RegressionMain.Show()

    End Sub

    Private Sub PictureBoxEx2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxEx2.Click

    End Sub

    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub PictureBoxEx1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBoxEx1.Click

    End Sub

    Private Sub ClickAndDragToZoomDoubleClickToResoreToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class