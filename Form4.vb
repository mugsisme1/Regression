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
Public Class Form4

    Private Sub KryptonLabel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub KryptonButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Me.Dispose()
        Me.Close()
        Me.Dispose()


    End Sub
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim CS_NOCLOSE As Integer = Int32.Parse("200", Globalization.NumberStyles.HexNumber)
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = CS_NOCLOSE
            Return cp
        End Get
    End Property
    Private Sub KryptonButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Dispose()

    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToParent()

        KryptonLabel6.Text = Regression.RegressionMain.gstrpathProduct & ":  " & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub KryptonButton1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KryptonButton1.Click
        'Select all
        For ix = 0 To kclbNewBench.Items.Count - 1
            kclbNewBench.SetItemChecked(ix, True)
        Next
    End Sub

    Private Sub KryptonButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KryptonButton2.Click
        'Clear all
        For ix = 0 To kclbNewBench.Items.Count - 1
            kclbNewBench.SetItemChecked(ix, False)
        Next
    End Sub

    Private Sub KryptonButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KryptonButton3.Click

        'Create new Benchmarks 

        Dim bNewBench As Boolean = True
        Dim ix As Integer = 0
        Dim iy As Integer = 0
        Dim iCount As Integer = 0

        If kclbNewBench.CheckedItems.Count = 0 Then
            MsgBox("Please select at least one case to create a new Benchmark from.", MsgBoxStyle.OkOnly)
            Return
        End If

        MsgBox("Clicking OK will create new Benchmarks for the selected cases.  The old values and illustration pages will be lost!", MsgBoxStyle.Exclamation)


        Dim message As String
        Dim title As String
        Dim value As String

        message = "Add notes for these new Benchmarks"
        title = "The Regressionator:  Benchmark Notes"
        value = InputBox(message, title)



        Dim oChkedItems As Object

        'Determine the items that are checked, by name, so can match up to correct folders
        For Each oChkedItems In kclbNewBench.CheckedItems

            Dim diOldCaseDir As DirectoryInfo = New DirectoryInfo(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & oChkedItems)
            Dim diNewCaseDir As DirectoryInfo = New DirectoryInfo(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\New" & oChkedItems & "\")
            Dim finext As FileInfo

            'list of .emf files
            Dim emfnew As FileInfo() = diNewCaseDir.GetFiles("*.emf")

            'relay.ini file
            Dim ininew As FileInfo() = diNewCaseDir.GetFiles("*.ini")

            'relay.out file
            Dim outnew As FileInfo() = diNewCaseDir.GetFiles("*.out")

            'copy the new files into the existing folders
            For Each finext In emfnew
                System.IO.File.Copy(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\New" & oChkedItems & "\" & finext.Name, Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & oChkedItems & "\" & finext.Name, True)
            Next
            For Each finext In ininew
                If finext.Name = "Gnawin.ini" Then
                Else
                    System.IO.File.Copy(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\New" & oChkedItems & "\" & finext.Name, Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & oChkedItems & "\" & finext.Name, True)
                End If
            Next
            For Each finext In outnew
                System.IO.File.Copy(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\New" & oChkedItems & "\" & finext.Name, Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & oChkedItems & "\" & finext.Name, True)
            Next


            ' write to the existing stats file, noting that a new benchmark was created
            Regression.RegressionMain.WriteMatchStatus(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\", Regression.RegressionMain.gstrCase(0), oChkedItems, Regression.RegressionMain.gstrCompCode, clsReadVAValues.bMisMatch, bNewBench, value)

            Regression.RegressionMain.WriteBenchMarkDate(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\", Regression.RegressionMain.gstrCase(0), oChkedItems)



            'delete new folder(s)
            Dim FSO As New FileSystemObject
            Directory.Delete(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\New" & oChkedItems, True)

            System.Windows.Forms.Application.DoEvents()

        Next oChkedItems

        'Delete items from checkbox

        For ix = kclbNewBench.Items.Count - 1 To 0 Step -1
            If kclbNewBench.GetItemChecked(ix) Then
                kclbNewBench.Items.RemoveAt(ix)
            End If
        Next ix

    End Sub

    Private Sub kbExitView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbExitView.Click

        'close this form
        Me.Close()
        Me.Dispose()
        'Reenable the controls on the main form
        Dim cCtrl As Control
        For Each cCtrl In Regression.RegressionMain.Controls
            If cCtrl.TabIndex = 9 Then
                cCtrl.Enabled = False
            Else
                cCtrl.Enabled = True
            End If
        Next

        'Clear the list box of mismatched cases on the main form
        Regression.RegressionMain.klbMismatchedClients.Items.Clear()

        'Populate the mistmatched cases list box on the main form with whatever cases are left
        For ix = 0 To kclbNewBench.Items.Count - 1
            Regression.RegressionMain.klbMismatchedClients.Items.Add(kclbNewBench.Items(ix))
        Next

        'If no cases are left, disable the unneeded controls on the main form
        If Regression.RegressionMain.klbMismatchedClients.Items.Count = 0 Then
            Regression.RegressionMain.klbMismatchedClients.Enabled = False
            'Regression.RegressionMain.kbDisplayMismatches.Enabled = False
            Regression.RegressionMain.kbViewIllustration.Enabled = False
            Regression.RegressionMain.kbNewBench.Enabled = False
            Regression.RegressionMain.kbDeleteFiles.Enabled = False
            Regression.RegressionMain.KryptonDataGridView1.RowCount = 0
            Regression.RegressionMain.kbSaveMismatchResults.Enabled = False
            Regression.RegressionMain.kbSaveMismatchResults.Enabled = False
        Else
            Regression.RegressionMain.KryptonDataGridView1.RowCount = 0
            'Regression.RegressionMain.klbMismatchedClients.SelectedIndex = 0
            Regression.RegressionMain.kbSaveMismatchResults.Enabled = False
            Regression.RegressionMain.kbViewIllustration.Visible = False
            Regression.RegressionMain.kbViewIllustration.Enabled = False
        End If
        For Each cCtrl In Regression.RegressionMain.Controls
            If Regression.RegressionMain.gstrpathProduct = "FIA" Then
                If cCtrl.Name = Regression.RegressionMain.kbPaloma.Name Or cCtrl.Name = Regression.RegressionMain.kbPalomaOnly.Name Then
                    cCtrl.Enabled = False
                End If
                'Else


                '    cCtrl.Enabled = True
                'End If
            End If
        Next
        'show the main form
        System.Windows.Forms.Application.DoEvents()
        Regression.RegressionMain.Show()

    End Sub

    Private Sub kclbNewBench_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kclbNewBench.SelectedIndexChanged

    End Sub
End Class