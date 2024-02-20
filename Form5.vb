Imports System.IO.FileSystemInfo
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Nini.Config
Imports Scripting
Imports System.Drawing.Imaging
Imports System.IO.Path
Imports System.Drawing
Imports System
Public Class Form5
    Public ifolders As Integer
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim CS_NOCLOSE As Integer = Int32.Parse("200", Globalization.NumberStyles.HexNumber)
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = CS_NOCLOSE
            Return cp
        End Get
    End Property
    Private Sub kbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbClose.Click
        Me.Close()
        Me.Dispose()
        'Reenable the controls on the main form
        'Dim cCtrl As Control
        For Each cCtrl In Regression.RegressionMain.Controls
            cCtrl.Enabled = True
        Next
        'If Regressionator was run first, and there were mismatches, re-enable the mismatch list, so you will be asked
        'to delete the test files before exiting....

        If Regression.RegressionMain.klbMismatchedClients.Items.Count > 0 Then
            Regression.RegressionMain.klbMismatchedClients.Enabled = True
        Else
            Regression.RegressionMain.klbMismatchedClients.Enabled = False
        End If

        'show the main form
        System.Windows.Forms.Application.DoEvents()
        Regression.RegressionMain.Show()
    End Sub

    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToParent()
        KryptonLabel6.Text = Regression.RegressionMain.gstrpathProduct & ":  " & Regression.RegressionMain.gstrCase(Regression.RegressionMain.klbCases.SelectedValue)
        'kbDisplayMismatches.Enabled = True
        If Regression.RegressionMain.gstrpathProduct = "SPIA" And Regression.RegressionMain.gbSPIAEffectiveDate = True Then
            kbBench.Enabled = False
        Else
            kbBench.Enabled = True
        End If

    End Sub
    Private Sub Recompare(ByVal icl As Integer)

        'Recompare each case that is benchmarked, individually

        If kdgvPalomaStatus.CurrentRow.Cells(2).Value = "--" Then
            MsgBox("Either the Test or Benchmark, or both do not run, so there is no PDF to compare", MsgBoxStyle.Information)
        Else
            Dim strProject As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.sdp" & """"
            Dim strControl As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.pdf" & """"
            Dim strTest As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Test\Test.pdf" & """"
            Dim strReport As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Report.pdf" & """"

            Dim pRun As New Process
            Dim pView As New Process

            'hide the cli
            pRun.StartInfo.UseShellExecute = False

            pRun.StartInfo.RedirectStandardOutput = True

            pRun.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdconapp.exe "
            pView.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdapp.exe "

            'pass the arguments to the commands
            pRun.StartInfo.Arguments = strProject & " -CD " & strControl & " -TD " & "" & strTest & " -CR -RT PDF -RP " & strReport & " -N -1"
            pView.StartInfo.Arguments = strProject & " " & strControl & " " & strTest

            'use a hidden window
            pRun.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

            pRun.Start()

            pRun.WaitForExit()
            'uncheck the client
            kdgvPalomaStatus.Rows(icl).Cells(0).Value = False
            'Color the text Green
            For ix = 1 To 4
                kdgvPalomaStatus.Rows(icl).Cells(ix).Style.ForeColor = Color.Green
            Next

        End If
    End Sub
    Private Sub Compare()

        'If kdgvPalomaStatus.CurrentRow.Cells(2).Value = "--" Then
        '    MsgBox("Either the Test or Benchmark, or both do not run, so there is no PDF to compare", MsgBoxStyle.Information)
        'Else
        'Dim strProject As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.CurrentRow.Cells(1).Value) & "\Control\Control.sdp" & """"
        'Dim strControl As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.CurrentRow.Cells(1).Value) & "\Control\Control.pdf" & """"
        'Dim strTest As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.CurrentRow.Cells(1).Value) & "\Test\Test.pdf" & """"
        'Dim strReport As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.CurrentRow.Cells(1).Value) & "\Control\Report.pdf" & """"

        'Dim pRun As New Process
        'Dim pView As New Process
        'Dim icl As Integer = 0
        Dim icount As Integer = 0

        'Dim ix As Integer
        'Dim s As String

        '****************************************************
        '****************************************************

        'determine the # of clients selected to run
        Dim icl As Integer = 0
        For icl = 0 To kdgvPalomaStatus.Rows.Count - 1
            If kdgvPalomaStatus.Rows(icl).Cells(0).Value = True Then
                icount = icount + 1
                
                'If kdgvPalomaStatus.Rows(icl).Cells(2).Value = "--" Then
                '    MsgBox("Either the Test or Benchmark, or both do not run, so there is no PDF to compare", MsgBoxStyle.Information)
                'Else
                Dim strProject As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.sdp" & """"
                Dim strControl As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.pdf" & """"
                Dim strTest As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Test\Test.pdf" & """"
                Dim strReport As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Report.pdf" & """"

                Dim pRun As New Process
                Dim pView As New Process

                'hide the cli
                pRun.StartInfo.UseShellExecute = False

                pRun.StartInfo.RedirectStandardOutput = True

                pRun.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdconapp.exe "
                pView.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdapp.exe "

                'pass the arguments to the commands
                pRun.StartInfo.Arguments = strProject & " -CD " & strControl & " -TD " & "" & strTest & " -CR -RT PDF -RP " & strReport & " -N -1"
                pView.StartInfo.Arguments = strProject & " " & strControl & " " & strTest

                'use a hidden window
                pRun.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

                pRun.Start()

                Dim errlev As String
                pRun.WaitForExit()
                errlev = (pRun.ExitCode)
                'If errlev = 0 Then
                '    MsgBox("This client matches the benchmark", MsgBoxStyle.OkOnly)
                'ElseIf errlev = 1 Then
                '    MsgBox("This client does not match the benchmark", MsgBoxStyle.Critical)
                'Else
                '    MsgBox("Some other error...", MsgBoxStyle.Critical)
                'End If

                'Change name in combobox to reflect current status
                If errlev = 1 Then
                    kdgvPalomaStatus.Rows(icl).Cells(2).Value = "Doesn't Match" 's.Replace("-Uncompared", "-Doesn't Match")
                    If InStr(kdgvPalomaStatus.Rows(icl).Cells(3).Value, "*") Then
                    Else
                        kdgvPalomaStatus.Rows(icl).Cells(3).Value = kdgvPalomaStatus.Rows(icl).Cells(3).Value & " *"
                    End If
                    'Uncheck the client
                    kdgvPalomaStatus.Rows(icl).Cells(0).Value = False
                    'Color the text Red
                    For ix = 1 To 4
                        kdgvPalomaStatus.Rows(icl).Cells(ix).Style.ForeColor = Color.Red
                    Next
                ElseIf errlev = 0 Then

                    kdgvPalomaStatus.Rows(icl).Cells(2).Value = "Matches"

                    If InStr(kdgvPalomaStatus.Rows(icl).Cells(3).Value, "*") Then
                    Else
                        kdgvPalomaStatus.Rows(icl).Cells(3).Value = kdgvPalomaStatus.Rows(icl).Cells(3).Value & " *"
                    End If
                    'Uncheck the Client
                    kdgvPalomaStatus.Rows(icl).Cells(0).Value = False
                    'Color the text Green
                    For ix = 1 To 4
                        kdgvPalomaStatus.Rows(icl).Cells(ix).Style.ForeColor = Color.Green
                    Next
                End If

                'If errlev = 0 Then
                'Else
                '    pView.Start()
                '    pView.WaitForExit()
                'End If
            End If

            'End If
        Next
        If icount = 0 Then
            MsgBox("Please choose at least one client to compare", MsgBoxStyle.Exclamation)
        End If


        For i = 0 To kdgvPalomaStatus.Rows.Count - 1
            kdgvPalomaStatus.Rows(i).Cells(0).Value = False
        Next i

        '****************************************************
        '****************************************************


        ''hide the cli
        'pRun.StartInfo.UseShellExecute = False

        'pRun.StartInfo.RedirectStandardOutput = True

        'pRun.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdconapp.exe "
        'pView.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdapp.exe "

        ''pass the arguments to the commands
        'pRun.StartInfo.Arguments = strProject & " -CD " & strControl & " -TD " & "" & strTest & " -CR -RT PDF -RP " & strReport & " -N -1"
        'pView.StartInfo.Arguments = strProject & " " & strControl & " " & strTest

        ''use a hidden window
        'pRun.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

        'pRun.Start()

        'Dim errlev As String
        'pRun.WaitForExit()
        'errlev = (pRun.ExitCode)
        'If errlev = 0 Then
        '    MsgBox("This client matches the benchmark", MsgBoxStyle.OkOnly)
        'ElseIf errlev = 1 Then
        '    MsgBox("This client does not match the benchmark", MsgBoxStyle.Critical)
        'Else
        '    MsgBox("Some other error...", MsgBoxStyle.Critical)

        'End If

        ''MsgBox(errlev)

        ''Change name in combobox to reflect current status
        'If errlev = 1 Then
        '    'ix = kdgvPalomaStatus.CurrentRow.Cells(0).Value
        '    's = kdgvPalomaStatus.CurrentRow.Cells(1).Value
        '    kdgvPalomaStatus.CurrentRow.Cells(2).Value = "Doesn't Match" 's.Replace("-Uncompared", "-Doesn't Match")
        '    If InStr(kdgvPalomaStatus.CurrentRow.Cells(3).Value, "*") Then
        '    Else

        '        kdgvPalomaStatus.CurrentRow.Cells(3).Value = kdgvPalomaStatus.CurrentRow.Cells(3).Value & " *"
        '    End If
        '    'kcbPalomaNewBench.Items.Add(kdgvPalomaStatus.CurrentRow.Cells(0).Value)
        '    'kcbPalomaCompare.Items.Remove(kcbPalomaCompare.Items(ix))
        '    'If kcbPalomaCompare.Items.Count > 0 Then
        '    '    kcbPalomaCompare.SelectedItem = kcbPalomaCompare.Items(0)
        '    'End If
        '    'If kcbPalomaCompare.Items.Count = 0 Then
        '    '    kcbPalomaCompare.ResetText()
        '    '    kcbPalomaCompare.Update()
        '    '    kcbPalomaCompare.Items.Add("No More Clients to Compare")
        '    '    kcbPalomaCompare.SelectedItem = "No More Clients to Compare"
        '    '    kbCompare.Enabled = False
        '    'End If


        'ElseIf errlev = 0 Then
        '    'ix = kcbPalomaCompare.SelectedIndex
        '    's = kcbPalomaCompare.Items(ix)
        '    'kcbPalomaCompare.Items(ix) = s.Replace("-Uncompared", "Matches")
        '    kdgvPalomaStatus.CurrentRow.Cells(2).Value = "Matches"
        '    'kdgvPalomaStatus.Columns(0).HeaderText = "*"
        '    If InStr(kdgvPalomaStatus.CurrentRow.Cells(3).Value, "*") Then
        '    Else
        '        kdgvPalomaStatus.CurrentRow.Cells(3).Value = kdgvPalomaStatus.CurrentRow.Cells(3).Value & " *"
        '    End If

        '    'kcbPalomaMatches.Items.Add(kcbPalomaCompare.Items(ix))
        '    'kcbPalomaCompare.Items.Remove(kcbPalomaCompare.Items(ix))
        '    'If kcbPalomaCompare.Items.Count > 0 Then
        '    '    kcbPalomaCompare.SelectedItem = kcbPalomaCompare.Items(0)
        '    'kcbPalomaMatches.SelectedItem = kcbPalomaMatches.Items(0)
        '    'End If
        '    'If kcbPalomaCompare.Items.Count = 0 Then
        '    '    kcbPalomaCompare.ResetText()
        '    '    kcbPalomaCompare.Update()
        '    '    kcbPalomaCompare.Items.Add("No More Clients to Compare")
        '    '    kcbPalomaCompare.SelectedItem = "No More Clients to Compare"
        '    '    'kcbPalomaCompare.SelectionMode = Toolkit.CheckedSelectionMode.None
        '    '    'kcbPalomaCompare.SelectedItem
        '    '    kbCompare.Enabled = False

        '    'End If

        'End If

        'If errlev = 0 Then
        'Else
        '    pView.Start()
        '    pView.WaitForExit()
        'End If
        'End If
    End Sub

    Private Sub kbCompare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbCompare.Click

        Compare()
        'Dim strProject As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.CurrentRow.Cells(0).Value) & "\Control\Control.sdp" & """"
        'Dim strControl As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.CurrentRow.Cells(0).Value) & "\Control\Control.pdf" & """"
        'Dim strTest As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.CurrentRow.Cells(0).Value) & "\Test\Test.pdf" & """"
        'Dim strReport As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.CurrentRow.Cells(0).Value) & "\Control\Report.pdf" & """"

        'Dim pRun As New Process
        'Dim pView As New Process

        ''Dim ix As Integer
        ''Dim s As String

        ''hide the cli
        'pRun.StartInfo.UseShellExecute = False

        'pRun.StartInfo.RedirectStandardOutput = True

        'pRun.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdconapp.exe "
        'pView.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdapp.exe "

        ''pass the arguments to the commands
        'pRun.StartInfo.Arguments = strProject & " -CD " & strControl & " -TD " & "" & strTest & " -CR -RT PDF -RP " & strReport & " -N -1"
        'pView.StartInfo.Arguments = strProject & " " & strControl & " " & strTest

        ''use a hidden window
        'pRun.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

        'pRun.Start()

        'Dim errlev As String
        'pRun.WaitForExit()
        'errlev = (pRun.ExitCode)
        'If errlev = 0 Then
        '    MsgBox("This client matches the benchmark", MsgBoxStyle.OkOnly)
        'ElseIf errlev = 1 Then
        '    MsgBox("This client does not match the benchmark", MsgBoxStyle.Exclamation)
        'Else
        '    MsgBox("Some other error...", MsgBoxStyle.Critical)

        'End If
        ''MsgBox(errlev)

        ''Change name in combobox to reflect current status
        'If errlev = 1 Then
        '    'ix = kdgvPalomaStatus.CurrentRow.Cells(0).Value
        '    's = kdgvPalomaStatus.CurrentRow.Cells(1).Value
        '    kdgvPalomaStatus.CurrentRow.Cells(1).Value = "Doesn't Match" 's.Replace("-Uncompared", "-Doesn't Match")
        '    'kcbPalomaNewBench.Items.Add(kdgvPalomaStatus.CurrentRow.Cells(0).Value)
        '    'kcbPalomaCompare.Items.Remove(kcbPalomaCompare.Items(ix))
        '    'If kcbPalomaCompare.Items.Count > 0 Then
        '    '    kcbPalomaCompare.SelectedItem = kcbPalomaCompare.Items(0)
        '    'End If
        '    'If kcbPalomaCompare.Items.Count = 0 Then
        '    '    kcbPalomaCompare.ResetText()
        '    '    kcbPalomaCompare.Update()
        '    '    kcbPalomaCompare.Items.Add("No More Clients to Compare")
        '    '    kcbPalomaCompare.SelectedItem = "No More Clients to Compare"
        '    '    kbCompare.Enabled = False
        '    'End If


        'ElseIf errlev = 0 Then
        '    'ix = kcbPalomaCompare.SelectedIndex
        '    's = kcbPalomaCompare.Items(ix)
        '    'kcbPalomaCompare.Items(ix) = s.Replace("-Uncompared", "Matches")
        '    kdgvPalomaStatus.CurrentRow.Cells(1).Value = "Matches"
        '    'kcbPalomaMatches.Items.Add(kcbPalomaCompare.Items(ix))
        '    'kcbPalomaCompare.Items.Remove(kcbPalomaCompare.Items(ix))
        '    'If kcbPalomaCompare.Items.Count > 0 Then
        '    '    kcbPalomaCompare.SelectedItem = kcbPalomaCompare.Items(0)
        '    'kcbPalomaMatches.SelectedItem = kcbPalomaMatches.Items(0)
        '    'End If
        '    'If kcbPalomaCompare.Items.Count = 0 Then
        '    '    kcbPalomaCompare.ResetText()
        '    '    kcbPalomaCompare.Update()
        '    '    kcbPalomaCompare.Items.Add("No More Clients to Compare")
        '    '    kcbPalomaCompare.SelectedItem = "No More Clients to Compare"
        '    '    'kcbPalomaCompare.SelectionMode = Toolkit.CheckedSelectionMode.None
        '    '    'kcbPalomaCompare.SelectedItem
        '    '    kbCompare.Enabled = False

        '    'End If

        'End If

        'If errlev = 0 Then
        'Else
        '    pView.Start()
        '    pView.WaitForExit()
        'End If

    End Sub

    Private Sub kbBench_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbBench.Click

        Dim icount As Integer = 0
        Dim icl As Integer = 0
        'Benchmark each client selected
        For icl = 0 To kdgvPalomaStatus.Rows.Count - 1
            If kdgvPalomaStatus.Rows(icl).Cells(0).Value = True Then
                icount = icount + 1

                'If kdgvPalomaStatus.Rows(icl).Cells(2).Value = "--" Then
                '    MsgBox("For at least one client, Either the Test or Benchmark, or both do not run, so there is no PDF to compare", MsgBoxStyle.Information)
                '    Exit For
                If kdgvPalomaStatus.Rows(icl).Cells(2).Value = "Uncompared" Then
                    MsgBox("Each client must be compared before making a new benchmark")
                    Exit For
                    'ElseIf kdgvPalomaStatus.Rows(icl).Cells(2).Value = "Matches" Then
                    '    MsgBox("At least one client already matches")
                    '    Exit For
                Else
                    'in order to create benchmarks, have to manually copy all of the test files to the control folder
                    Dim strBench As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\bench.txt" & ""
                    Dim strBenchTest As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Test\Test.pdf" & ""
                    Dim strBenchControl As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.pdf" & ""


                    Dim strpath As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\"
                    Dim strcase As String = Regression.RegressionMain.gstrCase(0) & "\"
                    Dim iclient As Integer = (kdgvPalomaStatus.Rows(icl).Cells(1).Value)


                    'Rename the old
                    System.IO.File.Copy(strBenchControl, "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\LastControl.pdf", overwrite:=True)

                    'Copy test to new bench
                    System.IO.File.Copy(strBenchTest, strBenchControl, overwrite:=True)

                    '"Touch" the bench.txt file to set the time stamp
                    System.IO.File.SetLastWriteTime(strBench, Date.Now)

                    'Rerun compare to confirm benchmark and to update log and report files
                    Recompare(icl)

                    kdgvPalomaStatus.Rows(icl).Cells(2).Value = "New Benchmark Made "
                    If InStr(kdgvPalomaStatus.Rows(icl).Cells(4).Value, "*") Then
                    Else
                        kdgvPalomaStatus.Rows(icl).Cells(4).Value = kdgvPalomaStatus.Rows(icl).Cells(4).Value & " *"
                    End If

                End If
            End If
        Next icl
        If icount = 0 Then
            MsgBox("Please choose at least one client to benchmark", MsgBoxStyle.Exclamation)
        End If
        'Regression.RegressionMain.WriteMatchStatus(Regression.RegressionMain.gstrpath & "\" & Regression.RegressionMain.gstrpathProduct & "\", Regression.RegressionMain.gstrCase(0), oChkedItems, Regression.RegressionMain.gstrCompCode, clsReadVAValues.bMisMatch, bNewBench, value)

        '"\Control\Control.pdf" & ""
        '        Regression.RegressionMain.WriteMatchStatus()


        '    Public Sub WriteMatchStatus(ByVal strpath As String, ByVal strcase As String, ByVal iclient As Integer, ByVal strcomp() As String, ByVal bMisMatch As Boolean, ByVal bNewBench As Boolean, Optional ByVal note As String = "")


    End Sub
    Private Sub KryptonTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub kdgvPalomaStatus_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles kdgvPalomaStatus.CellContentClick
        'If kdgvPalomaStatus.CurrentRow.Cells(1).Value = "Uncompared" Then
        '    kbView.Enabled = False
        'Else
        '    kbView.Enabled = True
        'End If
    End Sub

    Private Sub kbView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbView.Click


        Dim icl As Integer = 0
        Dim icount As Integer = 0

        'Check to see how many clients selected, but only allow one to be viewed at a time
        For icl = 0 To kdgvPalomaStatus.Rows.Count - 1
            If kdgvPalomaStatus.Rows(icl).Cells(0).Value = True Then
                icount = icount + 1
            End If
        Next icl
        If icount = 0 Then
            MsgBox("Please choose one client to view", MsgBoxStyle.Exclamation)
        ElseIf icount > 1 Then
            MsgBox("Please choose only one client to view at a time", MsgBoxStyle.Exclamation)

        Else
            For icl = 0 To kdgvPalomaStatus.Rows.Count - 1
                If kdgvPalomaStatus.Rows(icl).Cells(0).Value = True Then
                    Dim strProject As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.sdp" & """"
                    Dim strControl As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.pdf" & """"
                    Dim strTest As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Test\Test.pdf" & """"
                    Dim strReport As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Report.pdf" & """"

                    If kdgvPalomaStatus.Rows(icl).Cells(2).Value = "--" Then
                        MsgBox("Either the Test or Benchmark, or both do not run, so there is no PDF to compare", MsgBoxStyle.Information)
                    ElseIf kdgvPalomaStatus.Rows(icl).Cells(2).Value = "Uncompared" Then
                        MsgBox("must be compared before viewing")
                    Else
                        Dim pView As New Process

                        'hide the cli

                        pView.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdapp.exe "

                        'pass the arguments to the commands

                        pView.StartInfo.Arguments = strProject & " " & strControl & " " & strTest

                        pView.Start()
                        kbView.Enabled = False
                        pView.WaitForExit()
                        kbView.Enabled = True

                    End If
                End If
            Next icl
        End If
        'Dim strProject As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.sdp" & """"
        'Dim strControl As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.pdf" & """"
        'Dim strTest As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Test\Test.pdf" & """"
        'Dim strReport As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Report.pdf" & """"

        'If kdgvPalomaStatus.CurrentRow.Cells(2).Value = "--" Then
        '    MsgBox("Either the Test or Benchmark, or both do not run, so there is no PDF to compare", MsgBoxStyle.Information)
        'ElseIf kdgvPalomaStatus.CurrentRow.Cells(2).Value = "Uncompared" Then
        '    MsgBox("must be compared before viewing")
        'Else
        '    Dim pView As New Process

        ''hide the cli

        'pView.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdapp.exe "

        ''pass the arguments to the commands

        'pView.StartInfo.Arguments = strProject & " " & strControl & " " & strTest

        'pView.Start()
        'pView.WaitForExit()

        '    End If
        'End If
        '****************************************************
        '****************************************************
        'Dim icount As Integer = 0
        ''determine the # of clients selected to run, only allow one
        'Dim icl As Integer = 0
        'For icl = 0 To kdgvPalomaStatus.Rows.Count - 1
        '    If kdgvPalomaStatus.Rows(icl).Cells(0).Value = True Then
        '        icount = icount + 1

        '        If kdgvPalomaStatus.Rows(icl).Cells(2).Value = "--" Then
        '            MsgBox("Either the Test or Benchmark, or both do not run, so there is no PDF to compare", MsgBoxStyle.Information)
        '        Else
        '            Dim strProject As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.sdp" & """"
        '            Dim strControl As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Control.pdf" & """"
        '            Dim strTest As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Test\Test.pdf" & """"
        '            Dim strReport As String = """\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (kdgvPalomaStatus.Rows(icl).Cells(1).Value) & "\Control\Report.pdf" & """"

        '            Dim pRun As New Process
        '            Dim pView As New Process

        '            'hide the cli
        '            pRun.StartInfo.UseShellExecute = False

        '            pRun.StartInfo.RedirectStandardOutput = True

        '            pRun.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdconapp.exe "
        '            pView.StartInfo.FileName = "C:\Program Files\STREAMdiff\Bin\sdapp.exe "

        '            'pass the arguments to the commands
        '            pRun.StartInfo.Arguments = strProject & " -CD " & strControl & " -TD " & "" & strTest & " -CR -RT PDF -RP " & strReport & " -N -1"
        '            pView.StartInfo.Arguments = strProject & " " & strControl & " " & strTest

        '            'use a hidden window
        '            pRun.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

        '            pRun.Start()

        '            Dim errlev As String
        '            pRun.WaitForExit()
        '            errlev = (pRun.ExitCode)
        '            'If errlev = 0 Then
        '            '    MsgBox("This client matches the benchmark", MsgBoxStyle.OkOnly)
        '            'ElseIf errlev = 1 Then
        '            '    MsgBox("This client does not match the benchmark", MsgBoxStyle.Critical)
        '            'Else
        '            '    MsgBox("Some other error...", MsgBoxStyle.Critical)
        '            'End If

        '            'Change name in combobox to reflect current status
        '            If errlev = 1 Then
        '                kdgvPalomaStatus.Rows(icl).Cells(2).Value = "Doesn't Match" 's.Replace("-Uncompared", "-Doesn't Match")
        '                If InStr(kdgvPalomaStatus.Rows(icl).Cells(3).Value, "*") Then
        '                Else
        '                    kdgvPalomaStatus.Rows(icl).Cells(3).Value = kdgvPalomaStatus.Rows(icl).Cells(3).Value & " *"
        '                End If
        '                kdgvPalomaStatus.Rows(icl).Cells(0).Value = False
        '            ElseIf errlev = 0 Then

        '                kdgvPalomaStatus.Rows(icl).Cells(2).Value = "Matches"

        '                If InStr(kdgvPalomaStatus.Rows(icl).Cells(3).Value, "*") Then
        '                Else
        '                    kdgvPalomaStatus.Rows(icl).Cells(3).Value = kdgvPalomaStatus.Rows(icl).Cells(3).Value & " *"
        '                End If
        '                kdgvPalomaStatus.Rows(icl).Cells(0).Value = False
        '            End If

        '            'If errlev = 0 Then
        '            'Else
        '            '    pView.Start()
        '            '    pView.WaitForExit()
        '            'End If
        '        End If

        '    End If
        'Next
        'If icount = 0 Then
        '    MsgBox("Please choose at least one client to compare", MsgBoxStyle.Exclamation)
        'End If
        '****************************************************
        '****************************************************



    End Sub

    Private Sub KryptonBorderEdge3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles KryptonBorderEdge3.Paint

    End Sub

    Private Sub KryptonBorderEdge2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles KryptonBorderEdge2.Paint

    End Sub

    Private Sub KryptonBorderEdge1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles KryptonBorderEdge1.Paint

    End Sub

    Private Sub kbAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbAll.Click
        'Select all clients listed
        Dim i As Integer
        For i = 0 To kdgvPalomaStatus.Rows.Count - 1
            kdgvPalomaStatus.Rows(i).Cells(0).Value = True
        Next i
    End Sub

    Private Sub kbNone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbNone.Click
        'De-Select all clients listed
        Dim i As Integer
        For i = 0 To kdgvPalomaStatus.Rows.Count - 1
            kdgvPalomaStatus.Rows(i).Cells(0).Value = False
        Next i
    End Sub
    
    Sub KbCompareClick(sender As Object, e As EventArgs)
    	
    End Sub
End Class