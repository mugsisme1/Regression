Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
'Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook

Imports Excel = Microsoft.Office.Interop.Excel
'Imports Microsoft.Office.Interop.Excel
'Imports System.Runtime.InteropServices
'Imports Microsoft.Office.Interop.Excel.XlFileFormat

Public Class exporttoexcel

    'Private Sub exporttoexcel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    Dim objExcel As New Microsoft.Office.Interop.Excel.Application
    '    Dim objWorkBook As Microsoft.Office.Interop.Excel.Workbook
    '    Dim objWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
    '    Dim misValue As Object = System.Reflection.Missing.Value
    '    Dim i As Integer
    '    Dim j As Integer

    '    objExcel = New Excel.ApplicationClass
    '    objWorkBook = objExcel.Workbooks.Add(misValue)
    '    objWorkSheet = objWorkBook.Sheets("sheet1")

    '    For i = 0 To Regression.RegressionMain.KryptonDataGridView1.RowCount - 2
    '        For j = 0 To Regression.RegressionMain.KryptonDataGridView1.ColumnCount - 1
    '            objWorkSheet.Cells(i + 1, j + 1) = _
    '                Regression.RegressionMain.KryptonDataGridView1(j, i).Value.ToString()
    '        Next
    '    Next

    '    objWorkSheet.SaveAs("C:\vbexcel.xlsx")
    '    objWorkBook.Close()
    '    objExcel.Quit()

    '    'releaseObject(objExcel)
    '    'releaseObject(xlWorkBook)
    '    'releaseObject(xlWorkSheet)

    '    MsgBox("You can find the file C:\vbexcel.xlsx")
    'End Sub



End Class