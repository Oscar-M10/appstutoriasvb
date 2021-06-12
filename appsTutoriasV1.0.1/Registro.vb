Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Data
Public Class Registro
    Dim ApExcel = New Microsoft.Office.Interop.Excel.Application
    Dim Libro = ApExcel.Workbooks.add
    Dim Final As Long
    Private Sub Registro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        My.Computer.FileSystem.CreateDirectory("C:\basedatos")
        Libro.SaveAs(Filename:="C:\basedatos\test1.xlsx")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Libro.SaveAs(Filename:="C:\basedatos\test2.xlsx")
        Libro.Sheets(1).Cells(1, 1) = TextBox1.Text
        Libro.Sheets(1).Cells(1, 2) = TextBox2.Text
        Libro.Sheets(1).Cells(1, 3) = TextBox3.Text
        ApExcel.Quit()
        Libro = Nothing
        ApExcel = Nothing
        Me.Close()
    End Sub

    Private Sub Registro_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

    End Sub
End Class