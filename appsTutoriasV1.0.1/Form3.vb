Imports System.IO
Imports System.Data.OleDb
Imports System.ComponentModel
Imports Microsoft.Office.Core
Imports System.Data
Public Class Form3

    Dim ApExcel = New Microsoft.Office.Interop.Excel.Application
    Dim Libro = ApExcel.Workbooks.Open("C:\basedatos\test1.xlsx")

    Dim Final As Long
    Dim Numero As Long

    Dim semestre As String
    Dim carrera As String
    Dim alumnos As DataTable

    Public Sub Llenar()
        Dim cadena As String = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\basedatos\test1.xlsx';Extended Properties=Excel 8.0;"
        Dim conn As New OleDbConnection(cadena)
        conn.Open()

        Dim da As New OleDbDataAdapter("select * from [Sheet1$]", conn)
        Dim ds As New DataSet
        da.Fill(ds)

        Alumnos = New DataTable
        Alumnos = ds.Tables(0)

        Dim bs As New BindingSource
        bs.DataSource = alumnos

        bs.Filter = "Nombre like '%" & TextBox1.Text & "%'"

        DataGridView1.DataSource = bs

        conn.Close()
    End Sub
    Public Sub CalificacionColumnas()
        Dim total As Double = 0
        Dim total2 As Double = 0
        Dim total3 As Double = 0
        Dim fila As DataGridViewRow = New DataGridViewRow()
        For Each fila In DataGridView1.Rows
            total += Convert.ToDouble(fila.Cells("MAT1").Value)
            total2 += Convert.ToDouble(fila.Cells("Column4").Value)
            total3 += Convert.ToDouble(fila.Cells("Column5").Value)
        Next
        TextBox14.Text = Convert.ToString(total)
        TextBox15.Text = Convert.ToString(total2)
        TextBox16.Text = Convert.ToString(total3)
    End Sub



    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            Llenar()

        Catch ex As Exception
            MsgBox("NO HAY REGISTROS PREVIOS")
        End Try


    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Llenar()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CalificacionColumnas()
    End Sub
End Class