
Imports System.IO
Imports System.Data.OleDb
Imports System.ComponentModel
Imports Microsoft.Office.Core
Imports System.Data
Public Class Alumnos
    Dim ApExcel = New Microsoft.Office.Interop.Excel.Application
    Dim Libro = ApExcel.Workbooks.Open("C:\basedatos\test1.xlsx")
    Dim Final As Long
    Dim Numero As Long
    Dim semestre As String
    Dim carrera As String
    Dim alumnos As DataTable

    Private Sub verificaCamposSalida(sender As Object, e As EventArgs) Handles TextBox2.LostFocus, TextBox5.LostFocus, TextBox6.LostFocus, TextBox7.LostFocus, TextBox8.LostFocus
        Dim txt As New TextBox
        txt = sender
        If txt.Text = "" Then
            txt.Text = "<-complete los datos->"
        End If
        sender = txt
    End Sub

    Private Sub verificaCamposEntrada(sender As Object, e As EventArgs) Handles TextBox2.GotFocus, TextBox5.GotFocus, TextBox6.GotFocus, TextBox7.GotFocus, TextBox8.GotFocus
        Dim txt As New TextBox
        txt = sender
        If txt.Text = "<-complete los datos->" Then
            txt.Text = ""
        End If
        sender = txt
    End Sub

    Public Sub Llenar()
        Dim cadena As String = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\basedatos\test1.xlsx';Extended Properties=Excel 8.0;"
        Dim conn As New OleDbConnection(cadena)
        conn.Open()

        Dim da As New OleDbDataAdapter("select * from [Hoja1$]", conn)
        Dim ds As New DataSet
        da.Fill(ds)

        alumnos = New DataTable
        alumnos = ds.Tables(0)

        Dim bs As New BindingSource
        bs.DataSource = alumnos

        bs.Filter = "Nombres like '%" & TextBox13.Text & "%'"

        DataGridView1.DataSource = bs
        conn.Close()
    End Sub

    Private Sub Alumnos_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub
    Private Sub LlenaTexto()
        ' Label3.Text = DataGridView1.Item(0, 2).Value.ToString
        'Label3.Text = DataGridView1.Item(0, 3).Value.ToString
        'Label3.Text = DataGridView1.Item(0, 4).Value.ToString
        ' Label3.Text = DataGridView1.Item(0, 5).Value.ToString
    End Sub
    Private Sub LlenaTexto2()
        Try
            TextBox1.Text = DataGridView1.CurrentRow.Cells("ID").Value.ToString
            TextBox2.Text = DataGridView1.CurrentRow.Cells("NOMBRE").Value.ToString
            '   TextBox3.Text = DataGridView1.CurrentRow.Cells("SEMESTRE").Value.ToString
            '   TextBox4.Text = DataGridView1.CurrentRow.Cells("CARRERA").Value.ToString
            TextBox5.Text = DataGridView1.CurrentRow.Cells("MATERIA1").Value.ToString
            TextBox6.Text = DataGridView1.CurrentRow.Cells("MATERIA2").Value.ToString
            TextBox7.Text = DataGridView1.CurrentRow.Cells("MATERIA3").Value.ToString
            TextBox8.Text = DataGridView1.CurrentRow.Cells("MATERIA4").Value.ToString
            '  TextBox9.Text = DataGridView1.CurrentRow.Cells("MATERIA5").Value.ToString
            ' TextBox10.Text = DataGridView1.CurrentRow.Cells("MATERIA6").Value.ToString
            '  TextBox11.Text = DataGridView1.CurrentRow.Cells("MATERIA7").Value.ToString
            '  TextBox12.Text = DataGridView1.CurrentRow.Cells("MATERIA8").Value.ToString
            TextBox13.Text = DataGridView1.CurrentRow.Cells("MATERIA9").Value.ToString
            ActivarText()

        Catch ex As Exception
            MsgBox("No se puede mostrar los datos completos por falta de información")
        End Try

    End Sub
    Public Sub DesactivarbotonesMateria()
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        ' TextBox9.Enabled = False
        ' TextBox10.Enabled = False
        ' TextBox11.Enabled = False
        ' TextBox12.Enabled = False
        Label5.Enabled = False
        ' Button1.Enabled = False

    End Sub
    Public Sub ActivarbotonesMateria()
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        'TextBox9.Enabled = True
        ' TextBox10.Enabled = True
        ' TextBox11.Enabled = True
        ' TextBox12.Enabled = True
        Label5.Enabled = True
        ' Button1.Enabled = True
    End Sub
    Public Sub DesactivarbotonesNombre()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        '' TextBox3.Enabled = False
        '' TextBox4.Enabled = False
        Label1.Enabled = False
        Label2.Enabled = False
        '' Label3.Enabled = False
        '' Label4.Enabled = False
        Button2.Visible = False
        ComboBox1.Visible = False
        ComboBox2.Visible = False
        ComboBox3.Visible = False
    End Sub
    Public Sub ActivarbotonesNombre()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        '' TextBox3.Enabled = True
        '' TextBox4.Enabled = True
        Label1.Enabled = True
        Label2.Enabled = True
        '' Label3.Enabled = True
        '' Label4.Enabled = True
        Button2.Visible = True
    End Sub


    Public Sub ActivarText()
        TextBox1.Visible = False
        TextBox2.Visible = False
        ''TextBox3.Visible = False
        '' TextBox4.Visible = False
        TextBox5.Visible = False
        TextBox6.Visible = False
        ComboBox1.Visible = False
        ComboBox2.Visible = False
        ComboBox3.Visible = False
        'Button1.Visible = False
        Button2.Visible = False
    End Sub

    Public Sub LimpiarCampos()
        TextBox1.Clear()
        TextBox2.Clear()
        '' TextBox3.Clear()
        '' TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        ' TextBox9.Clear()
        ' TextBox10.Clear()
        ' TextBox11.Clear()
        ' TextBox12.Clear()
        TextBox5.Focus()
        TextBox1.Focus()
    End Sub
    Public Sub LimpiarCamposDos()
        TextBox1.Clear()
        TextBox2.Clear()
        '' TextBox3.Clear()
        ''TextBox4.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub EliminarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EliminarToolStripMenuItem.Click
        DataGridView1.Rows.Add(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text)
        'Final = nReg(Libro.Worksheets("Sheet1"), 1, 2)
        ' Libro.Worksheets("Sheet1").Cells(Final, 1) = TextBox2.Text
        ' MsgBox("Registro agregado a la fila " & Final)
    End Sub

    Private Sub ActualizarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActualizarToolStripMenuItem.Click
        ' DataGridView1.Rows.Add("1", "Nombre", TextBox5.Text, TextBox6.Text, TextBox7.Text)
        DataGridView1.Columns.Add("Column1", "ID")
        DataGridView1.Columns.Add("Column2", "ApellidoPaterno")
        DataGridView1.Columns.Add("Column3", "ApellidoMaterno")
        DataGridView1.Columns.Add("Column4", "Nombres")
        DataGridView1.Columns.Add("Column5", TextBox5.Text)
        DataGridView1.Columns.Add("Column6", TextBox6.Text)
        DataGridView1.Columns.Add("Column7", TextBox7.Text)
        DataGridView1.Columns.Add("Column8", TextBox8.Text)
        DataGridView1.Columns.Add("Column9", TextBox9.Text)
        DataGridView1.Columns.Add("Column10", TextBox10.Text)
        DataGridView1.Columns.Add("Column11", "Total")


    End Sub

    Private Sub GuardarToolStripMenuItem_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub NuevoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NuevoToolStripMenuItem.Click
        ActivarbotonesMateria()
        LimpiarCampos()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Items.Clear()

        carrera = ComboBox1.SelectedItem
        If carrera = "Tics" Then
            ComboBox2.Items.Add("Primer Semestre tics")
            ComboBox2.Items.Add("Segundo Semestre tics")
            ComboBox2.Items.Add("Tercer Semestre tics")
        ElseIf carrera = "Administración" Then
            ComboBox2.Items.Add("Primer Semestre administracion")
            ComboBox2.Items.Add("Segundo Semestre administracion")
            ComboBox2.Items.Add("Tercer Semestre administracion")
        End If


        ' TextBox1.Text,TextBox2.Text
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        semestre = ComboBox2.Text

        Select Case semestre
            Case "Primer Semestre tics", "Segundo Semestre tics", "Tercer Semestre tics"
                Select Case semestre
                    Case "Primer Semestre tics"
                        ComboBox3.Items.Add("Materia1")
                        ComboBox3.Items.Add("Materia2")
                        ComboBox3.Items.Add("Materia3 ")
                    Case "Segundo Semestre tics"
                        ComboBox3.Items.Add("Materia4")
                        ComboBox3.Items.Add("Materia5")
                        ComboBox3.Items.Add("Materia6 ")
                    Case "Tercer Semestre tics"
                        ComboBox3.Items.Add("Materia7")
                        ComboBox3.Items.Add("Materia8")
                        ComboBox3.Items.Add("Materia9")
                End Select
            Case "Primer Semestre administracion", "Segundo Semestre administracion", "Tercer Semestre administracion"
                Select Case semestre
                    Case "Primer Semestre administracion"
                        ComboBox3.Items.Add("Materia10")
                        ComboBox3.Items.Add("Materia11")
                        ComboBox3.Items.Add("Materia12")
                    Case "Segundo Semestre administracion"
                        ComboBox3.Items.Add("Materia13")
                        ComboBox3.Items.Add("Materia14")
                        ComboBox3.Items.Add("Materia15 ")
                    Case "Tercer Semestre administracion"
                        ComboBox3.Items.Add("Materia16")
                        ComboBox3.Items.Add("Materia17")
                        ComboBox3.Items.Add("Materia18")
                End Select
        End Select

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub Button3_Click_2(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim Num1, Num2, Num3, Num4, Num5, Num6, Total As String
            Num1 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(4).FormattedValue
            Num2 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(5).FormattedValue
            Num3 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(6).FormattedValue
            Num4 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(7).FormattedValue
            Num5 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(8).FormattedValue
            Num6 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(9).FormattedValue

            Total = (Val(Num1) + Val(Num2) + Val(Num3) / 3)
            'TextBox1.Text = Total
            DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(10).Value = Total
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub EDITARREGISTROSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EDITARREGISTROSToolStripMenuItem.Click
        Form3.Show()
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click

        Dim SAVE As New SaveFileDialog
        Dim ruta As String
        Dim xlApp As Object = CreateObject("Excel.Application")
        Dim pth As String = ""
        'crearemos una nueva hoja de calculo
        Dim xlwb As Object = xlApp.WorkBooks.add
        Dim xlws As Object = xlwb.WorkSheets(1)
        Try
            'exportaremos los caracteres de las columnas
            For c As Integer = 0 To DataGridView1.Columns.Count - 1
                xlws.cells(1, c + 1).value = DataGridView1.Columns(c).HeaderText
            Next
            'exportaremos las cabeceras de las calumnas
            For r As Integer = 0 To DataGridView1.RowCount - 1
                For c As Integer = 0 To DataGridView1.Columns.Count - 1
                    xlws.cells(r + 2, c + 1).value = Convert.ToString(DataGridView1.Item(c, r).Value)
                Next
            Next
            'guardamos la hoja de excel en la ruta especifica
            Dim SaveFileDialog1 As SaveFileDialog = New SaveFileDialog
            SaveFileDialog1.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            SaveFileDialog1.Filter = "Archivo Excel| *.xlsx"
            SaveFileDialog1.FilterIndex = 2
            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                ruta = SaveFileDialog1.FileName
                xlwb.saveas(ruta)
                xlws = Nothing
                xlwb = Nothing
                xlApp.quit()
                MsgBox("Exportado Correctamente", MsgBoxStyle.Information)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Try
            If e.ColumnIndex = 10 Then
                If Val(e.Value) <= 70 Then
                    e.CellStyle.ForeColor = Color.Black
                    e.CellStyle.BackColor = Color.Yellow
                    If Val(e.Value < 69) Then
                        e.CellStyle.ForeColor = Color.Black
                        e.CellStyle.BackColor = Color.Red
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox13_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox13.KeyPress

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        LlenaTexto()
    End Sub

    Private Sub GenerarNumerosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerarNumerosToolStripMenuItem.Click
        Dim total As Double = 0
        Dim total2 As Double = 0
        Dim total3 As Double = 0
        Dim total4 As Double = 0
        Dim total5 As Double = 0
        Dim total6 As Double = 0

        Dim fila As DataGridViewRow = New DataGridViewRow()


        For Each fila In DataGridView1.Rows
            total += Convert.ToDouble(fila.Cells("Column5").Value)
            total2 += Convert.ToDouble(fila.Cells("Column6").Value)
            total3 += Convert.ToDouble(fila.Cells("Column7").Value)
            total4 += Convert.ToDouble(fila.Cells("Column8").Value)
            total5 += Convert.ToDouble(fila.Cells("Column9").Value)
            total6 += Convert.ToDouble(fila.Cells("Column10").Value)
        Next
        TextBox11.Text = Convert.ToString(total)
        TextBox12.Text = Convert.ToString(total2)
        TextBox14.Text = Convert.ToString(total3)
        TextBox15.Text = Convert.ToString(total4)
        TextBox16.Text = Convert.ToString(total5)
        TextBox17.Text = Convert.ToString(total6)




    End Sub

    Private Sub SUBIRDATOSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SUBIRDATOSToolStripMenuItem.Click
        Dim Fila As Integer = 2
        Dim Columna As Integer = 1
        Dim RowCount As Integer = DataGridView1.Rows.Count - 2
        Dim ColumnCount As Integer = DataGridView1.Columns.Count - 1

        For nColumna As Integer = 0 To ColumnCount

            Libro.Worksheets("Hoja1").Cells(1, Columna) = DataGridView1.Columns(nColumna).HeaderText
            Libro.Worksheets("Hoja1").Cells(1, Columna).Font.Bold = True

            For nFila As Integer = 0 To RowCount
                Libro.Worksheets("Hoja1").Cells(Fila, Columna) = DataGridView1.Rows(nFila).Cells(nColumna).Value
                Fila = Fila + 1
            Next
            Columna = Columna + 1
            Fila = 2
        Next


        ''
        '  Libro.SaveAs(Filename:="C:\basedatos\test2.xlsx")
        ' Libro.Sheets(1).Cells(1, 1) = TextBox1.Text
        ' Libro.Sheets(1).Cells(1, 2) = TextBox2.Text
        ' Libro.Sheets(1).Cells(1, 3) = TextBox3.Text
        ''
        MsgBox("Los registros se exportaron satisfactoriamente")
        Libro.Save()
        MsgBox("Los cambios han sido guardados en el libro " & Libro.Name)
        ApExcel.Quit()
        Libro = Nothing
        ApExcel = Nothing
    End Sub

    Private Sub VerTablaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VerTablaToolStripMenuItem.Click
        Try
            Llenar()
        Catch ex As Exception
            MsgBox("No hay datos por mostrar")
        End Try
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        Llenar()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim columna1 As String
        Dim columna2 As String
        Dim columna3 As String
        Dim columna4 As String
        Dim columna5 As String
        Dim columna6 As String

        columna1 = Me.DataGridView1.Columns.Item(4).Name.ToString
        columna2 = Me.DataGridView1.Columns.Item(5).Name.ToString
        columna3 = Me.DataGridView1.Columns.Item(6).Name.ToString
        columna4 = Me.DataGridView1.Columns.Item(7).Name.ToString
        columna5 = Me.DataGridView1.Columns.Item(8).Name.ToString
        columna6 = Me.DataGridView1.Columns.Item(9).Name.ToString

        Dim VMaximo1 As Integer = 0
        Dim VMaximo2 As Integer = 0
        Dim VMaximo3 As Integer = 0
        Dim VMaximo4 As Integer = 0
        Dim VMaximo5 As Integer = 0
        Dim VMaximo6 As Integer = 0

        Dim VMinimo1 As Integer = Integer.MaxValue
        Dim VMinimo2 As Integer = Integer.MaxValue
        Dim VMinimo3 As Integer = Integer.MaxValue
        Dim VMinimo4 As Integer = Integer.MaxValue
        Dim VMinimo5 As Integer = Integer.MaxValue
        Dim VMinimo6 As Integer = Integer.MaxValue


        For Each Row As DataGridViewRow In DataGridView1.Rows
            If Not Row.IsNewRow Then

                If Convert.ToInt32(Row.Cells(columna1).Value) > VMaximo1 Then
                    VMaximo1 = Convert.ToInt32(Row.Cells(columna1).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna1).Value) < VMinimo1 Then
                    VMinimo1 = Convert.ToInt32(Row.Cells(columna1).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna2).Value) > VMaximo2 Then
                    VMaximo2 = Convert.ToInt32(Row.Cells(columna2).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna2).Value) < VMinimo2 Then
                    VMinimo2 = Convert.ToInt32(Row.Cells(columna2).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna3).Value) > VMaximo3 Then
                    VMaximo3 = Convert.ToInt32(Row.Cells(columna3).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna3).Value) < VMinimo3 Then
                    VMinimo3 = Convert.ToInt32(Row.Cells(columna3).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna4).Value) > VMaximo4 Then
                    VMaximo4 = Convert.ToInt32(Row.Cells(columna4).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna4).Value) < VMinimo4 Then
                    VMinimo4 = Convert.ToInt32(Row.Cells(columna4).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna5).Value) > VMaximo5 Then
                    VMaximo5 = Convert.ToInt32(Row.Cells(columna5).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna5).Value) < VMinimo5 Then
                    VMinimo5 = Convert.ToInt32(Row.Cells(columna5).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna6).Value) > VMaximo6 Then
                    VMaximo6 = Convert.ToInt32(Row.Cells(columna6).Value)
                End If
                If Convert.ToInt32(Row.Cells(columna6).Value) < VMinimo6 Then
                    VMinimo6 = Convert.ToInt32(Row.Cells(columna6).Value)
                End If
            End If
        Next
        TextBox18.Text = String.Format(VMinimo1)
        TextBox19.Text = String.Format(VMinimo2)
        TextBox20.Text = String.Format(VMinimo3)
        TextBox21.Text = String.Format(VMinimo4)
        TextBox22.Text = String.Format(VMinimo5)
        TextBox23.Text = String.Format(VMinimo6)

        TextBox24.Text = String.Format(VMaximo1)
        TextBox25.Text = String.Format(VMaximo2)
        TextBox26.Text = String.Format(VMaximo3)
        TextBox27.Text = String.Format(VMaximo4)
        TextBox28.Text = String.Format(VMaximo5)
        TextBox29.Text = String.Format(VMaximo6)


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        DataGridView1.Columns(0).Width = 20
        'DataGridView1.Columns(1).Width = 50
        'DataGridView1.Columns(2).Width = 50
        'DataGridView1.Columns(3).Width = 50
        'DataGridView1.Columns(4).Width = 50
        DataGridView1.Rows(2).Height = 50
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick

    End Sub

    Private Sub suma(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Try
            Dim Num1, Num2, Num3, Num4, Num5, Num6, Total, Num7 As String
            Num1 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(4).FormattedValue
            Num2 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(5).FormattedValue
            Num3 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(6).FormattedValue
            Num4 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(7).FormattedValue
            Num5 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(8).FormattedValue
            Num6 = DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(9).FormattedValue

            Num7 = (Val(Num1) + Val(Num2) + Val(Num3) + Val(Num4) + Val(Num5) + Val(Num6))
            Total = (Val(Num7) / 6)
            'TextBox1.Text = Total
            DataGridView1.Rows(DataGridView1.CurrentCellAddress.Y).Cells(10).Value = Total
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class