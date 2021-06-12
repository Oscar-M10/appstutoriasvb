Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports System.Windows.Forms
Module Reporte
    Dim xlibro As Microsoft.Office.Interop.Excel.Application
    Public FECHA As Date
    Public FACTURA As Integer
    Public RANGO As Integer = 11
    Public CLIENTE, CODIGO, DESCRIPCION, CANTIDAD, PRECIO, IMPUESTOS, SUBTOTAL As String

    Dim dtFactura As Data.DataTable = New Data.DataTable

    Public Sub AbrirDocumentoImprimir()
        'El siguiente codigo es para crear la ruta,entre comillas se pone la ruta donde esta el libro
        Dim Ruta As String = Path.Combine(Directory.GetCurrentDirectory(), "basedatos.xls")

        'El siguiente codigo es para abrir el libro y hacerlo visible, si se quiere dejar el libro oculto, se cambia la palabra True por False
        xlibro = CreateObject("Excel.Application")
        xlibro.Workbooks.Open(Ruta)
        xlibro.Visible = True

        xlibro.Sheets("factura").Select() 'Nombre del libro



    End Sub

    Public Sub LlenarCliente()
        Dim cadena As String = "provider=Microsoft.Jet.OLEDB.4.0;Data Source='basedatos.xls';Extended Properties=Excel 8.0;"
        Dim TotalFacturado As Double = 0

        Dim conn As New Data.OleDb.OleDbConnection(cadena)
        conn.Open()

        Dim da As New OleDbDataAdapter("select * from [detalle_ventas$]", conn)
        Dim ds As New DataSet
        da.Fill(ds)

        dtFactura = ds.Tables(0)

        Dim bs As New BindingSource
        bs.DataSource = dtFactura

        bs.Filter = "ID_FACTURA =" & FACTURA

        xlibro.Range("B6").Value = FACTURA
        xlibro.Range("B7").Value = FECHA
        xlibro.Range("B8").Value = CLIENTE

        For Each view As DataRowView In bs
            Dim row = view.Row
            CODIGO = row("CODIGO")
            DESCRIPCION = row("DESCRIPCION")
            CANTIDAD = row("CANTIDAD")
            PRECIO = row("PRECIO")
            IMPUESTOS = row("IMPUESTOS")

            xlibro.Range("A" & RANGO).Value = CODIGO
            xlibro.Range("B" & RANGO).Value = DESCRIPCION
            xlibro.Range("C" & RANGO).Value = CANTIDAD
            xlibro.Range("D" & RANGO).Value = PRECIO
            xlibro.Range("E" & RANGO).Value = IMPUESTOS
            xlibro.Range("F" & RANGO).Value = CANTIDAD * PRECIO + IMPUESTOS
            TotalFacturado = TotalFacturado + (CANTIDAD * PRECIO + IMPUESTOS)
            RANGO += 1

        Next

        RANGO += 1
        With xlibro.Range("A" & RANGO & ":" & "F" & RANGO)
            .Interior.Color = RGB(179, 179, 179)
            .Font.Bold = True
        End With



        xlibro.Range("E" & RANGO).Value = "Total:"
        xlibro.Range("F" & RANGO).Value = TotalFacturado

        'xlibro.ActiveWorkbook.Save()
        'xlibro.Quit()

        conn.Close()
    End Sub

    Private Sub KillExcelProcess()
        Try
            Dim Xcel() As Process = Process.GetProcessesByName("EXCEL")
            For Each Process As Process In Xcel
                Process.Kill()
            Next
        Catch ex As Exception
        End Try
    End Sub


    Public Sub EliminaFilasVacias()
        KillExcelProcess()
        AbrirDocumentoImprimir()

        xlibro.Sheets("cliente").Select()

        Dim lRow = xlibro.ActiveSheet.UsedRange.Rows.Count

        For fila As Integer = 1 To lRow
            If String.IsNullOrEmpty(xlibro.Range("A" & fila).Value) Then
                xlibro.Range("A" & fila).Select()
                xlibro.Range("A" & fila & ":Z" & fila).Delete()
                xlibro.ActiveWorkbook.Save()
            End If
        Next

        '---------------------------------------------------------------
        xlibro.Sheets("factura").Select()
        lRow = xlibro.ActiveSheet.UsedRange.Rows.Count

        For fila As Integer = 1 To lRow
            If String.IsNullOrEmpty(xlibro.Range("A" & fila).Value) Then
                xlibro.Range("A" & fila).Select()
                xlibro.Range("A" & fila & ":Z" & fila).Delete()
                xlibro.ActiveWorkbook.Save()
            End If
        Next

        '-------------------------------------------------------------
        xlibro.Sheets("productos").Select()
        lRow = xlibro.ActiveSheet.UsedRange.Rows.Count

        For fila As Integer = 1 To lRow
            If String.IsNullOrEmpty(xlibro.Range("A" & fila).Value) Then
                xlibro.Range("A" & fila).Select()
                xlibro.Range("A" & fila & ":Z" & fila).Delete()
                xlibro.ActiveWorkbook.Save()
            End If
        Next

        '----------------------------------------------------------------
        xlibro.Sheets("ventas").Select()
        lRow = xlibro.ActiveSheet.UsedRange.Rows.Count

        For fila As Integer = 1 To lRow
            If String.IsNullOrEmpty(xlibro.Range("A" & fila).Value) Then
                xlibro.Range("A" & fila).Select()
                xlibro.Range("A" & fila & ":Z" & fila).Delete()
                xlibro.ActiveWorkbook.Save()
            End If
        Next

        '--------------------------------------------------------------
        xlibro.Sheets("detalle_ventas").Select()
        lRow = xlibro.ActiveSheet.UsedRange.Rows.Count

        For fila As Integer = 1 To lRow
            If String.IsNullOrEmpty(xlibro.Range("A" & fila).Value) Then
                xlibro.Range("A" & fila).Select()
                xlibro.Range("A" & fila & ":Z" & fila).Delete()
                xlibro.ActiveWorkbook.Save()
            End If
        Next

    End Sub
    Public Function nReg(Hoja As Microsoft.Office.Interop.Excel.Worksheet, nFila As Long, nColumna As Long)
        Do Until Hoja.Cells(nFila, nColumna).Value = ""
            nFila = nFila + 1
        Loop
        Return nFila
    End Function
End Module
