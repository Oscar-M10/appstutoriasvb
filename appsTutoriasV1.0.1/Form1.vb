Imports Reporting
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub AlumnoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlumnoToolStripMenuItem.Click

        Try
            Alumnos.Show()
        Catch ex As Exception
            MsgBox("NO SE A ENCONTRADO SU REGISTRO")
        End Try

    End Sub

    Private Sub RegistrarseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegistrarseToolStripMenuItem.Click
        Registro.Show()
    End Sub
End Class
