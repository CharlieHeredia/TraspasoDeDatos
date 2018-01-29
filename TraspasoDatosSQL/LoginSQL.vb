Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class LoginSQL

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Falta llenar un campo")
        Else
            UsuarioBD = TextBox1.Text.Trim()
            Contra = TextBox2.Text.Trim()
            BaseDatos = TextBox3.Text.Trim()
            Tabla = TextBox4.Text.Trim()
            Instancia = ComboBox1.SelectedItem.ToString().Trim()
            'Instancia = "MACBOOKPRO\MACBOOKSQLEXP"
            DatosCompletos = True
            Me.Close()
        End If
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter
        'SIN USO
    End Sub

    Private Sub LoginSQL_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BackgroundWorker1.RunWorkerAsync()
        Control.CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim instan As SqlDataSourceEnumerator = SqlDataSourceEnumerator.Instance
        Dim talablainsta = New DataTable()
        talablainsta = instan.GetDataSources()

        For Each row As DataRow In talablainsta.Rows
            ComboBox1.Items.Add(row.Item(0) & "\" & row.Item(1))
        Next
    End Sub
End Class