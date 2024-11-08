Imports System.Data.SqlClient
Public Class FormFamSub
    Public Property Padre1 As Codificacion_Prodcutos
    Public conx As SqlConnection
    Dim Rs19, Rs20 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            conx.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub mostrarInformacionFamilia()
        abrir()
        DataGridView1.Rows.Clear()
        Dim i As Integer = 0
        Dim sql19 As String = "select codigo_fam,descrip_fam,rubro from custom_vianny.dbo.familias where  talm_fam ='" + Label3.Text + "' and ccia_fam ='" + Label2.Text + "'"
        Dim cmd19 As New SqlCommand(sql19, conx)
        Rs19 = cmd19.ExecuteReader()
        While Rs19.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i).Cells(0).Value = Rs19(0)
            DataGridView1.Rows(i).Cells(1).Value = Rs19(1)
            DataGridView1.Rows(i).Cells(2).Value = Rs19(2)
            i = i + 1
        End While
        Rs19.Close()
    End Sub
    Sub mostrarInformacionSubFamilia()
        abrir()
        DataGridView1.Rows.Clear()
        Dim i As Integer = 0
        Dim sql19 As String = "select codigo_sfam,descrip_sfam from custom_vianny.dbo.subfamilias where  talm_sfam='" + Label3.Text + "' and ccia_sfam ='" + Label2.Text + "' and familia_sfam ='" + Label5.Text + "'"
        Dim cmd19 As New SqlCommand(sql19, conx)
        Rs20 = cmd19.ExecuteReader()
        While Rs20.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i).Cells(0).Value = Rs20(0)
            DataGridView1.Rows(i).Cells(1).Value = Rs20(1)
            i = i + 1
        End While
        Rs20.Close()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Label4.Text = "2" Then
            Padre1.TextBox12.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
            Padre1.TextBox14.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
        Else

            Padre1.TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
            Padre1.TextBox13.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
            Padre1.TextBox15.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim()
        End If
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Dim filtroBase As String = ""
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        FiltrarDatos()
    End Sub
    Private Sub FiltrarDatos()
        Dim textoBusqueda As String = TextBox1.Text.Trim().ToLower()

        ' Recorrer todas las filas del DataGridView
        For Each fila As DataGridViewRow In DataGridView1.Rows
            ' Suponiendo que deseas filtrar por la primera o segunda columna (puedes ajustar las columnas)

            Dim descripcion As String = fila.Cells(1).Value.ToString().ToLower()

            ' Si el texto de búsqueda coincide con el código o descripción, mostramos la fila
            If descripcion.Contains(textoBusqueda) Then
                fila.Visible = True
            Else
                fila.Visible = False
            End If
        Next
    End Sub
    Private Sub FormFamSub_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        If Label4.Text = "2" Then
            mostrarInformacionSubFamilia()
        Else
            mostrarInformacionFamilia()
        End If
    End Sub
End Class