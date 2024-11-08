Imports System.Data.SqlClient
Public Class Ventas_Cliente
    Dim dt, dt1 As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.F1 Then
            If e.KeyCode = Keys.F1 Then
                Clientes.TextBox3.Text = "333"
                Clientes.Show()
            End If
        End If
    End Sub

    Private Sub Ventas_Cliente_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim jj As New fventas
        Dim kk As New vventas_n

        kk.gfichas = TextBox1.Text
        kk.gfini = DateTimePicker1.Value
        kk.gffin = DateTimePicker2.Value
        dt = jj.venta_cliente(kk)

        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(0).Width = 68
            DataGridView1.Columns(8).Width = 250
        Else
            DataGridView1.DataSource = ""
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim fg As New Exportar
        fg.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim j As New vtejeduria
        Dim ll As New ftejeduria

        j.gfechaini = DateTimePicker1.Value
        j.gfechafin = DateTimePicker2.Value

        dt1 = ll.buscar_ventas_m(j)
        If dt1.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt1
        End If
    End Sub
End Class