Imports System.Data.SqlClient
Public Class Agregar_Factura
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Agregar_Factura_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Liquidacion_Produccion.DataGridView1.Rows(Label4.Text).Cells(7).Value = Trim(TextBox3.Text)
        Dim cmd15 As New SqlCommand("insert into factura_servicio_descuento (op,empresa,area,factura) values (@op,@empresa,@area,@factura)", conx)
        cmd15.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
        Select Case Trim(TextBox2.Text)
            Case "CORTE" : cmd15.Parameters.AddWithValue("@area", "0300")
            Case "CONFECCION" : cmd15.Parameters.AddWithValue("@area", "0400")
            Case "ACABADOS" : cmd15.Parameters.AddWithValue("@area", "0500")
            Case "LAVANDERIA" : cmd15.Parameters.AddWithValue("@area", "0600")
            Case "APLICACIONES Y PIEZAS" : cmd15.Parameters.AddWithValue("@area", "0700")
        End Select

        cmd15.Parameters.AddWithValue("@empresa", Trim(Label5.Text))
        cmd15.Parameters.AddWithValue("@factura", Trim(TextBox3.Text))
        cmd15.ExecuteNonQuery()
        Me.Close()
    End Sub
End Class