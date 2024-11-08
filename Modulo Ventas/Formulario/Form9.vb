Imports System.Data.SqlClient
Public Class Form9
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado4 As SqlCommand
    Public respuesta4 As SqlDataReader
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Rpt_frecuencia_ventas.TextBox1.Text = TextBox3.Text
        Rpt_frecuencia_ventas.TextBox2.Text = TextBox2.Text
        Rpt_frecuencia_ventas.TextBox3.Text = TextBox5.Text
        Rpt_frecuencia_ventas.TextBox4.Text = ComboBox1.Text
        Rpt_frecuencia_ventas.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox5.Text = ComboBox1.SelectedValue.ToString
    End Sub

    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  rpm_ven='2'  and admin_ven in (0,2)", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "alias_ven"
            ComboBox1.ValueMember = "codigo_ven"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box2()

        TextBox5.Text = "0005"
    End Sub
End Class