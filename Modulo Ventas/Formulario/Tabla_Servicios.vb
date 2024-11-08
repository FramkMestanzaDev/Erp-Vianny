Imports System.Data.SqlClient
Public Class SERVICIOS
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

    Dim da As New DataTable
    Sub mostrar_informacion()
        abrir()
        da.Clear()
        DataGridView1.DataSource = Nothing
        If Trim(ComboBox1.Text) = "LAVANDERIA" Then
            Dim cmd As New SqlDataAdapter("SELECT fich_10 as RUC,nomb_10 AS EMPRESA,nombre_10 AS NOMBRE,apepat_10 AS APELLIDO_P,apemat_10 AS APELLIDO_M,emailfe_10 AS EMAIL,fono1_10 AS TELEFONO FROM custom_vianny.dbo.cag1000 WHERE Url_10='2'", conx)
            cmd.Fill(da)
            DataGridView1.DataSource = da
            DataGridView1.Columns(2).Width = 210
            DataGridView1.Columns(1).Width = 90
            DataGridView1.Columns(3).Width = 140
            DataGridView1.Columns(4).Width = 140
            DataGridView1.Columns(5).Width = 140
            DataGridView1.Columns(6).Width = 165
        Else
            If Trim(ComboBox1.Text) = "CONFECCION" Then
                Dim cmd As New SqlDataAdapter("SELECT fich_10 as RUC,nomb_10 AS EMPRESA,nombre_10 AS NOMBRE,apepat_10 AS APELLIDO_P,apemat_10 AS APELLIDO_M,emailfe_10 AS EMAIL,fono1_10 AS TELEFONO FROM custom_vianny.dbo.cag1000 WHERE Url_10='1'", conx)
                cmd.Fill(da)
                DataGridView1.DataSource = da
                DataGridView1.Columns(2).Width = 210
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(3).Width = 140
                DataGridView1.Columns(4).Width = 140
                DataGridView1.Columns(5).Width = 140
                DataGridView1.Columns(6).Width = 165
            End If
        End If
    End Sub
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        mostrar_informacion()

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Registro_Servicios.TextBox1.Text = Trim(DataGridView1.Rows(num1).Cells(1).Value)
            Registro_Servicios.TextBox2.Text = Trim(DataGridView1.Rows(num1).Cells(2).Value)
            Registro_Servicios.TextBox3.Text = Trim(DataGridView1.Rows(num1).Cells(3).Value)
            Registro_Servicios.TextBox4.Text = Trim(DataGridView1.Rows(num1).Cells(4).Value)
            Registro_Servicios.TextBox5.Text = Trim(DataGridView1.Rows(num1).Cells(5).Value)
            Registro_Servicios.TextBox6.Text = Trim(DataGridView1.Rows(num1).Cells(6).Value)
            Registro_Servicios.TextBox7.Text = Trim(DataGridView1.Rows(num1).Cells(7).Value)
            Registro_Servicios.ShowDialog()


        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub SERVICIOS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class