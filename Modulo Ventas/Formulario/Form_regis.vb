Imports System.Data.SqlClient
Public Class Form_regis
    Public conx As SqlConnection
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
    Private Sub Form_regis_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Label1.Text = MDIParent1.Label5.Text
        Label2.Text = MDIParent1.Label9.Text
        Dim cmd As New SqlDataAdapter("select ccia, ctab, cele AS REGISTRO,dele AS RUC, nele AS COMPROBANTE,codigo AS CUENTA,valo from tab0100 where valo ='500' and ccia ='" + Label2.Text + "' and ctab ='" + Label1.Text + "'", conx)

        Dim da As New DataTable

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(6).Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim u As Integer
        u = DataGridView1.Rows.Count
        For i = 0 To u - 1

            Dim cmd17 As New SqlCommand("update custom_vianny.dbo.cac3p00 set fuen_3a=@fuen_3a where ncom_3a =@ncom_3 and ccia_3a =@ccia_3 and cperiodo_3a =@cperiodo_3 and fich_3a =@ruc AND ccom_3a =@ccom_3 and ndoc_3a =@ndoc_3a and cuen_3a =@cuenta", conx)
            cmd17.Parameters.AddWithValue("@fuen_3a", "1")
            cmd17.Parameters.AddWithValue("@ncom_3", Trim(DataGridView1.Rows(i).Cells(2).Value))
            cmd17.Parameters.AddWithValue("@ccia_3", Trim(DataGridView1.Rows(i).Cells(0).Value))
            cmd17.Parameters.AddWithValue("@cperiodo_3", Trim(DataGridView1.Rows(i).Cells(1).Value))
            cmd17.Parameters.AddWithValue("@ruc", Trim(DataGridView1.Rows(i).Cells(3).Value))
            cmd17.Parameters.AddWithValue("@ndoc_3a", Trim(DataGridView1.Rows(i).Cells(4).Value))
            cmd17.Parameters.AddWithValue("@cuenta", Trim(DataGridView1.Rows(i).Cells(5).Value))
            cmd17.Parameters.AddWithValue("@ccom_3", Trim(DataGridView1.Rows(i).Cells(6).Value))
            'MsgBox(Trim(cco))
            cmd17.CommandTimeout = 500
            cmd17.ExecuteNonQuery()
        Next
        MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
        Me.Close()
        proveedores.TextBox1.Text = MDIParent1.Label5.Text
        proveedores.TextBox2.Text = MDIParent1.Label7.Text
        proveedores.Label1.Text = MDIParent1.Label9.Text
        proveedores.Show()
    End Sub
End Class