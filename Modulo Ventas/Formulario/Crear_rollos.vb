Imports System.Data.SqlClient
Public Class Crear_rollos
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
    Private Sub Crear_rollos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Add()
        abrir()
        Dim cmd As New SqlDataAdapter("exec custom_vianny.dbo.spu_cargatelasOT '01','" + TextBox1.Text + "','" + TextBox2.Text + "'", conx)

        Dim da As New DataTable

        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            DataGridView2.DataSource = da
            TextBox3.Text = DataGridView2.Rows(0).Cells(17).Value
            TextBox4.Text = DataGridView2.Rows(0).Cells(18).Value
            TextBox5.Text = DataGridView2.Rows(0).Cells(15).Value
            TextBox6.Text = DataGridView2.Rows(0).Cells(14).Value
            TextBox7.Text = DataGridView2.Rows(0).Cells(16).Value
            TextBox8.Text = DataGridView2.Rows(0).Cells(3).Value
            TextBox9.Text = DataGridView2.Rows(0).Cells(4).Value

            DataGridView1.Rows(0).Cells(1).Value = DataGridView2.Rows(0).Cells(8).Value
            DataGridView1.Rows(0).Cells(2).Value = DataGridView2.Rows(0).Cells(13).Value
            DataGridView1.Rows(0).Cells(3).Value = DataGridView2.Rows(0).Cells(9).Value
            DataGridView1.Rows(0).Cells(4).Value = DataGridView2.Rows(0).Cells(10).Value
            DataGridView1.Rows(0).Cells(5).Value = DataGridView2.Rows(0).Cells(11).Value
            DataGridView1.Rows(0).Cells(6).Value = DataGridView2.Rows(0).Cells(21).Value
            DataGridView1.Rows(0).Cells(7).Value = DataGridView2.Rows(0).Cells(22).Value
            DataGridView1.Rows(0).Cells(8).Value = DataGridView2.Rows(0).Cells(23).Value
        Else
            DataGridView2.DataSource = ""
        End If
        DataGridView1.Select()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.Rows(0).Cells(0).Value = True Then
            Dim po As String
            po = TextBox1.Text + TextBox2.Text
            Dim cmd15 As New SqlCommand("exec  custom_vianny.dbo.spu_generarollosOrdTejido '01',@po,@correlatico,@poc,'001',@pesorollo,@kgprogrmamr", conx)
            cmd15.Parameters.AddWithValue("@po", TextBox1.Text)
            cmd15.Parameters.AddWithValue("@correlatico", TextBox2.Text)
            cmd15.Parameters.AddWithValue("@poc", po)
            cmd15.Parameters.AddWithValue("@pesorollo", Trim(DataGridView1.Rows(0).Cells(4).Value))
            cmd15.Parameters.AddWithValue("@kgprogrmamr", Trim(DataGridView1.Rows(0).Cells(8).Value))

            cmd15.ExecuteNonQuery()
            MsgBox("SE GENERARON LOS ROLLOS CORRECTAMENTE")
            Me.Close()
        Else
            MsgBox("NO SELECCIONO NINGUN ITEM")
        End If

    End Sub
End Class