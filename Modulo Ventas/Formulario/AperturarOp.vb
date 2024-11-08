Imports System.Data.SqlClient
Public Class AperturarOp
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
    Private Sub AperturarOp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton2.Checked = True
        Button1.Visible = False
        SinAprobar()
    End Sub
    Sub SinAprobar()
        da.Clear()
        DataGridView1.DataSource = ""
        abrir()
        Dim cmd As New SqlDataAdapter("select ncom_3 as Op,c.nomb_10 as Cliente,descri_3 as Producto, cantp_3 Programado from custom_vianny.dbo.qag0300 q left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 = c.fich_10  where q.ccia ='" + Label4.Text + "' and ncom_3 like '01%' and finped_3 ='0'", conx)
        cmd.Fill(da)

        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(1).Width = 320
            DataGridView1.Columns(2).Width = 320
            DataGridView1.Columns(3).Width = 80
        End If
        Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
        Label7.Text = Convert.ToString(selectedRow.Cells("Op").Value.ToString())
    End Sub
    Sub Colores()
        Dim va As Integer
        va = DataGridView1.Rows.Count

        For i = 0 To va - 1
            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
            DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
        Next
    End Sub
    Sub Aprobar()
        da.Clear()
        DataGridView1.DataSource = ""
        abrir()
        Dim cmd As New SqlDataAdapter("select ncom_3 as Op,c.nomb_10 as Cliente,descri_3 as Producto, cantp_3 Programado from custom_vianny.dbo.qag0300 q left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 = c.fich_10  where q.ccia ='" + Label4.Text + "' and ncom_3 like '01%' and finped_3 ='1' and YEAR(fcom_3) = YEAR(GETDATE()) ", conx)
        cmd.Fill(da)

        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(1).Width = 320
            DataGridView1.Columns(2).Width = 320
            DataGridView1.Columns(3).Width = 80
        End If
        Colores()
        Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
        Label7.Text = Convert.ToString(selectedRow.Cells("Op").Value.ToString())
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Dim cmd20 As New SqlCommand("update custom_vianny.dbo.qag0300 set finped_3 ='0', fcome3_3 =@fcome3_3 WHERE ncom_3 =@ncom_3 AND ccia =@ccia ", conx)
        cmd20.Parameters.AddWithValue("@ncom_3", Trim(Label7.Text))
        cmd20.Parameters.AddWithValue("@ccia", Trim(Label4.Text))
        cmd20.Parameters.AddWithValue("@fcome3_3", DateTimePicker1.Value)
        cmd20.ExecuteNonQuery()
        MsgBox("SE ABRIO CORRECTAMENTE")
        Aprobar()
        If DataGridView1.Rows.Count > 0 Then
            DataGridView1.Rows(0).Selected = True
        End If
    End Sub
    Dim OP As String
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            Label7.Text = Convert.ToString(selectedRow.Cells("Op").Value.ToString())
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()

        Dim cmd20 As New SqlCommand("update custom_vianny.dbo.qag0300 set finped_3 ='1' WHERE ncom_3 =@ncom_3 AND ccia =@ccia ", conx)
        cmd20.Parameters.AddWithValue("@ncom_3", Trim(Label7.Text))
        cmd20.Parameters.AddWithValue("@ccia", Trim(Label4.Text))
        cmd20.ExecuteNonQuery()
        MsgBox("SE CERRO CORRECTAMENTE")
        SinAprobar()
        If DataGridView1.Rows.Count > 0 Then
            DataGridView1.Rows(0).Selected = True
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        If RadioButton2.Checked = True Then
            SinAprobar()
            Button1.Visible = False
            Button2.Visible = True
        Else
            If RadioButton1.Checked = True Then
                Aprobar()
                Button1.Visible = True
                Button2.Visible = False
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            Label7.Text = Convert.ToString(selectedRow.Cells("Op").Value.ToString())
        End If
    End Sub
    Private Sub buscar0()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Op" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 320
                DataGridView1.Columns(2).Width = 320
                DataGridView1.Columns(3).Width = 80
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Label7.Text = Convert.ToString(selectedRow.Cells("Op").Value.ToString())

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar0()
        If RadioButton1.Checked = True Then
            Colores()
        End If

    End Sub
End Class