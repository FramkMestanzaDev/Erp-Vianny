Imports System.Data.SqlClient
Public Class Lotes
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub buscar()
        Try

            Dim ds As New DataSet
            ds.Tables.Add(dg2.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "PARTIDA" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 200
            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception

        End Try



    End Sub
    Sub llenar_combo_box2()

        Try

            conn = New SqlDataAdapter(" select talm_15m as almacen, talm_15m + ' | '+ nomb_15m as descrip from custom_vianny.dbo.Mag1500 where ccia =01", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "descrip"
            ComboBox1.ValueMember = "almacen"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Dim dg2, da As New DataTable
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Dim Rsr22 As SqlDataReader
    Private Sub Lotes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box2()
        Dim almacen As String
        Dim cmd2 As New SqlDataAdapter("exec CaeSoft_ReporteStockFisico_ayuda '01','" + Trim(Label4.Text) + "','" + Trim(Label3.Text) + "','" + Trim(Label3.Text) + "',2,'1'", conx)

        cmd2.Fill(dg2)
        If dg2.Rows.Count <> 0 Then
            bs.DataSource = dg2

            DataGridView1.DataSource = bs
            DataGridView1.Columns(0).Width = 140
            DataGridView1.Columns(1).Width = 120
            DataGridView1.Columns(2).Width = 320
        Else
            DataGridView1.DataSource = ""
        End If
        abrir()

        Dim sql1022 As String = "select talm_17 from custom_vianny.dbo.Cag1700 where linea_17 ='" + Trim(Label3.Text) + "' and ccia ='01'"
        Dim cmd1022 As New SqlCommand(sql1022, conx)
        Rsr22 = cmd1022.ExecuteReader()
        If Rsr22.Read() = True Then
            almacen = Rsr22(0)
        Else
            almacen = "0"
        End If
        TextBox3.Text = almacen
        Rsr22.Close()

        enunciado2 = New SqlCommand("select talm_15m as almacen, talm_15m + ' | '+ nomb_15m as descrip from custom_vianny.dbo.Mag1500 where ccia ='01' and  talm_15m ='" + almacen + "'", conx)
        respuesta2 = enunciado2.ExecuteReader
        While respuesta2.Read
            ComboBox1.Text = respuesta2.Item("descrip")
        End While
        respuesta2.Close()
    End Sub
    Dim bs As New BindingSource()

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        bs.Filter = String.Format("Nombre LIKE '%{0}%'", TextBox3.Text)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox3.Text = ComboBox1.SelectedValue.ToString
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        bs.Filter = String.Format("Lote LIKE '%{0}%'", TextBox3.Text)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = ""
        da.Clear()
        abrir()
        llenar_combo_box2()
        Dim cmd As New SqlDataAdapter("exec CaeSoft_ReporteStockFisico_ayuda '01','" + Trim(Label4.Text) + "','" + Trim(Label3.Text) + "','" + Trim(Label3.Text) + "',2,'" + Trim(TextBox3.Text) + "'", conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(0).Width = 140
            DataGridView1.Columns(1).Width = 120
            DataGridView1.Columns(2).Width = 320

        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Otejeduria.DataGridView2.Rows(Label5.Text).Cells(5).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Otejeduria.DataGridView2.Rows(Label5.Text).Cells(6).Value = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        Otejeduria.DataGridView2.Rows(Label5.Text).Cells(6).Value = (Otejeduria.DataGridView1.Rows(0).Cells(10).Value * Otejeduria.DataGridView2.Rows(Label5.Text).Cells(2).Value) / 100
        'Otejeduria.DataGridView2.Rows(Label5.Text).Cells(6).Value = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        Me.Close()
    End Sub
End Class