Imports System.Data.SqlClient
Public Class Pedidos
    Dim dt As New DataTable
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
    Private Sub Pedidos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim gf As New fnopedido
        Dim gf1 As New vnapedido
        gf1.gperiodo = Label3.Text
        If Label5.Text = "01" Then
            gf1.galmacen = 1
        Else
            gf1.galmacen = "67"
        End If

        'dt = gf.validar_almacen2(gf1)
        'If dt.Rows.Count <> 0 Then
        '    DataGridView1.DataSource = dt
        '    Dim k As Integer
        '    k = DataGridView1.Rows.Count


        '    DataGridView1.Columns.Item(1).Visible = False
        '    DataGridView1.Columns.Item(2).Visible = False
        '    DataGridView1.Columns.Item(4).Visible = False
        '    DataGridView1.Columns.Item(6).Visible = False
        '    DataGridView1.Columns.Item(7).Visible = False
        '    DataGridView1.Columns.Item(8).Visible = False

        '    DataGridView1.Columns.Item(3).Width = 350

        'End If

        abrir()
        Dim cmd1 As New SqlDataAdapter("select numero_pedido as PEDIDO, fecha AS FECHA, codigo_cli AS RUC,c.r_social as CLIENTE,total AS TOTAL from nota_pedido n left join cliente c on n.codigo_cli = c.COD_CLI where codigo_cli ='" + Label7.Text + "' and n.vendedor ='" + Label6.Text + "' and year(fecha) ='" + Label3.Text + "'", conx)
        Dim da1 As New DataTable
        cmd1.Fill(da1)

        If da1.Rows.Count <> 0 Then
            DataGridView1.DataSource = da1
            DataGridView1.Columns(0).Width = 85
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(2).Width = 85
            DataGridView1.Columns(3).Width = 280
            DataGridView1.Columns(4).Width = 100

        End If
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PEDIDO" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns.Item(1).Visible = False
                DataGridView1.Columns.Item(2).Visible = False
                DataGridView1.Columns.Item(4).Visible = False
                DataGridView1.Columns.Item(6).Visible = False
                DataGridView1.Columns.Item(7).Visible = False
                DataGridView1.Columns.Item(8).Visible = False

                DataGridView1.Columns.Item(3).Width = 350
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns.Item(1).Visible = False
                DataGridView1.Columns.Item(2).Visible = False
                DataGridView1.Columns.Item(4).Visible = False
                DataGridView1.Columns.Item(6).Visible = False
                DataGridView1.Columns.Item(7).Visible = False
                DataGridView1.Columns.Item(8).Visible = False

                DataGridView1.Columns.Item(3).Width = 350
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If Label4.Text = 1 Then
            Registro_Facturas.TextBox34.Text = DataGridView1.CurrentRow.Cells(0).Value
        Else
            If Label4.Text = 2 Then
                Registro_Ventas.TextBox34.Text = DataGridView1.CurrentRow.Cells(0).Value
            End If

        End If


        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar1()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar2()
    End Sub
End Class