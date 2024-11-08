Imports System.Data.SqlClient
Public Class Separar_Partidas
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
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
    Function ELIMINAR1(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Dim agregar As String = "delete from pedido_separado where part_lote ='" + Trim(TextBox1.Text) + "' and EMPRESA ='" + Trim(Label6.Text) + "'"
        ELIMINAR1(agregar)

        Dim cmd200 As New SqlCommand("insert into pedido_separado(part_lote,estado,codigo,cantidad,numero_pedido,EMPRESA,ALMACEN) values (@part_lote,@estado,@codigo,@cantidad,@numero_pedido,@EMPRESA,@ALMACEN)", conx)
        cmd200.Parameters.AddWithValue("@part_lote", Trim(TextBox1.Text))
        cmd200.Parameters.AddWithValue("@codigo", Trim(TextBox2.Text))
        cmd200.Parameters.AddWithValue("@cantidad", Trim(TextBox5.Text))
        cmd200.Parameters.AddWithValue("@estado", 1)
        cmd200.Parameters.AddWithValue("@numero_pedido", "0000000000")
        cmd200.Parameters.AddWithValue("@EMPRESA", Trim(Label6.Text))
        cmd200.Parameters.AddWithValue("@ALMACEN", Trim(TextBox4.Text))
        cmd200.ExecuteNonQuery()
        MsgBox("SE SEPARO LA PARTIDA SOLICITADA")
        Me.Close()
    End Sub
    Dim dt10 As New DataTable
    Private Sub Separar_Partidas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd5 As New SqlDataAdapter("select * from pedido_separado WHERE numero_pedido ='0000000000'  and estado =1", conx)
        cmd5.Fill(dt10)
        If dt10.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt10
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Visible = False
            Dim y As Integer
            y = DataGridView1.Rows.Count
            For i = 0 To y - 1
                If DataGridView1.Rows(i).Cells(4).Value = 1 Then
                    DataGridView1.Rows(i).Cells(4).Value = "SEPARADO"
                End If
            Next


        Else
            DataGridView1.DataSource = ""
        End If


    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
        TextBox3.Text = ""
        TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value)
        TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
        Label9.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
        Me.TabControl1.SelectedIndex = 0
        Button1.Enabled = False
        Button2.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        Dim cmd15 As New SqlCommand("UPDATE pedido_separado SET estado =0 WHERE  id_pedido =@pedido", conx)
        cmd15.Parameters.AddWithValue("@pedido", Trim(Label9.Text))

        cmd15.ExecuteNonQuery()
        MsgBox("SE CANCELO LA SEPARACION DE PEDIDO")
        Me.Close()
    End Sub
End Class