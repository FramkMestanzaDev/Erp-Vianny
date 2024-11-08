Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class Despachos_adicionales
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim k As Integer

        k = DataGridView1.Rows.Count
        For i = 0 To k - 1

            Dim cmd20 As New SqlCommand("insert into despachos_adicionales(estado,fecha,chofer,vehiculo,observacion,estibador,pedido,dir_despacho,cliente)
values (@estado,@fecha,@chofer,@vehiculo,@observacion,@estibador,@pedido,@dir_despacho,@cliente)", conx)
            cmd20.Parameters.AddWithValue("@estado", 0)
            cmd20.Parameters.AddWithValue("@fecha", DBNull.Value)
            cmd20.Parameters.AddWithValue("@chofer", Trim(DataGridView1.Rows(i).Cells(0).Value))
            cmd20.Parameters.AddWithValue("@vehiculo", Trim(DataGridView1.Rows(i).Cells(1).Value))
            cmd20.Parameters.AddWithValue("@observacion", Trim(DataGridView1.Rows(i).Cells(2).Value))
            cmd20.Parameters.AddWithValue("@estibador", Trim(DataGridView1.Rows(i).Cells(3).Value))
            cmd20.Parameters.AddWithValue("@pedido", Trim(DataGridView1.Rows(i).Cells(4).Value))
            cmd20.Parameters.AddWithValue("@dir_despacho", Trim(DataGridView1.Rows(i).Cells(5).Value))
            cmd20.Parameters.AddWithValue("@cliente", Trim(DataGridView1.Rows(i).Cells(6).Value))
            cmd20.ExecuteNonQuery()
        Next
        buscar()
        DataGridView3.DataSource = ""
        DataGridView1.Rows.Clear()
    End Sub
    Sub buscar()
        DataGridView2.Rows.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select chofer,vehiculo,observacion,estibador,pedido,dir_despacho,cliente,id_despacho from  despachos_adicionales  where estado = 0", conx)

        Dim da As New DataTable

        cmd.Fill(da)

        DataGridView3.DataSource = da
        Dim j As Integer
        j = DataGridView3.Rows.Count
        If j > 0 Then
            DataGridView2.Rows.Add(j)
            llenar_combo_box3()
            llenar_combo_box4()
            For i = 0 To j - 1

                DataGridView2.Rows(i).Cells(2).Value = (Trim(DataGridView3.Rows(i).Cells(0).Value)).ToString
                DataGridView2.Rows(i).Cells(3).Value = (Trim(DataGridView3.Rows(i).Cells(1).Value)).ToString
                DataGridView2.Rows(i).Cells(4).Value = DataGridView3.Rows(i).Cells(2).Value
                DataGridView2.Rows(i).Cells(5).Value = (Trim(DataGridView3.Rows(i).Cells(3).Value)).ToString
                DataGridView2.Rows(i).Cells(6).Value = DataGridView3.Rows(i).Cells(4).Value
                DataGridView2.Rows(i).Cells(7).Value = DataGridView3.Rows(i).Cells(5).Value
                DataGridView2.Rows(i).Cells(8).Value = DataGridView3.Rows(i).Cells(6).Value
                DataGridView2.Rows(i).Cells(9).Value = DataGridView3.Rows(i).Cells(7).Value
            Next
        End If


    End Sub

    Sub llenar_combo_box3()
        Try
            Dim lsQuery As String = "select distinct chofer  from ingre_transportista"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            CHO.DataSource = loDataTable

            CHO.DisplayMember = "chofer"
            CHO.ValueMember = "chofer"
            CHO.DropDownWidth = 200
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box()
        Try
            Dim lsQuery As String = "select distinct chofer  from ingre_transportista"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            CHOFER.DataSource = loDataTable

            CHOFER.DisplayMember = "chofer"
            CHOFER.ValueMember = "chofer"
            CHOFER.DropDownWidth = 200
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box1()
        Try
            Dim lsQuery As String = "select distinct placa  from ingre_transportista"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            VEHICULO.DataSource = loDataTable

            VEHICULO.DisplayMember = "placa"
            VEHICULO.ValueMember = "placa"
            VEHICULO.DropDownWidth = 200
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box4()
        Try
            Dim lsQuery As String = "select distinct placa  from ingre_transportista"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            ve.DataSource = loDataTable

            ve.DisplayMember = "placa"
            ve.ValueMember = "placa"
            ve.DropDownWidth = 200
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim k As Integer

        k = DataGridView2.Rows.Count
        For i = 0 To k - 1
            If DataGridView2.Rows(i).Cells(0).Value = True Then
                Dim cmd20 As New SqlCommand("UPDATE despachos_adicionales SET estado=@estado, fecha=@fecha,chofer=@chofer,vehiculo=@vehiculo,observacion=@observacion,estibador=@estibador WHERE id_despacho =@ID", conx)
                cmd20.Parameters.AddWithValue("@estado", 1)
                cmd20.Parameters.AddWithValue("@fecha", Trim(DataGridView2.Rows(i).Cells(1).Value))
                cmd20.Parameters.AddWithValue("@chofer", Trim(DataGridView2.Rows(i).Cells(2).Value))
                cmd20.Parameters.AddWithValue("@vehiculo", Trim(DataGridView2.Rows(i).Cells(3).Value))
                cmd20.Parameters.AddWithValue("@observacion", Trim(DataGridView2.Rows(i).Cells(4).Value))
                cmd20.Parameters.AddWithValue("@estibador", Trim(DataGridView2.Rows(i).Cells(5).Value))
                cmd20.Parameters.AddWithValue("@ID", Trim(DataGridView2.Rows(i).Cells(9).Value))

                cmd20.ExecuteNonQuery()
            End If

        Next
        DataGridView2.Rows.Clear()
        buscar()
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        If e.ColumnIndex = 0 Then
            DataGridView2.Rows(e.RowIndex).Cells(1).Value = DateTime.Now.ToString("yyyy/MM/dd")

        End If
    End Sub

    Private Sub Despachos_adicionales_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Add(1)
        abrir()
        llenar_combo_box()
        llenar_combo_box1()
        buscar()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim U As Integer
        U = DataGridView1.Rows.Count
        If U = 0 Then
            MsgBox("NO HAY DATOS PARA ELIMINAR")
        Else
            Dim respuesta As DialogResult
            Dim I18, VAL As Integer
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

            End If
        End If

    End Sub
End Class