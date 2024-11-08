Imports System.Data.SqlClient
Public Class det_tesoreria
    Public conx As SqlConnection
    Public comando As SqlCommand
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Dim da As New DataTable
    Private Sub det_tesoreria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        If Label4.Text = 0 Then
            Dim cmd As New SqlDataAdapter(" SELECT * FROM RUBROS_TESORERIA ORDER BY NATURALEZA", conx)

            cmd.Fill(da)
            DataGridView1.DataSource = da
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Width = 300
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(3).Width = 220
        Else
            If Label4.Text = 3 Then
                Dim cmd As New SqlDataAdapter(" select PLACA as PLACA, MARCA AS MARCA, TIPO AS TIPO from  VEHICULO ORDER BY PLACA", conx)

                cmd.Fill(da)
                DataGridView1.DataSource = da

                DataGridView1.Columns(0).Width = 70
                DataGridView1.Columns(1).Width = 300
                DataGridView1.Columns(2).Width = 220
            Else
                Dim cmd As New SqlDataAdapter(" SELECT * FROM RUBROS_MANTENIMIENTO ORDER BY ID", conx)

                cmd.Fill(da)
                DataGridView1.DataSource = da

                DataGridView1.Columns(0).Width = 70
                DataGridView1.Columns(1).Width = 300
                DataGridView1.Columns(2).Width = 220
            End If

        End If

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick



    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "NOMBRE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(1).Width = 300
                DataGridView1.Columns(2).Width = 70
                DataGridView1.Columns(3).Width = 220
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.RowIndex <> -1 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            If Label4.Text = 0 Then



                Data_Bancos.DataGridView1.Rows(Label1.Text).Cells(1).Value = DataGridView1.Rows(num1).Cells(1).Value
                Data_Bancos.DataGridView1.Rows(Label1.Text).Cells(2).Value = DataGridView1.Rows(num1).Cells(0).Value
                If Trim(Data_Bancos.DataGridView1.Rows(Label1.Text).Cells(19).Value) <> "" Then
                    Data_Bancos.DataGridView1.Rows(Label1.Text).Cells(19).Value = ""
                End If
                Me.Close()


            Else
                If Label4.Text = 3 Then

                    Mantenimiento_carro.TextBox1.Text = Trim(DataGridView1.Rows(num1).Cells(1).Value)
                    Mantenimiento_carro.TextBox2.Text = Trim(DataGridView1.Rows(num1).Cells(0).Value)

                    Me.Close()


                Else

                    Mantenimiento_carro.TextBox3.Text = Trim(DataGridView1.Rows(num1).Cells(1).Value)
                    Mantenimiento_carro.TextBox8.Text = Trim(DataGridView1.Rows(num1).Cells(0).Value)
                    Mantenimiento_carro.TextBox9.Text = Trim(DataGridView1.Rows(num1).Cells(3).Value)
                    Mantenimiento_carro.DateTimePicker1.Enabled = True
                    Me.Close()
                End If

            End If
        End If
    End Sub
End Class