Imports System.Data.SqlClient
Public Class Seguimiento_reclamo
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da As New DataTable
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
    Dim Rsr2 As SqlDataReader
    Private Sub Seguimiento_reclamo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton2.Checked = True
        abrir()
        Dim cmd As New SqlDataAdapter("select numero AS NUMERO,ruc AS RUC,cliente AS CLIENTE,nota_pedido AS NOTA_PEDIDO,ap_am AS CONTACTO,lugar_visita AS LUGAR_VISITA, telefono AS TELEFONO, fecha AS FECHA,m.usuario AS VENDEDOR, '' AS CALIDAD from Hoja_Reclamo_tela t  left join usuario_modulo m on t.vendedor = m.id_usuario", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(1).Width = 95
        DataGridView1.Columns(2).Width = 95
        DataGridView1.Columns(3).Width = 200
        DataGridView1.Columns(4).Width = 95
        DataGridView1.Columns(5).Width = 200
        DataGridView1.Columns(6).Width = 200
        DataGridView1.Columns(7).Width = 80
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(8).Width = 80
        Dim va As Integer
        va = DataGridView1.Rows.Count

        For lp = 0 To va - 1
            abrir()
            Dim sql102 As String = "select COUNT(id_ca_rt) from calidad_reclamo_tela where numero ='" + DataGridView1.Rows(lp).Cells(1).Value + "'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            If Rsr2.Read() = True Then
                If Rsr2(0) >= 1 Then
                    DataGridView1.Rows(lp).Cells(10).Value = "REVISADO"
                    DataGridView1.Rows(lp).DefaultCellStyle.BackColor = Color.Yellow
                    DataGridView1.Rows(lp).DefaultCellStyle.ForeColor = Color.DarkRed
                Else
                    DataGridView1.Rows(lp).Cells(10).Value = "PENDIENTE"
                End If

            End If

            Rsr2.Close()
        Next

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count
                DataGridView1.DataSource = da
                DataGridView1.Columns(1).Width = 95
                DataGridView1.Columns(2).Width = 95
                DataGridView1.Columns(3).Width = 200
                DataGridView1.Columns(4).Width = 95
                DataGridView1.Columns(5).Width = 200
                DataGridView1.Columns(6).Width = 200
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Width = 80
                Dim va As Integer
                va = DataGridView1.Rows.Count

                For lp = 0 To va - 1
                    If Trim(DataGridView1.Rows(lp).Cells(10).Value) = "REVISADO" Then
                        DataGridView1.Rows(lp).DefaultCellStyle.BackColor = Color.Yellow
                        DataGridView1.Rows(lp).DefaultCellStyle.ForeColor = Color.DarkRed
                    End If
                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_VENDEDOR()
        Try
            Dim ds As New DataSet
            Dim K As Integer


            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "VENDEDOR" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count
                DataGridView1.DataSource = da
                DataGridView1.Columns(1).Width = 95
                DataGridView1.Columns(2).Width = 95
                DataGridView1.Columns(3).Width = 200
                DataGridView1.Columns(4).Width = 95
                DataGridView1.Columns(5).Width = 200
                DataGridView1.Columns(6).Width = 200
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Width = 80
                Dim va As Integer
                va = DataGridView1.Rows.Count

                For lp = 0 To va - 1
                    If Trim(DataGridView1.Rows(lp).Cells(10).Value) = "REVISADO" Then
                        DataGridView1.Rows(lp).DefaultCellStyle.BackColor = Color.Yellow
                        DataGridView1.Rows(lp).DefaultCellStyle.ForeColor = Color.DarkRed
                    End If
                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 0 Then
            Rpt_Reclamo.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            Rpt_Reclamo.Show()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar_VENDEDOR()
    End Sub
End Class