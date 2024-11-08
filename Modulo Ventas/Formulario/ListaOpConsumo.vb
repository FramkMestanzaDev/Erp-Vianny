Imports System.Data.SqlClient
Public Class ListaOpConsumo
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub ListaOpConsumo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PENDIENTE()
        RadioButton1.Checked = True
    End Sub
    Sub PENDIENTE()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select ncom_3 as OP, program_3 AS PO,descri_3 AS PRODUCTO,estilo_3 AS ESTILO, cantp_3 AS PROGRAMADO from custom_vianny.dbo.qag0300 where ncom_3 like '01%'and ccia='01' and ISNULL( ncorte_3,0) in (0,1) order by ncom_3", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(3).Width = 380
        Else
            DataGridView1.DataSource = Nothing
        End If

    End Sub
    Sub consumo()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select ncom_3 as OP, program_3 AS PO,descri_3 AS PRODUCTO,estilo_3 AS ESTILO, cantp_3 AS PROGRAMADO from custom_vianny.dbo.qag0300 where ncom_3 like '01%'and ccia='01' and ncorte_3 ='2' and YEAR(fcom_3) = '" + Label4.Text + "' order by ncom_3 desc", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(3).Width = 380
        Else
            DataGridView1.DataSource = Nothing
        End If

    End Sub
    Private Sub buscarop()
        Try
            Dim ds As New DataSet
            Dim K As Integer
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer
            dv.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(3).Width = 380
            Else

                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub buscarPO()
        Try
            Dim ds As New DataSet
            Dim K As Integer
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer
            dv.RowFilter = "PO" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(3).Width = 380
            Else

                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscarop()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscarPO()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            PENDIENTE()
        Else
            If RadioButton2.Checked = True Then
                consumo()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "C" Then
            FICHA_CONSUMO.Label8.Text = Label3.Text
            FICHA_CONSUMO.Label13.Text = Label5.Text
            FICHA_CONSUMO.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString.Trim
            If RadioButton1.Checked = True Then
                FICHA_CONSUMO.Button4.Enabled = False
                FICHA_CONSUMO.Button3.Enabled = True
            Else
                If RadioButton2.Checked = True Then
                    FICHA_CONSUMO.Button4.Enabled = True
                    FICHA_CONSUMO.Button3.Enabled = False
                    FICHA_CONSUMO.Button1.Enabled = True
                    FICHA_CONSUMO.Button2.Enabled = True
                    FICHA_CONSUMO.Button6.Enabled = True
                End If
            End If
            SendKeys.Send("{ENTER}")
            FICHA_CONSUMO.Show(Me)
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class