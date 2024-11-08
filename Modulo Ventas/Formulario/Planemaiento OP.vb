Imports System.Data.SqlClient
Public Class Planemaiento_OP
    Public conx As SqlConnection
    Public comando As SqlCommand
    Private Sub buscar()

        Try
            Dim ds As New DataSet

            ds.Tables.Add(da.Copy)


            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(3).Width = 350
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim da, da1 As New DataTable
    Private Sub Planemaiento_OP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PENDIENTE()
        RadioButton2.Checked = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            planeamiento.TextBox1.Text = DataGridView1.Rows(num1).Cells(1).Value
            planeamiento.Label9.Text = Label2.Text
            planeamiento.ShowDialog()

        End If
    End Sub
    Sub PENDIENTE()
        abrir()
        da.Clear()
        Dim cmd As New SqlDataAdapter("SELECT ncom_3 as OP,program_3 as PO, descri_3  as DESCRIPCION,cants_3 as SOLICITADO,cantp_3 AS PROGRAMADO, m.usuario AS VENDEDOR, fcom_3 AS F_REGISTRO,FCome1_3 as F_DESPACHO FROM custom_vianny.dbo.qag0300 q
	 LEFT JOIN usuario_modulo M ON q.merchan_3 = m.merchan_3
	 WHERE fcom_3 >='20220601' and cantp_3 =0  and ccia ='01' AND flag_3=1  and  q.merchan_3  in ('0001','0003','0009','0025','0032','0033','0023') AND ncom_3 <= '0300000000' order by ncom_3", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(3).Width = 350
        Label3.Text = 1
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If RadioButton2.Checked = True Then
            PENDIENTE()
        Else
            If RadioButton1.Checked = True Then
                PROGRAMADO()
            End If
        End If



    End Sub
    Sub PROGRAMADO()
        abrir()
        da.Clear()
        Dim cmd As New SqlDataAdapter("SELECT ncom_3 as OP,program_3 as PO, descri_3  as DESCRIPCION,cants_3 as SOLICITADO,cantp_3 AS PROGRAMADO, m.usuario AS VENDEDOR, fcom_3 AS F_REGISTRO,FCome1_3 as F_DESPACHO FROM custom_vianny.dbo.qag0300 q
	 LEFT JOIN usuario_modulo M ON q.merchan_3 = m.merchan_3
	 WHERE fcom_3 >='20220601' and cantp_3 >0  and ccia ='01' AND flag_3=1  and  q.merchan_3  in ('0001','0003','0009','0025','0032','0033','0023') AND ncom_3 <= '0300000000' order by ncom_3", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(3).Width = 350
        Label3.Text = 2
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class