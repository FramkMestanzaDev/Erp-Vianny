Imports System.Data.SqlClient
Public Class buscar_requerimitno
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
    Dim da, DA1 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        da.Clear()
        Dim cmd As New SqlDataAdapter("SELECT nu_requerimientio AS REQUERIMIENTO,cliente AS CLIENTE , po as PO, tipo as TIPO,VENDEDOR AS VENDEDOR, case when estado=0 then 'PENDIENTE' ELSE 'PROGRAMADO' END AS ESTADO FROM cab_req_comercial WHERE YEAR(fecha) ='" + Trim(ComboBox1.Text) + "' And estado ='0'", conx)



        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(1).Width = 115
            DataGridView1.Columns(2).Width = 310
            DataGridView1.Columns(1).Width = 110
        End If
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 0 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString

            REQ_C.TextBox1.Text = DataGridView1.Rows(num1).Cells(1).Value
            REQ_C.Show()


        End If
    End Sub
    Private Sub buscar()
        Try

            Dim ds As New DataSet
            Dim K As Integer
            ds.Tables.Add(da.Copy)

            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "REQUERIMIENTO" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then
                Dim jk As Integer
                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 125
                DataGridView1.Columns(2).Width = 350
                DataGridView1.Columns(1).Width = 125
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_requerimitno_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        DA1.Clear()
        Dim cmd As New SqlDataAdapter("SELECT nu_requerimientio AS REQUERIMIENTO,cliente AS CLIENTE , po as PO, tipo as TIPO,VENDEDOR AS VENDEDOR, case when estado=0 then 'PENDIENTE' ELSE 'PROGRAMADO' END AS ESTADO FROM cab_req_comercial WHERE YEAR(fecha) ='" + Trim(ComboBox1.Text) + "' And estado ='1'", conx)



        cmd.Fill(da1)
        If da1.Rows.Count > 0 Then
            DataGridView1.DataSource = da1
            DataGridView1.Columns(1).Width = 115
            DataGridView1.Columns(2).Width = 310
            DataGridView1.Columns(1).Width = 110
        End If
    End Sub
End Class