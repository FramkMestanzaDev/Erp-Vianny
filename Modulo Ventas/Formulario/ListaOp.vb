Imports System.Data.SqlClient
Public Class ListaOp
    Dim da As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim Rsr2, Rsr3, Rsr212 As SqlDataReader
    Private Sub ListaOp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        da.Rows.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT ncom_3 as OP,C.nomb_10 AS CLIENTE,descri_3 AS PRODUCTO, estilo_3 AS ESTILO,program_3 AS PM
        FROM custom_vianny.dbo.qag0300 Q
        LEFT JOIN custom_vianny.DBO.cag1000 C ON Q.ccia = C.ccia AND Q.fich_3 = C.fich_10
        where ncom_3 like '01%' and flag_3 ='1' order by ncom_3 desc", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(0).Width = 90
            DataGridView1.Columns(1).Width = 150
            DataGridView1.Columns(2).Width = 150
            DataGridView1.Columns(3).Width = 140
        End If
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

    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then


            With DataGridView1
                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If


                Op_Manufactura.TextBox1.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(0).Value)
                    Op_Manufactura.TextBox1.Focus()

                    SendKeys.Send("{ENTER}")

            End With
            Me.Close()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar0()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
    End Sub

    Private Sub buscar0()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 90
                DataGridView1.Columns(1).Width = 150
                DataGridView1.Columns(2).Width = 150
                DataGridView1.Columns(3).Width = 140
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
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 90
                DataGridView1.Columns(1).Width = 150
                DataGridView1.Columns(2).Width = 150
                DataGridView1.Columns(3).Width = 140
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
End Class