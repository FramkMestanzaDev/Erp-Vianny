Imports System.Data.SqlClient
Public Class PARTIDA_RAMA
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
    Private Sub PARTIDA_RAMA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("select DISTINCT ncom_3n ,regis_3n AS PARTIDA,r.cnomb_3c as COLOR,c.nomb_17 AS PRODUCTO,V.rollospesados AS ROLLOS,V.kgsobtenidos AS KILOS, m.lote_3r  from custom_vianny.dbo.qan0300 Q INNER JOIN  custom_vianny.dbo.qanp300 QA ON q.ccia_3n = qa.ccia_3p and q.regis_3n = qa.regis_3p INNER JOIN validar_partida V ON Q.regis_3n = V.partida
left join custom_vianny.dbo.cag1700 c on qa.ccia_3p = c.ccia and qa.linea_3p = c.linea_17 left join custom_vianny.dbo.Marg0001 m on q.ccia_3n = m.ccia_3r and q.regis_3n = m.partida_3r left join custom_vianny.dbo.Qarc0300 r  on  q.ccia_3n = r.ccia_3c and q.color_3n = r.ccolor_3c
where   q.ccia_3n ='01' AND V.estado ='1' 
", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(4).Width = 450
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar3()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(1).Visible = False
                DataGridView1.Columns(4).Width = 450
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, I2, SUMA As Integer
        Dim valor As Double
        i = DataGridView1.Rows.Count
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                rama.DataGridView1.Rows.Add(1)
                I2 = rama.DataGridView1.Rows.Count

                If I2 = 1 Then

                    rama.DataGridView1.Rows(0).Cells(0).Value = 1
                    rama.DataGridView1.Rows(0).Cells(1).Value = DataGridView1.Rows(0).Cells(1).Value
                    rama.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(0).Cells(2).Value
                    rama.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(0).Cells(3).Value
                    rama.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(0).Cells(4).Value
                    rama.DataGridView1.Rows(0).Cells(5).Value = DataGridView1.Rows(0).Cells(5).Value
                    rama.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(0).Cells(6).Value
                    rama.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(0).Cells(7).Value
                    rama.DataGridView1.Rows(0).Cells(10).Value = "-"
                    rama.DataGridView1.Rows(0).Cells(11).Value = 0
                    rama.DataGridView1.Rows(0).Cells(12).Value = "-"
                    rama.DataGridView1.Rows(0).Cells(13).Value = "0.00"
                    rama.DataGridView1.Rows(0).Cells(14).Value = 0
                    rama.DataGridView1.Rows(0).Cells(15).Value = "0.00"
                    rama.DataGridView1.Rows(0).Cells(16).Value = 0

                Else

                    If I2 > 1 Then

                        SUMA = rama.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                        rama.DataGridView1.Rows(SUMA - 1).Cells(0).Value = SUMA
                        rama.DataGridView1.Rows(SUMA - 1).Cells(1).Value = DataGridView1.Rows(a).Cells(1).Value
                        rama.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(a).Cells(2).Value
                        rama.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(a).Cells(3).Value
                        rama.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(a).Cells(4).Value
                        rama.DataGridView1.Rows(SUMA - 1).Cells(5).Value = DataGridView1.Rows(a).Cells(5).Value
                        rama.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(a).Cells(6).Value
                        rama.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(a).Cells(7).Value
                        rama.DataGridView1.Rows(SUMA - 1).Cells(10).Value = "-"
                        rama.DataGridView1.Rows(SUMA - 1).Cells(11).Value = 0
                        rama.DataGridView1.Rows(SUMA - 1).Cells(12).Value = "-"
                        rama.DataGridView1.Rows(SUMA - 1).Cells(13).Value = "0.00"
                        rama.DataGridView1.Rows(SUMA - 1).Cells(14).Value = 0
                        rama.DataGridView1.Rows(SUMA - 1).Cells(15).Value = "0.00"
                        rama.DataGridView1.Rows(SUMA - 1).Cells(16).Value = 0

                    End If
                End If
                valor = valor + DataGridView1.Rows(a).Cells(6).Value

            End If
        Next
        rama.TextBox2.Text = valor + rama.TextBox2.Text
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim j As Integer

        j = DataGridView1.Rows.Count
        If CheckBox1.Checked = True Then
            For i = 0 To j - 1
                Me.DataGridView1.Rows(i).Cells(0).Value = True
            Next
        Else
            For i = 0 To j - 1
                Me.DataGridView1.Rows(i).Cells(0).Value = False
            Next
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar3()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar2()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class