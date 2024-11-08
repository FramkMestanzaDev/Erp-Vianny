Imports System.Data.SqlClient
Public Class ListaOD
    Dim da As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim Rsr2, Rsr3, Rsr212 As SqlDataReader

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

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
    Private Sub buscar0()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "OD" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 90
                DataGridView1.Columns(1).Width = 65
                DataGridView1.Columns(2).Width = 150
                DataGridView1.Columns(3).Width = 150
                DataGridView1.Columns(4).Width = 185
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
                DataGridView1.Columns(1).Width = 65
                DataGridView1.Columns(2).Width = 150
                DataGridView1.Columns(3).Width = 150
                DataGridView1.Columns(4).Width = 185
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscarPm()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PM" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 90
                DataGridView1.Columns(1).Width = 65
                DataGridView1.Columns(2).Width = 150
                DataGridView1.Columns(3).Width = 150
                DataGridView1.Columns(4).Width = 185
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar0()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscarPm()
    End Sub

    Private Sub ListaOD_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        da.Rows.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT SUBSTRING(ncom_3,1,9) as OD,SUBSTRING(ncom_3,10,1) AS VERSION,C.nomb_10 AS CLIENTE,descri_3 AS PRODUCTO, estilo_3 AS ESTILO,program_3 AS PM,telaprinc_3
        FROM custom_vianny.dbo.qag0300 Q
        LEFT JOIN custom_vianny.DBO.cag1000 C ON Q.ccia = C.ccia AND Q.fich_3 = C.fich_10
        where ncom_3 like '03%' and flag_3 ='1' order by ncom_3 desc", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(0).Width = 90
            DataGridView1.Columns(1).Width = 65
            DataGridView1.Columns(2).Width = 150
            DataGridView1.Columns(3).Width = 150
            DataGridView1.Columns(4).Width = 140
            DataGridView1.Columns(6).Visible = False
        End If



    End Sub

    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then


            With DataGridView1
                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If

                If Label4.Text = 2 Then
                    OD.TextBox22.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(0).Value)
                    OD.TextBox23.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)
                    OD.TextBox23.Focus()

                    SendKeys.Send("{ENTER}")

                Else
                    If Label4.Text = 1 Then
                        Detalle_clonar.TextBox1.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(0).Value) & Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)
                        SendKeys.Send("{ENTER}")
                    Else
                        If Label4.Text = 3 Then
                            MatrizAviosOd.TextBox9.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(0).Value)
                            MatrizAviosOd.TextBox11.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)
                            MatrizAviosOd.TextBox11.Focus()
                            SendKeys.Send("{ENTER}")
                        Else
                            If Label4.Text = 4 Then

                                MatrizAviosOd.TextBox10.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(0).Value) & Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)
                                MatrizAviosOd.TextBox10.Focus()
                                SendKeys.Send("{ENTER}")
                            Else
                                If Label4.Text = 5 Then
                                    ReqTelaProduccion.TextBox1.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(0).Value) & Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)
                                    ReqTelaProduccion.TextBox2.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(5).Value)
                                    ReqTelaProduccion.TextBox3.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(3).Value)
                                    ReqTelaProduccion.TextBox6.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(6).Value)
                                    ReqTelaProduccion.TextBox1.Enabled = False
                                    ReqTelaProduccion.ComboBox1.SelectedIndex = 2
                                Else
                                    If Label4.Text = 6 Then
                                        RequerimientoAvios.TextBox1.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(0).Value) & Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)
                                        RequerimientoAvios.TextBox2.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(5).Value)
                                        RequerimientoAvios.TextBox3.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(3).Value)
                                        RequerimientoAvios.TextBox7.Enabled = True
                                        RequerimientoAvios.TextBox1.Enabled = False
                                        RequerimientoAvios.TextBox4.Enabled = False
                                        RequerimientoAvios.Label11.Visible = True
                                        RequerimientoAvios.Label13.Visible = True
                                        RequerimientoAvios.TextBox6.Visible = True
                                        RequerimientoAvios.Label12.Visible = True
                                        RequerimientoAvios.TextBox7.Visible = True
                                        RequerimientoAvios.ComboBox1.SelectedIndex = 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End With
            Me.Close()
        End If
    End Sub
End Class