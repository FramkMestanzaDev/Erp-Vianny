Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class det_ns_Prodc
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Dim Rsr2, Rsr20, Rsr21, Rsr22 As SqlDataReader

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar2()
    End Sub

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

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar3()
    End Sub

    Dim da As New DataTable
    Private Sub det_ns_Prodc_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        If Label3.Text = 1 Then

            Dim cmd As New SqlDataAdapter("SELECT dest_21e AS CODIGO,nomb_21e AS NOMBRE FROM custom_vianny.DBO.Qag210e A  Where A.empr_21e = '" + Label5.Text + "' AND A.almac_21e = '" + Trim(Label2.Text) + "'", conx)
            cmd.Fill(da)
            DataGridView1.DataSource = da
            DataGridView1.Columns(0).Width = 100
            DataGridView1.Columns(1).Width = 300
        Else
            If Label3.Text = 2 Then

                abrir()
                Dim cmd As New SqlDataAdapter("select ruc_10 AS RUC,nomb_10 as NOMBRE from custom_vianny.dbo.cag1000 where ccia ='01'", conx)
                cmd.Fill(da)
                DataGridView1.DataSource = da
                DataGridView1.Columns(0).Width = 100
                DataGridView1.Columns(1).Width = 300
                Label1.Text = "CLIENTE"
                Label6.Visible = True
                TextBox2.Enabled = True
            Else
                If Label3.Text = 3 Then

                    abrir()
                    Dim cmd As New SqlDataAdapter("SELECT A.dest_21s AS CODIGO,a.nomb_21s AS NOMBRE FROM custom_vianny.DBO.Qag210s A  Where a.empr_21s='" + Label5.Text + "' AND almac_21s='" + Trim(Label2.Text) + "'", conx)
                    cmd.Fill(da)
                    DataGridView1.DataSource = da
                    DataGridView1.Columns(0).Width = 100
                    DataGridView1.Columns(1).Width = 300
                Else
                    If Label3.Text = 4 Then

                        abrir()
                        Dim cmd As New SqlDataAdapter("select ruc_10 AS RUC,nomb_10 as NOMBRE from custom_vianny.dbo.cag1000 where ccia ='01' and tdeta_10 ='PR'", conx)
                        cmd.Fill(da)
                        DataGridView1.DataSource = da
                        DataGridView1.Columns(0).Width = 100
                        DataGridView1.Columns(1).Width = 300
                        Label6.Visible = True
                        TextBox2.Enabled = True
                    Else
                        If Label3.Text = 5 Then

                            abrir()
                            Dim cmd As New SqlDataAdapter("SELECT A.dest_21s as CODIGO,a.nomb_21s AS NOMBRE FROM custom_vianny.DBO.Qag210s A  Where a.empr_21s='" + Label5.Text + "' AND almac_21s='" + Trim(Label2.Text) + "'", conx)
                            cmd.Fill(da)
                            DataGridView1.DataSource = da
                            DataGridView1.Columns(0).Width = 100
                            DataGridView1.Columns(1).Width = 300
                        Else
                            If Label3.Text = 6 Then

                                abrir()
                                Dim cmd As New SqlDataAdapter("select ruc_10 AS RUC,nomb_10 AS NOMBRE,direcc_10  from custom_vianny.dbo.cag1000 where ccia ='01'  ", conx)
                                cmd.Fill(da)
                                DataGridView1.DataSource = da
                                DataGridView1.Columns(0).Width = 100
                                DataGridView1.Columns(1).Width = 300
                                DataGridView1.Columns(2).Visible = False
                                Label1.Text = "CLIENTES"
                                Label6.Visible = True
                                TextBox2.Enabled = True
                            Else
                                If Label3.Text = 8 Then


                                    abrir()
                                    Dim cmd As New SqlDataAdapter("select ruc_10 AS RUC,nomb_10 as NOMBRE from custom_vianny.dbo.cag1000 where ccia ='" + Label5.Text + "' and tdeta_10 ='PR' ", conx)
                                    cmd.Fill(da)
                                    DataGridView1.DataSource = da
                                    DataGridView1.Columns(0).Width = 100
                                    DataGridView1.Columns(1).Width = 300
                                    Label6.Visible = True
                                    TextBox2.Enabled = True
                                Else
                                    If Label3.Text = 4111 Then


                                        abrir()
                                        Dim cmd As New SqlDataAdapter("select ruc_10 AS RUC,nomb_10 as NOMBRE,direcc_10 as DIRECCION from custom_vianny.dbo.cag1000 where ccia ='" + Label5.Text + "' and tdeta_10 ='PR' ", conx)
                                        cmd.Fill(da)
                                        DataGridView1.DataSource = da
                                        DataGridView1.Columns(0).Width = 100
                                        Label6.Visible = True
                                        TextBox2.Enabled = True
                                        DataGridView1.Columns(1).Width = 300
                                    Else
                                        If Label3.Text = 9 Then
                                            abrir()
                                            Dim cmd As New SqlDataAdapter("select ruc_10 AS RUC,nomb_10 as NOMBRE from custom_vianny.dbo.cag1000 where ccia ='" + Label5.Text + "' and tdeta_10 ='PR'", conx)
                                            cmd.Fill(da)
                                            DataGridView1.DataSource = da
                                            DataGridView1.Columns(0).Width = 100
                                            DataGridView1.Columns(1).Width = 300
                                            Label1.Text = "CLIENTE"
                                            Label6.Visible = True
                                            TextBox2.Enabled = True
                                        Else
                                            If Label3.Text = 11 Then
                                                abrir()
                                                Dim cmd As New SqlDataAdapter("select ruc_10 AS RUC,nomb_10 as NOMBRE from custom_vianny.dbo.cag1000 where ccia ='" + Label5.Text + "' and tdeta_10 ='PR'", conx)
                                                cmd.Fill(da)
                                                DataGridView1.DataSource = da
                                                DataGridView1.Columns(0).Width = 100
                                                DataGridView1.Columns(1).Width = 300
                                                Label1.Text = "CLIENTE"
                                                Label6.Visible = True
                                                TextBox2.Enabled = True
                                            Else
                                                If Label3.Text = 12 Then

                                                    abrir()
                                                    Dim cmd As New SqlDataAdapter("select ruc_10 AS RUC,nomb_10 as NOMBRE from custom_vianny.dbo.cag1000 where ccia ='" + Label5.Text + "' and tdeta_10 ='PR' ", conx)
                                                    cmd.Fill(da)
                                                    DataGridView1.DataSource = da
                                                    DataGridView1.Columns(0).Width = 100
                                                    DataGridView1.Columns(1).Width = 350
                                                    Label6.Visible = True
                                                    TextBox2.Enabled = True
                                                Else
                                                    If Label3.Text = 14 Then
                                                        abrir()
                                                        Dim cmd As New SqlDataAdapter("select ruc_10 AS RUC,nomb_10 as NOMBRE from custom_vianny.dbo.cag1000 where ccia ='" + Label5.Text + "' and tdeta_10 ='PR' ", conx)
                                                        cmd.Fill(da)
                                                        DataGridView1.DataSource = da
                                                        DataGridView1.Columns(0).Width = 100
                                                        DataGridView1.Columns(1).Width = 350
                                                        Label6.Visible = True
                                                        TextBox2.Enabled = True
                                                    End If
                                                End If


                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    End If
                End If
            End If

        End If

    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "NOMBRE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 100
                DataGridView1.Columns(1).Width = 300
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


            dv.RowFilter = "RUC" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 100
                DataGridView1.Columns(1).Width = 300
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If Label3.Text = 1 Then
            Nia_Produccion.TextBox7.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            Nia_Produccion.TextBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            Nia_Produccion.TextBox8.Select()
            Me.Close()
        Else
            If Label3.Text = 2 Then
                Nia_Produccion.TextBox8.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                Nia_Produccion.TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                Nia_Produccion.TextBox15.Select()
                Me.Close()
            Else
                If Label3.Text = 3 Then
                    nsa_produccion.TextBox7.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                    nsa_produccion.TextBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                    nsa_produccion.TextBox8.Select()
                    Me.Close()
                Else
                    If Label3.Text = 4 Then
                        nsa_produccion.TextBox8.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                        nsa_produccion.TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                        nsa_produccion.TextBox15.Select()
                        Me.Close()
                    Else
                        If Label3.Text = 5 Then
                            guia_talleres.TextBox21.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                            guia_talleres.TextBox22.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                            guia_talleres.TextBox8.Select()
                            Me.Close()
                        Else
                            If Label3.Text = 6 Then
                                guia_talleres.TextBox8.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                                guia_talleres.TextBox9.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                                guia_talleres.TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
                                guia_talleres.TextBox15.Select()
                                Me.Close()
                            Else
                                If Label3.Text = 8 Then
                                    ORDEN_CORTE.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                                    ORDEN_CORTE.TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                                    ORDEN_CORTE.TextBox3.Select()
                                    Me.Close()
                                Else
                                    If Label3.Text = 4111 Then
                                        'Formato_solic_avios.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                                        'Formato_solic_avios.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                                        'Formato_solic_avios.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
                                        Me.Close()
                                    Else
                                        If Label3.Text = 9 Then
                                            Asignacion_Confeccion.Label3.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                                            Asignacion_Confeccion.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                                            Asignacion_Confeccion.TextBox2.Select()
                                            Me.Close()
                                        Else
                                            If Label3.Text = 11 Then
                                                Asignacion_Acabados.Label3.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                                                Asignacion_Acabados.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                                                Asignacion_Acabados.TextBox2.Select()
                                                Me.Close()
                                            Else

                                                If Label3.Text = 12 Then

                                                    Asignacion_Molde.TextBox14.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                                                    Asignacion_Molde.TextBox8.Select()
                                                    Me.Close()
                                                Else
                                                    If Label3.Text = 14 Then
                                                        RequerimientoAvios.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString())
                                                        RequerimientoAvios.TextBox8.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString())
                                                        Me.Close()
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        End If

    End Sub
End Class