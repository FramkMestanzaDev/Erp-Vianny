Imports System.Data.SqlClient
Public Class DET_FICHA_TECNICA
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
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
    Private Sub DET_FICHA_TECNICA_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        abrir()

        Dim j As String
        If Label3.Text = 1 Then
            If Label2.Text = "06" Then
                Dim cmd As New SqlDataAdapter("select linea_17 AS CODIGO,nomb_17 AS DESCRIPCION, unid_17 AS UM from custom_vianny.dbo.cag1700 where ccia='01' and interno2_17 IN (1,3) and talm_17 ='" + Trim(Label2.Text) + "'", conx)
                cmd.Fill(da)
            Else
                Dim cmd As New SqlDataAdapter("select linea_17 AS CODIGO,nomb_17 AS DESCRIPCION, unid_17 AS UM from custom_vianny.dbo.cag1700 where ccia='01'  and talm_17 ='" + Trim(Label2.Text) + "'", conx)
                cmd.Fill(da)
            End If


            DataGridView1.DataSource = da
                DataGridView1.Columns(0).Width = 170
                DataGridView1.Columns(1).Width = 380
            Else
                If Label3.Text = 2 Then
                Dim cmd1 As New SqlDataAdapter("select linea_17 AS CODIGO,nomb_17 AS DESCRIPCION, unid_17 AS UM from custom_vianny.dbo.cag1700 where ccia='01' and talm_17 IN (03,42)", conx)
                cmd1.Fill(da)
                DataGridView1.DataSource = da
                DataGridView1.Columns(0).Width = 170
                DataGridView1.Columns(1).Width = 380
            Else
                If Label3.Text = 3 Then
                    Label1.Text = "OP"
                    Dim cmd1 As New SqlDataAdapter("select ncom_3 AS OP, ISNULL( c.nomb_10,'') AS CLIENTE,descri_3 AS PRODUCTO,estilo_3 AS ESTILO from custom_vianny.dbo.qag0300 q left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 = c.fich_10 where q.ccia ='01'  and ncom_3 like '01%' and  (merchan_3 <> '0009' or  merchan_3 <> '0026')", conx)
                    cmd1.Fill(da)
                    DataGridView1.DataSource = da
                    DataGridView1.Columns(0).Width = 100
                    DataGridView1.Columns(1).Width = 180
                    DataGridView1.Columns(2).Width = 200
                    DataGridView1.Columns(3).Width = 170
                Else
                    If Label3.Text = 4 Then
                        Label1.Text = "OP"
                        Dim cmd1 As New SqlDataAdapter("select ncom_3 AS OP, ISNULL( c.nomb_10,'') AS CLIENTE,descri_3 AS PRODUCTO,estilo_3 AS ESTILO from custom_vianny.dbo.qag0300 q left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 = c.fich_10 where q.ccia ='01'  and ncom_3 like '01%' and  (merchan_3 <> '0009' or  merchan_3 <> '0026')", conx)
                        cmd1.Fill(da)
                        DataGridView1.DataSource = da
                        DataGridView1.Columns(0).Width = 100
                        DataGridView1.Columns(1).Width = 180
                        DataGridView1.Columns(2).Width = 200
                        DataGridView1.Columns(3).Width = 170
                    Else
                        If Label3.Text = 5 Then
                            Dim cmd1 As New SqlDataAdapter("select linea_17 AS CODIGO,nomb_17 AS DESCRIPCION, unid_17 AS UM from custom_vianny.dbo.cag1700 where ccia='01' and talm_17 IN (06,08)", conx)
                            cmd1.Fill(da)
                            DataGridView1.DataSource = da
                            DataGridView1.Columns(0).Width = 170
                            DataGridView1.Columns(1).Width = 380
                        Else
                            If Label3.Text = 6 Then
                                Label1.Text = "OP"
                                Dim cmd1 As New SqlDataAdapter("select ncom_3 AS OP, ISNULL( c.nomb_10,'') AS CLIENTE,descri_3 AS PRODUCTO,estilo_3 AS ESTILO from custom_vianny.dbo.qag0300 q left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 = c.fich_10 where q.ccia ='01' and year(fcom_3) >= '2021' and ncom_3 like '01%' and  (merchan_3 <> '0009' or  merchan_3 <> '0026') and  (broker_3 ='0001' or broker_3 ='0000') ", conx)
                                cmd1.Fill(da)
                                DataGridView1.DataSource = da
                                DataGridView1.Columns(0).Width = 100
                                DataGridView1.Columns(1).Width = 180
                                DataGridView1.Columns(2).Width = 200
                                DataGridView1.Columns(3).Width = 170
                            Else

                                If Label3.Text = 7 Then
                                    Label1.Text = "OP"
                                    Dim cmd1 As New SqlDataAdapter("select ncom_3 AS OP, ISNULL( c.nomb_10,'') AS CLIENTE,descri_3 AS PRODUCTO,estilo_3 AS ESTILO, program_3 AS PO from custom_vianny.dbo.qag0300 q left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 = c.fich_10 where q.ccia ='01' and year(fcom_3) >= '2021' and ncom_3 like '01%' and  (merchan_3 <> '0009' or  merchan_3 <> '0026') and  (broker_3 ='0001' or broker_3 ='0000') ", conx)
                                    cmd1.Fill(da)
                                    DataGridView1.DataSource = da
                                    DataGridView1.Columns(0).Width = 100
                                    DataGridView1.Columns(1).Width = 180
                                    DataGridView1.Columns(2).Width = 200
                                    DataGridView1.Columns(3).Width = 170
                                    DataGridView1.Columns(4).Visible = False
                                Else
                                    If Label3.Text = 8 Then
                                        Label1.Text = "PO"
                                        Dim cmd1 As New SqlDataAdapter("select DISTINCT program_3 AS PO, ISNULL( c.nomb_10,'') AS CLIENTE from custom_vianny.dbo.qag0300 q left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 = c.fich_10 where q.ccia ='01' and q.flag_3 ='1' and year(fcom_3) >='2023' and ncom_3 like '01%' and  (merchan_3 not in ( '0009',  '0026') )   group by program_3,c.nomb_10", conx)
                                        cmd1.Fill(da)
                                        DataGridView1.DataSource = da
                                        DataGridView1.Columns(0).Width = 150
                                        DataGridView1.Columns(1).Width = 400

                                    End If
                                End If
                            End If
                        End If
                    End If

                End If

            End If


        End If

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer
            If (Label3.Text = 3 Or Label3.Text = 4 Or Label3.Text = 6 Or Label3.Text = 7) Then
                dv.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"
            Else
                If (Label3.Text = 8) Then
                    dv.RowFilter = "PO" & " like '%" & TextBox1.Text & "%'"
                Else
                    dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox1.Text & "%'"
                End If

            End If


            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count
                DataGridView1.Columns(0).Width = 170
                DataGridView1.Columns(1).Width = 380
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer


            dv.RowFilter = "CODIGO" & " like '%" & TextBox2.Text & "%'"






            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count
                DataGridView1.Columns(0).Width = 170
                DataGridView1.Columns(1).Width = 380
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

    End Sub

    Dim Rs1, Rs11 As SqlDataReader
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then

        Else
            If Label3.Text = 1 Then
            FICHA_TECNICA.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
            FICHA_TECNICA.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                FICHA_TECNICA.TextBox1.Focus()
                SendKeys.Send("{ENTER}")
                FICHA_TECNICA.TextBox3.Focus()
                Me.Close()

        Else
                If Label3.Text = 2 Then
                    FICHA_TECNICA.DataGridView1.Rows(Label4.Text).Cells(2).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                    FICHA_TECNICA.DataGridView1.Rows(Label4.Text).Cells(3).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                    FICHA_TECNICA.TextBox3.Select()
                    Me.Close()
                Else
                    If Label3.Text = 3 Then
                        FICHA_TECNICA.DataGridView2.Rows(Label4.Text).Cells(1).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                        FICHA_TECNICA.DataGridView2.Rows(Label4.Text).Cells(2).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                        FICHA_TECNICA.DataGridView2.Rows(Label4.Text).Cells(3).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
                        FICHA_TECNICA.DataGridView2.Rows(Label4.Text).Cells(4).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)

                        Me.Close()
                    Else
                        If Label3.Text = 4 Then
                            'Asignacion_op_tela.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                            Me.Close()
                        Else
                            If Label3.Text = 5 Then
                                'Asignacion_op_tela.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                                'Asignacion_op_tela.TextBox1.Focus()
                                'SendKeys.Send("{ENTER}")
                                'Asignacion_op_tela.TextBox2.Select()
                                'Me.Close()
                            Else
                                If Label3.Text = 6 Then
                                    Asignacion_Molde.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                                    Me.Close()
                                Else
                                    If Label3.Text = 7 Then
                                        Ficha_udp.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                                        Ficha_udp.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
                                        Ficha_udp.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
                                        Ficha_udp.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)


                                        Dim sql1 As String = "SELECT count(op) as cant FROM Ficha_diseno where op='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value) + "'"
                                        Dim cmd1 As New SqlCommand(sql1, conx)
                                        Rs1 = cmd1.ExecuteReader()
                                        If Rs1.Read() = True Then
                                            If Rs1(0) = 1 Then
                                                Rs1.Close()
                                                Dim sql11 As String = "SELECT ruta,trabajador,f_costuta,f_avios,f_acabados,observacion,c.nomb_10,ruta_medida FROM Ficha_diseno F LEFT JOIN custom_vianny.DBO.cag1000 c on f.trabajador =c.ruc_10 where op='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value) + "'"
                                                Dim cmd11 As New SqlCommand(sql11, conx)
                                                Rs11 = cmd11.ExecuteReader()

                                                If Rs11.Read() = True Then
                                                    Ficha_udp.TextBox6.Text = Trim(Rs11(0))
                                                    Ficha_udp.TextBox8.Text = Trim(Rs11(1))
                                                    Ficha_udp.TextBox9.Text = Trim(Rs11(6))
                                                    Ficha_udp.TextBox7.Text = Trim(Rs11(7))
                                                    If Rs11(2) = 1 Then
                                                        Ficha_udp.CheckBox2.Checked = True
                                                    Else
                                                        Ficha_udp.CheckBox2.Checked = False
                                                    End If
                                                    If Rs11(3) = 1 Then
                                                        Ficha_udp.CheckBox3.Checked = True
                                                    Else
                                                        Ficha_udp.CheckBox3.Checked = False
                                                    End If
                                                    If Rs11(4) = 1 Then
                                                        Ficha_udp.CheckBox4.Checked = True
                                                    Else
                                                        Ficha_udp.CheckBox4.Checked = False
                                                    End If

                                                    Ficha_udp.TextBox11.Text = Trim(Rs11(5))
                                                End If
                                                Rs11.Close()
                                                Ficha_udp.Button1.Enabled = False
                                                Ficha_udp.Button2.Enabled = False
                                                Ficha_udp.Button3.Enabled = True
                                                Ficha_udp.Button5.Enabled = False
                                                Ficha_udp.TextBox8.Enabled = False
                                                Ficha_udp.TextBox11.Enabled = False
                                                Ficha_udp.CheckBox2.Enabled = False
                                                Ficha_udp.CheckBox3.Enabled = False
                                                Ficha_udp.CheckBox4.Enabled = False
                                                Ficha_udp.TextBox8.Select()
                                            Else
                                                Ficha_udp.Button5.Enabled = True
                                                Ficha_udp.Button1.Enabled = True
                                                Ficha_udp.TextBox8.Enabled = True
                                                Ficha_udp.TextBox11.Enabled = True
                                                Ficha_udp.CheckBox2.Enabled = True
                                                Ficha_udp.CheckBox3.Enabled = True
                                                Ficha_udp.CheckBox4.Enabled = True
                                                Ficha_udp.Button2.Enabled = True
                                            End If


                                        End If
                                        Rs1.Close()
                                        Me.Close()

                                    Else
                                        If Label3.Text = 8 Then
                                            Consumo_tela_Consolidado.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
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
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
    End Sub
End Class