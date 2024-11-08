Imports System.Data.SqlClient
Public Class Articulos_stock
    Dim GK, GK1 As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Private Sub Articulos_stock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim TG As New FSTOCK_PAR
        Dim HJ As New VSTOCK_PAR
        Dim K As Integer
        Try


            'Label5.Text = Nsalida.Label10.Text
            HJ.gCCIA = Label6.Text
            'HJ.gCCIA = "01"
            'HJ.gALMACEN = "61"
            HJ.gALMACEN = Label3.Text
            HJ.gMODAL = 2
            'HJ.gano = "2023"
            HJ.gano = Label5.Text
            If Label3.Text = "59" Then
                GK = TG.CaeSoft_ReporteStockFisico59(HJ)
            Else
                GK = TG.CaeSoft_ReporteStockFisico(HJ)
            End If

            DataGridView1.DataSource = GK

            DataGridView1.Columns(1).Width = 142
            DataGridView1.Columns(2).Width = 330
            DataGridView1.Columns(3).Width = 80
            DataGridView1.Columns(5).Width = 85
            DataGridView1.Columns(6).Width = 65
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(8).Width = 65
            DataGridView1.Columns(9).Width = 65
            DataGridView1.Columns(10).Width = 65
            DataGridView1.Columns(10).Visible = False
            K = DataGridView1.Rows.Count
            For i = 0 To K - 1

                If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                    DataGridView1.Rows(i).Cells(11).Value = "SI"
                    DataGridView1.Columns(11).ReadOnly = True
                End If
                If DataGridView1.Rows(i).Cells(10).Value = 1 Then
                    DataGridView1.Rows(i).Cells(11).Value = "NO"
                    DataGridView1.Columns(11).ReadOnly = True
                End If
                If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                    DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                    DataGridView1.Columns(11).ReadOnly = True
                End If
            Next
            If Label7.Text = "2" Then
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(14).Visible = False
            End If
        Catch ex As Exception
            MsgBox("NO EXISTE ESTOCK EN ESTE PERIODO")
        End Try

    End Sub
    Private Sub buscar()
        Try

            Dim ds As New DataSet
            Dim K As Integer
            ds.Tables.Add(GK.Copy)

            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then
                Dim jk As Integer
                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count

                DataGridView1.Columns(1).Width = 142
                DataGridView1.Columns(2).Width = 330
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(5).Width = 85
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Width = 65
                DataGridView1.Columns(9).Width = 65
                DataGridView1.Columns(10).Width = 65
                DataGridView1.Columns(10).Visible = False
                K = DataGridView1.Rows.Count
                For i = 0 To K - 1

                    If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                        DataGridView1.Rows(i).Cells(11).Value = "SI"
                        DataGridView1.Columns(11).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(9).Value = 1 Then
                        DataGridView1.Rows(i).Cells(11).Value = "NO"
                        DataGridView1.Columns(11).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                        DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                        DataGridView1.Columns(11).ReadOnly = True
                    End If

                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_PARTIDA()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(GK.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count

                DataGridView1.Columns(1).Width = 142
                DataGridView1.Columns(2).Width = 330
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(5).Width = 85
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Width = 65
                DataGridView1.Columns(9).Width = 65
                DataGridView1.Columns(10).Width = 65
                DataGridView1.Columns(10).Visible = False
                K = DataGridView1.Rows.Count
                For i = 0 To K - 1

                    If DataGridView1.Rows(i).Cells(10).Value = 2 Then
                        DataGridView1.Rows(i).Cells(11).Value = "SI"
                        DataGridView1.Columns(11).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(9).Value = 1 Then
                        DataGridView1.Rows(i).Cells(11).Value = "NO"
                        DataGridView1.Columns(11).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(10).Value = 0 Then
                        DataGridView1.Rows(i).Cells(11).Value = "X CONF."
                        DataGridView1.Columns(11).ReadOnly = True
                    End If


                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar_PARTIDA()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
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
    Dim Rsr22, Rsr2, Rsr21, Rsr212, Rsr217, Rsr218, Rsr2177, Rsr2178, Rsr300129 As SqlDataReader



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Label4.Text = 1 Then
            Dim i, I2, SUMA As Integer
            Dim stock As Double
            Dim TG1 As New FSTOCK_PAR
            Dim HJ1 As New VSTOCK_PAR
            i = DataGridView1.Rows.Count
            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then

                    Guia_remision.DataGridView1.Rows.Add(1)
                    I2 = Guia_remision.DataGridView1.Rows.Count

                    If I2 = 1 Then
                        HJ1.gpo = DataGridView1.Rows(a).Cells(3).Value
                        HJ1.gCCIA = Label6.Text
                        Guia_remision.DataGridView1.Rows(0).Cells(0).Value = 1
                        abrir()
                        Dim sql102177 As String = "select op_3p from custom_vianny.dbo.Ranp300 where regis_3p ='" + Microsoft.VisualBasic.Mid(Trim(DataGridView1.Rows(a).Cells(3).Value), 1, 6) + "'"
                        Dim cmd102177 As New SqlCommand(sql102177, conx)
                        Rsr2177 = cmd102177.ExecuteReader()
                        If Rsr2177.Read() = True Then
                            Guia_remision.DataGridView1.Rows(0).Cells(1).Value = Rsr2177(0)
                        Else
                            Guia_remision.DataGridView1.Rows(0).Cells(1).Value = ""
                        End If
                        Rsr2177.Close()

                        Guia_remision.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(a).Cells(1).Value
                        Guia_remision.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(a).Cells(2).Value
                        Guia_remision.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(a).Cells(4).Value
                        Guia_remision.DataGridView1.Rows(0).Cells(10).Value = DataGridView1.Rows(a).Cells(3).Value
                        Guia_remision.DataGridView1.Rows(0).Cells(11).Value = DataGridView1.Rows(a).Cells(5).Value
                        Guia_remision.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(a).Cells(3).Value
                        'Guia_remision.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(a).Cells(5).Value
                        ' buscar stock almacen ficticio

                        Dim sql1022139 As String = "select SUM(CanAfi) from Almacen_Ficticio where CodAfi='" + Trim(DataGridView1.Rows(a).Cells(1).Value) + "' and AlmAfi='" + Trim(Label3.Text) + "' and EmpAfi='" + Trim(Label6.Text) + "' and ParAfi ='" + Trim(DataGridView1.Rows(a).Cells(3).Value) + "'  group by CodAfi,AlmAfi,EmpAfi,ParAfi"
                        Dim cmd1022139 As New SqlCommand(sql1022139, conx)
                        Rsr300129 = cmd1022139.ExecuteReader()
                        Dim jo As Double
                        If Rsr300129.Read() Then
                            jo = Convert.ToDouble(Rsr300129(0))
                        Else
                            jo = 0
                        End If

                        Rsr300129.Close()
                        stock = Convert.ToDouble(DataGridView1.Rows(a).Cells(5).Value) - jo
                        Guia_remision.DataGridView1.Rows(0).Cells(7).Value = stock.ToString
                        Guia_remision.DataGridView1.Rows(0).Cells(11).Value = stock
                        ' fin almacen stock
                        Guia_remision.DataGridView1.Rows(0).Cells(8).Value = DataGridView1.Rows(a).Cells(6).Value
                        Guia_remision.DataGridView1.Rows(0).Cells(5).Value = "RLL"
                        Guia_remision.DataGridView1.Rows(0).Cells(22).Value = Trim(DataGridView1.Rows(a).Cells(15).Value)
                        abrir()
                        Dim sql102 As String = "SELECT Q1.nomb_10 FROM custom_vianny.dbo.qag0300 Q LEFT JOIN custom_vianny.dbo.cag1000 Q1 ON Q.fich_3=Q1.fich_10 AND Q.ccia=Q1.ccia  WHERE program_3 ='" + Trim(TG1.mostrar_op(HJ1)) + "'"
                        Dim cmd102 As New SqlCommand(sql102, conx)
                        Rsr2 = cmd102.ExecuteReader()
                        Dim i5 As Integer
                        i5 = 0
                        If Rsr2.Read() = True Then
                            Guia_remision.DataGridView1.Rows(0).Cells(13).Value = Rsr2(0)
                        Else
                            Guia_remision.DataGridView1.Rows(0).Cells(5).Value = ""
                        End If
                        Rsr2.Close()
                        Dim sql1021 As String = "select distinct ordtej_3r from custom_vianny.dbo.marg0001 where partida_3r ='" + Trim(DataGridView1.Rows(a).Cells(3).Value) + "'"
                        Dim cmd1021 As New SqlCommand(sql1021, conx)
                        Rsr21 = cmd1021.ExecuteReader()
                        If Rsr21.Read() = True Then
                            Guia_remision.DataGridView1.Rows(0).Cells(9).Value = Rsr21(0)
                        Else
                            Guia_remision.DataGridView1.Rows(0).Cells(9).Value = ""
                        End If
                        Rsr21.Close()
                        'Insertar Almacen Ficticio
                        Dim cmd16 As New SqlCommand("insert into Almacen_Ficticio (ItmAfi,CodAfi,AlmAfi,EmpAfi,GuiAfi,PerAfi,CanAfi,ParAfi) values (@ItmAfi,@CodAfi,@AlmAfi,@EmpAfi,@GuiAfi,@PerAfi,@CanAfi,@ParAfi)", conx)
                        cmd16.Parameters.AddWithValue("@ItmAfi", "1")
                        cmd16.Parameters.AddWithValue("@CodAfi", Trim(DataGridView1.Rows(a).Cells(1).Value))
                        cmd16.Parameters.AddWithValue("@AlmAfi", Trim(Label3.Text))
                        cmd16.Parameters.AddWithValue("@EmpAfi", Trim(Label6.Text))
                        cmd16.Parameters.AddWithValue("@GuiAfi", Trim(Label8.Text))
                        cmd16.Parameters.AddWithValue("@PerAfi", Trim(Label5.Text))
                        cmd16.Parameters.AddWithValue("@CanAfi", Convert.ToDouble(DataGridView1.Rows(a).Cells(5).Value))
                        cmd16.Parameters.AddWithValue("@ParAfi", Trim(DataGridView1.Rows(a).Cells(3).Value))
                        cmd16.ExecuteNonQuery()
                    Else
                        If I2 > 1 Then
                            HJ1.gpo = DataGridView1.Rows(a).Cells(3).Value
                            HJ1.gCCIA = Label6.Text
                            SUMA = Guia_remision.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(0).Value = SUMA
                            abrir()
                            Dim sql1021778 As String = "select op_3p from custom_vianny.dbo.Ranp300 where regis_3p ='" + Microsoft.VisualBasic.Mid(Trim(DataGridView1.Rows(a).Cells(3).Value), 1, 6) + "'"
                            Dim cmd1021778 As New SqlCommand(sql1021778, conx)
                            Rsr2178 = cmd1021778.ExecuteReader()
                            If Rsr2178.Read() = True Then
                                Guia_remision.DataGridView1.Rows(0).Cells(1).Value = Rsr2178(0)
                            Else
                                Guia_remision.DataGridView1.Rows(0).Cells(1).Value = ""
                            End If
                            Rsr2178.Close()
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(a).Cells(1).Value
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(a).Cells(2).Value
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(10).Value = DataGridView1.Rows(a).Cells(3).Value
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(11).Value = DataGridView1.Rows(a).Cells(5).Value
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(a).Cells(3).Value

                            'buscar stock almacen ficticio

                            Dim sql1022139 As String = "select SUM(CanAfi) from Almacen_Ficticio where CodAfi='" + Trim(DataGridView1.Rows(a).Cells(1).Value) + "' and AlmAfi='" + Trim(Label3.Text) + "' and EmpAfi='" + Trim(Label6.Text) + "' and ParAfi ='" + Trim(DataGridView1.Rows(a).Cells(3).Value) + "'  group by CodAfi,AlmAfi,EmpAfi,ParAfi"
                            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
                            Rsr300129 = cmd1022139.ExecuteReader()
                            Dim jo As Double
                            If Rsr300129.Read() Then
                                jo = Convert.ToDouble(Rsr300129(0))
                            Else
                                jo = 0
                            End If

                            Rsr300129.Close()

                            stock = Convert.ToDouble(DataGridView1.Rows(a).Cells(5).Value) - jo
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(7).Value = stock.ToString
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(11).Value = stock
                            'fin Almacen stock
                            'Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(a).Cells(5).Value
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(8).Value = DataGridView1.Rows(a).Cells(6).Value
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(a).Cells(4).Value
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "RLL"
                            Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(22).Value = Trim(DataGridView1.Rows(a).Cells(15).Value)
                            abrir()
                            Dim sql1022 As String = "SELECT Q1.nomb_10 FROM custom_vianny.dbo.qag0300 Q LEFT JOIN custom_vianny.dbo.cag1000 Q1 ON Q.fich_3=Q1.fich_10 AND Q.ccia=Q1.ccia  WHERE program_3 ='" + Trim(TG1.mostrar_op(HJ1)) + "'"
                            Dim cmd1022 As New SqlCommand(sql1022, conx)
                            Rsr22 = cmd1022.ExecuteReader()
                            Dim i5 As Integer
                            i5 = 0
                            If Rsr22.Read() = True Then
                                Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(13).Value = Rsr22(0)
                            Else
                                Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(5).Value = ""
                            End If
                            Rsr22.Close()
                            Dim sql10212 As String = "select distinct ordtej_3r from custom_vianny.dbo.marg0001 where partida_3r ='" + Trim(DataGridView1.Rows(a).Cells(3).Value) + "'"
                            Dim cmd10212 As New SqlCommand(sql10212, conx)
                            Rsr212 = cmd10212.ExecuteReader()
                            If Rsr212.Read() = True Then
                                Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(9).Value = Rsr212(0)
                            Else
                                Guia_remision.DataGridView1.Rows(SUMA - 1).Cells(9).Value = ""
                            End If
                            Rsr212.Close()
                            'Insertar Almacen Ficticio
                            Dim cmd16 As New SqlCommand("insert into Almacen_Ficticio (ItmAfi,CodAfi,AlmAfi,EmpAfi,GuiAfi,PerAfi,CanAfi,ParAfi) values (@ItmAfi,@CodAfi,@AlmAfi,@EmpAfi,@GuiAfi,@PerAfi,@CanAfi,@ParAfi)", conx)
                            cmd16.Parameters.AddWithValue("@ItmAfi", SUMA)
                            cmd16.Parameters.AddWithValue("@CodAfi", Trim(DataGridView1.Rows(a).Cells(1).Value))
                            cmd16.Parameters.AddWithValue("@AlmAfi", Trim(Label3.Text))
                            cmd16.Parameters.AddWithValue("@EmpAfi", Trim(Label6.Text))
                            cmd16.Parameters.AddWithValue("@GuiAfi", Trim(Label8.Text))
                            cmd16.Parameters.AddWithValue("@PerAfi", Trim(Label5.Text))
                            cmd16.Parameters.AddWithValue("@CanAfi", stock)
                            cmd16.Parameters.AddWithValue("@ParAfi", Trim(DataGridView1.Rows(a).Cells(3).Value))
                            cmd16.ExecuteNonQuery()
                        End If
                    End If
                    'DataGridView1.Rows(a).ReadOnly = True
                    'Me.DataGridView1.Rows(a).Cells(0).Value = False
                    'DataGridView1.Rows(a).DefaultCellStyle.BackColor = Color.Red
                End If
            Next
            Guia_remision.DataGridView1.Columns(0).Frozen = True
            Guia_remision.DataGridView1.Columns(1).Frozen = True
            Guia_remision.DataGridView1.Columns(2).Frozen = True


        End If
        If Label4.Text = 2 Then
            Dim i, I2, SUMA As Integer
            Dim suma2 As String
            Dim stock1 As Double
            suma2 = ""
            i = DataGridView1.Rows.Count
            For b = 0 To i - 1
                If Me.DataGridView1.Rows(b).Cells(0).Value = True Then
                    Nsalida.DataGridView1.Rows.Add(1)
                    I2 = Nsalida.DataGridView1.Rows.Count
                    If I2 = 1 Then
                        Nsalida.DataGridView1.Rows(0).Cells(0).Value = "001"
                        abrir()
                        Dim sql10217 As String = "select op_3p from custom_vianny.dbo.Ranp300 where regis_3p ='" + Microsoft.VisualBasic.Mid(Trim(DataGridView1.Rows(b).Cells(3).Value), 1, 6) + "'"
                        Dim cmd10217 As New SqlCommand(sql10217, conx)
                        Rsr217 = cmd10217.ExecuteReader()
                        If Rsr217.Read() = True Then
                            Nsalida.DataGridView1.Rows(0).Cells(1).Value = Rsr217(0)
                        Else
                            Nsalida.DataGridView1.Rows(0).Cells(1).Value = ""
                        End If
                        Rsr217.Close()
                        ' buscar stock almacen ficticio

                        Dim sql1022139 As String = "select SUM(CanAfi) from Almacen_Ficticio where CodAfi='" + Trim(DataGridView1.Rows(0).Cells(1).Value) + "' and AlmAfi='" + Trim(Label3.Text) + "' and EmpAfi='" + Trim(Label6.Text) + "' and ParAfi ='" + Trim(DataGridView1.Rows(0).Cells(3).Value) + "'  group by CodAfi,AlmAfi,EmpAfi,ParAfi"
                        Dim cmd1022139 As New SqlCommand(sql1022139, conx)
                        Rsr300129 = cmd1022139.ExecuteReader()
                        Dim jo As Double
                        If Rsr300129.Read() Then
                            jo = Convert.ToDouble(Rsr300129(0))
                        Else
                            jo = 0
                        End If

                        Rsr300129.Close()
                        stock1 = Convert.ToDouble(DataGridView1.Rows(0).Cells(5).Value) - jo
                        Nsalida.DataGridView1.Rows(0).Cells(5).Value = stock1.ToString
                        Nsalida.DataGridView1.Rows(0).Cells(7).Value = stock1
                        ' fin almacen stock
                        Nsalida.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                        Nsalida.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                        Nsalida.DataGridView1.Rows(0).Cells(5).Value = DataGridView1.Rows(b).Cells(5).Value
                        Nsalida.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(b).Cells(3).Value
                        Nsalida.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(b).Cells(5).Value
                        Nsalida.DataGridView1.Rows(0).Cells(8).Value = DataGridView1.Rows(b).Cells(6).Value
                        Nsalida.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        Nsalida.DataGridView1.Rows(0).Cells(10).Value = Trim(DataGridView1.Rows(b).Cells(15).Value)
                        'Insertar Almacen Ficticio
                        Dim cmd16 As New SqlCommand("insert into Almacen_Ficticio (ItmAfi,CodAfi,AlmAfi,EmpAfi,GuiAfi,PerAfi,CanAfi,ParAfi) values (@ItmAfi,@CodAfi,@AlmAfi,@EmpAfi,@GuiAfi,@PerAfi,@CanAfi,@ParAfi)", conx)
                        cmd16.Parameters.AddWithValue("@ItmAfi", "1")
                        cmd16.Parameters.AddWithValue("@CodAfi", Trim(DataGridView1.Rows(0).Cells(1).Value))
                        cmd16.Parameters.AddWithValue("@AlmAfi", Trim(Label3.Text))
                        cmd16.Parameters.AddWithValue("@EmpAfi", Trim(Label6.Text))
                        cmd16.Parameters.AddWithValue("@GuiAfi", Trim(Label8.Text))
                        cmd16.Parameters.AddWithValue("@PerAfi", Trim(Label5.Text))
                        cmd16.Parameters.AddWithValue("@CanAfi", Convert.ToDouble(DataGridView1.Rows(0).Cells(5).Value))
                        cmd16.Parameters.AddWithValue("@ParAfi", Trim(DataGridView1.Rows(0).Cells(3).Value))
                        cmd16.ExecuteNonQuery()

                    Else
                        If I2 > 1 Then
                            SUMA = Nsalida.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                            Select Case SUMA.ToString.Length
                                Case "1" : suma2 = "00" & "" & SUMA
                                Case "2" : suma2 = "0" & "" & SUMA
                                Case "3" : suma2 = SUMA
                            End Select
                            Nsalida.DataGridView1.Rows(SUMA - 1).Cells(0).Value = suma2
                            abrir()
                            Dim sql102178 As String = "select op_3p from custom_vianny.dbo.Ranp300 where regis_3p ='" + Microsoft.VisualBasic.Mid(Trim(DataGridView1.Rows(b).Cells(3).Value), 1, 6) + "'"
                            Dim cmd102178 As New SqlCommand(sql102178, conx)
                            Rsr218 = cmd102178.ExecuteReader()
                            If Rsr218.Read() = True Then
                                Nsalida.DataGridView1.Rows(SUMA - 1).Cells(1).Value = Rsr218(0)
                            Else
                                Nsalida.DataGridView1.Rows(SUMA - 1).Cells(1).Value = ""
                            End If
                            Rsr218.Close()
                            'buscar stock almacen ficticio

                            Dim sql1022139 As String = "select SUM(CanAfi) from Almacen_Ficticio where CodAfi='" + Trim(DataGridView1.Rows(b).Cells(1).Value) + "' and AlmAfi='" + Trim(Label3.Text) + "' and EmpAfi='" + Trim(Label6.Text) + "' and ParAfi ='" + Trim(DataGridView1.Rows(b).Cells(3).Value) + "'  group by CodAfi,AlmAfi,EmpAfi,ParAfi"
                            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
                            Rsr300129 = cmd1022139.ExecuteReader()
                            Dim jo As Double
                            If Rsr300129.Read() Then
                                jo = Convert.ToDouble(Rsr300129(0))
                            Else
                                jo = 0
                            End If

                            Rsr300129.Close()
                            'MsgBox(jo.ToString)
                            'MsgBox(DataGridView1.Rows(b).Cells(5).Value)
                            stock1 = Convert.ToDouble(DataGridView1.Rows(b).Cells(5).Value) - jo

                            'fin Almacen stock
                            Nsalida.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                            Nsalida.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                            'Nsalida.DataGridView1.Rows(SUMA - 1).Cells(5).Value = DataGridView1.Rows(b).Cells(5).Value
                            Nsalida.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(b).Cells(3).Value
                            'Nsalida.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(b).Cells(5).Value
                            Nsalida.DataGridView1.Rows(SUMA - 1).Cells(8).Value = DataGridView1.Rows(b).Cells(6).Value
                            Nsalida.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                            Nsalida.DataGridView1.Rows(SUMA - 1).Cells(5).Value = stock1
                            Nsalida.DataGridView1.Rows(SUMA - 1).Cells(7).Value = stock1
                            Nsalida.DataGridView1.Rows(SUMA - 1).Cells(10).Value = Trim(DataGridView1.Rows(b).Cells(15).Value)
                            'almacen ficticio
                            Dim cmd16 As New SqlCommand("insert into Almacen_Ficticio (ItmAfi,CodAfi,AlmAfi,EmpAfi,GuiAfi,PerAfi,CanAfi,ParAfi) values (@ItmAfi,@CodAfi,@AlmAfi,@EmpAfi,@GuiAfi,@PerAfi,@CanAfi,@ParAfi)", conx)
                            cmd16.Parameters.AddWithValue("@ItmAfi", SUMA)
                            cmd16.Parameters.AddWithValue("@CodAfi", Trim(DataGridView1.Rows(b).Cells(1).Value))
                            cmd16.Parameters.AddWithValue("@AlmAfi", Trim(Label3.Text))
                            cmd16.Parameters.AddWithValue("@EmpAfi", Trim(Label6.Text))
                            cmd16.Parameters.AddWithValue("@GuiAfi", Trim(Label8.Text))
                            cmd16.Parameters.AddWithValue("@PerAfi", Trim(Label5.Text))
                            cmd16.Parameters.AddWithValue("@CanAfi", stock1)
                            cmd16.Parameters.AddWithValue("@ParAfi", Trim(DataGridView1.Rows(b).Cells(3).Value))
                            cmd16.ExecuteNonQuery()
                        End If
                    End If
                    'DataGridView1.Rows(b).ReadOnly = True
                    'Me.DataGridView1.Rows(b).Cells(0).Value = True
                    'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.Red

                End If

            Next
            Nsalida.DataGridView1.Columns(0).Frozen = True
            Nsalida.DataGridView1.Columns(1).Frozen = True
            Nsalida.DataGridView1.Columns(2).Frozen = True


        End If

        If Label4.Text = 20 Then
            Dim i, I2, SUMA As Integer
            Dim suma2 As String
            suma2 = ""
            i = DataGridView1.Rows.Count
            For b = 0 To i - 1
                If Me.DataGridView1.Rows(b).Cells(0).Value = True Then
                    NotaIngreso.DataGridView1.Rows.Add(1)
                    I2 = NotaIngreso.DataGridView1.Rows.Count
                    If I2 = 1 Then
                        NotaIngreso.DataGridView1.Rows(0).Cells(0).Value = "001"
                        NotaIngreso.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                        NotaIngreso.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                        NotaIngreso.DataGridView1.Rows(0).Cells(5).Value = DataGridView1.Rows(b).Cells(5).Value
                        NotaIngreso.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(b).Cells(6).Value
                        NotaIngreso.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(b).Cells(3).Value

                        NotaIngreso.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        NotaIngreso.DataGridView1.Rows(0).Cells(11).Value = DataGridView1.Rows(b).Cells(15).Value

                    Else
                        If I2 > 1 Then
                            SUMA = NotaIngreso.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                            Select Case SUMA.ToString.Length
                                Case "1" : suma2 = "00" & "" & SUMA
                                Case "2" : suma2 = "0" & "" & SUMA
                                Case "3" : suma2 = SUMA
                            End Select
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(0).Value = suma2
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(5).Value = DataGridView1.Rows(b).Cells(5).Value
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(b).Cells(6).Value
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(b).Cells(3).Value

                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                            NotaIngreso.DataGridView1.Rows(SUMA - 1).Cells(11).Value = DataGridView1.Rows(b).Cells(15).Value
                        End If
                    End If
                    'DataGridView1.Rows(b).ReadOnly = True
                    'Me.DataGridView1.Rows(b).Cells(0).Value = True
                    'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.Red

                End If

            Next
            NotaIngreso.DataGridView1.Columns(0).Frozen = True
            NotaIngreso.DataGridView1.Columns(1).Frozen = True
            NotaIngreso.DataGridView1.Columns(2).Frozen = True


        End If
        If Label4.Text = 200 Then
            Dim i, I2, SUMA As Integer
            Dim suma2 As String
            suma2 = ""
            i = DataGridView1.Rows.Count
            For b = 0 To i - 1
                If Me.DataGridView1.Rows(b).Cells(0).Value = True Then
                    NiaHilo.DataGridView1.Rows.Add(1)
                    I2 = NiaHilo.DataGridView1.Rows.Count
                    If I2 = 1 Then
                        NiaHilo.DataGridView1.Rows(0).Cells(0).Value = "001"
                        NiaHilo.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                        NiaHilo.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                        NiaHilo.DataGridView1.Rows(0).Cells(5).Value = DataGridView1.Rows(b).Cells(5).Value
                        NiaHilo.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(b).Cells(6).Value
                        NiaHilo.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(b).Cells(3).Value

                        NiaHilo.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        'DataGridView1.Rows(b).ReadOnly = True
                        'Me.DataGridView1.Rows(b).Cells(0).Value = True
                        'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.BurlyWood
                    Else
                        If I2 > 1 Then
                            SUMA = NiaHilo.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                            Select Case SUMA.ToString.Length
                                Case "1" : suma2 = "00" & "" & SUMA
                                Case "2" : suma2 = "0" & "" & SUMA
                                Case "3" : suma2 = SUMA
                            End Select
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(0).Value = suma2
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = DataGridView1.Rows(b).Cells(5).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(b).Cells(6).Value
                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(b).Cells(3).Value

                            NiaHilo.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                        End If
                    End If
                    'DataGridView1.Rows(b).ReadOnly = True
                    'Me.DataGridView1.Rows(b).Cells(0).Value = True
                    'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.Red

                End If

            Next
            NiaHilo.DataGridView1.Columns(0).Frozen = True
            NiaHilo.DataGridView1.Columns(1).Frozen = True
            NiaHilo.DataGridView1.Columns(2).Frozen = True
        Else
            If Label4.Text = 400 Then
                Dim i, I2, SUMA As Integer
                Dim suma2 As String
                suma2 = ""
                i = DataGridView1.Rows.Count
                For b = 0 To i - 1
                    If Me.DataGridView1.Rows(b).Cells(0).Value = True Then
                        NsaHilo.DataGridView1.Rows.Add(1)
                        I2 = NsaHilo.DataGridView1.Rows.Count
                        If I2 = 1 Then
                            NsaHilo.DataGridView1.Rows(0).Cells(0).Value = "001"
                            NsaHilo.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                            NsaHilo.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                            Select Case Label3.Text
                                Case "41" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "KGS"
                                Case "03" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "CON"
                                Case "42" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "CON"
                                Case "43" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                                Case "44" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "UND"
                                Case "59" : NsaHilo.DataGridView1.Rows(0).Cells(5).Value = "KGS"
                            End Select
                            NsaHilo.DataGridView1.Rows(0).Cells(6).Value = DataGridView1.Rows(b).Cells(5).Value
                            NsaHilo.DataGridView1.Rows(0).Cells(10).Value = DataGridView1.Rows(b).Cells(5).Value
                            If Label3.Text = "59" Then
                                NsaHilo.DataGridView1.Rows(0).Cells(7).Value = DataGridView1.Rows(b).Cells(3).Value
                            Else
                                NsaHilo.DataGridView1.Rows(0).Cells(7).Value = ""
                            End If
                            NsaHilo.DataGridView1.Rows(0).Cells(8).Value = DataGridView1.Rows(b).Cells(3).Value
                            NsaHilo.DataGridView1.Rows(0).Cells(11).Value = DataGridView1.Rows(b).Cells(6).Value
                            NsaHilo.DataGridView1.Rows(0).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                            'DataGridView1.Rows(b).ReadOnly = True
                            'Me.DataGridView1.Rows(b).Cells(0).Value = True
                            'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.BurlyWood
                        Else
                            If I2 > 1 Then
                                SUMA = NsaHilo.DataGridView1.Rows(I2 - 2).Cells(0).Value + 1
                                Select Case SUMA.ToString.Length
                                    Case "1" : suma2 = "00" & "" & SUMA
                                    Case "2" : suma2 = "0" & "" & SUMA
                                    Case "3" : suma2 = SUMA
                                End Select
                                NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(0).Value = suma2
                                NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(2).Value = DataGridView1.Rows(b).Cells(1).Value
                                NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(3).Value = DataGridView1.Rows(b).Cells(2).Value
                                Select Case Label3.Text
                                    Case "41" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "KGS"
                                    Case "03" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "CON"
                                    Case "42" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "CON"
                                    Case "43" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                                    Case "44" : NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(5).Value = "UND"
                                End Select

                                NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(6).Value = DataGridView1.Rows(b).Cells(5).Value
                                NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(10).Value = DataGridView1.Rows(b).Cells(5).Value
                                If Label3.Text = "59" Then
                                    NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(7).Value = DataGridView1.Rows(b).Cells(3).Value
                                Else
                                    NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(7).Value = ""
                                End If
                                NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(8).Value = DataGridView1.Rows(b).Cells(3).Value
                                NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(11).Value = DataGridView1.Rows(b).Cells(6).Value
                                NsaHilo.DataGridView1.Rows(SUMA - 1).Cells(4).Value = DataGridView1.Rows(b).Cells(4).Value
                            End If
                        End If
                        'DataGridView1.Rows(b).ReadOnly = True
                        'Me.DataGridView1.Rows(b).Cells(0).Value = True
                        'DataGridView1.Rows(b).DefaultCellStyle.BackColor = Color.Red

                    End If

                Next
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim J As Integer

        J = DataGridView1.Rows.Count
        If CheckBox1.Checked = True Then
            For i = 0 To J - 1
                DataGridView1.Rows(i).Cells(0).Value = True
            Next
        Else
            For i = 0 To J - 1
                DataGridView1.Rows(i).Cells(0).Value = False
            Next
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'If e.ColumnIndex = Me.DataGridView1.Columns.Item("S").Index Then

        'End If
        'If e.ColumnIndex = Me.DataGridView1.Columns.Item("S").Index Then
        '    Dim value As Boolean = CType(Me.DataGridView1.CurrentCell.EditedFormattedValue, Boolean)
        '    If value = True Then
        '        'MessageBox.Show("CheckBox Marcado")
        '        Me.DataGridView1.Rows(e.RowIndex).Cells(7).Value = 1
        '        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Red
        '    Else
        '        'MessageBox.Show("CheckBox Desmarcado")
        '        Me.DataGridView1.Rows(e.RowIndex).Cells(7).Value = 2
        '        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
        '    End If
        'End If
    End Sub
End Class