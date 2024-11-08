Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Data.SqlClient

Public Class Ventas
    Dim DT, DT1, DT2, DT3, DT6 As New DataTable
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Private Sub Label9_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs)

    End Sub
    Dim FT1 As New DataTable

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs)

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

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE rpm_ven ='2'", conx)
            conn.Fill(ty2)
            ComboBox2.DataSource = ty2
            ComboBox2.DisplayMember = "alias_ven"
            ComboBox2.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Ventas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim gh As New fventas
        Dim gh1 As New vventas
        Dim ai, ai1 As New Integer
        Dim SUMA As Double
        Dim NOMBRE As String
        SUMA = 0
        If MDIParent1.Label9.Text = "02" Then
            ComboBox2.Items.Add("MCRUZADO")
            ComboBox2.Items.Add("TIENDA")
        Else
            abrir()
            llenar_combo_box2()
        End If


        ComboBox1.SelectedIndex = DateTime.Now.ToString("MM") - 1
        ComboBox2.SelectedIndex = 0
        gh1.gmes = DateTime.Now.ToString("MM")
        gh1.gANO = Label33.Text
        gh1.gccia = Label36.Text
        NOMBRE = MonthName(DateTime.Now.ToString("MM"))
        If Label36.Text = "01" Then
            DT = gh.buscar_ventas_mensuales(gh1)
            FT1 = gh.buscar_ventas_mensuales_negras(gh1)
            abrir()
            Dim cmd As New SqlDataAdapter("SELECT distinct  v.alias_ven FROM custom_vianny.DBO.Fag0300 A LEFT JOIN custom_vianny.DBO.Fap0300 B
ON A.CCia_3 = B.CCia_3a AND A.CPeriodo_3 = B.CPeriodo_3a AND A.Tienda_3 = B.Tienda_3a AND A.NCom_3 = B.NCom_3a					
left join custom_vianny.dbo.Vendedores v on a.vendedor_3 = v.codigo_ven
WHERE cperiodo_3 = '" + Label33.Text + "' AND ccia_3 ='" + Label36.Text + "'    AND flag_3 ='1' AND MONTH(fcom_3)= '" + DateTime.Now.ToString("MM") + "'  and v.rpm_ven <> '' and B.rubro_3a not in (0035,0036)
union
SELECT V.alias_ven    AS VENDEDOR FROM venta_cabecera FG  
INNER JOIN rventa_detalle FP ON  FG.cperiodo_3= FP.cperiodo_3a and fg.empresa= fp.empresa
AND FG.ncom_3 = FP.ncom_3a
LEFT JOIN custom_vianny.DBO.Vendedores V ON FG.vendedor_3 = V.codigo_ven
WHERE cperiodo_3 = '" + Label33.Text + "'  AND (tidoc_3 BETWEEN '001' AND '003')  AND flag_3 ='1' AND MONTH(fcom_3)= '" + DateTime.Now.ToString("MM") + "'
AND v.rpm_ven <> '' and fg.empresa ='" + Label36.Text + "' GROUP BY V.alias_ven", conx)

            Dim da12 As New DataTable

            cmd.Fill(da12)
            DataGridView5.DataSource = da12
        Else
            DT = gh.buscar_ventas_mensuales_GRAUS(gh1)
            FT1 = gh.buscar_ventas_mensuales_negras(gh1)
            abrir()
            Dim cmd As New SqlDataAdapter("SELECT distinct  v.alias_ven FROM custom_vianny.DBO.Fag0300 A LEFT JOIN custom_vianny.DBO.Fap0300 B
ON A.CCia_3 = B.CCia_3a AND A.CPeriodo_3 = B.CPeriodo_3a AND A.Tienda_3 = B.Tienda_3a AND A.NCom_3 = B.NCom_3a					
left join custom_vianny.dbo.Vendedores v on a.vendedor_3 = v.codigo_ven
WHERE cperiodo_3 = '" + Label33.Text + "' AND ccia_3 ='" + Label36.Text + "'    AND flag_3 ='1' AND MONTH(fcom_3)= '" + DateTime.Now.ToString("MM") + "'  and v.rpm_ven <> '' and B.rubro_3a not in (0035,0036)
union
SELECT V.alias_ven    AS VENDEDOR FROM venta_cabecera FG  
INNER JOIN rventa_detalle FP ON  FG.cperiodo_3= FP.cperiodo_3a and fg.empresa= fp.empresa
AND FG.ncom_3 = FP.ncom_3a
LEFT JOIN custom_vianny.DBO.Vendedores V ON FG.vendedor_3 = V.codigo_ven
WHERE cperiodo_3 = '" + Label33.Text + "'  AND (tidoc_3 BETWEEN '001' AND '003')  AND flag_3 ='1' AND MONTH(fcom_3)= '" + DateTime.Now.ToString("MM") + "'
AND v.rpm_ven <> '' and fg.empresa ='" + Label36.Text + "' GROUP BY V.alias_ven", conx)

            Dim da12 As New DataTable

            cmd.Fill(da12)
            DataGridView5.DataSource = da12
        End If

        Me.Chart1.Series.Clear()
        Chart1.ChartAreas.Clear()



        If DT Is Nothing Then
            MsgBox("NO SE REGISTRARON VENTAS ESTE MES")
        Else

            If DT.Rows.Count <> 0 Or FT1.Rows.Count <> 0 Then
                DataGridView2.DataSource = DT
                ai = DataGridView2.Rows.Count
                ai1 = DataGridView5.Rows.Count
                For i = 0 To ai1 - 1
                    DataGridView1.Columns.Add(DataGridView5.Rows(i).Cells(0).Value, DataGridView5.Rows(i).Cells(0).Value)
                    DataGridView1.Columns(i).HeaderText = DataGridView5.Rows(i).Cells(0).Value
                    DataGridView1.Columns(i).Name = DataGridView5.Rows(i).Cells(0).Value
                Next
                DataGridView1.Rows.Add(3)
                For o = 0 To ai1 - 1
                    For i = 0 To ai - 1
                        If DataGridView1.Columns(o).HeaderText.ToString.Trim = DataGridView2.Rows(i).Cells(0).Value.ToString.Trim Then
                            DataGridView1.Rows(0).Cells(o).Value = DataGridView2.Rows(i).Cells(1).Value
                            DataGridView1.Columns(o).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                            DataGridView1.Columns(o).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                            DataGridView1.Columns(o).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                        End If


                    Next
                Next


                Dim ad As Integer
                ad = Convert.ToInt32(DataGridView1.ColumnCount.ToString())

                For i = 0 To ad - 1
                    SUMA = SUMA + DataGridView1.Rows(0).Cells(i).Value
                Next

                TextBox1.Text = Format(SUMA, "##,##00.00")

                Dim KL, kl1, kl2 As New Integer
                DataGridView3.DataSource = FT1
                KL = DataGridView1.Rows.Count
                kl1 = DataGridView3.Rows.Count

                If KL = 0 Then

                    For l = 0 To kl1 - 1
                        DataGridView1.Columns.Add(DataGridView3.Rows(l).Cells(0).Value, DataGridView3.Rows(l).Cells(0).Value)
                        DataGridView1.Columns(l).HeaderText = DataGridView3.Rows(l).Cells(0).Value
                        DataGridView1.Columns(l).Name = DataGridView3.Rows(l).Cells(0).Value
                        DataGridView1.Rows(1).Cells(l).Value = DataGridView3.Rows(l).Cells(1).Value

                    Next

                Else
                    Dim v As Integer
                    kl2 = DataGridView2.Rows.Count

                    For o = 0 To ai1 - 1
                        For p = 0 To kl1 - 1

                            If DataGridView1.Columns(o).HeaderText.ToString.Trim = DataGridView3.Rows(p).Cells(0).Value.ToString.Trim Then

                                DataGridView1.Rows(1).Cells(o).Value = DataGridView3.Rows(p).Cells(1).Value
                            Else


                            End If
                        Next
                    Next

                    Dim op As Integer
                    op = DataGridView1.Columns.Count
                    DataGridView4.Rows.Add(op)
                    For p = 0 To op - 1
                        DataGridView4.Rows(p).Cells(0).Value = p + 1
                        DataGridView4.Rows(p).Cells(1).Value = DataGridView1.Columns(p).Name
                    Next


                    Dim AD1 As Integer
                    Dim SUMA20 As Double
                    Dim titulo As New Title(MonthName(gh1.gmes))
                    titulo.Font = New Font("Tahoma", 18, FontStyle.Bold)
                    Me.Chart1.Titles.Add(titulo)
                    Me.Chart1.Series.Add("VENTAS CF")
                    Me.Chart1.Series.Add("VENTAS SF")


                    Chart1.ChartAreas.Add("VENTAS")
                    AD1 = Convert.ToInt32(DataGridView1.ColumnCount.ToString())
                    Chart1.Series(0)("DrawingStyle") = "Cylinder"
                    For i = 0 To AD1 - 1
                        Me.Chart1.Series("VENTAS CF").Points.AddXY(i + 1, DataGridView1.Rows(0).Cells(i).Value)
                        Me.Chart1.Series("VENTAS SF").Points.AddXY(i + 1, DataGridView1.Rows(1).Cells(i).Value)
                        SUMA20 = SUMA20 + DataGridView1.Rows(1).Cells(i).Value

                    Next
                    TextBox2.Text = Format(SUMA20, "##,##00.00")
                    Dim jk, jk1 As Integer
                    Dim suma50 As Double

                    jk = DataGridView1.Rows.Count
                    jk1 = DataGridView5.Rows.Count

                    For h = 0 To jk1 - 1
                        suma50 = 0
                        For u = 0 To jk - 1
                            suma50 = suma50 + DataGridView1.Rows(u).Cells(h).Value
                        Next
                        DataGridView1.Rows(2).Cells(h).Value = suma50
                    Next
                    DataGridView1.Rows(2).DefaultCellStyle.BackColor = Color.IndianRed
                    Dim IO As Integer
                    Dim SUMA30 As Double
                    IO = Convert.ToInt32(DataGridView1.ColumnCount.ToString())

                    For W = 0 To IO - 1

                        SUMA30 = SUMA30 + DataGridView1.Rows(2).Cells(W).Value
                    Next
                    'Format(SUMA30, "##,##00.00")
                    TextBox3.Text = Format(SUMA30, "##,##00.00")

                End If
            Else
                Me.Chart1.Series.Clear()
                MsgBox("NO EXISTE INFORMACION PARA MOSTRAR")
            End If

        End If



    End Sub
    Dim Rsr2 As SqlDataReader
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox2.Text = "----SELECCIONAR----" Then
            MsgBox("NO SELECCIONO NINGUN VENDEDOR")
        Else
            Dim gh As New fventas
            Dim gh1 As New vventas
            Dim ai As New Integer
            Dim SUMA As Double
            'GroupBox3.Visible = True
            'GroupBox4.Visible = False
            SUMA = 0
            DataGridView4.Rows.Clear()
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            DataGridView2.DataSource = ""
            DataGridView3.DataSource = ""
            TextBox2.Text = ""
            TextBox3.Text = ""

            abrir()
            Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + ComboBox2.Text + "'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            If Rsr2.Read() = True Then
                gh1.gVENDEDOR = Rsr2(0)
            End If
            Rsr2.Close()
            'Select Case ComboBox2.Text
            '    Case "VPIZARRO" : gh1.gVENDEDOR = "0027"
            '    Case "VSILVERIO" : gh1.gVENDEDOR = "0005"
            '    Case "GBALVIN" : gh1.gVENDEDOR = "0028"
            '    Case "GBEDON" : gh1.gVENDEDOR = "0010"
            '    Case "VINCIO" : gh1.gVENDEDOR = "0022"
            '    Case "DBRAVO" : gh1.gVENDEDOR = "0023"
            '    Case "AMENDO" : gh1.gVENDEDOR = "0026"
            '    Case "JSALINAS" : gh1.gVENDEDOR = "0025"
            '    Case "VGRAUS" : gh1.gVENDEDOR = "0007"
            '    Case "WSALINAS" : gh1.gVENDEDOR = "0034"
            '    Case "MCRUZADO" : gh1.gVENDEDOR = "0035"
            'End Select
            gh1.gANO = Label33.Text
            gh1.gccia = Label36.Text
            DT2 = gh.buscar_ventas_VENDEDOR(gh1)
            DT6 = gh.buscar_ventas_VENDEDOR_NEGRAS(gh1)

            DataGridView2.DataSource = DT2
            DataGridView3.DataSource = DT6
            ai = DataGridView2.Rows.Count
            For i = 0 To ai - 1
                DataGridView1.Columns.Add(DataGridView2.Rows(i).Cells(0).Value, DataGridView2.Rows(i).Cells(0).Value)
                DataGridView1.Columns(i).HeaderText = DataGridView2.Rows(i).Cells(0).Value
                DataGridView1.Columns(i).Name = DataGridView2.Rows(i).Cells(0).Value
            Next
            DataGridView1.Rows.Add(3)
            For i = 0 To ai - 1
                DataGridView1.Rows(0).Cells(i).Value = DataGridView2.Rows(i).Cells(1).Value
                DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Next
            Me.Chart1.Series.Clear()
            Me.Chart1.Titles.Clear()
            Me.Chart1.Titles.Add(ComboBox2.Text)
            Me.Chart1.Series.Add("VENTAS CF")
            Me.Chart1.Series.Add("VENTAS SF")
            'Me.Chart1.Series.Add(ComboBox2.Text)
            'Me.Chart1.Series.Item(0).BorderColor = Color.DarkKhaki
            Dim ad As Integer


            ad = Convert.ToInt32(DataGridView1.ColumnCount.ToString())
            For i = 0 To ad - 1
                'Me.Chart1.Series(ComboBox2.Text).Points.AddXY(i + 1, DataGridView2.Rows(i).Cells(1).Value)

                SUMA = SUMA + DataGridView1.Rows(0).Cells(i).Value

            Next

            TextBox1.Text = Format(SUMA, "##,##00.00")
            Dim KL, kl1, kl2 As New Integer

            KL = DataGridView1.Rows.Count
            kl1 = DataGridView3.Rows.Count

            If KL = 0 Then

                For l = 0 To kl1 - 1
                    DataGridView1.Columns.Add(DataGridView3.Rows(l).Cells(0).Value, DataGridView3.Rows(l).Cells(0).Value)
                    DataGridView1.Columns(l).HeaderText = DataGridView3.Rows(l).Cells(0).Value
                    DataGridView1.Columns(l).Name = DataGridView3.Rows(l).Cells(0).Value
                    DataGridView1.Rows(1).Cells(l).Value = DataGridView3.Rows(l).Cells(1).Value
                Next
            Else


                DataGridView1.Rows(1).Cells(0).Value = DataGridView3.Rows(0).Cells(1).Value
                DataGridView1.Rows(1).Cells(1).Value = DataGridView3.Rows(1).Cells(1).Value
                DataGridView1.Rows(1).Cells(2).Value = DataGridView3.Rows(2).Cells(1).Value
                DataGridView1.Rows(1).Cells(3).Value = DataGridView3.Rows(3).Cells(1).Value
                DataGridView1.Rows(1).Cells(4).Value = DataGridView3.Rows(4).Cells(1).Value
                DataGridView1.Rows(1).Cells(5).Value = DataGridView3.Rows(5).Cells(1).Value
                DataGridView1.Rows(1).Cells(6).Value = DataGridView3.Rows(6).Cells(1).Value
                DataGridView1.Rows(1).Cells(7).Value = DataGridView3.Rows(7).Cells(1).Value
                DataGridView1.Rows(1).Cells(8).Value = DataGridView3.Rows(8).Cells(1).Value
                DataGridView1.Rows(1).Cells(9).Value = DataGridView3.Rows(9).Cells(1).Value
                DataGridView1.Rows(1).Cells(10).Value = DataGridView3.Rows(10).Cells(1).Value
                DataGridView1.Rows(1).Cells(11).Value = DataGridView3.Rows(11).Cells(1).Value



                DataGridView4.Rows.Add(12)
                For p = 0 To 11
                    DataGridView4.Rows(p).Cells(0).Value = p + 1
                Next
                DataGridView4.Rows(0).Cells(1).Value = "ENERO"
                DataGridView4.Rows(1).Cells(1).Value = "FEBRERO"
                DataGridView4.Rows(2).Cells(1).Value = "MARZO"
                DataGridView4.Rows(3).Cells(1).Value = "ABRIL"
                DataGridView4.Rows(4).Cells(1).Value = "MAYO"
                DataGridView4.Rows(5).Cells(1).Value = "JUNIO"
                DataGridView4.Rows(6).Cells(1).Value = "JULIO"
                DataGridView4.Rows(7).Cells(1).Value = "AGOSTO"
                DataGridView4.Rows(8).Cells(1).Value = "SETIEMBRE"
                DataGridView4.Rows(9).Cells(1).Value = "OCTUBRE"
                DataGridView4.Rows(10).Cells(1).Value = "NOVIEMBRE"
                DataGridView4.Rows(11).Cells(1).Value = "DICIEMBRE"


                Dim AD1 As Integer
                Dim SUMA20 As Double

                AD1 = Convert.ToInt32(DataGridView1.ColumnCount.ToString())

                For i = 0 To AD1 - 1
                    Me.Chart1.Series("VENTAS CF").Points.AddXY(i + 1, DataGridView1.Rows(0).Cells(i).Value)
                    Me.Chart1.Series("VENTAS SF").Points.AddXY(i + 1, DataGridView1.Rows(1).Cells(i).Value)
                    SUMA20 = SUMA20 + DataGridView1.Rows(1).Cells(i).Value

                Next
                TextBox2.Text = Format(SUMA20, "##,##00.00")
                Dim jk, jk1 As Integer
                Dim suma50 As Double

                jk = DataGridView1.Rows.Count
                jk1 = DataGridView2.Rows.Count

                For h = 0 To jk1 - 1
                    suma50 = 0
                    For u = 0 To jk - 1
                        suma50 = suma50 + DataGridView1.Rows(u).Cells(h).Value
                    Next
                    DataGridView1.Rows(2).Cells(h).Value = suma50
                Next
                DataGridView1.Rows(2).DefaultCellStyle.BackColor = Color.IndianRed
                Dim IO As Integer
                Dim SUMA30 As Double
                IO = Convert.ToInt32(DataGridView1.ColumnCount.ToString())

                For W = 0 To IO - 1

                    SUMA30 = SUMA30 + DataGridView1.Rows(2).Cells(W).Value
                Next
                TextBox3.Text = Format(SUMA30, "##,##00.00")

            End If

        End If

    End Sub
    Dim ft As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim gh As New fventas
        Dim gh1 As New vventas
        Dim ai As New Integer
        Dim SUMA As Double
        'GroupBox3.Visible = False
        'GroupBox4.Visible = True
        SUMA = 0
        DataGridView1.Rows.Clear()
        DataGridView4.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView2.DataSource = ""
        DataGridView3.DataSource = ""

        Select Case ComboBox1.Text
            Case "ENERO" : gh1.gmes = "01"
            Case "FEBRERO" : gh1.gmes = "02"
            Case "MARZO" : gh1.gmes = "03"
            Case "ABRIL" : gh1.gmes = "04"
            Case "MAYO" : gh1.gmes = "05"
            Case "JUNIO" : gh1.gmes = "06"
            Case "JULIO" : gh1.gmes = "07"
            Case "AGOSTO" : gh1.gmes = "08"
            Case "SETIEMBRE" : gh1.gmes = "09"
            Case "OCTUBRE" : gh1.gmes = "10"
            Case "NOVIEMBRE" : gh1.gmes = "11"
            Case "DICIEMBRE" : gh1.gmes = "12"
        End Select
        Me.Chart1.Titles.Clear()
        Me.Chart1.Series.Clear()
        Chart1.ChartAreas.Clear()

        Dim titulo As New Title(MonthName(gh1.gmes))
        titulo.Font = New Font("Tahoma", 18, FontStyle.Bold)
        Me.Chart1.Titles.Add(titulo)
        Me.Chart1.Series.Add("VENTAS CF")
        Me.Chart1.Series.Add("VENTAS SF")
        Chart1.ChartAreas.Add("VENTAS")
        'Me.Chart1.Palette = ChartColorPalette.Excel
        gh1.gANO = Label33.Text
        gh1.gccia = Label36.Text
        If Label36.Text = "01" Then
            DT1 = gh.buscar_ventas_mensuales(gh1)
            ft = gh.buscar_ventas_mensuales_negras(gh1)
        Else
            DT1 = gh.buscar_ventas_mensuales_GRAUS(gh1)
            ft = gh.buscar_ventas_mensuales_negras(gh1)
        End If
        'DT1 = gh.buscar_ventas_mensuales(gh1)
        'ft = gh.buscar_ventas_mensuales_negras(gh1)
        If DT1 Is Nothing Then
            MsgBox("NO SE REGISTRARON VENTAS ESTE MES")
        Else
            If DT1.Rows.Count <> 0 Or ft.Rows.Count <> 0 Then
                DataGridView2.DataSource = DT1
                ai = DataGridView2.Rows.Count
                For i = 0 To ai - 1
                    DataGridView1.Columns.Add(DataGridView2.Rows(i).Cells(0).Value, DataGridView2.Rows(i).Cells(0).Value)
                    DataGridView1.Columns(i).HeaderText = DataGridView2.Rows(i).Cells(0).Value
                    DataGridView1.Columns(i).Name = DataGridView2.Rows(i).Cells(0).Value
                Next
                DataGridView1.Rows.Add(3)
                For i = 0 To ai - 1
                    DataGridView1.Rows(0).Cells(i).Value = DataGridView2.Rows(i).Cells(1).Value
                    DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                Next
                'Me.Chart1.Series.Clear()
                'Me.Chart1.Series.Add(gh1.gmes)
                'Me.Chart1.Series.Item(0).BorderColor = Color.DarkKhaki
                Dim ad As Integer
                ad = Convert.ToInt32(DataGridView1.ColumnCount.ToString())
                For i = 0 To ad - 1
                    'Me.Chart1.Series(gh1.gmes).Points.AddXY(DataGridView2.Rows(i).Cells(0).Value, DataGridView2.Rows(i).Cells(1).Value)
                    SUMA = SUMA + DataGridView1.Rows(0).Cells(i).Value
                Next

                TextBox1.Text = Format(SUMA, "##,##00.00")




                Dim KL, kl1, kl2 As New Integer
                DataGridView3.DataSource = ft
                KL = DataGridView1.Rows.Count
                kl1 = DataGridView3.Rows.Count

                If KL = 0 Then

                    For l = 0 To kl1 - 1
                        DataGridView1.Columns.Add(DataGridView3.Rows(l).Cells(0).Value, DataGridView3.Rows(l).Cells(0).Value)
                        DataGridView1.Columns(l).HeaderText = DataGridView3.Rows(l).Cells(0).Value
                        DataGridView1.Columns(l).Name = DataGridView3.Rows(l).Cells(0).Value
                        DataGridView1.Rows(1).Cells(l).Value = DataGridView3.Rows(l).Cells(1).Value
                    Next

                Else

                    kl2 = DataGridView2.Rows.Count
                    For o = 0 To kl2 - 1


                        For p = 0 To kl1 - 1
                            If DataGridView1.Columns(o).HeaderText.ToString.Trim = DataGridView3.Rows(p).Cells(0).Value.ToString.Trim Then

                                DataGridView1.Rows(1).Cells(o).Value = DataGridView3.Rows(p).Cells(1).Value

                            End If
                        Next
                    Next

                    Dim op As Integer
                    op = DataGridView1.Columns.Count
                    DataGridView4.Rows.Add(op)
                    For p = 0 To op - 1
                        DataGridView4.Rows(p).Cells(0).Value = p + 1
                        DataGridView4.Rows(p).Cells(1).Value = DataGridView1.Columns(p).Name
                    Next


                    Dim AD1 As Integer
                    Dim SUMA20 As Double

                    AD1 = Convert.ToInt32(DataGridView1.ColumnCount.ToString())
                    'Chart1.Series(0).ChartType = SeriesChartType.Column
                    Chart1.Series(0)("DrawingStyle") = "Cylinder"
                    'Chart1.ChartAreas(0).AxisX.LabelStyle.Angle = -90

                    For i = 0 To AD1 - 1
                        Me.Chart1.Series("VENTAS CF").Points.AddXY(i + 1, DataGridView1.Rows(0).Cells(i).Value)
                        Me.Chart1.Series("VENTAS SF").Points.AddXY(i + 1, DataGridView1.Rows(1).Cells(i).Value)
                        SUMA20 = SUMA20 + DataGridView1.Rows(1).Cells(i).Value
                    Next

                    'Chart1.ChartAreas(0).AxisX.Interval = 200
                    TextBox2.Text = Format(SUMA20, "##,##00.00")
                    Dim jk, jk1 As Integer
                    Dim suma50 As Double

                    jk = DataGridView1.Rows.Count
                    jk1 = DataGridView2.Rows.Count

                    For h = 0 To jk1 - 1
                        suma50 = 0
                        For u = 0 To jk - 1
                            suma50 = suma50 + DataGridView1.Rows(u).Cells(h).Value
                        Next
                        DataGridView1.Rows(2).Cells(h).Value = suma50
                    Next
                    DataGridView1.Rows(2).DefaultCellStyle.BackColor = Color.IndianRed
                    Dim IO As Integer
                    Dim SUMA30 As Double
                    IO = Convert.ToInt32(DataGridView1.ColumnCount.ToString())

                    For W = 0 To IO - 1

                        SUMA30 = SUMA30 + DataGridView1.Rows(2).Cells(W).Value
                    Next
                    TextBox3.Text = Format(SUMA30, "##,##00.00")

                End If

            Else
                Me.Chart1.Series.Clear()
                MsgBox("NO EXISTE INFORMACION PARA MOSTRAR")
            End If

        End If





    End Sub

End Class