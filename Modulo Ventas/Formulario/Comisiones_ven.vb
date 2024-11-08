Imports System.Data.SqlClient
Public Class Comisiones_ven
    Dim dt As DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim da As New DataTable
    Dim Rsr2, Rsr3 As SqlDataReader
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim va11, va21 As String
        va11 = ""
        va21 = ""
        da.Rows.Clear()



        If Label15.Text = 0 Then

            abrir()
            Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + TextBox3.Text + "' and ccia_ven ='01'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            If Rsr2.Read() = True Then
                va11 = Rsr2(0)
            End If
            Rsr2.Close()
            'Select Case TextBox3.Text
            '    Case "GBEDON" : va11 = "0010"
            '    Case "VINCIO" : va11 = "0022"
            '    Case "DBRAVO" : va11 = "0023"
            '    Case "JSALINAS" : va11 = "0025"
            '    Case "GCUEVA" : va11 = "0029"
            '    Case "AMENDO" : va11 = "0026"
            '    Case "VGRAUS" : va11 = "0007"
            '    Case "VPIZARRO" : va11 = "0027"
            '    Case "GBALVIN" : va11 = "0028"
            '    Case "VSILVERIO" : va11 = "0005"
            '    Case "WSALINAS" : va11 = "0034"
            'End Select
        Else
            If Label15.Text = 1 Then
                abrir()
                Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + ComboBox1.Text + "' and ccia_ven ='01'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr3 = cmd102.ExecuteReader()
                If Rsr3.Read() = True Then
                    va11 = Rsr3(0)
                End If
                Rsr3.Close()
                'Select Case ComboBox1.Text
                '    Case "GBEDON" : va11 = "0010"
                '    Case "VINCIO" : va11 = "0022"
                '    Case "DBRAVO" : va11 = "0023"
                '    Case "JSALINAS" : va11 = "0025"
                '    Case "GCUEVA" : va11 = "0029"
                '    Case "AMENDO" : va11 = "0026"
                '    Case "VGRAUS" : va11 = "0007"
                '    Case "VPIZARRO" : va11 = "0027"
                '    Case "GBALVIN" : va11 = "0028"
                '    Case "VSILVERIO" : va11 = "0005"
                '    Case "WSALINAS" : va11 = "0034"
                'End Select
            End If
        End If


        Select Case ComboBox2.Text
            Case "ENERO" : va21 = "01"
            Case "FEBRERO" : va21 = "02"
            Case "MARZO" : va21 = "03"
            Case "ABRIL" : va21 = "04"
            Case "MAYO" : va21 = "05"
            Case "JUNIO" : va21 = "06"
            Case "JULIO" : va21 = "07"
            Case "AGOSTO" : va21 = "08"
            Case "SETIEMBRE" : va21 = "09"
            Case "OCTUBRE" : va21 = "10"
            Case "NOVIEMBRE" : va21 = "11"
            Case "DICIEMBRE" : va21 = "12"
        End Select

        abrir()
        Dim Rs As SqlDataReader

        Dim sql As String = ("select count(id) from comsion_vendedor where id ='" + Trim(Label7.Text) + Trim(Label8.Text) + va21 + va11 + "'")
        Dim cmd1 As New SqlCommand(sql, conx)
        Rs = cmd1.ExecuteReader()
        Rs.Read()
        If Rs(0) > 0 Then
            abrir()
            Dim cmd As New SqlDataAdapter("SELECT grupo AS GRUPO,cliente AS CLIENTE,rubro AS RUBRO,f_pago AS F_PAGO,serie AS SERIE,correlativo AS CORRELATIVO,precio_venta AS PRECIO_VENTA,valor_venta AS VALOR_VENTA,venta_descuento AS VENTA_DESC,porcentaje AS PORCENTAJE,comision AS COMISION, almacen AS ALMACEN,ventas_mes,'' AS PAGADO  from comsion_vendedor where id = '" + Trim(Label7.Text) + Trim(Label8.Text) + va21 + va11 + "'", conx)

            cmd.Fill(da)
            DataGridView1.DataSource = da

            Dim ui As Integer
            Dim suma, suma1 As Double
            ui = DataGridView1.Rows.Count

            Dim VA13, JU1 As Double
            For i = 0 To ui - 1
                suma1 = suma1 + DataGridView1.Rows(i).Cells(7).Value
            Next
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(1).Width = 90
            DataGridView1.Columns(2).Width = 140
            DataGridView1.Columns(3).Width = 120
            DataGridView1.Columns(4).Width = 90
            DataGridView1.Columns(6).Width = 60
            DataGridView1.Columns(7).Width = 90
            DataGridView1.Columns(12).Width = 80
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            TextBox2.Text = suma1.ToString("N2")


            VA13 = DataGridView1.Rows(0).Cells(12).Value
            JU1 = Convert.ToDouble(TextBox2.Text) - VA13
            TextBox5.Text = JU1.ToString("N2")
            If TextBox2.Text > 0 Then

                TextBox6.Visible = True
                TextBox6.Text = "LA VENTAS SUPERAN LA CUOTA MENSUAL,SI COMISIONARA "
            Else
                TextBox6.Visible = True
                TextBox6.Text = "LA VENTAS NO SUPERAN LA CUOTA MENSUAL,POR LO TANTO SOLO SE LE PAGARA SU BASICO"
                TextBox1.Text = "0.00"
                TextBox5.Text = "0.00"
            End If

            Label16.Text = DataGridView1.Rows(0).Cells(12).Value
            ui = DataGridView1.Rows.Count
            For i = 0 To ui - 1

                suma = suma + DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Columns(8).DefaultCellStyle.BackColor = Color.DarkGray
                DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Columns(10).DefaultCellStyle.BackColor = Color.DarkSeaGreen
                DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.DarkTurquoise
                DataGridView1.Columns(8).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Columns(9).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Columns(10).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Columns(11).DefaultCellStyle.ForeColor = Color.White
            Next

            TextBox1.Text = suma.ToString("N2")
            Button2.Enabled = True
            Button3.Enabled = True
        Else
            Button2.Enabled = False
            Button3.Enabled = False
            TextBox2.Text = ""
            TextBox5.Text = ""
            TextBox1.Text = ""
            MsgBox("EL JEFE COMERCIAL TODAVIA NO VALIDA SU COMISION")
            DataGridView1.DataSource = ""
        End If
        Rs.Close()




    End Sub
    Dim Rsr22, Rsr222 As SqlDataReader
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        Dim va11, va21 As String
        va11 = ""
        va21 = ""
        If Label15.Text = 0 Then

            abrir()
            Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + TextBox3.Text + "' and ccia_ven ='01'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            If Rsr2.Read() = True Then
                va11 = Rsr2(0)
            End If
            Rsr2.Close()
        Else
            If Label15.Text = 1 Then
                abrir()
                Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + ComboBox1.Text + "' and ccia_ven ='01'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr3 = cmd102.ExecuteReader()
                If Rsr3.Read() = True Then
                    va11 = Rsr3(0)
                End If
                Rsr3.Close()
            End If
        End If



        Select Case ComboBox2.Text
            Case "ENERO" : va21 = "01"
            Case "FEBRERO" : va21 = "02"
            Case "MARZO" : va21 = "03"
            Case "ABRIL" : va21 = "04"
            Case "MAYO" : va21 = "05"
            Case "JUNIO" : va21 = "06"
            Case "JULIO" : va21 = "07"
            Case "AGOSTO" : va21 = "08"
            Case "SETIEMBRE" : va21 = "09"
            Case "OCTUBRE" : va21 = "10"
            Case "NOVIEMBRE" : va21 = "11"
            Case "DICIEMBRE" : va21 = "12"
        End Select
        Dim idd As String
        idd = Trim(Label7.Text) + Trim(Label8.Text) + va21 + va11

        Dim sql1022 As String = "select distinct apgg from comsion_vendedor  where id ='" + idd + "'"
        Dim cmd1022 As New SqlCommand(sql1022, conx)
        Rsr222 = cmd1022.ExecuteReader()

        If Rsr222.Read() Then

            If Rsr222(0) = "1" Then

                Reporte_comisiones.TextBox1.Text = Trim(Label7.Text) + Trim(Label8.Text) + va21 + va11
                Reporte_comisiones.Show()
            Else
                MsgBox("LA COMISION NO ESTA APROBADA POR LA GERENCIA, POR LO QUE NO SE PUEDE IMPRIMIR")
            End If
        Else

        End If
        Rsr222.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox3.Text = " " Or ComboBox2.Text = " " Then
            MsgBox("FALTA SELECCIONAR UN VENDEDOR O MES")
        Else
            Dim jk As New Exportar
            jk.llenarExcel(DataGridView1)
        End If
    End Sub
    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE rpm_ven ='2' and ccia_ven ='" + Label8.Text + "'", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "alias_ven"
            ComboBox1.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Comisiones_ven_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box2()
        ComboBox2.SelectedIndex = DateTime.Now.ToString("MM") - 1
        If Label15.Text = 1 Then
            TextBox3.Visible = False
        Else
            If Label15.Text = 0 Then
                ComboBox1.Visible = False
            End If
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox4.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim ui As Integer
                Dim suma, suma1, VA As Double
                ui = DataGridView1.Rows.Count
                For i = 0 To ui - 1
                    VA = (DataGridView1.Rows(i).Cells(9).Value * DataGridView1.Rows(i).Cells(10).Value) / 100
                    DataGridView1.Rows(i).Cells(11).Value = VA.ToString("N2")
                    suma = suma + DataGridView1.Rows(i).Cells(11).Value
                    suma1 = suma1 + DataGridView1.Rows(i).Cells(9).Value
                    DataGridView1.Columns(8).DefaultCellStyle.BackColor = Color.DarkGray
                    DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
                    DataGridView1.Columns(10).DefaultCellStyle.BackColor = Color.DarkSeaGreen
                    DataGridView1.Columns(11).DefaultCellStyle.BackColor = Color.DarkTurquoise
                    DataGridView1.Columns(8).DefaultCellStyle.ForeColor = Color.White
                    DataGridView1.Columns(9).DefaultCellStyle.ForeColor = Color.White
                    DataGridView1.Columns(10).DefaultCellStyle.ForeColor = Color.White
                    DataGridView1.Columns(11).DefaultCellStyle.ForeColor = Color.White
                Next
                DataGridView1.Columns(0).Visible = False
                TextBox1.Text = suma
                TextBox2.Text = suma1
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        buscar()
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim fg As New vfactura
        Dim fg1 As New ffactura
        Dim oo As Integer
        Dim JJ As String
        oo = DataGridView1.Rows.Count
        For i = 0 To oo - 1

            If Trim(DataGridView1.Rows(i).Cells(0).Value) = "INGRESOS" Then
                If Mid(DataGridView1.Rows(i).Cells(4).Value, 1, 1) = "F" Then
                    fg.gdoc = "001"
                Else
                    If Mid(DataGridView1.Rows(i).Cells(4).Value, 1, 1) = "B" Then
                        fg.gdoc = "003"
                    End If

                End If

                fg.gndoc = Trim(DataGridView1.Rows(i).Cells(4).Value) + "-" + Trim(DataGridView1.Rows(i).Cells(5).Value)
                fg.gCLIENTE = Trim(DataGridView1.Rows(i).Cells(2).Value)
                fg.gperiodo = Label7.Text
                fg.gccia = Label8.Text

                JJ = fg1.mostrar_estado_facbol2(fg)
                If JJ = "False" Then
                    DataGridView1.Rows(i).Cells(13).Value = fg1.mostrar_estado_facbol(fg)
                Else
                    DataGridView1.Rows(i).Cells(13).Value = fg1.mostrar_estado_facbol2(fg)
                End If

            Else
                DataGridView1.Rows(i).Cells(13).Value = "CANCELADO"
            End If

        Next
    End Sub
End Class