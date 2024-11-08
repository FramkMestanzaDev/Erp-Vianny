Imports System.Windows.Forms.DataVisualization.Charting
Public Class Reporte_Cobranzas

    Dim dt, dt1, dt2, dt3 As New DataTable

    Private Sub ACTUALIZAR()
        Dim fd As New ffactura
        Dim jj As New vfactura
        Dim t As New Integer
        Dim AD1 As Integer

        jj.gccia = Label18.Text
        jj.gperiodo = Label17.Text
        dt = fd.buscar_facturas_cobranzas(jj)
        dt1 = fd.buscar_facturas_anuales(jj)
        DataGridView1.DataSource = dt
        DataGridView2.DataSource = dt1

        t = DataGridView1.Rows.Count

        For i = 0 To t - 1
            DataGridView3.Columns.Add(DataGridView1.Rows(i).Cells(0).Value, DataGridView1.Rows(i).Cells(0).Value)
            DataGridView3.Columns(i).HeaderText = DataGridView1.Rows(i).Cells(0).Value
            DataGridView3.Columns(i).Name = DataGridView1.Rows(i).Cells(0).Value

        Next
        DataGridView3.Rows.Add(2)

        For i = 0 To t - 1
            DataGridView3.Rows(1).Cells(i).Value = DataGridView1.Rows(i).Cells(1).Value
            DataGridView3.Rows(0).Cells(i).Value = DataGridView2.Rows(i).Cells(1).Value
        Next
        Me.Chart1.Series.Clear()
        Me.Chart1.Series.Add("VENTAS")
        Me.Chart1.Series.Add("COBRANZAS")


        AD1 = Convert.ToInt32(DataGridView3.ColumnCount.ToString())
        For i = 0 To AD1 - 1
            Me.Chart1.Series("VENTAS").Points.AddXY(i + 1, DataGridView3.Rows(0).Cells(i).Value)
            Me.Chart1.Series("COBRANZAS").Points.AddXY(i + 1, DataGridView3.Rows(1).Cells(i).Value)
            'SUMA20 = SUMA20 + DataGridView1.Rows(1).Cells(i).Value

        Next
        DataGridView3.Columns(0).Frozen = True
        DataGridView3.Columns(1).Frozen = True
        DataGridView3.Columns(2).Frozen = True
        DataGridView3.Columns(3).Frozen = True

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = ""
        DataGridView2.DataSource = ""
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()
        ACTUALIZAR()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.DataSource = ""
        DataGridView2.DataSource = ""
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()
        GroupBox3.Visible = True
        Dim fd As New fventas
        Dim hj As New vventas
        Dim t As New Integer
        Dim AD1 As Integer
        Select Case ComboBox1.Text.ToString.Trim
            Case "VPIZARRO" : hj.gVENDEDOR = "0027"
            Case "VSILVERIO" : hj.gVENDEDOR = "0005"
            Case "GBALVIN" : hj.gVENDEDOR = "0028"
            Case "GBEDON" : hj.gVENDEDOR = "0010"
            Case "VINCIO" : hj.gVENDEDOR = "0022"
            Case "DBRAVO" : hj.gVENDEDOR = "0023"
            Case "AMENDO" : hj.gVENDEDOR = "0026"
            Case "GCUEVA" : hj.gVENDEDOR = "0029"
            Case "JSALINAS" : hj.gVENDEDOR = "0025"
            Case "VGRAUS" : hj.gVENDEDOR = "0007"
            Case "WSALINAS" : hj.gVENDEDOR = "0034"
            Case "MCRUZAFO" : hj.gVENDEDOR = "0035"
        End Select
        hj.gccia = Label18.Text
        hj.gANO = Label17.Text
        dt = fd.buscar_cobranzas_VENDEDOR(hj)
        dt1 = fd.buscar_ventas_VENDEDOR(hj)
        DataGridView1.DataSource = dt
        DataGridView2.DataSource = dt1

        t = DataGridView1.Rows.Count

        For i = 0 To t - 1
            DataGridView3.Columns.Add(DataGridView1.Rows(i).Cells(0).Value, DataGridView1.Rows(i).Cells(0).Value)
            DataGridView3.Columns(i).HeaderText = DataGridView1.Rows(i).Cells(0).Value
            DataGridView3.Columns(i).Name = DataGridView1.Rows(i).Cells(0).Value

        Next
        DataGridView3.Rows.Add(2)

        For i = 0 To t - 1
            DataGridView3.Rows(1).Cells(i).Value = DataGridView1.Rows(i).Cells(1).Value
            DataGridView3.Rows(0).Cells(i).Value = DataGridView2.Rows(i).Cells(1).Value
        Next
        Me.Chart1.Series.Clear()
        Me.Chart1.ChartAreas.Clear()
        Me.Chart1.Series.Add("1")
        Chart1.Series.Add("2")
        Dim area As New ChartArea()
        Chart1.ChartAreas().Add(area)
        AD1 = Convert.ToInt32(DataGridView3.ColumnCount.ToString())
        For i = 0 To AD1 - 1
            Me.Chart1.Series("1").Points.AddXY(i + 1, DataGridView3.Rows(0).Cells(i).Value)
            Me.Chart1.Series("2").Points.AddXY(i + 1, DataGridView3.Rows(1).Cells(i).Value)
            'SUMA20 = SUMA20 + DataGridView1.Rows(1).Cells(i).Value

        Next
        DataGridView3.Columns(0).Frozen = True
        DataGridView3.Columns(1).Frozen = True
        DataGridView3.Columns(2).Frozen = True
        DataGridView3.Columns(3).Frozen = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DataGridView5.DataSource = ""
        DataGridView6.DataSource = ""
        DataGridView4.Rows.Clear()
        DataGridView4.Columns.Clear()
        ACTUALIZARS()
    End Sub

    Private Sub ACTUALIZARS()
        Dim fd As New ffactura
        Dim t As New Integer
        Dim AD1 As Integer


        dt2 = fd.buscar_facturas_anuales_NEGRAS
        dt3 = fd.buscar_facturas_anuales_NEGRAS_ventas
        DataGridView5.DataSource = dt2
        DataGridView6.DataSource = dt3

        t = DataGridView1.Rows.Count

        For i = 0 To t - 1
            DataGridView4.Columns.Add(DataGridView5.Rows(i).Cells(0).Value, DataGridView5.Rows(i).Cells(0).Value)
            DataGridView4.Columns(i).HeaderText = DataGridView5.Rows(i).Cells(0).Value
            DataGridView4.Columns(i).Name = DataGridView5.Rows(i).Cells(0).Value

        Next
        DataGridView4.Rows.Add(2)

        For i = 0 To t - 1
            DataGridView4.Rows(1).Cells(i).Value = DataGridView5.Rows(i).Cells(1).Value
            DataGridView4.Rows(0).Cells(i).Value = DataGridView6.Rows(i).Cells(1).Value
        Next
        Me.Chart2.Series.Clear()
        Me.Chart2.Series.Add("VENTAS")
        Me.Chart2.Series.Add("COBRANZAS")


        AD1 = Convert.ToInt32(DataGridView4.ColumnCount.ToString())
        For i = 0 To AD1 - 1
            Me.Chart2.Series("VENTAS").Points.AddXY(i + 1, DataGridView4.Rows(0).Cells(i).Value)
            Me.Chart2.Series("COBRANZAS").Points.AddXY(i + 1, DataGridView4.Rows(1).Cells(i).Value)
            'SUMA20 = SUMA20 + DataGridView1.Rows(1).Cells(i).Value

        Next
        DataGridView4.Columns(0).Frozen = True
        DataGridView4.Columns(1).Frozen = True
        DataGridView4.Columns(2).Frozen = True
        DataGridView4.Columns(3).Frozen = True

    End Sub
    Private Sub Reporte_Cobranzas_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Label18.Text = "02" Then
            ComboBox1.Items.Add("MCRUZADO")

        End If
        ComboBox1.SelectedIndex = 0
        ACTUALIZAR()
    End Sub
End Class