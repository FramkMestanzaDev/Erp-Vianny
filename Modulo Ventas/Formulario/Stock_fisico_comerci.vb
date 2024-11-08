Imports System.Data.SqlClient
Public Class Stock_fisico_comerci
    Dim GK, GK1, gk2 As New DataTable
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
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'If e.ColumnIndex = 0 Then

        '    ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
        '    Dim num11, num12 As Integer
        '    num11 = e.RowIndex.ToString
        '    num12 = e.ColumnIndex.ToString


        '    If Trim(DataGridView1.Rows(num11).Cells(12).Value) = "" Then
        '        MsgBox("NO EXISTE RUTA DEL DOCUMENTO A MOSTRAR")
        '    Else
        '        'Process.Start(Application.StartupPath & Trim(DataGridView1.Rows(num11).Cells(11).Value))
        '        'MsgBox(Trim(DataGridView1.Rows(num11).Cells(11).Value)
        '        Dim XL As New Microsoft.Office.Interop.Excel.Application 'Crea el objeto excel

        '        XL.Workbooks.Open(Trim(DataGridView1.Rows(num11).Cells(12).Value), , True) 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)

        '        XL.Visible = True

        '        XL.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized
        '    End If

        'End If
    End Sub
    Dim dt, DT1 As New DataTable
    Private Sub buscar3()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(GK.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim k As Integer
            Dim jk As Double

            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                k = DataGridView1.Rows.Count
                For i = 0 To k - 1
                    jk = jk + DataGridView1.Rows(i).Cells(6).Value
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
                TextBox1.Text = jk
                'DataGridView1.Columns(0).Width = 165
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(GK.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim k As Integer
            Dim jk As Double

            dv.RowFilter = "PARTIDA" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                k = DataGridView1.Rows.Count
                For i = 0 To k - 1
                    jk = jk + DataGridView1.Rows(i).Cells(6).Value
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
                TextBox1.Text = jk
                'DataGridView1.Columns(0).Width = 165
                DataGridView1.Columns(1).Width = 145
                DataGridView1.Columns(2).Width = 146
                DataGridView1.Columns(3).Width = 340
                DataGridView1.Columns(5).Width = 75
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 45
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 80
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(12).Visible = False

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar3()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim sd As New Exportar
        sd.llenarExcel(DataGridView1)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar2()
    End Sub

    Private Sub Stock_fisico_comerci_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Text = ""
        TextBox3.Text = ""
        Try
            If Label6.Text = "01" Or Label6.Text = "03" Then
                ComboBox1.SelectedIndex = 0
            Else
                ComboBox1.Items.Clear()
                ComboBox1.Items.Add("67 -- ALMACEN HILO - TELA G")
                ComboBox1.SelectedIndex = 0
            End If

        Catch ex As Exception

        End Try
    End Sub
    Dim KKK As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim TG As New FSTOCK_PAR
        Dim HJ, HJ3 As New VSTOCK_PAR
        Dim k As Integer
        Dim jk As Double
        Try
            GK.Clear()

            If Label6.Text = "01" Or Label6.Text = "03" Then
                If ComboBox1.Text = "PRIMERA" Then
                    HJ.gano = Label5.Text
                    HJ.gCCIA = Label6.Text
                    GK = TG.CaeSoft_ReporteStockFisico_COMERCIAL(HJ)
                    DataGridView1.DataSource = GK
                End If
                If ComboBox1.Text = "SEGUNDA" Then
                    HJ.gano = Label5.Text
                    HJ.gCCIA = Label6.Text
                    'GK1 = TG.CaeSoft_ReporteStockFisico_COMERCIA2L(HJ)
                    GK = TG.CaeSoft_ReporteStockFisico_COMERCIA2L(HJ)
                    DataGridView1.DataSource = GK
                End If
                If ComboBox1.Text = "HILO _TEÑIDO" Then
                    HJ.gALMACEN = "42"
                    HJ.gCCIA = Label6.Text
                    HJ.gMODAL = "02"
                    HJ.gano = Label5.Text
                    GK = TG.CaeSoft_ReporteStockFisico_3737(HJ)
                    'gk2 = TG.CaeSoft_ReporteStockFisico_3737(HJ)
                    DataGridView1.DataSource = GK
                End If
                If ComboBox1.Text = "TERCERA" Then

                    HJ.gano = Label5.Text
                    HJ.gCCIA = Label6.Text
                    'gk2 = TG.CaeSoft_ReporteStockFisico_tercera(HJ)
                    GK = TG.CaeSoft_ReporteStockFisico_tercera(HJ)
                    DataGridView1.DataSource = GK

                End If
                If ComboBox1.Text = "HILO_CRUDO" Then
                    HJ.gALMACEN = "03"
                    HJ.gCCIA = Label6.Text
                    HJ.gMODAL = "02"
                    HJ.gano = Label5.Text
                    'gk2 = TG.CaeSoft_ReporteStockFisico_3737(HJ)
                    GK = TG.CaeSoft_ReporteStockFisico_3737(HJ)
                    DataGridView1.DataSource = GK
                End If
                If ComboBox1.Text = "TELA_PLANA" Then
                    HJ.gALMACEN = "08"
                    HJ.gCCIA = Label6.Text
                    HJ.gMODAL = "02"
                    HJ.gano = Label5.Text
                    'gk2 = TG.CaeSoft_ReporteStockFisico_3737(HJ)
                    GK = TG.CaeSoft_ReporteStockFisico_3737(HJ)
                    DataGridView1.DataSource = GK
                End If
            Else
                ComboBox1.Items.Clear()
                ComboBox1.Items.Add("HILO TELA GRAUS")
                HJ3.gano = Label5.Text
                'KKK = TG.CaeSoft_ReporteStockFisico_COMERCIAL_GRAUS(HJ3)
                GK = TG.CaeSoft_ReporteStockFisico_COMERCIAL_GRAUS(HJ3)
                DataGridView1.DataSource = GK
            End If



            k = DataGridView1.Rows.Count
            For i = 0 To k - 1
                jk = jk + DataGridView1.Rows(i).Cells(6).Value
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
                If DataGridView1.Rows(i).Cells(14).Value = 1 Then

                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                End If
            Next
            TextBox1.Text = jk
            'DataGridView1.Columns(0).Width = 165
            DataGridView1.Columns(1).Width = 145
            DataGridView1.Columns(2).Width = 146
            DataGridView1.Columns(3).Width = 340
            DataGridView1.Columns(5).Width = 75
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 45
            DataGridView1.Columns(8).Width = 80
            DataGridView1.Columns(11).Width = 77
            DataGridView1.Columns(13).Width = 90
            DataGridView1.Columns(9).Width = 80

            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(14).Visible = False
            ComboBox1.Select()
        Catch ex As Exception
            MsgBox("NO HAY STOCK INICIAL DEL AÑO EN CURSO")
        End Try
    End Sub
    Dim Rsr225 As SqlDataReader
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 0 Then
            abrir()
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString

            Dim sql10225 As String = "select COUNT(partida) from cab_ingresop where partida ='" + Mid(Trim(DataGridView1.Rows(num1).Cells(4).Value), 1, 6) + "'"
            Dim cmd10225 As New SqlCommand(sql10225, conx)
            Rsr225 = cmd10225.ExecuteReader()
            Rsr225.Read()
            If Rsr225(0) > 0 Then

                Packing_nacional.TextBox1.Text = DataGridView1.Rows(num1).Cells(4).Value
                Packing_nacional.TextBox2.Text = Trim(ComboBox1.Text)
                Packing_nacional.Show()

            Else
                MsgBox("LA PARTIDA NO ESTA REGISTRADA EN EL SISTEMA")
            End If

        End If
    End Sub
End Class