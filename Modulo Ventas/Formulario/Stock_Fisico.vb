Imports System.Data.SqlClient
Public Class Stock_Fisico
    Dim GK, GK1 As New DataTable
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3, ty30 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim TG As New FSTOCK_PAR
        Dim HJ As New VSTOCK_PAR
        Dim k As Integer
        Dim jk As Double
        Try
            jk = 0
            HJ.gCCIA = Label6.Text

            HJ.gALMACEN = ComboBox1.SelectedValue.ToString
            HJ.gMODAL = 2
            HJ.gano = Label5.Text
            GK = TG.CaeSoft_ReporteStockFisico(HJ)
            If IsNothing(GK) Then
                MsgBox("NO EXISTE STOCK")
                DataGridView1.DataSource = Nothing
            Else
                If GK.Rows.Count > 0 Then
                    DataGridView1.DataSource = GK
                    k = DataGridView1.Rows.Count
                    For i = 0 To k - 1
                        jk = jk + DataGridView1.Rows(i).Cells(4).Value
                        If DataGridView1.Rows(i).Cells(9).Value = 2 Then
                            DataGridView1.Rows(i).Cells(10).Value = "SI"
                            DataGridView1.Columns(10).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(9).Value = 1 Then
                            DataGridView1.Rows(i).Cells(10).Value = "NO"
                            DataGridView1.Columns(10).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(9).Value = 0 Then
                            DataGridView1.Rows(i).Cells(10).Value = "X CONF."
                            DataGridView1.Columns(10).ReadOnly = True
                        End If
                        If DataGridView1.Rows(i).Cells(11).Value = 1 Then

                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                        End If
                        If DataGridView1.Rows(i).Cells(4).Value < 0 Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                            DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        End If
                        DataGridView1.Rows(i).Cells(13).Value = DataGridView1.Rows(i).Cells(4).Value - DataGridView1.Rows(i).Cells(12).Value
                        If HJ.gALMACEN = "41" Or HJ.gALMACEN = "44" Then
                            DataGridView1.Columns(2).Visible = False
                            DataGridView1.Columns(3).Visible = False
                            DataGridView1.Columns(4).Visible = False
                            DataGridView1.Columns(12).Visible = False
                        End If
                    Next
                    TextBox3.Text = jk
                    DataGridView1.Columns(0).Width = 165
                    DataGridView1.Columns(1).Width = 430
                    DataGridView1.Columns(2).Width = 80
                    DataGridView1.Columns(3).Width = 85
                    DataGridView1.Columns(4).Width = 65
                    DataGridView1.Columns(5).Visible = False
                    DataGridView1.Columns(6).Width = 65
                    DataGridView1.Columns(7).Width = 65

                    DataGridView1.Columns(8).Visible = False
                    DataGridView1.Columns(6).Visible = False
                    DataGridView1.Columns(7).Visible = False
                    DataGridView1.Columns(9).Visible = False
                    DataGridView1.Columns(10).Visible = False
                    DataGridView1.Columns(11).Visible = False
                Else

                    MsgBox("NO EXISTE STOCK")
                    DataGridView1.DataSource = Nothing
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar_PARTIDA()
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(GK.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "PRODUCTO" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then
                Dim jk As Double
                DataGridView1.DataSource = dv
                jk = 0
                Dim K As Integer
                DataGridView1.Columns(0).Width = 165
                DataGridView1.Columns(1).Width = 430
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Width = 65
                DataGridView1.Columns(8).Width = 65
                DataGridView1.Columns(8).Visible = False
                K = DataGridView1.Rows.Count
                For i = 0 To K - 1
                    jk = jk + Convert.ToDouble(DataGridView1.Rows(i).Cells(4).Value)
                    If DataGridView1.Rows(i).Cells(9).Value = 2 Then
                        DataGridView1.Rows(i).Cells(10).Value = "SI"
                        DataGridView1.Columns(10).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(9).Value = 1 Then
                        DataGridView1.Rows(i).Cells(10).Value = "NO"
                        DataGridView1.Columns(10).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(9).Value = 0 Then
                        DataGridView1.Rows(i).Cells(10).Value = "X CONF."
                        DataGridView1.Columns(10).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(11).Value = 1 Then

                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                    End If
                    If DataGridView1.Rows(i).Cells(4).Value < 0 Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    End If
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView1.Rows(i).Cells(4).Value - DataGridView1.Rows(i).Cells(12).Value
                Next
                TextBox3.Text = jk.ToString("N2")
                DataGridView1.Columns(0).Width = 165
                DataGridView1.Columns(1).Width = 430
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Width = 65

                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub
    Sub llenar_combo_alamcenes(ByVal ccia As String)

        Try

            conn = New SqlDataAdapter("select talm_15m + '  |  '+ nomb_15m as nombre, talm_15m as almacen from custom_vianny.dbo.Mag1500 where seriencr_15m in ('1','2','3','4','5') and flag_15m ='1' and ccia ='" + ccia + "' order by talm_15m", conx)
            conn.Fill(ty30)
            ComboBox1.DataSource = ty30
            ComboBox1.DisplayMember = "nombre"
            ComboBox1.ValueMember = "almacen"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Stock_Fisico_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ty30.Clear()
        abrir()
        llenar_combo_alamcenes(Trim(Label6.Text))

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim ext As New Exportar
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        buscar_clasificacion()
    End Sub

    Private Sub buscar_PARTIDA()
        Try
            Dim ds As New DataSet

            ds.Tables.Add(GK.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Double
            Dim K As Integer
            dv.RowFilter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = 0

                DataGridView1.Columns(0).Width = 165
                DataGridView1.Columns(1).Width = 430
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Width = 65
                DataGridView1.Columns(8).Width = 65
                DataGridView1.Columns(8).Visible = False
                K = DataGridView1.Rows.Count
                For i = 0 To K - 1
                    jk = jk + Convert.ToDouble(DataGridView1.Rows(i).Cells(4).Value)
                    If DataGridView1.Rows(i).Cells(9).Value = 2 Then
                        DataGridView1.Rows(i).Cells(10).Value = "SI"
                        DataGridView1.Columns(10).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(9).Value = 1 Then
                        DataGridView1.Rows(i).Cells(10).Value = "NO"
                        DataGridView1.Columns(10).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(9).Value = 0 Then
                        DataGridView1.Rows(i).Cells(10).Value = "X CONF."
                        DataGridView1.Columns(10).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(11).Value = 1 Then

                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                    End If
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView1.Rows(i).Cells(4).Value - DataGridView1.Rows(i).Cells(12).Value
                    If DataGridView1.Rows(i).Cells(4).Value < 0 Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    End If
                Next
                TextBox3.Text = jk.ToString("N2")
                DataGridView1.Columns(0).Width = 165
                DataGridView1.Columns(1).Width = 430
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Width = 65

                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar_clasificacion()
        Try
            Dim ds As New DataSet


            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(GK.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Double
            Dim K As Integer
            dv.RowFilter = "CLASIFICACION" & " like '%" & TextBox4.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = 0

                DataGridView1.Columns(0).Width = 165
                DataGridView1.Columns(1).Width = 430
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Width = 65
                DataGridView1.Columns(8).Width = 65
                DataGridView1.Columns(8).Visible = False
                K = DataGridView1.Rows.Count
                For i = 0 To K - 1
                    jk = jk + Convert.ToDouble(DataGridView1.Rows(i).Cells(4).Value)
                    If DataGridView1.Rows(i).Cells(9).Value = 2 Then
                        DataGridView1.Rows(i).Cells(10).Value = "SI"
                        DataGridView1.Columns(10).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(9).Value = 1 Then
                        DataGridView1.Rows(i).Cells(10).Value = "NO"
                        DataGridView1.Columns(10).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(9).Value = 0 Then
                        DataGridView1.Rows(i).Cells(10).Value = "X CONF."
                        DataGridView1.Columns(10).ReadOnly = True
                    End If
                    If DataGridView1.Rows(i).Cells(11).Value = 1 Then

                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                    End If
                    DataGridView1.Rows(i).Cells(13).Value = DataGridView1.Rows(i).Cells(4).Value - DataGridView1.Rows(i).Cells(12).Value
                    If DataGridView1.Rows(i).Cells(4).Value < 0 Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    End If
                Next
                TextBox3.Text = jk.ToString("N2")
                DataGridView1.Columns(0).Width = 165
                DataGridView1.Columns(1).Width = 430
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 65
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Width = 65

                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim j As Integer

        j = e.RowIndex
        Separar_Partidas.TextBox1.Text = Trim(DataGridView1.Rows(j).Cells(2).Value)
        Separar_Partidas.TextBox2.Text = Trim(DataGridView1.Rows(j).Cells(0).Value)
        Separar_Partidas.TextBox3.Text = Trim(DataGridView1.Rows(j).Cells(1).Value)

        Separar_Partidas.TextBox5.Text = Trim(DataGridView1.Rows(j).Cells(13).Value)
        'Select Case ComboBox1.Text
        '    Case "37 -- PRIMERA PTA2" : Separar_Partidas.TextBox4.Text = "37"
        '    Case "38 -- SEGUNDA PT2" : Separar_Partidas.TextBox4.Text = "38"
        '    Case "39 -- TERCERA PT2" : Separar_Partidas.TextBox4.Text = "39"
        '    Case "61 -- MUESTRAS" : Separar_Partidas.TextBox4.Text = "61"
        '    Case "10 -- PRIMERA CHILCA" : Separar_Partidas.TextBox4.Text = "10"
        '    Case "51 -- SEGUNDA CHILCA" : Separar_Partidas.TextBox4.Text = "51"
        '    Case "54 -- TERCERA CHILCA" : Separar_Partidas.TextBox4.Text = "54"
        '    Case "03 -- HILADO OPERATIVO" : Separar_Partidas.TextBox4.Text = "03"
        '    Case "42 -- HILADO TEÑIDO CHILCA" : Separar_Partidas.TextBox4.Text = "42"
        '    Case "59 -- SERVICIOS TERCEROS" : Separar_Partidas.TextBox4.Text = "59"
        '    Case "67 -- ALMACEN HILO - TELA G" : Separar_Partidas.TextBox4.Text = "67"
        '    Case "13 -- ALMACEN AVIOS" : Separar_Partidas.TextBox4.Text = "13"
        '    Case "68 -- ALMACEN HUACHIPA" : Separar_Partidas.TextBox4.Text = "68"
        '    Case "66 -- MERMA" : Separar_Partidas.TextBox4.Text = "66"
        '    Case "08 -- TELA PLANA" : Separar_Partidas.TextBox4.Text = "08"

        'End Select
        Separar_Partidas.TextBox4.Text = ComboBox1.SelectedValue.ToString()
        Separar_Partidas.Label6.Text = Label6.Text
        Separar_Partidas.Show()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then
            For i = 0 To DataGridView1.Columns.Count - 1

                DataGridView1.Columns.Item(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
        End If


    End Sub
End Class