Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Net.Mail
Public Class Form3
    Dim dt As New DataTable

    Sub ACTUALIZAR()
        Try


            Dim gh As New fregistro_venta
            Dim gh1 As New vregistroventa

            Button1.Visible = False
            gh1.ganal1_3 = 0
            gh1.gempresa = Label9.Text
            dt = gh.mostrar_venta(gh1)
            RadioButton2.Checked = True
            DataGridView1.DataSource = dt
            DataGridView1.Columns(4).Width = 250
            DataGridView1.Columns(5).Width = 82
            DataGridView1.Columns(6).Width = 200
            DataGridView1.Columns(0).Frozen = True
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(6).Frozen = True
            DataGridView1.Columns(7).Frozen = True
            DataGridView1.Columns(16).ReadOnly = True
            DataGridView1.Columns(6).DefaultCellStyle.BackColor = Color.Wheat
            DataGridView1.Columns(14).DefaultCellStyle.BackColor = Color.NavajoWhite

            Dim i As Integer
            Dim SUM, SUM2 As Double
            i = DataGridView1.Rows.Count
            For t = 0 To i - 1
                Me.DataGridView1.Columns(14).DefaultCellStyle.Format = "###,###,##0.00"
                Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "###,###,##0.00"


                DataGridView1.Rows(t).Cells(17).Value = DataGridView1.Rows(t).Cells(14).Value - DataGridView1.Rows(t).Cells(15).Value
                Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "###,###,##0.00"
                If DataGridView1.Rows(t).Cells(15).Value > "0.00" Then
                    DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.Wheat
                End If


                Select Case DataGridView1.Rows(t).Cells(9).Value
                    Case "0010" : DataGridView1.Rows(t).Cells(9).Value = "GBEDON"
                    Case "0022" : DataGridView1.Rows(t).Cells(9).Value = "VINCIO"
                    Case "0023" : DataGridView1.Rows(t).Cells(9).Value = "DBRAVO"
                    Case "0025" : DataGridView1.Rows(t).Cells(9).Value = "JSALINAS"
                    Case "0029" : DataGridView1.Rows(t).Cells(9).Value = "GCUEVA"
                    Case "0026" : DataGridView1.Rows(t).Cells(9).Value = "AMENDO"
                    Case "0007" : DataGridView1.Rows(t).Cells(9).Value = "VGRAUS"
                    Case "0027" : DataGridView1.Rows(t).Cells(9).Value = "VPIZARRO"
                    Case "0028" : DataGridView1.Rows(t).Cells(9).Value = "JBALVIN"
                    Case "0005" : DataGridView1.Rows(t).Cells(9).Value = "VSILVERIO"
                    Case "0030" : DataGridView1.Rows(t).Cells(9).Value = "RMEDINA"
                End Select
                Select Case DataGridView1.Rows(t).Cells(8).Value
                    Case "01" : DataGridView1.Rows(t).Cells(8).Value = "CONTADO"
                    Case "02" : DataGridView1.Rows(t).Cells(8).Value = "CRED 07 DIAS"
                    Case "03" : DataGridView1.Rows(t).Cells(8).Value = "CRED 15 DIAS"
                    Case "04" : DataGridView1.Rows(t).Cells(8).Value = "CRED 30 DIAS"
                    Case "05" : DataGridView1.Rows(t).Cells(8).Value = "CRED 45 DIAS"
                    Case "06" : DataGridView1.Rows(t).Cells(8).Value = "CRED 60 DIAS"
                    Case "07" : DataGridView1.Rows(t).Cells(8).Value = "CRED 90 DIAS"
                    Case "08" : DataGridView1.Rows(t).Cells(8).Value = "CRED 120 DIAS"
                    Case "09" : DataGridView1.Rows(t).Cells(8).Value = "CONTRA ENTREGA"

                End Select
                If DataGridView1.Rows(t).Cells(11).Value = "SOLES" Then
                    SUM = SUM + DataGridView1.Rows(t).Cells(14).Value
                Else
                    If DataGridView1.Rows(t).Cells(11).Value = "DOLARES" Then
                        SUM2 = SUM2 + DataGridView1.Rows(t).Cells(14).Value
                    End If

                End If

                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
                DataGridView1.Columns(10).ReadOnly = True
                DataGridView1.Columns(11).ReadOnly = True
                DataGridView1.Columns(12).ReadOnly = True
                DataGridView1.Columns(13).ReadOnly = True
                DataGridView1.Columns(15).ReadOnly = True
                DataGridView1.Columns(17).ReadOnly = True
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(19).Visible = False
            Next
            TextBox6.Text = SUM
            TextBox7.Text = SUM2
            If TextBox4.Text = 2 Then
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(3).Visible = False
            Else
                If TextBox4.Text = 1 Then
                    DataGridView1.Columns(2).Visible = True
                    DataGridView1.Columns(3).Visible = True
                End If

            End If
        Catch ex As Exception
            MsgBox("NO EXISTE DOCUMENTOS PENDIENTES POR CANCELAR")
        End Try

    End Sub
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ACTUALIZAR()
    End Sub
    Private Sub buscar_FPAGO()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "F_PAGO" & " like '%" & TextBox5.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(4).Width = 250
                DataGridView1.Columns(5).Width = 60
                DataGridView1.Columns(6).Width = 200
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(6).Frozen = True
                DataGridView1.Columns(7).Frozen = True
                DataGridView1.Columns(16).ReadOnly = True
                DataGridView1.Columns(6).DefaultCellStyle.BackColor = Color.Wheat
                DataGridView1.Columns(14).DefaultCellStyle.BackColor = Color.NavajoWhite
                DataGridView1.Columns(19).Visible = False

                Dim i As Integer
                Dim SUM, SUM2 As Double
                i = DataGridView1.Rows.Count
                For t = 0 To i - 1
                    DataGridView1.Rows(t).Cells(17).Value = DataGridView1.Rows(t).Cells(14).Value - DataGridView1.Rows(t).Cells(15).Value
                    If DataGridView1.Rows(t).Cells(15).Value > "0.00" Then
                        DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.Wheat
                    End If
                    If DataGridView1.Rows(t).Cells(11).Value = "SOLES" Then
                        SUM = SUM + DataGridView1.Rows(t).Cells(14).Value
                    Else
                        If DataGridView1.Rows(t).Cells(11).Value = "DOLARES" Then
                            SUM2 = SUM2 + DataGridView1.Rows(t).Cells(14).Value
                        End If

                    End If
                Next
                TextBox6.Text = SUM
                TextBox7.Text = SUM2
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
                DataGridView1.Columns(10).ReadOnly = True
                DataGridView1.Columns(11).ReadOnly = True
                DataGridView1.Columns(12).ReadOnly = True
                DataGridView1.Columns(13).ReadOnly = True
                DataGridView1.Columns(15).ReadOnly = True
                DataGridView1.Columns(17).ReadOnly = True
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(19).Visible = False

                If TextBox4.Text = 2 Then
                    DataGridView1.Columns(2).Visible = False
                    DataGridView1.Columns(3).Visible = False
                Else
                    If TextBox4.Text = 1 Then
                        DataGridView1.Columns(2).Visible = True
                        DataGridView1.Columns(3).Visible = True
                    End If

                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar_cliente()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(4).Width = 250
                DataGridView1.Columns(5).Width = 60
                DataGridView1.Columns(6).Width = 200
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(6).Frozen = True
                DataGridView1.Columns(7).Frozen = True
                DataGridView1.Columns(16).ReadOnly = True
                DataGridView1.Columns(6).DefaultCellStyle.BackColor = Color.Wheat
                DataGridView1.Columns(14).DefaultCellStyle.BackColor = Color.NavajoWhite
                DataGridView1.Columns(19).Visible = False
                Dim i As Integer
                Dim SUM, SUM2 As Double
                i = DataGridView1.Rows.Count
                For t = 0 To i - 1
                    DataGridView1.Rows(t).Cells(17).Value = DataGridView1.Rows(t).Cells(14).Value - DataGridView1.Rows(t).Cells(15).Value
                    If DataGridView1.Rows(t).Cells(15).Value > "0.00" Then
                        DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.Wheat
                    End If
                    If DataGridView1.Rows(t).Cells(11).Value = "SOLES" Then
                        SUM = SUM + DataGridView1.Rows(t).Cells(14).Value
                    Else
                        If DataGridView1.Rows(t).Cells(11).Value = "DOLARES" Then
                            SUM2 = SUM2 + DataGridView1.Rows(t).Cells(14).Value
                        End If

                    End If
                Next
                TextBox6.Text = SUM
                TextBox7.Text = SUM2
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
                DataGridView1.Columns(10).ReadOnly = True
                DataGridView1.Columns(11).ReadOnly = True
                DataGridView1.Columns(12).ReadOnly = True
                DataGridView1.Columns(13).ReadOnly = True
                DataGridView1.Columns(15).ReadOnly = True
                DataGridView1.Columns(17).ReadOnly = True
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(19).Visible = False
                If TextBox4.Text = 2 Then
                    DataGridView1.Columns(2).Visible = False
                    DataGridView1.Columns(3).Visible = False
                Else
                    If TextBox4.Text = 1 Then
                        DataGridView1.Columns(2).Visible = True
                        DataGridView1.Columns(3).Visible = True
                    End If
                End If
            Else
                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar_cORRLATIVO()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "COMPROBANTE" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(4).Width = 250
                DataGridView1.Columns(5).Width = 60
                DataGridView1.Columns(6).Width = 200
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(6).Frozen = True
                DataGridView1.Columns(7).Frozen = True
                DataGridView1.Columns(16).ReadOnly = True
                DataGridView1.Columns(6).DefaultCellStyle.BackColor = Color.Wheat
                DataGridView1.Columns(14).DefaultCellStyle.BackColor = Color.NavajoWhite
                DataGridView1.Columns(19).Visible = False
                Dim i As Integer
                Dim SUM, SUM2 As Double
                i = DataGridView1.Rows.Count
                For t = 0 To i - 1
                    DataGridView1.Rows(t).Cells(17).Value = DataGridView1.Rows(t).Cells(14).Value - DataGridView1.Rows(t).Cells(15).Value
                    If DataGridView1.Rows(t).Cells(15).Value > "0.00" Then
                        DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.Wheat
                    End If
                    If DataGridView1.Rows(t).Cells(11).Value = "SOLES" Then
                        SUM = SUM + DataGridView1.Rows(t).Cells(14).Value
                    Else
                        If DataGridView1.Rows(t).Cells(11).Value = "DOLARES" Then
                            SUM2 = SUM2 + DataGridView1.Rows(t).Cells(14).Value
                        End If

                    End If
                Next
                TextBox6.Text = SUM
                TextBox7.Text = SUM2
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
                DataGridView1.Columns(10).ReadOnly = True
                DataGridView1.Columns(11).ReadOnly = True
                DataGridView1.Columns(12).ReadOnly = True
                DataGridView1.Columns(13).ReadOnly = True
                DataGridView1.Columns(15).ReadOnly = True
                DataGridView1.Columns(17).ReadOnly = True
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(19).Visible = False
                If TextBox4.Text = 2 Then
                    DataGridView1.Columns(2).Visible = False
                    DataGridView1.Columns(3).Visible = False
                Else
                    If TextBox4.Text = 1 Then
                        DataGridView1.Columns(2).Visible = True
                        DataGridView1.Columns(3).Visible = True
                    End If
                End If
            Else
                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar_SALIDA()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "SALIDA" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(4).Width = 250
                DataGridView1.Columns(5).Width = 60
                DataGridView1.Columns(6).Width = 200
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(6).Frozen = True
                DataGridView1.Columns(7).Frozen = True
                DataGridView1.Columns(16).ReadOnly = True
                DataGridView1.Columns(6).DefaultCellStyle.BackColor = Color.Wheat
                DataGridView1.Columns(14).DefaultCellStyle.BackColor = Color.NavajoWhite
                DataGridView1.Columns(19).Visible = False
                Dim i As Integer
                Dim SUM, SUM2 As Double
                i = DataGridView1.Rows.Count
                For t = 0 To i - 1
                    DataGridView1.Rows(t).Cells(17).Value = DataGridView1.Rows(t).Cells(14).Value - DataGridView1.Rows(t).Cells(15).Value
                    If DataGridView1.Rows(t).Cells(15).Value > "0.00" Then
                        DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.Wheat
                    End If
                    If DataGridView1.Rows(t).Cells(11).Value = "SOLES" Then
                        SUM = SUM + DataGridView1.Rows(t).Cells(14).Value
                    Else
                        If DataGridView1.Rows(t).Cells(11).Value = "DOLARES" Then
                            SUM2 = SUM2 + DataGridView1.Rows(t).Cells(14).Value
                        End If

                    End If
                Next
                TextBox6.Text = SUM
                TextBox7.Text = SUM2
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
                DataGridView1.Columns(7).ReadOnly = True
                DataGridView1.Columns(8).ReadOnly = True
                DataGridView1.Columns(9).ReadOnly = True
                DataGridView1.Columns(10).ReadOnly = True
                DataGridView1.Columns(11).ReadOnly = True
                DataGridView1.Columns(12).ReadOnly = True
                DataGridView1.Columns(13).ReadOnly = True
                DataGridView1.Columns(15).ReadOnly = True
                DataGridView1.Columns(17).ReadOnly = True
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(19).Visible = False
                If TextBox4.Text = 2 Then
                    DataGridView1.Columns(2).Visible = False
                    DataGridView1.Columns(3).Visible = False
                Else
                    If TextBox4.Text = 1 Then
                        DataGridView1.Columns(2).Visible = True
                        DataGridView1.Columns(3).Visible = True
                    End If
                End If
            Else
                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar_cliente()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar_cORRLATIVO()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar_SALIDA()
    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 2 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            If DataGridView1.Rows(num1).Cells(3).Value = False Then
                DataGridView1.Rows(num1).Cells(3).Value = True
                DataGridView1.Columns(16).DefaultCellStyle.BackColor = Color.White
                DataGridView1.Rows(num1).Cells(16).Value = ""
                Button1.Visible = True
                DataGridView1.Rows(num1).Cells(16).Selected = True
                DataGridView1.Rows(num1).Cells(16).Value = DataGridView1.Rows(num1).Cells(14).Value

            Else
                If DataGridView1.Rows(num1).Cells(3).Value = True Then

                    DataGridView1.Rows(num1).Cells(3).Value = False
                    DataGridView1.Rows(num1).Cells(16).Value = "0.00"
                End If

            End If
        Else
            If e.ColumnIndex = 3 Then

                ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
                Dim num1, num2 As Integer
                num1 = e.RowIndex.ToString
                num2 = e.ColumnIndex.ToString
                DataGridView1.Rows(num1).Cells(2).Value = False
                DataGridView1.Rows(num1).Cells(16).ReadOnly = False
                DataGridView1.Columns(16).DefaultCellStyle.BackColor = Color.DarkTurquoise
                Button1.Visible = True
                DataGridView1.Rows(num1).Cells(16).Selected = True
            Else
                If e.ColumnIndex = 0 Then
                    Dim num1, num2 As Integer
                    num1 = e.RowIndex.ToString
                    num2 = e.ColumnIndex.ToString
                    Detale_Venta.TextBox2.Text = DataGridView1.Rows(num1).Cells(5).Value
                    Detale_Venta.TextBox1.Text = DataGridView1.Rows(num1).Cells(19).Value
                    Detale_Venta.Label1.Text = Label9.Text
                    Detale_Venta.Show()
                Else
                    If e.ColumnIndex = 1 Then

                        ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
                        Dim num1, num2 As Integer
                        num1 = e.RowIndex.ToString
                        num2 = e.ColumnIndex.ToString

                        Dim respuesta As DialogResult


                        respuesta = MessageBox.Show("DESEA IMPRIMIR EL REGISTRO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If (respuesta = Windows.Forms.DialogResult.Yes) Then
                            Form_ventan.TextBox1.Text = Mid(DataGridView1.Rows(num1).Cells(7).Value, 1, 4)
                            Form_ventan.TextBox2.Text = Mid(DataGridView1.Rows(num1).Cells(7).Value, 6, 8)

                            Form_ventan.Show()
                        End If
                    Else
                        If e.ColumnIndex = 7 Then
                            Dim num1, num2 As Integer
                            num1 = e.RowIndex.ToString
                            num2 = e.ColumnIndex.ToString
                            Dim respuesta As DialogResult


                            respuesta = MessageBox.Show("DESEA VISUALIZAR DETALLE DE LOS PAGOS?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                                Moviemiento_de_Comprobantes.Label1.Text = Trim(DataGridView1.Rows(num1).Cells(7).Value)
                                Moviemiento_de_Comprobantes.Label3.Text = Trim(DataGridView1.Rows(num1).Cells(12).Value)
                                Moviemiento_de_Comprobantes.Label4.Text = Trim(DataGridView1.Rows(num1).Cells(11).Value)
                                Moviemiento_de_Comprobantes.Label5.Text = Label9.Text
                                Moviemiento_de_Comprobantes.Label6.Text = Trim(DataGridView1.Rows(num1).Cells(5).Value)
                                Moviemiento_de_Comprobantes.Show()
                            End If

                        End If
                        End If
                End If


            End If
        End If






    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a As Integer
        Dim fh As New fregistro_venta
        Dim fh1 As New vfactura
        Dim fp As New fventasn
        Dim fp1 As New vventas_n
        i = DataGridView1.Rows.Count
        a = 0
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(2).Value = True Then
                fh1.gsfactu = Microsoft.VisualBasic.Left(DataGridView1.Rows(a).Cells(7).Value, 4)
                fh1.gnfactu = Microsoft.VisualBasic.Right(DataGridView1.Rows(a).Cells(7).Value, 8)
                fh1.gperiodo = DataGridView1.Rows(a).Cells(5).Value
                fh1.gccia = Label9.Text
                fh.actualizar_venta(fh1)
                If DataGridView1.Rows(a).Cells(14).Value > "0.00" Then

                    DataGridView1.Rows(a).DefaultCellStyle.BackColor = Color.BurlyWood
                End If

            End If
            If Me.DataGridView1.Rows(a).Cells(3).Value = True Then
                fh1.gsfactu = Microsoft.VisualBasic.Left(DataGridView1.Rows(a).Cells(7).Value, 4)
                fh1.gnfactu = Microsoft.VisualBasic.Right(DataGridView1.Rows(a).Cells(7).Value, 8)
                fh1.gperiodo = DataGridView1.Rows(a).Cells(5).Value
                fh1.gMONTO = DataGridView1.Rows(a).Cells(15).Value + DataGridView1.Rows(a).Cells(16).Value
                fh1.gccia = Label9.Text
                fh.actualizar_venta_MONTO(fh1)

                fp1.gperiodo = DataGridView1.Rows(a).Cells(5).Value
                fp1.gsalida = DataGridView1.Rows(a).Cells(19).Value
                fp1.gcomprobante = DataGridView1.Rows(a).Cells(7).Value
                fp1.gcliente = DataGridView1.Rows(a).Cells(4).Value
                fp1.gvendedor = DataGridView1.Rows(a).Cells(9).Value
                fp1.gfpago = DataGridView1.Rows(a).Cells(8).Value
                fp1.gvalor_venta = DataGridView1.Rows(a).Cells(12).Value
                fp1.gigv = DataGridView1.Rows(a).Cells(13).Value
                fp1.gtotal = DataGridView1.Rows(a).Cells(14).Value
                fp1.gpagado = DataGridView1.Rows(a).Cells(15).Value + DataGridView1.Rows(a).Cells(16).Value
                fp1.gparcial = DataGridView1.Rows(a).Cells(16).Value
                fp1.gobservacion = DataGridView1.Rows(a).Cells(18).Value
                fp1.gfecha = DateTimePicker1.Value
                fp1.gccia = Label9.Text
                fp.insertarventasn(fp1)
            End If

        Next
        MsgBox("SE ACTUALIZO LAS FACTURAS SOLICITADAS")
        ACTUALIZAR()
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PAGO_PARCIAL" Then
            Try
                Dim va, va1 As Double
                va = DataGridView1.Rows(e.RowIndex).Cells(16).Value
                va1 = Format(va, "##,##00.00")

                If va1 > DataGridView1.Rows(e.RowIndex).Cells(17).Value Then
                    DataGridView1.Rows(e.RowIndex).Cells(16).Value = 0.00

                    MsgBox("LA CANTIDAD INGRESADA EXEDE A LO QUE DEBE EL CLIENTE")
                Else
                    If va1 = DataGridView1.Rows(e.RowIndex).Cells(17).Value Then
                        DataGridView1.Rows(e.RowIndex).Cells(2).Value = True
                        MsgBox("CON LA CANTIDAD INGRESADA SE CANCELARA TODA LA FACTURA")
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(2).Value = False
                        TextBox8.Text = DataGridView1.Rows(e.RowIndex).Cells(14).Value - DataGridView1.Rows(e.RowIndex).Cells(16).Value
                    End If
                End If
            Catch ex As Exception
                MsgBox("FALTA INGRESAR EL PAGO PARCIAL")
            End Try

        End If
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If RadioButton1.Checked = True Then
            Try
                Dim gh As New fregistro_venta
                Dim gh1 As New vregistroventa

                Button1.Visible = False
                gh1.ganal1_3 = 1
                gh1.gempresa = Label9.Text
                dt = gh.mostrar_venta(gh1)
                RadioButton2.Checked = True
                DataGridView1.DataSource = dt
                DataGridView1.Columns(4).Width = 250
                DataGridView1.Columns(5).Width = 60
                DataGridView1.Columns(6).Width = 200
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(5).Frozen = True
                DataGridView1.Columns(6).Frozen = True
                DataGridView1.Columns(7).Frozen = True
                DataGridView1.Columns(15).ReadOnly = True
                DataGridView1.Columns(6).DefaultCellStyle.BackColor = Color.Wheat
                DataGridView1.Columns(13).DefaultCellStyle.BackColor = Color.NavajoWhite
                DataGridView1.Columns(19).Visible = False
                Dim i As Integer

                i = DataGridView1.Rows.Count
                For t = 0 To i - 1
                    DataGridView1.Rows(t).Cells(17).Value = DataGridView1.Rows(t).Cells(14).Value - DataGridView1.Rows(t).Cells(15).Value
                    If DataGridView1.Rows(t).Cells(15).Value > "0.00" Then
                        DataGridView1.Rows(t).DefaultCellStyle.BackColor = Color.Wheat
                    End If
                    Select Case DataGridView1.Rows(t).Cells(9).Value
                        Case "0010" : DataGridView1.Rows(t).Cells(9).Value = "GBEDON"
                        Case "0022" : DataGridView1.Rows(t).Cells(9).Value = "VINCIO"
                        Case "0023" : DataGridView1.Rows(t).Cells(9).Value = "DBRAVO"
                        Case "0025" : DataGridView1.Rows(t).Cells(9).Value = "JSALINAS"
                        Case "0029" : DataGridView1.Rows(t).Cells(9).Value = "GCUEVA"
                        Case "0026" : DataGridView1.Rows(t).Cells(9).Value = "AMENDO"
                        Case "0007" : DataGridView1.Rows(t).Cells(9).Value = "VGRAUS"
                        Case "0027" : DataGridView1.Rows(t).Cells(9).Value = "VPIZARRO"
                        Case "0028" : DataGridView1.Rows(t).Cells(9).Value = "JBALVIN"
                        Case "0005" : DataGridView1.Rows(t).Cells(9).Value = "VSILVERIO"
                        Case "0030" : DataGridView1.Rows(t).Cells(9).Value = "RMEDINA"
                    End Select
                    Select Case DataGridView1.Rows(t).Cells(8).Value
                        Case "01" : DataGridView1.Rows(t).Cells(8).Value = "CONTADO"
                        Case "02" : DataGridView1.Rows(t).Cells(8).Value = "CRED 07 DIAS"
                        Case "03" : DataGridView1.Rows(t).Cells(8).Value = "CRED 15 DIAS"
                        Case "04" : DataGridView1.Rows(t).Cells(8).Value = "CRED 30 DIAS"
                        Case "05" : DataGridView1.Rows(t).Cells(8).Value = "CRED 45 DIAS"
                        Case "06" : DataGridView1.Rows(t).Cells(8).Value = "CRED 60 DIAS"
                        Case "07" : DataGridView1.Rows(t).Cells(8).Value = "CRED 90 DIAS"
                        Case "08" : DataGridView1.Rows(t).Cells(8).Value = "CRED 120 DIAS"
                        Case "09" : DataGridView1.Rows(t).Cells(8).Value = "CONTRA ENTREGA"

                    End Select
                Next
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(3).Visible = False
                RadioButton1.Checked = True
            Catch ex As Exception
                MsgBox("NO HAY REGISTROS DE VENTAS NUEVAS INGREADAS")
            End Try
        Else
            If RadioButton2.Checked = True Then
                ACTUALIZAR()
            End If
        End If
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 18 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            observaciones_vn.TextBox2.Text = num1
            observaciones_vn.TextBox3.Text = num2
            observaciones_vn.TextBox1.Text = DataGridView1.Rows(num1).Cells(18).Value
            observaciones_vn.Show()
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        buscar_FPAGO()
    End Sub
End Class