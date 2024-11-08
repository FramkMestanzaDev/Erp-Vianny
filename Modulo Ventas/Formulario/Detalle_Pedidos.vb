Imports System.Data.SqlClient
Public Class Detalle_Pedidos
    Dim dt, dt3, dt4, dt5, dt6 As New DataTable
    Public conx As SqlConnection
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

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Form8.TextBox1.Text = Mid(Trim(DataGridView1.Item(9, num1).Value()), 1, 4)
            Form8.TextBox2.Text = Mid(Trim(DataGridView1.Item(9, num1).Value()), 5, 8)
            Form8.TextBox3.Text = Trim(DataGridView1.Item(10, num1).Value())
            Form8.TextBox4.Text = Label3.Text
            Form8.ShowDialog()
            'Reporte_Np.TextBox1.Text = DataGridView1.Item(3, num1).Value()

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar2()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar3()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ds As New DataSet
        ds.Tables.Add(dt.Copy)
        Dim dv As New DataView(ds.Tables(0))
        TextBox1.Text = DateTimePicker1.Value

        dv.RowFilter = "FECHA" & " >= '" & TextBox1.Text & "'"

        If dv.Count <> 0 Then

            DataGridView1.DataSource = dv



        Else

            DataGridView1.DataSource = Nothing
        End If
        Dim J As Integer

        J = DataGridView1.Rows.Count

        Try
            If J > 0 Then
                DataGridView1.Columns(2).Width = 85
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 280

                For I = 0 To J - 1


                    If Trim(DataGridView1.Rows(I).Cells(5).Value) = "AUTORIZADO" Then

                        DataGridView1.Rows(I).Cells(5).Style.BackColor = Color.DarkCyan
                        DataGridView1.Rows(I).Cells(5).Style.ForeColor = Color.White
                    End If
                    If Trim(DataGridView1.Rows(I).Cells(6).Value) = "VALIDADO" Then

                        DataGridView1.Rows(I).Cells(6).Style.BackColor = Color.DarkKhaki
                        DataGridView1.Rows(I).Cells(6).Style.ForeColor = Color.White


                    End If
                    If Trim(DataGridView1.Rows(I).Cells(7).Value) = "DESPAC PROG." Then

                        DataGridView1.Rows(I).Cells(7).Style.BackColor = Color.DarkGray
                        DataGridView1.Rows(I).Cells(7).Style.ForeColor = Color.White


                    End If
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
    Dim Rsr2 As SqlDataReader
    Private Sub Detalle_Pedidos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim hs As New fnopedido
        Dim HJ As New vnapedido

        abrir()
        Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + TextBox4.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            HJ.gvendedor = Rsr2(0)
        End If

        'Select Case TextBox4.Text
        '    Case "GBEDON" : HJ.gvendedor = "0010"
        '    Case "VGRAUS" : HJ.gvendedor = "0007"
        '    Case "VINCIO" : HJ.gvendedor = "0022"
        '    Case "DBRAVO" : HJ.gvendedor = "0023"
        '    Case "VPIZARRO" : HJ.gvendedor = "0027"
        '    Case "GBALVIN" : HJ.gvendedor = "0028"
        '    Case "JSALINAS" : HJ.gvendedor = "0025"
        '    Case "AMENDO" : HJ.gvendedor = "0026"
        '    Case "GCUEVA" : HJ.gvendedor = "0029"
        '    Case "WSALINAS" : HJ.gvendedor = "0034"
        '    Case "VSILVERIO" : HJ.gvendedor = "0005"
        'End Select
        HJ.gempresa = Label3.Text
        HJ.gperiodo = Label4.Text
        dt = hs.buscar_nota_pedido_comercial(HJ)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(2).Width = 85
            DataGridView1.Columns(3).Width = 85
            DataGridView1.Columns(4).Width = 280
            Dim J As Integer

            J = DataGridView1.Rows.Count
            For I = 0 To J - 1
                If DataGridView1.Rows(I).Cells(5).Value = 1 Then
                    DataGridView1.Rows(I).Cells(5).Value = "AUTORIZADO"
                    DataGridView1.Rows(I).Cells(5).Style.BackColor = Color.DarkCyan
                    DataGridView1.Rows(I).Cells(5).Style.ForeColor = Color.White
                Else
                    If DataGridView1.Rows(I).Cells(5).Value = 0 Then
                        DataGridView1.Rows(I).Cells(5).Value = "PENDIENTE"
                    End If

                End If
                If DataGridView1.Rows(I).Cells(6).Value = 1 Then
                    DataGridView1.Rows(I).Cells(6).Value = "VALIDADO"
                    DataGridView1.Rows(I).Cells(6).Style.BackColor = Color.DarkKhaki
                    DataGridView1.Rows(I).Cells(6).Style.ForeColor = Color.White
                Else
                    If DataGridView1.Rows(I).Cells(6).Value = 0 Then
                        DataGridView1.Rows(I).Cells(6).Value = "PENDIENTE"
                    End If

                End If
                If DataGridView1.Rows(I).Cells(7).Value = 1 Then
                    DataGridView1.Rows(I).Cells(7).Value = "DESPAC PROG."
                    DataGridView1.Rows(I).Cells(7).Style.BackColor = Color.DarkGray
                    DataGridView1.Rows(I).Cells(7).Style.ForeColor = Color.White
                Else
                    If DataGridView1.Rows(I).Cells(7).Value = 0 Then
                        DataGridView1.Rows(I).Cells(7).Value = "PENDIENTE"
                    End If

                End If
            Next
        End If
    End Sub

    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv


            Else

                DataGridView1.DataSource = Nothing
            End If
            Dim J As Integer

            J = DataGridView1.Rows.Count

            Try
                If J > 0 Then
                    DataGridView1.Columns(2).Width = 85
                    DataGridView1.Columns(3).Width = 85
                    DataGridView1.Columns(4).Width = 280

                    For I = 0 To J - 1


                        If Trim(DataGridView1.Rows(I).Cells(5).Value) = "AUTORIZADO" Then

                            DataGridView1.Rows(I).Cells(5).Style.BackColor = Color.DarkCyan
                            DataGridView1.Rows(I).Cells(5).Style.ForeColor = Color.White
                        End If
                        If Trim(DataGridView1.Rows(I).Cells(6).Value) = "VALIDADO" Then

                            DataGridView1.Rows(I).Cells(6).Style.BackColor = Color.DarkKhaki
                            DataGridView1.Rows(I).Cells(6).Style.ForeColor = Color.White


                        End If
                        If Trim(DataGridView1.Rows(I).Cells(7).Value) = "DESPAC PROG." Then

                            DataGridView1.Rows(I).Cells(7).Style.BackColor = Color.DarkGray
                            DataGridView1.Rows(I).Cells(7).Style.ForeColor = Color.White


                        End If
                    Next
                End If
            Catch ex As Exception

            End Try
        Catch ex As Exception


        End Try

    End Sub
    Private Sub buscar3()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PEDIDO" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim J As Integer

                J = DataGridView1.Rows.Count

                Try
                    If J > 0 Then
                        DataGridView1.Columns(2).Width = 85
                        DataGridView1.Columns(3).Width = 85
                        DataGridView1.Columns(4).Width = 280

                        For I = 0 To J - 1


                            If Trim(DataGridView1.Rows(I).Cells(5).Value) = "AUTORIZADO" Then

                                DataGridView1.Rows(I).Cells(5).Style.BackColor = Color.DarkCyan
                                DataGridView1.Rows(I).Cells(5).Style.ForeColor = Color.White
                            End If
                            If Trim(DataGridView1.Rows(I).Cells(6).Value) = "VALIDADO" Then

                                DataGridView1.Rows(I).Cells(6).Style.BackColor = Color.DarkKhaki
                                DataGridView1.Rows(I).Cells(6).Style.ForeColor = Color.White


                            End If
                            If Trim(DataGridView1.Rows(I).Cells(7).Value) = "DESPAC PROG." Then

                                DataGridView1.Rows(I).Cells(7).Style.BackColor = Color.DarkGray
                                DataGridView1.Rows(I).Cells(7).Style.ForeColor = Color.White


                            End If
                        Next
                    End If
                Catch ex As Exception

                End Try



            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
End Class