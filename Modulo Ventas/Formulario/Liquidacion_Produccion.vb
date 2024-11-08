Imports System.Data.SqlClient
Imports System.ComponentModel

Public Class Liquidacion_Produccion
    Dim ll As New DataTable
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
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
    Private Sub Liquidacion_Produccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub


    Dim Rsr2, Rsr299 As SqlDataReader

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim jk As New fop
        Dim kk As New vop
        DataGridView1.DataSource = Nothing
        kk.gcia = Label3.Text
        kk.gncom_3 = Trim(TextBox1.Text)
        ll = jk.liquidar_produccion(kk)

        If ll Is Nothing Then
            MsgBox("NO HAY REGISTROS")
        Else

            If ll.Rows.Count <> 0 Then
                DataGridView1.DataSource = ll
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(2).Width = 85
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 90
                DataGridView1.Columns(5).Width = 350
                DataGridView1.Columns(6).Width = 150
                DataGridView1.Columns(7).Width = 180
                DataGridView1.Columns(9).Width = 85
                DataGridView1.Columns(10).Width = 115
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(12).Visible = False
                Dim op As Integer
                op = DataGridView1.Rows.Count
                For i = 0 To op - 1

                    If Trim(DataGridView1.Rows(i).Cells(1).Value) = "CORTE" Then

                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Silver
                        'DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                            DataGridView1.Item(1, i).Style.BackColor = Color.Silver
                            DataGridView1.Item(6, i).Style.BackColor = Color.Silver
                            DataGridView1.Item(6, i).Style.BackColor = Color.Silver
                            DataGridView1.Item(9, i).Style.BackColor = Color.Silver
                            DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                            If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                                DataGridView1.Item(8, i).Style.BackColor = Color.Red
                                DataGridView1.Item(8, i).Style.ForeColor = Color.White
                            End If

                        End If
                        If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                            DataGridView1.Item(7, i).Style.BackColor = Color.Black
                            DataGridView1.Item(7, i).Style.ForeColor = Color.White
                        End If
                        If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                            DataGridView1.Item(10, i).Style.BackColor = Color.Black
                            DataGridView1.Item(10, i).Style.ForeColor = Color.White
                        End If
                    Else
                        If Trim(DataGridView1.Rows(i).Cells(1).Value) = "CONFECCION" Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Bisque
                            If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                                DataGridView1.Item(1, i).Style.BackColor = Color.Bisque
                                DataGridView1.Item(6, i).Style.BackColor = Color.Bisque
                                DataGridView1.Item(7, i).Style.BackColor = Color.Bisque
                                DataGridView1.Item(9, i).Style.BackColor = Color.Bisque
                                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                                If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                                    DataGridView1.Item(8, i).Style.BackColor = Color.Red
                                    DataGridView1.Item(8, i).Style.ForeColor = Color.White
                                End If

                            End If
                            If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                                DataGridView1.Item(7, i).Style.BackColor = Color.Black
                                DataGridView1.Item(7, i).Style.ForeColor = Color.White
                            End If
                            If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                                DataGridView1.Item(10, i).Style.BackColor = Color.Black
                                DataGridView1.Item(10, i).Style.ForeColor = Color.White
                            End If
                        Else
                            If Trim(DataGridView1.Rows(i).Cells(1).Value) = "LAVANDERIA" Then
                                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkSlateBlue
                                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                                    DataGridView1.Item(1, i).Style.BackColor = Color.DarkSlateBlue
                                    DataGridView1.Item(3, i).Style.BackColor = Color.DarkSlateBlue
                                    DataGridView1.Item(6, i).Style.BackColor = Color.DarkSlateBlue
                                    DataGridView1.Item(7, i).Style.BackColor = Color.DarkSlateBlue
                                    DataGridView1.Item(9, i).Style.BackColor = Color.DarkSlateBlue
                                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                    If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                                        DataGridView1.Item(8, i).Style.BackColor = Color.Red
                                        DataGridView1.Item(8, i).Style.ForeColor = Color.White
                                    Else
                                        DataGridView1.Item(8, i).Style.ForeColor = Color.Black
                                    End If

                                End If
                                If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                                    DataGridView1.Item(7, i).Style.BackColor = Color.Black
                                    DataGridView1.Item(7, i).Style.ForeColor = Color.White
                                End If
                                If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                                    DataGridView1.Item(10, i).Style.BackColor = Color.Black
                                    DataGridView1.Item(10, i).Style.ForeColor = Color.White
                                End If
                            Else
                                If Trim(DataGridView1.Rows(i).Cells(1).Value) = "ACABADOS" Then
                                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Peru
                                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                    If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                                        DataGridView1.Item(1, i).Style.BackColor = Color.Peru
                                        DataGridView1.Item(3, i).Style.BackColor = Color.Peru
                                        DataGridView1.Item(6, i).Style.BackColor = Color.Peru
                                        DataGridView1.Item(7, i).Style.BackColor = Color.Peru
                                        DataGridView1.Item(9, i).Style.BackColor = Color.Peru
                                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                        If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                                            DataGridView1.Item(8, i).Style.BackColor = Color.Red
                                            DataGridView1.Item(8, i).Style.ForeColor = Color.White
                                        Else
                                            DataGridView1.Item(8, i).Style.ForeColor = Color.Black
                                        End If

                                    End If
                                    If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                                        DataGridView1.Item(7, i).Style.BackColor = Color.Black
                                        DataGridView1.Item(7, i).Style.ForeColor = Color.White
                                    End If
                                    If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                                        DataGridView1.Item(10, i).Style.BackColor = Color.Black
                                        DataGridView1.Item(10, i).Style.ForeColor = Color.White
                                    End If
                                Else
                                    If Trim(DataGridView1.Rows(i).Cells(1).Value) = "APLICACIONES Y PIEZAS" Then
                                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Aqua
                                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                                        If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Black
                                            DataGridView1.Item(1, i).Style.BackColor = Color.Aqua
                                            DataGridView1.Item(3, i).Style.BackColor = Color.Aqua
                                            DataGridView1.Item(6, i).Style.BackColor = Color.Aqua
                                            DataGridView1.Item(7, i).Style.BackColor = Color.Aqua
                                            DataGridView1.Item(9, i).Style.BackColor = Color.Aqua
                                            DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                                            If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                                                DataGridView1.Item(8, i).Style.BackColor = Color.Red
                                                DataGridView1.Item(8, i).Style.ForeColor = Color.Black
                                            Else
                                                DataGridView1.Item(8, i).Style.ForeColor = Color.Black
                                            End If

                                        End If
                                        If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                                            DataGridView1.Item(7, i).Style.BackColor = Color.Black
                                            DataGridView1.Item(7, i).Style.ForeColor = Color.White
                                        End If
                                        If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                                            DataGridView1.Item(10, i).Style.BackColor = Color.Black
                                            DataGridView1.Item(10, i).Style.ForeColor = Color.White
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If



                Next


            Else
                MsgBox("LA OP NO EXISTE")
                DataGridView1.DataSource = Nothing
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub
    Public Sub NumConFrac(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False

        Else
            e.Handled = True
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.DataSource = Nothing
        TextBox1.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim jk As New fop
        Dim kk As New vop
        kk.gcia = Label3.Text
        kk.gncom_3 = Trim(DataGridView1.Rows(0).Cells(9).Value)
        ll = jk.liquidar_produccion(kk)
        DataGridView1.DataSource = Nothing
        If ll.Rows.Count <> 0 Then
            DataGridView1.DataSource = ll
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(2).Width = 85
            DataGridView1.Columns(3).Width = 85
            DataGridView1.Columns(4).Width = 90
            DataGridView1.Columns(5).Width = 350
            DataGridView1.Columns(6).Width = 150
            DataGridView1.Columns(7).Width = 180
            DataGridView1.Columns(9).Width = 85
            DataGridView1.Columns(10).Width = 115
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            Dim op As Integer
            op = DataGridView1.Rows.Count
            For i = 0 To op - 1

                If Trim(DataGridView1.Rows(i).Cells(1).Value) = "CORTE" Then

                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Silver
                    'DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                        DataGridView1.Item(1, i).Style.BackColor = Color.Silver
                        DataGridView1.Item(6, i).Style.BackColor = Color.Silver
                        DataGridView1.Item(9, i).Style.BackColor = Color.Silver
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                        If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                            DataGridView1.Item(8, i).Style.BackColor = Color.Red
                            DataGridView1.Item(8, i).Style.ForeColor = Color.White
                        End If

                    End If
                    If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                        DataGridView1.Item(7, i).Style.BackColor = Color.Black
                        DataGridView1.Item(7, i).Style.ForeColor = Color.White
                    End If
                    If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                        DataGridView1.Item(10, i).Style.BackColor = Color.Black
                        DataGridView1.Item(10, i).Style.ForeColor = Color.White
                    End If
                Else
                    If Trim(DataGridView1.Rows(i).Cells(1).Value) = "CONFECCION" Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Bisque
                        If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                            DataGridView1.Item(1, i).Style.BackColor = Color.Bisque
                            DataGridView1.Item(6, i).Style.BackColor = Color.Bisque
                            DataGridView1.Item(9, i).Style.BackColor = Color.Bisque
                            DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                            If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                                DataGridView1.Item(8, i).Style.BackColor = Color.Red
                                DataGridView1.Item(8, i).Style.ForeColor = Color.White
                            End If

                        End If
                        If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                            DataGridView1.Item(7, i).Style.BackColor = Color.Black
                            DataGridView1.Item(7, i).Style.ForeColor = Color.White
                        End If
                        If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                            DataGridView1.Item(10, i).Style.BackColor = Color.Black
                            DataGridView1.Item(10, i).Style.ForeColor = Color.White
                        End If
                    Else
                        If Trim(DataGridView1.Rows(i).Cells(1).Value) = "LAVANDERIA" Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkSlateBlue
                            DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                            If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                                DataGridView1.Item(1, i).Style.BackColor = Color.DarkSlateBlue
                                DataGridView1.Item(3, i).Style.BackColor = Color.DarkSlateBlue
                                DataGridView1.Item(6, i).Style.BackColor = Color.DarkSlateBlue
                                DataGridView1.Item(9, i).Style.BackColor = Color.DarkSlateBlue
                                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                                    DataGridView1.Item(8, i).Style.BackColor = Color.Red
                                    DataGridView1.Item(8, i).Style.ForeColor = Color.White
                                Else
                                    DataGridView1.Item(8, i).Style.ForeColor = Color.Black
                                End If

                            End If
                            If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                                DataGridView1.Item(7, i).Style.BackColor = Color.Black
                                DataGridView1.Item(7, i).Style.ForeColor = Color.White
                            End If
                            If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                                DataGridView1.Item(10, i).Style.BackColor = Color.Black
                                DataGridView1.Item(10, i).Style.ForeColor = Color.White
                            End If
                        Else
                            If Trim(DataGridView1.Rows(i).Cells(1).Value) = "ACABADOS" Then
                                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Peru
                                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                                    DataGridView1.Item(1, i).Style.BackColor = Color.Peru
                                    DataGridView1.Item(3, i).Style.BackColor = Color.Peru
                                    DataGridView1.Item(6, i).Style.BackColor = Color.Peru
                                    DataGridView1.Item(9, i).Style.BackColor = Color.Peru
                                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                    If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                                        DataGridView1.Item(8, i).Style.BackColor = Color.Red
                                        DataGridView1.Item(8, i).Style.ForeColor = Color.White
                                    Else
                                        DataGridView1.Item(8, i).Style.ForeColor = Color.Black
                                    End If

                                End If
                                If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                                    DataGridView1.Item(7, i).Style.BackColor = Color.Black
                                    DataGridView1.Item(7, i).Style.ForeColor = Color.White
                                End If
                                If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                                    DataGridView1.Item(10, i).Style.BackColor = Color.Black
                                    DataGridView1.Item(10, i).Style.ForeColor = Color.White
                                End If
                            Else
                                If Trim(DataGridView1.Rows(i).Cells(1).Value) = "APLICACIONES Y PIEZAS" Then
                                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Aqua
                                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                    If Trim(DataGridView1.Rows(i).Cells(3).Value) = "SALDO" Then
                                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                                        DataGridView1.Item(1, i).Style.BackColor = Color.Aqua
                                        DataGridView1.Item(3, i).Style.BackColor = Color.Aqua
                                        DataGridView1.Item(6, i).Style.BackColor = Color.Aqua
                                        DataGridView1.Item(7, i).Style.BackColor = Color.Aqua
                                        DataGridView1.Item(9, i).Style.BackColor = Color.Aqua
                                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                                        If Convert.ToInt32(DataGridView1.Rows(i).Cells(8).Value) < 0 Then
                                            DataGridView1.Item(8, i).Style.BackColor = Color.Red
                                            DataGridView1.Item(8, i).Style.ForeColor = Color.White
                                        Else
                                            DataGridView1.Item(8, i).Style.ForeColor = Color.Black
                                        End If

                                    End If
                                    If Trim(DataGridView1.Rows(i).Cells(11).Value) = "1" Then
                                        DataGridView1.Item(7, i).Style.BackColor = Color.Black
                                        DataGridView1.Item(7, i).Style.ForeColor = Color.White
                                    End If
                                    If Trim(DataGridView1.Rows(i).Cells(12).Value) = "1" Then
                                        DataGridView1.Item(10, i).Style.BackColor = Color.Black
                                        DataGridView1.Item(10, i).Style.ForeColor = Color.White
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If



            Next


        Else
            MsgBox("LA OP NO EISTE")
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            NumConFrac(TextBox1, e)
        Catch ex As Exception
            MsgBox("FALTA INGRESAR UN NUMERO")
        End Try
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Trim(DataGridView1.Columns(e.ColumnIndex).HeaderText) = "NIA_NSA" Then
            If Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) <> "" Then
                abrir()
                Dim hj, VA As String
                Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                    Case "CORTE" : hj = "0300"
                    Case "CONFECCION" : hj = "0400"
                    Case "ACABADOS" : hj = "0500"
                    Case "LAVANDERIA" : hj = "0600"
                    Case "APLICACIONES Y PIEZAS" : hj = "0700"
                End Select
                Dim sql102 As String = "SELECT revisado_3 FROM custom_vianny.dbo.Qag3000 WHERE ncom_3 ='" + Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 1, 8) + "' and Ano_3 ='" + Label2.Text + "' and talm_3 ='" + hj + "' and Empr_3 ='" + Label3.Text + "'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()
                Dim i5 As Integer
                i5 = DataGridView1.Rows.Count
                'MsgBox(Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 1, 8))
                'MsgBox(hj)
                'MsgBox(Label2.Text)
                'MsgBox(Label3.Text)
                If Rsr2.Read() = True Then
                    VA = Rsr2(0)
                End If
                Rsr2.Close()
                If VA = "0" Then

                    Dim cmd20 As New SqlCommand("update custom_vianny.dbo.Qag3000 set revisado_3 ='1'  WHERE ncom_3 =@ncom_3 and Ano_3 =@Ano_3 and talm_3 =@talm_3 and Empr_3 =@Empr_3", conx)
                    cmd20.Parameters.AddWithValue("@Empr_3", Trim(Label3.Text))
                    cmd20.Parameters.AddWithValue("@Ano_3", Trim(Label2.Text))
                    cmd20.Parameters.AddWithValue("@ncom_3", Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 1, 8))
                    cmd20.Parameters.AddWithValue("@talm_3", hj)
                    cmd20.ExecuteNonQuery()

                    'MsgBox(Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 12, 4))
                    'MsgBox(Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 17, 8))

                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value) = "ENVIO" Then
                        Dim cmd21 As New SqlCommand(" UPDATE custom_vianny.dbo.fag0400 SET cerrado_3 ='1' WHERE sfactu_3 =@sfactu_3 AND nfactu_3 =@nfactu_3 AND CIA_3 =@CIA_3", conx)
                        cmd21.Parameters.AddWithValue("@CIA_3", Trim(Label3.Text))
                        cmd21.Parameters.AddWithValue("@sfactu_3", Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 12, 4))
                        cmd21.Parameters.AddWithValue("@nfactu_3", Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 17, 8))
                        cmd21.ExecuteNonQuery()
                    End If


                    MsgBox("SE CERRO EL COMPROBANTE")
                    Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                        Case "CORTE" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.Silver
                        Case "CONFECCION" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.Bisque
                        Case "ACABADOS" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.Peru
                        Case "LAVANDERIA" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.DarkSlateBlue

                    End Select
                    DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.Black
                    DataGridView1.Item(7, e.RowIndex).Style.ForeColor = Color.White
                    'DataGridView1.Rows(e.RowIndex).Cells(8).Selected = True
                Else

                    Dim cmd20 As New SqlCommand("update custom_vianny.dbo.Qag3000 set revisado_3 ='0'  WHERE ncom_3 =@ncom_3 and Ano_3 =@Ano_3 and talm_3 =@talm_3 and Empr_3 =@Empr_3", conx)
                    cmd20.Parameters.AddWithValue("@Empr_3", Trim(Label3.Text))
                    cmd20.Parameters.AddWithValue("@Ano_3", Trim(Label2.Text))
                    cmd20.Parameters.AddWithValue("@ncom_3", Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 1, 8))
                    cmd20.Parameters.AddWithValue("@talm_3", hj)
                    cmd20.ExecuteNonQuery()
                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value) = "ENVIO" Then
                        Dim cmd21 As New SqlCommand(" UPDATE custom_vianny.dbo.fag0400 SET cerrado_3 ='1' WHERE sfactu_3 =@sfactu_3 AND nfactu_3 =@nfactu_3 AND CIA_3 =@CIA_3", conx)
                        cmd21.Parameters.AddWithValue("@CIA_3", Trim(Label3.Text))
                        cmd21.Parameters.AddWithValue("@sfactu_3", Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 12, 4))
                        cmd21.Parameters.AddWithValue("@nfactu_3", Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 17, 8))
                        cmd21.ExecuteNonQuery()
                    End If
                    MsgBox("SE APERTURO EL COMPROBANTE")
                    Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                        Case "CORTE" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.Silver
                        Case "CONFECCION" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.Bisque
                        Case "ACABADOS" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.Peru
                        Case "LAVANDERIA" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.DarkSlateBlue
                        Case "APLICACIONES Y PIEZAS" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.White
                    End Select

                    Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                        Case "CORTE" : DataGridView1.Item(7, e.RowIndex).Style.ForeColor = Color.Black
                        Case "CONFECCION" : DataGridView1.Item(7, e.RowIndex).Style.ForeColor = Color.Black
                        Case "ACABADOS" : DataGridView1.Item(7, e.RowIndex).Style.ForeColor = Color.White
                        Case "LAVANDERIA" : DataGridView1.Item(7, e.RowIndex).Style.ForeColor = Color.White
                        Case "APLICACIONES Y PIEZAS" : DataGridView1.Item(7, e.RowIndex).Style.BackColor = Color.Red
                    End Select

                    'DataGridView1.Item(7, e.RowIndex).Style.ForeColor = Color.Black

                    'DataGridView1.Rows(e.RowIndex).Cells(8).Selected = True
                End If
            Else
                Agregar_Factura.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value)
                Agregar_Factura.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                Agregar_Factura.Label4.Text = e.RowIndex
                Agregar_Factura.Label5.Text = Label3.Text
                Agregar_Factura.Show()
            End If


        End If
        If Trim(DataGridView1.Columns(e.ColumnIndex).HeaderText) = "ORDEN_SERVICIO" Then
            abrir()
            Dim hj As String
            Dim VAw As Integer

            Dim sql10299 As String = "SELECT cerrado_4 FROM custom_vianny.dbo.lag0400 where ncom_4 ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value) + "' AND ccia_4 ='" + Trim(Label3.Text) + "' "
            Dim cmd10299 As New SqlCommand(sql10299, conx)
            Rsr299 = cmd10299.ExecuteReader()
            Dim i5 As Integer
            i5 = DataGridView1.Rows.Count

            If Rsr299.Read() = True Then
                VAw = Rsr299(0)
            End If
            Rsr299.Close()

            If VAw = "0" Then
                Dim cmd2090 As New SqlCommand("UPDATE custom_vianny.dbo.lag0400 SET cerrado_4 ='1' WHERE ncom_4 =@ncom_4 AND ccia_4 =@ccia_4", conx)
                cmd2090.Parameters.AddWithValue("@ccia_4", Trim(Label3.Text))
                cmd2090.Parameters.AddWithValue("@ncom_4", Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value))
                cmd2090.ExecuteNonQuery()

                'MsgBox(Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 12, 4))
                'MsgBox(Mid(Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value), 17, 8))

                MsgBox("SE CERRO EL COMPROBANTE")
                Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                    Case "CORTE" : DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.Silver
                    Case "CONFECCION" : DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.Bisque
                    Case "ACABADOS" : DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.Peru
                    Case "LAVANDERIA" : DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.DarkSlateBlue

                End Select

                For i = 0 To i5 - 1
                    'MsgBox(Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value))
                    'MsgBox(Trim(DataGridView1.Rows(i).Cells(10).Value))
                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value) = Trim(DataGridView1.Rows(i).Cells(10).Value) Then
                        'MsgBox("paso")
                        DataGridView1.Item(10, i).Style.BackColor = Color.Black
                        DataGridView1.Item(10, i).Style.ForeColor = Color.White
                    End If
                Next
                DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.Black
                DataGridView1.Item(10, e.RowIndex).Style.ForeColor = Color.White

                'DataGridView1.Rows(e.RowIndex).Cells(8).Selected = True
            Else
                Dim cmd2098 As New SqlCommand("UPDATE custom_vianny.dbo.lag0400 SET cerrado_4 ='0' WHERE ncom_4 =@ncom_4 AND ccia_4 =@ccia_4", conx)
                cmd2098.Parameters.AddWithValue("@ccia_4", Trim(Label3.Text))
                cmd2098.Parameters.AddWithValue("@ncom_4", Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value))
                cmd2098.ExecuteNonQuery()
                MsgBox("SE APERTURO EL COMPROBANTE")

                For i = 0 To i5 - 1
                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(10).Value) = Trim(DataGridView1.Rows(i).Cells(10).Value) Then
                        Select Case Trim(DataGridView1.Rows(i).Cells(1).Value)
                            Case "CORTE" : DataGridView1.Item(10, i).Style.BackColor = Color.Silver

                            Case "CONFECCION" : DataGridView1.Item(10, i).Style.BackColor = Color.Bisque

                            Case "ACABADOS" : DataGridView1.Item(10, i).Style.BackColor = Color.Peru

                            Case "LAVANDERIA" : DataGridView1.Item(10, i).Style.BackColor = Color.DarkSlateBlue

                        End Select
                        Select Case Trim(DataGridView1.Rows(i).Cells(1).Value)

                            Case "CORTE" : DataGridView1.Item(10, i).Style.ForeColor = Color.Black

                            Case "CONFECCION" : DataGridView1.Item(10, i).Style.ForeColor = Color.Black

                            Case "ACABADOS" : DataGridView1.Item(10, i).Style.ForeColor = Color.White

                            Case "LAVANDERIA" : DataGridView1.Item(10, i).Style.ForeColor = Color.White
                        End Select

                    End If
                Next

                'Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                '    Case "CORTE" : DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.Silver
                '    Case "CORTE" : DataGridView1.Item(10, e.RowIndex).Style.ForeColor = Color.Black
                '    Case "CONFECCION" : DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.Bisque
                '    Case "CONFECCION" : DataGridView1.Item(10, e.RowIndex).Style.ForeColor = Color.Black
                '    Case "ACABADOS" : DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.Peru
                '    Case "ACABADOS" : DataGridView1.Item(10, e.RowIndex).Style.ForeColor = Color.White
                '    Case "LAVANDERIA" : DataGridView1.Item(10, e.RowIndex).Style.BackColor = Color.DarkSlateBlue
                '    Case "LAVANDERIA" : DataGridView1.Item(10, e.RowIndex).Style.ForeColor = Color.White
                'End Select
                'DataGridView1.Rows(e.RowIndex).Cells(8).Selected = True
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox1.Text).Length
                Case "1" : TextBox1.Text = "01" & "0000000" & TextBox1.Text
                Case "2" : TextBox1.Text = "01" & "000000" & TextBox1.Text
                Case "3" : TextBox1.Text = "01" & "00000" & TextBox1.Text
                Case "4" : TextBox1.Text = "01" & "0000" & TextBox1.Text
                Case "5" : TextBox1.Text = "01" & "000" & TextBox1.Text
                Case "6" : TextBox1.Text = "01" & "00" & TextBox1.Text
                Case "7" : TextBox1.Text = "01" & "0" & TextBox1.Text
                Case "8" : TextBox1.Text = "01" & TextBox1.Text

            End Select
        End If

    End Sub
End Class