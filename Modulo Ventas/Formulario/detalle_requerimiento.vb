Imports System.Data.SqlClient
Public Class detalle_requerimiento
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub detalle_requerimiento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT id_requeminieto_detalle AS ID,id_cabecera AS CABECERA,items AS ITEMS, linea AS CODIGO,detalle AS DETALLE,factor AS FACTOR,UM AS UM, cantidad AS CANTIDAD, TOTAL AS TOTAL,estado as ESTADO,estado1 FROM requerimiento_avios_detalle WHERE id_cabecera ='" + Label1.Text + "' ORDER BY linea", conx)
        Dim da As New DataTable
        cmd.Fill(da)
        DataGridView1.DataSource = da
        'DataGridView1.Columns(1).Visible = False
        'DataGridView1.Columns(2).Visible = False
        'DataGridView1.Columns(3).Visible = False
        'DataGridView1.Columns(10).Visible = False
        'DataGridView1.Columns(9).Visible = False

        DataGridView1.Columns(1).ReadOnly = True
        DataGridView1.Columns(2).ReadOnly = True
        DataGridView1.Columns(3).ReadOnly = True
        DataGridView1.Columns(4).ReadOnly = True
        DataGridView1.Columns(5).ReadOnly = True
        DataGridView1.Columns(6).ReadOnly = True
        DataGridView1.Columns(7).ReadOnly = True
        DataGridView1.Columns(4).Width = 135
        DataGridView1.Columns(5).Width = 400
        DataGridView1.Columns(6).Visible = False
        DataGridView1.Columns(8).Visible = False
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(3).Visible = False
        DataGridView1.Columns(10).Visible = False

        DataGridView1.Columns(11).Visible = False
        Dim o As Integer

        o = DataGridView1.Rows.Count
        If Label3.Text = 1 Then
            For i = 0 To o - 1

                If Trim(DataGridView1.Rows(i).Cells(10).Value) = "2" Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    DataGridView1.Rows(i).ReadOnly = True
                    End If

            Next
        Else
            If Label3.Text = "2" Then

                For i1 = 0 To o - 1
                    If Trim(DataGridView1.Rows(i1).Cells(11).Value) = "1" Then
                        DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                        DataGridView1.Rows(i1).ReadOnly = True
                    Else
                        If Trim(DataGridView1.Rows(i1).Cells(11).Value) = "2" Then

                            DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Black
                            DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                            DataGridView1.Rows(i1).ReadOnly = True
                        End If
                    End If

                Next



            End If
        End If



    End Sub
    Dim Rsr2127 As SqlDataReader
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim p, va As Integer
        va = 0
        p = DataGridView1.Rows.Count
        If Label3.Text = 1 Then
            For i = 0 To p - 1
                If DataGridView1.Rows(i).Cells(0).Value = True Then

                    If Trim(Convert.ToString(Agregar_requerimiento.DataGridView1.Rows(Label2.Text).Cells(11).Value)) = "" Then
                        Agregar_requerimiento.DataGridView1.Rows(Label2.Text).Cells(11).Value = Convert.ToString(DataGridView1.Rows(i).Cells(1).Value)
                    Else
                        Agregar_requerimiento.DataGridView1.Rows(Label2.Text).Cells(11).Value = Convert.ToString(Agregar_requerimiento.DataGridView1.Rows(Label2.Text).Cells(11).Value) + "," + Convert.ToString(DataGridView1.Rows(i).Cells(1).Value)
                    End If
                    abrir()
                    Dim cmd157 As New SqlCommand("update requerimiento_avios_detalle set estado = 1 where id_cabecera=@id_cabecera and id_requeminieto_detalle =@id_requeminieto_detalle ", conx)
                    cmd157.Parameters.AddWithValue("@id_cabecera", Trim(DataGridView1.Rows(i).Cells(2).Value))
                    cmd157.Parameters.AddWithValue("@id_requeminieto_detalle", Trim(DataGridView1.Rows(i).Cells(1).Value))
                    cmd157.ExecuteNonQuery()
                End If

            Next
            Agregar_requerimiento.DataGridView1.Rows(Label2.Text).Cells(0).Value = True
            Me.Close()
        Else
            If Label3.Text = 2 Then
                Dim ol As String
                Dim sql1021 As String = "SELECT max(correlativo) as Ma FROM requerimiento_avios_detalle where id_cabecera ='" + Trim(Label1.Text) + "'"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr2127 = cmd1021.ExecuteReader()
                If Rsr2127.Read Then
                    ol = Rsr2127(0) + 1

                End If
                Rsr2127.Close()

                For i = 0 To p - 1
                    If DataGridView1.Rows(i).Cells(0).Value = True Then


                        abrir()
                        Dim cmd16 As New SqlCommand("UPDATE requerimiento_avios_cabecera SET estado='1',farmado =@famado where id_requeminieto =@id_requeminieto", conx)
                        cmd16.Parameters.AddWithValue("@id_requeminieto", Trim(DataGridView1.Rows(i).Cells(2).Value))
                        cmd16.Parameters.AddWithValue("@famado", DateTimePicker1.Value)
                        cmd16.ExecuteNonQuery()




                        Dim cmd199 As New SqlCommand("update requerimiento_avios_detalle set estado1 =1,correlativo = '" + ol + "' where id_requeminieto_detalle =@id_requeminietodet", conx)
                        cmd199.Parameters.AddWithValue("@id_requeminietodet", DataGridView1.Rows(i).Cells(1).Value)
                        cmd199.ExecuteNonQuery()

                    End If

                Next

                'If p = va Then
                '    modulo_solicitud.DataGridView1.Rows(Label2.Text).Cells(1).Value = True
                'End If
                'modulo_solicitud.DataGridView1.Rows(Label2.Text).Cells(2).Value = True

                Me.Close()
            End If



        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim k As Integer
        k = DataGridView1.Rows.Count
        If CheckBox1.Checked = True Then
            For i = 0 To k - 1
                DataGridView1.Rows(i).Cells(0).Value = True
            Next
        Else
            For i = 0 To k - 1
                DataGridView1.Rows(i).Cells(0).Value = False
            Next
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim p As Integer

        p = DataGridView1.Rows.Count
        'For i = 0 To p - 1
        If Label4.Text <> "a" Then

            Dim cmd157 As New SqlCommand("update requerimiento_avios_detalle set estado1 = 0 where id_cabecera=@id_cabecera and id_requeminieto_detalle =@id_requeminieto_detalle ", conx)
            cmd157.Parameters.AddWithValue("@id_cabecera", Trim(DataGridView1.Rows(Label4.Text).Cells(2).Value))
            cmd157.Parameters.AddWithValue("@id_requeminieto_detalle", Trim(DataGridView1.Rows(Label4.Text).Cells(1).Value))
            cmd157.ExecuteNonQuery()
            MsgBox("SE ACTUALIZO LA INFORMACION")
            abrir()
            Dim cmd As New SqlDataAdapter("SELECT id_requeminieto_detalle AS ID,id_cabecera AS CABECERA,items AS ITEMS, linea AS CODIGO,detalle AS DETALLE,factor AS FACTOR,UM AS UM, cantidad AS CANTIDAD, TOTAL AS TOTAL,estado as ESTADO,estado1 FROM requerimiento_avios_detalle WHERE id_cabecera ='" + Label1.Text + "' ORDER BY ITEMS", conx)
            Dim da As New DataTable
            cmd.Fill(da)
            DataGridView1.DataSource = da
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(8).Visible = False

            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(10).Visible = False

            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).ReadOnly = True
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(4).Width = 135
            DataGridView1.Columns(5).Width = 400
            Dim o As Integer

            o = DataGridView1.Rows.Count
            If Label3.Text = 1 Then
                For i = 0 To o - 1
                    If Trim(DataGridView1.Rows(i).Cells(10).Value) = "1" Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Black
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        DataGridView1.Rows(i).ReadOnly = True
                    Else
                        If Trim(DataGridView1.Rows(i).Cells(10).Value) = "2" Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                            DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                            DataGridView1.Rows(i).ReadOnly = True
                        End If
                    End If
                Next
            Else
                If Label3.Text = "2" Then

                    For i1 = 0 To o - 1
                        If Trim(DataGridView1.Rows(i1).Cells(11).Value) = "1" Then

                            DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
                            DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                            DataGridView1.Rows(i1).ReadOnly = True
                        Else
                            If Trim(DataGridView1.Rows(i1).Cells(11).Value) = "2" Then

                                DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Black
                                DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                                DataGridView1.Rows(i1).ReadOnly = True
                            End If
                        End If

                    Next



                End If
            End If

        End If
        'Next

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Label4.Text = e.RowIndex


        If e.RowIndex = -1 Then
            For i = 0 To DataGridView1.Columns.Count - 1

                DataGridView1.Columns.Item(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            Dim o As Integer

            o = DataGridView1.Rows.Count

            If Label3.Text = 1 Then
                For i = 0 To o - 1

                    If Trim(DataGridView1.Rows(i).Cells(10).Value) = "2" Then

                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                        DataGridView1.Rows(i).ReadOnly = True
                    End If

                Next
            Else
                If Label3.Text = "2" Then

                    For i1 = 0 To o - 1

                        If Trim(DataGridView1.Rows(i1).Cells(11).Value) = "1" Then


                            DataGridView1.Rows(i1).ReadOnly = True
                            DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Red
                            DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                        End If

                    Next



                End If
            End If
        End If
    End Sub
End Class