Imports System.ComponentModel
Imports System.Data.SqlClient
Public Class Programa_Tejeduria
    Dim dt, dt1, dt2 As New DataTable
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
    Sub mostrar23()
        Dim JK As New fop
        dt = JK.PROGRAMA_TEJEDURIA()
        If Label3.Text = 2 Then
            Button1.Visible = False
        End If
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            Dim aq As Integer
            aq = DataGridView1.Rows.Count

            For i1 = 0 To aq - 1
                DataGridView1.Rows(i1).Cells(12).Value = DataGridView1.Rows(i1).Cells(10).Value - DataGridView1.Rows(i1).Cells(11).Value
                If DataGridView1.Rows(i1).Cells(12).Value < 0 Then
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Yellow

                End If
                If DataGridView1.Rows(i1).Cells(18).Value Is DBNull.Value Then
                    DataGridView1.Rows(i1).Cells(18).Value = Convert.ToString(DataGridView1.Rows(i1).Cells(18).Value)
                End If
            Next




            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(6).Width = 135

            DataGridView1.Columns(8).Width = 80
            DataGridView1.Columns(9).Width = 320
            DataGridView1.Columns(10).Width = 80
            DataGridView1.Columns(11).Width = 80
            DataGridView1.Columns(19).Width = 200
            DataGridView1.Columns(5).Width = 85
            DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.DarkKhaki
            DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
            DataGridView1.Columns(0).Frozen = True
            DataGridView1.Columns(1).Frozen = True
            DataGridView1.Columns(2).Frozen = True
            DataGridView1.Columns(3).Frozen = True
            DataGridView1.Columns(4).Frozen = True
            DataGridView1.Columns(5).Frozen = True
            DataGridView1.Columns(6).Frozen = True
            DataGridView1.Columns(2).Visible = False
            'DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(19).Visible = False
            DataGridView1.Columns(17).Visible = False
            If Label3.Text = 2 Then
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(2).Visible = False
            Else
                If Label3.Text = 3 Then
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(1).Visible = False
                    DataGridView1.Columns(2).Visible = False
                    DataGridView1.Columns(3).Visible = False
                    Button1.Visible = False

                    Button3.Visible = False
                End If
            End If
        Else
            MsgBox("NO HAY INFORMACION PARA MONSTRAR")
            DataGridView1.DataSource = ""
        End If
    End Sub
    Private Sub Programa_Tejeduria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrar23()

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            DataGridView1.Columns(8).Visible = True
        Else
            If CheckBox1.Checked = False Then
                DataGridView1.Columns(8).Visible = False
            End If
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim hg As New fop
        Dim kk As New vop
        Dim i As Integer
        i = DataGridView1.Rows.Count


        For i1 = 0 To i - 1
            If DataGridView1.Rows(i1).Cells(0).Value = True Then
                kk.gncom_3 = Trim(DataGridView1.Rows(i1).Cells(5).Value)
                hg.cerrar_op(kk)
            End If
        Next
        MsgBox("SE CERRO EL PEDIDO SOLICITADO")
        Dim JK As New fop
        DT4 = JK.PROGRAMA_TEJEDURIA()
        If Label3.Text = 2 Then
            Button1.Visible = False
        End If
        If DT4.Rows.Count <> 0 Then
            DataGridView1.DataSource = DT4
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(6).Width = 135
            DataGridView1.Columns(8).Width = 80

            DataGridView1.Columns(9).Width = 320
            DataGridView1.Columns(10).Width = 80
            DataGridView1.Columns(11).Width = 80
            DataGridView1.Columns(15).Width = 200
            DataGridView1.Columns(5).Width = 85

            DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.DarkKhaki
            DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
            If Label3.Text = 2 Then
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(2).Visible = False
            Else
                If Label3.Text = 3 Then
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(1).Visible = False
                    DataGridView1.Columns(2).Visible = False
                    Button1.Visible = False

                End If
            End If
        Else

            DataGridView1.DataSource = ""
        End If
    End Sub

    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(6).Width = 135
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 320
                DataGridView1.Columns(10).Width = 80
                DataGridView1.Columns(11).Width = 80
                DataGridView1.Columns(15).Width = 200
                DataGridView1.Columns(5).Width = 85
                DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
                If Label3.Text = 2 Then
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(2).Visible = False
                Else
                    If Label3.Text = 3 Then
                        DataGridView1.Columns(0).Visible = False
                        DataGridView1.Columns(1).Visible = False
                        DataGridView1.Columns(2).Visible = False
                        Button1.Visible = False

                    End If
                End If
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
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PO" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(6).Width = 135
                DataGridView1.Columns(8).Width = 80

                DataGridView1.Columns(9).Width = 320
                DataGridView1.Columns(10).Width = 80
                DataGridView1.Columns(11).Width = 80
                DataGridView1.Columns(15).Width = 200
                DataGridView1.Columns(5).Width = 85
                DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
                If Label3.Text = 2 Then
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(2).Visible = False
                Else
                    If Label3.Text = 3 Then
                        DataGridView1.Columns(0).Visible = False
                        DataGridView1.Columns(1).Visible = False
                        DataGridView1.Columns(2).Visible = False
                        Button1.Visible = False

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
        buscar1()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar2()
    End Sub
    Dim DT3, DT4 As DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim hg As New fop
        Dim kk As New vop
        Dim i As Integer
        i = DataGridView1.Rows.Count


        For i1 = 0 To i - 1
            If DataGridView1.Rows(i1).Cells(1).Value = True Then
                kk.gncom_3 = Trim(DataGridView1.Rows(i1).Cells(5).Value)
                hg.standby_op(kk)
            End If
        Next
        MsgBox("SE PUSO EN STAND BY EL PEDIDO SOLICITADO")
        Dim JK As New fop
        DT3 = JK.PROGRAMA_TEJEDURIA()
        If Label3.Text = 2 Then
            Button1.Visible = False
        End If
        If DT3.Rows.Count <> 0 Then
            DataGridView1.DataSource = DT3
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(6).Width = 135
            DataGridView1.Columns(8).Width = 80

            DataGridView1.Columns(9).Width = 320
            DataGridView1.Columns(10).Width = 80
            DataGridView1.Columns(11).Width = 80
            DataGridView1.Columns(15).Width = 200
            DataGridView1.Columns(5).Width = 85
            DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.DarkKhaki
            DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
            If Label3.Text = 2 Then
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(2).Visible = False
            Else
                If Label3.Text = 3 Then
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(1).Visible = False
                    DataGridView1.Columns(2).Visible = False
                    Button1.Visible = False

                End If
            End If
        Else

            DataGridView1.DataSource = ""
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim ext As New Exportar
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        abrir()
        Dim U As Integer
        U = DataGridView1.Rows.Count
        For i1 = 0 To U - 1
            If DataGridView1.Rows(i1).Cells(0).Value = True Then
                MsgBox(Trim(DataGridView1.Rows(i1).Cells(5).Value))
                Dim cmd As New SqlCommand("update custom_vianny.dbo.qag0300 set finped_3 = 0 where ncom_3 = @op and ccia ='01'", conx)
                cmd.Parameters.AddWithValue("@op", Trim(DataGridView1.Rows(i1).Cells(5).Value))
                cmd.ExecuteNonQuery()

            End If
        Next
        MsgBox("SE ACTUALIZO LO SOLICITADO")
        mostrar23()
        RadioButton2.Checked = True
        RadioButton1.Checked = False
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Dim JK As New fop
        dt2 = JK.BUSCAR_PROGRAMA_standby()
        If Label3.Text = 2 Then
            Button1.Visible = False
        End If
        If dt2.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt2
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(6).Width = 135
            DataGridView1.Columns(8).Width = 80

            DataGridView1.Columns(7).Width = 320
            DataGridView1.Columns(10).Width = 80
            DataGridView1.Columns(11).Width = 80
            'DataGridView1.Columns(15).Width = 200
            DataGridView1.Columns(5).Width = 85
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.DarkKhaki
            DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
            If Label3.Text = 2 Then
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(2).Visible = False
            Else
                If Label3.Text = 3 Then
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(1).Visible = False
                    DataGridView1.Columns(2).Visible = False
                    Button1.Visible = False

                End If
            End If
        Else

            DataGridView1.DataSource = ""
        End If
        Button3.Enabled = True
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Dim JK As New fop
        dt2 = JK.PROGRAMA_TEJEDURIA()
        If Label3.Text = 2 Then
            Button1.Visible = False
        End If
        If dt2.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt2
            Dim aq As Integer
            aq = DataGridView1.Rows.Count
            For i1 = 0 To aq - 1
                DataGridView1.Rows(i1).Cells(12).Value = DataGridView1.Rows(i1).Cells(10).Value - DataGridView1.Rows(i1).Cells(11).Value
                If DataGridView1.Rows(i1).Cells(12).Value < 0 Then
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Yellow

                End If
                If DataGridView1.Rows(i1).Cells(18).Value Is DBNull.Value Then
                    DataGridView1.Rows(i1).Cells(18).Value = Convert.ToString(DataGridView1.Rows(i1).Cells(18).Value)
                End If
            Next
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(6).Width = 135
            DataGridView1.Columns(8).Width = 80

            DataGridView1.Columns(9).Width = 320
            DataGridView1.Columns(11).Width = 80
            DataGridView1.Columns(15).Width = 200
            DataGridView1.Columns(5).Width = 85
            DataGridView1.Columns(0).Visible = True
            DataGridView1.Columns(1).Visible = True
            DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.DarkKhaki
            DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
            DataGridView1.Columns(2).Visible = False
            'DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(19).Visible = False
            DataGridView1.Columns(17).Visible = False
            If Label3.Text = 2 Then
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(2).Visible = False
            Else
                If Label3.Text = 3 Then
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(1).Visible = False
                    DataGridView1.Columns(2).Visible = False
                    Button1.Visible = False

                End If
            End If
        Else

            DataGridView1.DataSource = ""
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Dim JK As New fop
        dt1 = JK.PROGRAMA_TEJEDURIA_CERRADO()
        If Label3.Text = 2 Then
            Button1.Visible = False
        End If
        If dt1.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt1
            Dim aq As Integer
            aq = DataGridView1.Rows.Count
            For i1 = 0 To aq - 1
                DataGridView1.Rows(i1).Cells(12).Value = DataGridView1.Rows(i1).Cells(10).Value - DataGridView1.Rows(i1).Cells(11).Value
                If DataGridView1.Rows(i1).Cells(12).Value < 0 Then
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Yellow

                End If
                If DataGridView1.Rows(i1).Cells(18).Value Is DBNull.Value Then
                    DataGridView1.Rows(i1).Cells(18).Value = Convert.ToString(DataGridView1.Rows(i1).Cells(18).Value)
                End If
            Next
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(6).Width = 135
            'DataGridView1.Columns(8).Width = 80
            DataGridView1.Columns(7).Width = 320
            DataGridView1.Columns(10).Width = 80
            DataGridView1.Columns(11).Width = 80
            'DataGridView1.Columns(15).Width = 200
            DataGridView1.Columns(5).Width = 85
            DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.DarkKhaki
            DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.DarkKhaki
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(2).Visible = False
            'DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(19).Visible = False
            DataGridView1.Columns(17).Visible = False
            If Label3.Text = 2 Then
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(2).Visible = False
            Else
                If Label3.Text = 3 Then
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(1).Visible = False
                    DataGridView1.Columns(2).Visible = False
                    Button1.Visible = False

                Else
                    DataGridView1.Columns(0).Visible = True
                End If
            End If
        Else

            DataGridView1.DataSource = ""
        End If
        Button3.Enabled = True
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 3 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            F_DESPACJO.TextBox1.Text = Trim(DataGridView1.Rows(num1).Cells(5).Value)
            F_DESPACJO.Label3.Text = Label4.Text
            F_DESPACJO.Show()

            'Reporte__OP.TextBox1.Text = DataGridView1.Rows(num1).Cells(5).Value
            'Reporte__OP.Show()
        Else
            If e.ColumnIndex = 2 Then
                Dim num1, num2 As Integer
                num1 = e.RowIndex.ToString
                num2 = e.ColumnIndex.ToString
                OBSERVACION.TextBox2.Text = Trim(DataGridView1.Rows(num1).Cells(5).Value)
                OBSERVACION.Label1.Text = Label4.Text
                OBSERVACION.Show()
            End If
        End If
    End Sub
End Class