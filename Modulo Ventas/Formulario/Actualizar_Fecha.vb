Imports System.ComponentModel
Imports System.Data.SqlClient
Public Class Actualizar_Fecha
    Private dt As New DataTable

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
    Private Sub Actualizar_Fecha_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()

        Button2.Enabled = False
    End Sub
    Private Sub mostrar()
        Try
            Dim func As New vdata
            Dim func1 As New fadat


            func1.gpartida = TextBox1.Text
            func1.gempresa = Label2.Text
            dt = func.buscar_partida(func1)
            If dt.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt
                Dim u As Integer
                u = DataGridView1.Rows.Count
                DataGridView1.Columns(6).Visible = False
                For i = 0 To u - 1
                    If DataGridView1.Rows(i).Cells(6).Value = "1" Then
                        DataGridView1.Rows(i).Cells(5).Value = "FALTA PESAR"
                        DataGridView1.Rows(i).Cells(3).ReadOnly = True
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    Else
                        If DataGridView1.Rows(i).Cells(6).Value = "4" Then
                            DataGridView1.Rows(i).Cells(5).Value = "PROCESO RAMA"
                            DataGridView1.Rows(i).Cells(3).ReadOnly = True
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Azure
                        Else
                            If DataGridView1.Rows(i).Cells(6).Value = "2" Then
                                DataGridView1.Rows(i).Cells(5).Value = "PESADO"
                                DataGridView1.Rows(i).Cells(3).ReadOnly = False
                            Else
                                If DataGridView1.Rows(i).Cells(6).Value = "0" Then
                                    DataGridView1.Rows(i).Cells(5).Value = "ANULADO"
                                    DataGridView1.Rows(i).Cells(3).ReadOnly = True
                                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                                Else
                                    DataGridView1.Rows(i).Cells(5).Value = "AUDITADO"
                                    DataGridView1.Rows(i).Cells(3).ReadOnly = False
                                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Beige
                                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                                End If

                            End If
                        End If
                    End If
                Next
            Else
                MsgBox("La Partida no Existe")
            End If
        Catch ex As Exception

            MsgBox("NO EXISTE INFORMACION PARA MOSTRAR")

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        mostrar()
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        Button2.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim func As New vdata
        Dim func1 As New fplanea
        Dim i, a As Integer
        i = DataGridView1.Rows.Count
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                func1.grollo = Me.DataGridView1.Rows(a).Cells(2).Value
                func1.gpartida = Me.DataGridView1.Rows(a).Cells(1).Value
                func1.gfecha = DateTimePicker1.Value
                func1.gccia = Label2.Text
                func.actulizar_calidad(func1)

            End If

        Next
        MsgBox("LOS CAMPOS SE ACTUALIZARON CORRECTAMENTE")
        'DataGridView1.DataSource = ""
        'RadioButton1.Enabled = False
        'RadioButton2.Enabled = False
        'Button2.Enabled = False

        mostrar()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Dim i, a As Integer
        i = DataGridView1.Rows.Count
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = False Then
                Me.DataGridView1.Rows(a).Cells(0).Value = True
            End If

        Next
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Dim i, a As Integer
        i = DataGridView1.Rows.Count
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                Me.DataGridView1.Rows(a).Cells(0).Value = False
            End If

        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim func As New vdata
        Dim func1 As New fplanea
        Dim i, a As Integer
        i = DataGridView1.Rows.Count
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                func1.grollo = Me.DataGridView1.Rows(a).Cells(2).Value
                func1.gpartida = Me.DataGridView1.Rows(a).Cells(1).Value
                func1.gpeso = Me.DataGridView1.Rows(a).Cells(3).Value
                func1.gccia = Label2.Text
                func.actulizar_rollo(func1)

            End If

        Next
        MsgBox("LOS CAMPOS SE ACTUALIZARON CORRECTAMENTE")
        'DataGridView1.DataSource = ""
        'RadioButton1.Enabled = False
        'RadioButton2.Enabled = False
        'Button2.Enabled = False

        mostrar()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim j As Integer
        j = DataGridView1.Rows.Count
        For i = 0 To j - 1
            If Me.DataGridView1.Rows(i).Cells(0).Value = True Then
                abrir()
                Dim cmd As New SqlCommand("update custom_vianny.dbo.marg0001 set flag_3r =0  where  rollo_3r =@rollo and ccia_3r =@ccia", conx)
                cmd.Parameters.AddWithValue("@rollo", Trim(Me.DataGridView1.Rows(i).Cells(2).Value))
                cmd.Parameters.AddWithValue("@ccia", Trim(Label2.Text))
                cmd.ExecuteNonQuery()
            End If

        Next
        MsgBox("SE ANULARON LOS ROLLOS SOLICITADOS")
        mostrar()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim j As Integer
        j = DataGridView1.Rows.Count
        For i = 0 To j - 1
            If Me.DataGridView1.Rows(i).Cells(0).Value = True Then
                abrir()
                Dim cmd As New SqlCommand("update custom_vianny.dbo.marg0001 set flag_3r =2  where  rollo_3r =@rollo and ccia_3r =@ccia", conx)
                cmd.Parameters.AddWithValue("@rollo", Trim(Me.DataGridView1.Rows(i).Cells(2).Value))
                cmd.Parameters.AddWithValue("@ccia", Trim(Label2.Text))
                cmd.ExecuteNonQuery()
            End If

        Next
        MsgBox("SE ANULARON LOS ROLLOS SOLICITADOS")
        mostrar()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim j As Integer
        j = DataGridView1.Rows.Count
        For i = 0 To j - 1
            If Me.DataGridView1.Rows(i).Cells(0).Value = True Then
                abrir()
                Dim cmd As New SqlCommand("update custom_vianny.dbo.marg0001 set flag_3r =3  where  rollo_3r =@rollo and ccia_3r =@ccia", conx)
                cmd.Parameters.AddWithValue("@rollo", Trim(Me.DataGridView1.Rows(i).Cells(2).Value))
                cmd.Parameters.AddWithValue("@ccia", Trim(Label2.Text))
                cmd.ExecuteNonQuery()
            End If

        Next
        MsgBox("LOS ROLLOS SE AUDIRARON CORRECTAMENTE")
        mostrar()
    End Sub
End Class