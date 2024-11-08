Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class Od_Udp
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim Rsr2, Rsr3, Rsr212 As SqlDataReader
    Dim da As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub mostrar_rutas()
        abrir()
        DataGridView1.DataSource = ""
        da.Clear()
        Dim cmd As New SqlDataAdapter("SELECT * FROM  Ruta_Udp WHERE OdRut ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' order by EtaRut", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        Dim O As Integer
        O = DataGridView1.Rows.Count

        If O > 0 Then

            GroupBox5.Enabled = False
            Button20.Enabled = False
            TextBox60.Text = DataGridView1.Rows(0).Cells(5).Value
            TextBox61.Text = DataGridView1.Rows(1).Cells(5).Value
            TextBox62.Text = DataGridView1.Rows(2).Cells(5).Value
            TextBox63.Text = DataGridView1.Rows(3).Cells(5).Value
            TextBox64.Text = DataGridView1.Rows(4).Cells(5).Value
            TextBox65.Text = DataGridView1.Rows(5).Cells(5).Value
            TextBox66.Text = DataGridView1.Rows(6).Cells(5).Value
            TextBox56.Text = DataGridView1.Rows(7).Cells(5).Value
            ComboBox4.Text = DataGridView1.Rows(0).Cells(4).Value
            ComboBox6.Text = DataGridView1.Rows(1).Cells(4).Value
            TextBox83.Text = DataGridView1.Rows(3).Cells(4).Value
            TextBox84.Text = DataGridView1.Rows(4).Cells(4).Value
            TextBox85.Text = DataGridView1.Rows(5).Cells(4).Value

            DateTimePicker4.Value = DataGridView1.Rows(0).Cells(3).Value
            DateTimePicker5.Value = DataGridView1.Rows(1).Cells(3).Value
            DateTimePicker6.Value = DataGridView1.Rows(2).Cells(3).Value
            DateTimePicker7.Value = DataGridView1.Rows(3).Cells(3).Value
            DateTimePicker8.Value = DataGridView1.Rows(4).Cells(3).Value
            DateTimePicker9.Value = DataGridView1.Rows(5).Cells(3).Value
            DateTimePicker10.Value = DataGridView1.Rows(6).Cells(3).Value
            DateTimePicker11.Value = DataGridView1.Rows(7).Cells(3).Value
            DateTimePicker12.Value = DataGridView1.Rows(5).Cells(6).Value

            If Trim(DataGridView1.Rows(0).Cells(7).Value) = "1" Then
                Button26.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox57.Text = 1
                CheckBox1.Enabled = False
            Else
                Button26.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox1.Enabled = True
                TextBox57.Text = 0
            End If
            If Trim(DataGridView1.Rows(1).Cells(7).Value) = "1" Then
                Button27.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox67.Text = 1
                CheckBox2.Enabled = False
            Else
                Button27.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox2.Enabled = True
                TextBox67.Text = 0
            End If
            If Trim(DataGridView1.Rows(2).Cells(7).Value) = "1" Then
                Button28.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox68.Text = 1
                CheckBox3.Enabled = False

            Else
                Button28.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox3.Enabled = True
                TextBox68.Text = 0
            End If
            If Trim(DataGridView1.Rows(3).Cells(7).Value) = "1" Then
                Button29.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox89.Text = 1
                CheckBox4.Enabled = False
            Else
                Button29.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox4.Enabled = True
                TextBox89.Text = 0
            End If
            If Trim(DataGridView1.Rows(4).Cells(7).Value) = "1" Then
                Button30.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox90.Text = 1
                CheckBox5.Enabled = False
            Else
                Button30.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox5.Enabled = True
                TextBox90.Text = 0
            End If
            If Trim(DataGridView1.Rows(5).Cells(7).Value) = "1" Then
                Button31.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox91.Text = 1
                CheckBox6.Enabled = False
            Else
                Button31.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox6.Enabled = True
                TextBox91.Text = 0
            End If
            If Trim(DataGridView1.Rows(6).Cells(7).Value) = "1" Then
                Button32.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox92.Text = 1
                CheckBox7.Enabled = False
            Else
                Button32.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox7.Enabled = True
                TextBox92.Text = 0
            End If
            If Trim(DataGridView1.Rows(7).Cells(7).Value) = "1" Then
                Button33.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox93.Text = 1
                CheckBox8.Enabled = False
            Else
                Button33.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox8.Enabled = True
                TextBox93.Text = 0
            End If

        Else

            'TextBox60.Text = ""
            'TextBox61.Text = ""
            'TextBox62.Text = ""
            'TextBox63.Text = ""
            'TextBox64.Text = ""
            'TextBox65.Text = ""
            'TextBox66.Text = ""
            'TextBox56.Text = ""
            'TextBox80.Text = ""
            'TextBox83.Text = ""
            'TextBox84.Text = ""
            'TextBox85.Text = ""
            TextBox57.Text = "2"
            TextBox67.Text = "2"
            TextBox68.Text = "2"
            TextBox89.Text = "2"
            TextBox90.Text = "2"
            TextBox91.Text = "2"
            TextBox92.Text = "2"
            TextBox93.Text = "2"
            'Button26.Image = Nothing
            'Button27.Image = Nothing
            'Button28.Image = Nothing
            'Button29.Image = Nothing
            'Button30.Image = Nothing
            'Button31.Image = Nothing
            'Button32.Image = Nothing
            'Button33.Image = Nothing
            'DateTimePicker4.Value = DateTime.Now
            'DateTimePicker5.Value = DateTime.Now
            'DateTimePicker6.Value = DateTime.Now
            'DateTimePicker7.Value = DateTime.Now
            'DateTimePicker8.Value = DateTime.Now
            'DateTimePicker9.Value = DateTime.Now
            'DateTimePicker10.Value = DateTime.Now
            'DateTimePicker11.Value = DateTime.Now
            'DateTimePicker12.Value = DateTime.Now

        End If



    End Sub
    Sub mostrar_rutas_update()
        abrir()
        DataGridView1.DataSource = ""
        da.Clear()
        Dim cmd As New SqlDataAdapter("SELECT * FROM  Ruta_Udp WHERE OdRut ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' order by EtaRut", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        Dim O As Integer
        O = DataGridView1.Rows.Count

        If O > 0 Then


            TextBox60.Text = DataGridView1.Rows(0).Cells(5).Value
            TextBox61.Text = DataGridView1.Rows(1).Cells(5).Value
            TextBox62.Text = DataGridView1.Rows(2).Cells(5).Value
            TextBox63.Text = DataGridView1.Rows(3).Cells(5).Value
            TextBox64.Text = DataGridView1.Rows(4).Cells(5).Value
            TextBox65.Text = DataGridView1.Rows(5).Cells(5).Value
            TextBox66.Text = DataGridView1.Rows(6).Cells(5).Value
            TextBox56.Text = DataGridView1.Rows(7).Cells(5).Value
            ComboBox6.Text = DataGridView1.Rows(1).Cells(4).Value
            TextBox83.Text = DataGridView1.Rows(3).Cells(4).Value
            TextBox84.Text = DataGridView1.Rows(4).Cells(4).Value
            TextBox85.Text = DataGridView1.Rows(5).Cells(4).Value
            DateTimePicker4.Value = DataGridView1.Rows(0).Cells(3).Value
            DateTimePicker5.Value = DataGridView1.Rows(1).Cells(3).Value
            DateTimePicker6.Value = DataGridView1.Rows(2).Cells(3).Value
            DateTimePicker7.Value = DataGridView1.Rows(3).Cells(3).Value
            DateTimePicker8.Value = DataGridView1.Rows(4).Cells(3).Value
            DateTimePicker9.Value = DataGridView1.Rows(5).Cells(3).Value
            DateTimePicker10.Value = DataGridView1.Rows(6).Cells(3).Value
            DateTimePicker11.Value = DataGridView1.Rows(7).Cells(3).Value
            DateTimePicker12.Value = DataGridView1.Rows(5).Cells(6).Value

            If Trim(DataGridView1.Rows(0).Cells(7).Value) = "1" Then
                Button26.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox57.Text = 1
                CheckBox1.Enabled = False
            Else
                Button26.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox1.Enabled = True
                TextBox57.Text = 0
            End If
            If Trim(DataGridView1.Rows(1).Cells(7).Value) = "1" Then
                Button27.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox67.Text = 1
                CheckBox2.Enabled = False
            Else
                Button27.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox2.Enabled = True
                TextBox67.Text = 0
            End If
            If Trim(DataGridView1.Rows(2).Cells(7).Value) = "1" Then
                Button28.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox68.Text = 1
                CheckBox3.Enabled = False

            Else
                Button28.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox3.Enabled = True
                TextBox68.Text = 0
            End If
            If Trim(DataGridView1.Rows(3).Cells(7).Value) = "1" Then
                Button29.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox89.Text = 1
                CheckBox4.Enabled = False
            Else
                Button29.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox4.Enabled = True
                TextBox89.Text = 0
            End If
            If Trim(DataGridView1.Rows(4).Cells(7).Value) = "1" Then
                Button30.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox90.Text = 1
                CheckBox5.Enabled = False
            Else
                Button30.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox5.Enabled = True
                TextBox90.Text = 0
            End If
            If Trim(DataGridView1.Rows(5).Cells(7).Value) = "1" Then
                Button31.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox91.Text = 1
                CheckBox6.Enabled = False
            Else
                Button31.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox6.Enabled = True
                TextBox91.Text = 0
            End If
            If Trim(DataGridView1.Rows(6).Cells(7).Value) = "1" Then
                Button32.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox92.Text = 1
                CheckBox7.Enabled = False
            Else
                Button32.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox7.Enabled = True
                TextBox92.Text = 0
            End If
            If Trim(DataGridView1.Rows(7).Cells(7).Value) = "1" Then
                Button33.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox93.Text = 1
                CheckBox8.Enabled = False
            Else
                Button33.Image = Global.Modulo_Ventas.Resources.can_cer
                CheckBox8.Enabled = True
                TextBox93.Text = 0
            End If

        Else

            'TextBox60.Text = ""
            'TextBox61.Text = ""
            'TextBox62.Text = ""
            'TextBox63.Text = ""
            'TextBox64.Text = ""
            'TextBox65.Text = ""
            'TextBox66.Text = ""
            'TextBox56.Text = ""
            'TextBox80.Text = ""
            'TextBox83.Text = ""
            'TextBox84.Text = ""
            'TextBox85.Text = ""
            TextBox57.Text = "2"
            TextBox67.Text = "2"
            TextBox68.Text = "2"
            TextBox89.Text = "2"
            TextBox90.Text = "2"
            TextBox91.Text = "2"
            TextBox92.Text = "2"
            TextBox93.Text = "2"
            'Button26.Image = Nothing
            'Button27.Image = Nothing
            'Button28.Image = Nothing
            'Button29.Image = Nothing
            'Button30.Image = Nothing
            'Button31.Image = Nothing
            'Button32.Image = Nothing
            'Button33.Image = Nothing
            'DateTimePicker4.Value = DateTime.Now
            'DateTimePicker5.Value = DateTime.Now
            'DateTimePicker6.Value = DateTime.Now
            'DateTimePicker7.Value = DateTime.Now
            'DateTimePicker8.Value = DateTime.Now
            'DateTimePicker9.Value = DateTime.Now
            'DateTimePicker10.Value = DateTime.Now
            'DateTimePicker11.Value = DateTime.Now
            'DateTimePicker12.Value = DateTime.Now

        End If



    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        abrir()
        Dim cantod As String
        Dim sql1021 As String = "Select CiRut from Ruta_Udp where OdRut ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' and EtaRut ='" + Trim(TextBox40.Text) + "'"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr3 = cmd1021.ExecuteReader()
        If Rsr3.Read() = True Then
            cantod = Rsr3(0)
        Else
            cantod = "2"
        End If
        Rsr3.Close()

        If cantod = "0" Then
            Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
            cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
            cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox40.Text))
            cmd20.ExecuteNonQuery()
            TextBox57.Text = "1"
            CheckBox1.Enabled = False
            CheckBox1.Checked = False
            Button26.Image = Global.Modulo_Ventas.Resources.can_abi
            MsgBox("SE CERRO EL PROCESO ANALISIS DE PRENDA")
        Else
            If cantod = "1" Then
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                    cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                    cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox40.Text))
                    cmd20.ExecuteNonQuery()
                    TextBox57.Text = "0"
                    CheckBox1.Enabled = True
                    CheckBox1.Checked = True
                    Button26.Image = Global.Modulo_Ventas.Resources.can_cer
                    MsgBox("SE APERTURO EL PROCESO ANALISIS DE PRENDA")
                End If
            Else
                MsgBox("NO SE PUEDE EJECUTAR, PORQUE TODAVIA NO HA REGISTRADO INFORMACION DE RUTA")
            End If

        End If
        mostrar_rutas_update()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        If Trim(TextBox57.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO ANALISIS DE PRENDA")
        Else
            abrir()
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox1.Text) & (TextBox2.Text) + "' and EtaRut ='" + Trim(TextBox43.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            Else
                cantod = "2"
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox43.Text))
                cmd20.ExecuteNonQuery()
                Button27.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox67.Text = "1"
                CheckBox2.Enabled = False
                CheckBox2.Checked = False
                MsgBox("SE CERRO EL PROCESO CREACION DE MOLDE")
            Else
                If cantod = "1" Then
                    Dim respuesta As DialogResult
                    respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                        cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                        cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox43.Text))
                        cmd20.ExecuteNonQuery()
                        TextBox67.Text = "0"
                        CheckBox2.Enabled = True
                        CheckBox2.Checked = True
                        Button27.Image = Global.Modulo_Ventas.Resources.can_cer
                        MsgBox("SE APERTURO EL PROCESO CREACION DE MOLDE")
                    End If
                Else
                    MsgBox("NO SE PUEDE EJECUTAR, PORQUE TODAVIA NO HA REGISTRADO INFORMACION DE RUTA")
                End If

            End If
            mostrar_rutas_update()
        End If
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        If Trim(TextBox67.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO CREACION DE MOLDE")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox1.Text) & (TextBox2.Text) + "' and EtaRut ='" + Trim(TextBox45.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            Else
                cantod = "2"
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox45.Text))
                cmd20.ExecuteNonQuery()
                TextBox68.Text = "1"
                CheckBox3.Enabled = False
                CheckBox3.Checked = False
                Button28.Image = Global.Modulo_Ventas.Resources.can_abi

                MsgBox("SE CERRO EL PROCESO DE CORTE")
            Else
                If cantod = "1" Then
                    Dim respuesta As DialogResult
                    respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                        cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                        cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox45.Text))
                        cmd20.ExecuteNonQuery()
                        TextBox68.Text = "0"
                        CheckBox3.Enabled = True
                        CheckBox3.Checked = True
                        Button28.Image = Global.Modulo_Ventas.Resources.can_cer
                        MsgBox("SE APERTURO EL PROCESO DE CORTE")
                    End If
                Else
                    MsgBox("NO SE PUEDE EJECUTAR, PORQUE TODAVIA NO HA REGISTRADO INFORMACION DE RUTA")
                End If

            End If
            mostrar_rutas_update()
        End If
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        If Trim(TextBox68.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO CORTE")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox1.Text) & (TextBox2.Text) + "' and EtaRut ='" + Trim(TextBox47.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            Else
                cantod = "2"
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox47.Text))
                cmd20.ExecuteNonQuery()
                Button29.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox89.Text = "1"
                CheckBox4.Enabled = False
                CheckBox4.Checked = False
                MsgBox("SE CERRO EL PROCESO DE APLICACIONES")
            Else
                If cantod = "1" Then
                    Dim respuesta As DialogResult
                    respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                        cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                        cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox47.Text))
                        cmd20.ExecuteNonQuery()
                        TextBox89.Text = "0"
                        CheckBox4.Enabled = True
                        CheckBox4.Checked = True
                        Button29.Image = Global.Modulo_Ventas.Resources.can_cer
                        MsgBox("SE APERTURO EL PROCESO DE APLICACIONES")
                    End If
                Else
                    MsgBox("NO SE PUEDE EJECUTAR, PORQUE TODAVIA NO HA REGISTRADO INFORMACION DE RUTA")
                End If

            End If
            mostrar_rutas_update()
        End If
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        If Trim(TextBox89.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO APLICACIONES")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox1.Text) & (TextBox2.Text) + "' and EtaRut ='" + Trim(TextBox49.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            Else
                cantod = "2"
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox49.Text))
                cmd20.ExecuteNonQuery()
                TextBox90.Text = "1"
                CheckBox5.Enabled = False
                CheckBox5.Checked = False
                Button30.Image = Global.Modulo_Ventas.Resources.can_abi
                MsgBox("SE CERRO EL PROCESO DE CONFECCION")
            Else
                If cantod = "1" Then
                    Dim respuesta As DialogResult
                    respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                        cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                        cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox49.Text))
                        cmd20.ExecuteNonQuery()
                        TextBox90.Text = "0"
                        CheckBox5.Enabled = True
                        CheckBox5.Checked = True
                        Button30.Image = Global.Modulo_Ventas.Resources.can_cer
                        MsgBox("SE APERTURO EL PROCESO DE CONFECCION")
                    End If
                Else
                    MsgBox("NO SE PUEDE EJECUTAR, PORQUE TODAVIA NO HA REGISTRADO INFORMACION DE RUTA")
                End If

            End If
            mostrar_rutas_update()
        End If
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        If Trim(TextBox90.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO CONFECCIONES")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox1.Text) & (TextBox2.Text) + "' and EtaRut ='" + Trim(TextBox51.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            Else
                cantod = "2"
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox51.Text))
                cmd20.ExecuteNonQuery()
                TextBox91.Text = "1"
                CheckBox6.Enabled = False
                CheckBox6.Checked = False
                Button31.Image = Global.Modulo_Ventas.Resources.can_abi
                MsgBox("SE CERRO EL PROCESO DE LAVADO")
            Else
                If cantod = "1" Then
                    Dim respuesta As DialogResult
                    respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                        cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                        cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox51.Text))
                        cmd20.ExecuteNonQuery()
                        CheckBox6.Enabled = True
                        CheckBox6.Checked = True
                        TextBox91.Text = "0"
                        Button31.Image = Global.Modulo_Ventas.Resources.can_cer
                        MsgBox("SE APERTURO EL PROCESO DE LAVADO")
                    End If
                Else
                    MsgBox("NO SE PUEDE EJECUTAR, PORQUE TODAVIA NO HA REGISTRADO INFORMACION DE RUTA")
                End If

            End If
            mostrar_rutas_update()
        End If
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        If Trim(TextBox91.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO LAVADO")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox1.Text) & (TextBox2.Text) + "' and EtaRut ='" + Trim(TextBox53.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            Else
                cantod = "2"
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox53.Text))
                cmd20.ExecuteNonQuery()
                Button32.Image = Global.Modulo_Ventas.Resources.can_abi
                TextBox92.Text = "1"
                CheckBox7.Enabled = False
                CheckBox7.Checked = False
                MsgBox("SE CERRO EL PROCESO DE ACABADOS")
            Else
                If cantod = "1" Then
                    Dim respuesta As DialogResult
                    respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                        cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                        cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox53.Text))
                        cmd20.ExecuteNonQuery()
                        TextBox92.Text = "0"
                        CheckBox7.Enabled = True
                        CheckBox7.Checked = True
                        Button32.Image = Global.Modulo_Ventas.Resources.can_cer
                        MsgBox("SE APERTURO EL PROCESO DE ACABADOS")
                    End If
                Else
                    MsgBox("NO SE PUEDE EJECUTAR, PORQUE TODAVIA NO HA REGISTRADO INFORMACION DE RUTA")
                End If

            End If
            mostrar_rutas_update()
        End If
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        If Trim(TextBox92.Text) = "0" Then
            MsgBox("NO SE PUEDE CERRAR ESTE PROCESO, CIERRE PRIMERO ACABADOS")
        Else
            Dim cantod As String
            Dim sql1021 As String = "select CiRut from Ruta_Udp where OdRut ='" + (TextBox1.Text) & (TextBox2.Text) + "' and EtaRut ='" + Trim(TextBox55.Text) + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr3 = cmd1021.ExecuteReader()
            If Rsr3.Read() = True Then
                cantod = Rsr3(0)
            Else
                cantod = "2"
            End If
            Rsr3.Close()

            If cantod = 0 Then
                Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox55.Text))
                cmd20.ExecuteNonQuery()
                TextBox93.Text = "1"
                CheckBox8.Enabled = False
                CheckBox8.Checked = False
                Button33.Image = Global.Modulo_Ventas.Resources.can_abi
                MsgBox("SE CERRO EL PROCESO ENTREGA COMERCIAL")
            Else
                If cantod = "1" Then
                    Dim respuesta As DialogResult
                    respuesta = MessageBox.Show("EL PREOCESO ESTA CERRADO, DESEA APERTURARLO", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =0 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                        cmd20.Parameters.AddWithValue("@OdRut", (TextBox1.Text) & (TextBox2.Text))
                        cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox55.Text))
                        cmd20.ExecuteNonQuery()
                        TextBox93.Text = "0"
                        CheckBox8.Enabled = True
                        CheckBox8.Checked = True
                        Button33.Image = Global.Modulo_Ventas.Resources.can_cer
                        MsgBox("SE APERTURO EL PROCESO ENTREGA COMERCIAL")
                    End If
                Else
                    MsgBox("NO SE PUEDE EJECUTAR, PORQUE TODAVIA NO HA REGISTRADO INFORMACION DE RUTA")
                End If

            End If
            mostrar_rutas_update()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            DateTimePicker4.Enabled = True
            TextBox60.Enabled = True
            ComboBox6.Enabled = True
            TextBox60.Select()
        Else
            DateTimePicker4.Enabled = False
            ComboBox6.Enabled = True
            TextBox60.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            DateTimePicker5.Enabled = True
            ComboBox6.Enabled = True
            TextBox61.Enabled = True
            TextBox61.Select()
        Else
            DateTimePicker5.Enabled = False
            ComboBox6.Enabled = False
            TextBox61.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            DateTimePicker6.Enabled = True
            TextBox62.Enabled = True
            TextBox62.Select()
        Else
            DateTimePicker6.Enabled = False
            TextBox62.Enabled = False
        End If

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            DateTimePicker7.Enabled = True
            TextBox63.Enabled = True
            TextBox83.Enabled = True
            TextBox83.Select()
        Else
            DateTimePicker7.Enabled = False
            TextBox83.Enabled = False
            TextBox63.Enabled = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            DateTimePicker8.Enabled = True
            TextBox64.Enabled = True
            TextBox84.Enabled = True
            TextBox84.Select()
        Else
            DateTimePicker8.Enabled = False
            TextBox85.Enabled = True
            TextBox64.Enabled = False
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            DateTimePicker9.Enabled = True
            DateTimePicker12.Enabled = True
            TextBox65.Enabled = True
            TextBox85.Enabled = True
            TextBox85.Select()
        Else
            DateTimePicker9.Enabled = False
            DateTimePicker12.Enabled = False
            TextBox85.Enabled = True
            TextBox65.Enabled = False
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            DateTimePicker10.Enabled = True
            TextBox66.Enabled = True
            TextBox66.Select()
        Else
            DateTimePicker10.Enabled = False
            TextBox66.Enabled = False
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            DateTimePicker11.Enabled = True
            TextBox56.Enabled = True
            TextBox56.Select()
        Else
            DateTimePicker11.Enabled = False
            TextBox56.Enabled = False
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        abrir()
        'If Trim(TextBox80.Text).Length = 0 OrElse Trim(TextBox84.Text).Length = 0 OrElse Trim(TextBox85.Text).Length = 0 Then
        '    MsgBox("LOS CAMPOS DONDE SE REGISTRA EL TALLER O USUARIO QUE REALIZA LA CREACION DE MOLDE, CONFECCION, ACABADOS SON OBLIGATORIOS")

        'Else

        eliminar_Rutas()
            If Trim(TextBox57.Text) = "0" Or Trim(TextBox57.Text) = "2" Then
                Dim cmd20 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd20.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox40.Text))
                cmd20.Parameters.AddWithValue("@NomRut", Trim(TextBox41.Text))
                cmd20.Parameters.AddWithValue("@FechRut", DateTimePicker4.Value)
                cmd20.Parameters.AddWithValue("@FecIde", DateTimePicker5.Value)
            cmd20.Parameters.AddWithValue("@AsiRut", Trim(ComboBox4.Text.ToString()))
            cmd20.Parameters.AddWithValue("@ObsRut", Trim(TextBox60.Text))
            cmd20.Parameters.AddWithValue("@CiRut", "1")
            cmd20.ExecuteNonQuery()
            Else
                Dim cmd20 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd20.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd20.Parameters.AddWithValue("@EtaRut", Trim(TextBox40.Text))
                cmd20.Parameters.AddWithValue("@NomRut", Trim(TextBox41.Text))
                cmd20.Parameters.AddWithValue("@FechRut", DateTimePicker4.Value)
                cmd20.Parameters.AddWithValue("@FecIde", DateTimePicker5.Value)
            cmd20.Parameters.AddWithValue("@AsiRut", Trim(ComboBox4.Text.ToString()))
            cmd20.Parameters.AddWithValue("@ObsRut", Trim(TextBox60.Text))
                cmd20.Parameters.AddWithValue("@CiRut", "1")
                cmd20.ExecuteNonQuery()
            End If

            If Trim(TextBox67.Text) = "0" Or Trim(TextBox67.Text) = "2" Then
                Dim cmd21 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd21.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd21.Parameters.AddWithValue("@EtaRut", Trim(TextBox43.Text))
                cmd21.Parameters.AddWithValue("@NomRut", Trim(TextBox42.Text))
                cmd21.Parameters.AddWithValue("@FechRut", DateTimePicker5.Value)
                cmd21.Parameters.AddWithValue("@FecIde", DateTimePicker6.Value)
            cmd21.Parameters.AddWithValue("@AsiRut", Trim(ComboBox6.Text.ToString()))
            cmd21.Parameters.AddWithValue("@ObsRut", Trim(TextBox61.Text))
            cmd21.Parameters.AddWithValue("@CiRut", "1")
            cmd21.ExecuteNonQuery()
            Else
                Dim cmd21 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd21.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd21.Parameters.AddWithValue("@EtaRut", Trim(TextBox43.Text))
                cmd21.Parameters.AddWithValue("@NomRut", Trim(TextBox42.Text))
                cmd21.Parameters.AddWithValue("@FechRut", DateTimePicker5.Value)

                cmd21.Parameters.AddWithValue("@FecIde", DateTimePicker6.Value)
            cmd21.Parameters.AddWithValue("@AsiRut", Trim(ComboBox6.Text.ToString()))
            cmd21.Parameters.AddWithValue("@ObsRut", Trim(TextBox61.Text))
                cmd21.Parameters.AddWithValue("@CiRut", "1")
                cmd21.ExecuteNonQuery()
            End If

            If Trim(TextBox68.Text) = "0" Or Trim(TextBox68.Text) = "2" Then
                Dim cmd22 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd22.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd22.Parameters.AddWithValue("@EtaRut", Trim(TextBox45.Text))
                cmd22.Parameters.AddWithValue("@NomRut", Trim(TextBox44.Text))
                cmd22.Parameters.AddWithValue("@FechRut", DateTimePicker6.Value)
                If Trim(TextBox83.Text).Length > 0 Then
                    cmd22.Parameters.AddWithValue("@FecIde", DateTimePicker7.Value)
                Else
                    cmd22.Parameters.AddWithValue("@FecIde", DateTimePicker8.Value)
                End If

                cmd22.Parameters.AddWithValue("@AsiRut", "")
                cmd22.Parameters.AddWithValue("@ObsRut", Trim(TextBox62.Text))
                cmd22.Parameters.AddWithValue("@CiRut", "0")
                cmd22.ExecuteNonQuery()
            Else
                Dim cmd22 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd22.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd22.Parameters.AddWithValue("@EtaRut", Trim(TextBox45.Text))
                cmd22.Parameters.AddWithValue("@NomRut", Trim(TextBox44.Text))
                cmd22.Parameters.AddWithValue("@FechRut", DateTimePicker6.Value)
                If Trim(TextBox83.Text).Length > 0 Then
                    cmd22.Parameters.AddWithValue("@FecIde", DateTimePicker7.Value)
                Else
                    cmd22.Parameters.AddWithValue("@FecIde", DateTimePicker8.Value)
                End If
                cmd22.Parameters.AddWithValue("@AsiRut", "")
                cmd22.Parameters.AddWithValue("@ObsRut", Trim(TextBox62.Text))
                cmd22.Parameters.AddWithValue("@CiRut", "1")
                cmd22.ExecuteNonQuery()
            End If


            If Trim(TextBox89.Text) = "0" Or Trim(TextBox89.Text) = "2" Then
                Dim cmd24 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd24.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd24.Parameters.AddWithValue("@EtaRut", Trim(TextBox47.Text))
                cmd24.Parameters.AddWithValue("@NomRut", Trim(TextBox46.Text))
                cmd24.Parameters.AddWithValue("@FechRut", DateTimePicker7.Value)
                If Trim(TextBox83.Text).Length > 0 Then
                    cmd24.Parameters.AddWithValue("@FecIde", DateTimePicker8.Value)
                Else
                    cmd24.Parameters.AddWithValue("@FecIde", DateTimePicker7.Value)
                End If

                cmd24.Parameters.AddWithValue("@AsiRut", Trim(TextBox83.Text))
                cmd24.Parameters.AddWithValue("@ObsRut", Trim(TextBox63.Text))
                cmd24.Parameters.AddWithValue("@CiRut", "0")
                cmd24.ExecuteNonQuery()
            Else
                Dim cmd24 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd24.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd24.Parameters.AddWithValue("@EtaRut", Trim(TextBox47.Text))
                cmd24.Parameters.AddWithValue("@NomRut", Trim(TextBox46.Text))
                cmd24.Parameters.AddWithValue("@FechRut", DateTimePicker7.Value)
                If Trim(TextBox83.Text).Length > 0 Then
                    cmd24.Parameters.AddWithValue("@FecIde", DateTimePicker8.Value)
                Else
                    cmd24.Parameters.AddWithValue("@FecIde", DateTimePicker7.Value)
                End If
                cmd24.Parameters.AddWithValue("@AsiRut", Trim(TextBox83.Text))
                cmd24.Parameters.AddWithValue("@ObsRut", Trim(TextBox63.Text))
                cmd24.Parameters.AddWithValue("@CiRut", "1")
                cmd24.ExecuteNonQuery()
            End If

            If Trim(TextBox90.Text) = "0" Or Trim(TextBox90.Text) = "2" Then
                Dim cmd25 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd25.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd25.Parameters.AddWithValue("@EtaRut", Trim(TextBox49.Text))
                cmd25.Parameters.AddWithValue("@NomRut", Trim(TextBox48.Text))
                cmd25.Parameters.AddWithValue("@FechRut", DateTimePicker8.Value)
                cmd25.Parameters.AddWithValue("@FecIde", DateTimePicker9.Value)
                cmd25.Parameters.AddWithValue("@AsiRut", Trim(TextBox84.Text))
                cmd25.Parameters.AddWithValue("@ObsRut", Trim(TextBox64.Text))
                cmd25.Parameters.AddWithValue("@CiRut", "0")
                cmd25.ExecuteNonQuery()
            Else
                Dim cmd25 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd25.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd25.Parameters.AddWithValue("@EtaRut", Trim(TextBox49.Text))
                cmd25.Parameters.AddWithValue("@NomRut", Trim(TextBox48.Text))
                cmd25.Parameters.AddWithValue("@FechRut", DateTimePicker8.Value)
                cmd25.Parameters.AddWithValue("@FecIde", DateTimePicker9.Value)
                cmd25.Parameters.AddWithValue("@AsiRut", Trim(TextBox84.Text))
                cmd25.Parameters.AddWithValue("@ObsRut", Trim(TextBox64.Text))
                cmd25.Parameters.AddWithValue("@CiRut", "1")
                cmd25.ExecuteNonQuery()
            End If



            If Trim(TextBox91.Text) = "0" Or Trim(TextBox91.Text) = "2" Then
                Dim cmd27 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd27.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd27.Parameters.AddWithValue("@EtaRut", Trim(TextBox51.Text))
                cmd27.Parameters.AddWithValue("@NomRut", Trim(TextBox50.Text))
                cmd27.Parameters.AddWithValue("@FechRut", DateTimePicker9.Value)

                cmd27.Parameters.AddWithValue("@AsiRut", Trim(TextBox85.Text))
                cmd27.Parameters.AddWithValue("@ObsRut", Trim(TextBox65.Text))
                cmd27.Parameters.AddWithValue("@FecIde", DateTimePicker12.Value)
                cmd27.Parameters.AddWithValue("@CiRut", "0")
                cmd27.ExecuteNonQuery()
            Else
                Dim cmd27 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd27.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd27.Parameters.AddWithValue("@EtaRut", Trim(TextBox51.Text))
                cmd27.Parameters.AddWithValue("@NomRut", Trim(TextBox50.Text))
                cmd27.Parameters.AddWithValue("@FechRut", DateTimePicker9.Value)

                cmd27.Parameters.AddWithValue("@AsiRut", Trim(TextBox85.Text))
                cmd27.Parameters.AddWithValue("@ObsRut", Trim(TextBox65.Text))
                cmd27.Parameters.AddWithValue("@FecIde", DateTimePicker12.Value)
                cmd27.Parameters.AddWithValue("@CiRut", "1")
                cmd27.ExecuteNonQuery()
            End If

            If Trim(TextBox92.Text) = "0" Or Trim(TextBox92.Text) = "2" Then
                Dim cmd28 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd28.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd28.Parameters.AddWithValue("@EtaRut", Trim(TextBox53.Text))
                cmd28.Parameters.AddWithValue("@NomRut", Trim(TextBox52.Text))
                cmd28.Parameters.AddWithValue("@FechRut", DateTimePicker10.Value)
                cmd28.Parameters.AddWithValue("@FecIde", DateTimePicker11.Value)
                cmd28.Parameters.AddWithValue("@AsiRut", "")
                cmd28.Parameters.AddWithValue("@ObsRut", Trim(TextBox66.Text))
                cmd28.Parameters.AddWithValue("@CiRut", "0")
                cmd28.ExecuteNonQuery()
            Else
                Dim cmd28 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd28.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd28.Parameters.AddWithValue("@EtaRut", Trim(TextBox53.Text))
                cmd28.Parameters.AddWithValue("@NomRut", Trim(TextBox52.Text))
                cmd28.Parameters.AddWithValue("@FechRut", DateTimePicker10.Value)
                cmd28.Parameters.AddWithValue("@FecIde", DateTimePicker11.Value)
                cmd28.Parameters.AddWithValue("@AsiRut", "")
                cmd28.Parameters.AddWithValue("@ObsRut", Trim(TextBox66.Text))
                cmd28.Parameters.AddWithValue("@CiRut", "1")
                cmd28.ExecuteNonQuery()
            End If

            If Trim(TextBox93.Text) = "0" Or Trim(TextBox93.Text) = "2" Then
                Dim cmd29 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd29.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd29.Parameters.AddWithValue("@EtaRut", Trim(TextBox55.Text))
                cmd29.Parameters.AddWithValue("@NomRut", Trim(TextBox54.Text))
                cmd29.Parameters.AddWithValue("@FechRut", DateTimePicker11.Value)
                cmd29.Parameters.AddWithValue("@FecIde", DateTimePicker11.Value)
                cmd29.Parameters.AddWithValue("@AsiRut", "")
                cmd29.Parameters.AddWithValue("@ObsRut", Trim(TextBox56.Text))
                cmd29.Parameters.AddWithValue("@CiRut", "0")
                cmd29.ExecuteNonQuery()
            Else
                Dim cmd29 As New SqlCommand("insert into Ruta_Udp (OdRut,EtaRut,NomRut,FechRut,AsiRut,ObsRut,CiRut,FecIde) values (@OdRut,@EtaRut,@NomRut,@FechRut,@AsiRut,@ObsRut,@CiRut,@FecIde) ", conx)
                cmd29.Parameters.AddWithValue("@OdRut", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd29.Parameters.AddWithValue("@EtaRut", Trim(TextBox55.Text))
                cmd29.Parameters.AddWithValue("@NomRut", Trim(TextBox54.Text))
                cmd29.Parameters.AddWithValue("@FechRut", DateTimePicker11.Value)
                cmd29.Parameters.AddWithValue("@FecIde", DateTimePicker11.Value)
                cmd29.Parameters.AddWithValue("@AsiRut", "")
                cmd29.Parameters.AddWithValue("@ObsRut", Trim(TextBox56.Text))
                cmd29.Parameters.AddWithValue("@CiRut", "1")
                cmd29.ExecuteNonQuery()
            End If



        'MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
        'mostrar_rutas()
        '    bloquear_proceso()
        'End If
        Dim cmd31 As New SqlCommand("update custom_vianny.dbo.qag0300 set analista_3 =@analista,modelista_3 =@modelista where ncom_3 =@ncom_3 and ccia ='01'", conx)
        cmd31.Parameters.AddWithValue("@analista", ComboBox4.Text.ToString().Trim())
        cmd31.Parameters.AddWithValue("@modelista", ComboBox6.Text.ToString().Trim())
        cmd31.Parameters.AddWithValue("@ncom_3", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim())
        cmd31.ExecuteNonQuery()

        MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
        mostrar_rutas()
        bloquear_proceso()
        Cierre_Proc_Od.Button3.PerformClick()
    End Sub
    Sub bloquear_proceso()
        DateTimePicker4.Enabled = False
        DateTimePicker5.Enabled = False
        DateTimePicker6.Enabled = False
        DateTimePicker7.Enabled = False
        DateTimePicker8.Enabled = False
        DateTimePicker9.Enabled = False
        DateTimePicker10.Enabled = False
        DateTimePicker12.Enabled = False
        TextBox60.Enabled = False
        TextBox61.Enabled = False
        TextBox62.Enabled = False
        TextBox63.Enabled = False
        TextBox64.Enabled = False
        TextBox65.Enabled = False
        TextBox66.Enabled = False

        CheckBox1.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        CheckBox5.Enabled = False
        CheckBox6.Enabled = False
        CheckBox7.Enabled = False

    End Sub
    Sub habilitar_proceso()

        If Trim(TextBox57.Text) = "0" Or Trim(TextBox57.Text) = "2" Then
            CheckBox1.Enabled = True
            CheckBox1.Checked = True
        Else
            CheckBox1.Enabled = False
            CheckBox1.Checked = False
        End If
        If Trim(TextBox67.Text) = "0" Or Trim(TextBox67.Text) = "2" Then
            CheckBox2.Enabled = True
            CheckBox2.Checked = True
        Else
            CheckBox2.Enabled = False
            CheckBox2.Checked = False
        End If
        If Trim(TextBox68.Text) = "0" Or Trim(TextBox68.Text) = "2" Then
            CheckBox3.Enabled = True
            CheckBox3.Checked = True
        Else
            CheckBox3.Enabled = False
            CheckBox3.Checked = False
        End If
        If Trim(TextBox89.Text) = "0" Or Trim(TextBox89.Text) = "2" Then
            CheckBox4.Enabled = True
            CheckBox4.Checked = True
        Else
            CheckBox4.Enabled = False
            CheckBox4.Checked = False
        End If
        If Trim(TextBox90.Text) = "0" Or Trim(TextBox90.Text) = "2" Then
            CheckBox5.Enabled = True
            CheckBox5.Checked = True
        Else
            CheckBox5.Enabled = False
            CheckBox5.Checked = False
        End If
        If Trim(TextBox91.Text) = "0" Or Trim(TextBox91.Text) = "2" Then
            CheckBox6.Enabled = True
            CheckBox6.Checked = True
        Else
            CheckBox6.Enabled = False
            CheckBox6.Checked = False
        End If
        If Trim(TextBox92.Text) = "0" Or Trim(TextBox92.Text) = "2" Then
            CheckBox7.Enabled = True
            CheckBox7.Checked = True
        Else
            CheckBox7.Enabled = False
            CheckBox7.Checked = False
        End If
        If Trim(TextBox93.Text) = "0" Or Trim(TextBox93.Text) = "2" Then
            CheckBox8.Enabled = True
            CheckBox8.Checked = True
        Else
            CheckBox8.Enabled = False
            CheckBox8.Checked = False
        End If


    End Sub
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        habilitar_proceso()
        GroupBox5.Enabled = True
        Button20.Enabled = True
    End Sub

    Private Sub TextBox80_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Sub mostrar_telas()
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "select isnull(tela2_3,''),isnull(c.nomb_17,''),ISNULL( tela3_3,''),isnull(c1.nomb_17,''), isnull(tela4_3,''),isnull(c2.nomb_17,'') as nombre4, isnull(lavadoten_3,''),
			   isnull(otros1_3,''),isnull(otros2_3,''),isnull(estimg_3,''),isnull(otros3_3,''),tela1_3,telaprinc_3,descri_3,program_3,observ_3,mruta1_3,a.dele,q.tallador_3,Q.cants_3 from  
                                custom_vianny.DBO.qag0300  q 
                                left join custom_vianny.dbo.cag1700 c on q.ccia = c.ccia and q.tela2_3 = c.linea_17
                                left join custom_vianny.dbo.cag1700 c1 on q.ccia = c1.ccia and q.tela3_3 = c1.linea_17
                                left join custom_vianny.dbo.cag1700 c2 on q.ccia = c2.ccia and q.tela4_3 = c2.linea_17
                                left join custom_vianny.dbo.TAB0100 A on A.ccia + A.ctab='01MARCLI' and a.cele = q.marcacli_3
                                where ncom_3 = '" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' and q.ccia ='01'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        If Rsr1991.Read() Then
            TextBox26.Text = Rsr1991(0)
            TextBox27.Text = Rsr1991(1)
            TextBox28.Text = Rsr1991(2)
            TextBox29.Text = Rsr1991(3)
            TextBox30.Text = Rsr1991(4)
            TextBox31.Text = Rsr1991(5)
            TextBox35.Text = Rsr1991(6)
            TextBox36.Text = Rsr1991(7)
            TextBox37.Text = Rsr1991(8)
            TextBox38.Text = Rsr1991(9)
            TextBox24.Text = Rsr1991(11)
            TextBox25.Text = Rsr1991(12)
            TextBox6.Text = Rsr1991(12)
            TextBox39.Text = Rsr1991(13)
            TextBox5.Text = Rsr1991(13)
            TextBox94.Text = Rsr1991(14)
            TextBox8.Text = Rsr1991(15)
            TextBox9.Text = Rsr1991(16)
            TextBox10.Text = Rsr1991(17)
            Label12.Text = Rsr1991(18)
            TextBox12.Text = Rsr1991(19)
            If Rsr1991(10) = "BASICO" Then
                ComboBox5.SelectedIndex = 0
            Else
                If Rsr1991(10) = "SEMI MODA" Then
                    ComboBox5.SelectedIndex = 1
                Else
                    ComboBox5.SelectedIndex = 2
                End If
            End If
            GroupBox2.Enabled = False
            GroupBox3.Enabled = False
            GroupBox4.Enabled = False
            Button19.Enabled = False
        End If
        Rsr1991.Close()
    End Sub

    Private Sub TextBox84_TextChanged(sender As Object, e As EventArgs) Handles TextBox84.TextChanged

    End Sub

    Private Sub Od_Udp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox60.Text = ""
        TextBox61.Text = ""
        TextBox62.Text = ""
        TextBox63.Text = ""
        TextBox64.Text = ""
        TextBox65.Text = ""
        TextBox66.Text = ""
        TextBox56.Text = ""

        TextBox83.Text = ""
        TextBox84.Text = ""
        TextBox85.Text = ""
        TextBox57.Text = ""
        TextBox67.Text = ""
        TextBox68.Text = ""
        TextBox89.Text = ""
        TextBox90.Text = ""
        TextBox91.Text = ""
        TextBox92.Text = ""
        TextBox93.Text = ""
        Button26.Image = Nothing
        Button27.Image = Nothing
        Button28.Image = Nothing
        Button29.Image = Nothing
        Button30.Image = Nothing
        Button31.Image = Nothing
        Button32.Image = Nothing
        Button33.Image = Nothing
        DateTimePicker4.Value = DateTime.Now
        DateTimePicker5.Value = DateTime.Now
        DateTimePicker6.Value = DateTime.Now
        DateTimePicker7.Value = DateTime.Now
        DateTimePicker8.Value = DateTime.Now
        DateTimePicker9.Value = DateTime.Now
        DateTimePicker10.Value = DateTime.Now
        DateTimePicker11.Value = DateTime.Now
        DateTimePicker12.Value = DateTime.Now
        ComboBox2.SelectedIndex = 0
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        CheckBox6.Checked = False
        CheckBox7.Checked = False
        CheckBox8.Checked = False
        mostrar_rutas()
        mostrar_telas()
        mostrar_talla()
        FTY.Clear()
    End Sub
    Dim ll, FTY As DataTable
    Sub mostrar_talla()
        DataGridView4.DataSource = Nothing

        Dim ab As New vtallas
        Dim ab1 As New Padron_tallas
        ab.gcodigo = Trim(Label12.Text)
        ab.gccia = "01"
        FTY = ab1.mostrar_padrom_tallas_SELECCIONADO(ab)
        DataGridView4.DataSource = FTY
        TextBox11.Text = Trim(DataGridView4.Rows(0).Cells(13).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(14).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(15).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(16).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(17).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(18).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(19).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(20).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(21).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(22).Value) & "/" & Trim(DataGridView4.Rows(0).Cells(23).Value)

    End Sub
    Private Sub TextBox85_TextChanged(sender As Object, e As EventArgs) Handles TextBox85.TextChanged

    End Sub
    Function ELIMINAR(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Sub eliminar_Rutas()
        Dim agregar As String = "delete Ruta_Udp where OdRut ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "'"
        ELIMINAR(agregar)
    End Sub
    Private Sub TextBox80_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.F1 Then
            Detalle_ruta.Label2.Text = 3
            Detalle_ruta.ShowDialog()
        End If
    End Sub
    Sub bloquear()
        TextBox24.Enabled = False
        TextBox26.Enabled = False
        TextBox28.Enabled = False
        TextBox30.Enabled = False
        Button16.Enabled = False
        ComboBox5.Enabled = False
        TextBox38.Enabled = False
    End Sub
    Sub desbloquear()
        TextBox24.Enabled = True
        TextBox26.Enabled = True
        TextBox28.Enabled = True
        TextBox30.Enabled = True
        Button16.Enabled = True
        ComboBox5.Enabled = True
        TextBox38.Enabled = True
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        abrir()
        Dim cmd20 As New SqlCommand("update custom_vianny.dbo.qag0300 set tela1_3=@tela1_3,telaprinc_3=@telaprinc_3, tela2_3=@tela2_3, tela3_3=@tela3_3,tela4_3=@tela4_3,lavadoten_3=@lavadoten_3, otros1_3 =@otros1_3,otros2_3=@otros2_3, estimg_3=@estimg_3,otros3_3=@otros3_3  where ccia =@ccia  and ncom_3 =@ncom_3 ", conx)
        cmd20.Parameters.AddWithValue("@tela1_3", Trim(TextBox24.Text))
        cmd20.Parameters.AddWithValue("@telaprinc_3", Trim(TextBox25.Text))
        cmd20.Parameters.AddWithValue("@tela2_3", Trim(TextBox26.Text))
        cmd20.Parameters.AddWithValue("@tela3_3", Trim(TextBox28.Text))
        cmd20.Parameters.AddWithValue("@tela4_3", Trim(TextBox30.Text))
        cmd20.Parameters.AddWithValue("@lavadoten_3", Trim(TextBox35.Text))
        cmd20.Parameters.AddWithValue("@otros1_3", Trim(TextBox36.Text))
        cmd20.Parameters.AddWithValue("@otros2_3", Trim(TextBox37.Text))
        cmd20.Parameters.AddWithValue("@otros3_3", Trim(ComboBox5.Text))
        cmd20.Parameters.AddWithValue("@estimg_3", Trim(TextBox38.Text))
        cmd20.Parameters.AddWithValue("@ccia", "01")
        cmd20.Parameters.AddWithValue("@ncom_3", Trim(TextBox1.Text) & Trim(TextBox2.Text))
        cmd20.ExecuteNonQuery()
        MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
        bloquear()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        desbloquear()
        GroupBox2.Enabled = True
        GroupBox3.Enabled = True
        GroupBox4.Enabled = True
        Button19.Enabled = True
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Lavados.Label7.Text = "2"
        Lavados.ShowDialog()
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged

    End Sub

    Private Sub TextBox84_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox84.KeyDown
        If e.KeyCode = Keys.F1 Then
            Detalle_ruta.Label2.Text = 1
            Detalle_ruta.ShowDialog()
        End If
    End Sub

    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs) Handles TextBox26.TextChanged

    End Sub

    Private Sub TextBox85_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox85.KeyDown
        If e.KeyCode = Keys.F1 Then
            Detalle_ruta.Label2.Text = 2
            Detalle_ruta.ShowDialog()
        End If
    End Sub

    Private Sub TextBox28_TextChanged(sender As Object, e As EventArgs) Handles TextBox28.TextChanged

    End Sub

    Private Sub TextBox24_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = "01"
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent1.Label5.Text
            pproductos.FECHA.Text = Replace(DateTimePicker3.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "08"
            pproductos.Label3.Text = 7

            pproductos.ShowDialog()
        End If
    End Sub

    Private Sub TextBox30_TextChanged(sender As Object, e As EventArgs) Handles TextBox30.TextChanged

    End Sub

    Private Sub TextBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox26.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = "01"
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent1.Label5.Text
            pproductos.FECHA.Text = Replace(DateTimePicker3.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "08"
            pproductos.Label3.Text = 8

            pproductos.ShowDialog()
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        DataGridView3.Rows.Add()
        Dim P As Integer
        P = DataGridView3.Rows.Count
        DataGridView3.Rows(P - 1).Cells(0).Value = P
        DataGridView3.Rows(P - 1).Cells(1).Value = ComboBox2.Text
    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView3.Rows.Remove(DataGridView1.CurrentRow)
            I18 = DataGridView3.Rows.Count
            For i1 = 0 To I18 - 1
                VAL = i1 + 1
                DataGridView3.Rows(i1).Cells(0).Value = VAL

            Next
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Process.Start(Trim(TextBox9.Text))
    End Sub

    Private Sub TextBox28_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox28.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = "01"
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent1.Label5.Text
            pproductos.FECHA.Text = Replace(DateTimePicker3.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "08"
            pproductos.Label3.Text = 9

            pproductos.ShowDialog()
        End If
    End Sub

    Private Sub TextBox30_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox30.KeyDown
        If e.KeyCode = Keys.F1 Then
            pproductos.CCIA.Text = "01"
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = MDIParent1.Label5.Text
            pproductos.FECHA.Text = Replace(DateTimePicker3.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "08"
            pproductos.Label3.Text = 10

            pproductos.ShowDialog()
        End If
    End Sub
End Class