Imports System.Data.SqlClient

Public Class RequerimientoAvios
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public comando As SqlCommand
    Dim da As New DataTable
    Dim Rsr2, Rsr3, Rsr21 As SqlDataReader
    Dim loDataTable As New DataTable

    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()

            Select Case Trim(TextBox1.Text).Length
                Case "1" : TextBox1.Text = "01" & "0000000" & TextBox1.Text.ToString.Trim
                Case "2" : TextBox1.Text = "01" & "000000" & TextBox1.Text.ToString.Trim
                Case "3" : TextBox1.Text = "01" & "00000" & TextBox1.Text.ToString.Trim
                Case "4" : TextBox1.Text = "01" & "0000" & TextBox1.Text.ToString.Trim
                Case "5" : TextBox1.Text = "01" & "000" & TextBox1.Text.ToString.Trim
                Case "6" : TextBox1.Text = "01" & "00" & TextBox1.Text.ToString.Trim
                Case "7" : TextBox1.Text = "01" & "0" & TextBox1.Text.ToString.Trim
                Case "8" : TextBox1.Text = "01" & TextBox1.Text.ToString.Trim
                Case "9" : TextBox1.Text = "0" & TextBox1.Text.ToString.Trim
                Case "10" : TextBox1.Text = TextBox1.Text.ToString.Trim

            End Select
            Dim sql102 As String = "select ncom_3,program_3,ISNULL( a.nomb_17,q.descri_3) from custom_vianny.dbo.qag0300 q
            left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
            where ncom_3 ='" + Trim(TextBox1.Text) + "' and q.ccia ='" + Trim(Label5.Text) + "'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            If Rsr2.Read() = True Then
                TextBox2.Text = Rsr2(1)
                TextBox3.Text = Rsr2(2)
            End If
            Rsr2.Close()
            llenar_combo_box3()

        End If
        If e.KeyCode = Keys.F1 Then
            ListaOD.Label4.Text = 6
            ListaOD.ShowDialog()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)

    End Sub
    Sub asignar()
        abrir()
        Dim sql102213 As String = "select COUNT(op),proveedor from asignacion_cos_aca where area='02' and  op= '" + Trim(TextBox1.Text.ToString()) + "'  and corte ='" + ComboBox2.Text.ToString().Trim() + "' group by proveedor"
        Dim cmd102213 As New SqlCommand(sql102213, conx)
        Rsr3 = cmd102213.ExecuteReader()
        If Rsr3.Read() = True Then

            If Rsr3(0) > 0 Then
                MsgBox("YA EXISTE UN REGISTRO CON LA OP Y EL CORTE INGRESADOS ASIGNADOS  TALLER" + "      " + Rsr3(1))
                Rsr3.Close()
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("DESEAS REASIGNARLO LA OP?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Rsr3.Close()
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd16 As New SqlCommand("UPDATE asignacion_cos_aca SET proveedor =@proveedor,ruc=@ruc,Fecha =@Fecha WHERE area = @area AND op=@op AND corte =@corte", conx)
                    If ComboBox1.Text.ToString().Trim() = "ACABADOS" Then
                        cmd16.Parameters.AddWithValue("@area", "02")
                    Else
                        cmd16.Parameters.AddWithValue("@area", "01")
                    End If

                    cmd16.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
                    cmd16.Parameters.AddWithValue("@corte", ComboBox2.Text.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@ruc", TextBox4.Text().ToString().Trim())
                    cmd16.Parameters.AddWithValue("@proveedor", TextBox8.Text().ToString().Trim())
                    cmd16.Parameters.AddWithValue("@Fecha", DateTimePicker2.Value)
                    cmd16.ExecuteNonQuery()
                    MsgBox("Se Asigno el Proveedor a la Op")
                End If
            End If
        Else
            Rsr3.Close()
            Dim cmd16 As New SqlCommand("insert into asignacion_cos_aca (area,op,corte,ruc,proveedor,Fecha) values (@area,@op,@corte,@ruc,@proveedor,@Fecha)", conx)
            If ComboBox1.Text.ToString().Trim() = "ACABADOS" Then
                cmd16.Parameters.AddWithValue("@area", "02")
            Else
                cmd16.Parameters.AddWithValue("@area", "01")
            End If
            cmd16.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
            cmd16.Parameters.AddWithValue("@corte", ComboBox2.Text.ToString().Trim())
            cmd16.Parameters.AddWithValue("@ruc", TextBox4.Text().ToString().Trim())
            cmd16.Parameters.AddWithValue("@proveedor", TextBox8.Text().ToString().Trim())
            cmd16.Parameters.AddWithValue("@Fecha", DateTimePicker2.Value)
            cmd16.ExecuteNonQuery()
            MsgBox("Se Asigno el Proveedor a la Op")
        End If
    End Sub

    Sub llenar_combo_box3()
        Try
            loDataTable.Rows.Clear()

            ' Agregar fila vacía manualmente antes de ejecutar la consulta
            loDataTable.Columns.Clear()
            loDataTable.Columns.Add("ocorte_3cg", GetType(String))
            loDataTable.Rows.Add("")
            'loDataTable.Rows.Add("")
            Dim lsQuery As String = "select ocorte_3cg from custom_vianny.dbo.Qag300Cc where pedido_3cg ='" + Trim(TextBox1.Text) + "'"
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            ComboBox2.DataSource = loDataTable
            ComboBox2.DisplayMember = "ocorte_3cg"
            ComboBox2.ValueMember = "ocorte_3cg"
            ComboBox2.DropDownWidth = 90
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Convert.ToInt32(TextBox1.Text.ToString.Trim.Length) > 0 Then
            Dim valor As Integer
            Dim sql102 As String = "select COUNT(IdRAv) from RequerimientoAvios where OpRAv ='" + TextBox1.Text.ToString().Trim() + "' and CorRAv = '" + ComboBox2.Text.ToString().Trim() + "'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr21 = cmd102.ExecuteReader()
            If Rsr21.Read() = True Then
                valor = Convert.ToInt32(Rsr21(0))

            End If
            Rsr21.Close()
            If valor > 0 And TextBox5.Text.ToString().Trim().Length = 0 Then
                MsgBox("La Op y el Corte Ya esta Agregada en el Requerimiento de Avios")
            Else


                If TextBox5.Text.ToString.Trim.Length = 0 Then
                    Dim cmd20 As New SqlCommand("insert into RequerimientoAvios (OpRAv,PoRAv,ProRAv,FecRAv,AreRAv,RucRAv,RsoRAv,FeRRAv,CanRAv,EstRAv,CorRAv,PriRAv) values (@OpRAv,@PoRAv,@ProRAv,@FecRAv,@AreRAv,@RucRAv,@RsoRAv,@FeRRAv,@CanRAv,@EstRAv,@CorRAv,@PriRAv)", conx)
                    cmd20.Parameters.AddWithValue("@OpRAv", Trim(TextBox1.Text))
                    cmd20.Parameters.AddWithValue("@PoRAv", Trim(TextBox2.Text))
                    If Mid(TextBox1.Text.ToString().Trim(), 1, 2) = "03" Then
                        cmd20.Parameters.AddWithValue("@ProRAv", Trim(TextBox6.Text))
                    Else
                        cmd20.Parameters.AddWithValue("@ProRAv", Trim(TextBox3.Text))
                    End If
                    cmd20.Parameters.AddWithValue("@FecRAv", DateTimePicker1.Value)
                    cmd20.Parameters.AddWithValue("@AreRAv", ComboBox1.Text.ToString().Trim())
                    cmd20.Parameters.AddWithValue("@RucRAv", TextBox4.Text.ToString().Trim())
                    cmd20.Parameters.AddWithValue("@RsoRAv", TextBox8.Text.ToString().Trim())
                    cmd20.Parameters.AddWithValue("@FeRRAv", DateTimePicker2.Value)
                    cmd20.Parameters.AddWithValue("@CanRAv", Convert.ToDouble(TextBox7.Text))
                    cmd20.Parameters.AddWithValue("@EstRAv", "0")
                    cmd20.Parameters.AddWithValue("@CorRAv", ComboBox2.Text.ToString().Trim())
                    If ComboBox3.Text.ToString().Trim() = "NORMAL" Then
                        cmd20.Parameters.AddWithValue("@PriRAv", "1")
                    Else
                        cmd20.Parameters.AddWithValue("@PriRAv", "2")
                    End If

                    cmd20.ExecuteNonQuery()
                    MsgBox("Se Registro la Informacion Correctamnete")
                    asignar()
                    Listar_Informacion()
                    limpiar()
                Else
                    Dim cmd21 As New SqlCommand("update RequerimientoAvios set PriRAv=@PriRAv,OpRAv=@OpRAv,PoRAv=@PoRAv,ProRAv=@ProRAv,FecRAv=@FecRAv,AreRAv=@AreRAv,RucRAv=@RucRAv,RsoRAv=@RsoRAv,FeRRAv=@FeRRAv,CanRAv=@CanRAv,EstRAv=@EstRAv,CorRAv=@CorRAv where IdRAv = @IdRAv", conx)
                    cmd21.Parameters.AddWithValue("@OpRAv", Trim(TextBox1.Text))
                    cmd21.Parameters.AddWithValue("@PoRAv", Trim(TextBox2.Text))
                    If Mid(TextBox1.Text.ToString().Trim(), 1, 2) = "03" Then
                        cmd21.Parameters.AddWithValue("@ProRAv", Trim(TextBox6.Text))
                    Else
                        cmd21.Parameters.AddWithValue("@ProRAv", Trim(TextBox3.Text))
                    End If
                    cmd21.Parameters.AddWithValue("@FecRAv", DateTimePicker1.Value)
                    cmd21.Parameters.AddWithValue("@AreRAv", ComboBox1.Text.ToString().Trim())
                    cmd21.Parameters.AddWithValue("@RucRAv", TextBox4.Text.ToString().Trim())
                    cmd21.Parameters.AddWithValue("@RsoRAv", TextBox8.Text.ToString().Trim())
                    cmd21.Parameters.AddWithValue("@FeRRAv", DateTimePicker2.Value)
                    cmd21.Parameters.AddWithValue("@CanRAv", Convert.ToDouble(TextBox7.Text))
                    cmd21.Parameters.AddWithValue("@EstRAv", "0")
                    cmd21.Parameters.AddWithValue("@CorRAv", ComboBox2.Text.ToString().Trim())
                    cmd21.Parameters.AddWithValue("@IdRAv", TextBox5.Text.ToString().Trim())
                    If ComboBox3.Text.ToString().Trim() = "NOMRAL" Then
                        cmd21.Parameters.AddWithValue("@PriRAv", "1")
                    Else
                        cmd21.Parameters.AddWithValue("@PriRAv", "2")
                    End If
                    cmd21.ExecuteNonQuery()
                    MsgBox("Se Registro la Informacion Correctamnete")
                    asignar()
                    Listar_Informacion()
                    limpiar()
                End If
            End If
        Else
            MsgBox("Falta Ingresar Informacion en el Campo Op")
        End If
    End Sub
    Sub limpiar()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox8.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = "0"
    End Sub

    Private Sub RequerimientoAvios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Listar_Informacion()
        RadioButton1.Checked = True
        ComboBox1.SelectedIndex = 1
        If Label10.Text.ToString.Trim = "PRODUCCION" Then
            GroupBox1.Enabled = True
            DataGridView1.Columns(0).Visible = False
            TextBox7.Visible = False
            TextBox6.Visible = False
            Label12.Visible = False
            Label11.Visible = False
            Label13.Visible = False
            Label14.Visible = False
            Button2.Enabled = True
        Else
            If Label10.Text.ToString.Trim = "UDP" Then
                GroupBox1.Enabled = True
                DataGridView1.Columns(0).Visible = True
                Button2.Enabled = True
                TextBox4.Enabled = False
                TextBox7.Visible = True
                TextBox7.Enabled = True
                TextBox6.Visible = True
                Label12.Visible = True
                Label11.Visible = True
                Label13.Visible = True
                Label14.Visible = True
            Else
                GroupBox1.Enabled = False
                DataGridView1.Columns(0).Visible = True

                Button2.Enabled = False
            End If
        End If
        TextBox1.Select()
    End Sub
    Sub Listar_Informacion()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select FecRAv as FDespacho,OpRAv as Op,PoRAv as Po,ProRAv as Producto,
 CanRAv as Cantidad,case when EstRAv = '0' then 'PENDIENTE'  when EstRAv = '1' then 'DESPACHADO parcial' ELSE 'DESPACHADO' END as Estado,
RsoRAv as Taller,IdRAv,RucRAv,CorRAv as Corte,AreRAv as Area,case when PriRAv= '1' then 'NORMAL' else 'ALTA' end as Prioridad from RequerimientoAvios WHERE EstRAv=0 order by FecRAv", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(1).Width = 75
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).Width = 280
        DataGridView1.Columns(5).Width = 80
        DataGridView1.Columns(6).Width = 100
        DataGridView1.Columns(7).Width = 270
        DataGridView1.Columns(8).Visible = False
        DataGridView1.Columns(9).Visible = False
        DataGridView1.Columns(10).Width = 60
        DataGridView1.Columns(11).Width = 90
        DataGridView1.Columns(12).Width = 90
    End Sub
    Sub Listar_Informacion_Cerrada()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select FecRAv as FDespacho,OpRAv as Op,PoRAv as Po,ProRAv as Producto,
 CanRAv as Cantidad,case when EstRAv = '0' then 'PENDIENTE'  when EstRAv = '1' then 'DESPACHADO parcial' ELSE 'DESPACHADO' END as Estado,
RsoRAv as Taller,IdRAv,RucRAv,CorRAv as Corte,AreRAv as Area,case when PriRAv= '1' then 'NORMAL' else 'ALTA' end as Prioridad  from RequerimientoAvios WHERE EstRAv=1 order by FecRAv", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(1).Width = 90
        DataGridView1.Columns(1).Width = 90
        DataGridView1.Columns(4).Width = 280
        DataGridView1.Columns(5).Width = 90
        DataGridView1.Columns(6).Width = 120
        DataGridView1.Columns(7).Width = 280
        DataGridView1.Columns(8).Visible = False
        DataGridView1.Columns(9).Visible = False
        DataGridView1.Columns(10).Width = 60
        DataGridView1.Columns(11).Width = 90
        DataGridView1.Columns(12).Width = 90

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            Listar_Informacion()
        Else
            Listar_Informacion_Cerrada()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim selectedOp As String = DataGridView1.SelectedRows(0).Cells("Op").Value.ToString()


        If selectedOp.StartsWith("03") Then
            Nsalida.abrir()
            Nsalida.TextBox7.Text = Modulo_Almacen.Label6.Text.ToString().Trim()
            Nsalida.Label10.Text = Modulo_Almacen.Label8.Text.ToString().Trim()
            Nsalida.Label13.Text = Label5.Text.ToString().Trim()
            Nsalida.Label15.Text = "0"
            Nsalida.Show()
        End If

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CO" Then
            Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
            EnviarCorreo.TextBox1.Text = "Para Informarles los problemas que se tienes para realizar el armado de la op " + filaSeleccionada.Cells("Op").Value
            EnviarCorreo.Label3.Text = Label17.Text
            EnviarCorreo.Label4.Text = filaSeleccionada.Cells("Op").Value
            EnviarCorreo.Label5.Text = "1"
            EnviarCorreo.ShowDialog()
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString.Trim
        TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString.Trim
        TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value.ToString.Trim
        TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(9).Value.ToString.Trim
        TextBox8.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value.ToString.Trim
        TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(8).Value.ToString.Trim
        TextBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value.ToString.Trim
        TextBox7.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString.Trim
        If DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString.Trim.Length = 0 Then
            ComboBox2.Text = ""
        Else
            ComboBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString.Trim
        End If
        If DataGridView1.Rows(e.RowIndex).Cells(11).Value.ToString.Trim = "ACABADOS" Then
            ComboBox1.SelectedIndex = 1
        Else
            ComboBox1.SelectedIndex = 0
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim result As DialogResult = MessageBox.Show("¿Seguro que desea Cerar la fila seleccionada?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
            Dim cmd16 As New SqlCommand("update RequerimientoAvios set EstRAv ='1' where IdRAv =@IdRAv", conx)
            cmd16.Parameters.AddWithValue("@IdRAv", filaSeleccionada.Cells("IdRAv").Value)
            cmd16.ExecuteNonQuery()
            MessageBox.Show("Se cerro el requerimiento de Tela", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Button1.PerformClick()
        Else
            ' Si no hay una fila seleccionada, mostrar un mensaje al usuario
            MessageBox.Show("Por favor, seleccione una fila para cerrar el requerimiento.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        det_ns_Prodc.Label2.Text = TextBox1.Text
        det_ns_Prodc.Label3.Text = 14
        Select Case TextBox1.Text
            Case "0300" : det_ns_Prodc.Label4.Text = "1"
            Case "0400" : det_ns_Prodc.Label4.Text = "2"
            Case "0600" : det_ns_Prodc.Label4.Text = "3"
            Case "0500" : det_ns_Prodc.Label4.Text = "4"
        End Select
        det_ns_Prodc.Label5.Text = Trim(Label5.Text)
        det_ns_Prodc.Label1.Text = "PROVEEDOR"
        det_ns_Prodc.Show()
    End Sub
End Class