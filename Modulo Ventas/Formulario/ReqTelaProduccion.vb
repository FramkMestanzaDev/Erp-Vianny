Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class ReqTelaProduccion
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public comando As SqlCommand
    Dim da As New DataTable
    Dim Rsr2, Rsr3 As SqlDataReader
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
where ncom_3 ='" + Trim(TextBox1.Text) + "' and q.ccia ='" + Trim(Label5.Text) + "'
"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            If Rsr2.Read() = True Then
                TextBox2.Text = Rsr2(1)
                TextBox3.Text = Rsr2(2)
            End If
            Rsr2.Close()
            ComboBox1.Select()

        End If
        If e.KeyCode = Keys.F1 Then
            ListaOD.Label4.Text = 5
            ListaOD.ShowDialog()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If Convert.ToInt32(TextBox1.Text.ToString.Trim.Length) > 0 Then
            If TextBox5.Text.ToString.Trim.Length = 0 Then
                Dim cmd20 As New SqlCommand("insert into RequerimientoTelaProd (OpRtp,PoRtp,ProRtp,FEnRtp,TAdRtp,CanRtp,EstRtp,TallRtp,ObsRpt) values (@OpRtp,@PoRtp,@ProRtp,@FEnRtp,@TAdRtp,@CanRtp,@EstRtp,@TallRtp,@ObsRpt)", conx)
                cmd20.Parameters.AddWithValue("@OpRtp", Trim(TextBox1.Text))
                cmd20.Parameters.AddWithValue("@PoRtp", Trim(TextBox2.Text))
                If Mid(TextBox1.Text.ToString().Trim(), 1, 2) = "03" Or CheckBox2.Checked = True Then
                    cmd20.Parameters.AddWithValue("@ProRtp", Trim(TextBox6.Text))
                Else
                    cmd20.Parameters.AddWithValue("@ProRtp", Trim(TextBox3.Text))
                End If

                cmd20.Parameters.AddWithValue("@FEnRtp", DateTimePicker1.Value)
                If CheckBox1.Checked = True Then
                    cmd20.Parameters.AddWithValue("@TAdRtp", "1")
                Else
                    cmd20.Parameters.AddWithValue("@TAdRtp", "0")
                End If
                If Mid(TextBox1.Text.ToString().Trim(), 1, 2) = "03" Or CheckBox2.Checked = True Then
                    cmd20.Parameters.AddWithValue("@CanRtp", Convert.ToDouble(TextBox7.Text))
                Else
                    cmd20.Parameters.AddWithValue("@CanRtp", Convert.ToDouble(TextBox4.Text))
                End If

                cmd20.Parameters.AddWithValue("@EstRtp", "0")
                cmd20.Parameters.AddWithValue("@TallRtp", Trim(ComboBox1.Text.ToString))
                cmd20.Parameters.AddWithValue("@ObsRpt", Trim(TextBox10.Text.ToString))
                cmd20.ExecuteNonQuery()
                MsgBox("Se Registro la Informacion Correctamnete")
                Listar_Informacion()
                limpiar()
            Else
                Dim cmd21 As New SqlCommand("update RequerimientoTelaProd set ObsRpt=@ObsRpt ,FEnRtp = @FEnRtp,CanRtp =@CanRtp,TAdRtp=@TAdRtp,ProRtp =@ProRtp,TallRtp =@TallRtp,PoRtp = @PoRtp where IdRtp =@IdRtp", conx)
                cmd21.Parameters.AddWithValue("@OpRtp", Trim(TextBox1.Text))
                cmd21.Parameters.AddWithValue("@PoRtp", Trim(TextBox2.Text))
                cmd21.Parameters.AddWithValue("@ProRtp", Trim(TextBox3.Text))
                cmd21.Parameters.AddWithValue("@FEnRtp", DateTimePicker1.Value)
                If CheckBox1.Checked = True Then
                    cmd21.Parameters.AddWithValue("@TAdRtp", "1")
                Else
                    cmd21.Parameters.AddWithValue("@TAdRtp", "0")
                End If

                cmd21.Parameters.AddWithValue("@TallRtp", Trim(ComboBox1.Text.ToString))
                cmd21.Parameters.AddWithValue("@IdRtp", Trim(TextBox5.Text.ToString))
                cmd21.Parameters.AddWithValue("@ObsRpt", Trim(TextBox10.Text.ToString))
                If Mid(TextBox1.Text.ToString().Trim(), 1, 2) = "03" Or CheckBox2.Checked = True Then
                    cmd21.Parameters.AddWithValue("@CanRtp", Convert.ToDouble(TextBox7.Text))
                Else
                    cmd21.Parameters.AddWithValue("@CanRtp", Convert.ToDouble(TextBox4.Text))
                End If
                cmd21.ExecuteNonQuery()
                MsgBox("Se Registro la Informacion Correctamnete")
                Listar_Informacion()
                limpiar()
            End If
            TextBox1.Enabled = True
            TextBox1.Select()
        Else
                MsgBox("Falta Ingresar Informacion en el Campo Op")
        End If

    End Sub
    Sub limpiar()
        CheckBox2.Checked = False
        CheckBox1.Checked = False
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = "0"
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = "0"
        TextBox10.Text = ""
        CheckBox1.Checked = False
    End Sub
    Sub Listar_Informacion()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select FEnRtp as FEntergaCorte,case when SUBSTRING( OpRtp,1,2)='01' then OpRtp else SUBSTRING( OpRtp,1,9) end as Op,PoRtp as Po,ProRtp as Producto,case when TAdRtp=0 then 'NO' ELSE 'SI' END as 'Tela Adicional', CanRtp as Cantidad,case when CAST( CanRtp as decimal) = 0 then 'PRODUCCION'   ELSE 'UDP' END as Area,TallRtp as Taller,IdRtp,SelRtp,ObsRpt  AS OBSERVACION from RequerimientoTelaProd WHERE EstRtp  IN (0,1)", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(5).Width = 400
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 85
        DataGridView1.Columns(4).Width = 85
        DataGridView1.Columns(6).Width = 65
        DataGridView1.Columns(7).Width = 65
        DataGridView1.Columns(8).Width = 85
        DataGridView1.Columns(9).Width = 115
        DataGridView1.Columns(12).Width = 80
        DataGridView1.Columns(10).Visible = False
        DataGridView1.Columns(11).Visible = False
        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(11).Value.ToString.Trim = "1" Then
                DataGridView1.Rows(i).Cells(1).Value = True
            End If
        Next
    End Sub
    Sub Listar_Informacion_Despachado()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select FEnRtp as FEntergaCorte,case when SUBSTRING( OpRtp,1,2)='01' then OpRtp else SUBSTRING( OpRtp,1,9) end as Op,PoRtp as Po,ProRtp as Producto,case when TAdRtp=0 then 'NO' ELSE 'SI' END as 'Tela Adicional', CanRtp as Cantidad,case when CAST( CanRtp as decimal) = 0 then 'PRODUCCION'   ELSE 'UDP' END as Area,TallRtp as Taller,IdRtp,SelRtp,ObsRpt AS OBSERVACION from RequerimientoTelaProd WHERE EstRtp  IN (2)", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(5).Width = 400
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 85
        DataGridView1.Columns(4).Width = 85
        DataGridView1.Columns(6).Width = 65
        DataGridView1.Columns(7).Width = 65
        DataGridView1.Columns(8).Width = 85
        DataGridView1.Columns(9).Width = 115
        DataGridView1.Columns(12).Width = 80
        DataGridView1.Columns(10).Visible = False
        DataGridView1.Columns(11).Visible = False
        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(11).Value.ToString.Trim = "1" Then
                DataGridView1.Rows(i).Cells(1).Value = True
            End If
        Next
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox1.Checked = True Then
            TextBox4.Enabled = True
            TextBox4.Select()
        Else
            TextBox4.Enabled = False
        End If
    End Sub

    Private Sub ReqTelaProduccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Listar_Informacion()
        If Label10.Text.ToString.Trim = "PRODUCCION" Then
            GroupBox1.Enabled = True
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Visible = False
            TextBox7.Visible = False
            TextBox6.Visible = False
            Label12.Visible = False
            Label11.Visible = False
            Label13.Visible = False
            Label14.Visible = False
            Button2.Enabled = True
            Button9.Enabled = True
            Button8.Visible = False
            Button6.Visible = False
            Button5.Visible = False
        Else
            If Label10.Text.ToString.Trim = "UDP" Then
                GroupBox1.Enabled = True
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(1).Visible = False
                Button2.Enabled = True
                TextBox4.Enabled = False
                TextBox7.Visible = True
                TextBox7.Enabled = True
                TextBox6.Visible = True
                Label12.Visible = True
                Label11.Visible = True
                Label13.Visible = True
                Label14.Visible = True
                Button9.Enabled = True
                Button8.Visible = False
                Button6.Visible = False
                Button5.Visible = False
            Else
                GroupBox1.Enabled = False
                DataGridView1.Columns(0).Visible = True
                DataGridView1.Columns(1).Visible = True
                Button2.Enabled = False

                Button8.Visible = True
                Button6.Visible = True
                Button5.Visible = True
            End If

        End If
        ComboBox1.SelectedIndex = 0
        RadioButton1.Select()
        TextBox1.Select()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox8.Text = ""
        TextBox9.Text = ""
        If RadioButton1.Checked = True Then
            Listar_Informacion()
        Else
            Listar_Informacion_Despachado()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "TE" Then
            If Mid(DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString.Trim, 1, 2) = "03" Then
                MsgBox("ESTE REQUERIMIENTO ES UNA MUESTRA DE UDP POR LO QUE SE DESPACHA LO QUE APARECE NE LA DESCRIPCION")
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("DESEA CERRAR ESTE REQUERIMIENTO DE TELA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Dim cmd21 As New SqlCommand(" UPDATE RequerimientoTelaProd SET EstRtp = '2' WHERE IdRtp =@IdRtp", conx)
                    cmd21.Parameters.AddWithValue("@IdRtp", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value))
                    cmd21.ExecuteNonQuery()
                    MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
                    Listar_Informacion()
                End If
            Else
                    DetTelaProduccion.Label2.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString.Trim
                DetTelaProduccion.Label3.Text = Label5.Text
                DetTelaProduccion.Label4.Text = e.RowIndex
                DetTelaProduccion.Label5.Text = DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString.Trim
                DetTelaProduccion.ShowDialog()
            End If

        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "GUIA" Then
            If Mid(DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString.Trim, 1, 2) = "03" Then
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = True

            End If

        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Po" Then

            Dim valorcelda As String
            valorcelda = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString().Trim()
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OBSERVACION" Then
            Observacion2.TextBox1.Text = ""
            Observacion2.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(12).Value.ToString.Trim
            Observacion2.Label1.Text = "1"
            Observacion2.ShowDialog()
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString.Trim
        TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value.ToString.Trim
        TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString.Trim

        TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString.Trim
        DateTimePicker1.Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString.Trim
        If DataGridView1.Rows(e.RowIndex).Cells(9).Value.ToString.Trim = "TEXTIL HEGAR" Then
            ComboBox1.SelectedIndex = 0
        Else
            If DataGridView1.Rows(e.RowIndex).Cells(9).Value.ToString.Trim = "UDP" Then
                ComboBox1.SelectedIndex = 2
            Else
                ComboBox1.SelectedIndex = 1
            End If
        End If
        If Mid(DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString.Trim, 1, 2) = "01" And Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(7).Value) > 0 Then
            CheckBox2.Checked = True
            TextBox7.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value
            TextBox4.Text = "0"
        End If
        If Mid(DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString.Trim, 1, 2) = "03" Then
            TextBox7.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value
            CheckBox2.Checked = True
            TextBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value
        Else
            TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value
        End If
    End Sub

    Private Sub ReqTelaProduccion_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim id As String
        id = ""
        Dim alm As New Almacen_guia_nuevo
        alm.Label1.Text = Label8.Text
        alm.Label2.Text = Label5.Text
        alm.Label3.Text = Label9.Text
        alm.Label4.Text = "1"
        alm.Label6.Text = "1"
        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(1).Value = True Then
                If id = "" Then
                    id = DataGridView1.Rows(i).Cells(10).Value
                Else
                    id = id & "," & DataGridView1.Rows(i).Cells(10).Value
                End If
            End If
        Next
        alm.Label5.Text = id
        alm.ShowDialog()
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim id As String
        id = ""

        Dim selectedOp As String = DataGridView1.SelectedRows(0).Cells("Op").Value.ToString()


        If selectedOp.StartsWith("03") Then
            Nsalida.abrir()
            Nsalida.TextBox7.Text = Modulo_Almacen.Label6.Text.ToString().Trim()
            Nsalida.Label10.Text = Modulo_Almacen.Label8.Text.ToString().Trim()
            Nsalida.Label13.Text = Label5.Text.ToString().Trim()
            Nsalida.Label15.Text = "0"

            Nsalida.Show()
        Else
            Nsalida.abrir()
            Nsalida.TextBox7.Text = Modulo_Almacen.Label6.Text.ToString().Trim()
            Nsalida.Label10.Text = Modulo_Almacen.Label8.Text.ToString().Trim()
            Nsalida.Label13.Text = Label5.Text.ToString().Trim()
            Nsalida.TextBox4.Text = DateTimePicker1.Value.Month
            Nsalida.EstablecerAlmacen("08")
            Select Case Nsalida.TextBox4.Text.Length
                Case "1" : Nsalida.TextBox4.Text = "0" & DateTimePicker1.Value.Month
                Case "2" : Nsalida.TextBox4.Text = NotaIngreso.TextBox4.Text
            End Select
            'Nsalida.correlativo()
            'Nsalida.RecibirDatos(filasSeleccionadas)
            Nsalida.Label15.Text = "0"
            For i = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(1).Value = True Then
                    If id = "" Then
                        id = DataGridView1.Rows(i).Cells(10).Value
                    Else
                        id = "," & DataGridView1.Rows(i).Cells(10).Value
                    End If
                End If
            Next
            Nsalida.Label16.Text = id
            Nsalida.TextBox5.Focus()
            SendKeys.Send("{ENTER}")
            Nsalida.ShowDialog()
        End If



        'End If



    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = "0"
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox6.Text = "0"
        TextBox1.Enabled = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim result As DialogResult = MessageBox.Show("¿Seguro que desea Cerar la fila seleccionada?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then

            Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)


            Dim cmd16 As New SqlCommand("update RequerimientoTelaProd  set EstRtp='2' where IdRtp=@IdRtp", conx)
            cmd16.Parameters.AddWithValue("@IdRtp", filaSeleccionada.Cells("IdRtp").Value)

                cmd16.ExecuteNonQuery()
                MessageBox.Show("Se cerro el requerimiento de Tela", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Button1.PerformClick()
            Else
                ' Si no hay una fila seleccionada, mostrar un mensaje al usuario
                MessageBox.Show("Por favor, seleccione una fila para cerrar el requerimiento.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    Sub mostrar_tela()
        abrir()
        Dim sql102 As String = "select telaprinc_3 from custom_vianny.dbo.qag0300 where ncom_3 ='" + TextBox1.Text.ToString().Trim() + "' and ccia ='01' "
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr3 = cmd102.ExecuteReader()
        If Rsr3.Read() = True Then
            TextBox6.Text = Rsr3(0)
        End If
        Rsr3.Close()
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        buscar()
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer
            dv.RowFilter = "Po" & " like '%" & TextBox8.Text & "%'"
            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(5).Width = 400
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 85
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Width = 65
                DataGridView1.Columns(8).Width = 85
                DataGridView1.Columns(9).Width = 130
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                For i = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Rows(i).Cells(11).Value.ToString.Trim = "1" Then
                        DataGridView1.Rows(i).Cells(1).Value = True
                    End If
                Next
            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception


        End Try
    End Sub
    Function ELIMINAR(ByVal sql As String, ByVal idRtp As Integer) As Boolean
        Try
            abrir()
            Dim comando As New SqlCommand(sql, conx)
            comando.Parameters.AddWithValue("@IdRtp", idRtp)
            Dim i As Integer = comando.ExecuteNonQuery()
            If i > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox("Error al eliminar: " & ex.Message)
            Return False
        End Try
    End Function
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim respuesta As DialogResult
        Dim TAB As Integer
        TAB = DataGridView1.Rows.Count
        If TAB <> 0 Then
            respuesta = MessageBox.Show("¿DESEA ELIMINAR LA FILA?", "SALIR DEL PROGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If respuesta = Windows.Forms.DialogResult.Yes Then
                abrir()
                Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim idRtp As Integer = Convert.ToInt32(filaSeleccionada.Cells("IdRtp").Value)

                ' Usa una consulta parametrizada para evitar inyección SQL y errores de conversión
                Dim sqlEliminar As String = "DELETE FROM RequerimientoTelaProd WHERE IdRtp = @IdRtp"
                If ELIMINAR(sqlEliminar, idRtp) Then
                    DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
                Else
                    MsgBox("Error al eliminar la fila.")
                End If
            End If
        Else
            MsgBox("NO HAY PRODUCTOS A ELIMINAR")
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim seleccionado As DataGridViewRow = DataGridView1.SelectedRows(0)
        EnviarCorreo.TextBox1.Text = ""
        EnviarCorreo.TextBox2.Text = ""
        EnviarCorreo.TextBox1.Text = "Para Informarles que la Tela  " + seleccionado.Cells("Producto").Value.ToString().Trim() + "  Solicitada de la op  " + seleccionado.Cells("Op").Value.ToString().Trim() + "  no se tiene en Stock, Informar si se comprara o se cambiara por otra tela"
        EnviarCorreo.Label3.Text = Modulo_Almacen.Label6.Text.ToString().Trim()
        EnviarCorreo.Label5.Text = "0"
        EnviarCorreo.Label4.Text = seleccionado.Cells("Op").Value.ToString().Trim()
        EnviarCorreo.ShowDialog()
    End Sub
    Private Sub buscarAREA()
        Try
            Dim ds1 As New DataSet
            Dim K1 As Integer
            ds1.Tables.Add(da.Copy)
            Dim dv1 As New DataView(ds1.Tables(0))
            Dim jk1 As Integer
            dv1.RowFilter = "Area" & " like '%" & TextBox9.Text & "%'"
            If dv1.Count <> 0 Then

                DataGridView1.DataSource = dv1
                DataGridView1.Columns(5).Width = 400
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 85
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Width = 65
                DataGridView1.Columns(8).Width = 85
                DataGridView1.Columns(9).Width = 130
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                For i = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Rows(i).Cells(11).Value.ToString.Trim = "1" Then
                        DataGridView1.Rows(i).Cells(1).Value = True
                    End If
                Next
            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        buscarAREA()
    End Sub

    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button2.Select()
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)
        MsgBox("PASO")
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox4.Enabled = True
        Else
            TextBox4.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Label12.Visible = True
            TextBox7.Visible = True
            TextBox6.ReadOnly = False
            mostrar_tela()
            TextBox7.Select()
        End If
    End Sub
End Class