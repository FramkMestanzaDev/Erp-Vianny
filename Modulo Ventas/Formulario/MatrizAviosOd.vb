Imports System.Data.SqlClient
Public Class MatrizAviosOd
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim Rsr1991, rs1, rs2, rs3, rs4, rs6, rs5, rs7 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub clonar()
        abrir()
        TextBox9.Enabled = False
        Label14.Text = 1
        Select Case Trim(TextBox10.Text).Length
            Case "1" : TextBox10.Text = "01" & "0000000" & TextBox10.Text
            Case "2" : TextBox10.Text = "01" & "000000" & TextBox10.Text
            Case "3" : TextBox10.Text = "01" & "00000" & TextBox10.Text
            Case "4" : TextBox10.Text = "01" & "0000" & TextBox10.Text
            Case "5" : TextBox10.Text = "01" & "000" & TextBox10.Text
            Case "6" : TextBox10.Text = "01" & "00" & TextBox10.Text
            Case "7" : TextBox10.Text = "01" & "0" & TextBox10.Text
            Case "8" : TextBox10.Text = "01" & TextBox10.Text

        End Select
        Dim sql1011 As String = "SELECT COUNT(ncom_3) as Cantidad FROM custom_vianny.DBO.QAG0300   Where ccia='" + Label10.Text + "' AND NCOM_3='" + Trim(TextBox10.Text) + "'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        If Rsr1991.Read() = True And Rsr1991(0) > 0 Then
            Rsr1991.Close()

            Dim sql10112 As String = "SELECT COUNT(ncom_8) FROM  custom_vianny.DBO.Qamc800   Where ccia_8='" + Label10.Text + "' AND NCOM_8='" + Trim(TextBox10.Text) + "'"
            Dim cmd10112 As New SqlCommand(sql10112, conx)
            rs1 = cmd10112.ExecuteReader()
            If rs1.Read() = True And rs1(0) = 0 Then
                'MsgBox("paso")
                rs1.Close()

                habilitarGuardar()
                Dim sql101123 As String = "SELECT program_3,fcom_3,frequerida_3 ,estiloemp_3,estilo_3,descri_3, cantp_3,c.nomb_10  FROM custom_vianny.DBO.QAG0300 q left join custom_vianny.dbo.cag1000 c on '01'=c.ccia and q.fich_3 = c.fich_10  Where q.ccia='" + Label10.Text + "' AND NCOM_3='" + Trim(TextBox9.Text) + "'"
                Dim cmd101123 As New SqlCommand(sql101123, conx)
                rs2 = cmd101123.ExecuteReader()
                rs2.Read()
                TextBox1.Text = Trim(rs2(7))
                TextBox2.Text = Trim(rs2(0))
                TextBox3.Text = Trim(rs2(1))
                TextBox4.Text = Trim(rs2(2))
                TextBox8.Text = Trim(rs2(3))
                TextBox5.Text = Trim(rs2(4))
                TextBox6.Text = Trim(rs2(5))
                TextBox7.Text = Trim(rs2(6))
                rs2.Close()
            Else
                rs1.Close()
                mostrarCabeceraclonar()
                mostrarDetalleclonar()


            End If
            Button13.Enabled = True
            Button12.Enabled = True
            Button3.Enabled = True
        Else
            Rsr1991.Close()
            MsgBox("LA OP INGRESADA NO EXISTE")
        End If
    End Sub
    Private Function CantidadRegistros() As Integer
        Dim sql1 As String = "select count(nomb_8) from custom_vianny.dbo.qamc800 where ncom_8 ='" + TextBox9.Text.ToString.Trim & TextBox11.Text.ToString.Trim + "' and ccia_8 ='" + Label10.Text + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        rs5 = cmd1.ExecuteReader()
        Dim suma As Integer
        If rs5.Read() Then
            suma = Trim(rs5(0))

        Else
            suma = 0
        End If
        rs5.Close()
        Return suma
    End Function
    Private Function CantidadRegistrosCerrados() As Integer
        Dim sql1 As String = "select count(cierre_8) from custom_vianny.dbo.qamc800 where ncom_8 ='" + Trim(TextBox9.Text) + "' and ccia_8 ='" + Label10.Text + "' and cierre_8 ='1'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        rs6 = cmd1.ExecuteReader()
        Dim suma As Integer
        If rs6.Read() Then
            suma = Trim(rs6(0))
        Else
            suma = 0
        End If
        rs6.Close()
        Return suma
    End Function
    Private Sub mostrarCabeceraclonar()
        Dim sql1 As String = "EXEC custom_vianny.DBO.CaeSoft_GetallCabeceraMatrizConsumoAvios '" + Label10.Text + "','" + Trim(TextBox10.Text) + "',NULL,NULL"
        Dim cmd1 As New SqlCommand(sql1, conx)
        rs3 = cmd1.ExecuteReader()
        If rs3.Read() Then
            TextBox1.Text = Trim(rs3(2))
            TextBox2.Text = Trim(rs3(10))
            TextBox3.Text = Trim(rs3(3))
            TextBox4.Text = Trim(rs3(4))
            TextBox8.Text = Trim(rs3(7))
            TextBox5.Text = Trim(rs3(8))
            TextBox6.Text = Trim(rs3(6))
            TextBox7.Text = Trim(rs3(5))
        End If
        rs3.Close()
    End Sub
    Dim da55 As New DataTable
    Private Sub mostrarDetalleclonar()
        Dim sql1 As String = "Sp_ListarDetalleConsumo '" + Label10.Text + "','" + TextBox10.Text + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        rs4 = cmd1.ExecuteReader()
        Dim i As Integer
        While rs4.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i).Cells(0).Value = Trim(rs4(1))
            DataGridView1.Rows(i).Cells(6).Value = Trim(rs4(3))
            DataGridView1.Rows(i).Cells(7).Value = Trim(rs4(15))
            DataGridView1.Rows(i).Cells(8).Value = Trim(rs4(4))
            DataGridView1.Rows(i).Cells(2).Value = Trim(rs4(7))
            DataGridView1.Rows(i).Cells(12).Value = Trim(rs4(6))
            DataGridView1.Rows(i).Cells(9).Value = Trim(rs4(5))
            DataGridView1.Rows(i).Cells(1).Value = Trim(rs4(17))
            DataGridView1.Rows(i).Cells(13).Value = Trim(rs4(10))
            DataGridView1.Rows(i).Cells(11).Value = Trim(rs4(8))
            If Trim(rs4(9)) = True Then
                DataGridView1.Rows(i).Cells(3).Value = True
            End If
            i = i + 1
        End While
        rs4.Close()
    End Sub
    Private Sub mostrarCabecera()
        Dim sql1 As String = "EXEC custom_vianny.DBO.CaeSoft_GetallCabeceraMatrizConsumoAvios '" + Label10.Text + "','" + TextBox9.Text.ToString.Trim & TextBox11.Text.ToString.Trim + "',NULL,NULL"
        Dim cmd1 As New SqlCommand(sql1, conx)
        rs3 = cmd1.ExecuteReader()
        If rs3.Read() Then
            TextBox1.Text = Trim(rs3(2))
            TextBox2.Text = Trim(rs3(10))
            TextBox3.Text = Trim(rs3(3))
            TextBox4.Text = Trim(rs3(4))
            TextBox8.Text = Trim(rs3(7))
            TextBox5.Text = Trim(rs3(8))
            TextBox6.Text = Trim(rs3(6))
            TextBox7.Text = Trim(rs3(5))
        End If
        rs3.Close()
    End Sub

    Private Sub mostrarDetalle()
        Dim sql1 As String = "Sp_ListarDetalleConsumo '" + Label10.Text + "','" + TextBox9.Text.ToString.Trim & TextBox11.Text.ToString.Trim + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        rs4 = cmd1.ExecuteReader()
        Dim i As Integer
        While rs4.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i).Cells(0).Value = Trim(rs4(1))
            DataGridView1.Rows(i).Cells(6).Value = Trim(rs4(3))
            DataGridView1.Rows(i).Cells(7).Value = Trim(rs4(15))
            DataGridView1.Rows(i).Cells(8).Value = Trim(rs4(4))
            DataGridView1.Rows(i).Cells(2).Value = Trim(rs4(7))
            DataGridView1.Rows(i).Cells(12).Value = Trim(rs4(6))
            DataGridView1.Rows(i).Cells(9).Value = Trim(rs4(5))
            DataGridView1.Rows(i).Cells(1).Value = Trim(rs4(17))
            DataGridView1.Rows(i).Cells(13).Value = Trim(rs4(10))
            DataGridView1.Rows(i).Cells(11).Value = Trim(rs4(8))
            If Trim(rs4(9)) = True Then
                DataGridView1.Rows(i).Cells(3).Value = True
            End If
            i = i + 1
        End While
        rs4.Close()
    End Sub
    Sub habilitarGuardar()
        DataGridView1.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
        Button9.Enabled = True
        Button13.Enabled = True
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
    Sub InsertarLog(accion)
        Dim hj2 As New flog
        Dim hj1 As New vlog
        hj1.gmodulo = "MATRIZ-AVIOS"
        hj1.gcuser = MDIParent2.Label6.Text
        hj1.gaccion = accion
        hj1.gpc = My.Computer.Name
        hj1.gfecha = DateTimePicker1.Value
        hj1.gdetalle = Trim(TextBox9.Text)
        hj1.gccia = Label10.Text
        hj2.insertar_log(hj1)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        DataGridView1.Enabled = True
        Button9.Enabled = True
        Label14.Text = 2
        Button1.Enabled = True
        Button2.Enabled = True
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        MsgBox("ESTA FUNCION ESTA HABILITADA SOLO PARA EL ADMINISTRADOR DEL SISTEMA")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox9.Enabled = True
        Label13.Text = "ACTIVO"
        Label13.ForeColor = Color.Indigo
        limpiar()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        ReporteMatriz.TextBox1.Text = Label10.Text
        ReporteMatriz.TextBox2.Text = TextBox9.Text.Trim
        ReporteMatriz.ShowDialog()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        log.Label1.Text = Trim(TextBox9.Text)
        log.Label2.Text = "MATRIZ-AVIOS"
        log.Label3.Text = Label10.Text
        log.Show()
    End Sub
    Sub limpiar()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        DataGridView1.Rows.Clear()
    End Sub
    Sub limpiarInformacion()
        DataGridView1.Rows.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        Label13.Text = "ACTIVO"
        Label13.ForeColor = Color.Indigo
    End Sub
    Private Function validar_Articulo() As Boolean
        Dim dat1, suma As Integer
        dat1 = DataGridView1.Rows.Count
        For i = 0 To dat1 - 1
            If DataGridView1.Rows(i).Cells(6).Value Is Nothing Then
                suma = suma + 1
            End If
        Next
        If suma > 0 Then
            Return False
        End If
        Return True
    End Function
    Private Function validar_Factor() As Boolean
        Dim dat1, suma As Integer
        dat1 = DataGridView1.Rows.Count
        For i = 0 To dat1 - 1
            If DataGridView1.Rows(i).Cells(8).Value Is Nothing Then
                suma = suma + 1
            End If
        Next
        If suma > 0 Then
            Return False
        End If
        Return True
    End Function
    Private Function validar_um() As Boolean
        Dim dat1, suma As Integer
        dat1 = DataGridView1.Rows.Count
        For i = 0 To dat1 - 1
            If DataGridView1.Rows(i).Cells(9).Value Is Nothing Then
                suma = suma + 1
            End If
        Next
        If suma > 0 Then
            Return False
        End If
        Return True
    End Function
    Private Function validar_Generico() As Boolean
        Dim dat1, suma As Integer
        dat1 = DataGridView1.Rows.Count
        For i = 0 To dat1 - 1
            If DataGridView1.Rows(i).Cells(2).Value Is Nothing Then
                suma = suma + 1
            End If
        Next
        If suma > 0 Then
            Return False
        End If
        Return True
    End Function

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
        ' Ponemos la celda en modo de edición
        DataGridView1.BeginEdit(True)
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "A" Then
            Dim fechaActual As Date = Today

            pproductos.CCIA.Text = Label10.Text
            pproductos.ALMACEN.Text = "13"
            pproductos.ANO.Text = Label11.Text
            pproductos.FECHA.Text = Replace(fechaActual.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "13"
            pproductos.Label3.Text = 14
            pproductos.Label5.Text = e.RowIndex
            pproductos.ShowDialog()
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "B" Then
            If DataGridView1.Rows(e.RowIndex).Cells(3).Value = True Then

                DetalleMatrizAvios.Label1.Text = Label10.Text
                If CheckBox1.Checked = True Then
                    DetalleMatrizAvios.Label2.Text = TextBox10.Text
                Else
                    DetalleMatrizAvios.Label2.Text = TextBox9.Text
                End If

                DetalleMatrizAvios.Label3.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()

                DetalleMatrizAvios.ShowDialog()
            End If
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Obs" Then
            ObsMatriz.TextBox1.Text = ""
            If DataGridView1.Rows(e.RowIndex).Cells(11).Value Is Nothing Then
                ObsMatriz.TextBox1.Text = ""
            Else
                ObsMatriz.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(11).Value.ToString().Trim()
            End If

            ObsMatriz.Label1.Text = e.RowIndex
            ObsMatriz.Label2.Text = 3
            ObsMatriz.ShowDialog()
        End If
    End Sub
    Private WithEvents editingControl As Control
    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        If editingControl IsNot Nothing Then
            RemoveHandler editingControl.KeyDown, AddressOf EditingControl_KeyDown
        End If

        ' Asignar el nuevo control de edición y agregar el evento KeyDown
        editingControl = e.Control
        AddHandler editingControl.KeyDown, AddressOf EditingControl_KeyDown
    End Sub
    Private Sub EditingControl_KeyDown(sender As Object, e As KeyEventArgs)
        ' Verificar si la tecla presionada es F1
        If e.KeyCode = Keys.F1 Then
            ' Verificar si la celda actual está en la columna "Factor de Consumo"
            If DataGridView1.CurrentCell.ColumnIndex = 2 Then
                ' Cancelar la edición para evitar conflictos
                DataGridView1.EndEdit()
                ' Abrir el formulario deseado
                Form11.Label2.Text = DataGridView1.CurrentCell.RowIndex.ToString
                Form11.Label3.Text = "2"
                Form11.ShowDialog()
            End If
            If DataGridView1.CurrentCell.ColumnIndex = 1 Then
                ' Cancelar la edición para evitar conflictos
                DataGridView1.EndEdit()
                ' Abrir el formulario deseado
                EtapaMatriz.Label2.Text = DataGridView1.CurrentCell.RowIndex.ToString
                EtapaMatriz.Label3.Text = "2"
                EtapaMatriz.ShowDialog()
            End If
        End If
    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.F1 Then
            ListaOD.Label4.Text = 4
            ListaOD.ShowDialog()
        End If
        If e.KeyCode = Keys.Enter Then
            clonar()
            CheckBox1.Enabled = False
            TextBox10.Enabled = False
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If validar_Etapa() = True Then
            If validar_Generico() = True Then
                If validar_Articulo() = True Then
                    If validar_Factor() = True Then
                        If validar_um() = True Then
                            Dim agregar As String = "delete from custom_vianny.dbo.qamc800  where ncom_8 ='" + Trim(TextBox9.Text) & Trim(TextBox11.Text) + "' and ccia_8 ='" + Trim(Label10.Text) + "' "
                            ELIMINAR(agregar)
                            For i = 0 To DataGridView1.Rows.Count - 1
                                Dim cmd16 As New SqlCommand("insert into custom_vianny.dbo.qamc800(ccia_8,item_8,ncom_8,linea_8,factor_8,udm_8,gene_8,nomb_8,obser_8,talcol_8,area_8,cierre_8,cliente_8) 
                                                             values (@ccia_8,@item_8,@ncom_8,@linea_8,@factor_8,@udm_8,@gene_8,@nomb_8,@obser_8,@talcol_8,@area_8,@cierre_8,@cliente_8)", conx)
                                cmd16.Parameters.AddWithValue("@ccia_8", Trim(Label10.Text))
                                cmd16.Parameters.AddWithValue("@item_8", DataGridView1.Rows(i).Cells(0).Value.ToString().Trim())
                                cmd16.Parameters.AddWithValue("@ncom_8", Trim(TextBox9.Text) & Trim(TextBox11.Text))
                                cmd16.Parameters.AddWithValue("@linea_8", DataGridView1.Rows(i).Cells(6).Value.ToString().Trim())
                                cmd16.Parameters.AddWithValue("@factor_8", Convert.ToDouble(DataGridView1.Rows(i).Cells(8).Value))
                                cmd16.Parameters.AddWithValue("@udm_8", DataGridView1.Rows(i).Cells(9).Value.ToString().Trim())
                                cmd16.Parameters.AddWithValue("@gene_8", DataGridView1.Rows(i).Cells(12).Value.ToString().Trim())
                                cmd16.Parameters.AddWithValue("@nomb_8", DataGridView1.Rows(i).Cells(2).Value.ToString().Trim())
                                cmd16.Parameters.AddWithValue("@obser_8", DataGridView1.Rows(i).Cells(11).Value.ToString().Trim())
                                If DataGridView1.Rows(i).Cells(3).Value = True Then
                                    cmd16.Parameters.AddWithValue("@talcol_8", "1")
                                Else
                                    cmd16.Parameters.AddWithValue("@talcol_8", "0")
                                End If
                                cmd16.Parameters.AddWithValue("@area_8", DataGridView1.Rows(i).Cells(13).Value.ToString().Trim())
                                cmd16.Parameters.AddWithValue("@cierre_8", "0")
                                cmd16.Parameters.AddWithValue("@cliente_8", "")
                                cmd16.ExecuteNonQuery()
                                If Label14.Text = 1 Then
                                    InsertarLog(1)
                                Else
                                    InsertarLog(2)
                                End If
                            Next
                            limpiarInformacion()
                            TextBox10.Enabled = False
                            CheckBox1.Checked = False
                        Else
                            MsgBox("FALTA AGREGAR INFORMACION EN UNA FILA DEL CAMPO UM DE LA MATRIZ")
                        End If
                    Else
                        MsgBox("FALTA AGREGAR INFORMACION EN UNA FILA DEL CAMPO FACTOR DE LA MATRIZ")
                    End If
                Else
                    MsgBox("FALTA AGREGAR INFORMACION EN UNA FILA DEL CAMPO ARTICULO DE LA MATRIZ")
                End If
            Else
                MsgBox("FALTA AGREGAR INFORMACION EN UNA FILA DEL CAMPO GENERICO DE LA MATRIZ")
            End If
        Else
            MsgBox("FALTA AGREGAR INFORMACION EN UNA FILA DEL CAMPO ETAPA DE LA MATRIZ")
        End If
    End Sub
    Private Function validar_Etapa() As Boolean
        Dim dat1, suma As Integer
        dat1 = DataGridView1.Rows.Count
        For i = 0 To dat1 - 1
            If DataGridView1.Rows(i).Cells(1).Value Is Nothing Then
                suma = suma + 1
            End If
        Next
        If suma > 0 Then
            Return False
        End If
        Return True
    End Function
    Private Sub MatrizAviosOd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bloquear()
        TextBox9.Select()
    End Sub
    Sub bloquear()
        DataGridView1.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = False
        Button11.Enabled = False
        Button12.Enabled = False
        Button13.Enabled = False
    End Sub
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim od As String
            od = Trim(TextBox9.Text) & Trim(TextBox11.Text)
            CheckBox1.Enabled = True
            Label14.Text = 1

            Dim sql1011 As String = "SELECT COUNT(ncom_3) as Cantidad FROM custom_vianny.DBO.QAG0300   Where ccia='" + Label10.Text + "' AND NCOM_3='" + od + "'"
            Dim cmd1011 As New SqlCommand(sql1011, conx)
            Rsr1991 = cmd1011.ExecuteReader()
            If Rsr1991.Read() = True And Rsr1991(0) > 0 Then
                Rsr1991.Close()

                Dim sql10112 As String = "SELECT COUNT(ncom_8) FROM  custom_vianny.DBO.Qamc800   Where ccia_8='" + Label10.Text + "' AND NCOM_8='" + od + "'"
                Dim cmd10112 As New SqlCommand(sql10112, conx)
                rs1 = cmd10112.ExecuteReader()
                If rs1.Read() = True And rs1(0) = 0 Then

                    rs1.Close()
                    Dim agregar As String = "delete from custom_vianny.dbo.qamc800ct where ncom_8c ='" + od + "' and ccia_8c ='" + Trim(Label10.Text) + "'"
                    ELIMINAR(agregar)
                    habilitarGuardar()
                    Dim sql101123 As String = "SELECT program_3,fcom_3,frequerida_3 ,estiloemp_3,estilo_3,descri_3, cantp_3,c.nomb_10  FROM custom_vianny.DBO.QAG0300 q left join custom_vianny.dbo.cag1000 c on '01'=c.ccia and q.fich_3 = c.fich_10  Where q.ccia='" + Label10.Text + "' AND NCOM_3='" + od + "'"
                    Dim cmd101123 As New SqlCommand(sql101123, conx)
                    rs2 = cmd101123.ExecuteReader()
                    rs2.Read()
                    TextBox1.Text = Trim(rs2(7))
                    TextBox2.Text = Trim(rs2(0))
                    TextBox3.Text = Trim(rs2(1))
                    TextBox4.Text = Trim(rs2(2))
                    TextBox8.Text = Trim(rs2(3))
                    TextBox5.Text = Trim(rs2(4))
                    TextBox6.Text = Trim(rs2(5))
                    TextBox7.Text = Trim(rs2(6))
                    rs2.Close()
                Else

                    rs1.Close()
                    mostrarCabecera()
                    mostrarDetalle()
                    If CantidadRegistros() = CantidadRegistrosCerrados() Then
                        Button10.Enabled = False
                        Label13.Text = "CERRADO"
                        Label13.ForeColor = Color.Red
                    Else

                        Button10.Enabled = True
                    End If

                End If
                Button13.Enabled = True
                Button12.Enabled = True
                Button3.Enabled = True
            Else
                Rsr1991.Close()
                MsgBox("LA OP INGRESADA NO EXISTE")
            End If
        End If
        If e.KeyCode = Keys.F1 Then
            ListaOD.Label4.Text = 3
            ListaOD.ShowDialog()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim numero As Integer
        DataGridView1.Rows.Add(1)
        numero = DataGridView1.Rows.Count
        DataGridView1.Rows(numero - 1).Cells(0).Value = numero
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        End If

    End Sub
End Class