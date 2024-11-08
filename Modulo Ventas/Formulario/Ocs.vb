Imports System.Data.SqlClient
Imports System.Net.Mail
Public Class Ocs
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim Rsr20, Rsr21, Rsr22 As SqlDataReader
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox1.Text.Length
                Case "1" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = TextBox1.Text
            End Select
            Dim Rsr1991 As SqlDataReader
            Dim cco As String
            Dim sql1011 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAlluLTIMONumeroOrdendeCompra '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "'"
            Dim cmd1011 As New SqlCommand(sql1011, conx)
            Rsr1991 = cmd1011.ExecuteReader()
            Rsr1991.Read()
            TextBox2.Text = Rsr1991(0)
            Rsr1991.Close()
            TextBox2.Select()
        End If

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "3555"
            Clientes.Show()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then
            Select Case TextBox2.Text.Length
                Case "1" : TextBox2.Text = "00000" & "" & TextBox2.Text
                Case "2" : TextBox2.Text = "0000" & "" & TextBox2.Text
                Case "3" : TextBox2.Text = "000" & "" & TextBox2.Text
                Case "4" : TextBox2.Text = "00" & "" & TextBox2.Text
                Case "5" : TextBox2.Text = "0" & "" & TextBox2.Text
                Case "6" : TextBox2.Text = TextBox2.Text

            End Select
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Select()


            abrir()


            TextBox3.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            ComboBox1.Enabled = True
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True

            Button2.Enabled = True
            Button3.Enabled = True

            Button6.Enabled = True
            'DateTimePicker1.Enabled = True
            DataGridView1.Enabled = True
            TextBox3.Select()
            Button1.Enabled = True
            Dim sql10215 As String = "select COUNT(ncom_3) from custom_vianny.dbo.lap0300 p inner join custom_vianny.dbo.LAG0300 g on p.ccia_3a =g.ccia_3 and p.ncom_3a = g.ncom_3  where p.ncom_3a ='" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "' and p.ccia_3a ='" + Trim(Label16.Text) + "'"
            Dim cmd10215 As New SqlCommand(sql10215, conx)
            Rsr22 = cmd10215.ExecuteReader()
            Rsr22.Read()

            If Rsr22(0) = "0" Then
                'Label20.Text = 0
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""

                TextBox10.Text = ""
                TextBox11.Text = ""
                TextBox12.Text = ""
                DataGridView1.Rows.Clear()
                DataGridView1.ReadOnly = False
                Rsr22.Close()
            Else

                Button5.Enabled = True
                DataGridView1.Rows.Clear()
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                TextBox6.Enabled = False
                TextBox7.Enabled = False
                TextBox8.Enabled = False
                TextBox10.Enabled = False
                TextBox11.Enabled = False
                TextBox12.Enabled = False
                TextBox13.Enabled = False
                DataGridView1.ReadOnly = True
                Button4.Enabled = True
                ComboBox1.Enabled = False
                'TextBox14.Enabled = False
                'Label20.Text = 1

                Rsr22.Close()

                Dim sql1021 As String = "EXEC CaeSoft_GetAllOrdenesdeCompraCabecera '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "','" + Trim(TextBox2.Text) + "' "
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr21 = cmd1021.ExecuteReader()
                Dim i51 As Integer
                i51 = 0
                While Rsr21.Read()

                    TextBox3.Text = Rsr21(2)
                    TextBox4.Text = Rsr21(29)
                    TextBox5.Text = Rsr21(5)
                    TextBox6.Text = Rsr21(8)
                    TextBox7.Text = Trim(Rsr21(9))
                    TextBox8.Text = Trim(Rsr21(12))
                    'TextBox14.Text = Rsr21(27)
                    'TextBox15.Text = Rsr21(28)
                    TextBox9.Text = Rsr21(6)
                    'If Trim(Rsr21(14)) = "2" Then
                    '    Label21.Visible = True
                    'Else
                    '    Label21.Visible = False
                    'End If
                    If Rsr21(33) = "1" Then
                        CheckBox1.Checked = True
                    Else
                        CheckBox1.Checked = False
                    End If
                    If Rsr21(16) = "1" Then
                        CheckBox2.Checked = True
                    Else
                        CheckBox2.Checked = False
                    End If
                    If Rsr21(17) = "1" Then
                        CheckBox3.Checked = True
                    Else
                        CheckBox3.Checked = False
                    End If
                    If Rsr21(18) = "01" Then

                        TextBox10.Text = Rsr21(19)
                        TextBox11.Text = Rsr21(20)
                        TextBox12.Text = Rsr21(21)
                    Else
                        If Rsr21(17) = "02" Then
                            TextBox10.Text = Rsr21(22)
                            TextBox11.Text = Rsr21(23)
                            TextBox12.Text = Rsr21(24)
                        End If

                    End If

                    DateTimePicker1.Value = Rsr21(3)
                    DateTimePicker2.Value = Rsr21(4)
                    If Rsr21(15) = 1 Then
                        RadioButton1.Checked = True
                    Else
                        RadioButton2.Checked = True
                    End If


                    If Rsr21(18) = "01" Then
                        ComboBox2.SelectedIndex = 0
                    Else
                        ComboBox2.SelectedIndex = 1
                    End If

                    abrir()
                    enunciado2 = New SqlCommand("select nomb_15k as nomb_15k from custom_vianny.dbo.kag1500 where  ccia_15k ='" + Trim(Label16.Text) + "' AND cond_15k='" + Trim(Rsr21(7)) + "'", conx)
                    respuesta2 = enunciado2.ExecuteReader
                    While respuesta2.Read

                        ComboBox1.Text = respuesta2.Item("nomb_15k")
                    End While
                    respuesta2.Close()




                End While
                Rsr21.Close()

                Dim sql10212 As String = "EXEC CaeSoft_GetAllOrdenesdeCompraDetalle '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "','" + Trim(TextBox2.Text) + "'"
                Dim cmd10212 As New SqlCommand(sql10212, conx)
                Rsr20 = cmd10212.ExecuteReader()


                While Rsr20.Read()
                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i51).Cells(0).Value = Rsr20(3)
                    'Dim i4 As Integer
                    'i4 = Len(Rsr20(3))

                    'Select Case i4
                    '    Case "1" : DataGridView1.Rows(i51).Cells(0).Value = "00" & "" & Convert.ToString(Rsr20(3))
                    '    Case "2" : DataGridView1.Rows(i51).Cells(0).Value = "0" & "" & Convert.ToString(Rsr20(3))
                    '    Case "3" : DataGridView1.Rows(i51).Cells(0).Value = Convert.ToString(Rsr20(3))
                    'End Select
                    DataGridView1.Rows(i51).Cells(0).Value = Rsr20(20)

                    DataGridView1.Rows(i51).Cells(2).Value = Rsr20(2)
                    DataGridView1.Rows(i51).Cells(3).Value = Rsr20(27)
                    DataGridView1.Rows(i51).Cells(5).Value = Rsr20(3)
                    DataGridView1.Rows(i51).Cells(6).Value = Rsr20(4)
                    DataGridView1.Rows(i51).Cells(7).Value = Rsr20(18)
                    DataGridView1.Rows(i51).Cells(8).Value = Rsr20(6)
                    DataGridView1.Rows(i51).Cells(9).Value = Rsr20(7)
                    DataGridView1.Rows(i51).Cells(10).Value = Rsr20(8)
                    DataGridView1.Rows(i51).Cells(25).Value = Rsr20(21)

                    DataGridView1.Rows(i51).Cells(13).Value = Rsr20(22)
                    DataGridView1.Rows(i51).Cells(14).Value = Rsr20(23)


                    If ComboBox2.Text = "SOLES" Then
                        DataGridView1.Rows(i51).Cells(11).Value = Rsr20(9)
                        DataGridView1.Rows(i51).Cells(12).Value = Rsr20(10)
                    Else
                        DataGridView1.Rows(i51).Cells(11).Value = Rsr20(11)
                        DataGridView1.Rows(i51).Cells(12).Value = Rsr20(12)
                    End If
                    Dim cant10, sum10, cant9, sum9, cant8, sum8, cant101, sum101, cant91, sum91, cant81, sum81, cant102, sum102, cant92, sum92, cant82, sum82 As Double
                    If CheckBox2.Checked = True And CheckBox3.Checked = False Then
                        Dim i8 As Integer
                        i8 = DataGridView1.RowCount
                        CheckBox2.Enabled = True
                        CheckBox3.Checked = False
                        For i = 0 To i8 - 1
                            DataGridView1.Rows(i).Cells(15).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(11).Value

                            DataGridView1.Rows(i).Cells(17).Value = DataGridView1.Rows(i).Cells(15).Value * 1.18
                            DataGridView1.Rows(i).Cells(16).Value = DataGridView1.Rows(i).Cells(17).Value - DataGridView1.Rows(i).Cells(15).Value

                            Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"


                        Next
                        'For a9 = 0 To i8 - 1
                        cant10 = Val(DataGridView1.Rows(i51).Cells(15).Value)
                        sum10 = cant10 + Val(sum10)
                        cant9 = Val(DataGridView1.Rows(i51).Cells(16).Value)
                        sum9 = cant9 + Val(sum9)
                        cant8 = Val(DataGridView1.Rows(i51).Cells(17).Value)
                        sum8 = cant8 + Val(sum8)

                        'Next
                        TextBox10.Text = sum10.ToString("N2")
                        TextBox11.Text = sum9.ToString("N2")

                        TextBox12.Text = sum8.ToString("N2")
                    Else
                        If CheckBox2.Checked = False Then
                            Dim i9 As Integer
                            i9 = DataGridView1.RowCount
                            For i = 0 To i9 - 1
                                DataGridView1.Rows(i).Cells(17).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(11).Value
                                DataGridView1.Rows(i).Cells(15).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(11).Value
                                DataGridView1.Rows(i).Cells(16).Value = "0.00"
                                Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                                Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                                Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                            Next
                            'For a91 = 0 To i9 - 1
                            cant101 = Val(DataGridView1.Rows(i51).Cells(15).Value)
                            sum101 = cant101 + Val(sum101)
                            cant91 = Val(DataGridView1.Rows(i51).Cells(16).Value)
                            sum91 = cant91 + Val(sum91)
                            cant81 = Val(DataGridView1.Rows(i51).Cells(17).Value)
                            sum81 = cant81 + Val(sum81)

                            'Next
                            TextBox10.Text = sum101.ToString("N2")
                            TextBox11.Text = sum91.ToString("N2")

                            TextBox12.Text = sum81.ToString("N2")
                        Else
                            If CheckBox2.Checked = True And CheckBox3.Checked = True Then
                                Dim i10 As Integer
                                i10 = DataGridView1.RowCount

                                For i = 0 To i10 - 1
                                    DataGridView1.Rows(i).Cells(17).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(11).Value

                                    DataGridView1.Rows(i).Cells(15).Value = DataGridView1.Rows(i).Cells(17).Value / 1.18
                                    DataGridView1.Rows(i).Cells(16).Value = DataGridView1.Rows(i).Cells(17).Value - DataGridView1.Rows(i).Cells(15).Value

                                    Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                                    Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                                    Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                                Next
                                'For a92 = 0 To i10 - 1
                                'MsgBox(Val(DataGridView1.Rows(i51).Cells(12).Value))
                                cant102 = Val(DataGridView1.Rows(i51).Cells(15).Value)
                                sum102 = cant102 + Val(sum102)
                                cant92 = Val(DataGridView1.Rows(i51).Cells(16).Value)
                                sum92 = cant92 + Val(sum92)
                                cant82 = Val(DataGridView1.Rows(i51).Cells(17).Value)
                                sum82 = cant82 + Val(sum82)

                                'Next
                                TextBox10.Text = sum102.ToString("N2")
                                TextBox11.Text = sum92.ToString("N2")

                                TextBox12.Text = sum82.ToString("N2")
                            End If
                        End If


                    End If
                    i51 = i51 + 1
                End While

                Rsr20.Close()
            End If
        Else
            If e.KeyCode = Keys.F1 Then
                Det_oc.Label4.Text = Trim(Label16.Text)
                Det_oc.Label5.Text = TextBox1.Text
                Det_oc.Label6.Text = "1"
                Det_oc.Show()
            End If
        End If
    End Sub
    Dim ty2, ty1 As New DataTable
    Sub llenar_combo_box1()
        Try

            conn = New SqlDataAdapter("SELECT cond_15k AS CON , nomb_15k AS NOMBRE FROM  custom_vianny.dbo.kag1500 where  ccia_15k ='" + Trim(Label16.Text) + "'", conx)
            conn.Fill(ty1)
            ComboBox1.DataSource = ty1
            ComboBox1.DisplayMember = "NOMBRE"
            ComboBox1.ValueMember = "CON"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Ocs_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim func As New ftcambio
        Dim dts As New vtcambio
        dts.gfecha = DateTimePicker1.Text
        TextBox9.Text = func.mostrar_tipo_cambio(dts)
        ty1.Clear()
        abrir()
        llenar_combo_box1()
        CheckBox2.Checked = True
        ComboBox2.SelectedIndex = 0
        RadioButton1.Checked = True
        TextBox3.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        ComboBox1.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        CheckBox2.Enabled = False
        CheckBox3.Enabled = False
        CheckBox4.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False


        TextBox1.Select()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox3.Text.ToString.Trim().Length = 0 Then
            MsgBox("Falta Agregar El Cliente")
        Else
            DataGridView1.Rows.Add(1)
            Dim f, r As Integer
            f = DataGridView1.Rows.Count
            If f = 1 Then
                r = 1
                For i = 0 To f - 1
                    DataGridView1.Rows(i).Cells(0).Value = 1
                    DataGridView1.Rows(i).Cells(13).Value = 1
                    DataGridView1.Rows(i).Cells(11).Value = 0
                Next
            Else
                r = 0
                For i = 1 To f - 1
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView1.Rows(i - 1).Cells(0).Value + 1 + r
                    DataGridView1.Rows(i).Cells(13).Value = 1
                    DataGridView1.Rows(i).Cells(11).Value = 0
                    If DataGridView1.Rows(i - 1).Cells(2).Value IsNot Nothing AndAlso DataGridView1.Rows(i - 1).Cells(2).Value.ToString().Trim().Length > 0 Then
                        DataGridView1.Rows(i).Cells(2).Value = DataGridView1.Rows(i - 1).Cells(2).Value.ToString().Trim()
                    End If
                Next
            End If
        End If


    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 1 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Form6.Label3.Text = num1
            Form6.Label4.Text = Trim(Label16.Text)
            Form6.Show(Me)

        End If
        'If e.ColumnIndex = 3 Then

        '    ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
        '    Dim num1, num2 As Integer
        '    num1 = e.RowIndex.ToString
        '    num2 = e.ColumnIndex.ToString
        '    Lista_Productos.Label2.Text = Label16.Text
        '    Lista_Productos.Label3.Text = TextBox1.Text

        '    If Len(Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)) = 0 Then
        '        Lista_Productos.Label4.Text = ""
        '    Else

        '        Lista_Productos.Label4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
        '    End If



        '    Lista_Productos.Label5.Text = Label17.Text
        '    Lista_Productos.Label7.Text = e.RowIndex
        '    Lista_Productos.ShowDialog()

        'End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "C" Then
            pproductos.CCIA.Text = Label16.Text
            pproductos.ALMACEN.Text = Trim(TextBox1.Text)
            pproductos.ANO.Text = Label11.Text
            pproductos.FECHA.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = Trim(TextBox1.Text)
            pproductos.Label3.Text = 4
            pproductos.Label5.Text = e.RowIndex
            pproductos.Show()
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Obs" Then
            OB_SC.Label1.Text = e.RowIndex
            OB_SC.Label2.Text = 4
            OB_SC.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(18).Value)
            OB_SC.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(19).Value)
            OB_SC.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(20).Value)
            OB_SC.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(21).Value)
            OB_SC.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(22).Value)
            OB_SC.ShowDialog()
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PRC" Then


            If Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value) = "" Then
                    MsgBox("EL CAMPO CODIGO DE PRODUCTO ESTA VACIO, INGRESE LA INFORMACION DEL CODIGO DEL PRODUCTO")
                Else
                    Lista_ultimos_precios.Label1.Text = Label16.Text
                Lista_ultimos_precios.Label2.Text = e.RowIndex
                Lista_ultimos_precios.Label3.Text = Trim(ComboBox2.Text)
                Lista_ultimos_precios.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
                Lista_ultimos_precios.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
                    Lista_ultimos_precios.ShowDialog()
                End If



        End If

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
    Private Function validar() As Boolean
        If TextBox3.Text.ToString().Trim().Length = 0 Then
            MessageBox.Show("El Campo Cliente está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If TextBox9.Text.ToString().Trim().Length = 0 Then
            MessageBox.Show("El Campo Tipo de Cambio está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If

        Return True
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If validar() = True Then
            Dim cmd15 As New SqlCommand("IF NOT EXISTS( SELECT NCOM_3 FROM custom_vianny.dbo.LAG0300 WHERE CCIA_3 =@ccia_3 AND NCOM_3 = @ncom_3 )
INSERT INTO custom_vianny.dbo.lag0300 ( ccia_3 , cperiodo_3 , ncom_3     , fich_3        , fcom_3     , fcome_3    ,cond_3 , rubr_3 ,tcam_3 , mone_3 , tot1_3 , tot2_3 , vvta1_3 , vvta2_3 , igv1_3 , igv2_3 ,cigv_3 , refe_3 , gloa_3 , glob_3 , gloc_3 , glod_3 , aigv_3 , talm_3 , tipoc_3 , ordenp_3 ,lugar_3 ,flag_3 ,atencion_3 , cerrado_3 , program_3 , exceso_3 , termina_3, vpercep1_3, vpercep2_3, totg1_3, totg2_3 , aprobado_3, UserApro_3, Tipo_3 )
	          VALUES(@ccia_3 , ''     ,@ncom_3     ,@fich_3        ,@fcom_3     ,@fcome_3    ,@cond_3, ''     ,@tcam_3, @mone_3,@tot1_3 ,@tot2_3 ,@vvta1_3 ,@vvta2_3 ,@igv1_3 ,@igv2_3 ,@cigv_3, @refe_3,@gloa_3 ,@glob_3 , @gloc_3, ''     ,@aigv_3 , ''     , ''      , ''       ,@lugar_3,1      ,@atencion_3, 0         , ''        , 0        , 0        ,0.00       , 0.00      ,@totg1_3,@totg2_31 ,0         ,@UserApro_3,  '01')		ELSE
		   UPDATE  custom_vianny.dbo.LAG0300 SET fich_3 =@fich_3,
		    fcom_3  = @fcom_3 ,
		    fcome_3 = @fcome_3 ,
			cond_3 =@cond_3,
			rubr_3 = '',
			tcam_3 = @tcam_3,
			 mone_3 =@mone_3,
			tot1_3 =@tot1_3, 
			tot2_3 =@tot2_3 , 
             vvta1_3 =@vvta1_3 , 
             vvta2_3 =@vvta2_3 ,
             igv1_3 = @igv1_3, 
             igv2_3 =@igv2_3 , 
			cigv_3 = @cigv_3 , 
			refe_3 =@refe_3 ,
			gloa_3 = @gloa_3 ,
			glob_3 = @glob_3 ,
			gloc_3 = @gloc_3 ,
			glod_3 = '' ,
			aigv_3 = @aigv_3, 
			talm_3 ='' ,
			tipoc_3 =  '' , 
			ordenp_3 = '',
			lugar_3 =@lugar_3,
			flag_3 =1 ,
			atencion_3 =@atencion_3 , 
			cerrado_3 =0 , 
			program_3 ='        ' ,
			exceso_3 = 0 , 
			termina_3= 0,
			vpercep1_3 = 0.00, 
			vpercep2_3 = 0.00, 
			totg1_3    = @totg1_3, 
			totg2_3    = @totg2_31,
			aprobado_3 = 0,
			UserApro_3 = @UserApro_3,
			Tipo_3     = '01'
		    WHERE CCIA_3 = @ccia_3 AND NCOM_3 = @ncom_3", conx)
            cmd15.Parameters.AddWithValue("@ccia_3", Trim(Label16.Text))
            cmd15.Parameters.AddWithValue("@ncom_3", Trim(TextBox1.Text) & Trim(TextBox2.Text))
            cmd15.Parameters.AddWithValue("@fich_3", Trim(TextBox3.Text))
            cmd15.Parameters.AddWithValue("@fcom_3", DateTimePicker1.Value.Date)
            cmd15.Parameters.AddWithValue("@fcome_3", DateTimePicker2.Value.Date)
            cmd15.Parameters.AddWithValue("@cond_3", ComboBox1.SelectedValue.ToString)
            cmd15.Parameters.AddWithValue("@tcam_3", Convert.ToDouble(TextBox9.Text))
            If ComboBox2.Text = "SOLES" Then
                cmd15.Parameters.AddWithValue("@mone_3", 1)
            Else
                cmd15.Parameters.AddWithValue("@mone_3", 2)
            End If
            If ComboBox2.Text = "SOLES" Then
                cmd15.Parameters.AddWithValue("@tot1_3", Convert.ToDouble(TextBox12.Text))
                cmd15.Parameters.AddWithValue("@tot2_3", Convert.ToDouble(TextBox12.Text) / Convert.ToDouble(TextBox9.Text))
                cmd15.Parameters.AddWithValue("@vvta1_3", Convert.ToDouble(TextBox10.Text))
                cmd15.Parameters.AddWithValue("@vvta2_3", Convert.ToDouble(TextBox10.Text) / Convert.ToDouble(TextBox9.Text))
                cmd15.Parameters.AddWithValue("@igv1_3", Convert.ToDouble(TextBox11.Text))
                cmd15.Parameters.AddWithValue("@igv2_3", Convert.ToDouble(TextBox11.Text) / Convert.ToDouble(TextBox9.Text))
            Else
                cmd15.Parameters.AddWithValue("@tot1_3", Convert.ToDouble(TextBox12.Text) * Convert.ToDouble(TextBox9.Text))
                cmd15.Parameters.AddWithValue("@tot2_3", Convert.ToDouble(TextBox12.Text))
                cmd15.Parameters.AddWithValue("@vvta1_3", Convert.ToDouble(TextBox10.Text) * Convert.ToDouble(TextBox9.Text))
                cmd15.Parameters.AddWithValue("@vvta2_3", Convert.ToDouble(TextBox10.Text))
                cmd15.Parameters.AddWithValue("@igv1_3", Convert.ToDouble(TextBox11.Text) * Convert.ToDouble(TextBox9.Text))
                cmd15.Parameters.AddWithValue("@igv2_3", Convert.ToDouble(TextBox11.Text))
            End If
            If CheckBox3.Checked = True Then
                cmd15.Parameters.AddWithValue("@cigv_3", 1)
            Else
                cmd15.Parameters.AddWithValue("@cigv_3", 0)
            End If
            cmd15.Parameters.AddWithValue("@refe_3", Trim(TextBox6.Text))
            cmd15.Parameters.AddWithValue("@gloa_3", Trim(TextBox7.Text))
            cmd15.Parameters.AddWithValue("@glob_3", Trim(TextBox14.Text))
            cmd15.Parameters.AddWithValue("@gloc_3", Trim(TextBox15.Text))
            If CheckBox2.Checked = True Then
                cmd15.Parameters.AddWithValue("@aigv_3", 1)
            Else
                cmd15.Parameters.AddWithValue("@aigv_3", 0)
            End If

            cmd15.Parameters.AddWithValue("@lugar_3", Trim(TextBox8.Text))
            cmd15.Parameters.AddWithValue("@atencion_3", Trim(TextBox5.Text))
            If ComboBox2.Text = "SOLES" Then
                cmd15.Parameters.AddWithValue("@totg1_3", Convert.ToDouble(TextBox12.Text))
                cmd15.Parameters.AddWithValue("@totg2_31", Convert.ToDouble(TextBox12.Text) / Convert.ToDouble(TextBox9.Text))

            Else

                cmd15.Parameters.AddWithValue("@totg1_3", Convert.ToDouble(TextBox12.Text) * Convert.ToDouble(TextBox9.Text))
                cmd15.Parameters.AddWithValue("@totg2_31", Convert.ToDouble(TextBox12.Text))
            End If

            cmd15.Parameters.AddWithValue("@UserApro_3", "")
            cmd15.ExecuteNonQuery()





            abrir()
            Dim agregar As String = "DELETE FROM  custom_vianny.dbo.lap0300 Where CCIA_3a =  '" + Trim(Label16.Text) + "' And ncom_3a= '" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + "'"
            ELIMINAR(agregar)
            Dim u As Integer
            u = DataGridView1.Rows.Count

            For i = 0 To u - 1

                Dim cmd16 As New SqlCommand("INSERT INTO  custom_vianny.dbo.lap0300 ( ccia_3a , cperiodo_3a , ncom_3a   , Item_3a,cant_3a , linea_3a, preun1_3a, tot1_3a   , preun2_3a ,tot2_3a   , nombl_3a , ccos_3a , obser1_3a , obser2_3a , obser3_3a , obser4_3a,obser5_3a , pedido_3a , flag_3a , cante_3a ,unide_3a ,cante2_3a,unide2_3a ,percep_3a,padic_3a ,cantadi_3a,frequer_3a,ccosto_3a)
                                                                         VALUES( @ccia_3a, ''          ,@ncom_3a   ,@Item_3a,@cant_3a,@linea_3a,@preun1_3a,@tot1_3a   ,@preun2_3a ,@tot2_3a  ,@op       , ''      ,@obser1_3a ,@obser2_3a ,@obser3_3a ,@obser4_3a,@obser5_3a,@pedido_3a , 1       , @cante_3a,@unide_3a, 0.00    ,@unide2_3a,@percep_3a,@padic_3a,@cantadi_3a,NULL      ,'')", conx)
                cmd16.Parameters.AddWithValue("@ccia_3a", Trim(Label16.Text))
                cmd16.Parameters.AddWithValue("@ncom_3a", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd16.Parameters.AddWithValue("@Item_3a", Trim(DataGridView1.Rows(i).Cells(0).Value))
                cmd16.Parameters.AddWithValue("@cant_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value))
                cmd16.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(5).Value))
                If ComboBox2.Text = "SOLES" Then
                    cmd16.Parameters.AddWithValue("@preun1_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(11).Value))
                    cmd16.Parameters.AddWithValue("@tot1_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(12).Value))
                    cmd16.Parameters.AddWithValue("@preun2_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(11).Value) / Convert.ToDouble(TextBox9.Text))
                    cmd16.Parameters.AddWithValue("@tot2_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(12).Value) / Convert.ToDouble(TextBox9.Text))
                Else
                    cmd16.Parameters.AddWithValue("@preun1_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(11).Value) * Convert.ToDouble(TextBox9.Text))
                    cmd16.Parameters.AddWithValue("@tot1_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(12).Value) * Convert.ToDouble(TextBox9.Text))
                    cmd16.Parameters.AddWithValue("@preun2_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(11).Value))
                    cmd16.Parameters.AddWithValue("@tot2_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(12).Value))
                End If
                cmd16.Parameters.AddWithValue("@op", Trim(DataGridView1.Rows(i).Cells(3).Value))
                cmd16.Parameters.AddWithValue("@obser1_3a", Trim(DataGridView1.Rows(i).Cells(18).Value))
                cmd16.Parameters.AddWithValue("@obser2_3a", Trim(DataGridView1.Rows(i).Cells(19).Value))
                cmd16.Parameters.AddWithValue("@obser3_3a", Trim(DataGridView1.Rows(i).Cells(20).Value))
                cmd16.Parameters.AddWithValue("@obser4_3a", Trim(DataGridView1.Rows(i).Cells(21).Value))
                cmd16.Parameters.AddWithValue("@obser5_3a", Trim(DataGridView1.Rows(i).Cells(22).Value))
                cmd16.Parameters.AddWithValue("@pedido_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd16.Parameters.AddWithValue("@flag_4a", 1)
                cmd16.Parameters.AddWithValue("@percep_3a", Convert.ToByte(DataGridView1.Rows(i).Cells(25).Value))
                cmd16.Parameters.AddWithValue("@cante_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
                cmd16.Parameters.AddWithValue("@unide_3a", DataGridView1.Rows(i).Cells(8).Value)
                cmd16.Parameters.AddWithValue("@unide2_3a", DataGridView1.Rows(i).Cells(10).Value)
                cmd16.Parameters.AddWithValue("@padic_3a", DataGridView1.Rows(i).Cells(13).Value)
                cmd16.Parameters.AddWithValue("@cantadi_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(14).Value))
                cmd16.ExecuteNonQuery()
            Next
            MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
            Dim hj2 As New flog
            Dim hj1 As New vlog
            hj1.gmodulo = "ORDEN_COMPRA"
            hj1.gcuser = MDIParent1.Label3.Text
            If Label22.Text = "0" Then
                hj1.gaccion = 1
            Else
                hj1.gaccion = 2
            End If

            hj1.gpc = My.Computer.Name
            hj1.gfecha = DateTimePicker3.Value
            hj1.gdetalle = TextBox1.Text & TextBox2.Text
            hj1.gccia = Label16.Text
            hj2.insertar_log(hj1)
            MessageBox.Show("Datos registrados correctamente ", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Dim respuesta As DialogResult

            respuesta = MessageBox.Show("DESEA IMPRIMIR LA ORDEN DE COMPRA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                Rpt_oc_1.TextBox1.Text = Label16.Text
                Rpt_oc_1.TextBox2.Text = Trim(TextBox1.Text)
                Rpt_oc_1.TextBox3.Text = Trim(TextBox2.Text)
                Rpt_oc_1.TextBox4.Text = Trim(TextBox1.Text)
                Rpt_oc_1.TextBox5.Text = Trim(TextBox2.Text)
                Rpt_oc_1.TextBox6.Text = ""
                Rpt_oc_1.TextBox7.Text = 0
                Rpt_oc_1.TextBox8.Text = "\\172.21.0.1\erpcaesoft$\LIB.APPSV2\imagenes\"
                Rpt_oc_1.Show()
            End If

            abrir()
            TextBox1.Text = TextBox1.Text
            Dim Rsr1991 As SqlDataReader
            Dim sql1011 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAlluLTIMONumeroOrdendeCompra '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "'"
            Dim cmd1011 As New SqlCommand(sql1011, conx)
            Rsr1991 = cmd1011.ExecuteReader()
            Rsr1991.Read()
            TextBox2.Text = Rsr1991(0)
            Rsr1991.Close()
            TextBox2.Select()
            LIMPIAR()
            Label21.Visible = False
            TextBox1.Enabled = True
            DateTimePicker1.Value = DateTimePicker3.Value
            DateTimePicker2.Value = DateTimePicker3.Value
            TextBox2.Enabled = True
            TextBox2.Select()
        End If

    End Sub
    Sub LIMPIAR()
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""

        DataGridView1.Rows.Clear()

    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim func As New ftcambio
        Dim dts As New vtcambio

        dts.gfecha = DateTimePicker1.Text
        TextBox9.Text = func.mostrar_tipo_cambio(dts)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        log.Label1.Text = Trim(TextBox1.Text) & Trim(TextBox2.Text)
        log.Label2.Text = "ORDEN_COMPRA"
        log.Label3.Text = Label16.Text
        log.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Label20.Text = e.RowIndex
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Cant. Compra" Then
            Dim i, a9, fila As Integer
            Dim cant10, sum10, cant9, sum9, cant8, sum8, cant7, sum7 As Double
            i = DataGridView1.Rows.Count
            fila = e.RowIndex
            Try
                'DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                'DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                'DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(9).Value / 1.18
                'DataGridView1.Rows(fila).Cells(13).Value = DataGridView1.Rows(fila).Cells(9).Value - DataGridView1.Rows(fila).Cells(12).Value
                'Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                'Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                'Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                If CheckBox3.Checked = False And CheckBox2.Checked = True Then

                    DataGridView1.Rows(fila).Cells(12).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(11).Value)

                    DataGridView1.Rows(fila).Cells(15).Value = DataGridView1.Rows(fila).Cells(12).Value
                    DataGridView1.Rows(fila).Cells(17).Value = Val(DataGridView1.Rows(fila).Cells(12).Value * 1.18)
                    DataGridView1.Rows(fila).Cells(16).Value = DataGridView1.Rows(fila).Cells(17).Value - DataGridView1.Rows(fila).Cells(15).Value
                    DataGridView1.Rows(fila).Cells(14).Value = ((Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(13).Value)) / 100) + (Val(DataGridView1.Rows(fila).Cells(7).Value))
                    Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                Else
                    If CheckBox3.Checked = True And CheckBox2.Checked = True Then
                        DataGridView1.Rows(fila).Cells(12).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(11).Value)
                        DataGridView1.Rows(fila).Cells(17).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(11).Value)
                        DataGridView1.Rows(fila).Cells(15).Value = DataGridView1.Rows(fila).Cells(12).Value / 1.18
                        DataGridView1.Rows(fila).Cells(16).Value = DataGridView1.Rows(fila).Cells(17).Value - DataGridView1.Rows(fila).Cells(15).Value
                        DataGridView1.Rows(fila).Cells(14).Value = ((Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(13).Value)) / 100) + (Val(DataGridView1.Rows(fila).Cells(7).Value))

                        Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"

                    Else
                        If CheckBox3.Checked = False And CheckBox2.Checked = False Then
                            DataGridView1.Rows(fila).Cells(12).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(17).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(15).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(14).Value = ((Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(13).Value)) / 100) + (Val(DataGridView1.Rows(fila).Cells(7).Value))

                            DataGridView1.Rows(fila).Cells(16).Value = 0
                            Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                        End If
                    End If
                End If
                'DataGridView1.Focus()
                'DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(7)
                'DataGridView1.BeginEdit(True)


            Catch ex As Exception
                MsgBox("NO SELECCIONO UNA FILA CORRECTA")
            End Try
            For a9 = 0 To i - 1
                cant10 = Val(DataGridView1.Rows(a9).Cells(15).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(a9).Cells(16).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(a9).Cells(17).Value)
                sum8 = cant8 + Val(sum8)


            Next
            TextBox10.Text = sum10.ToString("N2")
            TextBox11.Text = sum9.ToString("N2")

            TextBox12.Text = sum8.ToString("N2")
        End If

        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "P. Unitario" Then
            Dim i, a9, fila As Integer
            Dim cant10, sum10, cant9, sum9, cant8, sum8, cant7, sum7 As Double
            i = DataGridView1.Rows.Count
            fila = e.RowIndex
            Try
                'DataGridView1.Rows(fila).Cells(9).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                'DataGridView1.Rows(fila).Cells(14).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                'DataGridView1.Rows(fila).Cells(12).Value = DataGridView1.Rows(fila).Cells(9).Value / 1.18
                'DataGridView1.Rows(fila).Cells(13).Value = DataGridView1.Rows(fila).Cells(9).Value - DataGridView1.Rows(fila).Cells(12).Value
                'Me.DataGridView1.Columns(9).DefaultCellStyle.Format = "n2"
                'Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                'Me.DataGridView1.Columns(13).DefaultCellStyle.Format = "n2"
                If CheckBox3.Checked = False And CheckBox2.Checked = True Then

                    DataGridView1.Rows(fila).Cells(12).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(11).Value)

                    DataGridView1.Rows(fila).Cells(15).Value = DataGridView1.Rows(fila).Cells(12).Value
                    DataGridView1.Rows(fila).Cells(17).Value = Val(DataGridView1.Rows(fila).Cells(12).Value * 1.18)
                    DataGridView1.Rows(fila).Cells(16).Value = DataGridView1.Rows(fila).Cells(17).Value - DataGridView1.Rows(fila).Cells(15).Value
                    DataGridView1.Rows(fila).Cells(14).Value = ((Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(13).Value)) / 100) + (Val(DataGridView1.Rows(fila).Cells(7).Value))
                    Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                Else
                    If CheckBox3.Checked = True And CheckBox2.Checked = True Then
                        DataGridView1.Rows(fila).Cells(12).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(11).Value)
                        DataGridView1.Rows(fila).Cells(17).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(11).Value)
                        DataGridView1.Rows(fila).Cells(15).Value = DataGridView1.Rows(fila).Cells(12).Value / 1.18
                        DataGridView1.Rows(fila).Cells(16).Value = DataGridView1.Rows(fila).Cells(17).Value - DataGridView1.Rows(fila).Cells(15).Value
                        DataGridView1.Rows(fila).Cells(14).Value = ((Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(13).Value)) / 100) + (Val(DataGridView1.Rows(fila).Cells(7).Value))

                        Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                        Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"

                    Else
                        If CheckBox3.Checked = False And CheckBox2.Checked = False Then
                            DataGridView1.Rows(fila).Cells(12).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(17).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(15).Value = Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(8).Value)
                            DataGridView1.Rows(fila).Cells(14).Value = ((Val(DataGridView1.Rows(fila).Cells(7).Value) * Val(DataGridView1.Rows(fila).Cells(13).Value)) / 100) + (Val(DataGridView1.Rows(fila).Cells(7).Value))

                            DataGridView1.Rows(fila).Cells(16).Value = 0
                            Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                            Me.DataGridView1.Columns(12).DefaultCellStyle.Format = "n2"
                        End If
                    End If
                End If
                'DataGridView1.Focus()
                'DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(7)
                'DataGridView1.BeginEdit(True)


            Catch ex As Exception
                MsgBox("NO SELECCIONO UNA FILA CORRECTA")
            End Try
            For a9 = 0 To i - 1
                cant10 = Val(DataGridView1.Rows(a9).Cells(15).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(a9).Cells(16).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(a9).Cells(17).Value)
                sum8 = cant8 + Val(sum8)


            Next
            TextBox10.Text = sum10.ToString("N2")
            TextBox11.Text = sum9.ToString("N2")

            TextBox12.Text = sum8.ToString("N2")
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            Label18.Visible = True
            TextBox13.Visible = True
            TextBox13.Select()
        Else
            Label18.Visible = False
            TextBox13.Visible = False
        End If
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
    Dim Rsr2 As SqlDataReader

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

    End Sub

    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim VAL As Integer
            Select Case Trim(TextBox13.Text).Length
                Case "1" : TextBox13.Text = "01" & "0000000" & TextBox13.Text
                Case "2" : TextBox13.Text = "01" & "000000" & TextBox13.Text
                Case "3" : TextBox13.Text = "01" & "00000" & TextBox13.Text
                Case "4" : TextBox13.Text = "01" & "0000" & TextBox13.Text
                Case "5" : TextBox13.Text = "01" & "000" & TextBox13.Text
                Case "6" : TextBox13.Text = "01" & "00" & TextBox13.Text
                Case "7" : TextBox13.Text = "01" & "0" & TextBox13.Text
                Case "8" : TextBox13.Text = "01" & TextBox13.Text
                Case "9" : TextBox13.Text = "0" & TextBox13.Text
            End Select
            Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, a.unid_17,program_3, cantp_3 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(TextBox13.Text) + "' and q.ccia ='" + Trim(Label16.Text) + "'
"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            Dim i5 As Integer
            i5 = DataGridView1.Rows.Count

            If Rsr2.Read() = True Then
                '    If i5 = 0 Then
                '        DataGridView1.Rows.Add()
                '        DataGridView1.Rows(i5).Cells(2).Value = Rsr2(4)
                '        DataGridView1.Rows(i5).Cells(3).Value = Rsr2(0)
                '        DataGridView1.Rows(i5).Cells(5).Value = Rsr2(1)
                '        DataGridView1.Rows(i5).Cells(6).Value = Rsr2(2)
                '        DataGridView1.Rows(i5).Cells(8).Value = Rsr2(3)
                '        DataGridView1.Rows(i5).Cells(10).Value = Rsr2(3)
                '        DataGridView1.Rows(i5).Cells(7).Value = Rsr2(5)
                '        DataGridView1.Rows(i5).Cells(9).Value = Rsr2(5)
                '        DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                '        DataGridView1.Rows(i5).Cells(13).Value = 0
                '        VAL = DataGridView1.Rows(i5).Cells(0).Value
                '        Select Case VAL.ToString.Length
                '            Case "1" : DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                '            Case "2" : DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                '            Case "3" : DataGridView1.Rows(i5).Cells(0).Value = VAL
                '        End Select

                '        TextBox13.Text = ""
                '    Else


                '        If Trim(DataGridView1.Rows(0).Cells(2).Value) = Trim(TextBox13.Text) Then
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i5).Cells(2).Value = Rsr2(4)
                DataGridView1.Rows(i5).Cells(3).Value = Rsr2(0)
                DataGridView1.Rows(i5).Cells(5).Value = Rsr2(1)
                DataGridView1.Rows(i5).Cells(6).Value = Rsr2(2)
                DataGridView1.Rows(i5).Cells(8).Value = Rsr2(3)
                DataGridView1.Rows(i5).Cells(10).Value = Rsr2(3)
                DataGridView1.Rows(i5).Cells(7).Value = Rsr2(5)
                DataGridView1.Rows(i5).Cells(9).Value = Rsr2(5)
                DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                DataGridView1.Rows(i5).Cells(13).Value = 0
                VAL = DataGridView1.Rows(i5).Cells(0).Value
                Select Case VAL.ToString.Length
                    Case "1" : DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                    Case "2" : DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                    Case "3" : DataGridView1.Rows(i5).Cells(0).Value = VAL
                End Select

                TextBox13.Text = ""
                Rsr2.Close()
            Else
                '        MsgBox("SOLO SE DEBE AGREGAR PRODUCTOS QUE PERTENECEN A LA MISMA PO")
                '    End If
                'End If
                'Rsr2.Close()
                'Else
                Rsr2.Close()
                MsgBox("LA OP INGRESADA NO EXISTE")
            End If
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged

    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        If CheckBox2.Checked = False Then

            Dim i8 As Integer
            i8 = DataGridView1.RowCount
            CheckBox2.Enabled = True
            CheckBox3.Checked = False
            CheckBox3.Enabled = False
            For i = 0 To i8 - 1

                DataGridView1.Rows(i).Cells(17).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(11).Value
                DataGridView1.Rows(i).Cells(15).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(11).Value
                Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                DataGridView1.Rows(i).Cells(16).Value = "0.00"

            Next

            For a9 = 0 To i8 - 1
                cant10 = Val(DataGridView1.Rows(a9).Cells(15).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(a9).Cells(16).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(a9).Cells(17).Value)
                sum8 = cant8 + Val(sum8)

            Next
            TextBox10.Text = sum10.ToString("N2")
            TextBox11.Text = sum9.ToString("N2")

            TextBox12.Text = sum8.ToString("N2")
        Else
            If CheckBox2.Checked = True Then

                CheckBox3.Enabled = True
                Dim i8 As Integer
                i8 = DataGridView1.RowCount
                For i = 0 To i8 - 1

                    DataGridView1.Rows(i).Cells(15).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(11).Value

                    DataGridView1.Rows(i).Cells(17).Value = DataGridView1.Rows(i).Cells(15).Value * 1.18
                    DataGridView1.Rows(i).Cells(16).Value = DataGridView1.Rows(i).Cells(17).Value - DataGridView1.Rows(i).Cells(15).Value
                    Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                Next

                For a9 = 0 To i8 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(15).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(16).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(17).Value)
                    sum8 = cant8 + Val(sum8)

                Next
                TextBox10.Text = sum10.ToString("N2")
                TextBox11.Text = sum9.ToString("N2")

                TextBox12.Text = sum8.ToString("N2")
            End If

        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        If CheckBox1.Checked = True Then
            MsgBox("NO SE PUEDE MODIFICAR LA ORDEN DE SERVICIO QUE ESTA APROBADA")
        Else
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            TextBox10.Enabled = True
            TextBox11.Enabled = True
            TextBox12.ReadOnly = False
            DataGridView1.ReadOnly = False
            DataGridView1.Enabled = True
            ComboBox1.Enabled = True
            TextBox13.Enabled = True
            Label22.Text = "1"
        End If


        DataGridView1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
        DataGridView1.BeginEdit(True)
    End Sub
    Sub ENVIAR_CORREO()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        'Dim corre As String
        'jj.gvendedor = TextBox8.Text
        'corre = fk.buscar_correo(jj)

        Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA ORDEN DE SERVICIO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> ANULADO <br/>
<FONT COLOR='blue'>* ALMACEN : </FONT>" + TextBox1.Text + " <br/>
<FONT COLOR='blue'>* ORDEN : </FONT>" + Trim(TextBox1.Text) & Trim(TextBox2.Text) + " <br/>
<FONT COLOR='blue'>* USUARIO : </FONT>" + Trim(MDIParent1.Label3.Text) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(MDIParent1.Label7.Text) + " <br/>
<FONT COLOR='blue'>* PROVEEDOR : </FONT>" + Trim(TextBox4.Text) + "<br/>
</body>

</html>"
        With message
            .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
            .To.Add("fmestanza@viannysac.com,hrivera@viannysac.com")
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = "ORDEN COMPRA N°" & Trim(TextBox1.Text) & Trim(TextBox2.Text)
            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp

            .EnableSsl = True
            .Port = "587"
            .Host = "smtppro.zoho.com"
            .Credentials = New Net.NetworkCredential("admin@viannysac.com", "20Via$&it2")
            .Send(message)
        End With

        Try
            MessageBox.Show("Su mensaje de correo ha sido enviado.", "Correo enviado",
                             MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
                             MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If CheckBox1.Checked = True Then
            MsgBox("NO SE PUEDE ANULAR LA ORDEN DE COMPRA QUE ESTA APROBADA")
        Else
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("QUIERES ANULAR LA ORDEN DE COMPRA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Dim cmd1577 As New SqlCommand("update custom_vianny.dbo.lag0300 set flag_3=0 where ncom_3 =@ncom4 and ccia_3=@ccia_4", conx)
                cmd1577.Parameters.AddWithValue("@ncom4", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd1577.Parameters.AddWithValue("@ccia_4", Trim(Label16.Text))
                cmd1577.ExecuteNonQuery()

                Dim cmd15776 As New SqlCommand("update custom_vianny.dbo.lap0300 set flag_3a=0 where ncom_3a =@ncom4 and ccia_3a=@ccia_4", conx)
                cmd15776.Parameters.AddWithValue("@ncom4", Trim(TextBox1.Text) & Trim(TextBox2.Text))
                cmd15776.Parameters.AddWithValue("@ccia_4", Trim(Label16.Text))
                cmd15776.ExecuteNonQuery()
                MsgBox("SE ANULO LA INFORMACION CORRECTAMENTE")
                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "ORDEN_COMPRA"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = 3
                hj1.gpc = My.Computer.Name
                hj1.gfecha = DateTimePicker3.Value
                hj1.gdetalle = TextBox1.Text & TextBox2.Text
                hj1.gccia = Label16.Text
                hj2.insertar_log(hj1)
                ENVIAR_CORREO()
                TextBox1.Text = TextBox1.Text
                abrir()
                Dim Rsr1991 As SqlDataReader
                Dim sql1011 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAlluLTIMONumeroOrdendeCompra '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "'"
                Dim cmd1011 As New SqlCommand(sql1011, conx)
                Rsr1991 = cmd1011.ExecuteReader()
                Rsr1991.Read()
                TextBox2.Text = Rsr1991(0)
                Rsr1991.Close()
                TextBox2.Select()

                LIMPIAR()
            End If
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Rpt_oc_1.TextBox1.Text = Label16.Text
        Rpt_oc_1.TextBox2.Text = Trim(TextBox1.Text)
        Rpt_oc_1.TextBox3.Text = Trim(TextBox2.Text)
        Rpt_oc_1.TextBox4.Text = Trim(TextBox1.Text)
        Rpt_oc_1.TextBox5.Text = Trim(TextBox2.Text)
        Rpt_oc_1.TextBox6.Text = ""
        Rpt_oc_1.TextBox7.Text = 0
        Rpt_oc_1.TextBox8.Text = "\\172.21.0.1\erpcaesoft$\LIB.APPSV2\imagenes\"
        Rpt_oc_1.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        LIMPIAR()
        TextBox1.Text = TextBox1.Text
        abrir()
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "EXEC custom_vianny.dbo.CaeSoft_GetAlluLTIMONumeroOrdendeCompra '" + Trim(Label16.Text) + "','" + Trim(TextBox1.Text) + "'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        Rsr1991.Read()
        TextBox2.Text = Rsr1991(0)
        Rsr1991.Close()


        Label21.Visible = False
        TextBox1.Enabled = True

        TextBox2.Enabled = True
        TextBox2.Select()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim func As New ftcambio
        Dim dts As New vtcambio

        dts.gfecha = DateTimePicker1.Text
        TextBox9.Text = func.mostrar_tipo_cambio(dts)
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


    Private Sub Ocs_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Me.Activate()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        Dim I18, VAL1 As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
            I18 = DataGridView1.Rows.Count
            For i1 = 0 To I18 - 1
                VAL1 = i1 + 1
                Select Case VAL1.ToString.Length
                    Case "1" : DataGridView1.Rows(i1).Cells(0).Value = "00" & "" & VAL1
                    Case "2" : DataGridView1.Rows(i1).Cells(0).Value = "0" & "" & VAL1
                    Case "3" : DataGridView1.Rows(i1).Cells(0).Value = VAL1
                End Select

                cant10 = Val(DataGridView1.Rows(i1).Cells(15).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(i1).Cells(16).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(i1).Cells(17).Value)
                sum8 = cant8 + Val(sum8)
            Next


            TextBox10.Text = sum10.ToString("N2")
            TextBox11.Text = sum9.ToString("N2")

            TextBox12.Text = sum8.ToString("N2")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        If e.ColumnIndex = 7 Or e.ColumnIndex = 9 Or e.ColumnIndex = 11 Then
            Dim cellValue As String = e.FormattedValue.ToString()
            If Not IsValidNumber(cellValue) Then

                MessageBox.Show("El valor ingresado no es válido. Debe ser un número entero o decimal con al menos un número antes del punto decimal.", "Error de Validación", MessageBoxButtons.OK, MessageBoxIcon.Error)

                e.Cancel = True
            End If
        End If
    End Sub
    Private Function IsValidNumber(value As String) As Boolean
        Dim pattern As String = "^\d+(\.\d+)?$"
        Return System.Text.RegularExpressions.Regex.IsMatch(value, pattern)
    End Function

    Private Sub CheckBox3_Click(sender As Object, e As EventArgs) Handles CheckBox3.Click
        Dim cant10, sum10, cant9, sum9, cant8, sum8 As Double
        If CheckBox3.Checked = True Then
            Dim i8 As Integer
            i8 = DataGridView1.RowCount
            For i = 0 To i8 - 1
                DataGridView1.Rows(i).Cells(17).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(11).Value
                DataGridView1.Rows(i).Cells(15).Value = DataGridView1.Rows(i).Cells(17).Value / 1.18
                DataGridView1.Rows(i).Cells(16).Value = DataGridView1.Rows(i).Cells(17).Value - DataGridView1.Rows(i).Cells(15).Value
                Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"

            Next

            For a9 = 0 To i8 - 1
                cant10 = Val(DataGridView1.Rows(a9).Cells(15).Value)
                sum10 = cant10 + Val(sum10)
                cant9 = Val(DataGridView1.Rows(a9).Cells(16).Value)
                sum9 = cant9 + Val(sum9)
                cant8 = Val(DataGridView1.Rows(a9).Cells(17).Value)
                sum8 = cant8 + Val(sum8)

            Next
            TextBox10.Text = sum10.ToString("N2")
            TextBox11.Text = sum9.ToString("N2")

            TextBox12.Text = sum8.ToString("N2")
        Else
            If CheckBox3.Checked = False And CheckBox2.Checked = True Then

                Dim i81 As Integer
                i81 = DataGridView1.RowCount
                For i = 0 To i81 - 1

                    DataGridView1.Rows(i).Cells(15).Value = DataGridView1.Rows(i).Cells(7).Value * DataGridView1.Rows(i).Cells(11).Value

                    DataGridView1.Rows(i).Cells(17).Value = DataGridView1.Rows(i).Cells(17).Value * 1.18
                    DataGridView1.Rows(i).Cells(16).Value = DataGridView1.Rows(i).Cells(17).Value - DataGridView1.Rows(i).Cells(15).Value
                    Me.DataGridView1.Columns(15).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(16).DefaultCellStyle.Format = "n2"
                    Me.DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
                Next

                For a9 = 0 To i81 - 1
                    cant10 = Val(DataGridView1.Rows(a9).Cells(15).Value)
                    sum10 = cant10 + Val(sum10)
                    cant9 = Val(DataGridView1.Rows(a9).Cells(16).Value)
                    sum9 = cant9 + Val(sum9)
                    cant8 = Val(DataGridView1.Rows(a9).Cells(17).Value)
                    sum8 = cant8 + Val(sum8)

                Next
                TextBox10.Text = sum10.ToString("N2")
                TextBox11.Text = sum9.ToString("N2")

                TextBox12.Text = sum8.ToString("N2")
            End If
        End If
    End Sub

    Private Sub RecibirItems(ByVal items As List(Of Item))
        Dim num As Integer = DataGridView1.Rows.Count
        DataGridView1.Rows.Add(items.Count - 1)
        Dim currentRowIndex As Integer = num - 1
        For Each item As Item In items
            item.RowIndex = currentRowIndex
            ' Verifica si el índice ya existe en el DataGridView
            If item.RowIndex >= 0 AndAlso item.RowIndex < DataGridView1.Rows.Count Then


                ' Actualiza la fila específica en el índice RowIndex
                DataGridView1.Rows(item.RowIndex).Cells("Po").Value = item.Po
                DataGridView1.Rows(item.RowIndex).Cells("umc").Value = item.um
                DataGridView1.Rows(item.RowIndex).Cells("Um").Value = item.um
                DataGridView1.Rows(item.RowIndex).Cells("Codigo").Value = item.Codigo
                DataGridView1.Rows(item.RowIndex).Cells("Descipcion").Value = item.Descripcion
                DataGridView1.Rows(item.RowIndex).Cells("Cantidad").Value = item.Cantidad
                DataGridView1.Rows(item.RowIndex).Cells("Camp").Value = item.Cantidad
            Else

                ' Si el índice no es válido, agrega una nueva fila
                'DataGridView1.Rows.Add(item.Codigo, item.Descripcion, item.Cantidad)
            End If
            currentRowIndex += 1
        Next
    End Sub
    Private Sub EditingControl_KeyDown(sender As Object, e As KeyEventArgs)
        ' Verificar si la tecla presionada es F1
        If e.KeyCode = Keys.F1 Then
            ' Verificar si la celda actual está en la columna "Factor de Consumo"
            If DataGridView1.CurrentCell.ColumnIndex = 5 Then
                If DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(2).Value IsNot Nothing Then
                    If DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(2).Value.ToString().Trim().Length() > 0 Then
                        ' Cancelar la edición para evitar conflictos
                        DataGridView1.EndEdit()
                        ' Abrir el formulario deseado
                        Dim form As New RequerimientoPo
                        form.DataGridView1.DataSource = Nothing
                        form.Label2.Text = DataGridView1.CurrentCell.RowIndex.ToString
                        form.Label1.Text = Trim(Label16.Text)
                        form.Label3.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(2).Value.ToString().Trim()
                        If TextBox1.Text.ToString().Trim() = "13" Then
                            form.Label4.Text = "AV"
                        Else
                            If TextBox1.Text.ToString().Trim() = "08" Or TextBox1.Text.ToString().Trim() = "10" Then
                                form.Label4.Text = "TC"
                            End If

                        End If
                        form.TextBox1.Select()
                        AddHandler form.PasarItems, AddressOf RecibirItems
                        form.ShowDialog()
                    End If
                End If




            End If

        End If
    End Sub
End Class