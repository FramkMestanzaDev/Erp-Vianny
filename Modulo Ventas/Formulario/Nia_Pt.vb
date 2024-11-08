Imports System.Data.SqlClient
Public Class Nia_Pt
    Public conx As SqlConnection
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public conn As SqlDataAdapter
    Dim ty2 As New DataTable
    Public comando As SqlCommand
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Dim Rsr2, Rsr20, Rsr205, Rsr22, Rsr23, Rsr2333, Rsr24 As SqlDataReader
    Private formModified As Boolean = False

    Private Sub TextBox_Modified(sender As Object, e As EventArgs)
        ' Establecer que el formulario ha sido modificado
        formModified = True

        ' Identificar cuál TextBox lanzó el evento
        Dim modifiedTextBox As TextBox = CType(sender, TextBox)

        ' Mostrar el nombre del TextBox modificado
        MessageBox.Show("Se modificó el TextBox: " & modifiedTextBox.Name)


    End Sub
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Select Case ComboBox3.Text
            Case "01" : TextBox2.Text = "PRODUCTO TERMINADO"
            Case "32" : TextBox2.Text = "TIENDA PLANTA 1"
            Case "50" : TextBox2.Text = "PRODUCTO TERMINADO SEGUNDA"
        End Select

        TextBox4.Enabled = True
        TextBox4.ReadOnly = False
        TextBox4.Text = DateTimePicker1.Value.Month
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & TextBox4.Text
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select
        ty2.Clear()
        abrir()
        llenar_combo_box2(ComboBox1, ComboBox3.Text)
        TextBox4.Select()
    End Sub
    Sub llenar_combo_box2(ByVal cb As ComboBox, ByVal alm As String)
        Try

            conn = New SqlDataAdapter("select dest_21e,rtrim(ltrim(nomb_21e)) as nomb_21e from custom_vianny.dbo.cag210e where  Empr_21e ='" + Trim(Label13.Text) + "' AND almac_21e ='" + alm + "' order by nomb_21e", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "nomb_21e"
            ComboBox1.ValueMember = "dest_21e"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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
    Dim da As New DataTable
    Function nia()
        Dim camposFaltantes As New List(Of String) ' Lista para almacenar los nombres de los campos faltantes

        ' Validar campos obligatorios
        If String.IsNullOrEmpty(TextBox8.Text) Then
            camposFaltantes.Add(" RUC ")
        End If

        If String.IsNullOrEmpty(TextBox4.Text) Then
            camposFaltantes.Add(" MES ")
        End If

        If String.IsNullOrEmpty(TextBox4.Text) Then
            camposFaltantes.Add(" CORRELATIVO ")
        End If

        ' Agregar más validaciones para otros campos obligatorios según sea necesario

        ' Verificar si hay campos faltantes
        If camposFaltantes.Count > 0 Then
            ' Mostrar mensaje de error indicando los campos faltantes
            Dim camposFaltantesStr As String = String.Join(", ", camposFaltantes)
            MessageBox.Show("Falta ingresar información en los siguientes campos: " & camposFaltantesStr, "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            abrir()
            da.Clear()

            Dim cmd17 As New SqlCommand("IF NOT EXISTS( SELECT NCom_3 FROM custom_vianny.DBO.Mag030f WHERE CCia_3 = @CCIA_3 AND CPeriodo_3 = @CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 = @ccom_3 AND NCom_3 = @ncom_3)
        INSERT INTO custom_vianny.DBO.MAG030F (CCIA_3 , CPERIODO_3, talm_3, ccom_3, ncom_3, fcom_3, gloa_3, glob_3,fich_3 , orig_3, nombe_3, adevol_3, transf_3, TDoc_3, grm_3, fase_3, flag_3, origen_3, usuario_3, fecha_3, hora_3, subfase_3, tidoc_3,sfactu_3, nfactu_3, transfer_3, pedorig_3, Auto_3) 
                     VALUES (@CCIA_3,@CPERIODO_3,@TAlm_3,@ccom_3,@ncom_3,@fcom_3,@gloa_3,'',@fich_3,@orig_3, '', 0, 0,@TDoc_3,@grm_3,'',1,@origen_3,@usuario_3,@fecha_3,@hora_3, '', '', '', '', 0, '', 0)
        				ELSE
        					UPDATE custom_vianny.DBO.MAG030F SET fcom_3  = @fcom_3, 
        					                   gloa_3  = @gloa_3, 
        					                   glob_3  = '', 
        					                   fich_3  = @fich_3, 
        					                   orig_3  = @orig_3, 
        					                   nombe_3 = '', 
        					                   adevol_3  = 0, 
        					                   transf_3  = 0, 
        					                   TDoc_3    =@TDoc_3,
        					                   grm_3     = @grm_3, 
        					                   fase_3    = '',
        					                   flag_3    = 1, 
        					                   origen_3  = @origen_3, 
        					                   usuario_3 = @usuario_3, 
        					                   fecha_3   =@fecha_3, 
        					                   hora_3    = @hora_3, 
        					                   subfase_3 = '', 
        					                   tidoc_3   = '', 
        					                   sfactu_3  = '', 
        					                   nfactu_3  = '', 
        					                   transfer_3= 0, 
        					                   pedorig_3 = '',
        					                   Auto_3 = 0
        									   WHERE CCIA_3 = @CCIA_3 And CPERIODO_3 =@CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 =@ccom_3 And ncom_3 = @ncom_3", conx)
            cmd17.Parameters.AddWithValue("@CCIA_3", Trim(Label13.Text))
            cmd17.Parameters.AddWithValue("@CPERIODO_3", Trim(Label11.Text))
            cmd17.Parameters.AddWithValue("@TAlm_3", Trim(ComboBox3.Text))
            cmd17.Parameters.AddWithValue("@ccom_3", "1")
            cmd17.Parameters.AddWithValue("@ncom_3", Trim(TextBox4.Text) & Trim(TextBox5.Text))
            cmd17.Parameters.AddWithValue("@fcom_3", DateTimePicker1.Value.Date)
            cmd17.Parameters.AddWithValue("@gloa_3", Trim(TextBox9.Text))
            cmd17.Parameters.AddWithValue("@orig_3", ComboBox1.SelectedValue.ToString)
            cmd17.Parameters.AddWithValue("@fich_3", Trim(TextBox8.Text))
            If Trim(ComboBox2.Text) = "INTERNA" Then
                cmd17.Parameters.AddWithValue("@origen_3", "INT")
            Else
                cmd17.Parameters.AddWithValue("@origen_3", "EXT")
            End If
            Dim doc As String
            Select Case ComboBox4.Text
                Case "GUIA" : doc = "009"
                Case "FACT" : doc = "001"
                Case "OTRO" : doc = "002"
                Case "" : doc = ""
            End Select
            cmd17.Parameters.AddWithValue("@TDoc_3", Trim(doc))
            cmd17.Parameters.AddWithValue("@grm_3", Trim(TextBox12.Text))
            cmd17.Parameters.AddWithValue("@usuario_3", Trim(TextBox7.Text))
            cmd17.Parameters.AddWithValue("@fecha_3", DateTimePicker1.Value.Date)
            cmd17.Parameters.AddWithValue("@hora_3", DateTime.Now.ToString("hh:mm:ss"))
            cmd17.ExecuteNonQuery()


            Dim p As Integer
            p = DataGridView1.Rows.Count


            'If Trim(TextBox22.Text) = "2" Then


            '    Dim cmd As New SqlDataAdapter("SELECT ccia_3b,cperiodo_3b,linea_3b, pedido_3b,talm_3b,cant_3b,cant1_3b,cant2_3b,cant3_3b,cant4_3b,cant5_3b,cant6_3b,cant7_3b,cant8_3b,cant9_3b,cant10_3b  FROM custom_vianny.DBO.Mat030f A 
            '  Where A.CCia_3b = '" + Trim(Label13.Text) + "' AND A.CPeriodo_3b = '" + Trim(Label11.Text) + "' AND A.TAlm_3b = '" + Trim(ComboBox3.Text) + "' AND A.CCom_3b = '1' AND A.NCom_3b = '" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "'", conx)
            '    cmd.Fill(da)
            '    DataGridView11.DataSource = da

            '    Dim oi As Integer
            '    oi = DataGridView11.Rows.Count

            '    For i1 = 0 To oi - 1

            '    Next
            'End If
            Dim agregar As String = "DELETE custom_vianny.DBO.Map030f Where CCIA_3a = '" + Trim(Label13.Text) + "' And CPERIODO_3a = '" + Trim(Label11.Text) + "' And talm_3a = '" + Trim(ComboBox3.Text) + "' And ccom_3a = '1' And ncom_3a = '" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "'"
            Dim agregar1 As String = "DELETE custom_vianny.DBO.Mat030f Where CCIA_3b = '" + Trim(Label13.Text) + "' And CPERIODO_3b = '" + Trim(Label11.Text) + "' And talm_3b = '" + Trim(ComboBox3.Text) + "' And ccom_3b = '1' And ncom_3b = '" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "'"
            ELIMINAR(agregar)
            ELIMINAR(agregar1)
            For i = 0 To p - 1
                Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.DBO.Mat030f (ccia_3b , cperiodo_3b, talm_3b, ccom_3b, ncom_3b, item_3b,linea_3b ,pedido_3b ,cant_3b ,cant1_3b ,cant2_3b ,cant3_3b ,cant4_3b ,cant5_3b ,cant6_3b ,cant7_3b ,cant8_3b ,cant9_3b ,cant10_3b,ordens_3b,flag_3b)
                VALUES (@ccia_3b,@cperiodo_3b,@talm_3b,@ccom_3b,@ncom_3b,@item_3b,@linea_3b,@pedido_3b,@cant_3b,@cant1_3b,@cant2_3b,@cant3_3b,@cant4_3b,@cant5_3b,@cant6_3b,@cant7_3b,@cant8_3b,@cant9_3b,@cant10_3b,'', 1)", conx)
                cmd15.Parameters.AddWithValue("@ccia_3b", Trim(Label13.Text))
                cmd15.Parameters.AddWithValue("@cperiodo_3b", Trim(Label11.Text))
                cmd15.Parameters.AddWithValue("@talm_3b", Trim(ComboBox3.Text))
                cmd15.Parameters.AddWithValue("@ccom_3b", 1)
                cmd15.Parameters.AddWithValue("@ncom_3b", Trim(TextBox4.Text) & Trim(TextBox5.Text))
                cmd15.Parameters.AddWithValue("@item_3b", Trim(DataGridView1.Rows(i).Cells(0).Value))
                cmd15.Parameters.AddWithValue("@linea_3b", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd15.Parameters.AddWithValue("@pedido_3b", Trim(DataGridView1.Rows(i).Cells(1).Value))
                cmd15.Parameters.AddWithValue("@cant_3b", DataGridView1.Rows(i).Cells(14).Value)
                cmd15.Parameters.AddWithValue("@cant1_3b", DataGridView1.Rows(i).Cells(4).Value)
                cmd15.Parameters.AddWithValue("@cant2_3b", DataGridView1.Rows(i).Cells(5).Value)
                cmd15.Parameters.AddWithValue("@cant3_3b", DataGridView1.Rows(i).Cells(6).Value)
                cmd15.Parameters.AddWithValue("@cant4_3b", DataGridView1.Rows(i).Cells(7).Value)
                cmd15.Parameters.AddWithValue("@cant5_3b", DataGridView1.Rows(i).Cells(8).Value)
                cmd15.Parameters.AddWithValue("@cant6_3b", DataGridView1.Rows(i).Cells(9).Value)
                cmd15.Parameters.AddWithValue("@cant7_3b", DataGridView1.Rows(i).Cells(10).Value)
                cmd15.Parameters.AddWithValue("@cant8_3b", DataGridView1.Rows(i).Cells(11).Value)
                cmd15.Parameters.AddWithValue("@cant9_3b", DataGridView1.Rows(i).Cells(12).Value)
                cmd15.Parameters.AddWithValue("@cant10_3b", DataGridView1.Rows(i).Cells(13).Value)
                cmd15.ExecuteNonQuery()





                For y = 4 To 13
                    If DataGridView1.Rows(i).Cells(y).Value > 0 Then
                        Dim cmd16 As New SqlCommand("INSERT INTO custom_vianny.DBO.MAP030F (CCIA_3A , CPERIODO_3A, talm_3a, ccom_3a, ncom_3a, item_3a,linea_3a ,sinon_3a,linea2_3a, cant_3a, unid_3a, cantu_3a, obser1_3a, obser2_3a, obser3_3a, cantk_3a, cantke_3a, pedido_3a, flag_3a, peso_3a,tpeso_3a,unidk_3a , pieza_3a, color_3a, ncolor_3a, parti_3a, ocorte_3a, ordens_3a, nrollo_3a, kgrollo_3a, lote_3a, maquina_3a, galga_3a,sigla_3a, sitcal_3a, cantc_3a, ancho_3a, densid_3a, pedod_3a, ordtej_3a  , tipapt_3a , combo_3a,ccosto_3a,PreUn1_3a,Tot1_3a  )
                   Values (@CCIA_3A,@CPERIODO_3A,@talm_3a,@ccom_3a,@ncom_3aa,@item_3a,@linea_3a, ''     ,'PT0101' ,@cantid ,@unid_3a,  @cantid,''        , ''       ,''        ,@cantid  , 0.00     ,@pedido_3a, 1      , 0.00000, 0.00   ,@unidk_3a,   ''    , ''      , ''       , ''      , ''       , ''       , ''       , 0.00      , ''     , ''        , 0.00    ,''      , 0        , 0.000   , 0.000   , 0.000    , ''      , ''         , ''        , ''      ,''       ,0.00000  , 0.00)", conx)
                        cmd16.Parameters.AddWithValue("@CCIA_3A", Trim(Label13.Text))
                        cmd16.Parameters.AddWithValue("@CPERIODO_3A", Trim(Label11.Text))
                        cmd16.Parameters.AddWithValue("@talm_3a", Trim(ComboBox3.Text))
                        cmd16.Parameters.AddWithValue("@ccom_3a", "1")
                        cmd16.Parameters.AddWithValue("@ncom_3aa", Trim(TextBox4.Text) & Trim(TextBox5.Text))
                        cmd16.Parameters.AddWithValue("@item_3a", Trim(DataGridView1.Rows(i).Cells(0).Value))
                        cmd16.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
                        cmd16.Parameters.AddWithValue("@cantid", DataGridView1.Rows(i).Cells(y).Value)
                        cmd16.Parameters.AddWithValue("@unid_3a", "UND")
                        cmd16.Parameters.AddWithValue("@pedido_3a", Trim(DataGridView1.Rows(i).Cells(1).Value))
                        cmd16.Parameters.AddWithValue("@unidk_3a", "UND")
                        cmd16.ExecuteNonQuery()
                    End If

                Next
            Next

            MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
            TextBox22.Text = "1"
            Dim respuesta As DialogResult

            respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                pTERMINADO.TextBox1.Text = Label13.Text
                pTERMINADO.TextBox2.Text = Label11.Text
                pTERMINADO.TextBox3.Text = ComboBox3.Text
                pTERMINADO.TextBox4.Text = "1"
                pTERMINADO.TextBox5.Text = Trim(TextBox4.Text) & Trim(TextBox5.Text)
                pTERMINADO.Show()
            End If
            correlativo()
            DataGridView11.DataSource = Nothing
            limpiar()
            BLOQUEARC()
        End If
    End Function
    Function nsa()
        abrir()
        da.Clear()
        Dim cmd17 As New SqlCommand("IF NOT EXISTS( SELECT NCom_3 FROM custom_vianny.DBO.Mag030f WHERE CCia_3 = @CCIA_3 AND CPeriodo_3 = @CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 = @ccom_3 AND NCom_3 = @ncom_3)
        INSERT INTO custom_vianny.DBO.MAG030F (CCIA_3 , CPERIODO_3, talm_3, ccom_3, ncom_3, fcom_3, gloa_3, glob_3,fich_3 , orig_3, nombe_3, adevol_3, transf_3, TDoc_3, grm_3, fase_3, flag_3, origen_3, usuario_3, fecha_3, hora_3, subfase_3, tidoc_3,sfactu_3, nfactu_3, transfer_3, pedorig_3, Auto_3,orides_3,motdes_3) 
                     VALUES (@CCIA_3,@CPERIODO_3,@TAlm_3,@ccom_3,@ncom_3,@fcom_3,@gloa_3,'',@fich_3,@orig_3, '', 0, 0,@TDoc_3,@grm_3,'',1,@origen_3,@usuario_3,@fecha_3,@hora_3, '', '', '', '', 0, '', 0,@orides_3,@motdes_3)
        				ELSE
        					UPDATE custom_vianny.DBO.MAG030F SET fcom_3  = @fcom_3, 
        					                   gloa_3  = @gloa_3, 
        					                   glob_3  = '', 
        					                   fich_3  = @fich_3, 
        					                   orig_3  = @orig_3, 
        					                   nombe_3 = '', 
        					                   adevol_3  = 0, 
        					                   transf_3  = 0, 
        					                   TDoc_3    =@TDoc_3,
        					                   grm_3     = @grm_3, 
        					                   fase_3    = '',
        					                   flag_3    = 1, 
        					                   origen_3  = @origen_3, 
        					                   usuario_3 = @usuario_3, 
        					                   fecha_3   =@fecha_3, 
        					                   hora_3    = @hora_3, 
        					                   subfase_3 = '', 
        					                   tidoc_3   = '', 
        					                   sfactu_3  = '', 
        					                   nfactu_3  = '', 
        					                   transfer_3= 0, 
        					                   pedorig_3 = '',
        					                   Auto_3 = 0
        									   WHERE CCIA_3 = @CCIA_3 And CPERIODO_3 =@CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 =@ccom_3 And ncom_3 = @ncom_3", conx)
        cmd17.Parameters.AddWithValue("@CCIA_3", Trim(Label13.Text))
        cmd17.Parameters.AddWithValue("@CPERIODO_3", Trim(Label11.Text))
        cmd17.Parameters.AddWithValue("@TAlm_3", Trim(ComboBox3.Text))
        cmd17.Parameters.AddWithValue("@ccom_3", "2")
        cmd17.Parameters.AddWithValue("@ncom_3", Trim(TextBox4.Text) & Trim(TextBox5.Text))
        cmd17.Parameters.AddWithValue("@fcom_3", DateTimePicker1.Value.Date)
        cmd17.Parameters.AddWithValue("@gloa_3", Trim(TextBox9.Text))
        cmd17.Parameters.AddWithValue("@orig_3", Trim(ComboBox1.SelectedValue.ToString))
        cmd17.Parameters.AddWithValue("@fich_3", Trim(TextBox8.Text))
        cmd17.Parameters.AddWithValue("orides_3", "32")
        cmd17.Parameters.AddWithValue("motdes_3", Trim(TextBox4.Text) & Trim(TextBox5.Text))
        If Trim(ComboBox2.Text) = "INTERNA" Then
            cmd17.Parameters.AddWithValue("@origen_3", "INT")
        Else
            cmd17.Parameters.AddWithValue("@origen_3", "EXT")
        End If
        Dim doc As String
        Select Case ComboBox4.Text
            Case "GUIA" : doc = "009"
            Case "FACT" : doc = "001"
            Case "OTRO" : doc = "002"
            Case "" : doc = ""
        End Select
        cmd17.Parameters.AddWithValue("@TDoc_3", Trim(doc))
        cmd17.Parameters.AddWithValue("@grm_3", Trim(TextBox12.Text))
        cmd17.Parameters.AddWithValue("@usuario_3", Trim(TextBox7.Text))
        cmd17.Parameters.AddWithValue("@fecha_3", DateTimePicker1.Value)
        cmd17.Parameters.AddWithValue("@hora_3", DateTime.Now.ToString("hh:mm:ss"))

        cmd17.ExecuteNonQuery()



        Dim agregar As String = "DELETE custom_vianny.DBO.Map030f Where CCIA_3a = '" + Trim(Label13.Text) + "' And CPERIODO_3a = '" + Trim(Label11.Text) + "' And talm_3a = '" + Trim(ComboBox3.Text) + "' And ccom_3a = '2' And ncom_3a = '" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "'"
        Dim agregar1 As String = "DELETE custom_vianny.DBO.Mat030f Where CCIA_3b = '" + Trim(Label13.Text) + "' And CPERIODO_3b = '" + Trim(Label11.Text) + "' And talm_3b = '" + Trim(ComboBox3.Text) + "' And ccom_3b = '2' And ncom_3b = '" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "'"
        ELIMINAR(agregar)
        ELIMINAR(agregar1)
        Dim p As Integer
        p = DataGridView1.Rows.Count

        For i = 0 To p - 1
            Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.DBO.Mat030f (ccia_3b , cperiodo_3b, talm_3b, ccom_3b, ncom_3b, item_3b,linea_3b ,pedido_3b ,cant_3b ,cant1_3b ,cant2_3b ,cant3_3b ,cant4_3b ,cant5_3b ,cant6_3b ,cant7_3b ,cant8_3b ,cant9_3b ,cant10_3b,ordens_3b,flag_3b)
                        VALUES (@ccia_3b,@cperiodo_3b,@talm_3b,@ccom_3b,@ncom_3b,@item_3b,@linea_3b,@pedido_3b,@cant_3b,@cant1_3b,@cant2_3b,@cant3_3b,@cant4_3b,@cant5_3b,@cant6_3b,@cant7_3b,@cant8_3b,@cant9_3b,@cant10_3b,'', 1)", conx)
            cmd15.Parameters.AddWithValue("@ccia_3b", Trim(Label13.Text))
            cmd15.Parameters.AddWithValue("@cperiodo_3b", Trim(Label11.Text))
            cmd15.Parameters.AddWithValue("@talm_3b", Trim(ComboBox3.Text))
            cmd15.Parameters.AddWithValue("@ccom_3b", 2)
            cmd15.Parameters.AddWithValue("@ncom_3b", Trim(TextBox4.Text) & Trim(TextBox5.Text))
            cmd15.Parameters.AddWithValue("@item_3b", Trim(DataGridView1.Rows(i).Cells(0).Value))
            cmd15.Parameters.AddWithValue("@linea_3b", Trim(DataGridView1.Rows(i).Cells(2).Value))
            cmd15.Parameters.AddWithValue("@pedido_3b", Trim(DataGridView1.Rows(i).Cells(1).Value))
            cmd15.Parameters.AddWithValue("@cant_3b", DataGridView1.Rows(i).Cells(14).Value)
            cmd15.Parameters.AddWithValue("@cant1_3b", DataGridView1.Rows(i).Cells(4).Value)
            cmd15.Parameters.AddWithValue("@cant2_3b", DataGridView1.Rows(i).Cells(5).Value)
            cmd15.Parameters.AddWithValue("@cant3_3b", DataGridView1.Rows(i).Cells(6).Value)
            cmd15.Parameters.AddWithValue("@cant4_3b", DataGridView1.Rows(i).Cells(7).Value)
            cmd15.Parameters.AddWithValue("@cant5_3b", DataGridView1.Rows(i).Cells(8).Value)
            cmd15.Parameters.AddWithValue("@cant6_3b", DataGridView1.Rows(i).Cells(9).Value)
            cmd15.Parameters.AddWithValue("@cant7_3b", DataGridView1.Rows(i).Cells(10).Value)
            cmd15.Parameters.AddWithValue("@cant8_3b", DataGridView1.Rows(i).Cells(11).Value)
            cmd15.Parameters.AddWithValue("@cant9_3b", DataGridView1.Rows(i).Cells(12).Value)
            cmd15.Parameters.AddWithValue("@cant10_3b", DataGridView1.Rows(i).Cells(13).Value)
            cmd15.ExecuteNonQuery()


            Dim cmd18 As New SqlCommand("IF LEN(@Codigo_la) > 0
                	BEGIN					
                			IF NOT EXISTS ( SELECT Codigo_la FROM custom_vianny.DBO.Cag1700_Almac_Lotes WHERE CCia_la = @CCia_la AND Almac_la = @Almac_la AND Codigo_la =@Codigo_la AND Lote_la = @op )
                				INSERT INTO custom_vianny.DBO.Cag1700_Almac_Lotes (CCia_la, Almac_la, Codigo_la, Lote_la, Stock_la, Stock1_la, Stock2_la, Stock3_la, Stock4_la, Stock5_la, Stock6_la, Stock7_la, Stock8_la, Stock9_la, Stock10_la)
                										                 VALUES (@CCia_la, @Almac_la, @Codigo_la, @op, @stock  , @stock1  , @stock2  , @stock3  , @stock4  , @stock5  , @stock6  , @stock7  , @stock8  , @stock9  ,@stock10)
                			ELSE
                				UPDATE custom_vianny.DBO.Cag1700_Almac_Lotes SET Stock_la = Stock_la - @stock,
                											   Stock1_la = Stock1_la - @stock1,
                											   Stock2_la = Stock2_la - @stock2,
                											   Stock3_la = Stock3_la - @stock3,
                											   Stock4_la = Stock4_la - @stock4,
                											   Stock5_la = Stock5_la - @stock5,
                											   Stock6_la = Stock6_la - @stock6,
                											   Stock7_la = Stock7_la - @stock7,
                											   Stock8_la = Stock8_la - @stock8,
                											   Stock9_la = Stock9_la - @stock9,
                											   Stock10_la = Stock10_la - @stock10
                									   WHERE CCia_la =@CCia_la AND Almac_la =@Almac_la AND Codigo_la =@Codigo_la AND Lote_la =@op
                	END", conx)
            cmd18.Parameters.AddWithValue("@Codigo_la", Trim(DataGridView1.Rows(i).Cells(2).Value))
            cmd18.Parameters.AddWithValue("@ccom", "2")
            cmd18.Parameters.AddWithValue("@CCia_la", Trim(Label13.Text))
            cmd18.Parameters.AddWithValue("@Almac_la", Trim(ComboBox3.Text))
            cmd18.Parameters.AddWithValue("@op", Trim(DataGridView1.Rows(i).Cells(1).Value))
            cmd18.Parameters.AddWithValue("@stock", DataGridView1.Rows(i).Cells(14).Value)
            cmd18.Parameters.AddWithValue("@stock1", DataGridView1.Rows(i).Cells(4).Value)
            cmd18.Parameters.AddWithValue("@stock2", DataGridView1.Rows(i).Cells(5).Value)
            cmd18.Parameters.AddWithValue("@stock3", DataGridView1.Rows(i).Cells(6).Value)
            cmd18.Parameters.AddWithValue("@stock4", DataGridView1.Rows(i).Cells(7).Value)
            cmd18.Parameters.AddWithValue("@stock5", DataGridView1.Rows(i).Cells(8).Value)
            cmd18.Parameters.AddWithValue("@stock6", DataGridView1.Rows(i).Cells(9).Value)
            cmd18.Parameters.AddWithValue("@stock7", DataGridView1.Rows(i).Cells(10).Value)
            cmd18.Parameters.AddWithValue("@stock8", DataGridView1.Rows(i).Cells(11).Value)
            cmd18.Parameters.AddWithValue("@stock9", DataGridView1.Rows(i).Cells(12).Value)
            cmd18.Parameters.AddWithValue("@stock10", DataGridView1.Rows(i).Cells(13).Value)
            cmd18.ExecuteNonQuery()


            Dim cmd19 As New SqlCommand("IF LEN(@Codigo_la) > 0
                BEGIN					
                      IF NOT EXISTS ( SELECT Codigo_st FROM custom_vianny.DBO.Cag1700_Stock WHERE CCia_st =@CCia_la AND Almac_st = @Almac_la AND Codigo_st = @Codigo_la)
                			INSERT INTO custom_vianny.DBO.Cag1700_Stock (CCia_st,
                									   Almac_st,
                									   Codigo_st,
                									   Stock_st,
                									   Stock1_st, Stock2_st, Stock3_st, Stock4_st, Stock5_st, Stock6_st, Stock7_st, Stock8_st, Stock9_st, Stock10_st)
                							  VALUES (@CCia_la,
                									  @Almac_la,
                									  @Codigo_la,
                									  -1 * @stock,
                									  -1 * @stock1, 
                									  -1 * @stock2, 
                									  -1 * @stock3, 
                									  -1 * @stock4, 
                									  -1 * @stock5, 
                									  -1 * @stock6, 
                									  -1 * @stock7, 
                									  -1 * @stock8, 
                									  -1 * @stock9, 
                									  -1 * @stock10)
                		ELSE
                			UPDATE custom_vianny.DBO.Cag1700_Stock SET Stock_st = Stock_st - @stock,
                									 Stock1_st = Stock1_st - @stock1,
                									 Stock2_st = Stock2_st - @stock2,
                									 Stock3_st = Stock3_st - @stock3,
                									 Stock4_st = Stock4_st - @stock4,
                									 Stock5_st = Stock5_st - @stock5,
                									 Stock6_st = Stock6_st - @stock6,
                									 Stock7_st = Stock7_st - @stock7,
                									 Stock8_st = Stock8_st - @stock8,
                									 Stock9_st = Stock9_st - @stock9,
                									 Stock10_st = Stock10_st - @stock10
                							     WHERE CCia_st = @CCia_la AND Almac_st = @Almac_la AND Codigo_st = @Codigo_la

                		UPDATE custom_vianny.DBO.Cag1700 SET Stock_17 = Stock_17 - @stock
                							   WHERE CCia = @CCia_la AND Linea_17 = @Codigo_la
                END", conx)
            cmd19.Parameters.AddWithValue("@Codigo_la", Trim(DataGridView1.Rows(i).Cells(2).Value))
            cmd19.Parameters.AddWithValue("@ccom", "2")
            cmd19.Parameters.AddWithValue("@CCia_la", Trim(Label13.Text))
            cmd19.Parameters.AddWithValue("@Almac_la", Trim(ComboBox3.Text))
            cmd19.Parameters.AddWithValue("@stock", DataGridView1.Rows(i).Cells(14).Value)
            cmd19.Parameters.AddWithValue("@stock1", DataGridView1.Rows(i).Cells(4).Value)
            cmd19.Parameters.AddWithValue("@stock2", DataGridView1.Rows(i).Cells(5).Value)
            cmd19.Parameters.AddWithValue("@stock3", DataGridView1.Rows(i).Cells(6).Value)
            cmd19.Parameters.AddWithValue("@stock4", DataGridView1.Rows(i).Cells(7).Value)
            cmd19.Parameters.AddWithValue("@stock5", DataGridView1.Rows(i).Cells(8).Value)
            cmd19.Parameters.AddWithValue("@stock6", DataGridView1.Rows(i).Cells(9).Value)
            cmd19.Parameters.AddWithValue("@stock7", DataGridView1.Rows(i).Cells(10).Value)
            cmd19.Parameters.AddWithValue("@stock8", DataGridView1.Rows(i).Cells(11).Value)
            cmd19.Parameters.AddWithValue("@stock9", DataGridView1.Rows(i).Cells(12).Value)
            cmd19.Parameters.AddWithValue("@stock10", DataGridView1.Rows(i).Cells(13).Value)
            cmd19.ExecuteNonQuery()


            For y = 4 To 13
                If DataGridView1.Rows(i).Cells(y).Value > 0 Then
                    Dim cmd16 As New SqlCommand("INSERT INTO custom_vianny.DBO.MAP030F (CCIA_3A , CPERIODO_3A, talm_3a, ccom_3a, ncom_3a, item_3a,linea_3a ,sinon_3a,linea2_3a, cant_3a, unid_3a, cantu_3a, obser1_3a, obser2_3a, obser3_3a, cantk_3a, cantke_3a, pedido_3a, flag_3a, peso_3a,tpeso_3a,unidk_3a , pieza_3a, color_3a, ncolor_3a, parti_3a, ocorte_3a, ordens_3a, nrollo_3a, kgrollo_3a, lote_3a, maquina_3a, galga_3a,sigla_3a, sitcal_3a, cantc_3a, ancho_3a, densid_3a, pedod_3a, ordtej_3a  , tipapt_3a , combo_3a,ccosto_3a,PreUn1_3a,Tot1_3a  )
                           Values (@CCIA_3A,@CPERIODO_3A,@talm_3a,@ccom_3a,@ncom_3aa,@item_3a,@linea_3a, ''     ,'PT0101' ,@cantid ,@unid_3a,  @cantid,''        , ''       ,''        ,@cantid  , 0.00     ,@pedido_3a, 1      , 0.00000, 0.00   ,@unidk_3a,   ''    , ''      , ''       , ''      , ''       , ''       , ''       , 0.00      , ''     , ''        , 0.00    ,''      , 0        , 0.000   , 0.000   , 0.000    , ''      , ''         , ''        , ''      ,''       ,0.00000  , 0.00)", conx)
                    cmd16.Parameters.AddWithValue("@CCIA_3A", Trim(Label13.Text))
                    cmd16.Parameters.AddWithValue("@CPERIODO_3A", Trim(Label11.Text))
                    cmd16.Parameters.AddWithValue("@talm_3a", Trim(ComboBox3.Text))
                    cmd16.Parameters.AddWithValue("@ccom_3a", "2")
                    cmd16.Parameters.AddWithValue("@ncom_3aa", Trim(TextBox4.Text) & Trim(TextBox5.Text))
                    cmd16.Parameters.AddWithValue("@item_3a", Trim(DataGridView1.Rows(i).Cells(0).Value))
                    cmd16.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
                    cmd16.Parameters.AddWithValue("@cantid", DataGridView1.Rows(i).Cells(y).Value)
                    cmd16.Parameters.AddWithValue("@unid_3a", "UND")
                    cmd16.Parameters.AddWithValue("@pedido_3a", Trim(DataGridView1.Rows(i).Cells(1).Value))
                    cmd16.Parameters.AddWithValue("@unidk_3a", "UND")
                    cmd16.ExecuteNonQuery()
                End If

            Next
        Next

        MsgBox("SE REGISTRO LA INFORMACION DE LA NSA CORRECTAMENTE")


        'DataGridView11.DataSource = Nothing
        'CheckBox1.Checked = False
        'TextBox22.Enabled = False
        'BLOQUEARC()
        'limpiar()
        'correlativo()
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If CheckBox1.Checked = True Then

        Else
            nia()
        End If

    End Sub
    Private Sub BLOQUEARC()
        TextBox13.Enabled = False
        DataGridView1.Enabled = False
        DataGridView1.Columns(7).ReadOnly = False

        TextBox9.Enabled = False
        DateTimePicker1.Enabled = False
        'TextBox12.Enabled = True
        TextBox12.ReadOnly = False

        TextBox20.Enabled = False

        Button2.Enabled = False
        Button5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox4.Enabled = False
        TextBox16.Enabled = False
        TextBox26.Enabled = False
        TextBox17.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Button1.Enabled = False
        TextBox8.Enabled = False
    End Sub
    Public Sub limpiar()
        TextBox8.Text = ""
        TextBox10.Text = ""
        TextBox9.Text = ""
        TextBox11.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox26.Text = ""
        TextBox23.Text = ""
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        pTERMINADO.TextBox1.Text = Label13.Text
        pTERMINADO.TextBox2.Text = Label11.Text
        pTERMINADO.TextBox3.Text = ComboBox3.Text
        pTERMINADO.TextBox4.Text = "1"
        pTERMINADO.TextBox5.Text = Trim(TextBox4.Text) & Trim(TextBox5.Text)
        pTERMINADO.Show()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub
    Sub correlativo()
        Dim JU As Integer
        Select Case TextBox4.Text.Length

            Case "1" : TextBox4.Text = "0" & "" & TextBox4.Text
            Case "2" : TextBox4.Text = TextBox4.Text
        End Select
        'Dim dtFecha As Date = DateSerial(Month(Date.Today), 1, 1)

        Dim ano As String
        ano = Convert.ToString(Year(DateTimePicker1.Value))
        'DateTimePicker1.Value = "01/" + TextBox4.Text + "/" + MDIParent1.Label5.Text
        abrir()
        TextBox5.Enabled = True
        'MsgBox("select top 1 ncom_3, substring(ncom_3,3,7) as ncom from custom_vianny.dbo.mag030f  where ccia_3 = '01' and cperiodo_3 =  YEAR(GETDATE()) and talm_3 = " + ComboBox3.Text + " " + "and ncom_3 like" + " " + TextBox4.Text + "% and ccom_3 = '2' order by ncom_3 desc")
        enunciado = New SqlCommand("select top 1 ncom_3, substring(ncom_3,3,7) as ncom from custom_vianny.dbo.mag030f  where ccia_3 =" + Label13.Text + " and cperiodo_3 = '" + Label11.Text + "' and talm_3 = '" + ComboBox3.Text + "' " + "and ncom_3 like" + " '" + TextBox4.Text + "%' and ccom_3 = '1' order by ncom_3 desc", conx)
        respuesta = enunciado.ExecuteReader
        While respuesta.Read
            JU = respuesta.Item("ncom")
        End While
        respuesta.Close()
        TextBox5.Text = JU + 1
        Select Case TextBox5.Text.Length

            Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
            Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
            Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
            Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
            Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
            Case "6" : TextBox5.Text = TextBox5.Text
        End Select
        TextBox5.Select()
        TextBox5.ReadOnly = False
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            correlativo()
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
            I18 = DataGridView1.Rows.Count
            For i1 = 0 To I18 - 1
                VAL = i1 + 1
                Select Case VAL.ToString.Length
                    Case "1" : DataGridView1.Rows(i1).Cells(0).Value = "00" & "" & VAL
                    Case "2" : DataGridView1.Rows(i1).Cells(0).Value = "0" & "" & VAL
                    Case "3" : DataGridView1.Rows(i1).Cells(0).Value = VAL
                End Select
            Next
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        Button2.Enabled = False
        TextBox22.Text = "1"
        DataGridView1.Enabled = False
        DataGridView1.Rows.Clear()
        Button7.Enabled = False
        Button5.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        RadioButton1.Checked = True
        limpiar()
        TextBox8.Enabled = False
        TextBox5.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox1.ReadOnly = False
        TextBox5.Enabled = True
        TextBox4.Enabled = True
        ComboBox3.Enabled = True
        TextBox5.Text = ""
        TextBox4.Select()
        ty2.Clear()
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add()
        Dim I18, VAL As Integer
        I18 = DataGridView1.Rows.Count
        For i1 = 0 To I18 - 1
            VAL = i1 + 1
            Select Case VAL.ToString.Length
                Case "1" : DataGridView1.Rows(i1).Cells(0).Value = "00" & "" & VAL
                Case "2" : DataGridView1.Rows(i1).Cells(0).Value = "0" & "" & VAL
                Case "3" : DataGridView1.Rows(i1).Cells(0).Value = VAL
            End Select
        Next
    End Sub
    Dim Rsr300 As SqlDataReader
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If RadioButton2.Checked = True Then
            MsgBox("LA NOTA DE INGRESO ESTA ANULADA NO SE PUEDE EDITAR")
        Else

            Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label13.Text + "' AND A.talm_3 ='" + ComboBox3.Text + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
            Dim cmd102213 As New SqlCommand(sql102213, conx)
            Rsr300 = cmd102213.ExecuteReader()
            Dim jo As Integer
            If Rsr300.Read() Then
                jo = Rsr300(0)
            Else
                jo = 0
            End If
            Rsr300.Close()
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("QUIERES EDITAR ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                If jo = 0 Then
                    TextBox13.Enabled = True
                    DataGridView1.Enabled = True
                    DataGridView1.Columns(7).ReadOnly = False

                    TextBox9.Enabled = True
                    DateTimePicker1.Enabled = True
                    'TextBox12.Enabled = True
                    TextBox12.ReadOnly = False
                    TextBox4.Enabled = False
                    TextBox20.Enabled = True
                    TextBox5.Enabled = False
                    Button2.Enabled = True
                    Button5.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox4.Enabled = True
                    TextBox16.Enabled = True
                    TextBox26.Enabled = True
                    TextBox17.Enabled = True
                    Button3.Enabled = True
                    Button1.Enabled = True
                    TextBox8.Enabled = True
                    TextBox22.Text = "2"
                Else
                    MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
                End If

            End If
        End If
    End Sub

    Private Sub Nia_Pt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
        TextBox22.Text = "1"
    End Sub

    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs) Handles TextBox26.TextChanged

    End Sub

    Private Sub TextBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox26.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim VAL As Integer
            Select Case Trim(TextBox26.Text).Length
                Case "1" : TextBox26.Text = "01" & "0000000" & TextBox26.Text
                Case "2" : TextBox26.Text = "01" & "000000" & TextBox26.Text
                Case "3" : TextBox26.Text = "01" & "00000" & TextBox26.Text
                Case "4" : TextBox26.Text = "01" & "0000" & TextBox26.Text
                Case "5" : TextBox26.Text = "01" & "000" & TextBox26.Text
                Case "6" : TextBox26.Text = "01" & "00" & TextBox26.Text
                Case "7" : TextBox26.Text = "01" & "0" & TextBox26.Text
                Case "8" : TextBox26.Text = "01" & TextBox26.Text

            End Select

            Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, q.tallador_3 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(TextBox26.Text) + "' and q.ccia ='" + Trim(Label13.Text) + "'
"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            Dim i5 As Integer
            i5 = DataGridView1.Rows.Count

            If Rsr2.Read() = True Then
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i5).Cells(1).Value = Rsr2(0)
                DataGridView1.Rows(i5).Cells(2).Value = Rsr2(1)
                DataGridView1.Rows(i5).Cells(3).Value = Rsr2(2)
                DataGridView1.Rows(i5).Cells(4).Value = 0
                DataGridView1.Rows(i5).Cells(5).Value = 0
                DataGridView1.Rows(i5).Cells(6).Value = 0
                DataGridView1.Rows(i5).Cells(7).Value = 0
                DataGridView1.Rows(i5).Cells(8).Value = 0
                DataGridView1.Rows(i5).Cells(9).Value = 0
                DataGridView1.Rows(i5).Cells(10).Value = 0
                DataGridView1.Rows(i5).Cells(11).Value = 0
                DataGridView1.Rows(i5).Cells(12).Value = 0
                DataGridView1.Rows(i5).Cells(13).Value = 0
                DataGridView1.Rows(i5).Cells(14).Value = 0
                DataGridView1.Rows(i5).Cells(15).Value = Rsr2(3).ToString().Trim()
                DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                VAL = DataGridView1.Rows(i5).Cells(0).Value
                Select Case VAL.ToString.Length
                    Case "1" : DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                    Case "2" : DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                    Case "3" : DataGridView1.Rows(i5).Cells(0).Value = VAL
                End Select


                'DataGridView1.Rows(i5).Cells(4).Selected = True
                DataGridView1.CurrentCell = DataGridView1.Rows(i5).Cells(4)
                DataGridView1.BeginEdit(True)
                Rsr2.Close()

                Dim sql1 As String = "         select  cast( ISNULL(SUM(cant_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_Total,
                 cast( isnull(SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1,
                 cast( isnull(SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_2,
                 cast( isnull(SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_3,
                 cast( isnull(SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_4,
                 cast( isnull(SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_5,
                 cast( isnull(SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_6,
                 cast( isnull(SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_7,
                 cast( isnull(SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_8,
                 cast( isnull(SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_9,
                 cast( isnull(SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_10
                from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label13.Text) + "'  and talm_3b ='" + Trim(ComboBox3.Text) + "' and pedido_3b ='" + Trim(DataGridView1.Rows(i5).Cells(1).Value) + "' and flag_3b ='1'  and cperiodo_3b ='" + Trim(Label11.Text) + "' "
                Dim cmd1 As New SqlCommand(sql1, conx)
                Rsr20 = cmd1.ExecuteReader()
                If Rsr20.Read() = True Then
                    TextBox11.Text = Rsr20(1)
                    TextBox13.Text = Rsr20(2)
                    TextBox14.Text = Rsr20(3)
                    TextBox15.Text = Rsr20(4)
                    TextBox16.Text = Rsr20(5)
                    TextBox17.Text = Rsr20(6)
                    TextBox18.Text = Rsr20(7)
                    TextBox19.Text = Rsr20(8)
                    TextBox20.Text = Rsr20(9)
                    TextBox21.Text = Rsr20(10)

                Else
                    TextBox11.Text = 0
                    TextBox13.Text = 0
                    TextBox14.Text = 0
                    TextBox15.Text = 0
                    TextBox16.Text = 0
                    TextBox17.Text = 0
                    TextBox18.Text = 0
                    TextBox19.Text = 0
                    TextBox20.Text = 0
                    TextBox21.Text = 0

                End If
                Rsr20.Close()

                Dim sql10235 As String = "SELECT A.*,
                		  ISNULL(B.Dele,'') AS NTalla1,
                		  ISNULL(C.Dele,'') AS NTalla2,
                		  ISNULL(D.Dele,'') AS NTalla3,
                		  ISNULL(E.Dele,'') AS NTalla4,
                		  ISNULL(F.Dele,'') AS NTalla5,
                		  ISNULL(G.Dele,'') AS NTalla6,
                		  ISNULL(H.Dele,'') AS NTalla7,
                		  ISNULL(I.Dele,'') AS NTalla8,
                		  ISNULL(J.Dele,'') AS NTalla9,
                		  ISNULL(K.Dele,'') AS NTalla10
                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
                 Where A.CCia_tl = '" + Trim(Label13.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(i5).Cells(15).Value) + "'
                ORDER BY  A.CCia_tl, A.Codigo_tl"
                Dim cmd10235 As New SqlCommand(sql10235, conx)
                Rsr205 = cmd10235.ExecuteReader()
                If Rsr205.Read() = True Then
                    DataGridView1.Columns(4).HeaderText = Trim(Rsr205(13).ToString())
                    DataGridView1.Columns(5).HeaderText = Trim(Rsr205(14).ToString())
                    DataGridView1.Columns(6).HeaderText = Trim(Rsr205(15).ToString())
                    DataGridView1.Columns(7).HeaderText = Trim(Rsr205(16).ToString())
                    DataGridView1.Columns(8).HeaderText = Trim(Rsr205(17).ToString())
                    DataGridView1.Columns(9).HeaderText = Trim(Rsr205(18).ToString())
                    DataGridView1.Columns(10).HeaderText = Trim(Rsr205(19).ToString())
                    DataGridView1.Columns(11).HeaderText = Trim(Rsr205(20).ToString())
                    DataGridView1.Columns(12).HeaderText = Trim(Rsr205(21).ToString())
                    DataGridView1.Columns(13).HeaderText = Trim(Rsr205(22).ToString())




                End If
                Rsr205.Close()
                TextBox26.Text = ""
                If Trim(DataGridView1.Columns(4).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(4).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(5).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(5).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(6).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(6).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(7).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(7).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(8).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(8).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(9).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(9).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(10).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(10).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(11).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(11).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(12).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(12).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(13).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(13).ReadOnly = True
                End If
            Else
                MsgBox("LA OP INGRESADA NO EXISTE")
            End If
            Rsr2.Close()
            TextBox26.Select()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        MsgBox("ESTA FUNCION NO ESTA HABILITADA")


    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        DataGridView1.BeginEdit(True)
    End Sub

    Dim dt1, dt2, HG As New DataTable
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox26.Enabled = True
            TextBox26.Select()

            ty2.Clear()
            abrir()
            llenar_combo_box2(ComboBox1, ComboBox3.Text)
            Dim ml As New vnia
            Dim mk As New fnia
            Dim i As Integer
            i = Len(TextBox5.Text)

            Select Case i
                Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                Case "6" : TextBox5.Text = TextBox5.Text

            End Select
            ml.gcomprobante = TextBox4.Text & TextBox5.Text

            ml.galmacen1 = ComboBox3.Text
            ml.gano = Label11.Text
            ml.gccia = Label13.Text
            If mk.mostrar_correlativo_nia1(ml) = True Then
                Dim jk As New fnia

                Button4.Enabled = True
                Button6.Enabled = True

                DataGridView1.Rows.Clear()
                dt1 = jk.mostrar_cabecera_nia(ml)
                dt2 = jk.mostrar_detalle_nia_pt(ml)
                If dt1.Rows.Count <> 0 Then
                    DataGridView3.DataSource = dt1

                    TextBox9.Text = DataGridView3.Rows(0).Cells(0).Value
                    TextBox8.Text = DataGridView3.Rows(0).Cells(4).Value
                    If Convert.ToString(DataGridView3.Rows(0).Cells(6).Value) Is DBNull.Value Then
                        TextBox10.Text = ""
                    Else
                        TextBox10.Text = Convert.ToString(DataGridView3.Rows(0).Cells(6).Value)
                    End If

                    DateTimePicker1.Value = DataGridView3.Rows(0).Cells(1).Value
                    TextBox1.Text = DataGridView3.Rows(0).Cells(2).Value

                    abrir()
                    enunciado2 = New SqlCommand("select rtrim(ltrim(nomb_21e)) as nomb_21e from custom_vianny.dbo.cag210e where  Empr_21e ='" + Trim(Label13.Text) + "' AND almac_21e='" + Trim(ComboBox3.Text) + "'  and dest_21e='" + Trim(DataGridView3.Rows(0).Cells(5).Value) + "'", conx)
                    respuesta2 = enunciado2.ExecuteReader
                    While respuesta2.Read
                        ComboBox1.Text = respuesta2.Item("nomb_21e")
                    End While
                    respuesta2.Close()

                    If DataGridView3.Rows(0).Cells(8).Value = "EXT" Then
                        ComboBox2.SelectedIndex = 1
                    Else
                        ComboBox2.SelectedIndex = 0
                    End If

                End If
                If dt2.Rows.Count <> 0 Then
                    Dim nu1 As Integer
                    DataGridView4.DataSource = dt2
                    nu1 = DataGridView4.Rows.Count
                    DataGridView1.Rows.Add(nu1)
                    For i1 = 0 To nu1 - 1

                        DataGridView1.Rows(i1).Cells(0).Value = DataGridView4.Rows(i1).Cells(5).Value
                        DataGridView1.Rows(i1).Cells(1).Value = DataGridView4.Rows(i1).Cells(8).Value
                        DataGridView1.Rows(i1).Cells(2).Value = DataGridView4.Rows(i1).Cells(6).Value

                        DataGridView1.Rows(i1).Cells(4).Value = DataGridView4.Rows(i1).Cells(10).Value
                        DataGridView1.Rows(i1).Cells(5).Value = DataGridView4.Rows(i1).Cells(11).Value
                        DataGridView1.Rows(i1).Cells(6).Value = DataGridView4.Rows(i1).Cells(12).Value
                        DataGridView1.Rows(i1).Cells(7).Value = DataGridView4.Rows(i1).Cells(13).Value
                        DataGridView1.Rows(i1).Cells(8).Value = DataGridView4.Rows(i1).Cells(14).Value
                        DataGridView1.Rows(i1).Cells(9).Value = DataGridView4.Rows(i1).Cells(15).Value
                        DataGridView1.Rows(i1).Cells(10).Value = DataGridView4.Rows(i1).Cells(16).Value
                        DataGridView1.Rows(i1).Cells(11).Value = DataGridView4.Rows(i1).Cells(17).Value
                        DataGridView1.Rows(i1).Cells(12).Value = DataGridView4.Rows(i1).Cells(18).Value
                        DataGridView1.Rows(i1).Cells(13).Value = DataGridView4.Rows(i1).Cells(19).Value
                        DataGridView1.Rows(i1).Cells(14).Value = DataGridView4.Rows(i1).Cells(9).Value

                        Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, q.tallador_3 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(DataGridView1.Rows(i1).Cells(1).Value) + "' and q.ccia ='" + Trim(Label13.Text) + "'
"
                        Dim cmd102 As New SqlCommand(sql102, conx)
                        Rsr2 = cmd102.ExecuteReader()


                        If Rsr2.Read() = True Then
                            DataGridView1.Rows(i1).Cells(3).Value = Rsr2(2)
                            DataGridView1.Rows(i1).Cells(15).Value = Rsr2(3)
                        End If


                        Rsr2.Close()


                    Next


                    DataGridView1.Columns(0).Frozen = True
                    DataGridView1.Columns(1).Frozen = True
                    DataGridView1.Columns(2).Frozen = True

                End If

                Dim kl, sum12 As Integer

                kl = DataGridView1.Rows.Count

                For ol = 0 To kl - 1
                    sum12 = sum12 + Convert.ToInt32(DataGridView1.Rows(ol).Cells(14).Value)
                Next
                TextBox23.Text = sum12
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                ComboBox3.Enabled = False
                ComboBox4.Enabled = False
                TextBox8.Enabled = False
                TextBox9.Enabled = False
                TextBox1.Enabled = False
                TextBox1.ReadOnly = False
                Button2.Enabled = False
                DataGridView1.Enabled = False
                Button1.Enabled = False
                Button5.Enabled = False
                Button7.Enabled = Enabled
                Button3.Enabled = False
                TextBox9.Enabled = False
                TextBox1.Enabled = False
                TextBox8.Enabled = False
                TextBox10.Enabled = False
                TextBox12.ReadOnly = False
            Else
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox3.Enabled = False
                ComboBox4.Enabled = True
                TextBox8.Enabled = True
                TextBox9.Enabled = True
                TextBox1.Enabled = True
                TextBox1.ReadOnly = False
                Button2.Enabled = True
                DataGridView1.Enabled = True
                Button1.Enabled = True
                Button5.Enabled = True
                Button3.Enabled = True
                DataGridView1.Rows.Clear()
                TextBox9.Text = ""
                TextBox1.Text = ""
                TextBox8.Text = ""
                TextBox10.Text = ""
                TextBox12.ReadOnly = False
                Button7.Enabled = Enabled
            End If
            ComboBox3.Enabled = False
            ComboBox4.SelectedIndex = -1
            DateTimePicker1.Select()

        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        If e.RowIndex >= 0 Then


            Dim sql1023 As String = " SELECT A.*,ISNULL(B.Nomb_17,'') AS Nomb_17,ISNULL(B.Unid_17,'') AS Unid_17,ISNULL(B.TipCtrl_17,'') AS TipCtrl_17,ISNULL(B.IdAlterno_17,'') AS IdAlterno_17,ISNULL(B.Peso_17,0) AS Peso_17,ISNULL(C.Nomb_15m,'') AS Nomb_15m,ISNULL(C.Sigla_15m,'') AS Sigla_15m
	   FROM  custom_vianny.dbo.Cag1700_Stock A FULL JOIN  custom_vianny.dbo.Cag1700 B ON A.CCia_st = B.CCia AND A.Codigo_st = B.Linea_17 FULL JOIN custom_vianny.dbo.Mag1500 C ON A.CCia_st = C.CCia AND A.Almac_st = C.TAlm_15m Where B.CCia = '" + Trim(Label13.Text) + "' AND B.Linea_17 = '" + Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value) + "' AND A.Almac_st ='" + Trim(ComboBox3.Text) + "' ORDER BY  A.ccia_st, A.Almac_st, A.codigo_st"
            Dim cmd1023 As New SqlCommand(sql1023, conx)
            Rsr22 = cmd1023.ExecuteReader()
            If Rsr22.Read() = True Then
                TextBox11.Text = Rsr22(4)
                TextBox13.Text = Rsr22(5)
                TextBox14.Text = Rsr22(6)
                TextBox15.Text = Rsr22(7)
                TextBox16.Text = Rsr22(8)
                TextBox17.Text = Rsr22(9)
                TextBox18.Text = Rsr22(10)
                TextBox19.Text = Rsr22(11)
                TextBox20.Text = Rsr22(12)
                TextBox21.Text = Rsr22(13)

            Else
                TextBox11.Text = 0
                TextBox13.Text = 0
                TextBox14.Text = 0
                TextBox15.Text = 0
                TextBox16.Text = 0
                TextBox17.Text = 0
                TextBox18.Text = 0
                TextBox19.Text = 0
                TextBox20.Text = 0
                TextBox21.Text = 0

            End If
            Rsr22.Close()

            Dim sql10235 As String = "SELECT A.*,
                		  ISNULL(B.Dele,'') AS NTalla1,
                		  ISNULL(C.Dele,'') AS NTalla2,
                		  ISNULL(D.Dele,'') AS NTalla3,
                		  ISNULL(E.Dele,'') AS NTalla4,
                		  ISNULL(F.Dele,'') AS NTalla5,
                		  ISNULL(G.Dele,'') AS NTalla6,
                		  ISNULL(H.Dele,'') AS NTalla7,
                		  ISNULL(I.Dele,'') AS NTalla8,
                		  ISNULL(J.Dele,'') AS NTalla9,
                		  ISNULL(K.Dele,'') AS NTalla10
                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
                 Where A.CCia_tl = '" + Trim(Label13.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(e.RowIndex).Cells(15).Value) + "'
                ORDER BY  A.CCia_tl, A.Codigo_tl"
            Dim cmd10235 As New SqlCommand(sql10235, conx)
            Rsr23 = cmd10235.ExecuteReader()
            If Rsr23.Read() = True Then
                DataGridView1.Columns(4).HeaderText = Trim(Rsr23(13))
                DataGridView1.Columns(5).HeaderText = Trim(Rsr23(14))
                DataGridView1.Columns(6).HeaderText = Trim(Rsr23(15))
                DataGridView1.Columns(7).HeaderText = Trim(Rsr23(16))
                DataGridView1.Columns(8).HeaderText = Trim(Rsr23(17))
                DataGridView1.Columns(9).HeaderText = Trim(Rsr23(18))
                DataGridView1.Columns(10).HeaderText = Trim(Rsr23(19))
                DataGridView1.Columns(11).HeaderText = Trim(Rsr23(20))
                DataGridView1.Columns(12).HeaderText = Trim(Rsr23(21))
                DataGridView1.Columns(13).HeaderText = Trim(Rsr23(22))
            End If
            Rsr23.Close()
        End If
    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "36080"
            Clientes.Show()
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = 4 Then
            Dim CANT1, CANT2, SUMA4, T1 As Integer
            SUMA4 = 0

            If Trim(DataGridView1.Columns(4).HeaderText).Length = 0 Then

            Else
                For i4 = 4 To 13
                    SUMA4 = SUMA4 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA4
                T1 = DataGridView1.Rows.Count
                For E4 = 0 To T1 - 1
                    CANT2 = CANT2 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT2
            End If



        End If
        If e.ColumnIndex = 5 Then
            Dim CANT15, CANT25, SUMA5, T2 As Integer


            If Trim(DataGridView1.Columns(5).HeaderText).Length = 0 Then

            Else
                For i4 = 4 To 13
                    SUMA5 = SUMA5 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA5
                T2 = DataGridView1.Rows.Count
                For E4 = 0 To T2 - 1
                    CANT25 = CANT25 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT25
            End If


        End If
        If e.ColumnIndex = 6 Then
            Dim CANT16, CANT26, SUMA6, T3 As Integer


            If Trim(DataGridView1.Columns(6).HeaderText).Length = 0 Then

            Else
                For i4 = 4 To 13
                    SUMA6 = SUMA6 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA6

                T3 = DataGridView1.Rows.Count
                For E4 = 0 To T3 - 1
                    CANT26 = CANT26 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT26
            End If


        End If
        If e.ColumnIndex = 7 Then
            Dim CANT17, CANT27, SUMA7, T4 As Integer


            If Trim(DataGridView1.Columns(7).HeaderText).Length = 0 Then

            Else
                For i4 = 4 To 13
                    SUMA7 = SUMA7 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA7
                T4 = DataGridView1.Rows.Count
                For E4 = 0 To T4 - 1
                    CANT27 = CANT27 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT27
            End If



        End If
        If e.ColumnIndex = 8 Then
            Dim CANT18, CANT28, SUMA8, T5 As Integer


            If Trim(DataGridView1.Columns(8).HeaderText).Length = 0 Then

            Else
                For i4 = 4 To 13
                    SUMA8 = SUMA8 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA8
                T5 = DataGridView1.Rows.Count
                For E4 = 0 To T5 - 1
                    CANT28 = CANT28 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT28
            End If



        End If
        If e.ColumnIndex = 9 Then
            Dim CANT19, CANT29, SUMA9, T6 As Integer


            If Trim(DataGridView1.Columns(9).HeaderText).Length = 0 Then

            Else
                For i4 = 4 To 13
                    SUMA9 = SUMA9 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA9

                T6 = DataGridView1.Rows.Count
                For E4 = 0 To T6 - 1
                    CANT29 = CANT29 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT29
            End If


        End If
        If e.ColumnIndex = "10" Then
            Dim CANT110, CANT210, SUMA10, T7 As Integer


            If Trim(DataGridView1.Columns(10).HeaderText).Length = 0 Then

            Else
                For i4 = 4 To 13
                    SUMA10 = SUMA10 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA10
                T7 = DataGridView1.Rows.Count
                For E4 = 0 To T7 - 1
                    CANT210 = CANT210 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT210
            End If



        End If
        If e.ColumnIndex = "11" Then
            Dim CANT111, CANT211, SUMA11, T8 As Integer


            If Trim(DataGridView1.Columns(11).HeaderText).Length = 0 Then

            Else
                For i4 = 4 To 13
                    SUMA11 = SUMA11 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA11
                T8 = DataGridView1.Rows.Count
                For E4 = 0 To T8 - 1
                    CANT211 = CANT211 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT211
            End If



        End If
        If e.ColumnIndex = "12" Then
            Dim CANT112, CANT212, SUMA12, T9 As Integer


            If Trim(DataGridView1.Columns(12).HeaderText).Length = 0 Then

            Else

                For i4 = 4 To 13
                    SUMA12 = SUMA12 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA12
                T9 = DataGridView1.Rows.Count
                For E4 = 0 To T9 - 1
                    CANT212 = CANT212 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT212
            End If


        End If
        If e.ColumnIndex = "13" Then
            Dim CANT113, CANT213, SUMA13, T10 As Integer



            If Trim(DataGridView1.Columns(13).HeaderText).Length = 0 Then

            Else
                For i4 = 4 To 13
                    SUMA13 = SUMA13 + CInt(DataGridView1.Rows(e.RowIndex).Cells(i4).Value)
                Next
                DataGridView1.Rows(e.RowIndex).Cells(14).Value = SUMA13
                T10 = DataGridView1.Rows.Count
                For E4 = 0 To T10 - 1
                    CANT213 = CANT213 + CInt(DataGridView1.Rows(E4).Cells(14).Value)
                Next
                TextBox23.Text = CANT213
            End If



        End If
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode = Keys.Up Then
            Dim sql1023 As String = " SELECT A.*,ISNULL(B.Nomb_17,'') AS Nomb_17,ISNULL(B.Unid_17,'') AS Unid_17,ISNULL(B.TipCtrl_17,'') AS TipCtrl_17,ISNULL(B.IdAlterno_17,'') AS IdAlterno_17,ISNULL(B.Peso_17,0) AS Peso_17,ISNULL(C.Nomb_15m,'') AS Nomb_15m,ISNULL(C.Sigla_15m,'') AS Sigla_15m
	   FROM  custom_vianny.dbo.Cag1700_Stock A FULL JOIN  custom_vianny.dbo.Cag1700 B ON A.CCia_st = B.CCia AND A.Codigo_st = B.Linea_17 FULL JOIN custom_vianny.dbo.Mag1500 C ON A.CCia_st = C.CCia AND A.Almac_st = C.TAlm_15m Where B.CCia = '" + Trim(Label13.Text) + "' AND B.Linea_17 = '" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(2).Value) + "' AND A.Almac_st ='" + Trim(ComboBox3.Text) + "' ORDER BY  A.ccia_st, A.Almac_st, A.codigo_st"
            Dim cmd1023 As New SqlCommand(sql1023, conx)
            Rsr22 = cmd1023.ExecuteReader()
            If Rsr22.Read() = True Then
                TextBox11.Text = Rsr22(4)
                TextBox13.Text = Rsr22(5)
                TextBox14.Text = Rsr22(6)
                TextBox15.Text = Rsr22(7)
                TextBox16.Text = Rsr22(8)
                TextBox17.Text = Rsr22(9)
                TextBox18.Text = Rsr22(10)
                TextBox19.Text = Rsr22(11)
                TextBox20.Text = Rsr22(12)
                TextBox21.Text = Rsr22(13)

            Else
                TextBox11.Text = 0
                TextBox13.Text = 0
                TextBox14.Text = 0
                TextBox15.Text = 0
                TextBox16.Text = 0
                TextBox17.Text = 0
                TextBox18.Text = 0
                TextBox19.Text = 0
                TextBox20.Text = 0
                TextBox21.Text = 0

            End If
            Rsr22.Close()

            Dim sql10235 As String = "SELECT A.*,
                		  ISNULL(B.Dele,'') AS NTalla1,
                		  ISNULL(C.Dele,'') AS NTalla2,
                		  ISNULL(D.Dele,'') AS NTalla3,
                		  ISNULL(E.Dele,'') AS NTalla4,
                		  ISNULL(F.Dele,'') AS NTalla5,
                		  ISNULL(G.Dele,'') AS NTalla6,
                		  ISNULL(H.Dele,'') AS NTalla7,
                		  ISNULL(I.Dele,'') AS NTalla8,
                		  ISNULL(J.Dele,'') AS NTalla9,
                		  ISNULL(K.Dele,'') AS NTalla10
                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
                 Where A.CCia_tl = '" + Trim(Label13.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(15).Value) + "'
                ORDER BY  A.CCia_tl, A.Codigo_tl"
            Dim cmd10235 As New SqlCommand(sql10235, conx)
            Rsr23 = cmd10235.ExecuteReader()
            If Rsr23.Read() = True Then
                DataGridView1.Columns(4).HeaderText = Trim(Rsr23(13))
                DataGridView1.Columns(5).HeaderText = Trim(Rsr23(14))
                DataGridView1.Columns(6).HeaderText = Trim(Rsr23(15))
                DataGridView1.Columns(7).HeaderText = Trim(Rsr23(16))
                DataGridView1.Columns(8).HeaderText = Trim(Rsr23(17))
                DataGridView1.Columns(9).HeaderText = Trim(Rsr23(18))
                DataGridView1.Columns(10).HeaderText = Trim(Rsr23(19))
                DataGridView1.Columns(11).HeaderText = Trim(Rsr23(20))
                DataGridView1.Columns(12).HeaderText = Trim(Rsr23(21))
                DataGridView1.Columns(13).HeaderText = Trim(Rsr23(22))



            End If
            Rsr23.Close()

            'TextBox37.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(4).Value
            'TextBox36.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(5).Value
            'TextBox35.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(6).Value
            'TextBox34.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(7).Value
            'TextBox33.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(8).Value
            'TextBox32.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(9).Value
            'TextBox31.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(10).Value
            'TextBox30.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(11).Value
            'TextBox29.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(12).Value
            'TextBox28.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(13).Value

        End If

        If e.KeyCode = Keys.Down Then
            Dim sql1023 As String = " SELECT A.*,ISNULL(B.Nomb_17,'') AS Nomb_17,ISNULL(B.Unid_17,'') AS Unid_17,ISNULL(B.TipCtrl_17,'') AS TipCtrl_17,ISNULL(B.IdAlterno_17,'') AS IdAlterno_17,ISNULL(B.Peso_17,0) AS Peso_17,ISNULL(C.Nomb_15m,'') AS Nomb_15m,ISNULL(C.Sigla_15m,'') AS Sigla_15m
	   FROM  custom_vianny.dbo.Cag1700_Stock A FULL JOIN  custom_vianny.dbo.Cag1700 B ON A.CCia_st = B.CCia AND A.Codigo_st = B.Linea_17 FULL JOIN custom_vianny.dbo.Mag1500 C ON A.CCia_st = C.CCia AND A.Almac_st = C.TAlm_15m Where B.CCia = '" + Trim(Label13.Text) + "' AND B.Linea_17 = '" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(2).Value) + "' AND A.Almac_st ='" + Trim(ComboBox3.Text) + "' ORDER BY  A.ccia_st, A.Almac_st, A.codigo_st"
            Dim cmd1023 As New SqlCommand(sql1023, conx)
            Rsr22 = cmd1023.ExecuteReader()
            If Rsr22.Read() = True Then
                TextBox11.Text = Rsr22(4)
                TextBox13.Text = Rsr22(5)
                TextBox14.Text = Rsr22(6)
                TextBox15.Text = Rsr22(7)
                TextBox16.Text = Rsr22(8)
                TextBox17.Text = Rsr22(9)
                TextBox18.Text = Rsr22(10)
                TextBox19.Text = Rsr22(11)
                TextBox20.Text = Rsr22(12)
                TextBox21.Text = Rsr22(13)

            Else
                TextBox11.Text = 0
                TextBox13.Text = 0
                TextBox14.Text = 0
                TextBox15.Text = 0
                TextBox16.Text = 0
                TextBox17.Text = 0
                TextBox18.Text = 0
                TextBox19.Text = 0
                TextBox20.Text = 0
                TextBox21.Text = 0

            End If
            Rsr22.Close()

            Dim sql10235 As String = "SELECT A.*,
                		  ISNULL(B.Dele,'') AS NTalla1,
                		  ISNULL(C.Dele,'') AS NTalla2,
                		  ISNULL(D.Dele,'') AS NTalla3,
                		  ISNULL(E.Dele,'') AS NTalla4,
                		  ISNULL(F.Dele,'') AS NTalla5,
                		  ISNULL(G.Dele,'') AS NTalla6,
                		  ISNULL(H.Dele,'') AS NTalla7,
                		  ISNULL(I.Dele,'') AS NTalla8,
                		  ISNULL(J.Dele,'') AS NTalla9,
                		  ISNULL(K.Dele,'') AS NTalla10
                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
                 Where A.CCia_tl = '" + Trim(Label13.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(15).Value) + "'
                ORDER BY  A.CCia_tl, A.Codigo_tl"
            Dim cmd10235 As New SqlCommand(sql10235, conx)
            Rsr23 = cmd10235.ExecuteReader()
            If Rsr23.Read() = True Then
                DataGridView1.Columns(4).HeaderText = Trim(Rsr23(13))
                DataGridView1.Columns(5).HeaderText = Trim(Rsr23(14))
                DataGridView1.Columns(6).HeaderText = Trim(Rsr23(15))
                DataGridView1.Columns(7).HeaderText = Trim(Rsr23(16))
                DataGridView1.Columns(8).HeaderText = Trim(Rsr23(17))
                DataGridView1.Columns(9).HeaderText = Trim(Rsr23(18))
                DataGridView1.Columns(10).HeaderText = Trim(Rsr23(19))
                DataGridView1.Columns(11).HeaderText = Trim(Rsr23(20))
                DataGridView1.Columns(12).HeaderText = Trim(Rsr23(21))
                DataGridView1.Columns(13).HeaderText = Trim(Rsr23(22))



            End If
            Rsr23.Close()
            'TextBox37.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(4).Value
            'TextBox36.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(5).Value
            'TextBox35.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(6).Value
            'TextBox34.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(7).Value
            'TextBox33.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(8).Value
            'TextBox32.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(9).Value
            'TextBox31.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(10).Value
            'TextBox30.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(11).Value
            'TextBox29.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(12).Value
            'TextBox28.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(13).Value
        End If
    End Sub

End Class