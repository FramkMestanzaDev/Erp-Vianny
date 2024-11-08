Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.ComponentModel

Public Class Codificacion_Prodcutos
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable

    'Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    '    If TextBox7.Text.Length > 0 Then
    '        Dim h As String
    '        h = Mid(TextBox7.Text, 4, 3)
    '        TextBox7.Text = ""
    '        TextBox7.Text = ComboBox1.SelectedValue.ToString + h + "N" + TextBox1.Text + TextBox3.Text
    '    Else
    '        TextBox7.Text = ComboBox1.SelectedValue.ToString
    '    End If
    '    TextBox8.Text = Trim(ComboBox1.Text) + " " + Trim(ComboBox2.Text) + " " + Trim(TextBox6.Text) + " COLOR: " + Trim(TextBox2.Text) + " ANCHO: " + Trim(TextBox4.Text) + " DENSIDAD:" + Trim(TextBox5.Text)
    '    'ty3.Clear()
    '    'llenar_combo_box3(ComboBox2)
    'End Sub

    'Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
    '    Dim fc As New fcodigo
    '    Dim kk As New vlogica

    '    kk.gCCia = "01"
    '    kk.gFamilia = ComboBox1.SelectedValue.ToString
    '    kk.gSubFamilia = ComboBox2.SelectedValue.ToString
    '    kk.gOrigen = "N"
    '    kk.gColor = TextBox1.Text

    '    TextBox3.Text = fc.mostrar_correlativo(kk)
    '    If TextBox7.Text.Length > 0 Then
    '        Dim h As String
    '        h = Mid(TextBox7.Text, 1, 3)
    '        TextBox7.Text = ""
    '        TextBox7.Text = h + ComboBox2.SelectedValue.ToString + "N" + TextBox1.Text + TextBox3.Text
    '    Else
    '        Dim h As String
    '        h = ComboBox2.SelectedValue.ToString
    '        TextBox7.Text = TextBox7.Text + h + "N" + TextBox1.Text + TextBox3.Text
    '    End If
    '    TextBox8.Text = Trim(ComboBox1.Text) + " " + Trim(ComboBox2.Text) + " " + Trim(TextBox6.Text) + " COLOR: " + Trim(TextBox2.Text) + " ANCHO: " + Trim(TextBox4.Text) + " DENSIDAD:" + Trim(TextBox5.Text)

    'End Sub

    Dim dt1, dt2, HG As New DataTable

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ko As New vcodigo
        Dim ll As New fcodigo

        If Me.ValidateChildren() And TextBox1.Text = String.Empty Or TextBox2.Text = String.Empty Or TextBox8.Text = String.Empty Then
            MsgBox("FALTA INGRESAR CAMPOS OBLIGATORIOS")
        Else

            ko.gccia = Label13.Text
            ko.glinea = TextBox7.Text
            ko.gfamilia = TextBox13.Text.ToString().Trim()
            ko.gsubfamilia = TextBox14.Text.ToString().Trim()
            ko.gcodcolor = TextBox1.Text
            ko.gcolor = TextBox2.Text
            ko.gancho = ""
            ko.gdensidad = 0
            ko.gdetalle = TextBox6.Text.ToString().Trim()
            ko.gdescripcion = TextBox8.Text
            ko.gnombre_comercial = TextBox10.Text
            ko.gum = ComboBox3.Text
            If ComboBox4.Text = "STOCK" Then
                ko.gTipCtrl_17 = "1"
            Else
                If ComboBox4.Text = "STOCK + LOTE" Then
                    ko.gTipCtrl_17 = "2"
                Else
                    ko.gTipCtrl_17 = "3"
                End If
            End If
            ko.galmacen = Label19.Text
            If TextBox1.Text = "C00001" Then
                ko.gtipo = 3
            Else
                ko.gtipo = 1
            End If
            ko.gbarra = ""
            ko.grubro = TextBox15.Text.ToString().Trim()
            '1 ES TELA
            '2 ES SERVICIO
            '3 ES MUESTRA
            If ll.insertar_codigo(ko) Then
                Dim K As Integer
                K = DataGridView2.Rows.Count
                For I = 0 To K - 1
                    Dim cmd15 As New SqlCommand("insert into codigo_barra (codigo,barra) values (@codigo,@barra)", conx)
                    cmd15.Parameters.AddWithValue("@codigo", Trim(TextBox7.Text))
                    cmd15.Parameters.AddWithValue("@barra", Trim(DataGridView2.Rows(I).Cells(0).Value))
                    cmd15.ExecuteNonQuery()
                Next
                Label7.Text = "0"
                MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("Quieres crear la Ficha Tecnica?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    FICHA_TECNICA.Close()
                    FICHA_TECNICA.Label39.Text = Label19.Text
                    FICHA_TECNICA.Label41.Text = Label20.Text
                    FICHA_TECNICA.TextBox1.Text = TextBox7.Text.ToString().Trim()
                    FICHA_TECNICA.TextBox1.Focus()
                    SendKeys.Send("{ENTER}")
                    FICHA_TECNICA.ShowDialog()
                End If
                Dim hj2 As New flog
                    Dim hj1 As New vlog
                    hj1.gmodulo = "CODIGO"
                    hj1.gcuser = MDIParent2.Label6.Text
                    If Label7.Text = "0" Then
                        hj1.gaccion = 1
                    Else
                        hj1.gaccion = 2
                    End If

                    hj1.gpc = My.Computer.Name
                    hj1.gfecha = DateTimePicker1.Value
                    hj1.gdetalle = TextBox7.Text.ToString().Trim()
                    hj1.gccia = Label13.Text
                    hj2.insertar_log(hj1)
                    TabPage2.Select()
                    TextBox6.Text = ""
                    TextBox8.Text = ""
                    TextBox10.Text = ""
                    TextBox7.Text = ""
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox11.Text = ""
                    TextBox12.Text = ""
                    TextBox13.Text = ""
                    TextBox14.Text = ""
                    DataGridView2.Rows.Clear()

                Else

                End If

        End If

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)
        TextBox8.Text = TextBox13.Text.ToString().Trim() + " " + TextBox14.Text.ToString().Trim() + " " + Trim(TextBox6.Text)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs)
        TextBox8.Text = TextBox13.Text.ToString().Trim() + " " + TextBox14.Text.ToString().Trim() + " " + Trim(TextBox6.Text)
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        TextBox8.Text = Trim(TextBox6.Text)

    End Sub

    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            If Label14.Text = 1 Then
                ds.Tables.Add(da.Copy)
            Else
                ds.Tables.Add(da1.Copy)
            End If

            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CODIGO" & " like '%" & TextBox4.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(1).Width = 180
                DataGridView1.Columns(2).Width = 490
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim da As New DataTable
    Dim da1 As New DataTable
    Sub ACTUALIZAR()
        abrir()
        da1.Clear()
                DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("select ccia,linea_17 AS CODIGO,nomb_17 AS DESCRIPCION,nombs_17,detalle_17 AS DETALLE,densidad_17 AS DENSIDAD,idcolor_17,ncolor_17, familia_17,subfam_17,unid_17 from custom_vianny.dbo.cag1700  where ccia ='01' AND talm_17 ='" + Label19.Text + "'", conx)

        cmd.Fill(da1)
        bs.DataSource = da1
        DataGridView1.DataSource = bs
        DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(1).Width = 180
                DataGridView1.Columns(2).Width = 490
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False




    End Sub
    Private Sub Codificacion_Prodcutos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bloquear()

        abrir()
        'llenar_combo_box2(ComboBox1)
        'llenar_combo_box3(ComboBox2)
        llenar_combo_box(ComboBox3)
        'TextBox7.Text = ComboBox1.SelectedValue.ToString + ComboBox2.SelectedValue.ToString + Trim(TextBox3.Text)
        ACTUALIZAR()

        ComboBox4.SelectedIndex = 0


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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    'Sub llenar_combo_box2(ByVal cb As ComboBox)
    '    If Label14.Text = 1 Then
    '        Try

    '            conn = New SqlDataAdapter("select codigo_fam,descrip_fam from custom_vianny.dbo.familias where  talm_fam ='06' and ccia_fam =" + Label13.Text, conx)
    '            conn.Fill(ty2)
    '            ComboBox1.DataSource = ty2
    '            ComboBox1.DisplayMember = "descrip_fam"
    '            ComboBox1.ValueMember = "codigo_fam"

    '            'respuesta = enunciado.ExecuteReader
    '            'While respuesta.Read
    '            '    cb.Items.Add(respuesta.Item("nomb_17f"))
    '            'End While
    '            'respuesta.Close()
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '    Else
    '        If Label14.Text = 2 Then

    '            Try

    '                conn = New SqlDataAdapter("select codigo_fam,descrip_fam from custom_vianny.dbo.familias where  talm_fam ='69' and ccia_fam =" + Label13.Text, conx)
    '                conn.Fill(ty2)
    '                ComboBox1.DataSource = ty2
    '                ComboBox1.DisplayMember = "descrip_fam"
    '                ComboBox1.ValueMember = "codigo_fam"

    '                'respuesta = enunciado.ExecuteReader
    '                'While respuesta.Read
    '                '    cb.Items.Add(respuesta.Item("nomb_17f"))
    '                'End While
    '                'respuesta.Close()
    '            Catch ex As Exception
    '                MsgBox(ex.Message)
    '            End Try

    '        End If
    '    End If

    'End Sub
    'Sub llenar_combo_box3(ByVal cb As ComboBox)
    '    If Label14.Text = 1 Then
    '        Try

    '            conn = New SqlDataAdapter("select codigo_sfam,descrip_sfam from custom_vianny.dbo.subfamilias where  talm_sfam='06' and ccia_sfam =" + Label13.Text, conx)
    '            conn.Fill(ty3)
    '            ComboBox2.DataSource = ty3
    '            ComboBox2.DisplayMember = "descrip_sfam"
    '            ComboBox2.ValueMember = "codigo_sfam"

    '            'respuesta = enunciado.ExecuteReader
    '            'While respuesta.Read
    '            '    cb.Items.Add(respuesta.Item("nomb_17f"))
    '            'End While
    '            'respuesta.Close()
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '    Else
    '        If Label14.Text = 2 Then
    '            Try

    '                conn = New SqlDataAdapter("select codigo_sfam,descrip_sfam from custom_vianny.dbo.subfamilias where  talm_sfam='69' and ccia_sfam =" + Label13.Text, conx)
    '                conn.Fill(ty3)
    '                ComboBox2.DataSource = ty3
    '                ComboBox2.DisplayMember = "descrip_sfam"
    '                ComboBox2.ValueMember = "codigo_sfam"

    '                'respuesta = enunciado.ExecuteReader
    '                'While respuesta.Read
    '                '    cb.Items.Add(respuesta.Item("nomb_17f"))
    '                'End While
    '                'respuesta.Close()
    '            Catch ex As Exception
    '                MsgBox(ex.Message)
    '            End Try
    '        End If
    '    End If

    'End Sub

    'Private Sub ComboBox2_Click(sender As Object, e As EventArgs) Handles ComboBox2.Click
    '    TextBox4.ReadOnly = False
    '    TextBox5.ReadOnly = False
    '    TextBox6.ReadOnly = False
    '    TextBox10.Enabled = True
    'End Sub

    Sub llenar_combo_box(ByVal cb As ComboBox)

        Try

            conn = New SqlDataAdapter(" SELECT unid_16m as unidad, nomb_16m as nombre FROM custom_vianny.dbo.MAG1600", conx)
            conn.Fill(ty)
            ComboBox3.DataSource = ty
            ComboBox3.DisplayMember = "unidad"
            ComboBox3.ValueMember = "nombre"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Dim f As New Ficha
            f.Padre1 = Me
            f.TextBox2.Text = 4
            f.Label2.Text = Label13.Text
            f.Show(Me)
        End If
    End Sub
    Dim bs As New BindingSource()
    Dim filtroBase As String = ""
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        Dim textoBusqueda As String = TextBox9.Text.Trim()

        ' Si el texto contiene el símbolo %, dividimos en la parte fija y la dinámica
        If textoBusqueda.Contains("%") Then
            Dim partes As String() = textoBusqueda.Split(New Char() {"%"c}, StringSplitOptions.None)
            filtroBase = partes(0).Trim() ' La parte antes del %

            If partes.Length > 1 Then
                ' Filtrar utilizando la parte congelada y lo que venga después del %
                AplicarFiltro(filtroBase, partes(1).Trim())
            Else
                ' Si solo está el %, usar solo el filtro base
                AplicarFiltro(filtroBase, "")
            End If
        Else
            ' Si no se encuentra el %, simplemente aplicar el filtro en tiempo real
            AplicarFiltro(textoBusqueda, "")
        End If
    End Sub
    Private Sub AplicarFiltro(filtroFijo As String, filtroDinamico As String)
        ' Verifica si el DataSource es un BindingSource
        If TypeOf DataGridView1.DataSource Is BindingSource Then
            Dim bs As BindingSource = CType(DataGridView1.DataSource, BindingSource)
            Dim dt As DataTable = CType(bs.DataSource, DataTable)

            If dt IsNot Nothing Then
                ' Escapamos caracteres especiales si es necesario
                filtroFijo = filtroFijo.Replace("'", "''")
                filtroDinamico = filtroDinamico.Replace("'", "''")

                ' Aplicamos el filtro base y luego el filtro dinámico si existe
                If filtroDinamico <> "" Then
                    dt.DefaultView.RowFilter = String.Format("DESCRIPCION LIKE '%{0}%' AND DESCRIPCION LIKE '%{1}%'", filtroFijo, filtroDinamico)
                Else
                    dt.DefaultView.RowFilter = String.Format("DESCRIPCION LIKE '%{0}%'", filtroFijo)
                End If
            End If
        End If
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'Label17.Text = e.RowIndex


    End Sub

    Private Sub TextBox4_Validating(sender As Object, e As CancelEventArgs)
        'If DirectCast(sender, TextBox).Text.Length > 0 Then
        '    Me.ErrorProvider1.SetError(sender, "")
        'Else
        '    Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL ANCHO")
        'End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox6.ReadOnly = False
        Label7.Text = "1"
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.DataSource = ""
        TextBox4.Text = ""
        TextBox9.Text = ""
        ACTUALIZAR()
    End Sub

    Private Sub TextBox5_Validating(sender As Object, e As CancelEventArgs)
        'If DirectCast(sender, TextBox).Text.Length > 0 Then
        '    Me.ErrorProvider1.SetError(sender, "")
        'Else
        '    Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR LA DENSIDAD")
        'End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DataGridView2.Rows.Add()


        DataGridView2.Rows(DataGridView2.Rows.Count - 1).Selected = True
        DataGridView2.CurrentCell = DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(0)
        DataGridView2.SelectionMode = DataGridViewSelectionMode.CellSelect
        DataGridView2.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
        DataGridView2.BeginEdit(True)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView2.Rows.Remove(DataGridView2.CurrentRow)


        End If
    End Sub

    Private Sub TextBox1_Validating(sender As Object, e As CancelEventArgs) Handles TextBox1.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL CODIGO COLOR")
        End If
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        Dim ffa As New FormFamSub
        ffa.Padre1 = Me
        TextBox12.Text = ""
        TextBox14.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox7.Text = ""
        ffa.Label2.Text = Label13.Text
        ffa.Label3.Text = Label19.Text
        ffa.Label4.Text = "1"
        ffa.Text = "Lista de Familia"
        ffa.ShowDialog()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If TextBox13.Text.ToString().Trim().Length = 0 Then
            MsgBox("Falta Ingresar Informacion de la Familia")
        Else
            Dim ffa1 As New FormFamSub
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox7.Text = ""
            ffa1.Padre1 = Me
            ffa1.Label2.Text = Label13.Text
            ffa1.Label3.Text = Label19.Text
            ffa1.Label4.Text = "2"
            ffa1.Label5.Text = TextBox13.Text.ToString().Trim()
            ffa1.Text = "Lista de Sub Familia"
            ffa1.ShowDialog()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Codificacion_Prodcutos_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        BringToFront()
        Activate()
    End Sub

    Private Sub TextBox2_Validating(sender As Object, e As CancelEventArgs) Handles TextBox2.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR EL COLOR")
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'TextBox8.Text = ""
        'Dim Rs19, Rs113 As SqlDataReader
        'Dim ol As Integer
        'ol = e.RowIndex
        'TextBox1.Text = DataGridView1.Rows(ol).Cells(6).Value
        'TextBox2.Text = DataGridView1.Rows(ol).Cells(7).Value
        'TextBox3.Text = Trim(Mid(DataGridView1.Rows(ol).Cells(1).Value, 13, 25))
        'TextBox7.Text = DataGridView1.Rows(ol).Cells(1).Value
        'TextBox8.Text = DataGridView1.Rows(ol).Cells(2).Value
        'TextBox10.Text = Trim(DataGridView1.Rows(ol).Cells(3).Value)
        'ComboBox3.Text = DataGridView1.Rows(ol).Cells(10).Value
        'abrir()
        'Dim sql19 As String = "select descrip_fam from custom_vianny.dbo.familias where codigo_fam='" + Trim(DataGridView1.Rows(ol).Cells(8).Value) + "'  AND talm_fam ='06'"
        'Dim cmd19 As New SqlCommand(sql19, conx)
        'Rs19 = cmd19.ExecuteReader()
        'Rs19.Read()
        'ComboBox1.Text = Rs19(0)
        'Rs19.Close()
        'llenar_combo_box3(ComboBox2)

        'Dim sql113 As String = "select descrip_sfam from custom_vianny.dbo.subfamilias where codigo_sfam='" + Trim(DataGridView1.Rows(ol).Cells(9).Value) + "' AND familia_sfam ='" + Trim(DataGridView1.Rows(ol).Cells(8).Value) + "' AND talm_sfam = '06'"
        'Dim cmd113 As New SqlCommand(sql113, conx)
        'Rs113 = cmd113.ExecuteReader()
        'Rs113.Read()
        'ComboBox2.Text = Rs113(0)
        'Rs113.Close()
        'TabControl1.SelectedIndex = 0
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            activar()
            TextBox6.Select()
        End If

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox4_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        buscar1()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        log.Label1.Text = TextBox7.Text.ToString().Trim()
        log.Label2.Text = "CODIGO"
        log.Label3.Text = Label13.Text
        log.Show(Me)
    End Sub



    Private Sub TextBox8_Validating(sender As Object, e As CancelEventArgs) Handles TextBox8.Validating
        If DirectCast(sender, TextBox).Text.Length > 0 Then
            Me.ErrorProvider1.SetError(sender, "")
        Else
            Me.ErrorProvider1.SetError(sender, "FALTA INGRESAR LA DESCRIPCION")
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Editar_Producto.Close()
        Editar_Producto.TextBox1.Text = ""
        Editar_Producto.TextBox2.Text = ""
        Editar_Producto.TextBox3.Text = ""
        Editar_Producto.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        Editar_Producto.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
        Editar_Producto.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
        abrir()
        enunciado2 = New SqlCommand("SELECT unid_17 FROM custom_vianny.dbo.cag1700 where linea_17 ='" + TextBox7.Text + "' and ccia ='" + Label13.Text + "'", conx)
        respuesta2 = enunciado2.ExecuteReader
        While respuesta2.Read

            Editar_Producto.ComboBox6.Text = respuesta2.Item("unid_17")
        End While
        respuesta2.Close()
        Editar_Producto.ComboBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString().Trim()
        Editar_Producto.TextBox2.Select()
        Editar_Producto.Label4.Text = Label13.Text
        Editar_Producto.Show(Me)
    End Sub
    Sub bloquear()

        TextBox6.Enabled = False
        TextBox10.Enabled = False
    End Sub
    Sub activar()

        TextBox6.Enabled = True
        TextBox10.Enabled = True
    End Sub
End Class