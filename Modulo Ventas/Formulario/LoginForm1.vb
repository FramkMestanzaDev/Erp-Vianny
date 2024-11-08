Public Class LoginForm1

    ' TODO: inserte el código para realizar autenticación personalizada usando el nombre de usuario y la contraseña proporcionada 
    ' (Consulte http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' El objeto principal personalizado se puede adjuntar al objeto principal del subproceso actual como se indica a continuación: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' donde CustomPrincipal es la implementación de IPrincipal utilizada para realizar la autenticación. 
    ' Posteriormente, My.User devolverá la información de identidad encapsulada en el objeto CustomPrincipal
    ' como el nombre de usuario, nombre para mostrar, etc.
    Private Sub AjustarImagen()
        If MdiParent IsNot Nothing AndAlso TypeOf MdiParent Is MDIParent2 Then
            Dim mdiParentForm As MDIParent2 = CType(MdiParent, MDIParent2)

            ' Verificar si hay una imagen de fondo asignada al panel
            If mdiParentForm.Panel2.BackgroundImage IsNot Nothing Then
                ' Establecer el modo de diseño de la imagen de fondo para que se ajuste al tamaño del panel
                mdiParentForm.Panel2.BackgroundImageLayout = ImageLayout.Stretch
            End If
        End If

    End Sub
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click


        Dim hu As New fusuario
        Dim hu1 As New vusuario
        Dim U As Integer
        hu1.gusuario = UCase(UsernameTextBox.Text.ToString.Trim)
        hu1.ggrupo = ComboBox1.Text
        hu1.gclave = UCase(PasswordTextBox.Text.ToString.Trim)
        U = hu.validar_login(hu1)
        If U = 1 Then
            Dim respuesta As DialogResult


            respuesta = MessageBox.Show("DESEAS INGRESAR A LA EMPRESA?   " + ComboBox3.Text, "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                MsgBox("CONEXION EXITOSA", MessageBoxIcon.Information)

                If Trim(ComboBox1.Text) = "UDP" Or Trim(ComboBox1.Text = "PRODUCCION") Then

                    Me.Visible = False
                    MDIParent2.Label6.Text = UsernameTextBox.Text
                    MDIParent2.Label7.Text = ComboBox1.Text
                    MDIParent2.Label8.Text = ComboBox2.Text
                    If ComboBox3.Text = "VIANNY" Then
                        MDIParent2.Label9.Text = ComboBox3.Text
                        MDIParent2.Label10.Text = "01"
                        MDIParent2.Panel2.BackgroundImage = Global.Modulo_Ventas.Resources.VIANNY
                        MDIParent2.Panel2.BackgroundImageLayout = ImageLayout.Stretch
                    Else
                        MDIParent2.Label9.Text = ComboBox3.Text
                        MDIParent2.Label10.Text = "02"

                    End If
                    AjustarImagen()
                    MDIParent2.Show()
                End If
                If Trim(ComboBox1.Text) = "COMERCIAL MANUFACTURA" Then
                    Me.Visible = False
                    ModuloComercial.Label6.Text = UsernameTextBox.Text
                    ModuloComercial.Label7.Text = ComboBox1.Text
                    ModuloComercial.Label8.Text = ComboBox2.Text
                    If ComboBox3.Text = "VIANNY" Then
                        ModuloComercial.Label9.Text = ComboBox3.Text
                        ModuloComercial.Label10.Text = "01"
                        ModuloComercial.Panel2.BackgroundImage = Global.Modulo_Ventas.Resources.VIANNY
                        MDIParent2.Panel2.BackgroundImageLayout = ImageLayout.Stretch
                    Else
                        ModuloComercial.Label9.Text = ComboBox3.Text
                        ModuloComercial.Label10.Text = "02"

                    End If
                    AjustarImagen()
                    ModuloComercial.Show()

                End If
                If Trim(ComboBox1.Text) = "ALMACEN" Then
                    Me.Visible = False
                    Modulo_Almacen.Label6.Text = UsernameTextBox.Text
                    Modulo_Almacen.Label7.Text = ComboBox1.Text
                    Modulo_Almacen.Label8.Text = ComboBox2.Text
                    If ComboBox3.Text = "VIANNY" Then
                        Modulo_Almacen.Label9.Text = ComboBox3.Text
                        Modulo_Almacen.Label10.Text = "01"
                        Modulo_Almacen.Panel2.BackgroundImage = Global.Modulo_Ventas.Resources.VIANNY
                        Modulo_Almacen.Panel2.BackgroundImageLayout = ImageLayout.Stretch
                    Else
                        Modulo_Almacen.Label9.Text = ComboBox3.Text
                        Modulo_Almacen.Label10.Text = "02"

                    End If
                    AjustarImagen()
                    Modulo_Almacen.Show()
                End If

                If Trim(ComboBox1.Text) = "GRAUS" Then
                    MDIParent1.Show()
                    MDIParent1.ALMACENToolStripMenuItem.Enabled = True
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = True
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False

                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False
                    MDIParent1.UDPToolStripMenuItem.Visible = False
                    MDIParent1.MANUFACTURAToolStripMenuItem.Visible = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False
                    MDIParent1.GRAUSToolStripMenuItem.Visible = True
                    MDIParent1.TesToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False

                End If
                If Trim(ComboBox1.Text) = "TEJEDURIA" Then
                    MDIParent1.Show()
                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False

                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False
                    MDIParent1.TesToolStripMenuItem.Enabled = False
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = True
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False
                    MDIParent1.TablasToolStripMenuItem1.Enabled = False
                    MDIParent1.ToolStripButton4.Enabled = False
                    MDIParent1.PreTejeduriaToolStripMenuItem.Enabled = False


                    MDIParent1.RequerimietoPorPOToolStripMenuItem.Enabled = False
                    MDIParent1.TelaAacabadaToolStripMenuItem.Enabled = False
                    MDIParent1.AlmacenToolStripMenuItem1.Enabled = False
                    MDIParent1.ReportesToolStripMenuItem1.Enabled = False
                End If
                If Trim(ComboBox1.Text) = "CALIDAD" Then
                    MDIParent1.Show()
                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False

                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False
                    MDIParent1.TesToolStripMenuItem.Enabled = False
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False
                    MDIParent1.ToolStripButton4.Enabled = True

                End If
                If Trim(ComboBox1.Text) = "COMERCIAL" Then
                    MDIParent1.Show()
                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False

                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    If UsernameTextBox.Text = "HFARJE" Or UsernameTextBox.Text = "EZIEGLER" Then
                        MDIParent1.ToolStripButton3.Enabled = True
                    Else
                        MDIParent1.ToolStripButton3.Enabled = True
                    End If

                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = True
                    MDIParent1.TesToolStripMenuItem.Enabled = False
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = True
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False


                End If
                If Trim(ComboBox1.Text) = "JEFATURA COMERCIAL" Then
                    MDIParent1.Show()

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = True

                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    If UsernameTextBox.Text = "HFARJE" Or UsernameTextBox.Text = "EZIEGLER" Then
                        MDIParent1.ToolStripButton3.Enabled = True


                    Else
                        MDIParent1.ToolStripButton3.Enabled = True

                    End If
                    MDIParent1.ToolStripButton1.Enabled = True
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = True
                    MDIParent1.TesToolStripMenuItem.Enabled = False
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = True
                    MDIParent1.ToolStripButton1.Enabled = True


                End If
                If Trim(ComboBox1.Text) = "COMERCIAL TEXTIL" Then
                    MDIParent1.Show()

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False

                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    If UsernameTextBox.Text = "HFARJE" Or UsernameTextBox.Text = "EZIEGLER" Then
                        MDIParent1.ToolStripButton3.Enabled = True


                    Else
                        MDIParent1.ToolStripButton3.Enabled = False

                    End If

                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = True
                    MDIParent1.TesToolStripMenuItem.Enabled = False
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = True
                    MDIParent1.ToolStripButton1.Enabled = False


                End If

                If Trim(ComboBox1.Text) = "COBRANZAS" Then
                    MDIParent1.Show()
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False

                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False

                    MDIParent1.TesToolStripMenuItem.Enabled = True
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False
                End If

                If Trim(ComboBox1.Text) = "TRANSPORTE" And Trim(ComboBox3.Text) = "VIANNY" Then
                    MDIParent1.Show()
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = True

                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False

                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False

                    MDIParent1.TesToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False

                End If
                If Trim(ComboBox1.Text) = "TRANSPORTE" And Trim(ComboBox3.Text) = "GRAUS" Then
                    MDIParent1.Show()
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = True

                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False

                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False

                    MDIParent1.TesToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = True

                End If
                If Trim(ComboBox1.Text) = "PRODUCCION CHILCA" Then
                    MDIParent1.Show()
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = True
                    If UsernameTextBox.Text = "HFARJE" Or UsernameTextBox.Text = "EZIEGLER" Then

                        MDIParent1.ToolStripButton3.Enabled = True
                    Else
                        MDIParent1.ToolStripButton3.Enabled = False

                    End If
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False

                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = True

                    MDIParent1.TesToolStripMenuItem.Enabled = False

                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False

                End If
                If Trim(ComboBox1.Text = "ADMINISTRADOR") Then
                    MDIParent1.Show()
                    If ComboBox3.Text = "VIANNY" Then
                        MDIParent1.COMERCIALToolStripMenuItem.Enabled = True

                        MDIParent1.ALMACENToolStripMenuItem.Enabled = True

                        MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = True
                        MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = True
                        MDIParent1.ToolStripButton2.Enabled = True
                        MDIParent1.ToolStripButton3.Enabled = True
                        MDIParent1.ToolStripButton3.Enabled = True
                    Else
                        MDIParent1.ALMACENToolStripMenuItem.Enabled = True
                        MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = True
                        MDIParent1.COMERCIALToolStripMenuItem.Enabled = True

                        MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = True
                        MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = True
                        MDIParent1.ToolStripButton2.Enabled = True
                        MDIParent1.UDPToolStripMenuItem.Visible = True
                        MDIParent1.MANUFACTURAToolStripMenuItem.Visible = True
                        MDIParent1.VENDEDORToolStripMenuItem.Enabled = True
                        MDIParent1.GRAUSToolStripMenuItem.Visible = True
                        MDIParent1.TesToolStripMenuItem.Enabled = True
                        MDIParent1.ToolStripButton1.Enabled = True
                        MDIParent1.ToolStripButton3.Enabled = True
                    End If

                End If

                If Trim(ComboBox1.Text = "GERENCIA") Then
                    MDIParent1.Show()
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = True

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = True

                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = True
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = True
                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False

                    MDIParent1.TesToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = True
                End If
                If Trim(ComboBox1.Text = "TESORERIA") Then
                    MDIParent1.Show()
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False

                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False
                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False

                    MDIParent1.TesToolStripMenuItem.Enabled = True
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton3.Enabled = False
                End If
                If Trim(ComboBox1.Text = "CONTABILIDAD") Then
                    MDIParent1.Show()
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False

                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False
                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = False

                    MDIParent1.TesToolStripMenuItem.Enabled = True
                    MDIParent1.ToolStripButton3.Enabled = False
                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = True
                End If
                If Trim(ComboBox1.Text = "CHILCA") Then
                    MDIParent1.REPORTESGERENCIAToolStripMenuItem.Enabled = False
                    MDIParent1.COMERCIALToolStripMenuItem.Enabled = False

                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False
                    If UsernameTextBox.Text = "HFARJE" Or UsernameTextBox.Text = "EZIEGLER" Then
                        MDIParent1.ToolStripButton3.Enabled = True


                    Else
                        MDIParent1.ToolStripButton3.Enabled = False

                    End If
                    MDIParent1.ADMINISTRADORToolStripMenuItem.Enabled = False
                    MDIParent1.ToolStripButton2.Enabled = False
                    MDIParent1.ALMACENToolStripMenuItem.Enabled = False
                    MDIParent1.VENDEDORToolStripMenuItem.Enabled = True

                    MDIParent1.TesToolStripMenuItem.Enabled = False

                    MDIParent1.ToolStripButton1.Enabled = False
                    MDIParent1.PRODUCCIONToolStripMenuItem.Enabled = False
                End If
                If ComboBox3.Text = "VIANNY" Then

                    'MDIParent1.PictureBox1.Image = Global.Modulo_Ventas.Resources.Logo_Vianny_fondo_índigo
                    'MDIParent2.Panel2.BackgroundImage = Global.Modulo_Ventas.Resources.VIANNY
                    'MDIParent2.Panel2.BackgroundImageLayout = ImageLayout.Zoom
                    MDIParent1.Text = "CONSORCIO TEXTIL VIANNY"
                    Me.Visible = False
                    MDIParent1.Label3.Text = UsernameTextBox.Text
                    MDIParent1.Label4.Text = ComboBox1.Text
                    MDIParent1.Label5.Text = ComboBox2.Text
                    MDIParent1.Label7.Text = ComboBox3.Text
                    MDIParent1.Label9.Text = "01"
                Else
                    If ComboBox3.Text = "GRAUS" Then
                        MDIParent1.PictureBox1.Image = Global.Modulo_Ventas.Resources.LOGO_DG_PNG
                        MDIParent1.Text = "GRAUS INDUSTRIAS TEXTIL"
                        MDIParent1.Label7.Text = ComboBox3.Text
                        MDIParent1.Label9.Text = "02"
                        Me.Visible = False
                        MDIParent1.Label3.Text = UsernameTextBox.Text
                        MDIParent1.Label4.Text = ComboBox1.Text
                        MDIParent1.Label5.Text = ComboBox2.Text
                    Else
                        MDIParent1.PictureBox1.Image = Global.Modulo_Ventas.Resources.diseñoy_textura
                        MDIParent1.Text = "DISEÑO Y TEXTURA"
                    End If
                End If

            Else
            MsgBox("LAS CREDENCIALES SON INCORRECTAS", MessageBoxIcon.Exclamation)
        End If

        Dim GH As New LoginForm1
            GH.Dispose()
        Else

            MsgBox("USUARIO O CONTRASEÑA INCORRECTA")
        End If


    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub LoginForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 5
    End Sub

    Private Sub PasswordTextBox_Enter(sender As Object, e As EventArgs) Handles PasswordTextBox.Enter
        Dim hu As New fusuario
        Dim hu1 As New vusuario
        hu1.ggrupo = UsernameTextBox.Text
        If hu.BUSCAR_GRUPO(hu1) = "1" Then
            UsernameTextBox.Select()
        Else

            ComboBox1.Items.Clear()
            ComboBox1.Items.Add(Trim(hu.BUSCAR_GRUPO(hu1)))

            ComboBox1.SelectedIndex = 0
            If ComboBox1.Text = "GRAUS" Then
                ComboBox3.SelectedIndex = 1
            Else
                ComboBox3.SelectedIndex = 0
            End If
        End If


    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub UsernameTextBox_TextChanged(sender As Object, e As EventArgs) Handles UsernameTextBox.TextChanged

    End Sub

    Private Sub UsernameLabel_Click(sender As Object, e As EventArgs) Handles UsernameLabel.Click

    End Sub

    Private Sub PasswordTextBox_TextChanged(sender As Object, e As EventArgs) Handles PasswordTextBox.TextChanged

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub
End Class
