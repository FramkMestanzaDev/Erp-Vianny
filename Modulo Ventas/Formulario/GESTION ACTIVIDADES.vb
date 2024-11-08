Imports System.Data.SqlClient
Public Class GESTION_ACTIVIDADES
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
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
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim hh As New vgestionact
        Dim kk As New fgestionact

        hh.gid_ga = TextBox8.Text
        hh.gfecha_g = DateTimePicker1.Value
        hh.gvededor_co = TextBox1.Text
        hh.gvendedor_nom = TextBox3.Text
        hh.gcliente_ruc = TextBox2.Text
        hh.gcliente_rs = TextBox4.Text
        hh.gprox_visita = DateTimePicker2.Value
        hh.gcontacto = TextBox5.Text
        hh.gestado = ComboBox1.Text
        hh.gdescripcion = TextBox6.Text
        hh.gubicacion = TextBox7.Text

        hh.gCORREO = TextBox10.Text
        hh.gCELULAR = TextBox11.Text
        hh.gCONTACTARON = ComboBox3.Text
        If ComboBox2.Text = "SOLES" Then
            hh.gMONEDA = 1
        Else
            If ComboBox2.Text = "DOLARES" Then
                hh.gMONEDA = 2
            End If
        End If
        If TextBox9.Text = "" Then
            hh.gMONTO = "0.00"
        Else
            hh.gMONTO = TextBox9.Text
        End If

        kk.insertar_gestoosn_actividades(hh)
        MsgBox("SE INGRESO LA INFORMACION CORRECTAMENTE")

        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        Dim func As New fgestionact
        Me.TextBox8.Text = func.correlativo_GESTION + 1
        Select Case TextBox8.Text.Length
            Case "1" : TextBox8.Text = "0000000" & "" & TextBox8.Text
            Case "2" : TextBox8.Text = "000000" & "" & TextBox8.Text
            Case "3" : TextBox8.Text = "00000" & "" & TextBox8.Text
            Case "4" : TextBox8.Text = "0000" & "" & TextBox8.Text
            Case "5" : TextBox8.Text = "000" & "" & TextBox8.Text
            Case "6" : TextBox8.Text = "00" & "" & TextBox8.Text
            Case "7" : TextBox8.Text = "0" & "" & TextBox8.Text
            Case "8" : TextBox8.Text = TextBox8.Text

        End Select
    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.SetToolTip(PictureBox1, "GUARDAR INFORMACION")
        ToolTip1.ToolTipTitle = "GESTION DE ACTIVIDADES"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info

    End Sub
    Dim Rsr2 As SqlDataReader
    Private Sub GESTION_ACTIVIDADES_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        Dim func As New fgestionact
        Me.TextBox8.Text = func.correlativo_GESTION + 1
        Select Case TextBox8.Text.Length
            Case "1" : TextBox8.Text = "0000000" & "" & TextBox8.Text
            Case "2" : TextBox8.Text = "000000" & "" & TextBox8.Text
            Case "3" : TextBox8.Text = "00000" & "" & TextBox8.Text
            Case "4" : TextBox8.Text = "0000" & "" & TextBox8.Text
            Case "5" : TextBox8.Text = "000" & "" & TextBox8.Text
            Case "6" : TextBox8.Text = "00" & "" & TextBox8.Text
            Case "7" : TextBox8.Text = "0" & "" & TextBox8.Text
            Case "8" : TextBox8.Text = TextBox8.Text

        End Select
        TextBox2.Select()

        abrir()
        Dim sql102 As String = "SELECT codigo_ven FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + TextBox3.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            TextBox1.Text = Rsr2(0)

        End If

        Rsr2.Close()
        'Select Case Trim(TextBox3.Text)
        '    Case "GBEDON" : TextBox1.Text = "0010"
        '    Case "VINCIO" : TextBox1.Text = "0022"
        '    Case "DBRAVO" : TextBox1.Text = "0023"
        '    Case "JSALINAS" : TextBox1.Text = "0025"
        '    Case "GCUEVA" : TextBox1.Text = "0029"
        '    Case "AMENDO" : TextBox1.Text = "0026"
        '    Case "VGRAUS" : TextBox1.Text = "0007"
        '    Case "VPIZARRO" : TextBox1.Text = "0027"
        '    Case "GBALVIN" : TextBox1.Text = "0028"
        '    Case "VSILVERIO" : TextBox1.Text = "0005"
        '    Case "WSALINAS" : TextBox1.Text = "0034"
        'End Select
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            Listar_Clientes.TextBox2.Text = 19
            Listar_Clientes.TextBox3.Text = Trim(TextBox1.Text)
            Listar_Clientes.Show()
        End If
    End Sub
End Class