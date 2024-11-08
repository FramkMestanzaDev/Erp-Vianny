Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class Op_Manufactura
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim Rsr2, Rsr3, Rsr222, Rsr212, Rsr2125 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Function correlativo()
        Dim va As New vproducot
        Dim va1 As New fproductos

        va.gcia = Label33.Text
        va.gt_pedido = "P"
        'TextBox1.Text = va1.op_comercial(va)
        Return va1.op_comercial(va)
    End Function
    Private Sub Op_Manufactura_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 2
        ComboBox5.SelectedIndex = 0
        RadioButton1.Checked = True
        TextBox17.Text = "OTROS"
        colorprenda.Text = "C00002"
        TextBox1.Text = correlativo()
        Dim hu As New fcliente
        Dim jk As New vcliente
        Dim va11, va21 As String
        va11 = ""
        va21 = ""
        abrir()
        Dim sql102 As String = "Select cele FROM custom_vianny.DBO.TAB0100 A  Where ccia='" + Label33.Text + "' AND CTAB='MERCHA' and codigo ='" + Trim(TextBox10.Text) + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            va11 = Rsr2(0)
        End If
        Rsr2.Close()

        TextBox26.Text = va11
        DateTimePicker2.Value = DateTimePicker1.Value.AddDays(30)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Ficha.TextBox2.Text = 2
        'Ficha.Label2.Text = Label33.Text
        'Ficha.Show()
        CPrenda.Label2.Text = Label33.Text
        CPrenda.Label3.Text = 2
        CPrenda.Label4.Text = 1
        CPrenda.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim famili As New FormFamilia
        famili.Label2.Text = 1
        famili.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim padronas As New Form_Padron
        padronas.Label2.Text = Label33.Text
        padronas.Label3.Text = 1
        padronas.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim color As New Ficha
        color.TextBox2.Text = 2
        color.Label2.Text = Label33.Text
        color.ShowDialog()

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim tallas As New Cantidad_Tallas
        tallas.DataGridView1.Rows.Add(11)
        'MsgBox(DataGridView2.Rows(0).Cells(12).Value)
        For i = 0 To 9
            tallas.DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(0).Cells(i + 13).Value

        Next
        For i = 0 To 9
            tallas.DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(1).Cells(i + 1).Value
            tallas.DataGridView1.Rows(i).Cells(2).Value = DataGridView4.Rows(1).Cells(i + 1).Value
        Next
        tallas.TextBox1.Text = TextBox21.Text
        tallas.ShowDialog()

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim div As New Form_Cliente
        div.Label1.Text = "DIVISION"
        div.TextBox2.Text = TextBox22.Text
        div.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim pago As New Form_Cliente
        pago.Label1.Text = "F. PAGO"
        pago.TextBox3.Text = 1

        pago.ShowDialog()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim pais As New Form_Cliente
        pais.Label1.Text = "PAIS"
        pais.TextBox2.Text = TextBox22.Text
        pais.ShowDialog()
    End Sub
    Private Sub limpiar()

        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox8.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox32.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox21.Text = ""
        TextBox19.Text = ""
        TextBox25.Text = ""
        TextBox20.Text = ""
        colorprenda.Text = ""
        colortela.Text = ""
        TextBox29.Text = ""
        TextBox27.Text = ""
        TextBox28.Text = ""
        TextBox30.Text = ""
        TextBox31.Text = ""
        DataGridView2.DataSource = ""
        DataGridView4.DataSource = ""
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        Form2.TextBox3.Text = 1
        Form2.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim cli As New Listar_Clientes
        cli.TextBox2.Text = 5
        cli.TextBox3.Text = "A"
        cli.ShowDialog()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Dim ll, FTY As DataTable
    Dim Rsr21, Rsr22 As SqlDataReader
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim gh As New fop
            Dim kk As New vop
            Dim lp As Integer
            TextBox1.Enabled = False
            Select Case Trim(TextBox1.Text).Length
                Case "1" : TextBox1.Text = "01" & "0000000" & TextBox1.Text
                Case "2" : TextBox1.Text = "01" & "000000" & TextBox1.Text
                Case "3" : TextBox1.Text = "01" & "00000" & TextBox1.Text
                Case "4" : TextBox1.Text = "01" & "0000" & TextBox1.Text
                Case "5" : TextBox1.Text = "01" & "000" & TextBox1.Text
                Case "6" : TextBox1.Text = "01" & "00" & TextBox1.Text
                Case "7" : TextBox1.Text = "01" & "0" & TextBox1.Text
                Case "8" : TextBox1.Text = "01" & TextBox1.Text

            End Select
            kk.gncom_3 = TextBox1.Text

            kk.gcia = Label33.Text
            ll = gh.buscar_op(kk)

            DataGridView3.DataSource = ll
            lp = DataGridView3.Rows.Count
            If lp <> 0 Then
                For i = 0 To lp - 1

                    If Trim(TextBox26.Text) = Trim(DataGridView3.Rows(0).Cells(44).Value) Or Trim(MDIParent1.Label4.Text) = "ADMINISTRADOR" Then
                        TextBox22.Text = DataGridView3.Rows(0).Cells(7).Value
                        TextBox2.Text = Trim(DataGridView3.Rows(0).Cells(8).Value)
                        TextBox26.Text = DataGridView3.Rows(0).Cells(44).Value
                        Select Case TextBox26.Text
                            Case "0001" : TextBox10.Text = "VSILVERIO"
                            Case "0003" : TextBox10.Text = "GBALVIN"
                            Case "0025" : TextBox10.Text = "VGRAUS"
                        End Select
                        TextBox3.Text = Trim(DataGridView3.Rows(0).Cells(10).Value)
                        TextBox23.Text = DataGridView3.Rows(0).Cells(79).Value
                        TextBox4.Text = DataGridView3.Rows(0).Cells(11).Value
                        TextBox24.Text = DataGridView3.Rows(0).Cells(80).Value
                        TextBox5.Text = DataGridView3.Rows(0).Cells(4).Value
                        TextBox32.Text = DataGridView3.Rows(0).Cells(5).Value & DataGridView3.Rows(0).Cells(6).Value
                        TextBox88.Text = Trim(DataGridView3.Rows(0).Cells(86).Value)
                        TextBox12.Text = Trim(DataGridView3.Rows(0).Cells(53).Value)
                        TextBox27.Text = Trim(DataGridView3.Rows(0).Cells(43).Value)
                        TextBox28.Text = Trim(DataGridView3.Rows(0).Cells(52).Value)
                        TextBox13.Text = DataGridView3.Rows(0).Cells(15).Value
                        TextBox14.Text = Trim(DataGridView3.Rows(0).Cells(51).Value)
                        TextBox15.Text = Trim(DataGridView3.Rows(0).Cells(30).Value) & "/" & Trim(DataGridView3.Rows(0).Cells(31).Value) & "/" & Trim(DataGridView3.Rows(0).Cells(32).Value) & "/" & Trim(DataGridView3.Rows(0).Cells(33).Value) & "/" & Trim(DataGridView3.Rows(0).Cells(34).Value) & "/" & Trim(DataGridView3.Rows(0).Cells(35).Value) & "/" & Trim(DataGridView3.Rows(0).Cells(36).Value) & "/" & Trim(DataGridView3.Rows(0).Cells(37).Value) & "/" & Trim(DataGridView3.Rows(0).Cells(38).Value) & "/" & Trim(DataGridView3.Rows(0).Cells(39).Value)
                        TextBox16.Text = Trim(DataGridView3.Rows(0).Cells(17).Value)
                        'TextBox30.Text = Trim(DataGridView3.Rows(0).Cells(89).Value) & " " & Trim(DataGridView3.Rows(0).Cells(90).Value) & " " & Trim(DataGridView3.Rows(0).Cells(91).Value)
                        TextBox31.Text = Trim(DataGridView3.Rows(0).Cells(82).Value)
                        TextBox30.Text = Trim(DataGridView3.Rows(0).Cells(83).Value)
                        TextBox6.Text = DataGridView3.Rows(0).Cells(9).Value
                        TextBox8.Text = DataGridView3.Rows(0).Cells(45).Value
                        TextBox21.Text = DataGridView3.Rows(0).Cells(27).Value
                        TextBox19.Text = DataGridView3.Rows(0).Cells(47).Value
                        TextBox25.Text = DataGridView3.Rows(0).Cells(46).Value
                        TextBox19.Text = DataGridView3.Rows(0).Cells(47).Value
                        TextBox17.Text = Trim(DataGridView3.Rows(0).Cells(24).Value)
                        colorprenda.Text = Trim(DataGridView3.Rows(0).Cells(48).Value)
                        TextBox18.Text = Trim(DataGridView3.Rows(0).Cells(85).Value)
                        colortela.Text = Trim(DataGridView3.Rows(0).Cells(48).Value)
                        TextBox20.Text = Trim(DataGridView3.Rows(0).Cells(42).Value)
                        DateTimePicker1.Value = DataGridView3.Rows(0).Cells(12).Value
                        DateTimePicker2.Value = DataGridView3.Rows(0).Cells(13).Value
                        TextBox29.Text = DataGridView3.Rows(0).Cells(48).Value.ToString().Trim()
                        RUTAMOLDE.Text = DataGridView3.Rows(0).Cells(92).Value.ToString().Trim()
                        ESTILOM.Text = DataGridView3.Rows(0).Cells(93).Value.ToString().Trim()
                        LAVADOM.Text = DataGridView3.Rows(0).Cells(94).Value.ToString().Trim()
                        TELAM.Text = DataGridView3.Rows(0).Cells(95).Value.ToString().Trim()
                        DESCRIPM.Text = DataGridView3.Rows(0).Cells(96).Value.ToString().Trim()
                        MODELISTA.Text = DataGridView3.Rows(0).Cells(97).Value.ToString().Trim()
                        ANALISTA.Text = DataGridView3.Rows(0).Cells(98).Value.ToString().Trim()
                        TextBox35.Text = DataGridView3.Rows(0).Cells(99).Value.ToString().Trim()
                        TextBox36.Text = DataGridView3.Rows(0).Cells(100).Value.ToString().Trim()
                        If DataGridView3.Rows(0).Cells(50).Value = 1 Then
                            RadioButton1.Checked = True
                        Else
                            If DataGridView3.Rows(0).Cells(50).Value = 0 Then
                                RadioButton2.Checked = True
                                Button11.Enabled = False
                                Button12.Enabled = False
                                Button13.Enabled = False
                            End If
                        End If
                        If Trim(DataGridView3.Rows(0).Cells(54).Value) = "N" Then
                            ComboBox3.SelectedIndex = 0
                        Else
                            ComboBox3.SelectedIndex = 1
                        End If
                        If Trim(DataGridView3.Rows(0).Cells(22).Value) = "Terrestre" Then
                            ComboBox4.SelectedIndex = 2
                        Else
                            If Trim(DataGridView3.Rows(0).Cells(22).Value) = "Aereo" Then
                                ComboBox4.SelectedIndex = 1
                            End If

                            ComboBox4.SelectedIndex = 0
                        End If
                        If Trim(DataGridView3.Rows(0).Cells(56).Value) = "01" Then
                            ComboBox5.SelectedIndex = 0
                        Else
                            ComboBox5.SelectedIndex = 1
                        End If
                        If Trim(DataGridView3.Rows(0).Cells(55).Value) = "PVT" Then
                            ComboBox1.SelectedIndex = 2
                        Else
                            If Trim(DataGridView3.Rows(0).Cells(55).Value) = "FOV" Then
                                ComboBox1.SelectedIndex = 1

                            Else

                                ComboBox1.SelectedIndex = 0
                            End If
                        End If

                        Dim ab As New vtallas
                        Dim ab1 As New Padron_tallas
                        ab.gcodigo = Trim(TextBox14.Text)
                        ab.gccia = Label33.Text
                        FTY = ab1.mostrar_padrom_tallas_SELECCIONADO(ab)
                        DataGridView2.DataSource = FTY
                        DataGridView4.DataSource = FTY
                        'TextBox20.Text = Trim(DataGridView2.Rows(0).Cells(13).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(14).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(15).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(16).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(17).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(18).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(19).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(20).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(21).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(22).Value) & "/" & Trim(DataGridView2.Rows(0).Cells(23).Value)
                        Dim sql1021 As String = "select * from custom_vianny.dbo.qag160d where ncom_16d ='" + (TextBox1.Text) + "' and ccia ='" + Label33.Text + "' order by correl_16d"
                        Dim cmd1021 As New SqlCommand(sql1021, conx)
                        Rsr21 = cmd1021.ExecuteReader()
                        Dim i51 As Integer
                        i51 = 1

                        While Rsr21.Read()

                            DataGridView2.Rows(1).Cells(i51).Value = Rsr21(5)

                            i51 = i51 + 1
                        End While
                        Rsr21.Close()

                        Dim sql10217 As String = "select * from custom_vianny.dbo.qag160c where ncom_16c ='" + (TextBox1.Text) + "' and ccia ='" + Label33.Text + "' order by correl_16c"
                        Dim cmd10217 As New SqlCommand(sql10217, conx)
                        Rsr22 = cmd10217.ExecuteReader()
                        Dim i517 As Integer
                        i517 = 1

                        While Rsr22.Read()

                            DataGridView4.Rows(1).Cells(i517).Value = Rsr22(6)

                            i517 = i517 + 1
                        End While
                        Rsr22.Close()

                        TextBox22.Enabled = False
                        TextBox2.Enabled = False
                        TextBox5.Enabled = False
                        TextBox6.Enabled = False
                        TextBox8.Enabled = False
                        TextBox17.Enabled = False
                        TextBox18.Enabled = False
                        TextBox21.Enabled = False
                        TextBox19.Enabled = False
                        TextBox25.Enabled = False
                        TextBox13.Enabled = False

                        TextBox16.Enabled = False
                        TextBox20.Enabled = False
                        colorprenda.Enabled = False
                        colortela.Enabled = False
                        TextBox29.Enabled = False
                        Button10.Enabled = False
                        Button3.Enabled = False
                        Button1.Enabled = False
                        Button4.Enabled = False
                        Button5.Enabled = False
                        Button2.Enabled = False
                        Button6.Enabled = False
                        Button13.Enabled = True
                        Button11.Enabled = False
                    Else
                        MsgBox("LA OP A COSULTAR NO ES DE SERVICIO O NO PERTENECE AL VENDEDOR")
                        correlativo()
                        TextBox22.Enabled = False
                        TextBox2.Enabled = False
                        TextBox5.Enabled = False
                        TextBox6.Enabled = False
                        TextBox8.Enabled = False
                        TextBox17.Enabled = False
                        TextBox18.Enabled = False
                        TextBox21.Enabled = False
                        TextBox19.Enabled = False
                        TextBox25.Enabled = False
                        TextBox13.Enabled = False

                        TextBox16.Enabled = False
                        TextBox20.Enabled = False
                        colorprenda.Enabled = False
                        colortela.Enabled = False
                        TextBox29.Enabled = False
                        Button10.Enabled = False
                        Button11.Enabled = False
                        Button3.Enabled = False
                        Button1.Enabled = False
                        Button4.Enabled = False
                        Button5.Enabled = False
                        Button2.Enabled = False
                        Button6.Enabled = False
                        ComboBox1.Enabled = False
                        ComboBox2.Enabled = False
                        ComboBox3.Enabled = False
                        ComboBox4.Enabled = False
                        ComboBox5.Enabled = False


                        TextBox3.Enabled = False
                        TextBox4.Enabled = False
                        TextBox5.Enabled = False
                        TextBox6.Enabled = False
                        TextBox7.Enabled = False
                    End If
                Next


            Else
                'TextBox2.ReadOnly = False
                'TextBox3.ReadOnly = False
                'TextBox4.ReadOnly = False
                TextBox33.Enabled = True
                TextBox5.ReadOnly = False
                TextBox6.ReadOnly = False
                TextBox7.ReadOnly = False
                TextBox8.ReadOnly = False
                TextBox16.ReadOnly = False
                TextBox30.ReadOnly = False
                TextBox31.ReadOnly = False
                TextBox20.ReadOnly = False
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox3.Enabled = True
                ComboBox4.Enabled = True
                ComboBox5.Enabled = True
                DateTimePicker1.Enabled = True
                DateTimePicker2.Enabled = True
                Button1.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True
                Button5.Enabled = True
                Button6.Enabled = True
                Button7.Enabled = True
                Button8.Enabled = True
                TextBox5.Enabled = True
                Button10.Enabled = True
                Button11.Enabled = True
                TextBox6.Enabled = True
                TextBox8.Enabled = True
                TextBox33.Select()
            End If


        End If
        If e.KeyCode = Keys.F1 Then

            ListaOp.ShowDialog()
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub
    Private Function Validar_Textobox() As Boolean
        If TextBox8.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Precio está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If

        If TextBox12.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Familia está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If

        If TextBox3.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Division está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If

        If TextBox4.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Pais está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If TextBox5.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Pedido está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If TextBox21.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Total de Prendas está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If TextBox19.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Forma de Pago está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If TextBox14.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Padron de Tallas está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If TextBox2.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Cliente está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If TextBox1.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Op está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If TextBox18.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Color de Prenda está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        If TextBox32.Text.Trim().Length = 0 Then
            MessageBox.Show("El Campo Od está vacío", "Cuadro de Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
        Return True
    End Function
    Sub ENVIAR_CORREO(opf)
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim cabecera As String

        Dim usuario, contras As String
        usuario = ""
        contras = ""
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "SELECT correo,contrasena FROM usuario_modulo where merchan_3 ='" + Trim(TextBox26.Text.ToString()) + "'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        If Rsr1991.Read() Then
            usuario = Rsr1991(0)
            contras = Rsr1991(1)
        End If
        Rsr1991.Close()
        If Label43.Text = 1 Then
            cabecera = "SE EDITO LA OP"
        Else
            cabecera = "SE CREO LA OP"
        End If

        Dim Mailz As String = "" &
      "<!DOCTYPE html>
<html xmlns='http://www.w3.org/1999/xhtml'>
<head>
    <title>" + cabecera + "</title>
</head>
<body>
    <font color='green'>" + cabecera + "</font><br/><br/>
    <table border='1' cellspacing='0' cellpadding='5'>
        <thead>
            <tr style='background-color: #e0f7fa;'>
                <th align='center'><font color='blue'>OP</font></th>            
                <th align='center'><font color='blue'>PO</font></th>
                <th align='center'><font color='blue'>Familia</font></th>
                <th align='center'><font color='blue'>Cliente</font></th>
                <th align='center'><font color='blue'>Estilo</font></th>
                <th align='center'><font color='blue'>Producto</font></th>
                <th align='center'><font color='blue'>Fecha de Entrega</font></th>

                <th align='center'><font color='blue'>Cantidad Solicitada</font></th>
                <th align='center'><font color='blue'>Tela</font></th>

                <th align='center'><font color='blue'>Lavado</font></th>
                <th align='center'><font color='blue'>Observación</font></th>
                <th align='center'><font color='blue'>Ficha</font></th>
            </tr>
        </thead>
        <tbody>
            <tr>              
                <td  align='center'> " + opf + "</td>                       
                      
                <td  align='center'>" + TextBox5.Text.ToString.Trim + "</td>              
                <td  align='center'>" + TextBox12.Text.ToString.Trim + " </td>          
                <td  align='center'>" + TextBox2.Text.ToString.Trim + " </td>                
                <td  align='center'>" + TextBox6.Text.ToString.Trim + " </td>  
                 <td  align='center'>" + TextBox16.Text.ToString.Trim + " </td>  
                <td  align='center'> " + DateTimePicker2.Value.ToString() + " </td>  
                
                <td  align='center'>" + TextBox21.Text.ToString.Trim + " </td>                   
                <td  align='center'> " + TextBox31.Text.ToString.Trim + " </td> 
                
                <td  align='center'> " + TextBox30.Text.ToString.Trim + " </td>           
                <td  align='center'>" + TextBox20.Text.ToString.Trim + " </td>           
                <td  align='center'><a href=" + Trim(TextBox88.Text) + ">" + Trim(TextBox88.Text) + "</a> </td>
            </tr>
        </tbody>
    </table>
</body>
</html>"
        Dim Rs As SqlDataReader
        Dim sql As String = "select correo from lista_correos where formulario ='PRODUCCION          '"
        Dim cmd As New SqlCommand(sql, conx)
        Rs = cmd.ExecuteReader()
        Rs.Read()

        With message

            .From = New System.Net.Mail.MailAddress(usuario)
            .To.Add(Rs(0))
            '.To.Add("fmestanza@viannysac.com")
            Rs.Close()
            .IsBodyHtml = True
            .Body = Mailz
            .Subject = cabecera & TextBox1.Text & " --- " & "PO " & TextBox5.Text
            .Priority = System.Net.Mail.MailPriority.Normal
        End With

        With smtp

            .EnableSsl = True
            .Port = "587"
            .Host = "smtppro.zoho.com"
            .Credentials = New Net.NetworkCredential(usuario, contras)

            .Send(message)
        End With
        Try
            MessageBox.Show("EL mensaje de correo ha sido enviado.", "Correo enviado",
                             MessageBoxButtons.OK)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
                             MessageBoxButtons.OK)
        End Try
    End Sub
    Dim op, opf As String
    Sub IngresarNotificacion()
        Dim info As String = ""
        If Label43.Text = "1" Then
            info = "Creacion"
        Else
            info = "Edicion"
        End If
        Dim cmd16 As New SqlCommand("INSERT INTO Notificaciones(FecNot,DetNot,JCoNot,SisNot,Co1Not,Co2Not,JPrNot,APrNot,Pr1Not,Pr2Not,Pr3Not,An1Not,An2Not,JudNot,ComNot,EmpNot,ModNot) 
                                                       VALUES (@FecNot,@DetNot,@JCoNot,@SisNot,@Co1Not,@Co2Not,@JPrNot,@APrNot,@Pr1Not,@Pr2Not,@Pr3Not,@An1Not,@An2Not,@JudNot,@ComNot,@EmpNot,@ModNot)", conx)
        cmd16.Parameters.AddWithValue("@FecNot", DateTimePicker3.Value)
        cmd16.Parameters.AddWithValue("@DetNot", "" + info + " de la Po " & TextBox5.Text.ToString().Trim() & " Op -" & TextBox1.Text.ToString().Trim())
        cmd16.Parameters.AddWithValue("@JCoNot", "0")
        cmd16.Parameters.AddWithValue("@SisNot", "0")
        cmd16.Parameters.AddWithValue("@Co1Not", "0")
        cmd16.Parameters.AddWithValue("@Co2Not", "0")
        cmd16.Parameters.AddWithValue("@JPrNot", "0")
        cmd16.Parameters.AddWithValue("@APrNot", "0")
        cmd16.Parameters.AddWithValue("@Pr1Not", "0")
        cmd16.Parameters.AddWithValue("@Pr2Not", "0")
        cmd16.Parameters.AddWithValue("@Pr3Not", "0")
        cmd16.Parameters.AddWithValue("@An1Not", "0")
        cmd16.Parameters.AddWithValue("@An2Not", "0")
        cmd16.Parameters.AddWithValue("@JudNot", "0")
        cmd16.Parameters.AddWithValue("@ComNot", "0")
        cmd16.Parameters.AddWithValue("@EmpNot", "01")
        cmd16.Parameters.AddWithValue("@ModNot", "OP")
        cmd16.ExecuteNonQuery()
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim ki As New vopmanufac
        Dim ki1 As New FOPMANUF
        Dim ju As New vopmanufac
        Dim hj As New insertar_codigo
        Dim ghj, codigo As String
        Dim uo, ko, mn, lñ As New vresto_op
        Dim j, SUMA, na As Integer
        abrir()
        op = TextBox1.Text.ToString().Trim()
        Dim sql1 As String = "select count(ncom_3) from custom_vianny.dbo.qag0300 where ncom_3 ='" + op + "' and ccia ='" + Label33.Text + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        Rsr2125 = cmd1.ExecuteReader()
        If Rsr2125.Read() = True And Convert.ToInt32(Rsr2125(0)) > 0 Then
            If Label43.Text = 1 Then
                opf = op
            Else
                opf = correlativo()
                MsgBox("LA OP INGRESADA YA EXISTE, LA NUEVA OP ES : " + opf)
            End If
        Else
            opf = op
        End If
        Rsr2125.Close()



        If Validar_Textobox() = True Then



            ki.gncom_3 = opf
            ki.gfich_3 = TextBox22.Text
            ki.gfcom_3 = DateTimePicker1.Value
            If ComboBox2.Text = "S/." Then
                ki.gmone_3 = 1
            Else
                If ComboBox2.Text = "US$" Then
                    ki.gmone_3 = 2
                End If
            End If


            ki.gfamilia = TextBox27.Text

            ki.gccia = Label33.Text
            ki.gtallador = TextBox14.Text
            ki.gFCome1_3 = DateTimePicker2.Value
            ki.gfrequerida_3 = DateTimePicker2.Value
            ki.gtcam_3 = 0.0000
            ki.gprogram_3 = TextBox5.Text
            ki.gflag_3 = 1
            ki.gdescri_3 = Trim(TextBox16.Text)
            ki.galterno_3 = Trim(TextBox30.Text)
            ki.gzona = TextBox24.Text
            ki.gestilo_3 = TextBox6.Text
            ki.gpfob_3 = TextBox8.Text
            ki.gcantp_3 = 0.00
            ki.gcants_3 = TextBox21.Text
            ki.gtipped_3 = "P"
            ki.gmruta1_3 = ""
            ki.gmruta5_3 = TextBox35.Text.ToString().Trim()
            ki.gmruta6_3 = TextBox36.Text.ToString().Trim()
            ki.gotros1_3 = ""
            ki.gotros2_3 = ""
            ki.gtela1_3 = Trim(TextBox34.Text)
            ki.glavadoten_3 = ""
            Select Case ComboBox1.Text
                Case "P. VENTA" : ki.gmdvent_3 = "PVT"
                Case "F.O.V." : ki.gmdvent_3 = "FOV"
                Case "C.I.F." : ki.gmdvent_3 = "CIF"
            End Select
            Select Case ComboBox3.Text
                Case "EXTRANJERO" : ki.gtipo_mercado = "E"
                Case "NACIONAL" : ki.gtipo_mercado = "N"
            End Select
            Dim va11, va21 As String
            va11 = ""
            va21 = ""
            ki.gmerchan_3 = Trim(TextBox26.Text)
            ki.gfpago = Trim(TextBox25.Text)
            ki.gbroker_3 = "0001"
            Select Case ComboBox4.Text
                Case "MARITIMO" : ki.gviatra = "01"
                Case "AEREO" : ki.gviatra = "02"
                Case "TERRESTRE" : ki.gviatra = "03"
            End Select
            ki.gdivision = TextBox23.Text
            ki.gestilo_empresa = ""
            ki.gmarcacli = "0114"
            Select Case ComboBox5.Text
                Case "PLANO" : ki.gtipo_tejido = "01"
                Case "PUNTO" : ki.gtipo_tejido = "02"
            End Select
            ki.gobserv_3 = Trim(TextBox20.Text)
            ki.gcod_color = colortela.Text
            ki.gcolor = colortela.Text
            ki.gtela = Trim(TextBox31.Text)
            'ELIMINAR OP
            Dim JJ As New vop

            Dim vg As New fop
            If Label43.Text = 1 Then
                JJ.gncom_3 = opf
                JJ.glinea_3 = TextBox13.Text.ToString().Trim()
                JJ.gcia = Label33.Text
                vg.ELIMINAR_OP2(JJ)
            End If



            ju.gfamilia = Trim(TextBox27.Text)
            ju.gSubFamilia = Trim(TextBox28.Text)
            ju.gOrigen = ki.gtipo_mercado
            ju.gcolor = Trim(colortela.Text)
            ju.gccia = Trim(Label33.Text)

            ghj = ki1.CORRELTIVO_PRODUCTO(ju)


            ki.gOd = Mid(Trim(TextBox32.Text), 1, 9)
            ki.gOdV = Mid(Trim(TextBox32.Text), 10, 1)

            j = Convert.ToInt32(DataGridView2.ColumnCount.ToString())

            For iP = 12 To j - 1

                If DataGridView2.Rows(0).Cells(iP).Value.ToString.Trim <> "" Then
                    SUMA = SUMA + 1
                End If
            Next

            For k = 1 To SUMA
                ko.gncom = opf
                Select Case k.ToString.Count
                    Case "1" : ko.gcorrel = "0" & k
                    Case "2" : ko.gcorrel = k

                End Select

                ko.gcantidad = DataGridView2.Rows(1).Cells(k).Value
                ko.gcodcom = colortela.Text
                ko.gcodtol = DataGridView2.Rows(0).Cells(k + 1).Value
                ko.gccia = Label33.Text
                ki1.insertar_qag160d(ko)
            Next
            For kh = 1 To SUMA
                mn.gncom4 = opf
                Select Case kh.ToString.Count
                    Case "1" : mn.gcorrelq = "0" & kh
                    Case "2" : mn.gcorrelq = kh

                End Select

                mn.gtalla = DataGridView2.Rows(0).Cells(kh + 12).Value
                mn.gcodtal = DataGridView2.Rows(0).Cells(kh + 1).Value
                mn.gccia = Label33.Text
                If Trim(DataGridView4.Rows(1).Cells(kh).Value) = "" Then
                    mn.gratio = "0"
                Else
                    mn.gratio = DataGridView4.Rows(1).Cells(kh).Value
                End If

                ki1.insertar_qag160c(mn)
            Next

            lñ.gncom4 = opf
            lñ.gcorrelq = "01"
            lñ.gcolor = TextBox18.Text
            lñ.gcantidad = TextBox21.Text
            lñ.gcodcom = colortela.Text
            lñ.gcolorrt = colorprenda.Text
            lñ.gccia = Label33.Text
            ki1.insertar_qag160b(lñ)


            'If Trim(TextBox13.Text).Length = 0 Then
            Dim codigo2 As String
            codigo2 = TextBox27.Text.ToString.Trim & TextBox28.Text.ToString.Trim & ki.gtipo_mercado & colortela.Text.ToString.Trim & ghj
            'Dim va As Integer
            'Dim sql102 As String = "SELECT  COUNT(linea_17) as cant FROM custom_vianny.DBO.Cag1700 where linea_17 ='" + codigo2 + "' and ccia ='" + Label33.Text + "'"
            'Dim cmd102 As New SqlCommand(sql102, conx)
            'Rsr222 = cmd102.ExecuteReader()
            'If Rsr222.Read() = True And Convert.ToInt32(Rsr222(0)) > 0 Then
            '    va = Rsr222(0)
            'End If
            'Rsr222.Close()
            'If va >= 1 Then
            Dim agregar As String = "delete from custom_vianny.DBO.cag1700 where linea_17 ='" + (TextBox13.Text.ToString.Trim) + "' and ccia ='" + Label33.Text + "'"
            ELIMINAR(agregar)

            Dim cor As Integer


            cor = Convert.ToInt32(Microsoft.VisualBasic.Right(codigo2, 6)) + 1

            Select Case cor.ToString.Count
                Case "1" : codigo = TextBox27.Text.ToString.Trim & TextBox28.Text.ToString.Trim & ki.gtipo_mercado & colortela.Text.ToString.Trim & "00000" & cor
                Case "2" : codigo = TextBox27.Text.ToString.Trim & TextBox28.Text.ToString.Trim & ki.gtipo_mercado & colortela.Text.ToString.Trim & "0000" & cor
                Case "3" : codigo = TextBox27.Text.ToString.Trim & TextBox28.Text.ToString.Trim & ki.gtipo_mercado & colortela.Text.ToString.Trim & "000" & cor
                Case "4" : codigo = TextBox27.Text.ToString.Trim & TextBox28.Text.ToString.Trim & ki.gtipo_mercado & colortela.Text.ToString.Trim & "00" & cor
                Case "5" : codigo = TextBox27.Text.ToString.Trim & TextBox28.Text.ToString.Trim & ki.gtipo_mercado & colortela.Text.ToString.Trim & "0" & cor
                Case "6" : codigo = TextBox27.Text.ToString.Trim & TextBox28.Text.ToString.Trim & ki.gtipo_mercado & colortela.Text.ToString.Trim & cor
            End Select
            If TextBox13.Text.ToString().Trim().Length = 0 Then
                hj.glinea_17 = codigo
            Else
                'MsgBox("paso")
                hj.glinea_17 = Trim(TextBox13.Text.ToString())
            End If

            hj.gnomb_17 = TextBox16.Text
            hj.gunid_17 = "UND"
            hj.gfamilia_17 = TextBox27.Text
            hj.gsubfam_17 = TextBox28.Text
            hj.gtallaje_17 = TextBox14.Text
            hj.gorigen_17 = ki.gtipo_mercado
            hj.gidcolor_17 = colortela.Text
            hj.garticulo_17 = TextBox16.Text
            hj.gdmarca_17 = "0114"
            hj.gcodprod_17 = TextBox6.Text
            hj.gncolor_17 = TextBox18.Text
            hj.gccia = Label33.Text
            ki1.insertar_cag1700(hj)





            ki.glinea_3 = codigo
            ki.gmruta7_3 = RUTAMOLDE.Text.ToString().Trim()
            ki.gmruta8_3 = ESTILOM.Text.ToString().Trim()
            ki.gscobs1_3 = LAVADOM.Text.ToString().Trim()
            ki.gscobs2_3 = TELAM.Text.ToString().Trim()
            ki.gscobs3_3 = DESCRIPM.Text.ToString().Trim()
            ki.gmodelista_3 = MODELISTA.Text.ToString().Trim()
            ki.ganalista_3 = ANALISTA.Text.ToString().Trim()

            Try
                    Button11.Enabled = False
                    ki1.insertar_op(ki)
                Catch ex As Exception
                    MessageBox.Show("Ocurrió un error al guardar los datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    Button11.Enabled = True
                End Try

                For t = 1 To SUMA
                    uo.gncom_3 = opf
                    uo.glinea_3 = ki.glinea_3
                    Select Case t.ToString.Count
                        Case "1" : uo.gcorel = "010" & t
                        Case "2" : uo.gcorel = "01" & t

                    End Select

                    uo.gcolor = TextBox29.Text.ToString.Trim & " " & DataGridView2.Rows(0).Cells(11 + t).Value
                    uo.gccia = Label33.Text
                    ki1.insertar_cag170l(uo)
                Next
                MsgBox("SE REGISTRO LA ORDEN DE PRODUCCION SON EXITO")
                Dim respuesta As DialogResult
            'IngresarNotificacion()
            'ENVIAR_CORREO(opf)

            respuesta = MessageBox.Show("QUIERES IMPRIMIR LA ORDEN DE PEDIDO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then

                    Select Case opf.Length

                        Case "1" : opf = "010000000" & "" & opf
                        Case "2" : opf = "01000000" & "" & opf
                        Case "3" : opf = "0100000" & "" & opf
                        Case "4" : opf = "010000" & "" & opf
                        Case "5" : opf = "01000" & "" & opf
                        Case "6" : opf = "0100" & "" & opf
                        Case "7" : opf = "010" & "" & opf
                        Case "8" : opf = "01" & "" & opf
                        Case "9" : opf = "0" & opf
                    End Select
                    Rpt_Op.TextBox1.Text = Label33.Text
                    Rpt_Op.TextBox2.Text = opf
                    Rpt_Op.TextBox3.Text = opf
                    Rpt_Op.TextBox4.Text = 0
                    Rpt_Op.TextBox5.Text = 1
                    Rpt_Op.Show()


                    limpiar()
                    'Me.Close()
                End If


                limpiar()
                TextBox1.Text = correlativo()
                'Me.Close()
                TextBox1.Enabled = True
                Label43.Text = "0"

            Else
                MsgBox("FALTA INGRESAR ALGUNOS CAMPOS OBLIGATORIOS")
        End If
        'End If
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
    Dim Rsr199 As SqlDataReader
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        Dim valor As Integer = 0
        Dim sql101 As String = "select count(ncom_3) from custom_vianny.dbo.qag0300  WHERE ( LEN( modelista_3) >0 or  acabados_3 >1  ) and ncom_3 ='" + TextBox1.Text.ToString().Trim() + "' and ccia ='" + Label33.Text + "'"
        Dim cmd101 As New SqlCommand(sql101, conx)
        Rsr199 = cmd101.ExecuteReader()
        If Rsr199.Read() Then
            valor = Rsr199(0)
        Else
            valor = 0
        End If
        Rsr199.Close()

        'If valor = 1 Then
        '    MsgBox("NO SE PUEDE EDITAR ESTA OP, PORQUE YA SE HAN REALIZADO VARIOS PROCESOS QUE SE MODIFICARIAN SI EDITA - CONSULTAR CON EL AREA DE SISTEMAS")
        'Else
        Label43.Text = "1"
            TextBox16.Enabled = True
            TextBox8.ReadOnly = False
            TextBox5.ReadOnly = False
            TextBox8.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox6.ReadOnly = False
            TextBox16.ReadOnly = False
            TextBox20.ReadOnly = False
            TextBox30.ReadOnly = False
            TextBox31.ReadOnly = False
            TextBox32.ReadOnly = False
            TextBox20.Enabled = True
            Button10.Enabled = True
            Button11.Enabled = True
            Button16.Enabled = True
            Button7.Enabled = True
            Button8.Enabled = True
            Button5.Enabled = True
            Button2.Enabled = True
            Button1.Enabled = True
            Button6.Enabled = True
            Button3.Enabled = True
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True
        'End If

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim FF As New fop
        Dim GG As New vop

        GG.gncom_3 = TextBox1.Text
        GG.gcia = Label33.Text
        FF.actualizar_OP(GG)
        MsgBox("LA OP SOLICITADA SE ANULO CON EXITO")
        Me.Close()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Select Case TextBox1.Text.Length

            Case "1" : TextBox1.Text = "010000000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "01000000" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "0100000" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = "010000" & "" & TextBox1.Text
            Case "5" : TextBox1.Text = "01000" & "" & TextBox1.Text
            Case "6" : TextBox1.Text = "0100" & "" & TextBox1.Text
            Case "7" : TextBox1.Text = "010" & "" & TextBox1.Text
            Case "8" : TextBox1.Text = "01" & "" & TextBox1.Text
            Case "9" : TextBox1.Text = "0" & TextBox1.Text
        End Select
        Rpt_Op.TextBox1.Text = Label33.Text
        Rpt_Op.TextBox2.Text = TextBox1.Text
        Rpt_Op.TextBox3.Text = TextBox1.Text
        Rpt_Op.TextBox4.Text = 0
        Rpt_Op.TextBox5.Text = 1
        Rpt_Op.Show()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        limpiar()
        TextBox1.Enabled = True
        TextBox5.Enabled = True
        abrir()
        TextBox1.Text = correlativo()
        TextBox1.Select()

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'DateTimePicker1.Value = DateTimePicker1.Value.AddDays(30)
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        Process.Start(Trim(TextBox88.Text))
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Process.Start(Trim(TextBox35.Text))
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Process.Start(Trim(TextBox36.Text))
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click

    End Sub

    Private Sub TextBox33_TextChanged(sender As Object, e As EventArgs) Handles TextBox33.TextChanged

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Si no es un número y no es la tecla de retroceso, entonces no permitir la entrada
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        ' Verificar si la tecla presionada es un número, un punto decimal o la tecla de retroceso
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "." AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Si no es un número, un punto decimal ni la tecla de retroceso, entonces no permitir la entrada
            e.Handled = True
        End If

        ' Verificar que no haya más de un punto decimal en el texto
        If e.KeyChar = "." AndAlso TextBox1.Text.IndexOf(".") > -1 Then
            e.Handled = True
        End If
    End Sub
    Dim Rsr2229 As SqlDataReader
    Private Sub TextBox33_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox33.KeyDown


        If e.KeyCode = Keys.Enter Then

            Dim gh As New fop
            Dim kk As New vop
            Dim lp As Integer
            kk.gncom_3 = Trim(TextBox33.Text)
            kk.gcia = "01"
            ll = gh.buscar_op(kk)
            Dim num As Integer
            Dim p1 As String
            num = 0
            DataGridView6.DataSource = ll
            lp = DataGridView6.Rows.Count
            If lp <> 0 Then
                For i = 0 To lp - 1

                    If Trim(TextBox26.Text) = Trim(DataGridView6.Rows(0).Cells(44).Value) Or Trim(MDIParent1.Label4.Text) = "ADMINISTRADOR" Then
                        TextBox22.Text = DataGridView6.Rows(0).Cells(7).Value
                        TextBox2.Text = DataGridView6.Rows(0).Cells(8).Value
                        Dim bpo As String
                        bpo = Microsoft.VisualBasic.Mid(Trim(DataGridView6.Rows(0).Cells(4).Value), 2, 5)
                        'MsgBox(bpo)
                        Dim sql102 As String = "select top 1 program_3, CAST( substring(program_3,7,4) as int) as numero,substring(program_3,1,5) from custom_vianny.dbo.qag0300  where  program_3 like  '" + bpo + "%' and ccia='01' AND SUBSTRING(ncom_3,1,2) IN (01)  order by numero desc"
                        Dim cmd102 As New SqlCommand(sql102, conx)
                        Rsr2229 = cmd102.ExecuteReader()
                        If Rsr2229.Read() Then
                            num = Convert.ToInt32(Rsr2229(1)) + 1
                            p1 = Rsr2229(2).ToString().Trim()
                        Else
                            num = 1
                            p1 = bpo
                        End If
                        Rsr2229.Close()
                        Select Case num.ToString().Length
                            Case 1 : TextBox5.Text = p1 & "0000" & num
                            Case 2 : TextBox5.Text = p1 & "000" & num
                            Case 3 : TextBox5.Text = p1 & "00" & num
                            Case 4 : TextBox5.Text = p1 & "0" & num
                            Case 5 : TextBox5.Text = p1 & num

                        End Select
                        'num = Convert.ToInt32(Microsoft.VisualBasic.Mid(Trim(DataGridView6.Rows(0).Cells(4).Value), 7, 4))
                        'Select Case num.ToString().Length
                        '    Case 1 : TextBox5.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView6.Rows(0).Cells(4).Value), 2, 5) & "0000" & num
                        '    Case 2 : TextBox5.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView6.Rows(0).Cells(4).Value), 2, 5) & "000" & num
                        '    Case 3 : TextBox5.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView6.Rows(0).Cells(4).Value), 2, 5) & "00" & num
                        '    Case 4 : TextBox5.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView6.Rows(0).Cells(4).Value), 2, 5) & "0" & num
                        '    Case 5 : TextBox5.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView6.Rows(0).Cells(4).Value), 2, 5) & num

                        'End Select
                        'TextBox5.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView6.Rows(0).Cells(4).Value), 2, 5)
                        TextBox3.Text = DataGridView6.Rows(0).Cells(10).Value
                        TextBox4.Text = DataGridView6.Rows(0).Cells(11).Value
                        TextBox6.Text = DataGridView6.Rows(0).Cells(81).Value
                        TextBox26.Text = DataGridView6.Rows(0).Cells(44).Value
                        TextBox8.Text = DataGridView6.Rows(0).Cells(45).Value
                        TextBox8.Text = DataGridView6.Rows(0).Cells(45).Value
                        TextBox27.Text = DataGridView6.Rows(0).Cells(43).Value
                        TextBox28.Text = DataGridView6.Rows(0).Cells(52).Value
                        TextBox12.Text = DataGridView6.Rows(0).Cells(53).Value
                        TextBox32.Text = TextBox33.Text
                        Select Case TextBox26.Text
                            Case "0001" : TextBox10.Text = "VSILVERIO"
                            Case "0003" : TextBox10.Text = "GBALVIN"
                            Case "0025" : TextBox10.Text = "VGRAUS"
                        End Select
                        TextBox8.Text = DataGridView6.Rows(0).Cells(45).Value
                        TextBox23.Text = DataGridView6.Rows(0).Cells(79).Value
                        TextBox24.Text = DataGridView6.Rows(0).Cells(80).Value
                        TextBox31.Text = DataGridView6.Rows(0).Cells(82).Value
                        TextBox30.Text = Trim(DataGridView6.Rows(0).Cells(83).Value) & Trim(DataGridView6.Rows(0).Cells(90).Value) & Trim(DataGridView6.Rows(0).Cells(91).Value)
                        TextBox16.Text = Trim(DataGridView6.Rows(0).Cells(17).Value)
                        TextBox34.Text = DataGridView6.Rows(0).Cells(88).Value




                        TextBox88.Text = Trim(DataGridView6.Rows(0).Cells(86).Value)

                        If Trim(DataGridView6.Rows(0).Cells(54).Value) = "N" Then
                            ComboBox3.SelectedIndex = 0
                        Else
                            ComboBox3.SelectedIndex = 1
                        End If
                    End If
                Next
                TextBox33.Text = ""
            End If
        End If
        If e.KeyCode = Keys.F1 Then
            Copiar_Op.Label1.Text = "OP"
            Copiar_Op.ShowDialog()
        End If
    End Sub

    Private Sub TextBox33_Invalidated(sender As Object, e As InvalidateEventArgs) Handles TextBox33.Invalidated

    End Sub
End Class