Imports System.ComponentModel
Imports System.Data.SqlClient
Public Class Letra
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            Factura_Letra.ShowDialog()
        End If
    End Sub
    Public Function CantidadALetras(ByVal XNumero As Decimal, ByVal XMoneda As String) As String
        Dim XlnEntero As Integer
        Dim XlcRetorno As String
        Dim XlnTerna As Integer
        Dim XlcCadena As String
        Dim XlnUnidades As Integer
        Dim XlnDecenas As Integer
        Dim XlnCentenas As Integer
        Dim XlnFraccion As Integer
        Dim Xresultado As String

        XlnEntero = Math.Truncate(XNumero)
        XlnFraccion = (XNumero - XlnEntero) * 100

        XlcRetorno = ""
        XlnTerna = 1

        While (XlnEntero > 0)

            'Recorro terna por terna
            XlcCadena = ""
            XlnUnidades = XlnEntero Mod 10
            XlnEntero = Math.Floor(XlnEntero / 10)
            XlnDecenas = XlnEntero Mod 10
            XlnEntero = Math.Floor(XlnEntero / 10)
            XlnCentenas = XlnEntero Mod 10
            XlnEntero = Math.Floor(XlnEntero / 10)

            'Analizo las unidades
            Select Case True ' UNIDADES
                Case XlnUnidades = 1 And XlnTerna = 1
                    XlcCadena = "UN " & XlcCadena
                Case XlnUnidades = 1 And XlnTerna <> 1
                    XlcCadena = XlcCadena
                Case XlnUnidades = 2
                    XlcCadena = "DOS " & XlcCadena
                Case XlnUnidades = 3
                    XlcCadena = "TRES " & XlcCadena
                Case XlnUnidades = 4
                    XlcCadena = "CUATRO " & XlcCadena
                Case XlnUnidades = 5
                    XlcCadena = "CINCO " & XlcCadena
                Case XlnUnidades = 6
                    XlcCadena = "SEIS " & XlcCadena
                Case XlnUnidades = 7
                    XlcCadena = "SIETE " & XlcCadena
                Case XlnUnidades = 8
                    XlcCadena = "OCHO " & XlcCadena
                Case XlnUnidades = 9
                    XlcCadena = "NUEVE " & XlcCadena
                Case Else
                    XlcCadena = XlcCadena
            End Select 'UNIDADES

            'Analizo las decenas
            Select Case True 'DECENAS
                Case XlnDecenas = 1
                    Select Case XlnUnidades
                        Case 0
                            XlcCadena = "DIEZ "
                        Case 1
                            XlcCadena = "ONCE "
                        Case 2
                            XlcCadena = "DOCE "
                        Case 3
                            XlcCadena = "TRECE "
                        Case 4
                            XlcCadena = "CATORCE "
                        Case 5
                            XlcCadena = "QUINCE"
                        Case Else
                            XlcCadena = "DIECI" & XlcCadena
                    End Select
                Case XlnDecenas = 2 And XlnUnidades = 0
                    XlcCadena = "VEINTE " & XlcCadena
                Case XlnDecenas = 2 And XlnUnidades <> 0
                    XlcCadena = "VEINTI" & XlcCadena
                Case XlnDecenas = 3 And XlnUnidades = 0
                    XlcCadena = "TREINTA " & XlcCadena
                Case XlnDecenas = 3 And XlnUnidades <> 0
                    XlcCadena = "TREINTA Y " & XlcCadena
                Case XlnDecenas = 4 And XlnUnidades = 0
                    XlcCadena = "CUARENTA " & XlcCadena
                Case XlnDecenas = 4 And XlnUnidades <> 0
                    XlcCadena = "CUARENTA Y " & XlcCadena
                Case XlnDecenas = 5 And XlnUnidades = 0
                    XlcCadena = "CINCUENTA " & XlcCadena
                Case XlnDecenas = 5 And XlnUnidades <> 0
                    XlcCadena = "CINCUENTA Y " & XlcCadena
                Case XlnDecenas = 6 And XlnUnidades = 0
                    XlcCadena = "SESENTA " & XlcCadena
                Case XlnDecenas = 6 And XlnUnidades <> 0
                    XlcCadena = "SESENTA Y " & XlcCadena
                Case XlnDecenas = 7 And XlnUnidades = 0
                    XlcCadena = "SETENTA " & XlcCadena
                Case XlnDecenas = 7 And XlnUnidades <> 0
                    XlcCadena = "SETENTA Y " & XlcCadena
                Case XlnDecenas = 8 And XlnUnidades = 0
                    XlcCadena = "OCHENTA " & XlcCadena
                Case XlnDecenas = 8 And XlnUnidades <> 0
                    XlcCadena = "OCHENTA Y " & XlcCadena
                Case XlnDecenas = 9 And XlnUnidades = 0
                    XlcCadena = "NOVENTA " & XlcCadena
                Case XlnDecenas = 9 And XlnUnidades <> 0
                    XlcCadena = "NOVENTA Y " & XlcCadena
                Case Else
                    XlcCadena = XlcCadena
            End Select 'DECENAS

            ' Analizo las centenas
            Select Case True ' CENTENAS
                Case XlnCentenas = 1 And XlnUnidades = 0 And XlnDecenas = 0
                    XlcCadena = "CIEN " & XlcCadena
                Case XlnCentenas = 1 And Not (XlnUnidades = 0 And XlnDecenas = 0)
                    XlcCadena = "CIENTO " & XlcCadena
                Case XlnCentenas = 2
                    XlcCadena = "DOSCIENTOS " & XlcCadena
                Case XlnCentenas = 3
                    XlcCadena = "TRESCIENTOS " & XlcCadena
                Case XlnCentenas = 4
                    XlcCadena = "CUATROCIENTOS " & XlcCadena
                Case XlnCentenas = 5
                    XlcCadena = "QUINIENTOS " & XlcCadena
                Case XlnCentenas = 6
                    XlcCadena = "SEISCIENTOS " & XlcCadena
                Case XlnCentenas = 7
                    XlcCadena = "SETECIENTOS " & XlcCadena
                Case XlnCentenas = 8
                    XlcCadena = "OCHOCIENTOS " & XlcCadena
                Case XlnCentenas = 9
                    XlcCadena = "NOVECIENTOS " & XlcCadena
                Case Else
                    XlcCadena = XlcCadena
            End Select 'CENTENAS

            ' Analizo la terna
            Select Case True ' TERNA
                Case XlnTerna = 1
                    XlcCadena = XlcCadena
                Case XlnTerna = 2 And (XlnUnidades + XlnDecenas + XlnCentenas <> 0)
                    XlcCadena = XlcCadena & "MIL "
                Case XlnTerna = 3 And (XlnUnidades + XlnDecenas + XlnCentenas <> 0) And XlnUnidades = 1 And XlnDecenas = 0 And XlnCentenas = 0
                    XlcCadena = XlcCadena & "MILLON "
                Case XlnTerna = 3 And (XlnUnidades + XlnDecenas + XlnCentenas <> 0) And Not (XlnUnidades = 1 And XlnDecenas = 0 And XlnCentenas = 0)
                    XlcCadena = XlcCadena & "MILLONES "
                Case XlnTerna = 4 And (XlnUnidades + XlnDecenas + XlnCentenas <> 0)
                    XlcCadena = XlcCadena & "MIL MILLONES "
                Case XlnTerna = 5 And (XlnUnidades + XlnDecenas + XlnCentenas <> 0)
                    XlcCadena = XlcCadena & "BILLON "
                Case XlnTerna = 6 And (XlnUnidades + XlnDecenas + XlnCentenas <> 0)
                    XlcCadena = XlcCadena & "BILLONES "
                Case Else
                    XlcCadena = ""
            End Select 'TERNA

            'Armo el retorno terna a terna
            XlcRetorno = XlcCadena & XlcRetorno
            XlnTerna = XlnTerna + 1
        End While ' While

        If XlnTerna = 1 Then XlcRetorno = "CERO"
        If ComboBox1.Text = "S/" Then
            If XlnFraccion = 0 Then
                Xresultado = XlcRetorno.ToString.TrimEnd.ToLower & " con " & XlnFraccion.ToString.TrimStart & "0/100 SOLES"
            Else
                Xresultado = XlcRetorno.ToString.TrimEnd.ToLower & " con " & XlnFraccion.ToString.TrimStart & "/100 SOLES"
            End If
        Else
            If XlnFraccion = 0 Then
                Xresultado = XlcRetorno.ToString.TrimEnd.ToLower & " con " & XlnFraccion.ToString.TrimStart & "0/100 DOLARES"
            Else
                Xresultado = XlcRetorno.ToString.TrimEnd.ToLower & " con " & XlnFraccion.ToString.TrimStart & "/100 DOLARES"
            End If
        End If



        Return Xresultado.ToUpper()
    End Function

    Private Sub Letra_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox19.Enabled = False
        TextBox2.Select()
        ComboBox1.SelectedItem = 0
        CORRELATIVO()
    End Sub
    Dim Rsr21 As SqlDataReader
    Sub CORRELATIVO()
        abrir()
        Dim va As String
        Dim sql102 As String = "Select top 1 NumLet,substring(cast(NumLet as char(6)),1,2)   from  Letra where NumLet like   '" + Trim(TextBox6.Text) + "' +'%' order by NumLet desc "
        Dim cmd1021 As New SqlCommand(sql102, conx)
        Rsr21 = cmd1021.ExecuteReader()

        If Rsr21.Read() = True Then
            TextBox1.Text = Mid(Rsr21(0), 3, 4) + 1
            Select Case TextBox1.Text.Length
                Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = TextBox1.Text
            End Select
        Else
            TextBox1.Text = 1
            Select Case TextBox1.Text.Length
                Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = TextBox1.Text
            End Select
        End If
        Rsr21.Close()

    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged


        If Label21.Text = 1 Then
                Dim mon As String
            If ComboBox1.Text = "S/" Then
                mon = "SOLES"
            Else
                mon = "DOLARES"
            End If
                Dim numeroEnLetras As String = CantidadALetras(Convert.ToDouble(TextBox4.Text), mon)
                TextBox5.Text = numeroEnLetras
                If Trim(TextBox2.Text) = "" Then
                    TextBox4.Text = ""
                    MsgBox("FALTA INGRESAR EL DOCUMENTO DE REFERENCIA ")
                End If
            End If



    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clients_Prove.Label1.Text = "1"
            Clients_Prove.ShowDialog()
        End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clients_Prove.Label1.Text = "2"
            Clients_Prove.ShowDialog()
        End If
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox19.Enabled = True
            TextBox19.Select()
        Else
            TextBox19.Enabled = False
        End If

    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged

    End Sub

    Private Sub TextBox18_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox18.KeyPress
        ' Verificar si el carácter presionado es un número o una tecla de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Si no es un número ni una tecla de control, suprimir el carácter ingresado
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        ' Verificar si el carácter presionado es un número o una tecla de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Si no es un número ni una tecla de control, suprimir el carácter ingresado
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox19_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox19.KeyPress
        ' Verificar si el carácter presionado es un número o una tecla de control
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            ' Si no es un número ni una tecla de control, suprimir el carácter ingresado
            e.Handled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim camposFaltantes As New List(Of String) ' Lista para almacenar los nombres de los campos faltantes

        ' Validar campos obligatorios
        If String.IsNullOrEmpty(TextBox1.Text) Then
            camposFaltantes.Add(" Numero de Letra ")
        End If

        If String.IsNullOrEmpty(TextBox2.Text) Then
            camposFaltantes.Add(" Documento de Referencia ")
        End If

        If String.IsNullOrEmpty(TextBox18.Text) Then
            camposFaltantes.Add(" Dias de la Letra ")
        End If

        If String.IsNullOrEmpty(TextBox4.Text) Then
            camposFaltantes.Add(" Importe ")
        End If
        If String.IsNullOrEmpty(TextBox11.Text) Then
            camposFaltantes.Add(" Ruc Aceptante ")
        End If
        'If String.IsNullOrEmpty(TextBox14.Text) Then
        '    camposFaltantes.Add(" Ruc Aval ")
        'End If
        ' Agregar más validaciones para otros campos obligatorios según sea necesario

        ' Verificar si hay campos faltantes
        If camposFaltantes.Count > 0 Then
            Dim camposFaltantesStr As String = String.Join(", ", camposFaltantes)
            MessageBox.Show("Falta ingresar información en los siguientes campos: " & camposFaltantesStr, "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim i2 As Integer
            Dim CodIni, CodFin, Num As String
            Label21.Text = 2
            If CheckBox1.Checked = True Then
                CodIni = TextBox6.Text & TextBox1.Text
                For i2 = 0 To TextBox19.Text - 1
                    guardar()
                    CORRELATIVO()

                Next
                MsgBox("SE REGISTRO LAS  " & TextBox19.Text & " LETRAS CORRECTAMENTE")


                Num = Convert.ToInt32(TextBox1.Text) + Convert.ToInt32(TextBox19.Text) - 1
                Select Case Num.Length
                    Case "1" : CodFin = TextBox6.Text & "000" & "" & Num
                    Case "2" : CodFin = TextBox6.Text & "00" & "" & Num
                    Case "3" : CodFin = TextBox6.Text & "0" & "" & Num
                    Case "4" : CodFin = TextBox6.Text & Num
                End Select
                'MsgBox(CodIni)
                'MsgBox(CodFin)
                Reporte_Letra.TextBox1.Text = CodIni
                Reporte_Letra.TextBox2.Text = CodFin
                Reporte_Letra.Show()
                LIMPIAR()
                CORRELATIVO()

            Else
                guardar()
                MsgBox("SE REGISTRO LA LETRA CORRECTAMENTE")
                Dim respuesta As DialogResult
                respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    Reporte_Letra.TextBox1.Text = TextBox6.Text & TextBox1.Text
                    Reporte_Letra.TextBox2.Text = TextBox6.Text & TextBox1.Text
                    Reporte_Letra.Show()
                End If
                LIMPIAR()
                CORRELATIVO()
            End If

        End If
    End Sub
    Sub guardar()
        abrir()
        Dim cmd15 As New SqlCommand("Insert into Letra (NumLet,DoRLet,FegLet,DiaLet,LugLet,FeVLet,MonLet,ImpLet,EstLet,ImLLet,RenLet,ReDLet,Empresa)
                                             values (@NumLet,@DoRLet,@FegLet,@DiaLet,@LugLet,@FeVLet,@MonLet,@ImpLet,@EstLet,@ImLLet,@RenLet,@ReDLet,@Empresa)", conx)
        cmd15.Parameters.AddWithValue("@NumLet", TextBox6.Text & TextBox1.Text)
        cmd15.Parameters.AddWithValue("@DoRLet", TextBox2.Text)
        cmd15.Parameters.AddWithValue("@FegLet", DateTimePicker1.Value)
        cmd15.Parameters.AddWithValue("@DiaLet", TextBox18.Text)
        cmd15.Parameters.AddWithValue("@LugLet", TextBox3.Text)
        cmd15.Parameters.AddWithValue("@FeVLet", DateTimePicker2.Value)
        cmd15.Parameters.AddWithValue("@MonLet", ComboBox1.Text)
        cmd15.Parameters.AddWithValue("@ImpLet", Convert.ToDouble(TextBox4.Text))
        cmd15.Parameters.AddWithValue("@EstLet", "1")
        cmd15.Parameters.AddWithValue("@ImLLet", Trim(TextBox5.Text))
        cmd15.Parameters.AddWithValue("@RenLet", "0")
        cmd15.Parameters.AddWithValue("@ReDLet", "0")
        cmd15.Parameters.AddWithValue("@Empresa", Trim(TextBox7.Text))
        cmd15.ExecuteNonQuery()

        Dim cmd16 As New SqlCommand("insert into Aceptante (idAce,NomAce,DomAce,LocAce,RucAce,TelAce) values (@idAce,@NomAce,@DomAce,@LocAce,@RucAce,@TelAce)", conx)
        cmd16.Parameters.AddWithValue("@idAce", TextBox6.Text & TextBox1.Text)
        cmd16.Parameters.AddWithValue("@NomAce", Trim(TextBox8.Text))
        cmd16.Parameters.AddWithValue("@DomAce", Trim(TextBox9.Text))
        cmd16.Parameters.AddWithValue("@LocAce", Trim(TextBox10.Text))
        cmd16.Parameters.AddWithValue("@RucAce", Trim(TextBox11.Text))
        cmd16.Parameters.AddWithValue("@TelAce", TextBox12.Text)
        cmd16.ExecuteNonQuery()
        Dim cmd17 As New SqlCommand("insert into Aval(idAva,NomAva,DomAav,LocAva,RucAva,TelAva) values (@idAva,@NomAva,@DomAav,@LocAva,@RucAva,@TelAva)", conx)
        cmd17.Parameters.AddWithValue("@idAva", TextBox6.Text & TextBox1.Text)
        cmd17.Parameters.AddWithValue("@NomAva", Trim(TextBox17.Text))
        cmd17.Parameters.AddWithValue("@DomAav", Trim(TextBox16.Text))
        cmd17.Parameters.AddWithValue("@LocAva", Trim(TextBox15.Text))
        cmd17.Parameters.AddWithValue("@RucAva", Trim(TextBox14.Text))
        cmd17.Parameters.AddWithValue("@TelAva", TextBox13.Text)
        cmd17.ExecuteNonQuery()
        Dim cmd18 As New SqlCommand("insert into Aval2(idAva2,NomAva2,DomAav2,LocAva2,RucAva2,TelAva2) values (@idAva1,@NomAva1,@DomAav1,@LocAva1,@RucAva1,@TelAva1)", conx)
        cmd18.Parameters.AddWithValue("@idAva1", TextBox6.Text & TextBox1.Text)
        cmd18.Parameters.AddWithValue("@NomAva1", Trim(TextBox24.Text))
        cmd18.Parameters.AddWithValue("@DomAav1", Trim(TextBox23.Text))
        cmd18.Parameters.AddWithValue("@LocAva1", Trim(TextBox22.Text))
        cmd18.Parameters.AddWithValue("@RucAva1", Trim(TextBox21.Text))
        cmd18.Parameters.AddWithValue("@TelAva1", TextBox20.Text)
        cmd18.ExecuteNonQuery()


    End Sub
    Sub LIMPIAR()
        TextBox2.Text = ""
        TextBox18.Text = ""

        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox4.Text = ""
        TextBox16.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox13.Text = ""
        CheckBox1.Checked = False
        TextBox19.Text = ""
        TextBox19.Enabled = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Reporte_Letra.TextBox1.Text = "230382"
        Reporte_Letra.TextBox2.Text = "230384"
        Reporte_Letra.Show()
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Clonar_letra.Label3.Text = 3
        Clonar_letra.Show()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox18_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox18.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
            DateTimePicker2.Value = DateTimePicker1.Value.AddDays(Trim(TextBox18.Text))
            TextBox11.Select()
        End If
    End Sub
End Class