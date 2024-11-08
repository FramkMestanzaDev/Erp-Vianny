Imports System.Data.SqlClient
Imports System.Globalization

Public Class Form10
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
    Dim Rsr21 As SqlDataReader

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Dim dg1 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = ""
        dg1.Clear()

        If Trim(TextBox1.Text) <> "" And Trim(TextBox2.Text) <> "" Then
            abrir()


            Dim cmd As New SqlDataAdapter("select NumLet AS LETRA,FegLet as FECREGISTRO,'' AS DIAS  ,FeVLet AS FECVENC,ImpLet AS IMPORTE,ImLLet AS IMPLETRA,MonLet AS MOENDA from Letra where NumLet between '" + Trim(TextBox1.Text) + "' and '" + Trim(TextBox2.Text) + "' order by NumLet", conx)

            cmd.Fill(dg1)
            If dg1.Rows.Count <> 0 Then
                DataGridView1.DataSource = dg1
            Else
                DataGridView1.DataSource = ""
            End If
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).Width = 100
            DataGridView1.Columns(2).Width = 65
            DataGridView1.Columns(4).Width = 80
            DataGridView1.Columns(5).Width = 470
            DataGridView1.Columns(6).Width = 70
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(1).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView1.Columns(3).DefaultCellStyle.Format = "dd/MM/yyyy"

        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "IMPORTE" Then
            Dim mon As String
            If Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value) = "S/" Then
                mon = "SOLES"
            Else
                mon = "DOLARES"
            End If
            Dim numeroEnLetras As String = CantidadALetras(Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(4).Value), mon)
            DataGridView1.Rows(e.RowIndex).Cells(5).Value = numeroEnLetras
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "DIAS" Then
            Dim FECHA As Date
            FECHA = Date.Parse(DataGridView1.Rows(e.RowIndex).Cells(1).Value)

            DataGridView1.Rows(e.RowIndex).Cells(3).Value = FECHA.AddDays(Convert.ToInt16(DataGridView1.Rows(e.RowIndex).Cells(2).Value))
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        abrir()
        Dim I1 As Integer
        I1 = DataGridView1.Rows.Count
        Dim FEC As String
        For I = 0 To I1 - 1


            Dim cmd15 As New SqlCommand("UPDATE Letra SET FeVLet =@FeVLet,ImpLet=@ImpLet,ImLLet =@ImLLet WHERE NumLet =@NumLet", conx)
            cmd15.Parameters.AddWithValue("@FeVLet", (DataGridView1.Rows(I).Cells(3).Value))
            cmd15.Parameters.AddWithValue("@ImpLet", Trim(DataGridView1.Rows(I).Cells(4).Value))
            cmd15.Parameters.AddWithValue("@ImLLet", Trim(DataGridView1.Rows(I).Cells(5).Value))
            cmd15.Parameters.AddWithValue("@NumLet", Trim(DataGridView1.Rows(I).Cells(0).Value))
            cmd15.ExecuteNonQuery()
        Next
        MsgBox("SE ACTUALIZO LA INFORMACION CORRECTAMENTE")
        DataGridView1.DataSource = ""

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
        If Trim(XMoneda) = "SOLES" Then
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        If e.ColumnIndex = 1 Then
            Dim valor As String = e.FormattedValue.ToString()

            ' Define la longitud máxima permitida (por ejemplo, 10 caracteres)
            Dim longitudMaxima As Integer = 10

            ' Valida la longitud del texto
            If valor.Length > longitudMaxima Then
                MessageBox.Show("La longitud del texto no puede ser mayor de " & longitudMaxima & " caracteres.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                DataGridView1.CancelEdit() ' Cancela la edición
            End If
            Dim valor1 As String = e.FormattedValue.ToString()

            ' Intenta analizar el valor ingresado como fecha
            Dim fecha As DateTime
            If Not DateTime.TryParse(valor1, fecha) Then
                MessageBox.Show("El valor ingresado no es una fecha válida.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True ' Cancela la edición
            End If
            Dim valor2 As String = e.FormattedValue.ToString()

            ' Define el formato de fecha deseado
            Dim formatoFecha As String = "dd/MM/yyyy"

            ' Intenta analizar el valor ingresado como fecha en el formato deseado
            Dim fecha2 As DateTime
            If Not DateTime.TryParseExact(valor2, formatoFecha, CultureInfo.InvariantCulture, DateTimeStyles.None, fecha2) Then
                MessageBox.Show("El valor ingresado no cumple con el formato de fecha deseado (año/mes/día).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True ' Cancela la edición
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clonar_letra.Label3.Text = 1
            Clonar_letra.Show()

        End If
    End Sub

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clonar_letra.Label3.Text = 2
            Clonar_letra.Show()

        End If
    End Sub
End Class