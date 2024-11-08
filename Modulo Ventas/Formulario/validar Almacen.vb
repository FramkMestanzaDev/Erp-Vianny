Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class validar_Almacen
    Dim dt, dt2 As New DataTable
    Public conx As SqlConnection
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
    Private Sub validar_Almacen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim gf As New fnopedido
        Dim gf1 As New vnapedido
        gf1.gperiodo = Label6.Text
        gf1.gempresa = Label7.Text
        dt = gf.validar_almacen(gf1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            Dim k As Integer
            k = DataGridView1.Rows.Count

            For i = 0 To k - 1

                Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                    Case "10" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                    Case "37" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                    Case "38" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                    Case "51" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                    Case "42" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                    Case "60" : DataGridView1.Rows(i).Cells(10).Value = "SERVICIO CHILCA"
                    Case "67" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"

                End Select
            Next
            DataGridView1.Columns.Item(9).Visible = False
            DataGridView1.Columns.Item(9).Visible = False
            DataGridView1.Columns.Item(11).Visible = False
            DataGridView1.Columns.Item(5).Width = 262
            If TextBox3.Text = 1 Then
                DataGridView1.Columns.Item(0).Visible = False
                Button2.Visible = False
            End If
        End If

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "RAZON_SOCIAL" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim k As Integer
                k = DataGridView1.Rows.Count

                For i = 0 To k - 1

                    Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                        Case "10" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "37" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                        Case "38" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                        Case "51" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "42" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "60" : DataGridView1.Rows(i).Cells(10).Value = "SERVICIO CHILCA"
                        Case "67" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"

                    End Select
                Next
                DataGridView1.Columns.Item(9).Visible = False
                DataGridView1.Columns.Item(5).Width = 262
                DataGridView1.Columns.Item(11).Visible = False
                DataGridView1.Columns.Item(5).Width = 262
                If TextBox3.Text = 1 Then
                    DataGridView1.Columns.Item(0).Visible = False
                    Button2.Visible = False
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PEDIDO" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim k As Integer
                k = DataGridView1.Rows.Count

                For i = 0 To k - 1

                    Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                        Case "10" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "37" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                        Case "38" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                        Case "51" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "42" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "60" : DataGridView1.Rows(i).Cells(10).Value = "SERVICIO CHILCA"
                        Case "67" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                    End Select
                Next
                DataGridView1.Columns.Item(9).Visible = False
                DataGridView1.Columns.Item(5).Width = 262
                DataGridView1.Columns.Item(11).Visible = False
                DataGridView1.Columns.Item(5).Width = 262
                If TextBox3.Text = 1 Then
                    DataGridView1.Columns.Item(0).Visible = False
                    Button2.Visible = False
                End If
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Dim jkk As New DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim J As Integer

        J = DataGridView1.Rows.Count

        For i = 0 To J - 1

            If DataGridView1.Rows(i).Cells(0).Value = True Then
                Dim ju As New fnopedido
                Dim ju1 As New vnapedido

                ju1.gnumero_pedido = DataGridView1.Rows(i).Cells(2).Value
                ju1.gempresa = Label7.Text
                ju.actualizar_estado_almacen(ju1)

                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "PEDIDO-COMERCIAL"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = 8 'VALIDO ALMACEN
                hj1.gpc = My.Computer.Name
                hj1.gfecha = String.Format("{0:G}", DateTime.Now)
                hj1.gdetalle = Trim(DataGridView1.Rows(i).Cells(2).Value)
                hj1.gccia = Label7.Text
                hj2.insertar_log(hj1)
                Dim message As New MailMessage
                Dim smtp As New SmtpClient
                Dim fk As New fnopedido
                Dim jj As New vnapedido
                Dim corre As String
                jj.gvendedor = Trim(DataGridView1.Rows(i).Cells(11).Value)

                corre = fk.buscar_correo(jj)
                abrir()
                Dim Rs As SqlDataReader
                Dim sql As String = "select correo from lista_correos where formulario ='PEDIDOS_COMERCIAL   '"
                Dim cmd As New SqlCommand(sql, conx)
                Rs = cmd.ExecuteReader()
                Rs.Read()
                Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA NOTA DE PEDIDO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT>VALIDAD POR EL AREA DE ALMACEN<br/>
<FONT COLOR='blue'>* NOTA DE PEDIDO : </FONT>" + Trim(DataGridView1.Rows(i).Cells(2).Value) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(DataGridView1.Rows(i).Cells(7).Value) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(Label8.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(DataGridView1.Rows(i).Cells(5).Value) + "<br/>
<FONT COLOR='blue'>* USUARIO QUE APROBO : </FONT>" + Trim(Label9.Text) + "<br/>

</body>

</html>"
                With message
                    .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
                    .To.Add(Rs(0) & "," & corre)
                    .IsBodyHtml = True
                    .Body = Mailz
                    .Subject = "Nota Pedido N°" & Trim(DataGridView1.Rows(i).Cells(2).Value)
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
                abrir()
                Dim cmd200 As New SqlCommand("update pedido_separado SET estado = 0 where numero_pedido =@numero_pedido AND EMPRESA = @EMPRESA", conx)
                cmd200.Parameters.AddWithValue("@numero_pedido", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd200.Parameters.AddWithValue("@EMPRESA", Trim(Label7.Text))
                cmd200.ExecuteNonQuery()



            End If

        Next

        MsgBox("SE ACTUALIZO LOS DATOS SOLICITADOS")
        Dim gf As New fnopedido
        Dim gf1 As New vnapedido


        gf1.gperiodo = Label6.Text
        gf1.gempresa = Label7.Text
        jkk = gf.validar_almacen(gf1)


        DataGridView1.DataSource = jkk
        Dim k As Integer
        k = DataGridView1.Rows.Count

        For i = 0 To k - 1

            Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                Case "10" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                Case "37" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                Case "38" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                Case "51" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                Case "42" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                Case "60" : DataGridView1.Rows(i).Cells(10).Value = "SERVICIO CHILCA"
            End Select
        Next
        DataGridView1.Columns.Item(9).Visible = False
        DataGridView1.Columns.Item(5).Width = 262
        DataGridView1.Columns.Item(11).Visible = False
        If TextBox3.Text = 1 Then
            DataGridView1.Columns.Item(0).Visible = False
            Button2.Visible = False
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If RadioButton1.Checked = True Then
            Dim gf As New fnopedido
            Dim gf1 As New vnapedido

            gf1.gperiodo = Label6.Text
            gf1.galmacen = "10"
            dt = gf.validar_almacen2(gf1)
            If dt.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt
                Dim k As Integer
                k = DataGridView1.Rows.Count

                For i = 0 To k - 1

                    Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                        Case "10" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "37" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                        Case "38" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                        Case "51" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "42" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "60" : DataGridView1.Rows(i).Cells(10).Value = "SERVICIO CHILCA"
                        Case "68" : DataGridView1.Rows(i).Cells(10).Value = "HUACHIPA"
                    End Select
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                Next
                DataGridView1.Columns.Item(0).Visible = False
                DataGridView1.Columns.Item(9).Visible = False
                DataGridView1.Columns.Item(5).Width = 262
                DataGridView1.Columns.Item(11).Visible = False
            Else
                DataGridView1.DataSource = ""
            End If

        End If
        If Label7.Text = "01" Then
            If RadioButton2.Checked = True Then
                Dim gf As New fnopedido
                Dim gf1 As New vnapedido
                gf1.gperiodo = Label6.Text
                gf1.galmacen = "68"
                dt = gf.validar_almacen2(gf1)
                If dt.Rows.Count <> 0 Then
                    DataGridView1.DataSource = dt
                    Dim k As Integer
                    k = DataGridView1.Rows.Count

                    For i = 0 To k - 1

                        Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                            Case "10" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                            Case "68" : DataGridView1.Rows(i).Cells(10).Value = "HUACHIPA"
                            Case "38" : DataGridView1.Rows(i).Cells(10).Value = "HUACHIPA"
                            Case "51" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                            Case "42" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                            Case "60" : DataGridView1.Rows(i).Cells(10).Value = "SERVICIO CHILCA"
                            Case "68" : DataGridView1.Rows(i).Cells(10).Value = "HUACHIPA"
                        End Select
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    Next
                    DataGridView1.Columns.Item(0).Visible = False
                    DataGridView1.Columns.Item(9).Visible = False
                    DataGridView1.Columns.Item(5).Width = 262
                    'DataGridView1.Columns.Item(11).Visible = False
                    If TextBox3.Text = 1 Then
                        DataGridView1.Columns.Item(0).Visible = False
                        Button2.Visible = False
                    End If
                Else
                    DataGridView1.DataSource = ""
                End If

            End If

        Else
            If RadioButton2.Checked = True Then
                Dim gf As New fnopedido
                Dim gf1 As New vnapedido
                gf1.gperiodo = Label6.Text
                gf1.galmacen = "67"
                dt = gf.validar_almacen2(gf1)
                If dt.Rows.Count <> 0 Then
                    DataGridView1.DataSource = dt
                    Dim k As Integer
                    k = DataGridView1.Rows.Count

                    For i = 0 To k - 1

                        Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                            Case "67" : DataGridView1.Rows(i).Cells(10).Value = "HUACHIPA"

                        End Select
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    Next
                    DataGridView1.Columns.Item(0).Visible = False
                    DataGridView1.Columns.Item(9).Visible = False
                    DataGridView1.Columns.Item(5).Width = 262
                    'DataGridView1.Columns.Item(11).Visible = False
                    If TextBox3.Text = 1 Then
                        DataGridView1.Columns.Item(0).Visible = False
                        Button2.Visible = False
                    End If
                Else
                    DataGridView1.DataSource = ""
                End If

            End If

        End If


    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 1 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            REPORTE_NP.TextBox1.Text = DataGridView1.Rows(num1).Cells(2).Value
            REPORTE_NP.TextBox2.Text = Label7.Text
            REPORTE_NP.Show()


        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar2()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim gf As New fnopedido
        Dim gf1 As New vnapedido
        gf1.gperiodo = Label6.Text
        gf1.gempresa = Label7.Text
        dt = gf.validar_almacen(gf1)

        DataGridView1.DataSource = dt
        Dim k As Integer
        k = DataGridView1.Rows.Count

        For i = 0 To k - 1

            Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                Case "10" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                Case "37" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                Case "38" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                Case "51" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                Case "42" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                Case "60" : DataGridView1.Rows(i).Cells(10).Value = "SERVICIO CHILCA"
                Case "67" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
            End Select
        Next
        DataGridView1.Columns.Item(9).Visible = False
        DataGridView1.Columns.Item(5).Width = 262
        DataGridView1.Columns.Item(11).Visible = False
        If TextBox3.Text = 1 Then
            DataGridView1.Columns.Item(0).Visible = False
            Button2.Visible = False
        Else
            DataGridView1.Columns.Item(0).Visible = True
            Button2.Visible = True
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If RadioButton3.Checked = True Then
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "ALMACEN   = '10' OR ALMACEN   = '51' OR ALMACEN   = '42' OR ALMACEN   = '60' "

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim k As Integer
                k = DataGridView1.Rows.Count

                For i = 0 To k - 1

                    Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                        Case "10" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "37" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                        Case "38" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                        Case "51" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "42" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                        Case "60" : DataGridView1.Rows(i).Cells(10).Value = "SERVICIO CHILCA"
                    End Select
                Next
                DataGridView1.Columns.Item(9).Visible = False
                DataGridView1.Columns.Item(5).Width = 262

            Else

                DataGridView1.DataSource = Nothing
            End If

        Else
            If RadioButton4.Checked = True Then
                Dim ds As New DataSet
                ds.Tables.Add(dt.Copy)
                Dim dv As New DataView(ds.Tables(0))


                dv.RowFilter = "ALMACEN   = '37' OR ALMACEN   = '38' OR ALMACEN   = '67'  "

                If dv.Count <> 0 Then

                    DataGridView1.DataSource = dv
                    Dim k As Integer
                    k = DataGridView1.Rows.Count

                    For i = 0 To k - 1

                        Select Case Trim(DataGridView1.Rows(i).Cells(9).Value)
                            Case "10" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                            Case "37" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                            Case "38" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                            Case "51" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                            Case "42" : DataGridView1.Rows(i).Cells(10).Value = "CHILCA"
                            Case "60" : DataGridView1.Rows(i).Cells(10).Value = "SERVICIO CHILCA"
                            Case "67" : DataGridView1.Rows(i).Cells(10).Value = "PLANTA 2"
                        End Select
                    Next
                    DataGridView1.Columns.Item(9).Visible = False
                    DataGridView1.Columns.Item(5).Width = 262
                Else

                    DataGridView1.DataSource = Nothing
                End If


            End If
        End If



    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.SetToolTip(PictureBox1, "ACTUALIZAR ESTADO PEDIDOS")
        ToolTip1.ToolTipTitle = "VALIDAR ALMACEN"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
        ToolTip1.SetToolTip(Button2, "GUARDAR INFORMACION")
        ToolTip1.ToolTipTitle = "VALIDAR ALMACEN"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
        ToolTip1.SetToolTip(Button3, "BUSCAR PEDIDOS PENDIENTE")
        ToolTip1.ToolTipTitle = "VALIDAR ALMACEN"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
        ToolTip1.SetToolTip(Button4, "BUSCAR PEDIDOS PENDIENTE")
        ToolTip1.ToolTipTitle = "VALIDAR ALMACEN"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub
End Class