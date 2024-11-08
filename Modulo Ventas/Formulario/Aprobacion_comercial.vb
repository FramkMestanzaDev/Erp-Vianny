Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class Aprobacion_comercial
    Dim dt, dt3, dt4, dt5, dt6 As New DataTable
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim func10 As New fnopedido
        Dim kl As New vnapedido
        Dim kl1 As New vnapedido
        TextBox1.Text = ""
        TextBox2.Text = ""
        If RadioButton2.Checked = True Then
            Button3.Visible = False
            Button4.Visible = True
            kl.gempresa = Label4.Text
            dt3 = func10.mostrar_pedidos1(kl)
            If dt3.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt3
                Dim i As Integer
                i = DataGridView1.Rows.Count
                DataGridView1.Columns(9).Visible = False
                'DataGridView1.Columns(10).Visible = False
                For I1 = 0 To i - 1
                    If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                        DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"

                    Else
                        If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                            DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"
                        End If

                    End If


                Next

            Else
                DataGridView1.DataSource = ""
            End If

        Else
            If RadioButton1.Checked = True Then
                Button4.Visible = False
                Button3.Visible = True
                kl1.gempresa = Label4.Text
                dt4 = func10.mostrar_pedidos_aprobado1(kl1)
                If dt4.Rows.Count <> 0 Then
                    DataGridView1.DataSource = dt4
                    Dim i As Integer
                    i = DataGridView1.Rows.Count
                    DataGridView1.Columns(9).Visible = False
                    DataGridView1.Columns(10).Visible = False
                    For I1 = 0 To i - 1

                        If DataGridView1.Rows(I1).Cells(10).Value = 1 Then
                            DataGridView1.Rows(I1).Cells(8).Value = "ANULADO"
                            DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.Black
                            DataGridView1.Rows(I1).DefaultCellStyle.ForeColor = Color.White
                        Else
                            If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                                DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"
                                DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.DarkRed
                                DataGridView1.Rows(I1).DefaultCellStyle.ForeColor = Color.White
                            Else
                                If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                                    DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"

                                End If

                            End If
                        End If

                    Next
                    DataGridView1.Columns(7).Visible = False
                Else
                    DataGridView1.DataSource = ""
                End If
            End If
        End If
    End Sub

    Private Sub Aprobacion_comercial_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hs As New fnopedido
        Dim hu As New vnapedido
        hs.anular_pedidos_automaticos()
        hu.gempresa = Label4.Text
        dt = hs.mostrar_pedidos1(hu)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).Width = 80
            DataGridView1.Columns(4).Width = 250
            DataGridView1.Columns(6).Width = 130
            DataGridView1.Columns(7).Width = 90
            DataGridView1.Columns(8).Width = 100
            DataGridView1.Columns(7).Visible = False
            'DataGridView1.Columns(7).Visible = False
            'DataGridView1.Columns(2).Width = 405
            Dim i As Integer
            i = DataGridView1.Rows.Count
            For I1 = 0 To i - 1
                If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                    DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"

                Else
                    If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                        DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"
                    End If

                End If

            Next
            DataGridView1.Columns(9).Visible = False
        End If
        RadioButton2.Checked = True
        Button3.Visible = False
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If RadioButton1.Checked = True Then
            buscar4()
        Else
            buscar0()
        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If RadioButton1.Checked = True Then
            buscar2()
        Else
            buscar1()
        End If

    End Sub



    Private Sub DataGridView1_ColumnAdded(sender As Object, e As DataGridViewColumnEventArgs) Handles DataGridView1.ColumnAdded
        e.Column.SortMode = DataGridViewColumnSortMode.NotSortable
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 1 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            REPORTE_NP.TextBox1.Text = DataGridView1.Rows(num1).Cells(2).Value
            REPORTE_NP.TextBox2.Text = Label4.Text
            REPORTE_NP.Show()


        End If
    End Sub
    Sub ENVIAR_CORREO3()

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim i As Integer
        i = DataGridView1.Rows.Count
        Dim fg As New fnopedido

        For i1 = 0 To i - 1
            If DataGridView1.Rows(i1).Cells(0).Value = True Then
                Dim func As New vnapedido

                Dim ag As New vacciones
                Dim ag1 As New faccion

                func.gnumero_pedido = DataGridView1.Rows(i1).Cells(2).Value

                func.gestado = 1
                func.gempresa = Label4.Text

                ag.gAccion_realiza = "MODIFICO APROBAR"
                ag.gCodigo_iden = DataGridView1.Rows(i1).Cells(3).Value
                ag.gfyh = String.Format("{0:G}", DateTime.Now)
                ag.gmaquina = My.Computer.Name
                ag.gNombre_fomulario = "MODULO APROBACION-COMERCIAL"
                ag.gusuario = MDIParent1.Label3.Text
                ag1.insertar_acciones(ag)
                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "PEDIDO-COMERCIAL"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = 4 'APROBO
                hj1.gpc = My.Computer.Name
                hj1.gfecha = String.Format("{0:G}", DateTime.Now)
                hj1.gdetalle = Trim(DataGridView1.Rows(i1).Cells(2).Value)
                hj1.gccia = Label4.Text
                hj2.insertar_log(hj1)
                fg.actualizar_ESTADO_comercial(func)
                abrir()
                Dim Rs1 As SqlDataReader
                Dim sql1 As String = "select correo from lista_correos where formulario ='PEDIDOS_COMERCIAL   '"
                Dim cmd1 As New SqlCommand(sql1, conx)
                Rs1 = cmd1.ExecuteReader()
                Rs1.Read()
                Dim message As New MailMessage
                Dim smtp As New SmtpClient
                Dim fk As New fnopedido
                Dim jj As New vnapedido

                Dim corre As String
                jj.gvendedor = DataGridView1.Rows(i1).Cells(9).Value
                corre = fk.buscar_correo(jj)
                Dim Mailz As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA NOTA DE PEDIDO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT> APROBADO POR COMERCIAL <br/>
<FONT COLOR='blue'>* NOTA DE PEDIDO : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(2).Value) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(5).Value) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(Label5.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(4).Value) + "<br/>
<FONT COLOR='blue'>* FORMA DE PAGO : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(6).Value) + "<br/>
<FONT COLOR='blue'>* USUARIO QUE APROBO : </FONT>" + Trim(Label7.Text) + "<br/>

</body>

</html>"
                With message
                    .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
                    .To.Add(Rs1(0) & "," & corre)
                    .IsBodyHtml = True
                    .Body = Mailz
                    .Subject = "Nota Pedido N°" & DataGridView1.Rows(i1).Cells(2).Value
                    .Priority = System.Net.Mail.MailPriority.Normal
                End With

                With smtp

                    .EnableSsl = True
                    .Port = "587"
                    .Host = "smtppro.zoho.com"
                    .Credentials = New Net.NetworkCredential("admin@viannysac.com", "20Via$&it2")
                End With

                Try
                    MessageBox.Show("Su mensaje de correo ha sido enviado.", "Correo enviado",
                                     MessageBoxButtons.OK)
                Catch ex As Exception
                    MessageBox.Show("Error: " & ex.Message, "Error al enviar correo",
                                     MessageBoxButtons.OK)
                End Try

            End If

        Next
        MsgBox("SE REALIZO LA ACCION SOLICITADA")
        Dim jj1 As New vnapedido
        jj1.gempresa = Label4.Text
        dt5 = fg.mostrar_pedidos1(jj1)

        DataGridView1.DataSource = dt5
        Dim i6 As Integer
        i6 = DataGridView1.Rows.Count
        For I1 = 0 To i6 - 1
            If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"

            Else
                If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                    DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"
                End If

            End If

        Next
        DataGridView1.Columns(9).Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i As Integer
        i = DataGridView1.Rows.Count
        Dim fg As New fnopedido


        For i1 = 0 To i - 1

            If DataGridView1.Rows(i1).Cells(0).Value = True Then

                Dim Rsr23 As SqlDataReader
                abrir()
                Dim sql1023 As String = "select  isnull(estado_cobranzas,0) from nota_pedido where numero_pedido ='" + Trim(DataGridView1.Rows(i1).Cells(2).Value) + "'"
                Dim cmd1023 As New SqlCommand(sql1023, conx)
                Rsr23 = cmd1023.ExecuteReader()
                Rsr23.Read()
                If Rsr23(0) = 1 Then
                    MsgBox("LA NOTA DE PEDIDO  " + Trim(DataGridView1.Rows(i1).Cells(2).Value) + "   YA ESTA APROBADA POR COBRANZAS, EL AREA DE COMERCIAL NO PUEDE DESAPROBAR")
                Else


                    Dim func12 As New vnapedido

                    Dim ag As New vacciones
                    Dim ag1 As New faccion

                    func12.gnumero_pedido = DataGridView1.Rows(i1).Cells(2).Value


                    func12.gestado = 0
                    func12.gempresa = Label4.Text

                    ag.gAccion_realiza = "MODIFICO DESAPROBO"
                    ag.gCodigo_iden = DataGridView1.Rows(i1).Cells(3).Value
                    ag.gfyh = String.Format("{0:G}", DateTime.Now)
                    ag.gmaquina = My.Computer.Name
                    ag.gNombre_fomulario = "MODULO APROBACION-COMERCIAL"
                    ag.gusuario = MDIParent1.Label3.Text
                    ag1.insertar_acciones(ag)
                    Dim hj2 As New flog
                    Dim hj1 As New vlog
                    hj1.gmodulo = "PEDIDO-COMERCIAL"
                    hj1.gcuser = MDIParent1.Label3.Text
                    hj1.gaccion = 5 'DESAPROBO
                    hj1.gpc = My.Computer.Name
                    hj1.gfecha = String.Format("{0:G}", DateTime.Now)
                    hj1.gdetalle = Trim(DataGridView1.Rows(i1).Cells(2).Value)
                    hj1.gccia = Label4.Text
                    hj2.insertar_log(hj1)
                    fg.actualizar_ESTADO_comercial(func12)
                    abrir()
                    Dim Rs As SqlDataReader
                    Dim sql As String = "select correo from lista_correos where formulario ='PEDIDOS_COMERCIAL   '"
                    Dim cmd As New SqlCommand(sql, conx)
                    Rs = cmd.ExecuteReader()
                    Rs.Read()
                    Dim message As New MailMessage
                    Dim smtp As New SmtpClient
                    Dim fk As New fnopedido
                    Dim jj As New vnapedido
                    Dim corre As String
                    jj.gvendedor = Trim(DataGridView1.Rows(i1).Cells(9).Value)

                    corre = fk.buscar_correo(jj)
                    Dim Mailz1 As String = "" &
      " <!DOCTYPE html>" &
"<html xmlns='http://www.w3.org/1999/xhtml'>" &
"<head>" &
"    <title></title>" &
"</head>" &
"<body>
     <FONT COLOR='green'>INFORMACION DE LA NOTA DE PEDIDO</FONT><br/><br/>

<FONT COLOR='blue'>* ESTADO : </FONT>DESAPROBADO POR COMERCIAL <br/>
<FONT COLOR='blue'>* NOTA DE PEDIDO : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(2).Value) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(5).Value) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(Label5.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(4).Value) + "<br/>
<FONT COLOR='blue'>* FORMA DE PAGO : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(6).Value) + "<br/>
<FONT COLOR='blue'>* USUARIO QUE DESAPROBO : </FONT>" + Trim(Label7.Text) + "<br/>

</body>

</html>"
                    With message
                        .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
                        .To.Add(Rs(0) & "," & corre)
                        .IsBodyHtml = True
                        .Body = Mailz1
                        .Subject = "Nota Pedido N°" & Trim(DataGridView1.Rows(i1).Cells(2).Value)
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
                    MsgBox("SE REALIZO LA ACCION SOLICITADA")



                End If

            End If

        Next
        DataGridView1.DataSource = ""
        Dim jj1 As New vnapedido
        jj1.gempresa = Label4.Text
        dt4 = fg.mostrar_pedidos_aprobado1(jj1)

        DataGridView1.DataSource = dt4
        Dim i7 As Integer
        i7 = DataGridView1.Rows.Count
        For i11 = 0 To i7 - 1
            If DataGridView1.Rows(i11).Cells(8).Value = 1 Then
                DataGridView1.Rows(i11).Cells(8).Value = "APROBADO"
                DataGridView1.Rows(i11).DefaultCellStyle.BackColor = Color.DarkRed
                DataGridView1.Rows(i11).DefaultCellStyle.ForeColor = Color.White
            Else
                If DataGridView1.Rows(i11).Cells(8).Value = 0 Then
                    DataGridView1.Rows(i11).Cells(8).Value = "PENDIENTE"
                End If

            End If

        Next
        DataGridView1.Columns(9).Visible = False
    End Sub
    Private Sub buscar0()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            If RadioButton1.Checked = True Then

                else
            End If
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Cliente" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 250
                DataGridView1.Columns(6).Width = 130
                DataGridView1.Columns(7).Width = 90
                DataGridView1.Columns(8).Width = 100
                DataGridView1.Columns(7).Visible = False
                'DataGridView1.Columns(7).Visible = False
                'DataGridView1.Columns(2).Width = 405
                Dim i As Integer
                i = DataGridView1.Rows.Count
                'For I1 = 0 To i - 1
                '    If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                '        DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"

                '    Else
                '        If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                '            DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"
                '        End If

                '    End If

                'Next
                DataGridView1.Columns(9).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Pedido" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 250
                DataGridView1.Columns(6).Width = 130
                DataGridView1.Columns(7).Width = 90
                DataGridView1.Columns(8).Width = 100
                DataGridView1.Columns(7).Visible = False
                'DataGridView1.Columns(7).Visible = False
                'DataGridView1.Columns(2).Width = 405

                'i = DataGridView1.Rows.Count
                'For I1 = 0 To i - 1
                '    If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                '        DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"

                '    Else
                '        If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                '            DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"
                '        End If

                '    End If

                'Next
                DataGridView1.Columns(9).Visible = False

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt4.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Pedido" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 250
                DataGridView1.Columns(6).Width = 130
                DataGridView1.Columns(7).Width = 90
                DataGridView1.Columns(8).Width = 100
                DataGridView1.Columns(7).Visible = False
                Dim i As Integer
                i = DataGridView1.Rows.Count
                For I1 = 0 To i - 1
                    If DataGridView1.Rows(I1).Cells(8).Value = "APROBADO" Then

                        DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.DarkRed
                        DataGridView1.Rows(I1).DefaultCellStyle.ForeColor = Color.White

                    End If

                Next
                DataGridView1.Columns(9).Visible = False

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar4()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt4.Copy)
            If RadioButton1.Checked = True Then

            Else
            End If
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Cliente" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 250
                DataGridView1.Columns(6).Width = 130
                DataGridView1.Columns(7).Width = 90
                DataGridView1.Columns(8).Width = 100
                DataGridView1.Columns(7).Visible = False
                Dim i As Integer
                i = DataGridView1.Rows.Count
                For I1 = 0 To i - 1
                    If DataGridView1.Rows(I1).Cells(8).Value = "APROBADO" Then

                        DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.DarkRed
                        DataGridView1.Rows(I1).DefaultCellStyle.ForeColor = Color.White

                    End If

                Next
                DataGridView1.Columns(9).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
End Class