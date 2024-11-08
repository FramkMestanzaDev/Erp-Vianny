Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class Aprobar_notapedido
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
                func.gempresa = Label3.Text

                ag.gAccion_realiza = "MODIFICO APROBAR"
                ag.gCodigo_iden = DataGridView1.Rows(i1).Cells(2).Value
                ag.gfyh = String.Format("{0:G}", DateTime.Now)
                ag.gmaquina = My.Computer.Name
                ag.gNombre_fomulario = "MODULO APROBACION"
                ag.gusuario = MDIParent1.Label3.Text

                ag1.insertar_acciones(ag)
                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "PEDIDO-COMERCIAL"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = 6 'APROBO cobranzas
                hj1.gpc = My.Computer.Name
                hj1.gfecha = String.Format("{0:G}", DateTime.Now)
                hj1.gdetalle = Trim(DataGridView1.Rows(i1).Cells(2).Value)
                hj1.gccia = Label3.Text
                hj2.insertar_log(hj1)
                fg.actualizar_ESTADO(func)
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

<FONT COLOR='blue'>* ESTADO : </FONT>APROBADO POR COBRANZAS <br/>
<FONT COLOR='blue'>* NOTA DE PEDIDO : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(2).Value) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(5).Value) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(Label5.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(4).Value) + "<br/>
<FONT COLOR='blue'>* FORMA DE PAGO : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(6).Value) + "<br/>
<FONT COLOR='blue'>* USUARIO QUE APROBO : </FONT>" + Trim(Label6.Text) + "<br/>

</body>

</html>"
                With message
                    .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
                    .To.Add(Rs(0) & "," & corre)
                    Rs.Close()
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
                    .Send(message)
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
        MsgBox("SE APROBO LA NOTA DE PEDIDO")
        Dim JK As New vnapedido
        JK.gempresa = Label3.Text
        dt5 = fg.mostrar_pedidos2(JK)

        DataGridView1.DataSource = dt5

        Dim i9 As Integer
        i9 = DataGridView1.Rows.Count
        For I1 = 0 To i9 - 1

            If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"


            Else
                If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                    DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"
                End If

            End If
        Next
        DataGridView1.Columns(4).Width = 230
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i As Integer
        i = DataGridView1.Rows.Count
        Dim fg As New fnopedido

        For i1 = 0 To i - 1
            If DataGridView1.Rows(i1).Cells(0).Value = True Then
                Dim func As New vnapedido

                Dim ag As New vacciones
                Dim ag1 As New faccion

                func.gnumero_pedido = DataGridView1.Rows(i1).Cells(2).Value
                func.gempresa = Label3.Text
                func.gestado = 0

                ag.gAccion_realiza = "MODIFICO DESAPROBO"
                ag.gCodigo_iden = DataGridView1.Rows(i1).Cells(2).Value
                ag.gfyh = String.Format("{0:G}", DateTime.Now)
                ag.gmaquina = My.Computer.Name
                ag.gNombre_fomulario = "MODULO APROBACION"
                ag.gusuario = MDIParent1.Label3.Text
                ag1.insertar_acciones(ag)
                Dim hj2 As New flog
                Dim hj1 As New vlog
                hj1.gmodulo = "PEDIDO-COMERCIAL"
                hj1.gcuser = MDIParent1.Label3.Text
                hj1.gaccion = 7 'desaprobo cobranzas
                hj1.gpc = My.Computer.Name
                hj1.gfecha = String.Format("{0:G}", DateTime.Now)
                hj1.gdetalle = Trim(DataGridView1.Rows(i1).Cells(2).Value)
                hj1.gccia = Label3.Text
                hj2.insertar_log(hj1)
                fg.actualizar_ESTADO(func)
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

<FONT COLOR='blue'>* ESTADO : </FONT>DESAPROBADO POR COBRANZAS <br/>
<FONT COLOR='blue'>* NOTA DE PEDIDO : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(2).Value) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(5).Value) + " <br/>
<FONT COLOR='blue'>* EMPRESA : </FONT>" + Trim(Label5.Text) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(4).Value) + "<br/>
<FONT COLOR='blue'>* FORMA DE PAGO : </FONT>" + Trim(DataGridView1.Rows(i1).Cells(6).Value) + "<br/>
<FONT COLOR='blue'>* USUARIO QUE DESAPROBO : </FONT>" + Trim(Label6.Text) + "<br/>

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
                    .Send(message)
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
        MsgBox("SE DESAPROBO LA NOTA DE PEDIDO")
        Dim JK As New vnapedido
        JK.gempresa = Label3.Text
        dt6 = fg.mostrar_pedidos_aprobado2(JK)


        DataGridView1.DataSource = dt6

        Dim i9 As Integer
        i9 = DataGridView1.Rows.Count
        For I1 = 0 To i9 - 1

            If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"


            Else
                If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                    DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"
                End If

            End If
            If DataGridView1.Rows(I1).Cells(8).Value = "APROBADO" Then

                DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.IndianRed
            End If
        Next


        DataGridView1.Columns(4).Width = 230
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar0()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 1 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            REPORTE_NP.TextBox1.Text = DataGridView1.Rows(num1).Cells(2).Value
            REPORTE_NP.TextBox2.Text = Label3.Text
            REPORTE_NP.Show()


        End If
    End Sub

    Private Sub Aprobar_notapedido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hs As New fnopedido
        Dim JK As New vnapedido
        dt.Clear()
        hs.anular_pedidos_automaticos()
        JK.gempresa = Label3.Text
        dt = hs.mostrar_pedidos2(JK)
        If hs.saved = True Then
            If dt.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt
                'DataGridView1.Columns(4).Visible = False
                'DataGridView1.Columns(5).Visible = False
                'DataGridView1.Columns(6).Visible = False
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
                DataGridView1.Columns("TOTAL").DefaultCellStyle.Format = "N2"
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(4).Width = 230
            End If
        End If


        RadioButton2.Checked = True
        Button3.Visible = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim hs As New fnopedido
        If hs.saved = True Then
            dt.Clear()
        End If

        Dim func10 As New fnopedido
        Dim JK As New vnapedido
        JK.gempresa = Label3.Text
        If RadioButton2.Checked = True Then
            Button3.Visible = False
            Button4.Visible = True
            dt = func10.mostrar_pedidos2(JK)
            If dt.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt
                Dim i8 As Integer
                i8 = DataGridView1.Rows.Count
                For I1 = 0 To i8 - 1
                    If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                        DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"

                    Else
                        If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                            DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"
                        End If

                    End If

                Next
                DataGridView1.Columns(9).Visible = False
                DataGridView1.Columns(4).Width = 230
            Else
                DataGridView1.DataSource = ""
            End If

        Else
            If RadioButton1.Checked = True Then
                Button4.Visible = False
                Button3.Visible = True
                dt = func10.mostrar_pedidos_aprobado2(JK)
                If dt.Rows.Count <> 0 Then
                    DataGridView1.DataSource = dt
                    Dim i As Integer
                    i = DataGridView1.Rows.Count

                    For I1 = 0 To i - 1
                        If DataGridView1.Rows(I1).Cells(8).Value = 1 Then

                            DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.IndianRed
                        End If

                    Next
                    Dim i6 As Integer
                    i6 = DataGridView1.Rows.Count
                    For I1 = 0 To i6 - 1
                        If DataGridView1.Rows(I1).Cells(11).Value = 1 Then
                            DataGridView1.Rows(I1).Cells(8).Value = "ANULADO"
                            DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.Black
                            DataGridView1.Rows(I1).DefaultCellStyle.ForeColor = Color.White
                        Else
                            If DataGridView1.Rows(I1).Cells(8).Value = 1 Then
                                DataGridView1.Rows(I1).Cells(8).Value = "APROBADO"

                            Else
                                If DataGridView1.Rows(I1).Cells(8).Value = 0 Then
                                    DataGridView1.Rows(I1).Cells(8).Value = "PENDIENTE"
                                End If

                            End If
                        End If
                    Next
                    DataGridView1.Columns(4).Width = 230

                Else
                    DataGridView1.DataSource = ""
                End If
            End If
        End If
    End Sub
    Private Sub buscar0()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Cliente" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim i6 As Integer
                i6 = DataGridView1.Rows.Count
                For I1 = 0 To i6 - 1
                    If DataGridView1.Rows(I1).Cells(8).Value = "APROBADO" Then

                        DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.IndianRed
                    End If

                Next

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


            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
End Class