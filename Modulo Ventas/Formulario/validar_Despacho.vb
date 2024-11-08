Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class validar_Despacho
    Dim dt As New DataTable
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
    Private Sub validar_Despacho_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim gf As New fnopedido
        Dim gf1 As New vnapedido
        gf1.gperiodo = Label2.Text
        gf1.gempresa = Label1.Text
        dt = gf.validar_despacho(gf1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            Dim k As Integer
            k = DataGridView1.Rows.Count


            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(9).Width = 220
            DataGridView1.Columns(14).Visible = False
            DataGridView1.Columns(15).DefaultCellStyle.BackColor = Color.Beige
            DataGridView1.Columns(16).DefaultCellStyle.BackColor = Color.Azure
            DataGridView1.Columns(17).DefaultCellStyle.BackColor = Color.Coral
            For i = 0 To k - 1

                If Trim(DataGridView1.Rows(i).Cells(15).Value) = "PENDIENTE" Then
                    DataGridView1.Rows(i).Cells(15).Style.BackColor = Color.DarkCyan
                    DataGridView1.Rows(i).Cells(15).Style.ForeColor = Color.White
                End If
                If Trim(DataGridView1.Rows(i).Cells(16).Value) = "PENDIENTE" Then
                    DataGridView1.Rows(i).Cells(16).Style.BackColor = Color.DarkKhaki
                    DataGridView1.Rows(i).Cells(16).Style.ForeColor = Color.White
                End If
                If Trim(DataGridView1.Rows(i).Cells(17).Value) = "PENDIENTE" Then
                    DataGridView1.Rows(i).Cells(17).Style.BackColor = Color.DarkGray
                    DataGridView1.Rows(i).Cells(17).Style.ForeColor = Color.White
                End If
                'abrir()
                'llenar_combo_box()
                'llenar_combo_box1()
            Next
        End If

    End Sub
    'Sub llenar_combo_box()
    '    Try
    '        Dim lsQuery As String = "select distinct chofer  from ingre_transportista"
    '        Dim loDataTable As New DataTable
    '        Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
    '        loDataAdapter.Fill(loDataTable)
    '        CHOFER.DataSource = loDataTable

    '        CHOFER.DisplayMember = "chofer"
    '        CHOFER.ValueMember = "chofer"
    '        CHOFER.DropDownWidth = 200
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub
    'Sub llenar_combo_box1()
    '    Try
    '        Dim lsQuery As String = "select distinct placa  from ingre_transportista"
    '        Dim loDataTable As New DataTable
    '        Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
    '        loDataAdapter.Fill(loDataTable)
    '        VEHICULO.DataSource = loDataTable

    '        VEHICULO.DisplayMember = "placa"
    '        VEHICULO.ValueMember = "placa"
    '        VEHICULO.DropDownWidth = 200
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

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




                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(9).Width = 220
                DataGridView1.Columns(14).Visible = False
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






                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(12).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(9).Width = 220
                DataGridView1.Columns(14).Visible = False

            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)



    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar2()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim gf As New fnopedido
        Dim gf1 As New vnapedido
        gf1.gperiodo = Label2.Text
        gf1.gempresa = Label1.Text
        dt = gf.validar_despacho(gf1)

        DataGridView1.DataSource = dt
        Dim k As Integer
        k = DataGridView1.Rows.Count

        DataGridView1.Columns(11).Visible = False
        DataGridView1.Columns(12).Visible = False
        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(9).Width = 220
        DataGridView1.Columns(14).Visible = False
        DataGridView1.Columns(15).DefaultCellStyle.BackColor = Color.Beige
        DataGridView1.Columns(16).DefaultCellStyle.BackColor = Color.Azure
        DataGridView1.Columns(17).DefaultCellStyle.BackColor = Color.Coral
        Dim k1 As Integer
        k1 = DataGridView1.Rows.Count
        For i = 0 To k - 1
            If Trim(DataGridView1.Rows(i).Cells(15).Value) = "PENDIENTE" Then
                DataGridView1.Rows(i).Cells(15).Style.BackColor = Color.DarkCyan
                DataGridView1.Rows(i).Cells(15).Style.ForeColor = Color.White
            End If
            If Trim(DataGridView1.Rows(i).Cells(16).Value) = "PENDIENTE" Then
                DataGridView1.Rows(i).Cells(16).Style.BackColor = Color.DarkKhaki
                DataGridView1.Rows(i).Cells(16).Style.ForeColor = Color.White
            End If
            If Trim(DataGridView1.Rows(i).Cells(17).Value) = "PENDIENTE" Then
                DataGridView1.Rows(i).Cells(17).Style.BackColor = Color.DarkGray
                DataGridView1.Rows(i).Cells(17).Style.ForeColor = Color.White
            End If
        Next


    End Sub
    Dim jkk As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim J As Integer

        J = DataGridView1.Rows.Count

        For i = 0 To J - 1

            If DataGridView1.Rows(i).Cells(1).Value = True Then
                Dim ju As New fnopedido
                Dim ju1 As New vnapedido

                ju1.gnumero_pedido = DataGridView1.Rows(i).Cells(2).Value

                abrir()
                Dim cmd17 As New SqlCommand("update nota_pedido set estado_tranporte=1  WHERE numero_pedido =@pedido and empresa =@ccia", conx)
                cmd17.Parameters.AddWithValue("@ccia", Trim(Label1.Text))
                cmd17.Parameters.AddWithValue("@pedido", Trim(DataGridView1.Rows(i).Cells(7).Value))
                cmd17.ExecuteNonQuery()
                Dim message As New MailMessage
                Dim smtp As New SmtpClient
                Dim fk As New fnopedido
                Dim jj As New vnapedido
                Dim corre As String

                jj.gvendedor = Trim(DataGridView1.Rows(i).Cells(13).Value)

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

<FONT COLOR='blue'>* ESTADO : </FONT> DESPACHO <br/>
<FONT COLOR='blue'>* NOTA DE PEDIDO : </FONT>" + Trim(DataGridView1.Rows(i).Cells(7).Value) + " <br/>
<FONT COLOR='blue'>* VENDEDOR : </FONT>" + Trim(DataGridView1.Rows(i).Cells(13).Value) + " <br/>
<FONT COLOR='blue'>* CHOFER : </FONT>" + Trim(DataGridView1.Rows(i).Cells(3).Value) + " <br/>
<FONT COLOR='blue'>* VEHICULO : </FONT>" + Trim(DataGridView1.Rows(i).Cells(4).Value) + " <br/>
<FONT COLOR='blue'>* CLIENTE : </FONT>" + Trim(DataGridView1.Rows(i).Cells(9).Value) + "<br/>
<FONT COLOR='blue'>* FECHA DESPACHO : </FONT>" + Trim(DataGridView1.Rows(i).Cells(2).Value) + "<br/>

</body>

</html>"
                With message
                    .From = New System.Net.Mail.MailAddress("admin@viannysac.com")
                    .To.Add(Rs(0) & "," & corre)
                    .IsBodyHtml = True
                    .Body = Mailz
                    .Subject = "Nota Pedido N°" & Trim(DataGridView1.Rows(i).Cells(7).Value)
                    .Priority = System.Net.Mail.MailPriority.Normal
                End With

                With smtp
                    .EnableSsl = True
                    .Port = "25"
                    .Host = "mail.onehostingperu.com"
                    .Credentials = New Net.NetworkCredential("admin@viannysac.com", "I9!?@ni2;+go")
                    .Send(message)
                End With

                Try
                    MessageBox.Show("Su mensaje de correo ha sido enviado.", "Correo enviado",
                                     MessageBoxButtons.OK)
                Catch ex As Exception
                    MessageBox.Show("Error:  " & ex.Message, "Error al enviar correo",
                                     MessageBoxButtons.OK)
                End Try
                abrir()
                Dim cmd16 As New SqlCommand("update nota_pedido set fecha_transporte=@fecha_transporte,chofer_unidad=@chofer_unidad,unidad =@unidad,OBSER_TRANSPORTE=@OBSER_TRANSPORTE,estibador=@estibador WHERE numero_pedido =@pedido and empresa =@ccia", conx)
                cmd16.Parameters.AddWithValue("@fecha_transporte", Trim(DataGridView1.Rows(i).Cells(2).Value))
                If Trim(DataGridView1.Rows(i).Cells(3).Value) = "" Then
                    cmd16.Parameters.AddWithValue("@chofer_unidad", DBNull.Value)
                Else
                    cmd16.Parameters.AddWithValue("@chofer_unidad", Trim(DataGridView1.Rows(i).Cells(3).Value))
                End If
                cmd16.Parameters.AddWithValue("@ccia", Trim(Label1.Text))
                cmd16.Parameters.AddWithValue("@pedido", Trim(DataGridView1.Rows(i).Cells(7).Value))
                If Trim(DataGridView1.Rows(i).Cells(4).Value) = "" Then
                    cmd16.Parameters.AddWithValue("@unidad", DBNull.Value)
                Else
                    cmd16.Parameters.AddWithValue("@unidad", Trim(DataGridView1.Rows(i).Cells(4).Value))
                End If

                cmd16.Parameters.AddWithValue("@OBSER_TRANSPORTE", Trim(DataGridView1.Rows(i).Cells(5).Value))
                cmd16.Parameters.AddWithValue("@estibador", Trim(DataGridView1.Rows(i).Cells(6).Value))
                cmd16.ExecuteNonQuery()
            End If

        Next

        MsgBox("SE ACTUALIZO LOS DATOS SOLICITADOS")
        Dim gf As New fnopedido
        Dim gf1 As New vnapedido


        gf1.gperiodo = Label2.Text
        gf1.gempresa = Label1.Text
        jkk = gf.validar_despacho(gf1)

        DataGridView1.DataSource = jkk
        Dim k As Integer
        k = DataGridView1.Rows.Count




        DataGridView1.Columns(11).Visible = False
        DataGridView1.Columns(12).Visible = False
        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(9).Width = 220
        DataGridView1.Columns(14).Visible = False
    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.SetToolTip(PictureBox1, "ACTUALIZAR PEND-CANC")
        ToolTip1.ToolTipTitle = "PEDIDOS COMERCIAL"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub


    Private Sub Button1_MouseHover(sender As Object, e As EventArgs) Handles Button1.MouseHover
        ToolTip3.SetToolTip(Button1, "CERRAR DESPACHO")
        ToolTip3.ToolTipTitle = "PEDIDOS COMERCIAL"
        ToolTip3.ToolTipIcon = ToolTipIcon.Info
    End Sub
    Dim DateTimePicker1 As New DateTimePicker()
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then
            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            REPORTE_NP.TextBox1.Text = DataGridView1.Rows(num1).Cells(7).Value
            REPORTE_NP.TextBox2.Text = Label1.Text
            REPORTE_NP.Show()
        Else
            If e.ColumnIndex = 1 Then
                DataGridView1.Rows(e.RowIndex).Cells(2).Value = DateTime.Now.ToString("yyyy/MM/dd")
                DataGridView1.Columns(2).ReadOnly = False
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.Columns(1).Visible = False
        Dim gf As New fnopedido
        Dim gf1 As New vnapedido
        gf1.gperiodo = Label2.Text
        gf1.gempresa = Label1.Text
        dt = gf.validar_despacho2(gf1)

        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            Dim k As Integer
            k = DataGridView1.Rows.Count
            For i = 0 To k - 1
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                If Trim(DataGridView1.Rows(i).Cells(10).Value) <> "" Then


                    abrir()
                    Dim Rs3 As SqlDataReader
                    Dim sql1 As String = "select  fecha_transporte,chofer_unidad,unidad,OBSER_TRANSPORTE,estibador from nota_pedido where numero_pedido ='" + Trim(DataGridView1.Rows(i).Cells(7).Value) + "' and empresa =" + Trim(Label1.Text) + " and YEAR(fecha) ='" + Label2.Text + "' AND fecha >='20220704'"
                    Dim cmd1 As New SqlCommand(sql1, conx)
                    Rs3 = cmd1.ExecuteReader()
                    Rs3.Read()

                    If Rs3(0) Is DBNull.Value Then
                        DataGridView1.Rows(i).Cells(2).Value = ""
                    Else
                        DataGridView1.Rows(i).Cells(2).Value = Rs3(0)
                    End If
                    If Rs3(1) Is DBNull.Value Then
                        DataGridView1.Rows(i).Cells(3).Value = ""
                    Else

                        DataGridView1.Rows(i).Cells(3).Value = Rs3(1).ToString
                    End If


                    If Rs3(2) Is DBNull.Value Then
                        DataGridView1.Rows(i).Cells(4).Value = ""
                    Else

                        DataGridView1.Rows(i).Cells(4).Value = Rs3(2).ToString
                    End If




                    If Rs3(3) Is DBNull.Value Then
                        DataGridView1.Rows(i).Cells(5).Value = ""
                    Else
                        DataGridView1.Rows(i).Cells(5).Value = Rs3(3)
                    End If
                    If Rs3(4) Is DBNull.Value Then
                        DataGridView1.Rows(i).Cells(6).Value = ""
                    Else
                        DataGridView1.Rows(i).Cells(6).Value = Rs3(4)
                    End If
                    Rs3.Close()
                Else
                    '    abrir()
                    '    Dim Rs32 As SqlDataReader
                    '    Dim sql12 As String = "select fecha,chofer,vehiculo,observacion,cast(estibador as char(2)) from  despachos_adicionales where id_despacho ='" + Trim(DataGridView1.Rows(i).Cells(13).Value) + "' "
                    '    Dim cmd12 As New SqlCommand(sql12, conx)
                    '    Rs32 = cmd12.ExecuteReader()
                    '    Rs32.Read()
                    '    If Rs32(0) Is DBNull.Value Then
                    '        DataGridView1.Rows(i).Cells(2).Value = ""
                    '    Else
                    '        DataGridView1.Rows(i).Cells(2).Value = Rs32(0)
                    '    End If
                    '    If Rs32(1) Is DBNull.Value Then
                    '        DataGridView1.Rows(i).Cells(3).Value = ""
                    '    Else


                    '        DataGridView1.Rows(i).Cells(3).Value = Rs32(1)
                    '        End If


                    '    If Rs32(2) Is DBNull.Value Then
                    '        DataGridView1.Rows(i).Cells(4).Value = ""

                    '    Else
                    '            DataGridView1.Rows(i).Cells(4).Value = Rs32(2)


                    '    End If
                    '    If Rs32(3) Is DBNull.Value Then

                    '        DataGridView1.Rows(i).Cells(5).Value = ""
                    '    Else
                    '        DataGridView1.Rows(i).Cells(5).Value = Rs32(3)
                    '    End If
                    '    If Rs32(4) Is DBNull.Value Then
                    '        DataGridView1.Rows(i).Cells(6).Value = ""
                    '    Else
                    '        DataGridView1.Rows(i).Cells(6).Value = Rs32(4)
                    '    End If
                    '    Rs32.Close()
                End If
            Next

            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(9).Width = 220
            DataGridView1.Columns(14).Visible = False
        End If
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Despachos_adicionales.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim DF As New Exportar

        DF.llenarExcel(DataGridView1)
    End Sub
End Class