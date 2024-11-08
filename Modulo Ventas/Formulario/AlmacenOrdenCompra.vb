Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class AlmacenOrdenCompra
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public comando As SqlCommand
    Dim da As New DataTable
    Dim Rsr2 As SqlDataReader
    Private WithEvents dtp As New DateTimePicker()
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub AlmacenOrdenCompra_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Listar_Informacion()

        TextBox1.Select()
    End Sub
    Sub Listar_Informacion()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select ncom_3  as 'OC', fich_3 AS RUC,c.nomb_10 AS EMPRESA,fcom_3 AS FECHA,case when cigv_3='1' then 'SI' else  'NO' end AS IGV,case when mone_3 ='1' then 'S/.' else '$' end AS MONEDA, case when mone_3 ='1' then tot1_3 else tot2_3 end AS TOTAL,'' AS ESTADO from  custom_vianny.dbo.lag0300 l
            left join custom_vianny.dbo.cag1000 c on '01' = c.ccia and l.fich_3 = c.fich_10
            where year(fcom_3) ='" + Label4.Text + "' and ccia_3 ='" + Label3.Text + "' and flag_3='1' and termina_3 ='0'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        If DataGridView1.Rows.Count > 0 Then
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(8).ReadOnly = True
            DataGridView1.Columns(9).ReadOnly = True
            DataGridView1.Columns(10).ReadOnly = True
            DataGridView1.Columns(11).ReadOnly = True
            DataGridView1.Columns(12).ReadOnly = True
            DataGridView1.Columns(13).ReadOnly = True
            DataGridView1.Columns(2).Width = 75
            DataGridView1.Columns(4).Width = 75
            DataGridView1.Columns(6).Width = 65
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(8).Width = 350
            DataGridView1.Columns(9).Width = 75
            DataGridView1.Columns(10).Width = 45
            DataGridView1.Columns(11).Width = 65
            DataGridView1.Columns(12).Width = 85
        End If

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = 1 Then
            If IsDBNull(DataGridView1.Rows(e.RowIndex).Cells(1).Value) OrElse DataGridView1.Rows(e.RowIndex).Cells(1).Value Is Nothing OrElse String.IsNullOrWhiteSpace(DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()) Then

            Else
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().ToUpper()
                ' Obtener el valor de la celda editada
                Dim celda As String = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString()
                ' Verificar si la celda contiene un guion
                If celda.Contains("-") Then
                    ' Separar el valor antes y después del guion
                    Dim partes() As String = celda.Split("-"c)
                    If partes.Length = 2 Then
                        ' Obtener la parte después del guion
                        Dim despuesDelGuion As String = partes(1)

                        ' Verificar si la parte después del guion tiene menos de 8 caracteres
                        If despuesDelGuion.Length < 8 Then
                            ' Completar con ceros a la izquierda
                            despuesDelGuion = despuesDelGuion.PadLeft(8, "0"c)
                        End If

                        ' Reasignar el valor completo a la celda
                        DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = partes(0) & "-" & despuesDelGuion
                    End If
                Else
                    MsgBox("La Celda Tiene que contener Un Guion para Serapara la Serie del Correlativo")
                End If
            End If

        End If
        If e.ColumnIndex = 3 Then
            If IsDBNull(DataGridView1.Rows(e.RowIndex).Cells(2).Value) OrElse DataGridView1.Rows(e.RowIndex).Cells(1).Value Is Nothing OrElse String.IsNullOrWhiteSpace(DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString()) Then

            Else
                DataGridView1.Rows(e.RowIndex).Cells(3).Value = DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString().ToUpper()
                ' Obtener el valor de la celda editada
                Dim celda As String = DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString()
                ' Verificar si la celda contiene un guion
                If celda.Contains("-") Then
                    ' Separar el valor antes y después del guion
                    Dim partes() As String = celda.Split("-"c)
                    If partes.Length = 2 Then
                        ' Obtener la parte después del guion
                        Dim despuesDelGuion As String = partes(1)

                        ' Verificar si la parte después del guion tiene menos de 8 caracteres
                        If despuesDelGuion.Length < 8 Then
                            ' Completar con ceros a la izquierda
                            despuesDelGuion = despuesDelGuion.PadLeft(8, "0"c)
                        End If

                        ' Reasignar el valor completo a la celda
                        DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = partes(0) & "-" & despuesDelGuion
                    End If
                Else
                    MsgBox("La Celda Tiene que contener Un Guion para Serapara la Serie del Correlativo")
                End If
            End If

        End If
        If e.ColumnIndex = 5 Then
            If Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(5).Value) > Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(12).Value) Then
                MsgBox("EL TOTAL DE LA FACTURA ES MAYOR A LA ORDEN DE COMPRA")
                DataGridView1.Rows(e.RowIndex).Cells(5).Value = 0
            End If
        End If

        If e.ColumnIndex = 2 Then
            ' Asignar el valor seleccionado en el DateTimePicker a la celda
            DataGridView1.CurrentCell.Value = dtp.Value.ToString("dd/MM/yyyy")

            ' Ocultar el DateTimePicker
            dtp.Visible = False
        End If
        If e.ColumnIndex = 4 Then
            ' Asignar el valor seleccionado en el DateTimePicker a la celda
            DataGridView1.CurrentCell.Value = dtp.Value.ToString("dd/MM/yyyy")

            ' Ocultar el DateTimePicker
            dtp.Visible = False
        End If



    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Function EsCeldaNulla(fila As Integer, columna As Integer) As Boolean
        Dim valorCelda = DataGridView1.Rows(fila).Cells(columna).Value

        If IsDBNull(valorCelda) OrElse valorCelda Is Nothing OrElse String.IsNullOrWhiteSpace(valorCelda.ToString()) Then
            Return True ' La celda es nulla o vacía
        Else

            Return False
        End If

    End Function
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If e.ColumnIndex = 1 Or e.ColumnIndex = 2 Or e.ColumnIndex = 3 Or e.ColumnIndex = 4 Then
                DataGridView1.BeginEdit(True)
            End If

            If e.ColumnIndex = 0 Then
                If EsCeldaNulla(e.RowIndex, 1) = False And EsCeldaNulla(e.RowIndex, 2) = False And Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(5).Value) > 0 Then
                    DetalleOrden.Label1.Text = Label3.Text
                    DetalleOrden.Label2.Text = DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString().Trim()
                    DetalleOrden.TextBox1.Text = Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(5).Value).ToString()
                    DetalleOrden.Label3.Text = Label4.Text
                    DetalleOrden.Label4.Text = Label5.Text
                    If DataGridView1.Rows(e.RowIndex).Cells(11).Value.ToString().Trim() = "S/." Then
                        DetalleOrden.Label7.Text = "1"
                    Else
                        DetalleOrden.Label7.Text = "2"
                    End If
                    'MsgBox(e.RowIndex)
                    'MsgBox(Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString().Trim(), 1, 2))
                    DetalleOrden.Label8.Text = Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString().Trim(), 1, 2)
                    DetalleOrden.Label9.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value.ToString().Trim()
                    DetalleOrden.Label10.Text = DataGridView1.Rows(e.RowIndex).Cells(8).Value.ToString().Trim()
                    DetalleOrden.Label11.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                    DetalleOrden.Label12.Text = DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString().Trim()
                    DetalleOrden.ShowDialog()
                Else
                    MsgBox("Es Obligatorio Ingresar De La Guia-Remision y Factura")
                End If

            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscarn()
    End Sub
    Private Sub buscarn()
        Try
            Dim ds As New DataSet
            Dim K As Integer


            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer
            If String.IsNullOrEmpty(TextBox1.Text) Then
                DataGridView1.DataSource = ds.Tables(0)
            Else

                dv.RowFilter = "OC" & " like '%" & TextBox1.Text & "%'"

                If dv.Count <> 0 Then
                    DataGridView1.DataSource = dv
                Else

                    DataGridView1.DataSource = ds.Tables(0)
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub buscare()
        Try
            Dim ds As New DataSet
            Dim K As Integer


            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "EMPRESA" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv
            Else

                'DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscare()
    End Sub

    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        ' Verificar si la columna es la que queremos que tenga el DateTimePicker (por ejemplo, la columna 1)
        If e.ColumnIndex = 2 Then
            ' Configurar el DateTimePicker
            dtp.Format = DateTimePickerFormat.Short ' O puedes usar otro formato
            dtp.Visible = True ' Inicialmente no visible

            ' Posicionar el DateTimePicker sobre la celda
            Dim rect As Rectangle = DataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
            dtp.Location = rect.Location
            dtp.Size = rect.Size
            dtp.Visible = True

            ' Asignar el valor de la celda al DateTimePicker, si hay un valor
            If DataGridView1.CurrentCell.Value IsNot Nothing Then
                dtp.Value = Convert.ToDateTime(DataGridView1.CurrentCell.Value)
            End If

            ' Añadir el DateTimePicker al DataGridView
            DataGridView1.Controls.Add(dtp)
        End If
        If e.ColumnIndex = 4 Then
            ' Configurar el DateTimePicker
            dtp.Format = DateTimePickerFormat.Short ' O puedes usar otro formato
            dtp.Visible = True ' Inicialmente no visible

            ' Posicionar el DateTimePicker sobre la celda
            Dim rect As Rectangle = DataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
            dtp.Location = rect.Location
            dtp.Size = rect.Size
            dtp.Visible = True

            ' Asignar el valor de la celda al DateTimePicker, si hay un valor
            If DataGridView1.CurrentCell.Value IsNot Nothing Then
                dtp.Value = Convert.ToDateTime(DataGridView1.CurrentCell.Value)
            End If

            ' Añadir el DateTimePicker al DataGridView
            DataGridView1.Controls.Add(dtp)
        End If
    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        dtp.Visible = False
    End Sub
    Sub ENVIAR_CORREO()
        abrir()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim fk As New fnopedido
        Dim jj As New vnapedido
        Dim cabecera, usuario, contras As String
        usuario = ""
        contras = ""
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "SELECT correo,contrasena FROM usuario_modulo where usuario ='" + Trim(Label5.Text.ToString().Trim()) + "'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        If Rsr1991.Read() Then
            usuario = Rsr1991(0)
            contras = Rsr1991(1)
        End If
        Rsr1991.Close()
        cabecera = "INFORMACION DEL INGRESO DE TELA"
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
                <th align='center'><font color='blue'>PEDIDO</font></th>
                <th align='center'><font color='blue'>CODIGO</font></th>
                <th align='center'><font color='blue'>PRODUCTO</font></th>
                <th align='center'><font color='blue'>CANTIDAD</font></th>
                <th align='center'><font color='blue'>UNIDAD</font></th>
                <th align='center'><font color='blue'>ORDEN COMPRA</font></th>
            </tr>
        </thead>
        <tbody>"

        ' Comienza el bucle while para leer los registros del lector
        Dim Rsr1993 As SqlDataReader
        Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim sql1013 As String = "select item_3a as ITEMS,pedido_3a AS PEDIDO,linea_17 AS CODIGO,cg.nomb_17 AS PRODUCTO, cante_3a AS CANT,unid_17 AS UND,ncom_3a AS OC from custom_vianny.dbo.lap0300 lp
                                        left join custom_vianny.dbo.cag1700 cg on lp.ccia_3a = cg.ccia and lp.linea_3a = cg.linea_17
                                        where ccia_3a ='01' and ncom_3a ='" + filaSeleccionada.Cells("OC").Value.ToString().Trim() + "' and flag_3a ='1'"
        Dim cmd1013 As New SqlCommand(sql1013, conx)
        Rsr1993 = cmd1013.ExecuteReader()

        While Rsr1993.Read()
            Mailz &= "<tr>" &
             "<td align='center'>" + Rsr1993(1).ToString().Trim() + "</td>" &
             "<td align='center'>" + Rsr1993(2).ToString().Trim() + "</td>" &
             "<td align='center'>" + Rsr1993(3).ToString().Trim() + "</td>" &
             "<td align='center'>" + Rsr1993(4).ToString().Trim() + "</td>" &
             "<td align='center'>" + Rsr1993(5).ToString().Trim() + "</td>" &
             "<td align='center'>" + Rsr1993(6).ToString().Trim() + "</td>" &
             "</tr>"
        End While
        Rsr1993.Close()
        ' Cierra la tabla y el cuerpo del HTML
        Mailz &= "
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
            .Subject = cabecera
            .Priority = System.Net.Mail.MailPriority.Normal
        End With


        With smtp

            .EnableSsl = True
            .Port = "587"
            If usuario = "JMEDINA" Then
                .Host = "smtp.gmail.com"
            Else
                .Host = "smtppro.zoho.com"
            End If

            '.Credentials = New Net.NetworkCredential("admin@viannysac.com", "20Via$&it2")

            .Credentials = New Net.NetworkCredential(usuario, contras)
            .DeliveryMethod = SmtpDeliveryMethod.Network
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
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ENVIAR_CORREO()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.ColumnIndex = 6 Then
            Dim seleccionado As DataGridViewRow = DataGridView1.SelectedRows(0)

            Rpt_oc_1.TextBox1.Text = Label3.Text
            Rpt_oc_1.TextBox2.Text = Mid(seleccionado.Cells("OC").Value, 1, 2).ToString().Trim()
            Rpt_oc_1.TextBox3.Text = Mid(seleccionado.Cells("OC").Value, 3, 6).ToString().Trim()
            Rpt_oc_1.TextBox4.Text = Mid(seleccionado.Cells("OC").Value, 1, 2).ToString().Trim()
            Rpt_oc_1.TextBox5.Text = Mid(seleccionado.Cells("OC").Value, 3, 6).ToString().Trim()
            Rpt_oc_1.TextBox6.Text = ""
            Rpt_oc_1.TextBox7.Text = 0
            Rpt_oc_1.TextBox8.Text = "\\172.21.0.1\erpcaesoft$\LIB.APPSV2\imagenes\"
            Rpt_oc_1.Show()
        End If
    End Sub
End Class