Imports System.Data.SqlClient
Public Class buscar_Packing
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Dim Rsr2, Rsr20, Rsr21, Rsr22 As SqlDataReader

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub buscar_Packing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        MOSTRAR()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar2()
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim k As Integer
            Dim jk As Double

            dv.RowFilter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(2).Width = 70
                DataGridView1.Columns(3).Width = 500
                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 70
                DataGridView1.Columns(6).Width = 70
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(8).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub insertar_nia()
        Dim x, mj, valo, mp, mop, final As String
        Dim pp As Integer
        Dim fun As New vinsertar_nia
        Dim func As New fnia
        Dim func1 As New insertardetallenia
        Dim func2 As New fnia
        Dim con As String
        Dim hn As New vnia
        Dim fg As New fnia
        Dim registro As String

        'con = TextBox4.Text & TextBox5.Text
        'hn.gccia = Label2.Text
        'hn.gcomprobante = con
        'hn.galmacen = "10"
        'hn.gncom = 1
        'hn.gano = Label2.Text
        'fg.eliminar_nia(hn)
        registro = correlativo()
        fun.gccia = Label4.Text
        fun.gncom = registro
        fun.gglosa = "REGISTRO AUTOMATICO DEL PACKING"
        fun.gfecha = DateTimePicker1.Value
        fun.gguia = ""
        fun.gano = Label2.Text
        fun.gdoc = ""
        fun.galmacen = "10"
        fun.gempresa = "20508740361"
        fun.gint = "INT"
        fun.gorigen = "1001"
        fun.gtidoc = ""
        fun.gusuario = Label1.Text
        fun.gfase = ""
        Dim i, num2 As Integer

        num2 = DataGridView1.Rows.Count
        For i = 0 To num2 - 1
            If Trim(DataGridView1.Rows(i).Cells(2).Value) = True Then
                Dim nu As String
                Dim jj As New fingresopac
                Dim aa As New vpackingtela
                func1.gccia = Label4.Text
                func1.gncom = registro
                Select Case Convert.ToString(i).Length
                    Case "1" : nu = "00" & "" & i
                    Case "2" : nu = "0" & "" & i
                    Case "3" : nu = i
                End Select
                func1.gitem = nu
                func1.glinea = DataGridView1.Rows(i).Cells(9).Value
                func1.gop = DataGridView1.Rows(i).Cells(10).Value
                func1.gund = "KGS"
                func1.gcantidad = DataGridView1.Rows(i).Cells(5).Value
                func1.gcantidad1 = DataGridView1.Rows(i).Cells(6).Value
                func1.gpartida = DataGridView1.Rows(i).Cells(2).Value
                func1.galmacen = "10"
                func1.gumpresentacion = "RLL"
                func1.gotejeduria = ""
                func1.goc = ""
                func1.gano = Label2.Text
                func1.glote = Trim(DataGridView1.Rows(i).Cells(2).Value)
                func1.gcenvio = DataGridView1.Rows(i).Cells(5).Value
                Select Case Trim(DataGridView1.Rows(i).Cells(7).Value)
                    Case "PRIMERA" : func1.gclasif = "1"
                    Case "SEGUNDA" : func1.gclasif = "2"
                    Case "TERCERA" : func1.gclasif = "3"

                End Select
                aa.gpartida = DataGridView1.Rows(i).Cells(2).Value

                aa.galmacen = "10"
                aa.ges = 1
                aa.gpartida = Trim(DataGridView1.Rows(i).Cells(2).Value)
                aa.gcalidad = Trim(DataGridView1.Rows(i).Cells(7).Value)
                jj.actualizar_seleccionado(aa)
                func2.insertar_detalle_nia(func1)
            End If
        Next
        If Func.insertar_cabecera_nia(fun) Then

            MessageBox.Show("Datos registrados correctamente en la NIA", "Guardando registros", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Dim hj2 As New flog
            Dim hj1 As New vlog
            hj1.gmodulo = "NIA-ALMACEN"
            hj1.gcuser = MDIParent1.Label3.Text
            hj1.gaccion = 1
            hj1.gpc = My.Computer.Name
            hj1.gfecha = DateTimePicker1.Value
            hj1.gdetalle = "10" & registro
            hj1.gccia = Label4.Text
            hj2.insertar_log(hj1)
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE INGRESO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                Reporte_Nia_Nsa.TextBox1.Text = "10"
                Reporte_Nia_Nsa.TextBox2.Text = 1
                Reporte_Nia_Nsa.TextBox3.Text = registro
                Reporte_Nia_Nsa.TextBox4.Text = Label2.Text
                Reporte_Nia_Nsa.TextBox5.Text = Label4.Text
                Reporte_Nia_Nsa.Show()
            End If
        End If
    End Sub
    Private Function correlativo()
        Dim corre As Integer
        Dim Val_corr As String
        Dim func As New fnia
        Dim dts As New vnia
        Dim me12 As String
        me12 = Month(DateTimePicker1.Value)
        Val_corr = ""
        Select Case me12.Length
            Case "1" : dts.gmes = "0" & "" & me12
            Case "2" : dts.gmes = me12

        End Select
        dts.galmacen = "10"
        dts.gano = Label2.Text
        dts.gccia = Label4.Text
        corre = func.buscar_nia(dts) + 1
        Select Case corre.ToString.Length

            Case "1" : Val_corr = "00000" & "" & corre
            Case "2" : Val_corr = "0000" & "" & corre
            Case "3" : Val_corr = "000" & "" & corre
            Case "4" : Val_corr = "00" & "" & corre
            Case "5" : Val_corr = "0" & "" & corre
            Case "6" : Val_corr = corre
        End Select

        Return dts.gmes & Val_corr

    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("SE GENERARA LA NOTA DE INGRESO ESTA SEGURO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then

            insertar_nia()
        End If

    End Sub

    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim da As New DataTable
    Private Sub MOSTRAR()
        abrir()
        DataGridView1.DataSource = ""
        da.Clear()
        Dim calidad As String

        Dim estad As String
        estad = ""

        Dim cmd As New SqlDataAdapter("SELECT partida AS PARTIDA,des_articulo AS PRODUCTO,numero_rollos AS CALIDAD,ROLLOS AS ROLLOS,total AS KGS, numero_rollos as CLASIFICACION ,
orden,articulo,q.Op_3p FROM cab_ingresop c
inner join custom_vianny.dbo.Qanp300 q on q.ccia_3p ='01' and c.partida = q.regis_3p
WHERE seleccionado = 0ORDER BY PARTIDA", conx)
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(2).Width = 70
            DataGridView1.Columns(3).Width = 500
            DataGridView1.Columns(4).Width = 70
            DataGridView1.Columns(5).Width = 70
            DataGridView1.Columns(6).Width = 70
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).ReadOnly = True
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(8).ReadOnly = True
            DataGridView1.Columns(9).ReadOnly = True
            DataGridView1.Columns(10).ReadOnly = True
        Else
            DataGridView1.DataSource = ""
        End If
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 0 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            If Trim(DataGridView1.Rows(num1).Cells(8).Value) = "" Then

                Packing_nacional.TextBox1.Text = DataGridView1.Rows(num1).Cells(2).Value
                Packing_nacional.TextBox2.Text = DataGridView1.Rows(num1).Cells(4).Value
                Packing_nacional.TextBox3.Text = "10"
                Packing_nacional.Show()
            Else
                Packing_exportacion.TextBox1.Text = DataGridView1.Rows(num1).Cells(2).Value
                Packing_exportacion.TextBox2.Text = DataGridView1.Rows(num1).Cells(4).Value
                Packing_exportacion.Show()
            End If

        End If
    End Sub
End Class