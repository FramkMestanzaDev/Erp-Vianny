Imports System.Data.SqlClient
Public Class DetalleOrden
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim Rsr21 As SqlDataReader
    Dim da1 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub DetalleOrden_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            Me.BeginInvoke(New MethodInvoker(Sub()
                                                 mostrar()
                                             End Sub))
    End Sub
    Sub mostrar()
        da1.Rows.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select item_3a as ITEMS,pedido_3a AS PEDIDO,linea_17 AS CODIGO,cg.nomb_17 AS PRODUCTO,cant_3a AS 'CANT PRES',unide2_3a AS 'UND PRES', cante_3a AS CANT,unid_17 AS UND,ncom_3a AS OC,case when '" + Label7.Text + "' = '1' THEN preun1_3a else preun2_3a END MON ,Cast( 0 as numeric) AS INGRESAR,0 AS valor from custom_vianny.dbo.lap0300 lp
                                        left join custom_vianny.dbo.cag1700 cg on lp.ccia_3a = cg.ccia and lp.linea_3a = cg.linea_17
                                        where ccia_3a ='" + Label1.Text + "' and ncom_3a ='" + Label2.Text + "' and flag_3a ='1' and percep_3a in(0)", conx)
        cmd.Fill(da1)
        If da1.Rows.Count > 0 Then
            DataGridView1.DataSource = da1
            DataGridView1.Columns(0).ReadOnly = False
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).ReadOnly = True
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(8).ReadOnly = True
            DataGridView1.Columns(9).ReadOnly = True
            DataGridView1.Columns(10).ReadOnly = True
            DataGridView1.Columns(11).ReadOnly = True
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(1).Width = 48
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).Width = 132
            DataGridView1.Columns(4).Width = 350
            DataGridView1.Columns(5).Width = 75
            DataGridView1.Columns(6).Width = 75
            DataGridView1.Columns(7).Width = 75
            DataGridView1.Columns(8).Width = 75
            DataGridView1.Columns(9).Width = 65
            DataGridView1.Columns(10).Width = 65
            DataGridView1.Columns(11).Width = 75
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    'Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
    '    If DataGridView1.IsCurrentCellDirty Then
    '        DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
    '    End If
    'End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick



    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Convert.ToDouble(TextBox1.Text) <= Convert.ToDouble(TextBox2.Text) Then


            Dim filasSeleccionadas As New List(Of DataGridViewRow)
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Convert.ToBoolean(row.Cells("S").Value) = True Then
                    filasSeleccionadas.Add(row)
                End If
            Next

            ' Llamamos al método en el formulario NotaIngreso para pasar los datos
            If filasSeleccionadas.Count > 0 Then
                If Label8.Text.ToString().Trim() = "08" Or Label8.Text.ToString().Trim() = "10" Then
                    NotaIngreso.abrir()
                    NotaIngreso.TextBox7.Text = Label4.Text
                    NotaIngreso.Label11.Text = Label3.Text
                    NotaIngreso.Label14.Text = Label1.Text

                    NotaIngreso.TextBox4.Text = DateTimePicker1.Value.Month
                    'NotaIngreso.llenar_combo_alamcenes("1")
                    NotaIngreso.EstablecerAlmacen(Label8.Text.ToString().Trim())
                    Select Case NotaIngreso.TextBox4.Text.Length
                        Case "1" : NotaIngreso.TextBox4.Text = "0" & DateTimePicker1.Value.Month
                        Case "2" : NotaIngreso.TextBox4.Text = NotaIngreso.TextBox4.Text
                    End Select
                    NotaIngreso.correlativo()
                    NotaIngreso.ComboBox2.SelectedIndex = 1
                    NotaIngreso.ComboBox4.SelectedIndex = 0
                    'NotaIngreso.TextBox5.Focus()
                    'SendKeys.Send("{ENTER}")
                    NotaIngreso.TextBox9.Select()
                    NotaIngreso.ComboBox2.SelectedIndex = 1


                    NotaIngreso.RecibirDatos(filasSeleccionadas)
                    NotaIngreso.Label16.Text = "1"
                    NotaIngreso.Show(Me)
                    NotaIngreso.TextBox8.Text = Label9.Text
                    NotaIngreso.TextBox10.Text = Label10.Text
                    NotaIngreso.TextBox12.Text = Label11.Text
                    NotaIngreso.Button6.Enabled = True
                    NotaIngreso.compras()
                Else
                    NiaHilo.abrir()
                    NiaHilo.TextBox7.Text = Label4.Text
                    NiaHilo.Label11.Text = Label3.Text
                    NiaHilo.Label13.Text = Label1.Text
                    NiaHilo.TextBox14.Text = 2

                    NiaHilo.TextBox4.Text = DateTimePicker1.Value.Month
                    NiaHilo.TextBox8.Text = Label9.Text
                    NiaHilo.TextBox10.Text = Label10.Text
                    NiaHilo.TextBox12.Text = Label11.Text
                    'NotaIngreso.llenar_combo_alamcenes("1")
                    'MsgBox(Label8.Text.ToString().Trim())
                    If Label8.Text.ToString().Trim() = "03" Then
                        NiaHilo.TextBox16.Text = 2
                        NiaHilo.EstablecerAlmacen(Label8.Text.ToString().Trim())
                    Else
                        NiaHilo.TextBox16.Text = 5
                        NiaHilo.LLenarAlmacenAviosSuministros(5, Label8.Text.ToString().Trim())
                    End If

                    Select Case NiaHilo.TextBox4.Text.Length
                            Case "1" : NiaHilo.TextBox4.Text = "0" & DateTimePicker1.Value.Month
                            Case "2" : NiaHilo.TextBox4.Text = NiaHilo.TextBox4.Text
                        End Select
                        'NiaHilo.ComboBox1.SelectedIndex = 0

                        NiaHilo.correlativo()
                        NiaHilo.ComboBox2.SelectedIndex = 1
                        NiaHilo.RecibirDatos(filasSeleccionadas)
                        NiaHilo.Label16.Text = "1"

                        NiaHilo.Show(Me)
                        NiaHilo.ComboBox1.Enabled = True
                        NiaHilo.llenar_combo_box2(Label8.Text.ToString().Trim())
                        'NiaHilo.TextBox5.Focus()
                        'SendKeys.Send("{ENTER}")

                        NiaHilo.compras()
                    End If

                    Else
                MessageBox.Show("Por favor, seleccione al menos un ítem.")
            End If
        Else
            MsgBox("LA CANTIDAS INGRESADA EXCEDE EN EL TOTAL DE LA ORDEN DE COMPRA")
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        If e.RowIndex >= 0 AndAlso e.ColumnIndex = 11 Then

            Dim isCellChecked As Boolean = Convert.ToBoolean(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
            If isCellChecked Then
                'MsgBox("Ingrese lo llegado de acuerdo a la guia-remision")
                DataGridView1.Rows(e.RowIndex).Cells(11).ReadOnly = False
                'DataGridView1.CurrentCell = DataGridView1(11, e.RowIndex)
                DataGridView1.BeginEdit(True)
            Else
                DataGridView1.Rows(e.RowIndex).Cells(11).ReadOnly = True
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = 11 Then
            If Convert.ToDecimal(DataGridView1.Rows(e.RowIndex).Cells(11).Value) > Convert.ToDecimal(DataGridView1.Rows(e.RowIndex).Cells(7).Value) Then
                MsgBox("LA CANTIDAD A INGRESAR ES MAYOR A LA ORDEN DE COMPRA - COORDINE CON EL AREA DE COMPRAS LA MODIFICACION DE LA ORDEN")
                DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0
                DataGridView1.Rows(e.RowIndex).Cells(12).Value = 0
            Else
                Dim suma As Decimal = 0
                Dim total As Decimal = 0
                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim isCellChecked As Boolean = Convert.ToBoolean(DataGridView1.Rows(i).Cells(0).Value)
                    If isCellChecked Then
                        If Label12.Text.ToString().Trim() = "NO" Then
                            total = Convert.ToDecimal(DataGridView1.Rows(i).Cells(11).Value) * Convert.ToDecimal(DataGridView1.Rows(i).Cells(10).Value) * 1.18
                        Else
                            total = Convert.ToDecimal(DataGridView1.Rows(i).Cells(11).Value) * Convert.ToDecimal(DataGridView1.Rows(i).Cells(10).Value)
                        End If
                        suma = suma + total
                    End If
                Next
                TextBox2.Text = suma.ToString("N2")
                If Convert.ToDecimal(DataGridView1.Rows(e.RowIndex).Cells(11).Value) = Convert.ToDecimal(DataGridView1.Rows(e.RowIndex).Cells(7).Value) Then
                    DataGridView1.Rows(e.RowIndex).Cells(12).Value = 2
                Else
                    DataGridView1.Rows(e.RowIndex).Cells(12).Value = 1
                End If
            End If
        End If

    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim txt As TextBox = TryCast(e.Control, TextBox)
        If DataGridView1.CurrentCell.ColumnIndex = 11 Then
            RemoveHandler txt.KeyPress, AddressOf txt_KeyPress
            AddHandler txt.KeyPress, AddressOf txt_KeyPress
        End If
    End Sub
    Private Sub txt_KeyPress(sender As Object, e As KeyPressEventArgs)
        Dim txt As TextBox = CType(sender, TextBox)
        ' Permitir solo números, teclas de control (backspace, etc.) y un solo punto decimal
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "." Then
            e.Handled = True ' Ignora el carácter si no es un número o punto
        End If
        ' Permitir un solo punto decimal
        If e.KeyChar = "." AndAlso txt.Text.Contains(".") Then
            e.Handled = True ' Ignora si ya existe un punto decimal en el texto
        End If
    End Sub
End Class