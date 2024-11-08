Imports System.Data.SqlClient
Public Class Lista_Solicitud
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim Rsr2, Rsr22, Rsr3, Rsr34, Rsr341, Rsr346, Rsr29, Rsr244 As SqlDataReader

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'DT.DefaultView.RowFilter = $"OP LIKE '{TextBox1.Text}%'"
        buscar()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim DA As New DataTable
    Private Sub Lista_Solicitud_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
        abrir()
        'Dim sql102213 As String = "SELECT id_requeminieto as REQUERIMIENTO, op AS OP,cliente AS CLIENTE FROM requerimiento_avios_cabecera WHERE  periodo ='" + Label2.Text + "' AND SUBSTRING( id_requeminieto,1,2) = '" + Label3.Text + "'"
        'Dim cmd102213 As New SqlCommand(sql102213, conx)
        'Rsr3 = cmd102213.ExecuteReader()
        'Dim i51 As Integer
        'i51 = 0
        'While Rsr3.Read()
        '    DataGridView1.Rows.Add()
        '    DataGridView1.Rows(i51).Cells(0).Value = Rsr3(0)
        '    DataGridView1.Rows(i51).Cells(1).Value = Rsr3(1)
        '    DataGridView1.Rows(i51).Cells(2).Value = Rsr3(2)
        '    i51 = i51 + 1


        'End While

        'Rsr3.Close()


        Dim cmd As New SqlDataAdapter("SELECT id_requeminieto as REQUERIMIENTO, op AS OP,corte as CORTE,po AS PO FROM requerimiento_avios_cabecera WHERE  periodo ='" + Label2.Text + "' AND SUBSTRING( id_requeminieto,1,2) = '" + Label3.Text + "'", conx)



        cmd.Fill(DA)
        DataGridView1.DataSource = DA
        DataGridView1.Columns(0).Width = 120
        DataGridView1.Columns(1).Width = 120
        DataGridView1.Columns(2).Width = 120
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(DA.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(0).Width = 120
                DataGridView1.Columns(1).Width = 120
                DataGridView1.Columns(2).Width = 420
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub
    Dim nombre_columna As String
    Sub filtrar(DataGrid As DataGridView, nombre_columna As String)

        Dim DataTable As New DataTable()
        DataTable = DataGrid.DataSource


        Dim vista_datos As New DataView()
        vista_datos = DataTable.DefaultView
        Dim filtro As String
        filtro = "CONVERT(" + "(" + nombre_columna + ")" + ", System.String) LIKE '%" + TextBox1.Text.Trim() + "%'"


        vista_datos.RowFilter = filtro
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Formato_solic_avios.TextBox2.Text = Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(0).Value, 3, 8)
        Formato_solic_avios.TextBox2.Focus()
        SendKeys.Send("{ENTER}")
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        'If DataGridView1.Columns(e.ColumnIndex).HeaderText = "REQUERIMIENTO" Then
        '    Formato_solic_avios.TextBox2.Text = Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(0).Value, 3, 8)
        '    Formato_solic_avios.TextBox2.Focus()
        '    SendKeys.Send("{ENTER}")
        '    Me.Close()
        'Else
        '    If DataGridView1.Columns(e.ColumnIndex).HeaderText = "OP" Then
        '        Formato_solic_avios.TextBox2.Text = Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(0).Value, 3, 8)
        '        Formato_solic_avios.TextBox2.Focus()
        '        SendKeys.Send("{ENTER}")
        '        Me.Close()
        '    Else
        '        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CORTE" Then
        '            Formato_solic_avios.TextBox2.Text = Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(0).Value, 3, 8)
        '            Formato_solic_avios.TextBox2.Focus()
        '            SendKeys.Send("{ENTER}")
        '            Me.Close()
        '        Else

        '            If DataGridView1.Columns(e.ColumnIndex).HeaderText = "PO" Then
        '                Formato_solic_avios.TextBox2.Text = Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(0).Value, 3, 8)
        '                Formato_solic_avios.TextBox2.Focus()
        '                SendKeys.Send("{ENTER}")
        '                Me.Close()
        '            End If
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode = Keys.Enter Then

            Formato_solic_avios.TextBox2.Text = Microsoft.VisualBasic.Mid(Convert.ToString(DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(0).Value), 3, 8)
            e.Handled = True
            Formato_solic_avios.TextBox2.Focus()
            SendKeys.Send("{ENTER}")
            Me.Close()
        End If
    End Sub

    Private Sub DataGridView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DataGridView1.KeyPress
        'If (e.KeyChar = (Convert.ToChar(Keys.Enter))) Then


        'End If        'If (e.KeyChar = (Convert.ToChar(Keys.Enter))) Then


        'End If
    End Sub
End Class