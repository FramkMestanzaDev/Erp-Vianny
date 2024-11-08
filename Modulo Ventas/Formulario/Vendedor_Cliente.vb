Imports System.Data.SqlClient
Public Class Vendedor_Cliente
    Dim DT As New DataTable

    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE rpm_ven ='2' union all
select  '1' ,'TODOS'  ", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "alias_ven"
            ComboBox1.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box3()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE rpm_ven ='2'", conx)
            conn.Fill(ty3)
            ComboBox2.DataSource = ty3
            ComboBox2.DisplayMember = "alias_ven"
            ComboBox2.ValueMember = "codigo_ven"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FG As New fcliente
        Dim JK As New vcliente
        Dim ll As String
        If ComboBox1.Text = "" Then
            MsgBox("NO HAY NINGUN VENDEDOR SELECCIONADO")
        Else
            TextBox1.Text = ""
            TextBox2.Text = ""
            JK.gVendedor = ComboBox1.SelectedValue.ToString
            DT = FG.buscar_VENDEDOR_CLIENTE12(JK)

            If DT.Rows.Count > 0 Then
                DataGridView1.DataSource = DT
                DataGridView1.Columns(2).Width = 220
                DataGridView1.Columns(3).Width = 220
                DataGridView1.Columns(6).Width = 180
                DataGridView1.Columns(7).Width = 130
                DataGridView1.Columns(4).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Columns(4).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.Bisque
                DataGridView1.Columns(5).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Columns(6).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Columns(6).DefaultCellStyle.ForeColor = Color.White
            Else

                MsgBox("NO HAY INFORMACION PARA MOSTRAR")
            End If
        End If

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(2).Width = 220
                DataGridView1.Columns(3).Width = 220
                DataGridView1.Columns(6).Width = 180
                DataGridView1.Columns(7).Width = 130
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "RUC" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(2).Width = 220
                DataGridView1.Columns(3).Width = 220
                DataGridView1.Columns(6).Width = 180
                DataGridView1.Columns(7).Width = 130
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim hj As New ExportarCom
        hj.llenarExcel(DataGridView1)
    End Sub



    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.SetToolTip(PictureBox1, "EXPORTAR INFORMACION")
        ToolTip1.ToolTipTitle = "CLIENTES x VENDEDOR"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1 As Integer
            num1 = e.RowIndex.ToString


            Cambiar_Vendedor.TextBox1.Text = DataGridView1.Rows(num1).Cells(2).Value
            Cambiar_Vendedor.TextBox2.Text = DataGridView1.Rows(num1).Cells(3).Value
            Cambiar_Vendedor.TextBox3.Text = DataGridView1.Rows(num1).Cells(1).Value
            Cambiar_Vendedor.Show()


        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Label4.Visible = True
            ComboBox2.Visible = True
            Button2.Visible = True
        Else
            Label4.Visible = False
            ComboBox2.Visible = False
            Button2.Visible = False
        End If
    End Sub
    Dim Rsr2 As SqlDataReader
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim hj As New vcliente
        Dim jh As New fcliente
        Dim vendedor As String
        abrir()
        Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + ComboBox2.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            vendedor = Rsr2(0)

        End If

        Rsr2.Close()


        Dim cmd19 As New SqlCommand("update cliente set Vendedor =@vendedor ,fcom_fin=@fcom_fin  where COD_CLI  in(" + Trim(TextBox3.Text) + ")", conx)
        'cmd19.Parameters.AddWithValue("@ruc", Trim(TextBox3.Text))
        cmd19.Parameters.AddWithValue("@vendedor", vendedor)
        cmd19.Parameters.AddWithValue("@fcom_fin", DateTimePicker1.Value.Date)
        cmd19.ExecuteNonQuery()



        MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
        TextBox3.Text = ""
        Label4.Visible = False
        ComboBox2.Visible = False
        Button2.Visible = False
        CheckBox1.Checked = False

        Dim FG As New fcliente
        Dim JK As New vcliente
        Dim ll As String
        If ComboBox1.Text = "" Then
            MsgBox("NO HAY NINGUN VENDEDOR SELECCIONADO")
        Else
            'Select Case ComboBox1.Text

            '    Case "VSILVERIO" : JK.gVendedor = "0005"
            '    Case "GBALVIN" : JK.gVendedor = "0028"
            '    Case "GBEDON" : JK.gVendedor = "0010"
            '    Case "DBRAVO" : JK.gVendedor = "0023"
            '    Case "AMENDO" : JK.gVendedor = "0026"
            '    Case "VGRAUS" : JK.gVendedor = "0007"
            '    Case "WSALINAS" : JK.gVendedor = "0034"
            '    Case "MGRAUS" : JK.gVendedor = "0037"
            '    Case "LHUAYTA" : JK.gVendedor = "0039"
            '    Case "TODOS" : JK.gVendedor = 1
            'End Select
            JK.gVendedor = ComboBox1.SelectedValue.ToString
            DT = FG.buscar_VENDEDOR_CLIENTE12(JK)

            If DT.Rows.Count > 0 Then
                DataGridView1.DataSource = DT
                DataGridView1.Columns(2).Width = 220
                DataGridView1.Columns(3).Width = 220
                DataGridView1.Columns(6).Width = 180
                DataGridView1.Columns(7).Width = 130
                DataGridView1.Columns(4).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Columns(4).DefaultCellStyle.ForeColor = Color.White
                DataGridView1.Columns(5).DefaultCellStyle.BackColor = Color.Bisque
                DataGridView1.Columns(5).DefaultCellStyle.ForeColor = Color.Black
                DataGridView1.Columns(6).DefaultCellStyle.BackColor = Color.DarkKhaki
                DataGridView1.Columns(6).DefaultCellStyle.ForeColor = Color.White
            Else

                MsgBox("NO HAY INFORMACION PARA MOSTRAR")
            End If
        End If
    End Sub

    Private Sub Vendedor_Cliente_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box2()
        llenar_combo_box3()

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 2 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1 As Integer
            num1 = e.RowIndex.ToString
            If InStr(TextBox3.Text, Trim(Convert.ToString(DataGridView1.Rows(num1).Cells(2).Value))) Then
                Dim respuesta As DialogResult

                respuesta = MessageBox.Show("EL RUC YA ESTA AGREGADO, DESEA ELIMINARLO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    TextBox3.Text = Replace(TextBox3.Text, ",'" & Trim(Convert.ToString(DataGridView1.Rows(num1).Cells(2).Value)) & "'", "")
                    TextBox3.Text = Replace(TextBox3.Text, "'" & Trim(Convert.ToString(DataGridView1.Rows(num1).Cells(2).Value)) & "',", "")
                    TextBox3.Text = Replace(TextBox3.Text, "'" & Trim(Convert.ToString(DataGridView1.Rows(num1).Cells(2).Value)) & "'", "")
                End If


                'MsgBox("EL RUC YA ESTA AGREGADO")
            Else
                    If Trim(Convert.ToString(TextBox3.Text)) = "" Then
                    TextBox3.Text = "'" & Trim(Convert.ToString(DataGridView1.Rows(num1).Cells(2).Value)) & "'"
                Else
                    TextBox3.Text = Convert.ToString(Trim(TextBox3.Text)) + "," + "'" & Trim(Convert.ToString(DataGridView1.Rows(num1).Cells(2).Value)) & "'"
                End If
            End If
        End If

    End Sub
End Class