Imports System.Data.SqlClient
Public Class Data_Bancos
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

    Dim da As New DataTable
    Dim ar As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        da.Clear()
        ComboBox2.Items.Clear()

        If ComboBox1.Text = "SOLES" Then
            ar = 1
        Else
            ar = 2
        End If
        Dim DT, DT1 As String
        DT = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        DT1 = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
        abrir()
        Dim cmd As New SqlDataAdapter("exec data_flujo_caja '" + Trim(Label7.Text) + "','" + Trim(TextBox4.Text) + "','" + DT + "','" + DT1 + "','" + Trim(TextBox2.Text) + "','" + Trim(TextBox3.Text) + "','" + ar + "'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Frozen = True
        DataGridView1.Columns(1).Frozen = True
        DataGridView1.Columns(16).Visible = True
        DataGridView1.Columns(17).Visible = True
        DataGridView1.Columns(18).Visible = False
        DataGridView1.Columns(19).Visible = False
        DataGridView1.Columns(20).Visible = False
        DataGridView1.Columns(21).Visible = False
        DataGridView1.Columns(22).Visible = False
        Dim u As Integer
        u = DataGridView1.Rows.Count
        Label13.Text = 2
        For i = 0 To u - 1
            DataGridView1.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(19).Value
            DataGridView1.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(20).Value
        Next
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(3).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(4).Name.ToString)
            ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(5).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(6).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(7).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(8).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(9).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(10).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(11).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(12).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(13).Name.ToString)
        ComboBox2.Items.Add(Me.DataGridView1.Columns.Item(14).Name.ToString)
    End Sub

    Private Sub Data_Bancos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0

    End Sub

    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "GLOSA" & " like '%" & TextBox5.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(16).Visible = True
                DataGridView1.Columns(17).Visible = True
                DataGridView1.Columns(18).Visible = False
                DataGridView1.Columns(19).Visible = False
                DataGridView1.Columns(20).Visible = False
                DataGridView1.Columns(21).Visible = False
                DataGridView1.Columns(22).Visible = False
                Dim u As Integer
                u = DataGridView1.Rows.Count
                Label13.Text = 2
                For i = 0 To u - 1
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(19).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(20).Value
                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then

            For i = 0 To DataGridView1.Columns.Count - 1

                DataGridView1.Columns.Item(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next

        Else
            If e.ColumnIndex = 0 Then

                ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
                Dim num1, num2 As Integer
                num1 = e.RowIndex.ToString
                num2 = e.ColumnIndex.ToString
                det_tesoreria.Label1.Text = num1
                det_tesoreria.Label2.Text = num2
                det_tesoreria.Label4.Text = 0
                det_tesoreria.Show()



            End If
            Label13.Text = 0
            If Label13.Text = 0 Then
                Dim J As Integer

                J = DataGridView1.Rows.Count
                For I = 0 To J - 1
                    If Me.DataGridView1.CurrentRow.Index.ToString() = I Then

                        DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.DarkKhaki
                    Else

                        DataGridView1.Rows(I).DefaultCellStyle.BackColor = Color.White
                        DataGridView1.Rows(I).Cells(1).Style.BackColor = Color.Honeydew
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        Dim P As Integer
        P = DataGridView1.Rows.Count
        For i = 0 To P - 1
            If Trim(DataGridView1.Rows(i).Cells(2).Value) <> "" And Len(Trim(DataGridView1.Rows(i).Cells(19).Value)) = 0 Then
                Dim fun As New VSTOCK_PAR
                Dim func As New FSTOCK_PAR
                fun.gocompra = DataGridView1.Rows(i).Cells(2).Value
                fun.gCCIA = Label7.Text
                fun.gcuenta = DataGridView1.Rows(i).Cells(6).Value
                fun.gregistro = DataGridView1.Rows(i).Cells(5).Value
                fun.gperiodo = TextBox4.Text
                fun.gdh = DataGridView1.Rows(i).Cells(7).Value
                fun.gccom = DataGridView1.Rows(i).Cells(18).Value
                func.actualizar_data_bancos(fun)
                'Dim cmd15 As New SqlCommand("update custom_vianny.dbo.cac3p00 set ocompra_3a =@ocompra where  ccia_3a =@ccia and cuen_3a =@cuenta and ncom_3a =@registro  and cperiodo_3a =@periodo and codh_3a =@dh AND ccom_3a =@ccom", conx)
                'cmd15.CommandTimeout = 450
                'cmd15.Parameters.AddWithValue("@ocompra", DataGridView1.Rows(i).Cells(2).Value)
                ''cmd15.Parameters.AddWithValue("@ccom", DataGridView1.Rows(i).Cells(2).Value)
                'cmd15.Parameters.AddWithValue("@ccia", Label7.Text)
                'cmd15.Parameters.AddWithValue("@cuenta", DataGridView1.Rows(i).Cells(6).Value)
                'cmd15.Parameters.AddWithValue("@registro", DataGridView1.Rows(i).Cells(5).Value)
                'cmd15.Parameters.AddWithValue("@periodo", TextBox4.Text)
                'cmd15.Parameters.AddWithValue("@dh", DataGridView1.Rows(i).Cells(7).Value)
                'cmd15.Parameters.AddWithValue("@ccom", DataGridView1.Rows(i).Cells(18).Value)
                'cmd15.ExecuteNonQuery()
            End If
        Next
        da.Rows.Clear()
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        MsgBox("SE AGREGO LA INFORMACION CORRECTAMENTE")
        DataGridView1.DataSource = ""
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "CCOM" Then
            Me.DataGridView1.Columns("CCOM").Visible = False
            ComboBox2.Items.Remove("CCOM")
            ComboBox3.Items.Add("CCOM")
        End If
        If ComboBox2.Text = "CDOC" Then
            Me.DataGridView1.Columns("CDOC").Visible = False
            ComboBox2.Items.Remove("CDOC")
            ComboBox3.Items.Add("CDOC")
        End If
        If ComboBox2.Text = "NCOM" Then
            Me.DataGridView1.Columns("NCOM").Visible = False
            ComboBox2.Items.Remove("NCOM")
            ComboBox3.Items.Add("NCOM")
        End If
        If ComboBox2.Text = "CUENTA" Then
            Me.DataGridView1.Columns("CUENTA").Visible = False
            ComboBox2.Items.Remove("CUENTA")
            ComboBox3.Items.Add("CUENTA")
        End If
        If ComboBox2.Text = "OPERACION" Then
            Me.DataGridView1.Columns("OPERACION").Visible = False
            ComboBox2.Items.Remove("OPERACION")
            ComboBox3.Items.Add("OPERACION")
        End If
        If ComboBox2.Text = "FECHA" Then
            Me.DataGridView1.Columns("FECHA").Visible = False
            ComboBox2.Items.Remove("FECHA")
            ComboBox3.Items.Add("FECHA")
        End If
        If ComboBox2.Text = "FECHA_EMISION" Then
            Me.DataGridView1.Columns("FECHA_EMISION").Visible = False
            ComboBox2.Items.Remove("FECHA_EMISION")
            ComboBox3.Items.Add("FECHA_EMISION")
        End If
        If ComboBox2.Text = "FECHA_UPDATE" Then
            Me.DataGridView1.Columns("FECHA_UPDATE").Visible = False
            ComboBox2.Items.Remove("FECHA_UPDATE")
            ComboBox3.Items.Add("FECHA_UPDATE")
        End If
        If ComboBox2.Text = "MONEDA" Then
            Me.DataGridView1.Columns("MONEDA").Visible = False
            ComboBox2.Items.Remove("MONEDA")
            ComboBox3.Items.Add("MONEDA")
        End If
        If ComboBox2.Text = "FECHA_UPDATE" Then
            Me.DataGridView1.Columns("FECHA_UPDATE").Visible = False
            ComboBox2.Items.Remove("FECHA_UPDATE")
            ComboBox3.Items.Add("FECHA_UPDATE")
        End If
        If ComboBox2.Text = "GLOSA" Then
            Me.DataGridView1.Columns("GLOSA").Visible = False
            ComboBox2.Items.Remove("GLOSA")
            ComboBox3.Items.Add("GLOSA")
        End If
        If ComboBox2.Text = "DEBE" Then
            Me.DataGridView1.Columns("DEBE").Visible = False
            ComboBox2.Items.Remove("DEBE")
            ComboBox3.Items.Add("DEBE")
        End If
        If ComboBox2.Text = "HABER" Then
            Me.DataGridView1.Columns("HABER").Visible = False
            ComboBox2.Items.Remove("HABER")
            ComboBox3.Items.Add("HABER")
        End If
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.Text = "CCOM" Then
            Me.DataGridView1.Columns("CCOM").Visible = True
            ComboBox3.Items.Remove("CCOM")
            ComboBox2.Items.Add("CCOM")
        End If
        If ComboBox3.Text = "CDOC" Then
            Me.DataGridView1.Columns("CDOC").Visible = True
            ComboBox3.Items.Remove("CDOC")
            ComboBox2.Items.Add("CDOC")
        End If
        If ComboBox3.Text = "NCOM" Then
            Me.DataGridView1.Columns("NCOM").Visible = True
            ComboBox3.Items.Remove("NCOM")
            ComboBox2.Items.Add("NCOM")
        End If
        If ComboBox3.Text = "CUENTA" Then
            Me.DataGridView1.Columns("CUENTA").Visible = True
            ComboBox3.Items.Remove("CUENTA")
            ComboBox2.Items.Add("CUENTA")
        End If
        If ComboBox3.Text = "OPERACION" Then
            Me.DataGridView1.Columns("OPERACION").Visible = True
            ComboBox3.Items.Remove("OPERACION")
            ComboBox2.Items.Add("OPERACION")
        End If
        If ComboBox3.Text = "FECHA" Then
            Me.DataGridView1.Columns("FECHA").Visible = True
            ComboBox3.Items.Remove("FECHA")
            ComboBox2.Items.Add("FECHA")
        End If
        If ComboBox3.Text = "FECHA_EMISION" Then
            Me.DataGridView1.Columns("FECHA_EMISION").Visible = True
            ComboBox3.Items.Remove("FECHA_EMISION")
            ComboBox2.Items.Add("FECHA_EMISION")
        End If
        If ComboBox3.Text = "FECHA_UPDATE" Then
            Me.DataGridView1.Columns("FECHA_UPDATE").Visible = True
            ComboBox3.Items.Remove("FECHA_UPDATE")
            ComboBox2.Items.Add("FECHA_UPDATE")
        End If
        If ComboBox3.Text = "GLOSA" Then
            Me.DataGridView1.Columns("GLOSA").Visible = True
            ComboBox3.Items.Remove("GLOSA")
            ComboBox2.Items.Add("GLOSA")
        End If
        If ComboBox3.Text = "DEBE" Then
            Me.DataGridView1.Columns("DEBE").Visible = True
            ComboBox3.Items.Remove("DEBE")
            ComboBox2.Items.Add("DEBE")
        End If
        If ComboBox3.Text = "HABER" Then
            Me.DataGridView1.Columns("HABER").Visible = True
            ComboBox3.Items.Remove("HABER")
            ComboBox2.Items.Add("HABER")
        End If

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        buscar()
    End Sub

    Dim Rsr20 As SqlDataReader
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CODIGO" Then
            Select Case Len(Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value))

                Case "1" : DataGridView1.Rows(e.RowIndex).Cells(2).Value = "0000000" & "" & DataGridView1.Rows(e.RowIndex).Cells(2).Value
                Case "2" : DataGridView1.Rows(e.RowIndex).Cells(2).Value = "000000" & "" & DataGridView1.Rows(e.RowIndex).Cells(2).Value
                Case "3" : DataGridView1.Rows(e.RowIndex).Cells(2).Value = "00000" & "" & DataGridView1.Rows(e.RowIndex).Cells(2).Value
                Case "4" : DataGridView1.Rows(e.RowIndex).Cells(2).Value = "0000" & "" & DataGridView1.Rows(e.RowIndex).Cells(2).Value
                Case "5" : DataGridView1.Rows(e.RowIndex).Cells(2).Value = "000" & "" & DataGridView1.Rows(e.RowIndex).Cells(2).Value
                Case "6" : DataGridView1.Rows(e.RowIndex).Cells(2).Value = "00" & "" & DataGridView1.Rows(e.RowIndex).Cells(2).Value
                Case "7" : DataGridView1.Rows(e.RowIndex).Cells(2).Value = "0" & "" & DataGridView1.Rows(e.RowIndex).Cells(2).Value
                Case "8" : DataGridView1.Rows(e.RowIndex).Cells(2).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value
            End Select
            Dim sql1020 As String = "select NOMBRE from RUBROS_TESORERIA where CODIGO = '" + DataGridView1.Rows(e.RowIndex).Cells(2).Value + "'"
            Dim cmd1020 As New SqlCommand(sql1020, conx)
            Rsr20 = cmd1020.ExecuteReader()
            If Rsr20.Read() Then
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = Rsr20(0)
            Else
                MsgBox("EL CODIGO INGRESADO NO EXISTE")
                DataGridView1.Rows(e.RowIndex).Cells(2).Value = ""
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = ""
            End If
            Rsr20.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim GH As New Exportar
        GH.llenarExcel(DataGridView1)
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        buscar1()
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "NOMBRE" & " like '%" & TextBox6.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(21).Visible = False
                DataGridView1.Columns(22).Visible = False
                DataGridView1.Columns(18).Visible = False
                DataGridView1.Columns(19).Visible = False
                DataGridView1.Columns(20).Visible = False
                Dim u As Integer
                u = DataGridView1.Rows.Count

                For i = 0 To u - 1
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(19).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(20).Value
                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CDOC" & " like '%" & TextBox7.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(21).Visible = False
                DataGridView1.Columns(22).Visible = False
                DataGridView1.Columns(18).Visible = False
                DataGridView1.Columns(19).Visible = False
                DataGridView1.Columns(20).Visible = False
                Dim u As Integer
                u = DataGridView1.Rows.Count

                For i = 0 To u - 1
                    DataGridView1.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(19).Value
                    DataGridView1.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(20).Value
                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged


    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        buscar2()
    End Sub
End Class