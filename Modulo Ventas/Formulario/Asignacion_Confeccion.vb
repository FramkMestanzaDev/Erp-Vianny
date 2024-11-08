Imports System.Data.SqlClient
Public Class Asignacion_Confeccion
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
    Dim loDataTable As New DataTable
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            det_ns_Prodc.Label2.Text = TextBox1.Text
            det_ns_Prodc.Label3.Text = 9
            Select Case TextBox1.Text
                Case "0300" : det_ns_Prodc.Label4.Text = "1"
                Case "0400" : det_ns_Prodc.Label4.Text = "2"
                Case "0600" : det_ns_Prodc.Label4.Text = "3"
                Case "0500" : det_ns_Prodc.Label4.Text = "4"
            End Select
            det_ns_Prodc.Label5.Text = "01"
            det_ns_Prodc.Show()
        End If
    End Sub
    Sub llenar_combo_box3()

        Try
            loDataTable.Rows.Clear()
            Dim lsQuery As String = "select ocorte_3cg from custom_vianny.dbo.Qag300Cc where pedido_3cg ='" + Trim(TextBox1.Text) + "'"

            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            ComboBox1.DataSource = loDataTable

            ComboBox1.DisplayMember = "ocorte_3cg"
            ComboBox1.ValueMember = "ocorte_3cg"
            ComboBox1.DropDownWidth = 90
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Dim Rsr210, Rsr21, Rsr2105 As SqlDataReader
    Sub mostrar_informacion()
        Dim i51 As Integer


        Dim sql1021 As String = "select  producto AS PRODUCTO,po AS PO,op AS OP,corte   from requerimiento_avios_cabecera where estado IN (0,1,2)  and area=1"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr21 = cmd1021.ExecuteReader()
        Dim i5 As Integer
        i5 = 0
        While Rsr21.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i5).Cells(0).Value = Rsr21(0)
            DataGridView1.Rows(i5).Cells(1).Value = Rsr21(1)
            DataGridView1.Rows(i5).Cells(2).Value = Rsr21(2)
            DataGridView1.Rows(i5).Cells(3).Value = Rsr21(3)
            i5 = i5 + 1
        End While
        Rsr21.Close()

        i51 = DataGridView1.Rows.Count
        For i = 0 To i51 - 1
            Dim sql10215 As String = "select ruc,proveedor from asignacion_cos_aca where area ='01' and op='" + Trim(DataGridView1.Rows(i).Cells(2).Value) + "' and corte ='" + Trim(DataGridView1.Rows(i).Cells(3).Value) + "'"

            Dim cmd10215 As New SqlCommand(sql10215, conx)
            Rsr2105 = cmd10215.ExecuteReader()

            If Rsr2105.Read() Then
                DataGridView1.Rows(i).Cells(4).Value = Rsr2105(0)
            Else
                Rsr2105.Close()
                Dim sql10214 As String = "SELECT distinct fich_3,c.nomb_10 FROM custom_vianny.dbo.Qap3000 q 
                    inner join custom_vianny.dbo.Qag3000 g on q.Empr_3a = g.Empr_3 and q.Ano_3a=g.Ano_3 and q.talm_3a =g.talm_3 and q.ccom_3a =g.ccom_3 and q.ncom_3a = g.ncom_3  
                    left join custom_vianny.dbo.cag1000 c on q.Empr_3a = c.ccia And g.fich_3 = c.fich_10
                       where pedido_3a ='" + Trim(DataGridView1.Rows(i).Cells(2).Value) + "' and q.talm_3a ='0400'"
                Dim cmd10214 As New SqlCommand(sql10214, conx)
                Rsr210 = cmd10214.ExecuteReader()

                If Rsr210.Read() Then

                    DataGridView1.Rows(i).Cells(4).Value = Rsr210(0)

                Else
                    DataGridView1.Rows(i).Cells(4).Value = ""

                End If
                Rsr210.Close()
            End If
            Rsr2105.Close()

        Next
        For i2 = 0 To i51 - 1
            If Trim(DataGridView1.Rows(i2).Cells(4).Value) = "" Then

            Else
                DataGridView1.Rows(i2).Visible = False
            End If
        Next
    End Sub
    Private Sub Asignacion_Confeccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box3()
        mostrar_informacion()
        TextBox2.Select()


    End Sub
    Dim Rsr3 As SqlDataReader
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(TextBox1.Text) = "" Or Trim(TextBox2.Text) = "" Then
            MsgBox("FALTA INGRESAR LA OP O EL TALLER")

        Else
            Dim CORTE As String
            If CheckBox1.Checked = True Then
                CORTE = ComboBox1.SelectedValue.ToString
            Else
                CORTE = ""
            End If

            Dim sql102213 As String = "select COUNT(op),proveedor from asignacion_cos_aca where area='01' and  op= '" + Trim(TextBox1.Text) + "' and corte ='" + CORTE + "' GROUP BY proveedor"
            Dim cmd102213 As New SqlCommand(sql102213, conx)
            Rsr3 = cmd102213.ExecuteReader()
            If Rsr3.Read() = True Then

                If Rsr3(0) > 0 Then
                    MsgBox("YA EXISTE UN REGISTRO CON LA OP Y EL CORTE INGRESADOS ASIGNADOS  TALLER" + "      " + Rsr3(1))
                    Dim respuesta As DialogResult
                    respuesta = MessageBox.Show("DESEAS REASIGNARLO LA OP?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    Rsr3.Close()
                    If (respuesta = Windows.Forms.DialogResult.Yes) Then
                        Dim cmd16 As New SqlCommand("UPDATE asignacion_cos_aca SET proveedor =@proveedor,ruc=@ruc,Fecha=@Fecha WHERE area = @area AND op=@op AND corte =@corte", conx)
                        cmd16.Parameters.AddWithValue("@area", "01")
                        cmd16.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
                        If CheckBox1.Checked = True Then
                            cmd16.Parameters.AddWithValue("@corte", ComboBox1.SelectedValue.ToString)
                        Else
                            cmd16.Parameters.AddWithValue("@corte", "")
                        End If

                        cmd16.Parameters.AddWithValue("@ruc", Trim(Label3.Text))
                        cmd16.Parameters.AddWithValue("@proveedor", Trim(TextBox2.Text))
                        cmd16.Parameters.AddWithValue("@Fecha", DateTimePicker1.Value)
                        cmd16.ExecuteNonQuery()
                        MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")

                    End If

                End If
            Else
                Rsr3.Close()
                Dim cmd16 As New SqlCommand("insert into asignacion_cos_aca (area,op,corte,ruc,proveedor,Fecha) values (@area,@op,@corte,@ruc,@proveedor,@Fecha)", conx)
                cmd16.Parameters.AddWithValue("@area", "01")
                cmd16.Parameters.AddWithValue("@op", Trim(TextBox1.Text))
                If CheckBox1.Checked = True Then
                    cmd16.Parameters.AddWithValue("@corte", ComboBox1.SelectedValue.ToString)
                Else
                    cmd16.Parameters.AddWithValue("@corte", "")
                End If

                cmd16.Parameters.AddWithValue("@ruc", Trim(Label3.Text))
                cmd16.Parameters.AddWithValue("@proveedor", Trim(TextBox2.Text))
                cmd16.Parameters.AddWithValue("@Fecha", DateTimePicker1.Value)
                cmd16.ExecuteNonQuery()
                MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
            End If

            Rsr3.Close()
            mostrar_informacion()
            TextBox1.Text = ""
            TextBox2.Text = ""
            ComboBox1.Text = ""
            TextBox2.Select()

            'Me.Close()
            'Me.Show()
        End If



    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Select Case TextBox1.Text.Length

            Case "1" : TextBox1.Text = "010000000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "01000000" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "0100000" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = "010000" & "" & TextBox1.Text
            Case "5" : TextBox1.Text = "01000" & "" & TextBox1.Text
            Case "6" : TextBox1.Text = "0100" & "" & TextBox1.Text
            Case "7" : TextBox1.Text = "010" & "" & TextBox1.Text
            Case "8" : TextBox1.Text = "01" & "" & TextBox1.Text
            Case "8" : TextBox1.Text = "0" & "" & TextBox1.Text
        End Select
        loDataTable.Clear()
        abrir()
        llenar_combo_box3()
        CheckBox1.Checked = True
        TextBox2.Select()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "010000000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "01000000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0100000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "010000" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "01000" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = "0100" & "" & TextBox1.Text
                Case "7" : TextBox1.Text = "010" & "" & TextBox1.Text
                Case "8" : TextBox1.Text = "01" & "" & TextBox1.Text
                Case "8" : TextBox1.Text = "0" & "" & TextBox1.Text
            End Select
            loDataTable.Clear()
            abrir()
            llenar_combo_box3()
            TextBox2.Select()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Detalle_Asignacion.Label2.Text = "01"
        Detalle_Asignacion.Show()
    End Sub
End Class