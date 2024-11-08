Imports System.Data.SqlClient
Public Class Requerimiento_op
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
    Function insertar(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function ELIMINAR(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Dim dt, dt2 As New DataTable
    Sub CALCULAR()
        DataGridView1.Rows.Clear()
        Dim k As New vprogramacion
        Dim k1 As New vprogramacion
        Dim l As New fprogramacion
        DataGridView3.DataSource = ""
        DataGridView2.DataSource = ""


        k1.gccia = "01"
        k1.gop = TextBox1.Text
        dt2 = l.buscar_FICHA(k1)
        If dt2.Rows.Count <> 0 Then
            DataGridView3.DataSource = dt2
            DataGridView1.Rows.Add()
            DataGridView1.Rows(0).Cells(0).Value = "01"
            DataGridView1.Rows(0).Cells(1).Value = DataGridView3.Rows(0).Cells(5).Value
            DataGridView1.Rows(0).Cells(4).Value = TextBox3.Text
            DataGridView1.Rows(0).Cells(5).Value = TextBox12.Text
            DataGridView1.Rows(0).Cells(6).Value = TextBox11.Text
            DataGridView1.Rows(0).Cells(7).Value = TextBox12.Text
            DataGridView1.Rows(0).Cells(8).Value = TextBox12.Text
            DataGridView1.Rows(0).Cells(9).Value = TextBox12.Text
            k.gccia = "01"
            k.gcodigo = DataGridView1.Rows(0).Cells(1).Value
            dt = l.buscar_tejido(k)
            If dt.Rows.Count <> 0 Then
                DataGridView2.DataSource = dt
            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CALCULAR()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 2 Then
            Ficha.TextBox2.Text = 3
            Ficha.Label2.Text = "01"
            Ficha.Show()
        End If
    End Sub
    Dim Rsr20, Rsr201 As SqlDataReader
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button2.Enabled = True
    End Sub
    Dim da As New DataTable
    Private Sub Requerimiento_op_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("EXEC custom_vianny.dbo.GetAllRequerimientoTextilDet '01','" + Trim(TextBox1.Text) + "'", conx)
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then

            DataGridView4.DataSource = da
            DataGridView1.Rows.Add()
            DataGridView1.Rows(0).Cells(0).Value = "01"
            DataGridView1.Rows(0).Cells(1).Value = DataGridView4.Rows(0).Cells(3).Value
            DataGridView1.Rows(0).Cells(3).Value = DataGridView4.Rows(0).Cells(16).Value
            DataGridView1.Rows(0).Cells(4).Value = DataGridView4.Rows(0).Cells(6).Value
            DataGridView1.Rows(0).Cells(5).Value = DataGridView4.Rows(0).Cells(7).Value
            DataGridView1.Rows(0).Cells(6).Value = DataGridView4.Rows(0).Cells(8).Value
            DataGridView1.Rows(0).Cells(7).Value = DataGridView4.Rows(0).Cells(9).Value
            DataGridView1.Rows(0).Cells(8).Value = DataGridView4.Rows(0).Cells(10).Value
            DataGridView1.Rows(0).Cells(9).Value = DataGridView4.Rows(0).Cells(11).Value
            DataGridView1.Rows(0).Cells(10).Value = DataGridView4.Rows(0).Cells(4).Value

            Dim sql1020 As String = "EXEC custom_vianny.dbo.GetallRequerimientoTextilCab '01','" + Trim(TextBox1.Text) + "'"
            Dim cmd1020 As New SqlCommand(sql1020, conx)
            Rsr20 = cmd1020.ExecuteReader()
            Rsr20.Read()
            TextBox7.Text = Rsr20(15)
            TextBox8.Text = Rsr20(16)
            TextBox9.Text = Rsr20(18)
            TextBox10.Text = Rsr20(17)
            TextBox11.Text = Rsr20(9)

        Else


            DataGridView4.DataSource = ""
            TextBox7.Text = "0.00"
            TextBox8.Text = "0.00"
            TextBox9.Text = "0.00"
            TextBox10.Text = "0.00"
            TextBox11.Text = "1.00"
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim vf As Integer
        vf = TextBox8.Text
        Dim agregar1 As String = "delete from custom_vianny.dbo.RequerimientoTextil_Det where ccia_det ='01' and pedido_det = '" + Trim(TextBox1.Text) + "'"
        ELIMINAR(agregar1)
        abrir()
        Dim cmd13 As New SqlCommand("INSERT INTO custom_vianny.dbo.RequerimientoTextil_Det (CCia_Det,Pedido_Det,Combi_Det,Tela_Det,Color_Det,Pieza_Det,Solicitado_Det,Programado_Det,Consumo_Det,TelaAcabada_Det,TelaCruda_Det,Hilado_Det,mermatint_det)
                                                                                   VALUES ( '01'   ,@a        ,'01'     ,@b      ,@j       ,''       ,@c            ,@d            ,@e         ,@f             ,@g           ,@h        , @i)", conx)
        cmd13.Parameters.AddWithValue("@a", TextBox1.Text)
        cmd13.Parameters.AddWithValue("@b", DataGridView1.Rows(0).Cells(1).Value)
        cmd13.Parameters.AddWithValue("@c", DataGridView1.Rows(0).Cells(4).Value)
        cmd13.Parameters.AddWithValue("@d", DataGridView1.Rows(0).Cells(5).Value)
        cmd13.Parameters.AddWithValue("@e", DataGridView1.Rows(0).Cells(6).Value)
        cmd13.Parameters.AddWithValue("@f", DataGridView1.Rows(0).Cells(7).Value)
        cmd13.Parameters.AddWithValue("@g", DataGridView1.Rows(0).Cells(8).Value)
        cmd13.Parameters.AddWithValue("@h", DataGridView1.Rows(0).Cells(9).Value)
        cmd13.Parameters.AddWithValue("@i", TextBox8.Text)
        cmd13.Parameters.AddWithValue("@j", DataGridView1.Rows(0).Cells(10).Value)
        cmd13.ExecuteNonQuery()
        abrir()

            Dim cmd As New SqlCommand("UPDATE custom_vianny.dbo.QAG0300 SET  pmartint_3 = @pmartint_3,pmarteñ_3 = @pmarteñ_3,pmartej_3 = @pmartej_3,Kgsprenda_3 =@Kgsprenda_3 WHERE ccia = '01' And ncom_3 = '" + TextBox1.Text + "'", conx)
            cmd.Parameters.AddWithValue("@pmartint_3", TextBox8.Text)
            cmd.Parameters.AddWithValue("@pmarteñ_3", TextBox7.Text)
            cmd.Parameters.AddWithValue("@pmartej_3", TextBox9.Text)
            cmd.Parameters.AddWithValue("@Kgsprenda_3", TextBox11.Text)

            cmd.ExecuteNonQuery()
            Dim vk As New vprogramacion
            Dim vk1 As New vprogramacion
            Dim vk2 As New vprogramacion
            Dim ko As New fprogramacion

            vk.gop = TextBox1.Text
            vk.gfecha = DateTimePicker1.Value
            ko.insertar_Qarg0300(vk)
            Dim k As Integer
            k = DataGridView2.Rows.Count
            Dim agregar2 As String = "delete custom_vianny.dbo.QARP0300 where ccia_3p ='01' and pedido_3p= '" + Trim(TextBox1.Text) + "'"
            ELIMINAR(agregar2)
            For i = 0 To k - 1
                vk1.gop = TextBox1.Text
                'vk1.gcodigo = DataGridView1.Rows(0).Cells(1).Value
                vk1.gfecha = DateTimePicker1.Value
                'vk1.gcolor = DataGridView1.Rows(0).Cells(10).Value
                vk1.gcodigo1 = DataGridView2.Rows(i).Cells(2).Value
                'vk1.gcodigo2 = DataGridView2.Rows(1).Cells(2).Value
                'vk1.gcodigo3 = DataGridView2.Rows(2).Cells(2).Value
                'vk1.gvalor2 = DataGridView2.Rows(1).Cells(12).Value * DataGridView2.Rows(0).Cells(9).Value / 100
                vk1.gvalor1 = DataGridView2.Rows(i).Cells(12).Value * DataGridView1.Rows(0).Cells(9).Value / 100
                'vk1.gvalor3 = DataGridView2.Rows(2).Cells(12).Value * DataGridView2.Rows(0).Cells(9).Value / 100
                'vk1.gvalor = DataGridView1.Rows(0).Cells(8).Value
                'vk1.gvalor4 = DataGridView1.Rows(0).Cells(7).Value
                ko.insertar_Qarp0300(vk1)
            Next
            vk2.gop = TextBox1.Text
            vk2.gcodigo = DataGridView1.Rows(0).Cells(1).Value
            vk2.gfecha = DateTimePicker1.Value
            vk2.gcolor = DataGridView1.Rows(0).Cells(10).Value
            vk2.gvalor = DataGridView1.Rows(0).Cells(8).Value
            vk2.gvalor4 = DataGridView1.Rows(0).Cells(7).Value
            ko.insertar_Qarp0300_det(vk2)

            Dim jk As String

            jk = Mid(DataGridView1.Rows(0).Cells(1).Value, 1, 7) & Trim(DataGridView1.Rows(0).Cells(10).Value) & Mid(Trim(DataGridView1.Rows(0).Cells(1).Value), 14, 25)

            Dim sql10201 As String = "select count(linea_17) from custom_vianny.dbo.cag1700 where  linea_17 ='" + jk + "'"
            Dim cmd10201 As New SqlCommand(sql10201, conx)
            Rsr201 = cmd10201.ExecuteReader()
            Rsr201.Read()
        If Rsr201(0) > 0 Then

        Else
            Dim fg As New fcodigo
            Dim jk1 As New vcodigo
            jk1.glinea = DataGridView1.Rows(0).Cells(1).Value
            jk1.glinea2 = jk
            jk1.gcolor = DataGridView1.Rows(0).Cells(3).Value
            If fg.insertar_codigo_planemiento(jk1) Then
                MsgBox("SE REGISTRO EL CODIGO CORRECTAMENTE")
            End If

        End If
            MessageBox.Show("DATOS INGRESADOS CORRECTAMENTE")
            Button2.Enabled = False



    End Sub
End Class