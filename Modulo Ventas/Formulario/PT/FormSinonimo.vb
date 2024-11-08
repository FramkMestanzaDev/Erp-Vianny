Imports System.Data.SqlClient
Public Class FormSinonimo
    Public conx As SqlConnection
    Dim Rsr301, Rsr302 As SqlDataReader
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
    Dim Rsr3018 As SqlDataReader
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        Dim num As Integer
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()

            Try
                abrir()
                Select Case Trim(TextBox1.Text).Length
                    Case "1" : TextBox1.Text = "01" & "0000000" & TextBox1.Text
                    Case "2" : TextBox1.Text = "01" & "000000" & TextBox1.Text
                    Case "3" : TextBox1.Text = "01" & "00000" & TextBox1.Text
                    Case "4" : TextBox1.Text = "01" & "0000" & TextBox1.Text
                    Case "5" : TextBox1.Text = "01" & "000" & TextBox1.Text
                    Case "6" : TextBox1.Text = "01" & "00" & TextBox1.Text
                    Case "7" : TextBox1.Text = "01" & "0" & TextBox1.Text
                    Case "8" : TextBox1.Text = "01" & TextBox1.Text
                    Case "9" : TextBox1.Text = "0" & TextBox1.Text
                    Case "10" : TextBox1.Text = TextBox1.Text

                End Select

                Dim sql1022137 As String = "select count(c.nomb_sin) from custom_vianny.dbo.qag0300 q
                left join custom_vianny.dbo.Cag1700_Sinonimos c on q.ccia = c.ccia_sin and q.linea_3 = c.codigo_sin 
                where ncom_3='" + TextBox1.Text.ToString().Trim() + "' and q.ccia='" + Label3.Text.ToString().Trim() + "' and c.item_sin ='000'"
                Dim cmd1022137 As New SqlCommand(sql1022137, conx)
                Rsr3018 = cmd1022137.ExecuteReader()

                If Rsr3018.Read() Then
                    num = Convert.ToInt32(Rsr3018(0))

                End If
                Rsr3018.Close()

                Dim sql102213 As String = "select linea_3,isnull(c.nomb_17,q.descri_3) from custom_vianny.dbo.qag0300 q
                left join custom_vianny.dbo.cag1700 c on q.ccia = c.ccia and q.linea_3 = c.linea_17
                where ncom_3='" + TextBox1.Text.ToString().Trim() + "' and q.ccia='" + Label3.Text.ToString().Trim() + "'"
                Dim cmd102213 As New SqlCommand(sql102213, conx)
                Rsr301 = cmd102213.ExecuteReader()
                Dim jo As Integer
                If Rsr301.Read() Then
                    TextBox2.Text = Rsr301(0).ToString().Trim()
                    TextBox3.Text = Rsr301(1).ToString().Trim()
                End If
                Rsr301.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            If num = 0 Then
                registrarItems0()
            End If

            cargarInformacion()
        End If
    End Sub
    Sub registrarItems0()
        Dim cmd169 As New SqlCommand("insert into custom_vianny.dbo.Cag1700_Sinonimos(ccia_sin,codigo_sin,item_sin,nomb_sin) values (@ccia_sin,@codigo_sin,@item_sin,@nomb_sin)", conx)
        cmd169.Parameters.AddWithValue("@ccia_sin", Label3.Text.ToString().Trim())
        cmd169.Parameters.AddWithValue("@codigo_sin", TextBox2.Text.ToString().Trim())
        cmd169.Parameters.AddWithValue("@item_sin", "000")
        cmd169.Parameters.AddWithValue("@nomb_sin", TextBox3.Text.ToString().Trim())
        Try
            cmd169.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fila As Integer = DataGridView1.Rows.Count
        Dim filaf As String
        Select Case fila.ToString().Length
            Case 1 : filaf = "00" & fila
            Case 2 : filaf = "0" & fila
            Case 3 : filaf = fila
        End Select
        DataGridView1.Rows.Add()
        DataGridView1.Rows(fila).Cells(0).Value = filaf
        DataGridView1.Rows(fila).Cells(1).Value = TextBox2.Text.ToString().Trim()
        DataGridView1.Rows(fila).Cells(2).Value = TextBox4.Text.ToString().Trim()
        Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.Cag1700_Sinonimos(ccia_sin,codigo_sin,item_sin,nomb_sin) values (@ccia_sin,@codigo_sin,@item_sin,@nomb_sin)", conx)
        cmd168.Parameters.AddWithValue("@ccia_sin", Label3.Text.ToString().Trim())
        cmd168.Parameters.AddWithValue("@codigo_sin", TextBox2.Text.ToString().Trim())
        cmd168.Parameters.AddWithValue("@item_sin", filaf.ToString().Trim())
        cmd168.Parameters.AddWithValue("@nomb_sin", TextBox4.Text.ToString().Trim())
        Try
            cmd168.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try




    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        DataGridView1.Rows.Clear()
    End Sub
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
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim respuesta As DialogResult

        Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim agregar As String = "delete custom_vianny.dbo.Cag1700_Sinonimos  where codigo_sin ='" + filaSeleccionada.Cells("CODIGO").Value + "' and item_sin ='" + filaSeleccionada.Cells("SINO").Value + "'"
            ELIMINAR(agregar)
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        End If
    End Sub

    Sub cargarInformacion()
        Try
            abrir()
            Dim sql102213 As String = "select item_sin,codigo_sin, nomb_sin from custom_vianny.dbo.Cag1700_Sinonimos where codigo_sin ='" + TextBox2.Text.ToString().Trim() + "' and ccia_sin ='" + Label3.Text.ToString().Trim() + "'"
            Dim cmd102213 As New SqlCommand(sql102213, conx)
            Rsr302 = cmd102213.ExecuteReader()
            Dim i As Integer = 0
            While Rsr302.Read()
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i).Cells(0).Value = Rsr302(0).ToString().Trim()
                DataGridView1.Rows(i).Cells(1).Value = Rsr302(1).ToString().Trim()
                DataGridView1.Rows(i).Cells(2).Value = Rsr302(2).ToString().Trim()
                i = i + 1
            End While
            Rsr302.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class