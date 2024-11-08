Imports System.Data.SqlClient
Public Class Mantenimiento_carro
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

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.F1 Then
            det_tesoreria.Label4.Text = 1
            det_tesoreria.Show()
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If Label12.Text <> 1 Then

            Dim fe As Integer
            fe = TextBox9.Text
            TextBox7.Text = DateAdd(DateInterval.Day, fe, DateTimePicker1.Value).ToString("yyyy/MM/dd")


        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            det_tesoreria.Label4.Text = 3
            det_tesoreria.Show()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        If Trim(TextBox2.Text) = "" Or Trim(TextBox7.Text) = "" Then
            MsgBox("FALTA REGISTRAR LA PLACA O LA FECHA DE PROGRAMACION")
        Else

            Dim agregar10 As String = "DELETE FROM MANTENIMIENTO WHERE id_mantenimiento = '" + Label10.Text + "'"
            ELIMINAR(agregar10)
            Dim cmd15 As New SqlCommand("INSERT INTO MANTENIMIENTO (UNIDAD,PLACA,RUBRO,FECHA,COSTO,FACTURA,PROGRAMACION,OBSERVACION,PERIODO) VALUES (@UNIDAD,@PLACA,@RUBRO,@FECHA,@COSTO,@FACTURA,@PROGRAMACION,@OBSERVACION,@PERIODO)", conx)
            cmd15.Parameters.AddWithValue("@UNIDAD", Trim(TextBox1.Text))
            cmd15.Parameters.AddWithValue("@PLACA", Trim(TextBox2.Text))
            cmd15.Parameters.AddWithValue("@RUBRO", Trim(TextBox8.Text))
            cmd15.Parameters.AddWithValue("@FECHA", DateTimePicker1.Value)
            cmd15.Parameters.AddWithValue("@COSTO", TextBox4.Text)
            cmd15.Parameters.AddWithValue("@FACTURA", Trim(TextBox5.Text))
            cmd15.Parameters.AddWithValue("@PROGRAMACION", Trim(TextBox7.Text))
            cmd15.Parameters.AddWithValue("@OBSERVACION", Trim(TextBox6.Text))
            cmd15.Parameters.AddWithValue("@PERIODO", Year(DateTimePicker1.Value))
            cmd15.ExecuteNonQuery()
            MsgBox("SE AGREGO LA INFORMACION SOLICITADA")
            MOSTRAR()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            Button2.Enabled = False
            Button4.Enabled = False
            Label12.Text = 0
            Label10.Text = 0
        End If
    End Sub
      Dim Rsr19 As SqlDataReader
    Sub MOSTRAR()
        abrir()
        DataGridView1.Rows.Clear()
        Dim i5 As Integer

        Dim sql101 As String = "SELECT M.*,R.ID,R.NOMBRE,R.DIAS FROM MANTENIMIENTO M LEFT JOIN RUBROS_MANTENIMIENTO R ON M.RUBRO =R.ID"
        Dim cmd101 As New SqlCommand(sql101, conx)
        Rsr19 = cmd101.ExecuteReader()

        While Rsr19.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i5).Cells(0).Value = Rsr19(0)
            DataGridView1.Rows(i5).Cells(1).Value = Rsr19(1)
            DataGridView1.Rows(i5).Cells(2).Value = Rsr19(2)
            DataGridView1.Rows(i5).Cells(3).Value = Rsr19(11)
            DataGridView1.Rows(i5).Cells(4).Value = Rsr19(4)
            DataGridView1.Rows(i5).Cells(5).Value = Rsr19(5)
            DataGridView1.Rows(i5).Cells(6).Value = Rsr19(6)
            DataGridView1.Rows(i5).Cells(7).Value = Rsr19(7)
            DataGridView1.Rows(i5).Cells(8).Value = Rsr19(8)
            DataGridView1.Rows(i5).Cells(9).Value = Rsr19(9)
            DataGridView1.Rows(i5).Cells(10).Value = Rsr19(10)
            DataGridView1.Rows(i5).Cells(11).Value = Rsr19(11)
            DataGridView1.Rows(i5).Cells(12).Value = Rsr19(12)
            i5 = i5 + 1
            'DataGridView2.Rows(i5).Cells(6).Value = (Rsr2(12) * DataGridView1.Rows(0).Cells(10).Value) / 100
        End While

        Rsr19.Close()
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()

        Dim agregar1 As String = "DELETE FROM MANTENIMIENTO WHERE id_mantenimiento = '" + Label10.Text + "'"

        ELIMINAR(agregar1)
        MsgBox("SE ELIMINO EL REGISTRO SOLICITADO")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        MOSTRAR()
        Label12.Text = 0
        Label10.Text = 0
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Label12.Text = 1
        TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
        TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(11).Value)
        TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
        TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
        TextBox6.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(8).Value)
        TextBox7.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value)
        DateTimePicker1.Value = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        TextBox10.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(9).Value)
        Label10.Text = Trim((DataGridView1.Rows(e.RowIndex).Cells(0).Value))
        TextBox8.Text = Trim((DataGridView1.Rows(e.RowIndex).Cells(10).Value))
        TextBox9.Text = Trim((DataGridView1.Rows(e.RowIndex).Cells(12).Value))
        Button2.Enabled = True
        Button4.Enabled = True
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
    End Sub

    Private Sub Mantenimiento_carro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MOSTRAR()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        buscar()
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            dt = TryCast(DataGridView1.DataSource, DataTable)
            ds.Tables.Add(dt)

            ds.Tables.Add(DataGridView1.DataSource)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PLACA" & " like '%" & TextBox11.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
End Class