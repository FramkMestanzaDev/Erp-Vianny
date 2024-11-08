Imports System.Data.SqlClient
Public Class vehuculo
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        If Trim(TextBox1.Text) = "" Or Trim(TextBox2.Text) = "" Then
            MsgBox("FALTA REGISTRAR LA PLACA O LA MARCA DEL VEHICULO")
        Else
            Dim agregar2 As String = "DELETE FROM VEHICULO WHERE PLACA = '" + Trim(TextBox1.Text) + "'"
            ELIMINAR(agregar2)
            Dim cmd15 As New SqlCommand("INSERT INTO VEHICULO (PLACA,MARCA,TIPO,AÑO,OBSERVACION) VALUES (@PLACA,@MARCA,@TIPO,@ANO,@OBSERVACION)", conx)
            cmd15.Parameters.AddWithValue("@PLACA", Trim(TextBox1.Text))
            cmd15.Parameters.AddWithValue("@MARCA", Trim(TextBox2.Text))
            cmd15.Parameters.AddWithValue("@TIPO", Trim(TextBox5.Text))
            cmd15.Parameters.AddWithValue("@ANO", Trim(TextBox4.Text))
            cmd15.Parameters.AddWithValue("@OBSERVACION", Trim(TextBox3.Text))
            cmd15.ExecuteNonQuery()
            MsgBox("SE AGREGO LA INFORMACION SOLICITADA")
            MOSTRAR()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            Button2.Enabled = False
            Button3.Enabled = False
        End If
    End Sub
    Sub MOSTRAR()
        abrir()
        DataGridView1.Rows.Clear()
        Dim i5 As Integer
        Dim Rsr19 As SqlDataReader
        Dim sql101 As String = "select * from  VEHICULO"
        Dim cmd101 As New SqlCommand(sql101, conx)
        Rsr19 = cmd101.ExecuteReader()

        While Rsr19.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i5).Cells(0).Value = Rsr19(0)
            DataGridView1.Rows(i5).Cells(1).Value = Rsr19(1)
            DataGridView1.Rows(i5).Cells(2).Value = Rsr19(2)
            DataGridView1.Rows(i5).Cells(3).Value = Rsr19(3)
            DataGridView1.Rows(i5).Cells(4).Value = Rsr19(4)
            i5 = i5 + 1
            'DataGridView2.Rows(i5).Cells(6).Value = (Rsr2(12) * DataGridView1.Rows(0).Cells(10).Value) / 100
        End While

        Rsr19.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        Dim agregar1 As String = "DELETE FROM VEHICULO WHERE PLACA = '" + Trim(TextBox1.Text) + "'"
        ELIMINAR(agregar1)
        MsgBox("SE ELIMINO EL REGISTRO SOLICITADO")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        MOSTRAR()
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
        TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
        TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
        TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
    End Sub

    Private Sub vehuculo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MOSTRAR()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
    End Sub
End Class