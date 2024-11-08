Imports System.Data.SqlClient
Public Class rubros_manteminiento
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
    Private Sub rubros_manteminiento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
        ComboBox1.SelectedIndex = 0
        MOSTRAR()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        If Trim(TextBox1.Text) = "" Then
            MsgBox("FALTA REGISTRAR EL NOMBRE DEL RUBRO")
        Else

            Dim cmd15 As New SqlCommand("insert into RUBROS_MANTENIMIENTO (NOMBRE,TIPO,DIAS) values (@NOMBRE,@TIPO,@DIAS)", conx)
            cmd15.Parameters.AddWithValue("@NOMBRE", Trim(TextBox1.Text))
            cmd15.Parameters.AddWithValue("@TIPO", Trim(ComboBox1.Text))
            cmd15.Parameters.AddWithValue("@DIAS", Trim(TextBox2.Text))
            cmd15.ExecuteNonQuery()
            MsgBox("SE AGREGO LA INFORMACION SOLICITADA")
            MOSTRAR()
            TextBox1.Text = ""
            TextBox2.Text = ""

        End If
    End Sub
    Sub MOSTRAR()
        abrir()
        DataGridView1.Rows.Clear()
        Dim i5 As Integer
        Dim Rsr19 As SqlDataReader
        Dim sql101 As String = "select * from RUBROS_MANTENIMIENTO"
        Dim cmd101 As New SqlCommand(sql101, conx)
        Rsr19 = cmd101.ExecuteReader()

        While Rsr19.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i5).Cells(0).Value = Rsr19(0)
            DataGridView1.Rows(i5).Cells(1).Value = Rsr19(1)
            DataGridView1.Rows(i5).Cells(2).Value = Rsr19(2)
            DataGridView1.Rows(i5).Cells(3).Value = Rsr19(3)
            i5 = i5 + 1
            'DataGridView2.Rows(i5).Cells(6).Value = (Rsr2(12) * DataGridView1.Rows(0).Cells(10).Value) / 100
        End While

        Rsr19.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()

        Dim agregar1 As String = "delete from RUBROS_MANTENIMIENTO where ID = '" + Label3.Text + "'"

        ELIMINAR(agregar1)
        MsgBox("SE ELIMINO EL REGISTRO SOLICITADO")
        TextBox1.Text = ""
        MOSTRAR()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox1.Select()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
        ComboBox1.Text = Trim((DataGridView1.Rows(e.RowIndex).Cells(2).Value)).ToString
        Label3.Text = Trim((DataGridView1.Rows(e.RowIndex).Cells(0).Value))
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class