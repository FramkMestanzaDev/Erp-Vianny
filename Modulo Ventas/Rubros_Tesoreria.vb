Imports System.Data.SqlClient
Public Class Rubros_Tesoreria
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
        If Trim(TextBox1.Text) = "" Then
            MsgBox("FALTA REGISTRAR EL NOMBRE DEL RUBRO")
        Else
            Dim agregar2 As String = "delete from RUBROS_TESORERIA where CODIGO = '" + TextBox2.Text + "'"
            ELIMINAR(agregar2)
            Dim cmd15 As New SqlCommand("insert into RUBROS_TESORERIA (CODIGO,NOMBRE,NATURALEZA,ACTIVIDAD) values (@CODIGO,@NOMBRE,@NATURALEZA,@ACTIVIDAD)", conx)
            cmd15.Parameters.AddWithValue("@CODIGO", Trim(TextBox2.Text))
            cmd15.Parameters.AddWithValue("@NOMBRE", Trim(TextBox1.Text))
            cmd15.Parameters.AddWithValue("@NATURALEZA", Trim(ComboBox1.Text))
            cmd15.Parameters.AddWithValue("@ACTIVIDAD", Trim(ComboBox2.Text))
            cmd15.ExecuteNonQuery()
            MsgBox("SE AGREGO LA INFORMACION SOLICITADA")
            CORRELATIVO()
            MOSTRAR()
            TextBox1.Text = ""
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()

        Dim agregar1 As String = "delete from RUBROS_TESORERIA where CODIGO = '" + TextBox2.Text + "'"

        ELIMINAR(agregar1)
        MsgBox("SE ELIMINO EL REGISTRO SOLICITADO")
        TextBox2.Text = ""
        MOSTRAR()
        CORRELATIVO()
    End Sub
    Sub CORRELATIVO()
        abrir()
        Dim Rsr2 As SqlDataReader
        Dim sql102 As String = "SELECT TOP 1  CODIGO   AS VAL FROM RUBROS_TESORERIA   order by CODIGO desc"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            TextBox2.Text = Rsr2(0) + 1
        Else
            TextBox2.Text = 1
        End If

        Select Case Trim(TextBox2.Text).Length
            Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
            Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
            Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
            Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
            Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
            Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
            Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
            Case "8" : TextBox2.Text = TextBox2.Text

        End Select
        Rsr2.Close()
    End Sub

    Sub MOSTRAR()
        DataGridView1.Rows.Clear()
        Dim i5 As Integer
        Dim Rsr19 As SqlDataReader
        Dim sql101 As String = "select * from RUBROS_TESORERIA"
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
    Private Sub Rubros_Tesoreria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        CORRELATIVO()
        MOSTRAR()
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex <> -1 Then
            TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            ComboBox1.Text = Trim((DataGridView1.Rows(e.RowIndex).Cells(3).Value)).ToString
            ComboBox2.SelectedItem = Trim((DataGridView1.Rows(e.RowIndex).Cells(2).Value)).ToString
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        CORRELATIVO()
        TextBox1.Text = ""
        TextBox1.Select()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
End Class