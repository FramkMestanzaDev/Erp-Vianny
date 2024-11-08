Imports System.Data.SqlClient
Public Class tabla_rubros
    Public conx As SqlConnection
    Public conn As SqlDataAdapter

    Public comando As SqlCommand
    Dim da As New DataTable
    Dim bs As New BindingSource()
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
        If Trim(TextBox2.Text) = "" Or Trim(TextBox3.Text) = "" Or Trim(TextBox4.Text) = "" Then
            MsgBox("LOS CAMPOS * ESTAN VACIOS")
        Else
            Dim agregar As String = "delete from  TABLA_RUBROS where codigo ='" + Trim(TextBox1.Text) + "'"
            ELIMINAR(agregar)
            Dim cmd15 As New SqlCommand("insert into TABLA_RUBROS (codigo,nombre,codigo_contabilidad,descripcion_contab,estado) values(@codigo,@nombre,@codigo_contabilidad,@descripcion_contab,@estado)", conx)
            cmd15.Parameters.AddWithValue("@codigo", Trim(TextBox1.Text))
            cmd15.Parameters.AddWithValue("@nombre", Trim(TextBox2.Text))
            cmd15.Parameters.AddWithValue("@codigo_contabilidad", Trim(TextBox3.Text))
            cmd15.Parameters.AddWithValue("@descripcion_contab", Trim(TextBox4.Text))
            cmd15.Parameters.AddWithValue("@estado", "1")
            cmd15.ExecuteNonQuery()
        End If
        abrir()
        correlativo()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        DataGridView1.Rows.Clear()
        MOSTRAR()
        TextBox2.Select()
    End Sub
    Sub MOSTRAR()

        Dim Rsr19912 As SqlDataReader
        Dim i51 As Integer
        Dim sql10112 As String = "select * from TABLA_RUBROS"
        Dim cmd10112 As New SqlCommand(sql10112, conx)
        Rsr19912 = cmd10112.ExecuteReader()
        i51 = 0

        While Rsr19912.Read()
            DataGridView1.Rows.Add(1)
            DataGridView1.Rows(i51).Cells(0).Value = Trim(Rsr19912(0))
            DataGridView1.Rows(i51).Cells(1).Value = Trim(Rsr19912(1))
            DataGridView1.Rows(i51).Cells(2).Value = Trim(Rsr19912(2))
            DataGridView1.Rows(i51).Cells(3).Value = Trim(Rsr19912(3))
            If Trim(Rsr19912(4)) = 1 Then
                DataGridView1.Rows(i51).Cells(4).Value = "ACTIVO"
            Else
                DataGridView1.Rows(i51).Cells(4).Value = "ANULADO"
            End If


            i51 = i51 + 1
        End While
        Rsr19912.Close()
    End Sub
    Sub correlativo()
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "select top 1 codigo from TABLA_RUBROS  order by codigo desc"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        Dim l As Integer
        l = DataGridView1.Rows.Count

        If Rsr1991.Read() Then

            TextBox1.Text = Rsr1991(0) + 1
        Else
            TextBox1.Text = 1
        End If
        Select Case Trim(TextBox1.Text).Length

            Case "1" : TextBox1.Text = "000" & TextBox1.Text
            Case "2" : TextBox1.Text = "00" & TextBox1.Text
            Case "3" : TextBox1.Text = "0" & TextBox1.Text
            Case "4" : TextBox1.Text = TextBox1.Text

        End Select
        Rsr1991.Close()
    End Sub

    Private Sub tabla_rubros_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        correlativo()
        MOSTRAR()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        Rubro.Label1.Text = MDIParent1.Label9.Text
        Rubro.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim cmd15 As New SqlCommand("update TABLA_RUBROS set estado ='0' where codigo =@codigo", conx)
        cmd15.Parameters.AddWithValue("@codigo", Trim(TextBox1.Text))
        cmd15.ExecuteNonQuery()
        MOSTRAR()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        TextBox2.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox2.Enabled = True
    End Sub
End Class