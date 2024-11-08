Imports System.Data.SqlClient
Public Class Codigo_barra
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim Da1 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DataGridView1.Rows.Add()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)


        End If
    End Sub

    Private Sub Codigo_barra_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Da1.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select barra as CODIGO from codigo_barra where codigo ='" + TextBox1.Text + "'", conx)
        cmd.Fill(Da1)
        DataGridView2.DataSource = Da1

        Dim p As Integer

        p = DataGridView2.Rows.Count
        If p > 0 Then
            DataGridView1.Rows.Add(p)
            For i = 0 To p - 1
                DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(0).Value
                DataGridView1.Columns(0).Width = 270
            Next
        Else

        End If




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
        Dim K As Integer
        K = DataGridView1.Rows.Count
        abrir()

        Dim agregar As String = "DELETE FROM codigo_barra where codigo ='" + Trim(TextBox1.Text) + "'  "
        ELIMINAR(agregar)
        For I = 0 To K - 1
            Dim cmd15 As New SqlCommand("insert into codigo_barra (codigo,barra) values (@codigo,@barra)", conx)
            cmd15.Parameters.AddWithValue("@codigo", Trim(TextBox1.Text))
            cmd15.Parameters.AddWithValue("@barra", Trim(DataGridView1.Rows(I).Cells(0).Value))
            cmd15.ExecuteNonQuery()
        Next
        MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
        Me.Close()
    End Sub
End Class