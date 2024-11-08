Public Class DESPACHO
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim v As New vdetdireccion
        Dim v1 As New fdireccion
        v.gcodigo = TextBox3.Text
        v.gdirecion = TextBox2.Text

        v1.insertar_direccion(v)
        MsgBox("SE REGISTRO LA DIRECCION DE DESPACHO CORRECTAMENTE")

        TextBox2.Text = ""

    End Sub

    Private Sub DESPACHO_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class