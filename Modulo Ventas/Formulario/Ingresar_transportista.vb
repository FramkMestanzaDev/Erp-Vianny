Public Class Ingresar_transportista
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("ALGUNO DE LOS CAMPOS ESTA VACIO")
        Else
            Dim vg As New fingreso_tranposrtista
            Dim vg1 As New vingre_tranportistas

            vg1.gcodigo = TextBox1.Text
            vg1.gchofer = TextBox2.Text
            vg1.gmarca = TextBox3.Text
            vg1.gplaca = TextBox4.Text
            vg.insertar_ingre_transportista(vg1)
            MsgBox("SE INGRESO CORRECTAMENTE LA INFORMACION SOLICITADA")
        End If

    End Sub
End Class