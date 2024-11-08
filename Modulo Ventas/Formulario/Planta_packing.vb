Public Class Planta_packing
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton2.Checked = True Then
            Ingresar_Packing.Label7.Text = "LIMA"
            Ingresar_Packing.Show()
        End If

    End Sub

End Class