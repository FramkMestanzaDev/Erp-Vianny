Public Class Observacion_rama

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim hg As New foberrama
        Dim ll As New vrama

        ll.gPROGRAMA = TextBox2.Text
        ll.gfecha = DateTimePicker1.Value
        ll.ghora = DateTimePicker2.Value
        ll.gMOTIVO = ComboBox1.Text
        ll.gCOMENTARIO = TextBox1.Text
        hg.INSERTAR_OBSERVACIONES_RAMA(ll)
        MsgBox("SE REGISTRO LA INFORMACION SOLICITADA")
        Me.Close()
    End Sub

    Private Sub Observacion_rama_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class