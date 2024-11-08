Public Class Direccion_Despacho
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, conta As Integer
        i = DataGridView1.Rows.Count
        a = 0
        conta = 0
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                conta = conta + 1
            End If
        Next
        If conta > 1 Then
            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
        Else
            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                    Nota_Pedido.TextBox7.Text = DataGridView1.Rows(a).Cells(1).Value
                End If
            Next
        End If

        Close()
    End Sub

    Private Sub Direccion_Despacho_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class