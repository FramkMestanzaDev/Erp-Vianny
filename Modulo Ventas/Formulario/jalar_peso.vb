Imports System.IO
Imports System.Text

Public Class jalar_peso
    Private Sub jalar_peso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Using reader As New StreamReader("C:\Program Files (x86)\HMG\WS PESO\peso.wei", Encoding.Default)
            TextBox1.Text = reader.ReadToEnd()
        End Using
    End Sub

End Class