Public Class OB_SC
    Private Sub OB_SC_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Label2.Text = 1 Then
            os.DataGridView1.Rows(Label1.Text).Cells(11).Value = TextBox1.Text
            os.DataGridView1.Rows(Label1.Text).Cells(15).Value = TextBox2.Text
            os.DataGridView1.Rows(Label1.Text).Cells(16).Value = TextBox3.Text
            os.DataGridView1.Rows(Label1.Text).Cells(17).Value = TextBox4.Text
            os.DataGridView1.Rows(Label1.Text).Cells(18).Value = TextBox5.Text
        Else
            If Label2.Text = 2 Then
                Guia_remision.DataGridView1.Rows(Label1.Text).Cells(17).Value = TextBox1.Text
                Guia_remision.DataGridView1.Rows(Label1.Text).Cells(18).Value = TextBox2.Text
                Guia_remision.DataGridView1.Rows(Label1.Text).Cells(19).Value = TextBox3.Text
                Guia_remision.DataGridView1.Rows(Label1.Text).Cells(20).Value = TextBox4.Text
                Guia_remision.DataGridView1.Rows(Label1.Text).Cells(21).Value = TextBox5.Text
            Else
                If Label2.Text = 3 Then
                    Guia_hilo.DataGridView1.Rows(Label1.Text).Cells(16).Value = TextBox1.Text
                    Guia_hilo.DataGridView1.Rows(Label1.Text).Cells(17).Value = TextBox2.Text
                    Guia_hilo.DataGridView1.Rows(Label1.Text).Cells(18).Value = TextBox3.Text
                    Guia_hilo.DataGridView1.Rows(Label1.Text).Cells(19).Value = TextBox4.Text
                    Guia_hilo.DataGridView1.Rows(Label1.Text).Cells(20).Value = TextBox5.Text
                Else
                    If Label2.Text = 4 Then
                        Ocs.DataGridView1.Rows(Label1.Text).Cells(18).Value = Trim(TextBox1.Text)
                        Ocs.DataGridView1.Rows(Label1.Text).Cells(19).Value = Trim(TextBox2.Text)
                        Ocs.DataGridView1.Rows(Label1.Text).Cells(20).Value = Trim(TextBox3.Text)
                        Ocs.DataGridView1.Rows(Label1.Text).Cells(21).Value = Trim(TextBox4.Text)
                        Ocs.DataGridView1.Rows(Label1.Text).Cells(22).Value = Trim(TextBox5.Text)
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub OB_SC_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class