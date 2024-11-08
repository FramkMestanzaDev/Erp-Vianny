Public Class Packing_hilo
    Private Sub Packing_hilo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If RadioButton1.Checked = True Then
            TextBox3.ReadOnly = False
            TextBox2.ReadOnly = True
        Else
            If RadioButton2.Checked = True Then
                TextBox2.ReadOnly = False
                TextBox3.ReadOnly = True
            End If


        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim bh As New vpackingtela
        Dim ch As New vpackingtela
        Dim bh1 As New fingresopac
        Dim as1 As Integer
        bh.gpartida = TextBox1.Text
        bh.garticulo = TextBox4.Text
        bh.gcodigo_trab = TextBox5.Text
        bh.gnombre_trab = TextBox6.Text
        bh.gnumero_rollos = TextBox3.Text
        bh.gunidad = "RLL"
        bh1.insertar_packing_tela(bh)

        as1 = DataGridView1.Rows.Count
        For i = 0 To as1 - 1
            ch.gnrollo = DataGridView1.Rows(i).Cells(0).Value
            ch.gpeso = DataGridView1.Rows(i).Cells(1).Value
            ch.gposicion = DataGridView1.Rows(i).Cells(2).Value
            ch.galmacen = DataGridView1.Rows(i).Cells(3).Value
            ch.garticulo2 = DataGridView1.Rows(i).Cells(4).Value
            bh1.insertar_packing_tela_detalle(ch)
        Next
        MsgBox("SE REGISTRO CON EXITO")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        DataGridView1.DataSource = ""
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            TextBox3.ReadOnly = False
            TextBox2.ReadOnly = True
        Else
            If RadioButton2.Checked = True Then
                TextBox2.ReadOnly = False
                TextBox3.ReadOnly = True
            End If


        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton1.Checked = True Then
            TextBox3.ReadOnly = False
            TextBox2.ReadOnly = True
        Else
            If RadioButton2.Checked = True Then
                TextBox2.ReadOnly = False
                TextBox3.ReadOnly = True
            End If


        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub
End Class