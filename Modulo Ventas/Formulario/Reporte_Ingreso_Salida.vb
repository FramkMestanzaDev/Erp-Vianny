Public Class Reporte_Ingreso_Salida
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged


        Button1.Enabled = (TextBox1.TextLength > 0)
    End Sub

    Private Sub Reporte_Ingreso_Salida_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Dim HH As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim jh As New fnia
        Dim kl As New vnia
        kl.gpartida = TextBox1.Text
        kl.glinea = ""
        kl.gano = TextBox4.Text
        kl.gmes = ""
        kl.gccom = TextBox5.Text
        kl.gccia = Label4.Text
        kl.gop = TextBox6.Text
        HH = jh.PARTE_INGRESO(kl)
        If HH.Rows.Count <> 0 Then
            DataGridView1.DataSource = HH
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).Width = 150
            DataGridView1.Columns(4).Width = 70
            DataGridView1.Columns(5).Width = 50
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 50
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(7).Width = 80
        Else
            DataGridView1.DataSource = ""
            MsgBox("NO EXITE INFORMACION LA DE PARTIDA O LOTE SOLICITAD0")
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Button2.Enabled = (TextBox2.TextLength > 0)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Button3.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim jh As New fnia
        Dim kl As New vnia
        kl.gpartida = ""
        kl.glinea = ""
        kl.gano = TextBox4.Text
        kl.gccom = TextBox5.Text
        kl.gop = TextBox6.Text
        Select Case ComboBox1.Text
            Case "ENERO" : kl.gmes = "01"
            Case "FEBRERO" : kl.gmes = "02"
            Case "MARZO" : kl.gmes = "03"
            Case "ABRIL" : kl.gmes = "04"
            Case "MAYO" : kl.gmes = "05"
            Case "JUNIO" : kl.gmes = "06"
            Case "JULIO" : kl.gmes = "07"
            Case "AGOSTO" : kl.gmes = "08"
            Case "SETIEMBRE" : kl.gmes = "09"
            Case "OCTUBRE" : kl.gmes = "10"
            Case "NOVIEMBRE" : kl.gmes = "11"
            Case "DICIEMBRE" : kl.gmes = "12"
        End Select
        kl.gccia = Label4.Text
        HH = jh.PARTE_INGRESO(kl)
        If HH.Rows.Count <> 0 Then
            DataGridView1.DataSource = HH
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).Width = 150
            DataGridView1.Columns(4).Width = 70
            DataGridView1.Columns(5).Width = 50
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 50
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(7).Width = 80
        Else
            DataGridView1.DataSource = ""
            MsgBox("NO EXITE INFORMACION LA DE PARTIDA O LOTE SOLICITAD0")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim jh As New fnia
        Dim kl As New vnia
        kl.gpartida = ""
        kl.glinea = TextBox2.Text
        kl.gano = TextBox4.Text
        kl.gmes = ""
        kl.gccom = TextBox5.Text
        kl.gccia = Label4.Text
        kl.gop = TextBox6.Text
        HH = jh.PARTE_INGRESO(kl)

        If HH.Rows.Count <> 0 Then
            DataGridView1.DataSource = HH
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).Width = 150
            DataGridView1.Columns(4).Width = 70
            DataGridView1.Columns(5).Width = 50
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 50
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(7).Width = 80
        Else
            DataGridView1.DataSource = ""
            MsgBox("NO EXITE INFORMACION LA DE PARTIDA O LOTE SOLICITAD0")
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            Dim fechaActual As Date = Today
            pproductos.CCIA.Text = Label4.Text
            pproductos.ALMACEN.Text = "08"
            pproductos.ANO.Text = TextBox4.Text
            pproductos.FECHA.Text = Replace(fechaActual.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "13"
            pproductos.Label3.Text = 155
            pproductos.Label5.Text = 0
            pproductos.TextBox2.Text = ""
            pproductos.TextBox3.Text = ""
            pproductos.Show(Me)
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Reporte_Nia_Nsa.TextBox1.Text = Trim(DataGridView1.Rows(num1).Cells(1).Value)
            Reporte_Nia_Nsa.TextBox2.Text = TextBox5.Text
            Reporte_Nia_Nsa.TextBox3.Text = Trim(DataGridView1.Rows(num1).Cells(2).Value)
            Reporte_Nia_Nsa.TextBox4.Text = TextBox4.Text
            Reporte_Nia_Nsa.TextBox5.Text = Label4.Text

            Reporte_Nia_Nsa.Show()
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox6.Text.ToString().Trim().Length() = 0 Then
            MsgBox("FALTA INGRESAR UNA PO")
        Else
            Dim jh As New fnia
            Dim kl As New vnia
            kl.gpartida = TextBox1.Text
            kl.glinea = ""
            kl.gano = TextBox4.Text
            kl.gmes = ""
            kl.gccom = TextBox5.Text
            kl.gccia = Label4.Text
            kl.gop = TextBox6.Text
            HH = jh.PARTE_INGRESO(kl)
            If HH.Rows.Count <> 0 Then
                DataGridView1.DataSource = HH
                DataGridView1.Columns(1).Width = 80
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 150
                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 50
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 50
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(7).Width = 80
            Else
                DataGridView1.DataSource = ""
                MsgBox("NO EXITE INFORMACION LA DE PARTIDA O LOTE SOLICITAD0")
            End If
        End If

    End Sub
End Class