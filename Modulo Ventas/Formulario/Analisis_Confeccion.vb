Public Class Analisis_Confeccion
    Dim DT, DT1 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim GH As New forden_costura
        DataGridView1.DataSource = ""
        Dim GH1 As New vorden_costura
        If TextBox1.Text = "" Then
            MsgBox("EL CAMPO PO ESTA VACIO", vbExclamation)
        Else
            GH1.gop = TextBox1.Text
            DT = GH.ANALISIS_CONFECCION(GH1)
            If DT.Rows.Count <> 0 Then
                DataGridView1.DataSource = DT
                DataGridView1.Columns(0).Width = 330
                DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Else
                MsgBox("NO EXISTE INFORMACION DE LA OP SOLICITADA")
            End If
        End If
    End Sub



    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            BUSCAR_OP.TextBox2.Text = 1
            BUSCAR_OP.Show()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim GH As New forden_costura
        Dim GH1 As New vorden_costura
        DataGridView2.DataSource = ""
        If TextBox2.Text = "" Then
            MsgBox("EL CAMPO PO ESTA VACIO", vbExclamation)
        Else
            GH1.gop = TextBox2.Text
            DT1 = GH.ANALISIS_CONFECCIONPO(GH1)

            If DT1.Rows.Count <> 0 Then
                DataGridView2.DataSource = DT1
                DataGridView2.Columns(1).Width = 260
                DataGridView2.Columns(2).Width = 80
                DataGridView2.Columns(3).Width = 100
                DataGridView2.Columns(4).Width = 100
                DataGridView2.Columns(5).Width = 100
                DataGridView2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Else
                MsgBox("NO EXISTE INFORMACION DE LA OP SOLICITADA")
            End If
        End If

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Analisis_Confeccion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            BUSCAR_OP.TextBox2.Text = 2
            BUSCAR_OP.Show()
        End If
    End Sub
End Class