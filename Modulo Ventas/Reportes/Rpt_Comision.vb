Public Class Rpt_Comision
    Private Sub Rpt_Comision_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Select Case ComboBox1.Text
            Case "ENERO" : Comisiones_Mes.TextBox1.Text = "01"
            Case "FEBRERO" : Comisiones_Mes.TextBox1.Text = "02"
            Case "MARZO" : Comisiones_Mes.TextBox1.Text = "03"
            Case "ABRIL" : Comisiones_Mes.TextBox1.Text = "04"
            Case "MAYO" : Comisiones_Mes.TextBox1.Text = "05"
            Case "JUNIO" : Comisiones_Mes.TextBox1.Text = "06"
            Case "JULIO" : Comisiones_Mes.TextBox1.Text = "07"
            Case "AGOSTO" : Comisiones_Mes.TextBox1.Text = "08"
            Case "SETIEMBRE" : Comisiones_Mes.TextBox1.Text = "09"
            Case "OCTUBRE" : Comisiones_Mes.TextBox1.Text = "10"
            Case "NOVIEMBRE" : Comisiones_Mes.TextBox1.Text = "11"
            Case "DICIEMBRE" : Comisiones_Mes.TextBox1.Text = "12"
        End Select
        Comisiones_Mes.TextBox2.Text = Label3.Text
        Comisiones_Mes.TextBox3.Text = Label3.Text
        Comisiones_Mes.Show()
    End Sub
    Dim dt As New DataTable
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim fg As New vfactura
        Dim fg1 As New ffactura
        Select Case ComboBox1.Text
            Case "ENERO" : fg.gmes = "01"
            Case "FEBRERO" : fg.gmes = "02"
            Case "MARZO" : fg.gmes = "03"
            Case "ABRIL" : fg.gmes = "04"
            Case "MAYO" : fg.gmes = "05"
            Case "JUNIO" : fg.gmes = "06"
            Case "JULIO" : fg.gmes = "07"
            Case "AGOSTO" : fg.gmes = "08"
            Case "SETIEMBRE" : fg.gmes = "09"
            Case "OCTUBRE" : fg.gmes = "10"
            Case "NOVIEMBRE" : fg.gmes = "11"
            Case "DICIEMBRE" : fg.gmes = "12"
        End Select
        fg.gperiodo = Label3.Text
        fg.gccia = Label4.Text
        dt = fg1.comisiones_vendedor_detallado(fg)
        DataGridView1.DataSource = dt
        Dim jk As New Exportar
        jk.llenarExcel(DataGridView1)
    End Sub
End Class