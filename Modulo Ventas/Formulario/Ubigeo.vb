Public Class Ubigeo
    Dim dy, DY2, DY3 As New DataTable
    Private Sub Ubigeo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fl As New fcliente

        dy = fl.buscar_departamento()
        DataGridView1.DataSource = dy
        DataGridView1.Columns(0).Width = 50
        DataGridView1.Columns(1).Width = 120
        DataGridView1.Columns(0).ReadOnly = True
        DataGridView1.Columns(1).ReadOnly = True
    End Sub

    Private Sub Ubigeo_FormClosed(sender As Object, e As FormClosedEventArgs)

    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim HJ As New fcliente
        Dim HG As New vcliente
        TextBox1.Text = Trim(DataGridView1.SelectedCells.Item(0).Value) & "-" & Trim(DataGridView1.SelectedCells.Item(1).Value)
        HG.gdepartamento = Trim(DataGridView1.SelectedCells.Item(0).Value)
        DY2 = HJ.buscar_PROVINCIA(HG)
        DataGridView2.DataSource = DY2
        DataGridView2.Columns(0).Width = 70
        DataGridView2.Columns(1).Width = 140
        DataGridView2.Columns(0).ReadOnly = True
        DataGridView2.Columns(1).ReadOnly = True
    End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form_Clientes.TextBox9.Text = Trim(DataGridView1.SelectedCells.Item(1).Value)
        Form_Clientes.TextBox10.Text = Trim(DataGridView2.SelectedCells.Item(1).Value)
        Form_Clientes.TextBox11.Text = Trim(DataGridView3.SelectedCells.Item(1).Value)
        Form_Clientes.TextBox19.Text = TextBox3.Text
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim HJ As New fcliente
        Dim HG As New vcliente
        TextBox2.Text = Trim(DataGridView2.SelectedCells.Item(0).Value) & " -- " & Trim(DataGridView1.SelectedCells.Item(1).Value) & " / " & Trim(DataGridView2.SelectedCells.Item(1).Value)
        HG.gprovincia = Trim(DataGridView2.SelectedCells.Item(0).Value)
        DY3 = HJ.BUSCAR_DISTRITO(HG)
        DataGridView3.DataSource = DY3
        DataGridView3.Columns(0).Width = 80
        DataGridView3.Columns(1).Width = 160
        DataGridView3.Columns(0).ReadOnly = True
        DataGridView3.Columns(1).ReadOnly = True
    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        TextBox3.Text = Trim(DataGridView3.SelectedCells.Item(0).Value)
    End Sub
End Class