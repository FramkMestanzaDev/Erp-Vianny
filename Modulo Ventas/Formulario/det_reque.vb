Imports System.Data.SqlClient
Public Class det_reque
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim dt10 As New DataTable
    Private Sub det_reque_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        If Label2.Text = 2 Then
            Dim cmd5 As New SqlDataAdapter("select det_req,nu_requerimientio,prducto AS PRODUCTO, cantidad as CANTIDAD,estado from  det_req_comercial  WHERE  estado =0  and nu_requerimientio ='" + Trim(TextBox1.Text) + "'", conx)

            cmd5.Fill(dt10)
            DataGridView1.DataSource = dt10
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(2).Width = 400
            DataGridView1.Columns(3).Width = 100
            Label1.Visible = True
            TextBox1.Visible = True
        Else
            Dim cmd5 As New SqlDataAdapter("Select  nu_requerimientio As REQUERIMIENTO, PO As po,cliente As cliente from cab_req_comercial where estado=0", conx)

            cmd5.Fill(dt10)
            DataGridView1.DataSource = dt10
            'DataGridView1.Columns(0).Visible = False
            'DataGridView1.Columns(1).Visible = False
            'DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(0).Width = 100
            DataGridView1.Columns(1).Width = 100
            DataGridView1.Columns(2).Width = 300
            Label1.Visible = False
            TextBox1.Visible = False
        End If


    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim nun1, num2 As Integer
        nun1 = e.RowIndex
        num2 = e.ColumnIndex
        If Label2.Text = 2 Then
            Orden_Produccion.TextBox23.Text = DataGridView1.Rows(nun1).Cells(0).Value
            Orden_Produccion.TextBox6.Text = DataGridView1.Rows(nun1).Cells(3).Value
            Orden_Produccion.TextBox24.Text = DataGridView1.Rows(nun1).Cells(2).Value
            Me.Close()
        Else
            Orden_Produccion.TextBox22.Text = DataGridView1.Rows(nun1).Cells(0).Value
            Orden_Produccion.TextBox23.Enabled = True
            Orden_Produccion.TextBox23.Select()
            Me.Close()
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class