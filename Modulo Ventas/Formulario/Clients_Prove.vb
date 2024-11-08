Imports System.Data.SqlClient
Public Class Clients_Prove
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
    Dim da As New DataTable
    Private Sub Clients_Prove_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("select  fich_10 as RUC,nomb_10 AS NOMBRE,direcc_10 as DIRECCION from custom_vianny.dbo.cag1000 where ccia ='01' and LEN(ruc_10) >0", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 71
        DataGridView1.Columns(1).Width = 300
        DataGridView1.Columns(2).Width = 350

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If Label1.Text = "1" Then
            Letra.TextBox8.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Letra.TextBox9.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
            Letra.TextBox11.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)

            TextBox1.Text = ""
            TextBox2.Text = ""
            Me.Close()
        Else
            Letra.TextBox17.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Letra.TextBox16.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
            Letra.TextBox14.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)

            TextBox1.Text = ""
            TextBox2.Text = ""
            Me.Close()
        End If

    End Sub
    Dim dt As New DataTable
    Private Sub buscarruc()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "RUC" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 71
                DataGridView1.Columns(1).Width = 300
                DataGridView1.Columns(2).Width = 350

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscarnombre()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "NOMBRE" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 71
                DataGridView1.Columns(1).Width = 300
                DataGridView1.Columns(2).Width = 350
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscarruc()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscarnombre()
    End Sub
End Class