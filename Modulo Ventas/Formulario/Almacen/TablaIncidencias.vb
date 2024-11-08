Imports System.Data.SqlClient
Public Class TablaIncidencias
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public comando As SqlCommand
    Dim da As New DataTable
    Dim Rsr2, Rsr3 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub Listar_InformacionMotivo()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT CodMpb as CODIGO,DesMpb AS DESCRIPCION FROM MotivosProblemas", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(0).Width = 55
            DataGridView1.Columns(1).Width = 310
        Else
            DataGridView1.DataSource = Nothing
        End If

    End Sub
    Sub Listar_InformacionRubro()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select CodRpb as CODIGO,DesRpb AS DESCRIPCION  from RubroProblemas", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(0).Width = 55
            DataGridView1.Columns(1).Width = 310
        Else
            DataGridView1.DataSource = Nothing
        End If

    End Sub
    Sub buscar()
        Dim ds As DataSet
        ds.Tables.Add(da.Copy)
        Dim dv As New DataView(ds.Tables(0))
        dv.RowFilter = "DESCRIPCION" & "LIKE %" & TextBox1.Text & "%'"
        If dv.Count > 0 Then
            DataGridView1.DataSource = dv
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Label2.Text = "1" Then
            RegisroInconvenientes.TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
            RegisroInconvenientes.TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
        Else
            RegisroInconvenientes.TextBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim()
            RegisroInconvenientes.TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
        End If
        Me.Close()
    End Sub

    Private Sub TablaIncidencias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Label2.Text = "1" Then
            Listar_InformacionMotivo()
        Else
            Listar_InformacionRubro()
        End If

    End Sub
End Class