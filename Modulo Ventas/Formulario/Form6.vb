Imports System.Data.SqlClient
Public Class Form6
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
    Dim DT14 As New DataTable
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()

        Dim cmd6 As New SqlDataAdapter("SELECT distinct A.program_3 AS PO, ISNULL(B.nomb_10,'') AS CLIENTE  FROM custom_vianny.dbo.QAG0300 A LEFT JOIN custom_vianny.dbo.CAG1000 B  ON '01' = B.CCIA AND A.FICH_3 = B.FICH_10  Where A.CCia='" + Label4.Text + "' AND YEAR(fcom_3) >= '2021' AND broker_3 IN (0001,0002) AND LEN(program_3) >0  AND  merchan_3 <>'0026' ORDER BY program_3
", conx)

        cmd6.Fill(DT14)
        DataGridView1.DataSource = DT14
        DataGridView1.Columns(0).Width = 110
        DataGridView1.Columns(1).Width = 400
        TextBox1.Select()
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex > 0 Then
            Ocs.DataGridView1.Rows(Label3.Text).Cells(2).Value = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            DT14.Clear()
            Me.Close()
        End If

    End Sub
    Sub BUSCAR()
        Dim ds As New DataSet
        ds.Tables.Add(DT14.Copy)
        Dim dv As New DataView(ds.Tables(0))
        dv.RowFilter = "PO" & " like '%" & TextBox1.Text & "%'"
        If dv.Count <> 0 Then
            DataGridView1.DataSource = dv
            DataGridView1.Columns(1).Width = 110
            DataGridView1.Columns(1).Width = 400
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub
    Sub BUSCAR1()
        Dim ds As New DataSet
        ds.Tables.Add(DT14.Copy)
        Dim dv As New DataView(ds.Tables(0))
        dv.RowFilter = "CLIENTE" & " like '%" & TextBox2.Text & "%'"
        If dv.Count <> 0 Then
            DataGridView1.DataSource = dv
            DataGridView1.Columns(1).Width = 110
            DataGridView1.Columns(1).Width = 400
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        BUSCAR()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        BUSCAR1()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class