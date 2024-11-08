Imports System.Data.SqlClient
Public Class TallasTa
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim Rsr20 As SqlDataReader
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
    Private Sub TallasTa_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MOSTRAR_INFORMACION()
    End Sub
    Sub BUSCAR()
        Dim ds As New DataSet
        ds.Tables.Add(DT14.Copy)
        Dim dv As New DataView(ds.Tables(0))
        dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox2.Text & "%'"
        If dv.Count <> 0 Then
            DataGridView1.DataSource = dv
            DataGridView1.Columns(0).Width = 110
            DataGridView1.Columns(1).Width = 340
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub
    Public Sub MOSTRAR_INFORMACION()
        abrir()
        DataGridView1.DataSource = ""
        DT14.Clear()
        Dim cmd6 As New SqlDataAdapter(" SELECT cele as CODIGO,dele AS DESCRIPCION FROM custom_vianny.DBO.TAB0100   Where ccia='01' AND CTAB='TALLAS' order by cast(cele as integer)", conx)
        cmd6.Fill(DT14)
        DataGridView1.DataSource = DT14
        DataGridView1.Columns(0).Width = 110
        DataGridView1.Columns(1).Width = 340
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Select Case Label1.Text
            Case 1

                PadronTalla.Label13.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Case 2
                PadronTalla.Label14.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Case 3
                PadronTalla.Label15.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Case 4
                PadronTalla.Label16.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Case 5
                PadronTalla.Label17.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox6.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Case 6
                PadronTalla.Label18.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox7.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Case 7
                PadronTalla.Label19.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox8.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Case 8
                PadronTalla.Label20.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox9.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Case 9
                PadronTalla.Label21.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox10.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Case 10
                PadronTalla.Label22.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                PadronTalla.TextBox11.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)

        End Select
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        BUSCAR()
    End Sub

End Class