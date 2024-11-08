Imports System.Data.SqlClient
Public Class Lista_Memorandum
    Dim dt As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Lista_Memorandum_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("select CodMem as MEMORANDUM, DeGMen,CarGMem,AReMem AS TRABAJADOR,CarRMem AS CARGO,AsuMem AS ASUNTO, FecMem AS FECHA from Memorandum_cabecera", conx)
        Dim da As New DataTable
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Width = 220
            DataGridView1.Columns(4).Width = 160
            DataGridView1.Columns(5).Width = 250
        Else

            DataGridView1.DataSource = ""
        End If

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "TRABAJADOR" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Visible = False
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(3).Width = 220
                DataGridView1.Columns(4).Width = 160
                DataGridView1.Columns(5).Width = 250
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then


            With DataGridView1
                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)

                    Memorándum.TextBox2.Text = Mid(DataGridView1.Rows(hti.RowIndex).Cells(0).Value, 5, 6)

                    Memorándum.TextBox2.Focus()
                    SendKeys.Send("{ENTER}")
                    Memorándum.TextBox3.Select()
                End If
            End With
            Me.Close()
        End If
    End Sub
End Class