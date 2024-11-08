Imports System.Data.SqlClient

Public Class MatrizAviosMonitor
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub mostrarinformacion()
        abrir()
        da.Clear()
        Dim cmd As New SqlDataAdapter("EXEC Sp_CabeceraMatrizConsumoAvios '" + Label3.Text + "',NULL,'" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "','" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "'", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da

            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(0).Width = 85
            DataGridView1.Columns(1).Width = 300
            DataGridView1.Columns(3).Width = 130
            DataGridView1.Columns(4).Width = 80
            DataGridView1.Columns(5).Width = 74
            DataGridView1.Columns(5).ReadOnly = True
            'DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Visible = False
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub
    Private Sub MatrizAviosMonitor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = DateTimePicker1.Value.AddDays(-90)
        mostrarinformacion()
    End Sub
    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        If DataGridView1.Columns(e.ColumnIndex).Name = "ESTADO" Then
            If e.Value IsNot Nothing AndAlso e.Value.ToString().Trim() = "CERRADO" Then
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.DarkRed
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White
            End If
        End If
    End Sub
    Private Sub buscarop()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            dv.RowFilter = "OP" & " like '%" & TextBox1.Text & "%'"
            If dv.Count <> 0 Then
                DataGridView1.DataSource = dv

                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(0).Width = 85
                DataGridView1.Columns(1).Width = 300
                DataGridView1.Columns(3).Width = 130
                DataGridView1.Columns(4).Width = 80
                DataGridView1.Columns(5).Width = 74
                DataGridView1.Columns(5).ReadOnly = True
                'DataGridView1.Columns(5).Frozen = True
                'DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Visible = False

            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscarop()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        mostrarinformacion()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex > 0 Then

        End If
        If DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString().Trim() = "CERRADO" Then
            Button2.Text = "Abrir Matriz Avios"
        Else
            Button2.Text = "Cerrar Matriz Avios"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim valor As String = ""
        Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
        If filaSeleccionada.Cells("ESTADO").Value.ToString().Trim() = "CERRADO" Then
            valor = "0"
        Else
            valor = "1"
        End If
        abrir()
        MsgBox(valor)
        Dim cmd15776 As New SqlCommand("UPDATE custom_vianny.dbo.QAG0300 SET cierrematcon_3 = @estado Where ccia= @ccia And ncom_3 = @ncom_3", conx)
        cmd15776.Parameters.AddWithValue("@ccia", Label3.Text)
        cmd15776.Parameters.AddWithValue("@estado", valor)
        cmd15776.Parameters.AddWithValue("@ncom_3", filaSeleccionada.Cells("OP").Value.ToString().Trim())
        cmd15776.ExecuteNonQuery()
        Dim cmd15775 As New SqlCommand("UPDATE custom_vianny.dbo.Qamc800 SET Cierre_8 = @estado WHERE CCia_8 = @ccia AND NCom_8 =  @ncom_3", conx)
        cmd15775.Parameters.AddWithValue("@ccia", Label3.Text)
        cmd15775.Parameters.AddWithValue("@estado", valor)
        cmd15775.Parameters.AddWithValue("@ncom_3", filaSeleccionada.Cells("OP").Value.ToString().Trim())
        cmd15775.ExecuteNonQuery()

        MsgBox("Se Actualizo la Informacion solicitada")
        mostrarinformacion()
    End Sub
    Sub limpiar()
        TextBox1.Text = ""
    End Sub
End Class