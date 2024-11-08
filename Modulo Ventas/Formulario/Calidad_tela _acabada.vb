Imports System.Data.SqlClient
Public Class Calidad_tela__acabada
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
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dg2.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 85
                DataGridView1.Columns(2).Width = 85

                DataGridView1.Columns(3).Width = 300
                DataGridView1.Columns(4).Width = 586
                DataGridView1.Columns(3).DefaultCellStyle.BackColor = Color.Beige
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Dim dg2 As New DataTable
    Private Sub Calidad_tela__acabada_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 1
        abrir()
        dg2.Clear()
        Dim cmd2 As New SqlDataAdapter("SELECT partida as PARTIDA, fecha AS FECHAS,descripcion AS PRODUCTO,observacion AS OBSERVACION FROM CALIDAD_PARTIDA where YEAR(fecha) ='" + Trim(Label3.Text) + "' and len(partida) >0 order by fecha", conx)

        cmd2.Fill(dg2)
        If dg2.Rows.Count <> 0 Then
            DataGridView1.DataSource = dg2
            'DataGridView3.Columns(2).Width = 350
            DataGridView1.Columns(1).Width = 85
            DataGridView1.Columns(2).Width = 85

            DataGridView1.Columns(3).Width = 300
            DataGridView1.Columns(4).Width = 586
            DataGridView1.Columns(3).DefaultCellStyle.BackColor = Color.Beige
        Else
            DataGridView1.DataSource = ""
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString

            Form4.TextBox1.Text = DataGridView1.Rows(num1).Cells(1).Value
            Form4.Label22.Text = 2
            Form4.TextBox1.Select()
            Form4.TextBox1.Focus()
            SendKeys.Send("{ENTER}")
            Form4.Show()

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = ""
        dg2.Clear()
        abrir()

        Dim cmd2 As New SqlDataAdapter("SELECT partida as PARTIDA, fecha AS FECHAS,descripcion AS PRODUCTO,observacion AS OBSERVACION FROM CALIDAD_PARTIDA where YEAR(fecha) ='" + Trim(ComboBox1.Text) + "'", conx)

        cmd2.Fill(dg2)
        If dg2.Rows.Count <> 0 Then
            DataGridView1.DataSource = dg2
            'DataGridView3.Columns(2).Width = 350
            DataGridView1.Columns(1).Width = 85
            DataGridView1.Columns(2).Width = 85

            DataGridView1.Columns(3).Width = 300
            DataGridView1.Columns(4).Width = 586
            DataGridView1.Columns(3).DefaultCellStyle.BackColor = Color.Beige
        Else
            DataGridView1.DataSource = ""
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class