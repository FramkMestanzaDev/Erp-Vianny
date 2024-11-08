Imports System.Data.SqlClient
Public Class Lista_reclamo
    Dim DT, DT2 As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim ty, ty2, ty3 As New DataTable
    Public conn As SqlDataAdapter

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Dim Rsr23 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'buscar
    End Sub

    Private Sub Lista_reclamo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim sql1023 As String = "select numero,cliente,nota_pedido,fecha from Hoja_Reclamo_tela where vendedor ='" + Trim(Label3.Text) + "'"
        Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr23 = cmd1023.ExecuteReader()
        Dim i51 As Integer
        i51 = 0
        While Rsr23.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i51).Cells(0).Value = Trim(Rsr23(0))
            DataGridView1.Rows(i51).Cells(1).Value = Trim(Rsr23(1))
            DataGridView1.Rows(i51).Cells(2).Value = Trim(Rsr23(2))
            DataGridView1.Rows(i51).Cells(3).Value = Trim(Rsr23(3))
            i51 = i51 + 1
        End While
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Hoja_Reclamos.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
        Hoja_Reclamos.TextBox1.Focus()
        SendKeys.Send("{ENTER}")
        Me.Close()
    End Sub
    'Private Sub buscar()
    '    Try
    '        Dim ds As New DataSet
    '        ds.Tables.Add(da.Copy)
    '        Dim dv As New DataView(ds.Tables(0))


    '        dv.RowFilter = "ORDEN" & " like '%" & TextBox1.Text & "%'"

    '        If dv.Count <> 0 Then

    '            DataGridView1.DataSource = dv
    '            Dim i6 As Integer
    '            i6 = DataGridView1.Rows.Count
    '            For I1 = 0 To i6 - 1
    '                If DataGridView1.Rows(I1).Cells(6).Value = "1" Then
    '                    DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.IndianRed
    '                End If

    '            Next

    '        Else

    '            DataGridView1.DataSource = Nothing
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message)

    '    End Try
    'End Sub
End Class