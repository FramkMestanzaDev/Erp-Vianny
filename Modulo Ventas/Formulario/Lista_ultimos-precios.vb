Imports System.Data.SqlClient
Public Class Lista_ultimos_precios
    Public conx As SqlConnection
    Dim da As New DataTable
    Dim bs As New BindingSource()
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim Rsr20 As SqlDataReader
    Private Sub Lista_ultimos_precios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        da.Clear()
        abrir()
        Dim sql10212 As String = "EXEC custom_vianny.DBO.CaeSoft_ListadoUltimoPrecios '" + Label1.Text + "', NULL, '" + TextBox1.Text + "      ', '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
        Dim cmd10212 As New SqlCommand(sql10212, conx)
        Rsr20 = cmd10212.ExecuteReader()
        Dim i51 As Integer
        While Rsr20.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i51).Cells(0).Value = Rsr20(0)
            DataGridView1.Rows(i51).Cells(1).Value = Rsr20(2)
            DataGridView1.Rows(i51).Cells(2).Value = Rsr20(3)
            DataGridView1.Rows(i51).Cells(3).Value = Rsr20(5)
            DataGridView1.Rows(i51).Cells(4).Value = Rsr20(10)
            DataGridView1.Rows(i51).Cells(5).Value = Rsr20(9)
            DataGridView1.Rows(i51).Cells(6).Value = Rsr20(12)
            DataGridView1.Rows(i51).Cells(7).Value = Rsr20(13)

            i51 = i51 + 1
        End While


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Trim(Label3.Text) = "SOLES" Then
            Ocs.DataGridView1.Rows(Label2.Text).Cells(11).Value = DataGridView1.Rows(e.RowIndex).Cells(6).Value
        Else
            Ocs.DataGridView1.Rows(Label2.Text).Cells(11).Value = DataGridView1.Rows(e.RowIndex).Cells(7).Value
        End If

        Me.Close()
    End Sub
End Class