Imports System.Data.SqlClient
Public Class detalle_reclamo
    Dim DT, DT2 As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim ty, ty2, ty3 As New DataTable



    Public conn As SqlDataAdapter
    Dim bs As New BindingSource()
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        bs.Filter = String.Format("PEDIDO LIKE '%{0}%'", TextBox1.Text)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub detalle_reclamo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim Rsr19912 As SqlDataReader
        Dim i51 As Integer
        Dim sql10112 As String = "select numero_pedido AS PEDIDO,C.nomb_10 AS CLIENTE, total AS TOTAL,fecha,ISNULL(sfactu_3 +'-'+ nfactu_3,'') as FACTURA,ISNULL(CONVERT( VARCHAR(10),F.fcom_3, 103),'') as Fecha from nota_pedido N INNER JOIN custom_vianny.DBO.cag1000 C ON N.codigo_cli = C.fich_10 AND N.empresa = C.ccia left JOIN  custom_vianny.DBO.faG0300 f  ON n.numero_pedido = f.tdanoper3_3 and n.empresa = F.ccia_3 AND N.codigo_cli = F.fich_3 where  year(fecha) = '" + Label2.Text + "' AND ccia ='" + Label3.Text + "'  AND vendedor ='" + Label4.Text + "' and n.codigo_cli='" + Label5.Text + "' ORDER BY numero_pedido"
        Dim cmd10112 As New SqlCommand(sql10112, conx)
        Rsr19912 = cmd10112.ExecuteReader()
        i51 = 0

        While Rsr19912.Read()
            DataGridView1.Rows.Add(1)
            DataGridView1.Rows(i51).Cells(0).Value = Rsr19912(0)
            DataGridView1.Rows(i51).Cells(1).Value = Rsr19912(1)
            DataGridView1.Rows(i51).Cells(2).Value = Rsr19912(2)
            DataGridView1.Rows(i51).Cells(3).Value = Rsr19912(4)
            DataGridView1.Rows(i51).Cells(4).Value = Rsr19912(5)
            'DataGridView1.Rows(i51).Cells(0).Value = bs
            i51 = i51 + 1
        End While
        Rsr19912.Close()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Hoja_Reclamos.TextBox12.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Hoja_Reclamos.TextBox8.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value) <> "" Then
            Hoja_Reclamos.DateTimePicker2.Value = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        End If

        Me.Close()
    End Sub
End Class