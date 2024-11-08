Imports System.Data.SqlClient
Public Class Detalle_ms
    Public conx As SqlConnection
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Dim Rsr2127 As SqlDataReader

    Private Sub Detalle_ms_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        abrir()
        Dim sql1021 As String = "SELECT distinct correlativo FROM requerimiento_avios_detalle where id_cabecera ='" + Label1.Text + "'  and correlativo <>0"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr2127 = cmd1021.ExecuteReader()
        Dim i51 As Integer
        i51 = 0
        While Rsr2127.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i51).Cells(0).Value = Rsr2127(0)
            i51 = i51 + 1
        End While



    End Sub



    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Rpt_armadoA.TextBox1.Text = Label1.Text
        Rpt_armadoA.TextBox2.Text = Label5.Text
        Rpt_armadoA.TextBox3.Text = Label6.Text
        Rpt_armadoA.TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Rpt_armadoA.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class