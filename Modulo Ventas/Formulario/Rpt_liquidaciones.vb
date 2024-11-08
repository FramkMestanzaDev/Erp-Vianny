Imports System.Data.SqlClient
Public Class Rpt_liquidaciones
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Dim Rsr2 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Rpt_liquidaciones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("EXEC custom_vianny.dbo.Spu_StatusManufacturaEnvioRetornoxProceso '" + Label5.Text + "','" + Trim(TextBox3.Text) + "','" + Trim(TextBox4.Text) + "',NULL,NULL,'" + Trim(TextBox2.Text) + "','" + Trim(TextBox2.Text) + "','" + Label6.Text + "','" + Label7.Text + "'", conx)
        Dim da As New DataTable
        cmd.Fill(da)
        DataGridView2.DataSource = da

        Dim hj, hj1, hj2 As Integer
        hj = DataGridView2.Rows.Count
        hj1 = 0
        hj2 = 0
        Dim ruc As String

        For i = 0 To hj - 1
            If DataGridView2.Rows(i).Cells(2).Value = "1" Then
                ruc = DataGridView2.Rows(i).Cells(10).Value
                hj1 = hj1 + 1
            End If
            If DataGridView2.Rows(i).Cells(2).Value = "2" Then

            End If
            hj2 = hj2 + 1

        Next

        If hj1 > hj2 Then
            DataGridView1.Rows.Add(hj1 + 1)
        Else
            DataGridView1.Rows.Add(hj2 + 1)
        End If


    End Sub
End Class