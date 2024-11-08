Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Public Class OdxPm
    Dim da As New DataTable
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Label4.Text = "1" And (Microsoft.VisualBasic.Mid(TextBox2.Text.ToString.Trim, 1, 1) = "M") Then
            ReportePm.TextBox1.Text = Label3.Text
            ReportePm.TextBox2.Text = TextBox2.Text
            ReportePm.TextBox3.Text = Label4.Text
            ReportePm.ShowDialog()
        Else
            If Label4.Text = "2" And (Microsoft.VisualBasic.Mid(TextBox2.Text.ToString.Trim, 2, 1) <> "M") Then
                ListadoOpxPo.TextBox1.Text = Label3.Text
                ListadoOpxPo.TextBox2.Text = TextBox2.Text
                ListadoOpxPo.ShowDialog()
            Else
                If Label4.Text = "1" Then
                    MsgBox("ESTE REPORTE ES SOLO DE PROGRAMA DE MUESTRA")
                Else
                    MsgBox("ESTE REPORTE ES SOLO DE PROGRAMA DE PRODUCCION")
                End If

            End If
        End If


    End Sub
    Private Sub MOSTRAR()
        DataGridView1.DataSource = ""
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("exec Sp_ReporteOpsxPO '" + Label3.Text + "','" + Trim(TextBox2.Text) + "'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        If da.Rows.Count > 0 Then


            Dim EXP = New ExportarSoloData
            EXP.llenarExcel(DataGridView1)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MOSTRAR()

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub



    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            If Label4.Text = "1" Then
                TodOds.ShowDialog()
            Else
                ListaPO.Label3.Text = "1"
                ListaPO.ShowDialog()
            End If

        End If
    End Sub


End Class