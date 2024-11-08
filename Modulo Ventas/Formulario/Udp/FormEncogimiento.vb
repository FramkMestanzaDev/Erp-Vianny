Imports System.Data.SqlClient
Public Class FormEncogimiento
    Dim da As New DataTable
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Mostrar_Informacion()
    End Sub
    Sub Mostrar_Informacion()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("exec Sp_Encogimiento '" + TextBox1.Text.ToString().Trim() + "','" + Label5.Text.ToString().Trim() + "','0' ", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(1).Width = 70
        DataGridView1.Columns(2).Width = 77
        DataGridView1.Columns(3).Width = 77
        DataGridView1.Columns(4).Width = 30
        DataGridView1.Columns(5).Width = 120
        DataGridView1.Columns(6).Width = 120
        DataGridView1.Columns(7).Width = 120
        DataGridView1.Columns(8).Width = 120
        DataGridView1.Columns(9).Width = 30
        DataGridView1.Columns(10).Width = 70
        DataGridView1.Columns(11).Width = 70
        DataGridView1.Columns(12).Width = 70
        DataGridView1.Columns(13).Width = 70
        DataGridView1.Columns(14).Width = 70
        MsgBox("Se cargo la Informacion de Encogimientos Pendientes Correctamente")
    End Sub
    Sub Mostrar_InformacionCerrada()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("exec Sp_Encogimiento '" + TextBox1.Text.ToString().Trim() + "','" + Label5.Text.ToString().Trim() + "','1' ", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(1).Width = 70
        DataGridView1.Columns(2).Width = 77
        DataGridView1.Columns(3).Width = 77
        DataGridView1.Columns(4).Width = 30
        DataGridView1.Columns(5).Width = 120
        DataGridView1.Columns(6).Width = 120
        DataGridView1.Columns(7).Width = 120
        DataGridView1.Columns(8).Width = 120
        DataGridView1.Columns(9).Width = 30
        DataGridView1.Columns(10).Width = 70
        DataGridView1.Columns(11).Width = 70
        DataGridView1.Columns(12).Width = 70
        DataGridView1.Columns(13).Width = 70
        DataGridView1.Columns(14).Width = 70
        MsgBox("Se cargo la Informacio de Encogimientos Cerrados Correctamente")
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim p As Integer
        p = DataGridView1.Rows.Count
        If p > 0 Then
            For i = 0 To p - 1
                If DataGridView1.Rows(i).Cells(0).Value = True Then
                    Dim cmd20 As New SqlCommand("update custom_vianny.dbo.qag0300 set  acabados_3=1 where ncom_3 =@ncom_3 and ccia ='01'", conx)
                    cmd20.Parameters.AddWithValue("@ncom_3", Trim(DataGridView1.Rows(i).Cells(3).Value).ToString())
                    cmd20.ExecuteNonQuery()
                End If
            Next
            MsgBox("SE CERRO LA OD")
            Mostrar_Informacion()
        Else
            MsgBox("NO HAY INFORMACION PARA CERRAR")

        End If
    End Sub
    Private Sub buscarOp()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "OP" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 75
                DataGridView1.Columns(2).Width = 85
                DataGridView1.Columns(3).Width = 75
                DataGridView1.Columns(4).Width = 75
                DataGridView1.Columns(5).Width = 160
                DataGridView1.Columns(6).Width = 160
                DataGridView1.Columns(7).Width = 160
                DataGridView1.Columns(8).Width = 160
                DataGridView1.Columns(9).Width = 35
                DataGridView1.Columns(10).Width = 75
                DataGridView1.Columns(11).Width = 75
                DataGridView1.Columns(12).Width = 75
                DataGridView1.Columns(13).Width = 75
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscarPo()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PO" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(1).Width = 75
                DataGridView1.Columns(2).Width = 85
                DataGridView1.Columns(3).Width = 75
                DataGridView1.Columns(4).Width = 75
                DataGridView1.Columns(5).Width = 160
                DataGridView1.Columns(6).Width = 160
                DataGridView1.Columns(7).Width = 160
                DataGridView1.Columns(8).Width = 160
                DataGridView1.Columns(9).Width = 35
                DataGridView1.Columns(10).Width = 75
                DataGridView1.Columns(11).Width = 75
                DataGridView1.Columns(12).Width = 75
                DataGridView1.Columns(13).Width = 75
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscarOp()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscarPo()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ext As New ExportaruDP
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub FormEncogimiento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Mostrar_Informacion()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Mostrar_InformacionCerrada()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Label5.Text.ToString().Trim() = "01" Then
            RptEncogimiento.TextBox1.Text = "VIANNY"
        Else
            RptEncogimiento.TextBox1.Text = "GRAUS"
        End If

        RptEncogimiento.TextBox2.Text = Label5.Text.ToString().Trim()
        RptEncogimiento.TextBox3.Text = TextBox1.Text.ToString().Trim()
        RptEncogimiento.Show()
    End Sub
End Class