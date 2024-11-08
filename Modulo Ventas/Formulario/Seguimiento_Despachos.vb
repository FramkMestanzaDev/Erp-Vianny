Imports System.Data.SqlClient
Public Class Seguimiento_Despachos
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
    Private Sub Seguimiento_Despachos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim gf As New fnopedido
        Dim gf1 As New vnapedido
        gf1.gperiodo = Label4.Text
        gf1.gempresa = Label3.Text
        dt = gf.validar_despacho(gf1)
        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            Dim k As Integer
            k = DataGridView1.Rows.Count


            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(2).Width = 220
            DataGridView1.Columns(7).Visible = False

            DataGridView1.Columns(8).DefaultCellStyle.BackColor = Color.Beige
            DataGridView1.Columns(9).DefaultCellStyle.BackColor = Color.Azure
            DataGridView1.Columns(10).DefaultCellStyle.BackColor = Color.Coral
            For i = 0 To k - 1

                If Trim(DataGridView1.Rows(i).Cells(8).Value) = "PENDIENTE" Then
                    DataGridView1.Rows(i).Cells(8).Style.BackColor = Color.DarkCyan
                    DataGridView1.Rows(i).Cells(8).Style.ForeColor = Color.White
                End If
                If Trim(DataGridView1.Rows(i).Cells(9).Value) = "PENDIENTE" Then
                    DataGridView1.Rows(i).Cells(9).Style.BackColor = Color.DarkKhaki
                    DataGridView1.Rows(i).Cells(9).Style.ForeColor = Color.White
                End If
                If Trim(DataGridView1.Rows(i).Cells(10).Value) = "PENDIENTE" Then
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.DarkGray
                    DataGridView1.Rows(i).Cells(10).Style.ForeColor = Color.White
                End If

            Next
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "RAZON_SOCIAL" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim k As Integer
                k = DataGridView1.Rows.Count




                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(2).Width = 220
                DataGridView1.Columns(7).Visible = False
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PEDIDO" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim k As Integer
                k = DataGridView1.Rows.Count






                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(2).Width = 220
                DataGridView1.Columns(7).Visible = False

            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar2()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar()
    End Sub
End Class