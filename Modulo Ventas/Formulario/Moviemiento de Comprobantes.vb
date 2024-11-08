Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Net.Mail

Public Class Moviemiento_de_Comprobantes
    Public conx As SqlConnection
    Dim dt, dt2 As New DataTable

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
    Private Sub ACTUALIZAR()
        Dim MiDataSet As New DataTable
        conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
        Dim Comando As New SqlDataAdapter("select id_vn as ID, comprobante AS COMPROBANTE, total, parcial AS PG_PARCIAL, observacion AS OBSERVACION, fecha AS FECHA from ventas_negras_reg where comprobante ='" + Trim(Label1.Text) + "' and empresa='" + Trim(Label5.Text) + "'", conx)

        Comando.Fill(MiDataSet)

        Me.DataGridView1.DataSource = MiDataSet
        DataGridView1.Columns(4).Width = 400
        DataGridView1.Columns(0).Width = 50
        DataGridView1.Columns(1).Width = 140
        DataGridView1.Columns(2).Visible = False
        DataGridView1.Enabled = False
        Button1.Enabled = False
    End Sub

    Private Sub Moviemiento_de_Comprobantes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ACTUALIZAR()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.Enabled = True
        Button1.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a As Integer
        Dim sumas As Double
        Dim ok As New fcontailidad
        Dim ok1 As New vcontablidad
        i = DataGridView1.Rows.Count
        a = 0
        If i = 0 Then
            MsgBox("NO HAY NINGUNA FILA PARA REGISTRAR")
        Else
            For a = 0 To i - 1
                ok1.gCOMPROBANTE = Trim(DataGridView1.Rows(a).Cells(1).Value)
                ok1.gPARCIAL = Trim(DataGridView1.Rows(a).Cells(3).Value)
                If Convert.ToString(Trim(DataGridView1.Rows(a).Cells(4).Value)) = "" Then
                    ok1.gOBSERVACION = ""
                Else
                    ok1.gOBSERVACION = Trim(DataGridView1.Rows(a).Cells(4).Value)
                End If

                ok1.gID = Trim(DataGridView1.Rows(a).Cells(0).Value)
                ok1.gcia = Label5.Text
                ok.actualizar_PAGO_VN(ok1)
                sumas = sumas + Trim(DataGridView1.Rows(a).Cells(3).Value)
            Next
            abrir()

            Dim cmd As New SqlCommand("update venta_cabecera set aelanto = @monto where  sfactu_3 =@sfactu AND cperiodo_3 =@ano AND nfactu_3 = @nfactu", conx)
            cmd.Parameters.AddWithValue("@monto", sumas)
            cmd.Parameters.AddWithValue("@sfactu", Mid(Label1.Text, 1, 4))
            cmd.Parameters.AddWithValue("@ano", Label6.Text)
            cmd.Parameters.AddWithValue("@nfactu", Mid(Label1.Text, 6, 8))

            cmd.ExecuteNonQuery()
        End If
        MsgBox("SE ACTUALIZO LA INFORMACION SOLICITADA")
        ACTUALIZAR()
    End Sub
End Class