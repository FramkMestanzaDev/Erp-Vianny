Imports System.Data.SqlClient
Public Class merma
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
    Dim Rs, rs1, rs2, Rs12, Rs121 As SqlDataReader

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim ext As New Exportar
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add(8)
            DataGridView1.Rows(0).Cells(0).Value = "PROGRAMADO"
            DataGridView1.Rows(1).Cells(0).Value = "HILO DESPACHADO"
            DataGridView1.Rows(2).Cells(0).Value = "TEJIDO"
            DataGridView1.Rows(3).Cells(0).Value = "VALIDADO TEJEDURIA"
            DataGridView1.Rows(4).Cells(0).Value = "RAMEADO"
            DataGridView1.Rows(6).Cells(0).Value = "DEVOLUCION HILO"
            DataGridView1.Rows(7).Cells(0).Value = "MERMA TOTAL"
            DataGridView1.Rows(0).Cells(4).Value = TextBox1.Text
            DataGridView1.Rows(1).Cells(4).Value = TextBox1.Text
            DataGridView1.Rows(2).Cells(4).Value = TextBox1.Text
            DataGridView1.Rows(3).Cells(4).Value = TextBox1.Text
            DataGridView1.Rows(4).Cells(4).Value = TextBox1.Text
            abrir()
            Dim sql12 As String = "select SUM(cantk_3r),SUM(cantkmv_3r), ISNULL( kgsobtenidos,0) ,isnull((select SUM(cantk_3a) from custom_vianny.dbo.map030f where lote_3a LIKE '" + TextBox1.Text + "%' and talm_3a in (10,51,54) and ccom_3a ='1' and CCIA_3A ='" + Label4.Text + "' and CPERIODO_3A ='" + Label5.Text + "'),0),ISNULL((select SUM(cantk_3a) from custom_vianny.dbo.map030f where parti_3a = '" + TextBox1.Text + "' and talm_3a in (03,25,40,42) AND ccom_3a ='2' AND CCIA_3A ='" + Label4.Text + "'),0) from custom_vianny.dbo.marg0001 m
left join validar_partida  v on m.partida_3r = v.partida
where partida_3r ='" + TextBox1.Text + "' and flag_3r >1 group by  kgsobtenidos"
            Dim cmd12 As New SqlCommand(sql12, conx)
            Rs12 = cmd12.ExecuteReader()

            If Rs12.Read() Then
                DataGridView1.Rows(0).Cells(1).Value = Rs12(0)
                DataGridView1.Rows(1).Cells(1).Value = Rs12(4) 'aca me pone el primer campo del select
                DataGridView1.Rows(2).Cells(1).Value = Rs12(1)
                DataGridView1.Rows(3).Cells(1).Value = Rs12(2)
                DataGridView1.Rows(4).Cells(1).Value = Rs12(3)
            Else
                MsgBox("PARTIDA NO EXISTE")
            End If
            DataGridView1.Rows(1).Cells(2).Value = Convert.ToDecimal(DataGridView1.Rows(0).Cells(1).Value) - Convert.ToDecimal(DataGridView1.Rows(2).Cells(1).Value)
            DataGridView1.Rows(2).Cells(2).Value = Convert.ToDecimal(DataGridView1.Rows(1).Cells(1).Value) - Convert.ToDecimal(DataGridView1.Rows(2).Cells(1).Value)
            DataGridView1.Rows(3).Cells(2).Value = Convert.ToDecimal(DataGridView1.Rows(2).Cells(1).Value) - Convert.ToDecimal(DataGridView1.Rows(3).Cells(1).Value)
            DataGridView1.Rows(4).Cells(2).Value = Convert.ToDecimal(DataGridView1.Rows(3).Cells(1).Value) - Convert.ToDecimal(DataGridView1.Rows(4).Cells(1).Value)

            If DataGridView1.Rows(1).Cells(1).Value = 0 Then
                DataGridView1.Rows(2).Cells(3).Value = 0
                DataGridView1.Rows(3).Cells(3).Value = 0
                DataGridView1.Rows(4).Cells(3).Value = 0
            Else
                DataGridView1.Rows(2).Cells(3).Value = Math.Round(((Convert.ToDecimal(DataGridView1.Rows(1).Cells(1).Value) - Convert.ToDecimal(DataGridView1.Rows(1).Cells(2).Value)) / Convert.ToDecimal(DataGridView1.Rows(1).Cells(1).Value)) * 100, 2)
                DataGridView1.Rows(3).Cells(3).Value = Math.Round(((Convert.ToDecimal(DataGridView1.Rows(1).Cells(1).Value) - Convert.ToDecimal(DataGridView1.Rows(1).Cells(3).Value)) / Convert.ToDecimal(DataGridView1.Rows(1).Cells(1).Value)) * 100, 2)
                DataGridView1.Rows(4).Cells(3).Value = Math.Round(((Convert.ToDecimal(DataGridView1.Rows(1).Cells(1).Value) - Convert.ToDecimal(DataGridView1.Rows(1).Cells(4).Value)) / Convert.ToDecimal(DataGridView1.Rows(1).Cells(1).Value)) * 100, 2)

            End If
            Rs12.Close()

            Dim sql121 As String = "select isnull(SUM(p.cantk_3a),0) from custom_vianny.dbo.mag030f g inner join custom_vianny.dbo.map030f p on g.CCIA_3 =p.CCIA_3A and g.CPERIODO_3 = p.CPERIODO_3A and g.talm_3 = p.talm_3a and g.ncom_3 =p.ncom_3a where  p.lote_3a ='" + TextBox1.Text + "%' and CCIA_3 ='" + Label4.Text + "' and orig_3 ='0002' and ccom_3 ='1' and CPERIODO_3 ='" + Label5.Text + "' "
            Dim cmd121 As New SqlCommand(sql121, conx)
            Rs121 = cmd121.ExecuteReader()

            If Rs121.Read() Then
                DataGridView1.Rows(6).Cells(1).Value = Rs121(0)
            End If
            Rs121.Close()

        End If
    End Sub
End Class