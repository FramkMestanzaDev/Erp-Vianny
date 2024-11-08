Imports System.Data.SqlClient
Public Class Det_Os
    Public comando As SqlCommand
    Public conx As SqlConnection
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub buscar0()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da1.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "ORDEN" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Dim i6 As Integer
                i6 = DataGridView1.Rows.Count
                For I1 = 0 To i6 - 1
                    If DataGridView1.Rows(I1).Cells(6).Value = "1" Then
                        DataGridView1.Rows(I1).DefaultCellStyle.BackColor = Color.IndianRed
                    End If

                Next

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim da1 As New DataTable
    Private Sub Det_Os_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = DateTimePicker2.Value.AddDays(-30)
        abrir()
        If Label6.Text = "1" Then
            Dim cmd55 As New SqlDataAdapter("SELECT A.ncom_4 as ORDEN , A.fich_4 AS RUC, ISNULL(C.Nomb_10,SPACE(80)) AS Nomb_10, A.fcom_4 AS FECHA ,case when A.mone_4 ='1' then 'S/' else 'US$' end as MONEDA ,case when mone_4 ='1' then A.tot1_4 else  A.tot2_4 end as TOTAL , A.Aprobado_4 AS APROBADO
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '01' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '01' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '01' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			    Where A.CCia_4 = '" + Label4.Text + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND LEFT(A.NCom_4,2) = '" + Label5.Text + "'  Order By A.NCOM_4", conx)



            cmd55.Fill(da1)
            If da1.Rows.Count > 0 Then
                DataGridView1.DataSource = da1
            Else
                DataGridView1.DataSource = Nothing
            End If
            Dim uo As Integer
            uo = DataGridView1.Rows.Count

            For i = 0 To uo - 1
                DataGridView1.Columns(0).Width = 75
                DataGridView1.Columns(1).Width = 110
                DataGridView1.Columns(2).Width = 320
                DataGridView1.Columns(3).Width = 72

                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 70
                DataGridView1.Columns(6).Width = 73
                DataGridView1.Columns(6).Visible = False
                If DataGridView1.Rows(i).Cells(6).Value = "1" Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.IndianRed
                End If
            Next
        Else
            Dim cmd55 As New SqlDataAdapter(" SELECT A.ncom_4 as ORDEN , A.fich_4 AS RUC, ISNULL(C.Nomb_10,SPACE(80)) AS Nomb_10, A.fcom_4 AS FECHA ,case when A.mone_4 ='1' then 'S/' else 'US$' end as MONEDA ,case when mone_4 ='1' then A.tot1_4 else  A.tot2_4 end as TOTAL , A.Aprobado_4 AS APROBADO
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '01' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '01' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '01' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			 Where A.CCia_4 = '" + Label4.Text + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "'  AND LEFT(A.NCom_4,2) = '" + Label5.Text + "' AND ruc_10 ='" + Trim(Label7.Text) + "' Order By A.NCOM_4", conx)
            cmd55.Fill(da1)
            If da1.Rows.Count > 0 Then
                DataGridView1.DataSource = da1
            Else
                DataGridView1.DataSource = Nothing
            End If
            Dim uo As Integer
            uo = DataGridView1.Rows.Count

            For i = 0 To uo - 1
                DataGridView1.Columns(0).Width = 75
                DataGridView1.Columns(1).Width = 110
                DataGridView1.Columns(2).Width = 320
                DataGridView1.Columns(3).Width = 72

                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 70
                DataGridView1.Columns(6).Width = 73
                DataGridView1.Columns(6).Visible = False
                If DataGridView1.Rows(i).Cells(6).Value = "1" Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.IndianRed
                End If
            Next
        End If


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        da1.Clear()
        DataGridView1.DataSource = Nothing
        abrir()
        If Label6.Text = "1" Then
            Dim cmd55 As New SqlDataAdapter(" SELECT A.ncom_4 as ORDEN , A.fich_4 AS RUC, ISNULL(C.Nomb_10,SPACE(80)) AS Nomb_10, A.fcom_4 AS FECHA ,case when A.mone_4 ='1' then 'S/' else 'US$' end as MONEDA ,case when mone_4 ='1' then A.tot1_4 else  A.tot2_4 end as TOTAL , A.Aprobado_4 AS APROBADO
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '01' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '01' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '01' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			    Where A.CCia_4 = '" + Label4.Text + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND LEFT(A.NCom_4,2) = '" + Label5.Text + "'  Order By A.NCOM_4", conx)



            cmd55.Fill(da1)
            If da1.Rows.Count > 0 Then
                DataGridView1.DataSource = da1
            Else
                DataGridView1.DataSource = Nothing
            End If
            Dim uo As Integer
            uo = DataGridView1.Rows.Count

            For i = 0 To uo - 1
                DataGridView1.Columns(0).Width = 75
                DataGridView1.Columns(1).Width = 110
                DataGridView1.Columns(2).Width = 320
                DataGridView1.Columns(3).Width = 72

                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 70
                DataGridView1.Columns(6).Width = 73
                DataGridView1.Columns(6).Visible = False
                If DataGridView1.Rows(i).Cells(6).Value = "1" Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.IndianRed
                End If
            Next
        Else
            Dim cmd55 As New SqlDataAdapter(" SELECT A.ncom_4 as ORDEN , A.fich_4 AS RUC, ISNULL(C.Nomb_10,SPACE(80)) AS Nomb_10, A.fcom_4 AS FECHA ,case when A.mone_4 ='1' then 'S/' else 'US$' end as MONEDA ,case when mone_4 ='1' then A.tot1_4 else  A.tot2_4 end as TOTAL , A.Aprobado_4 AS APROBADO
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '01' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '01' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '01' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			 Where A.CCia_4 = '" + Label4.Text + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "'  AND LEFT(A.NCom_4,2) = '" + Label5.Text + "' AND ruc_10 ='" + Trim(Label7.Text) + "' Order By A.NCOM_4", conx)
            cmd55.Fill(da1)
            If da1.Rows.Count > 0 Then
                DataGridView1.DataSource = da1
            Else
                DataGridView1.DataSource = Nothing
            End If
            Dim uo As Integer
            uo = DataGridView1.Rows.Count

            For i = 0 To uo - 1
                DataGridView1.Columns(0).Width = 75
                DataGridView1.Columns(1).Width = 110
                DataGridView1.Columns(2).Width = 320
                DataGridView1.Columns(3).Width = 72

                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 70
                DataGridView1.Columns(6).Width = 73
                DataGridView1.Columns(6).Visible = False
                If DataGridView1.Rows(i).Cells(6).Value = "1" Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.IndianRed
                End If
            Next
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar0()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Label6.Text = "1" Then
            os.TextBox2.Text = Microsoft.VisualBasic.Right(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 6)
            os.TextBox2.Focus()
            SendKeys.Send("{ENTER}")
            Me.Close()
        Else
            Nia_Produccion.DataGridView1.Rows(Label8.Text).Cells(9).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
            Me.Close()
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub
End Class