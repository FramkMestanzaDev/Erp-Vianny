Imports System.Data.SqlClient
Public Class Apro_Oc
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Dim da, da1 As New DataTable
    Private Sub Apro_Oc_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = DateTimePicker1.Value.AddMonths(-1)
        RadioButton2.Checked = True
        If RadioButton2.Checked = True Then
            Button2.Visible = False
        Else
            Button2.Visible = True
        End If
        da.Clear()
        DataGridView1.DataSource = ""
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT A.CCIA_3     ,A.NCOM_3  AS ORDEN   ,A.FICH_3 AS RUC ,ISNULL(B.Nomb_10,'') AS PROVEEDOR,    A.FCOM_3 AS F_EMISION,
				ISNULL(A.FCOME_3,'') AS F_ENTREGA,
				case when A.MONE_3 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.MONE_3 = '1' then	A.TOT1_3 else A.TOT2_3 end AS TOTAL,
		cast(aprobado_3  as char(1)) as ap
				
				FROM custom_vianny.dbo.LAG0300 A LEFT OUTER JOIN custom_vianny.dbo.Cag1000 B
							 ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Fich_3 = B.Fich_10
							   LEFT JOIN custom_vianny.dbo.Kag1500 C
							 ON '" + Trim(Label7.Text) + "' = C.CCia_15k AND A.Cond_3 = C.Cond_15k
				 Where A.CCia_3 = '" + Trim(Label7.Text) + "'  AND A.FCom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_3 = '0' 
				 Order By A.NCOM_3", conx)

        cmd.Fill(da)

        DataGridView1.DataSource = da
        If DataGridView1.Rows.Count = 0 Then
            Label11.Text = 0
        Else
            Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
        End If

        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(10).Visible = False
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(5).Width = 300
        DataGridView1.Columns(6).Width = 80
        DataGridView1.Columns(7).Width = 80
        DataGridView1.Columns(8).Width = 70
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        da.Clear()
        DataGridView1.DataSource = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        If RadioButton2.Checked = True Then
            Button2.Visible = False
            Button1.Visible = True
            abrir()
            Dim cmd As New SqlDataAdapter("SELECT A.CCIA_3     ,A.NCOM_3  AS ORDEN   ,A.FICH_3 AS RUC ,ISNULL(B.Nomb_10,'') AS PROVEEDOR,    A.FCOM_3 AS F_EMISION,
				ISNULL(A.FCOME_3,'') AS F_ENTREGA,
				case when A.MONE_3 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.MONE_3 = '1' then	A.TOT1_3 else A.TOT2_3 end AS TOTAL,
	cast(aprobado_3  as char(1)) as ap
				
				FROM custom_vianny.dbo.LAG0300 A LEFT OUTER JOIN custom_vianny.dbo.Cag1000 B
							 ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Fich_3 = B.Fich_10
							   LEFT JOIN custom_vianny.dbo.Kag1500 C
							 ON '" + Trim(Label7.Text) + "' = C.CCia_15k AND A.Cond_3 = C.Cond_15k
				 Where A.CCia_3 = '" + Trim(Label7.Text) + "'  AND A.FCom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_3 = '0' 
				 Order By A.NCOM_3", conx)

            cmd.Fill(da)
            DataGridView1.DataSource = da
            Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
        Else
            If RadioButton1.Checked = True Then
                Button2.Visible = True
                Button1.Visible = False
                abrir()
                Dim cmd As New SqlDataAdapter("SELECT A.CCIA_3     ,A.NCOM_3  AS ORDEN   ,A.FICH_3 AS RUC ,ISNULL(B.Nomb_10,'') AS PROVEEDOR,    A.FCOM_3 AS F_EMISION,
				ISNULL(A.FCOME_3,'') AS F_ENTREGA,
				case when A.MONE_3 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.MONE_3 = '1' then	A.TOT1_3 else A.TOT2_3 end AS TOTAL,
	cast(aprobado_3  as char(1)) as ap
				
				FROM custom_vianny.dbo.LAG0300 A LEFT OUTER JOIN custom_vianny.dbo.Cag1000 B
							 ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Fich_3 = B.Fich_10
							   LEFT JOIN custom_vianny.dbo.Kag1500 C
							 ON '" + Trim(Label7.Text) + "' = C.CCia_15k AND A.Cond_3 = C.Cond_15k
				 Where A.CCia_3 = '" + Trim(Label7.Text) + "'  AND A.FCom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_3 = '1' 
				 Order By A.NCOM_3", conx)

                cmd.Fill(da)

                DataGridView1.DataSource = da
                Dim II As Integer
                II = DataGridView1.Rows.Count
                For I1 = 0 To II - 1
                    DataGridView1.Rows(I1).Cells(0).Value = True
                Next
                Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
            Else
                Button2.Visible = True
                Button1.Visible = True
                abrir()
                Dim cmd As New SqlDataAdapter("SELECT A.CCIA_3     ,A.NCOM_3  AS ORDEN   ,A.FICH_3 AS RUC ,ISNULL(B.Nomb_10,'') AS PROVEEDOR,    A.FCOM_3 AS F_EMISION,
				ISNULL(A.FCOME_3,'') AS F_ENTREGA,
				case when A.MONE_3 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.MONE_3 = '1' then	A.TOT1_3 else A.TOT2_3 end AS TOTAL,
			cast(aprobado_3  as char(1)) as ap
				
				FROM custom_vianny.dbo.LAG0300 A LEFT OUTER JOIN custom_vianny.dbo.Cag1000 B
							 ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Fich_3 = B.Fich_10
							   LEFT JOIN custom_vianny.dbo.Kag1500 C
							 ON '" + Trim(Label7.Text) + "' = C.CCia_15k AND A.Cond_3 = C.Cond_15k
				 Where A.CCia_3 = '" + Trim(Label7.Text) + "'  AND A.FCom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "'  
				 Order By A.NCOM_3", conx)

                cmd.Fill(da)
                DataGridView1.DataSource = da
                Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
            End If

        End If
        Dim y As Integer
        y = DataGridView1.Rows.Count
        For i = 0 To y - 1

            If Trim(DataGridView1.Rows(i).Cells(10).Value) = "1" Then

                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
            End If


        Next
        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(10).Visible = False
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(5).Width = 300
        DataGridView1.Columns(6).Width = 80
        DataGridView1.Columns(7).Width = 80
        DataGridView1.Columns(8).Width = 70

    End Sub
    Dim dv1 As DataView
    Private Sub buscar()
        Try

            Dim K As Integer
            Dim ds As New DataSet

            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "ORDEN" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value

                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(5).Width = 300
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(8).Width = 70
                Dim y As Integer
                y = DataGridView1.Rows.Count
                For i = 0 To y - 1

                    If Trim(DataGridView1.Rows(i).Cells(10).Value) = "1" Then

                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    End If


                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            Dim K As Integer


            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "PROVEEDOR" & " like '%" & TextBox2.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
                DataGridView1.Columns(2).Visible = False
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(5).Width = 300
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Width = 80
                DataGridView1.Columns(8).Width = 70
                Dim y As Integer
                y = DataGridView1.Rows.Count
                For i = 0 To y - 1

                    If Trim(DataGridView1.Rows(i).Cells(10).Value) = "1" Then

                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                    End If


                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Label10.Text = e.RowIndex
        Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Label10.Text = e.RowIndex

        If DataGridView1.Rows(e.RowIndex).Cells(10).Value = "1" Then
            Button2.Visible = True
            Button1.Visible = False
        Else
            Button1.Visible = True
            Button2.Visible = False
        End If
        If e.ColumnIndex = 1 Then
            Rpt_oc_1.TextBox1.Text = Label7.Text
            Rpt_oc_1.TextBox2.Text = Trim(Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(3).Value, 1, 2))
            Rpt_oc_1.TextBox3.Text = Trim(Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(3).Value, 3, 6))
            Rpt_oc_1.TextBox4.Text = Trim(Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(3).Value, 1, 2))
            Rpt_oc_1.TextBox5.Text = Trim(Microsoft.VisualBasic.Mid(DataGridView1.Rows(e.RowIndex).Cells(3).Value, 3, 6))
            Rpt_oc_1.TextBox6.Text = ""
            Rpt_oc_1.TextBox7.Text = 0
            Rpt_oc_1.TextBox8.Text = "\\172.21.0.1\erpcaesoft$\LIB.APPSV2\imagenes\"
            Rpt_oc_1.Show()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        abrir()



        Dim cmd20 As New SqlCommand("UPDATE custom_vianny.dbo.Lag0300 SET Aprobado_3 = 1, UserApro_3 = @UserApro_3 WHERE CCia_3 = @CCia_3 AND NCom_3 = @NCom_3", conx)
                cmd20.Parameters.AddWithValue("@UserApro_3", Trim(MDIParent1.Label3.Text))
        cmd20.Parameters.AddWithValue("@NCom_3", Trim(Label11.Text))
        cmd20.Parameters.AddWithValue("@CCia_3", Trim(Label7.Text))

                cmd20.ExecuteNonQuery()

        MsgBox("SE APROBO CORRECTAMENTE")
        TextBox1.Text = ""
        TextBox2.Text = ""
        da.Clear()
        DataGridView1.DataSource = ""

        Dim cmd As New SqlDataAdapter("SELECT A.CCIA_3     ,A.NCOM_3  AS ORDEN   ,A.FICH_3 AS RUC ,ISNULL(B.Nomb_10,'') AS PROVEEDOR,    A.FCOM_3 AS F_EMISION,
				ISNULL(A.FCOME_3,'') AS F_ENTREGA,
				case when A.MONE_3 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.MONE_3 = '1' then	A.TOT1_3 else A.TOT2_3 end AS TOTAL,
	cast(aprobado_3  as char(1)) as ap
				
				FROM custom_vianny.dbo.LAG0300 A LEFT OUTER JOIN custom_vianny.dbo.Cag1000 B
							 ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Fich_3 = B.Fich_10
							   LEFT JOIN custom_vianny.dbo.Kag1500 C
							 ON '" + Trim(Label7.Text) + "' = C.CCia_15k AND A.Cond_3 = C.Cond_15k
				 Where A.CCia_3 = '" + Trim(Label7.Text) + "'  AND A.FCom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_3 = '0' 
				 Order By A.NCOM_3", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        Dim y As Integer
        y = DataGridView1.Rows.Count
        If y > 0 Then
            For i = 0 To y - 1
                If Trim(DataGridView1.Rows(i).Cells(10).Value) = "1" Then

                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
                End If
            Next
            Label11.Text = DataGridView1.Rows(0).Cells(3).Value
        End If

        Label10.Text = 0

        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(10).Visible = False
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(5).Width = 300
        DataGridView1.Columns(6).Width = 80
        DataGridView1.Columns(7).Width = 80
        DataGridView1.Columns(8).Width = 70
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        abrir()
        Dim j As Integer

        Dim cmd20 As New SqlCommand("UPDATE custom_vianny.dbo.Lag0300 SET Aprobado_3 = 0, UserApro_3 = @UserApro_3 WHERE CCia_3 = @CCia_3 AND NCom_3 = @NCom_3", conx)
        cmd20.Parameters.AddWithValue("@UserApro_3", Trim(MDIParent1.Label3.Text))
        cmd20.Parameters.AddWithValue("@NCom_3", Trim(Label11.Text))
        cmd20.Parameters.AddWithValue("@CCia_3", Trim(Label7.Text))
        cmd20.ExecuteNonQuery()

        MsgBox("SE DESAPROBO CORRECTAMENTE")
        TextBox1.Text = ""
        TextBox2.Text = ""
        da.Clear()
        DataGridView1.DataSource = ""

        Dim cmd As New SqlDataAdapter("SELECT A.CCIA_3     ,A.NCOM_3  AS ORDEN   ,A.FICH_3 AS RUC ,ISNULL(B.Nomb_10,'') AS PROVEEDOR,    A.FCOM_3 AS F_EMISION,
        			ISNULL(A.FCOME_3,'') AS F_ENTREGA,
        			case when A.MONE_3 ='1' then 'S/.' else 'US$' end AS MONEDA , 
        		case when  A.MONE_3 = '1' then	A.TOT1_3 else A.TOT2_3 end AS TOTAL,
        cast(aprobado_3  as char(1)) as ap

        			FROM custom_vianny.dbo.LAG0300 A LEFT OUTER JOIN custom_vianny.dbo.Cag1000 B
        						 ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Fich_3 = B.Fich_10
        						   LEFT JOIN custom_vianny.dbo.Kag1500 C
        						 ON '" + Trim(Label7.Text) + "' = C.CCia_15k AND A.Cond_3 = C.Cond_15k
        			 Where A.CCia_3 = '" + Trim(Label7.Text) + "'  AND A.FCom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_3 = '1' 
        			 Order By A.NCOM_3", conx)

        cmd.Fill(da)

        DataGridView1.DataSource = da
        Dim y As Integer
        y = DataGridView1.Rows.Count
        For i = 0 To y - 1

            If Trim(DataGridView1.Rows(i).Cells(10).Value) = "1" Then

                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White
            End If


        Next
        DataGridView1.Columns(2).Visible = False
        DataGridView1.Columns(10).Visible = False
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(5).Width = 300
        DataGridView1.Columns(6).Width = 80
        DataGridView1.Columns(7).Width = 80
        DataGridView1.Columns(8).Width = 70
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode = Keys.Up Then
            Dim oc = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(3).Value
            Label11.Text = oc
        End If
        If e.KeyCode = Keys.Down Then
            Dim oc1 = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(3).Value
            Label11.Text = oc1
        End If
    End Sub
End Class