Imports System.Data.SqlClient
Public Class Aprob_os
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

    Dim da As New DataTable
    Private Sub Aprob_os_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim cmd As New SqlDataAdapter("SELECT 
		       A.ccia_4     ,A.ncom_4  AS ORDEN   ,A.fich_4 AS RUC ,ISNULL(C.nomb_10,'') AS PROVEEDOR,    A.fcom_4 AS F_EMISION,
				ISNULL(A.fcome_4,'') AS F_ENTREGA,
				case when A.mone_4 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.mone_4 = '1' then	A.tot1_4 else A.tot2_4 end AS TOTAL,
		cast(aprobado_4  as char(1)) as ap
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '" + Trim(Label7.Text) + "' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '" + Trim(Label7.Text) + "' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			    Where A.CCia_4 = '" + Trim(Label7.Text) + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_4 = '0'  Order By A.NCOM_4", conx)

        cmd.Fill(da)
        DataGridView1.DataSource = da
        If DataGridView1.Rows.Count > 0 Then
            Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).Width = 80
            DataGridView1.Columns(5).Width = 300
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(8).Width = 70
        End If

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
            Dim cmd As New SqlDataAdapter("SELECT 
		       A.ccia_4     ,A.ncom_4  AS ORDEN   ,A.fich_4 AS RUC ,ISNULL(C.nomb_10,'') AS PROVEEDOR,    A.fcom_4 AS F_EMISION,
				ISNULL(A.fcome_4,'') AS F_ENTREGA,
				case when A.mone_4 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.mone_4 = '1' then	A.tot1_4 else A.tot2_4 end AS TOTAL,
		cast(aprobado_4  as char(1)) as ap
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '" + Trim(Label7.Text) + "' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '" + Trim(Label7.Text) + "' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			    Where A.CCia_4 = '" + Trim(Label7.Text) + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_4 = '0'  Order By A.NCOM_4", conx)

            cmd.Fill(da)
            DataGridView1.DataSource = da
            Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
        Else
            If RadioButton1.Checked = True Then
                Button2.Visible = True
                Button1.Visible = False
                abrir()
                Dim cmd As New SqlDataAdapter("SELECT 
		       A.ccia_4     ,A.ncom_4  AS ORDEN   ,A.fich_4 AS RUC ,ISNULL(C.nomb_10,'') AS PROVEEDOR,    A.fcom_4 AS F_EMISION,
				ISNULL(A.fcome_4,'') AS F_ENTREGA,
				case when A.mone_4 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.mone_4 = '1' then	A.tot1_4 else A.tot2_4 end AS TOTAL,
		cast(aprobado_4  as char(1)) as ap
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '" + Trim(Label7.Text) + "' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '" + Trim(Label7.Text) + "' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			    Where A.CCia_4 = '" + Trim(Label7.Text) + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_4 = '1'  Order By A.NCOM_4", conx)

                cmd.Fill(da)

                DataGridView1.DataSource = da
                Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
            Else
                Button2.Visible = True
                Button1.Visible = True
                abrir()
                Dim cmd As New SqlDataAdapter("SELECT 
		       A.ccia_4     ,A.ncom_4  AS ORDEN   ,A.fich_4 AS RUC ,ISNULL(C.nomb_10,'') AS PROVEEDOR,    A.fcom_4 AS F_EMISION,
				ISNULL(A.fcome_4,'') AS F_ENTREGA,
				case when A.mone_4 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.mone_4 = '1' then	A.tot1_4 else A.tot2_4 end AS TOTAL,
		cast(aprobado_4  as char(1)) as ap
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '" + Trim(Label7.Text) + "' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '" + Trim(Label7.Text) + "' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			    Where A.CCia_4 = '" + Trim(Label7.Text) + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "'   Order By A.NCOM_4", conx)

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        abrir()

        Dim cmd20 As New SqlCommand("UPDATE custom_vianny.dbo.Lag0400 SET Aprobado_4 = 1 ,UserApro_4 = @UserApro_4 WHERE CCia_4 = @CCia_4 AND NCom_4 = @NCom_4", conx)
                cmd20.Parameters.AddWithValue("@UserApro_4", Trim(MDIParent1.Label3.Text))
        cmd20.Parameters.AddWithValue("@NCom_4", Trim(Label11.Text))
        cmd20.Parameters.AddWithValue("@CCia_4", Trim(Label7.Text))

                cmd20.ExecuteNonQuery()

        MsgBox("SE APROBO CORRECTAMENTE")
        TextBox1.Text = ""
        TextBox2.Text = ""
        da.Clear()
        DataGridView1.DataSource = ""

        Dim cmd As New SqlDataAdapter("SELECT 
		       A.ccia_4     ,A.ncom_4  AS ORDEN   ,A.fich_4 AS RUC ,ISNULL(C.nomb_10,'') AS PROVEEDOR,    A.fcom_4 AS F_EMISION,
				ISNULL(A.fcome_4,'') AS F_ENTREGA,
				case when A.mone_4 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.mone_4 = '1' then	A.tot1_4 else A.tot2_4 end AS TOTAL,
		cast(aprobado_4  as char(1)) as ap
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '" + Trim(Label7.Text) + "' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '" + Trim(Label7.Text) + "' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			    Where A.CCia_4 = '" + Trim(Label7.Text) + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_4 = '0'  Order By A.NCOM_4", conx)

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

        Dim cmd21 As New SqlCommand("UPDATE custom_vianny.dbo.Lag0400 SET Aprobado_4 = 0, UserApro_4 = @UserApro_4 WHERE CCia_4 = @CCia_4 AND NCom_4 = @NCom_4", conx)
                cmd21.Parameters.AddWithValue("@UserApro_4", Trim(MDIParent1.Label3.Text))
        cmd21.Parameters.AddWithValue("@NCom_4", Trim(Label11.Text))
        cmd21.Parameters.AddWithValue("@CCia_4", Trim(Label7.Text))

                cmd21.ExecuteNonQuery()

        MsgBox("SE DESAPROBO CORRECTAMENTE")
        TextBox1.Text = ""
        TextBox2.Text = ""
        da.Clear()
        DataGridView1.DataSource = ""
        Dim cmd As New SqlDataAdapter("SELECT 
		       A.ccia_4     ,A.ncom_4  AS ORDEN   ,A.fich_4 AS RUC ,ISNULL(C.nomb_10,'') AS PROVEEDOR,    A.fcom_4 AS F_EMISION,
				ISNULL(A.fcome_4,'') AS F_ENTREGA,
				case when A.mone_4 ='1' then 'S/.' else 'US$' end AS MONEDA , 
			case when  A.mone_4 = '1' then	A.tot1_4 else A.tot2_4 end AS TOTAL,
		cast(aprobado_4  as char(1)) as ap
			   FROM custom_vianny.dbo.lag0400 A  LEFT OUTER JOIN custom_vianny.dbo.Qag0400 B
			   				   ON '" + Trim(Label7.Text) + "' = B.CCia AND A.Subfase_4 = B.fase_4 
			   				   LEFT OUTER JOIN custom_vianny.dbo.Cag1000 C
			   				   ON '" + Trim(Label7.Text) + "' = C.CCia AND A.Fich_4 = C.Fich_10
			   				   LEFT JOIN custom_vianny.dbo.Kag1500 D
			   				   ON '" + Trim(Label7.Text) + "' = D.CCia_15k AND A.Cond_4 = D.Cond_15k
			    Where A.CCia_4 = '" + Trim(Label7.Text) + "'  AND A.FCom_4 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.FCom_4 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "' AND A.Aprobado_4 = '1'  Order By A.NCOM_4", conx)

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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim j As Integer

        j = DataGridView1.Rows.Count

        For i = 0 To j - 1

            FORM_OS.TextBox1.Text = Trim(DataGridView1.Rows(Label10.Text).Cells(2).Value)
            FORM_OS.TextBox2.Text = Trim(DataGridView1.Rows(Label10.Text).Cells(2).Value)
            FORM_OS.TextBox3.Text = Label7.Text
            FORM_OS.TextBox4.Text = "\\172.21.0.1\erpcaesoft$\LIB.APPSV2\imagenes\"

            FORM_OS.Show()
        Next


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Label10.Text = e.RowIndex
        Label10.Text = DataGridView1.CurrentRow.Index
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Label10.Text = e.RowIndex
        Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
        If DataGridView1.Rows(e.RowIndex).Cells(10).Value = "1" Then
            Button2.Visible = True
            Button1.Visible = False
        Else
            Button1.Visible = True
            Button2.Visible = False
        End If

        If e.ColumnIndex = 1 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString

            FORM_OS.TextBox1.Text = Trim(DataGridView1.Rows(num1).Cells(3).Value)
            FORM_OS.TextBox2.Text = Trim(DataGridView1.Rows(num1).Cells(3).Value)
            FORM_OS.TextBox3.Text = Label7.Text
            FORM_OS.TextBox4.Text = "\\172.21.0.1\erpcaesoft$\LIB.APPSV2\imagenes\"

            FORM_OS.Show()


        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer


            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "ORDEN" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                Label11.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value
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
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        buscar1()
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

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub
End Class