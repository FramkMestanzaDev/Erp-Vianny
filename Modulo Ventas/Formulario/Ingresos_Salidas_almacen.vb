Imports System.Data.SqlClient
Public Class Ingresos_Salidas_almacen
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
    Private Sub Ingresos_Salidas_almacen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        If Label7.Text = "1" Or Label7.Text = "2" Then
            Dim cmd As New SqlDataAdapter(" SELECT A.ncom_3 as REGISTRO, A.fich_3 AS RUC,C.nomb_10 AS EMPRESA,E.nomb_21e AS MOTIVO, A.fcom_3 AS FECHA,A.grm_3 AS DOCUMENTO FROM custom_vianny.dbo.Mag030f A 
		   LEFT JOIN custom_vianny.dbo.cag1000 C ON A.CCIA_3 = C.ccia AND A.fich_3 = C.fich_10
		   LEFT JOIN custom_vianny.dbo.cag210e E ON A.CCIA_3 = E.Empr_21e AND A.orig_3 = E.dest_21e AND A.talm_3 = E.almac_21e
		   Where A.ccia_3 + a.cperiodo_3 + A.talm_3  ='" + Label3.Text & Label4.Text & Label5.Text + "'AND a.ccom_3 ='1' AND Transfer_3='0' AND A.ncoM_3 LIKE '" + Label6.Text + "%' order by ncom_3", conx)
            Dim da As New DataTable
            cmd.Fill(da)
            DataGridView1.DataSource = da

            DataGridView1.Columns(2).Width = 280
            DataGridView1.Columns(3).Width = 250
        Else
            If Label7.Text = "3" Or Label7.Text = "4" Then
                Dim cmd As New SqlDataAdapter("SELECT A.ncom_3 as REGISTRO, A.fich_3 AS RUC,C.nomb_10 AS EMPRESA,E.nomb_21s AS MOTIVO, A.fcom_3 AS FECHA,A.grm_3 AS DOCUMENTO FROM custom_vianny.dbo.Mag030f A 
		   LEFT JOIN custom_vianny.dbo.cag1000 C ON A.CCIA_3 = C.ccia AND A.fich_3 = C.fich_10
		   LEFT JOIN custom_vianny.dbo.cag210s E ON A.CCIA_3 = E.Empr_21s AND A.orig_3 = E.dest_21s AND A.talm_3 = E.almac_21s
		   Where A.ccia_3 + a.cperiodo_3 + A.talm_3  ='" + Label3.Text & Label4.Text & Label5.Text + "'AND a.ccom_3 ='2' AND Transfer_3='0' AND A.ncoM_3 LIKE '" + Label6.Text + "%' order by ncom_3", conx)
                Dim da As New DataTable
                cmd.Fill(da)
                DataGridView1.DataSource = da
                DataGridView1.Columns(2).Width = 280
                DataGridView1.Columns(3).Width = 250
            Else
                If Label7.Text = "5" Then
                    Dim cmd As New SqlDataAdapter("SELECT A.ncom_3 as REGISTRO, A.fich_3 AS RUC,C.nomb_10 AS EMPRESA,E.nomb_21s AS MOTIVO, A.fcom_3 AS FECHA,A.grm_3 AS DOCUMENTO FROM custom_vianny.dbo.qag3000 A 
		   LEFT JOIN custom_vianny.dbo.cag1000 C ON A.Empr_3 = C.ccia AND A.fich_3 = C.fich_10
		   LEFT JOIN custom_vianny.dbo.Qag210s E ON A.Empr_3 = E.Empr_21s AND A.orig_3 = E.dest_21s AND A.talm_3 = E.almac_21s
		   Where  A.Empr_3 + a.Ano_3 + A.talm_3  ='" + Label3.Text & Label4.Text & Label5.Text + "'AND a.ccom_3 ='2'  AND MONTH(A.fcom_3) = '" + Label6.Text + "' order by ncom_3", conx)
                    Dim da As New DataTable
                    cmd.Fill(da)
                    DataGridView1.DataSource = da
                    DataGridView1.Columns(2).Width = 280
                    DataGridView1.Columns(3).Width = 250
                Else
                    If Label7.Text = "6" Then
                        Dim cmd As New SqlDataAdapter("SELECT A.ncom_3 as REGISTRO, A.fich_3 AS RUC,C.nomb_10 AS EMPRESA,E.nomb_21e AS MOTIVO, A.fcom_3 AS FECHA,A.grm_3 AS DOCUMENTO FROM custom_vianny.dbo.qag3000 A 
		   LEFT JOIN custom_vianny.dbo.cag1000 C ON A.Empr_3 = C.ccia AND A.fich_3 = C.fich_10
		   LEFT JOIN custom_vianny.dbo.Qag210e E ON A.Empr_3 = E.Empr_21e AND A.orig_3 = E.dest_21e AND A.talm_3 = E.almac_21e
		   Where  A.Empr_3 + a.Ano_3 + A.talm_3  ='" + Label3.Text & Label4.Text & Label5.Text + "'AND a.ccom_3 ='1'  AND MONTH(A.fcom_3) = '" + Label6.Text + "' order by ncom_3", conx)
                        Dim da As New DataTable
                        cmd.Fill(da)
                        DataGridView1.DataSource = da
                        DataGridView1.Columns(2).Width = 280
                        DataGridView1.Columns(3).Width = 250
                    End If
                End If
            End If

        End If


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If Label7.Text = "2" Then
            NiaHilo.TextBox4.Text = Microsoft.VisualBasic.Left(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 2)
            NiaHilo.TextBox5.Text = Microsoft.VisualBasic.Right(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 6)

            NiaHilo.TextBox5.Focus()
            SendKeys.Send("{ENTER}")
            Me.Close()

        Else
            If Label7.Text = "1" Then
                NotaIngreso.TextBox4.Text = Microsoft.VisualBasic.Left(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 2)
                NotaIngreso.TextBox5.Text = Microsoft.VisualBasic.Right(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 6)
                NotaIngreso.TextBox5.Focus()
                SendKeys.Send("{ENTER}")
                Me.Close()
            Else

                If Label7.Text = "3" Then
                    NsaHilo.TextBox4.Text = Microsoft.VisualBasic.Left(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 2)
                    NsaHilo.TextBox5.Text = Microsoft.VisualBasic.Right(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 6)
                    NsaHilo.TextBox5.Focus()
                    SendKeys.Send("{ENTER}")
                    Me.Close()
                Else
                    If Label7.Text = "4" Then
                        Nsalida.TextBox4.Text = Microsoft.VisualBasic.Left(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 2)
                        Nsalida.TextBox5.Text = Microsoft.VisualBasic.Right(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 6)
                        Nsalida.TextBox5.Focus()
                        SendKeys.Send("{ENTER}")
                        Me.Close()
                    Else
                        If Label7.Text = "5" Then
                            nsa_produccion.TextBox4.Text = Microsoft.VisualBasic.Left(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 2)
                            nsa_produccion.TextBox5.Text = Microsoft.VisualBasic.Right(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 6)
                            nsa_produccion.TextBox5.Focus()
                            SendKeys.Send("{ENTER}")
                            Me.Close()
                        Else
                            If Label7.Text = "6" Then
                                Nia_Produccion.TextBox4.Text = Microsoft.VisualBasic.Left(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 2)
                                Nia_Produccion.TextBox5.Text = Microsoft.VisualBasic.Right(Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value), 6)
                                Nia_Produccion.TextBox5.Focus()
                                SendKeys.Send("{ENTER}")
                                Me.Close()
                            End If
                        End If
                    End If
                End If
            End If

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        If Label7.Text = "1" Or Label7.Text = "2" Then
            Dim cmd As New SqlDataAdapter(" SELECT A.ncom_3 as REGISTRO, A.fich_3 AS RUC,C.nomb_10 AS EMPRESA,E.nomb_21e AS MOTIVO, A.fcom_3 AS FECHA,A.grm_3 AS DOCUMENTO FROM custom_vianny.dbo.Mag030f A 
		   LEFT JOIN custom_vianny.dbo.cag1000 C ON A.CCIA_3 = C.ccia AND A.fich_3 = C.fich_10
		   LEFT JOIN custom_vianny.dbo.cag210e E ON A.CCIA_3 = E.Empr_21e AND A.orig_3 = E.dest_21e AND A.talm_3 = E.almac_21e
		   Where A.ccia_3 + a.cperiodo_3 + A.talm_3  ='" + Label3.Text & Label4.Text & Label5.Text + "'AND a.ccom_3 ='1' AND Transfer_3='0' AND (fcom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' and fcom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "') order by ncom_3", conx)
            Dim da As New DataTable
            cmd.Fill(da)
            DataGridView1.DataSource = da

            DataGridView1.Columns(2).Width = 280
            DataGridView1.Columns(3).Width = 250
        Else
            Dim cmd As New SqlDataAdapter(" SELECT A.ncom_3 as REGISTRO, A.fich_3 AS RUC,C.nomb_10 AS EMPRESA,E.nomb_21e AS MOTIVO, A.fcom_3 AS FECHA,A.grm_3 AS DOCUMENTO FROM custom_vianny.dbo.Mag030f A 
		   LEFT JOIN custom_vianny.dbo.cag1000 C ON A.CCIA_3 = C.ccia AND A.fich_3 = C.fich_10
		   LEFT JOIN custom_vianny.dbo.cag210e E ON A.CCIA_3 = E.Empr_21e AND A.orig_3 = E.dest_21e AND A.talm_3 = E.almac_21e
		   Where A.ccia_3 + a.cperiodo_3 + A.talm_3  ='" + Label3.Text & Label4.Text & Label5.Text + "'AND a.ccom_3 ='2' AND Transfer_3='0' AND (fcom_3 >= '" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "' and fcom_3 <= '" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "') order by ncom_3", conx)
            Dim da As New DataTable
            cmd.Fill(da)
            DataGridView1.DataSource = da

            DataGridView1.Columns(2).Width = 280
            DataGridView1.Columns(3).Width = 250
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class