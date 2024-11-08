Imports System.Data.SqlClient
Public Class Almacen_guia_nuevo
    Public conx As SqlConnection
    Public conn As SqlDataAdapter

    Public comando As SqlCommand
    Dim da As New DataTable
    Dim bs As New BindingSource()
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Almacen_guia_nuevo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        da.Clear()
        abrir()
        Dim CONSULTA As String
        If Label4.Text = "2" Then
            CONSULTA = "select ccia,talm_15m AS ALMACEN,nomb_15m AS DESCRIPCION  from custom_vianny.dbo.Mag1500 where ccia ='" + Label2.Text + "' and  seriencr_15m ='2' and flag_15m ='1'"
        Else
            If Label4.Text = "1" Then
                CONSULTA = "select ccia,talm_15m AS ALMACEN,nomb_15m AS DESCRIPCION  from custom_vianny.dbo.Mag1500 where ccia ='" + Label2.Text + "' and  seriencr_15m ='1' and flag_15m ='1'"
            Else
                If Label4.Text = "4" Then
                    CONSULTA = "select ccia,talm_15m AS ALMACEN,nomb_15m AS DESCRIPCION  from custom_vianny.dbo.Mag1500 where ccia ='" + Label2.Text + "' and  seriencr_15m ='4' and flag_15m ='1'"
                Else
                    CONSULTA = "select ccia,talm_15m AS ALMACEN,nomb_15m AS DESCRIPCION  from custom_vianny.dbo.Mag1500 where ccia ='" + Label2.Text + "' and  seriencr_15m ='3' and flag_15m ='1'"

                End If

            End If
        End If
        Dim cmd As New SqlDataAdapter(CONSULTA, conx)
        cmd.SelectCommand.CommandTimeout = 300
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).Width = 350
            DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridView1.Columns(2).DefaultCellStyle.BackColor = Color.AliceBlue
            DataGridView1.Columns(2).DefaultCellStyle.ForeColor = Color.Red
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) = "03" Or Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) = "42" Or Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) = "41" Or Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) = "14" Or Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) = "45" Then
            Guia_hilo.Close()
            Guia_hilo.Label23.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Guia_hilo.Label25.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
            Guia_hilo.TextBox19.Text = Label3.Text
            Guia_hilo.Label29.Text = Label1.Text
            Guia_hilo.Label30.Text = Label2.Text
            Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                Case "03" : Guia_hilo.TextBox1.Text = "T003"
                Case "42" : Guia_hilo.TextBox1.Text = "T003"
                Case "14" : Guia_hilo.TextBox1.Text = "T005"
                Case "41" : Guia_hilo.TextBox1.Text = "T005"
                Case "59" : Guia_hilo.TextBox1.Text = "T006"
                Case "45" : Guia_hilo.TextBox1.Text = "T007"
            End Select
            Guia_hilo.Show()

        Else
            Guia_remision.Close()
            Guia_remision.Label23.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Guia_remision.Label25.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
            Guia_remision.TextBox19.Text = Label3.Text
            Guia_remision.Label29.Text = Label1.Text
            Guia_remision.Label30.Text = Label2.Text
            Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)

                Case "10" : Guia_remision.TextBox3.Text = "JIRON LOS OLMOS 358 URB. CANTOB BELLO"


            End Select
            Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)

                Case "10" : Guia_remision.TextBox1.Text = "T004"
                Case "06" : Guia_remision.TextBox1.Text = "T004"
                Case "08" : Guia_remision.TextBox1.Text = "T004"
                Case "24" : Guia_remision.TextBox1.Text = "T005"
            End Select
            Guia_remision.Label35.Text = Label6.Text
            Guia_remision.Label36.Text = Label5.Text
            Guia_remision.ShowDialog()
        End If

    End Sub
End Class