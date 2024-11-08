
Imports System.Data.SqlClient
Public Class Despacho_Pedido
    Dim fecha1, fecha2 As Date
    Dim empresa As String
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        fecha1 = DateTimePicker1.Value.ToString("yyyy-MM-dd")
        fecha2 = DateTimePicker2.Value.ToString("yyyy-MM-dd")
        'empresa = Label5.Text
        TextBox1.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        TextBox2.Text = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
        Dim objreporte As New rp_despacho
        objreporte.SetDatabaseLogon("sa", "Vi@Gr@Tex2005%")
        objreporte.SetParameterValue("@CCIA", Label5.Text)
        objreporte.SetParameterValue("@FECHAFIN", Trim(TextBox2.Text))
        objreporte.SetParameterValue("@FECHAINI", Trim(TextBox1.Text))
        If Trim(ComboBox1.Text) = "TODOS" Then

            objreporte.SetParameterValue("@CHOFER", DBNull.Value)
        Else
            objreporte.SetParameterValue("@CHOFER", Trim(ComboBox1.Text))
        End If
        If Trim(ComboBox2.Text) = "TODOS" Then
            objreporte.SetParameterValue("@placa", DBNull.Value)
        Else
            objreporte.SetParameterValue("@placa", Trim(ComboBox2.Text))
        End If


        CrystalReportViewer1.ReportSource = objreporte

    End Sub

    Private Sub Despacho_Pedido_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box()
        llenar_combo_box1()
    End Sub
    Sub llenar_combo_box()
        Try
            Dim lsQuery As String = "select 'TODOS' as chofer
union all
select distinct chofer  from ingre_transportista"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            ComboBox1.DataSource = loDataTable

            ComboBox1.DisplayMember = "chofer"
            ComboBox1.ValueMember = "chofer"
            ComboBox1.DropDownWidth = 200
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box1()
        Try
            Dim lsQuery As String = "select 'TODOS' as placa
union all
select distinct placa  from ingre_transportista"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            ComboBox2.DataSource = loDataTable

            ComboBox2.DisplayMember = "placa"
            ComboBox2.ValueMember = "placa"
            ComboBox2.DropDownWidth = 200
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load


    End Sub
End Class