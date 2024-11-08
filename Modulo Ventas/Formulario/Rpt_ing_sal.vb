Imports System.Data.SqlClient
Public Class Rpt_ing_sal
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
    Dim da As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        da.Clear()
        Dim r, part, ruc As String
        If Trim(TextBox1.Text) = "" Then
            ruc = "NULL"
        Else
            ruc = Trim(TextBox1.Text)
        End If
        If Trim(TextBox2.Text) = "" Then
            part = "NULL"
        Else
            part = Trim(TextBox1.Text)
        End If
        Dim fini As Date
        Dim fsal As Date
        Select Case ComboBox2.Text
            Case "INGRESO" : r = "1"
            Case "SALIDA" : r = "2"

        End Select
        fini = DateTimePicker1.Value
        fsal = DateTimePicker2.Value

        If ComboBox1.Text = "TELA PLANTA HUACHIPA" Then

            Dim cmd As New SqlDataAdapter("EXEC CaeSoft_ReportePartedeIngresoSalida_2021 '" + Label7.Text + "','68','38','39','" + r + "','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "','" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "',NULL,NULL", conx)
            cmd.Fill(da)
        Else
            If ComboBox1.Text = "HILADO TEÑIDO" Then
                Dim cmd As New SqlDataAdapter("EXEC CaeSoft_ReportePartedeIngresoSalida_2021 '" + Label7.Text + "','42',NULL,NULL,'" + r + "','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "','" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "',NULL,NULL", conx)
                cmd.Fill(da)

            Else
                If ComboBox1.Text = "HILA CRUDO" Then
                    Dim cmd As New SqlDataAdapter("EXEC CaeSoft_ReportePartedeIngresoSalida_2021 '" + Label7.Text + "','06',NULL,NULL,'" + r + "','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "','" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "',NULL,NULL", conx)
                    cmd.Fill(da)
                Else
                    If ComboBox1.Text = "AVIOS" Then
                        Dim cmd As New SqlDataAdapter("EXEC CaeSoft_ReportePartedeIngresoSalida_2021 '" + Label7.Text + "','13',NULL,NULL,'" + r + "','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "','" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "',NULL,NULL", conx)
                        cmd.Fill(da)
                    Else
                        Dim cmd As New SqlDataAdapter("EXEC CaeSoft_ReportePartedeIngresoSalida_2021 '" + Label7.Text + "','10','51','54','" + r + "','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "','" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "',NULL,NULL", conx)
                        cmd.Fill(da)
                    End If

                End If

            End If

        End If


            If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da

        End If
        Dim gh As New Exportar
        gh.llenarExcel(DataGridView1)
    End Sub

    Private Sub Rpt_ing_sal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
    End Sub
End Class