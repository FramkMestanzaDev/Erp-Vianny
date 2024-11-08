Imports System.ComponentModel
Imports System.Data.SqlClient
Public Class Ingreso_salida_Produ
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
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.Text
            Case "CORTE" : TextBox2.Text = "0300"
            Case "APLICACIONES Y PIEZAS" : TextBox2.Text = "0700"
            Case "CONFECCION" : TextBox2.Text = "0400"
            Case "LAVANDERIA" : TextBox2.Text = "0600"
            Case "ACABADOS" : TextBox2.Text = "0500"
        End Select
    End Sub
    Dim da As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        DataGridView1.DataSource = ""
        da.Clear()
        If Trim(ComboBox1.Text) = "" Or Trim(ComboBox2.Text) = "" Then
            MsgBox("FALTA SELECCIONAR EL AREA O EL QUIEBRE")
        Else
            Dim com As String
            If ComboBox2.Text = "INGRESO" Then
                com = "1"
            Else
                com = "2"
            End If
            Dim DT, DT1 As String
            DT = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
            DT1 = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
            abrir()
            If Trim(TextBox1.Text) = "" Then
                Dim cmd As New SqlDataAdapter("EXEC custom_vianny.DBO.CaeSoft_ReportePartedeIngresoSalidaProduccion '" + Label6.Text + "','" + TextBox2.Text + "','" + com + "','" + DT + "','" + DT1 + "',NULL,NULL,NULL,'INTEXT',NULL,1,6", conx)
                cmd.Fill(da)
            Else

                MsgBox(TextBox1.Text)
                Dim cmd As New SqlDataAdapter("EXEC custom_vianny.DBO.CaeSoft_ReportePartedeIngresoSalidaProduccion '" + Label6.Text + "','" + TextBox2.Text + "','" + com + "','" + DT + "','" + DT1 + "','" + TextBox1.Text + "',NULL,NULL,'INTEXT',NULL,6,2", conx)
                cmd.Fill(da)
            End If


            DataGridView1.DataSource = da
                Dim ex As New Exportar
                ex.llenarExcel(DataGridView1)
            End If


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim com As String
        If ComboBox2.Text = "INGRESO" Then
            com = "1"
        Else
            com = "2"
        End If
        Dim DT, DT1 As String
        DT = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        DT1 = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
        Ingreso_salida_Prod.TextBox1.Text = Label6.Text
        Ingreso_salida_Prod.TextBox2.Text = TextBox2.Text
        Ingreso_salida_Prod.TextBox3.Text = com
        Ingreso_salida_Prod.TextBox4.Text = DT
        Ingreso_salida_Prod.TextBox5.Text = DT1
        Ingreso_salida_Prod.TextBox6.Text = TextBox1.Text
        Ingreso_salida_Prod.Show()
    End Sub

    Private Sub Ingreso_salida_Produ_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class