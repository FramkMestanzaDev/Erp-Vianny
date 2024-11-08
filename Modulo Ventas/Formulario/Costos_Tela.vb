Imports System.Data.SqlClient
Public Class Costos_Tela
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
    Dim da As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim DT, DT1, PERIODO, CCIA As String
        DataGridView1.DataSource = ""
        da.Clear()
        'DT = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        'DT1 = Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "")
        Select Case Trim(ComboBox1.Text)
            Case "ENERO" : DT = "01"
            Case "FEBRERO" : DT = "02"
            Case "MARZO" : DT = "03"
            Case "ABRIL" : DT = "04"
            Case "MAYO" : DT = "05"
            Case "JUNIO" : DT = "06"
            Case "JULIO" : DT = "07"
            Case "AGOSTO" : DT = "08"
            Case "SETIEMBRE" : DT = "09"
            Case "OCTUBRE" : DT = "10"
            Case "NOVIEMBRE" : DT = "11"
            Case "DICIEMBRE" : DT = "12"
        End Select
        Select Case Trim(ComboBox2.Text)
            Case "ENERO" : DT1 = "01"
            Case "FEBRERO" : DT1 = "02"
            Case "MARZO" : DT1 = "03"
            Case "ABRIL" : DT1 = "04"
            Case "MAYO" : DT1 = "05"
            Case "JUNIO" : DT1 = "06"
            Case "JULIO" : DT1 = "07"
            Case "AGOSTO" : DT1 = "08"
            Case "SETIEMBRE" : DT1 = "09"
            Case "OCTUBRE" : DT1 = "10"
            Case "NOVIEMBRE" : DT1 = "11"
            Case "DICIEMBRE" : DT1 = "12"
        End Select
        PERIODO = Label3.Text
        CCIA = Label4.Text

        abrir()

        Dim cmd As New SqlDataAdapter("exec Costos_tela '" + DT + "','" + DT1 + "','" + PERIODO + "','" + CCIA + "'", conx)
        cmd.Fill(da)

        DataGridView1.DataSource = da
        Dim ex As New Exportar3
        ex.llenarExcel(DataGridView1)
    End Sub
End Class