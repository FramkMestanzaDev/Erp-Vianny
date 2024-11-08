Imports System.Data.SqlClient
Public Class ObsMolde
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
    Dim Rsr21 As SqlDataReader
    Private Sub ObsMolde_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        abrir()
        Dim cadena As String = ""
        Dim sql1021 As String = "SELECT 
        '   Op: ' + RTRIM(ltrim( qg.ncom_3)) + char(10)+ 
        '   Modelista: ' + RTRIM(ltrim( cg.nomb_10)) + char(10)+ 
        '   Fecha de Registro: ' + CONVERT(VARCHAR(10), fcome4_3, 120) + 
        '   Ruta del Molde: ' + CAST(mruta7_3 AS VARCHAR(MAX)) + 
        '   Lavandería: ' + CAST(scobs1_3 AS VARCHAR(MAX))
            FROM custom_vianny.dbo.qag0300 qg
            LEFT JOIN custom_vianny.dbo.cag1000 cg ON qg.ccia = cg.ccia 
            AND qg.modelista_3 = cg.ruc_10
            WHERE qg.ccia = '01' 
           AND ncom_3 = '" + Label1.Text + "'"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr21 = cmd1021.ExecuteReader()
        Dim i51 As Integer
        i51 = 0

        If Rsr21.Read() Then
            cadena = Rsr21(0).ToString().Trim()
        End If

        Rsr21.Close()
        TextBox1.Text = cadena
    End Sub
End Class