Imports System.Data.SqlClient
Public Class TablaCLE
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Dim Rsr20, Rsr21, Rsr22 As SqlDataReader
    Dim DT14, DT15, DT16 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim deta As String
        deta = ""
        If Label3.Text = "1" Then
            deta = "UDPLAV"
        Else
            If Label3.Text = "2" Then
                deta = "UDPCOL"
            Else
                If Label3.Text = "3" Then
                    deta = "UDPEFE"
                End If
            End If
        End If
        Dim cmd20 As New SqlCommand("INSERT INTO custom_vianny.dbo.TAB0100 (ccia,ctab,cele,dele,nele,codigo,codigo2) VALUES (@ccia,@ctab,@cele,@dele,@nele,@codigo,@codigo2)", conx)
        cmd20.Parameters.AddWithValue("@ccia", Trim(Label5.Text))
        cmd20.Parameters.AddWithValue("@ctab", deta)
        cmd20.Parameters.AddWithValue("@cele", TextBox5.Text)
        cmd20.Parameters.AddWithValue("@dele", TextBox2.Text.ToString.Trim)
        cmd20.Parameters.AddWithValue("@nele", 0)
        cmd20.Parameters.AddWithValue("@codigo", TextBox2.Text.ToString.Trim)
        cmd20.Parameters.AddWithValue("@codigo2", "")
        cmd20.ExecuteNonQuery()
        MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
        Lavados.actualizarinformacion()
        LIMPIAR()
        Me.Close()
    End Sub
    Sub LIMPIAR()
        TextBox2.Text = ""
        TextBox4.Text = ""
    End Sub
    Sub CorrelativoLavado()
        abrir()
        Dim sql10212 As String = "SELECT MAX(cele)   FROM custom_vianny.dbo.TAB0100 A  Where A.CCia = '01' AND A.CTab = 'UDPLAV'"
        Dim cmd10212 As New SqlCommand(sql10212, conx)
        Rsr20 = cmd10212.ExecuteReader()
        If Rsr20.Read() = True Then
            TextBox5.Text = Convert.ToInt64(Rsr20(0)) + 1
        End If
        Select Case TextBox5.Text.Length
            Case "1" : TextBox5.Text = "000" & TextBox5.Text
            Case "2" : TextBox5.Text = "00" & TextBox5.Text
            Case "3" : TextBox5.Text = "0" & TextBox5.Text
            Case "4" : TextBox5.Text = TextBox5.Text
        End Select
    End Sub
    Sub CorrelativoColor()
        abrir()
        Dim sql10212 As String = "SELECT MAX(cele)   FROM custom_vianny.dbo.TAB0100 A  Where A.CCia = '01' AND A.CTab = 'UDPCOL'"
        Dim cmd10212 As New SqlCommand(sql10212, conx)
        Rsr21 = cmd10212.ExecuteReader()
        If Rsr21.Read() = True Then
            TextBox5.Text = Convert.ToInt64(Rsr21(0)) + 1
        End If
        Select Case TextBox5.Text.Length
            Case "1" : TextBox5.Text = "000" & TextBox5.Text
            Case "2" : TextBox5.Text = "00" & TextBox5.Text
            Case "3" : TextBox5.Text = "0" & TextBox5.Text
            Case "4" : TextBox5.Text = TextBox5.Text
        End Select
    End Sub
    Sub CorrelativoEfecto()
        abrir()
        Dim sql10212 As String = "SELECT MAX(cele)   FROM custom_vianny.dbo.TAB0100 A  Where A.CCia = '01' AND A.CTab = 'UDPEFE'"
        Dim cmd10212 As New SqlCommand(sql10212, conx)
        Rsr22 = cmd10212.ExecuteReader()
        If Rsr22.Read() = True Then
            TextBox5.Text = Convert.ToInt64(Rsr22(0)) + 1
        End If
        Select Case TextBox5.Text.Length
            Case "1" : TextBox5.Text = "000" & TextBox5.Text
            Case "2" : TextBox5.Text = "00" & TextBox5.Text
            Case "3" : TextBox5.Text = "0" & TextBox5.Text
            Case "4" : TextBox5.Text = TextBox5.Text
        End Select
    End Sub
    Public Sub MOSTRAR_INFORMACIONLAVADO()
        abrir()
        DataGridView1.DataSource = ""
        DT14.Clear()
        Dim cmd6 As New SqlDataAdapter("SELECT ctab AS AREA,cele AS CODIGO,DELE AS DETALLE FROM custom_vianny.dbo.TAB0100 A  Where A.CCia = '01' AND A.CTab = 'UDPLAV'", conx)
        cmd6.Fill(DT14)
        DataGridView1.DataSource = DT14
        DataGridView1.Columns(2).Width = 220
        DataGridView1.Columns(1).Width = 220
        DataGridView1.Columns(0).Width = 85

    End Sub
    Public Sub MOSTRAR_INFORMACIONCOLOR()
        abrir()
        DataGridView1.DataSource = ""
        DT15.Clear()
        Dim cmd6 As New SqlDataAdapter("SELECT ctab AS AREA,cele AS CODIGO,DELE AS DETALLE FROM custom_vianny.dbo.TAB0100 A  Where A.CCia = '01' AND A.CTab = 'UDPCOL'", conx)
        cmd6.Fill(DT15)
        DataGridView1.DataSource = DT15
        DataGridView1.Columns(2).Width = 220
        DataGridView1.Columns(1).Width = 220
        DataGridView1.Columns(0).Width = 85

    End Sub
    Public Sub MOSTRAR_INFORMACIONEFECTO()
        abrir()
        DataGridView1.DataSource = ""
        DT15.Clear()
        Dim cmd6 As New SqlDataAdapter("SELECT ctab AS AREA,cele AS CODIGO,DELE AS DETALLE FROM custom_vianny.dbo.TAB0100 A  Where A.CCia = '01' AND A.CTab = 'UDPEFE'", conx)
        cmd6.Fill(DT15)
        DataGridView1.DataSource = DT15
        DataGridView1.Columns(2).Width = 220
        DataGridView1.Columns(1).Width = 220
        DataGridView1.Columns(0).Width = 85

    End Sub
    Private Sub TablaCLE_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Label3.Text = "1" Then
            CorrelativoLavado()
            MOSTRAR_INFORMACIONLAVADO()
        Else
            If Label3.Text = "2" Then
                CorrelativoColor()
                MOSTRAR_INFORMACIONCOLOR()
            Else
                If Label3.Text = "3" Then
                    CorrelativoEfecto()
                    MOSTRAR_INFORMACIONEFECTO()
                End If
            End If
        End If
    End Sub
End Class