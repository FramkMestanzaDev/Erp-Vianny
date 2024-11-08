Imports System.Data.SqlClient
Public Class Valorizr_Salidas
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
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
        Motivos_valorizar.Label1.Text = Label4.Text
        Motivos_valorizar.Label2.Text = TextBox3.Text
        Motivos_valorizar.Label3.Text = 1
        Motivos_valorizar.ShowDialog()
    End Sub

    Private Sub Valorizr_Salidas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box2()
        ComboBox2.SelectedIndex = 0
        TextBox3.Text = "00"
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox3.Text = ComboBox1.SelectedValue.ToString
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
    Dim GK As New DataTable
    Dim data As New DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim TG As New fvsalidas
        Dim HJ1 As New vvsalidas
        Dim HJ As New vvsalidas
        HJ1.gmes = Trim(TextBox4.Text)
        HJ1.gano = Label5.Text
        HJ1.galmacen = Trim(TextBox3.Text)
        HJ1.gempresa = Label4.Text

        'HJ.gmotivos = Trim(TextBox1.Text)

        data = TG.listar_codigos(HJ1)
        'MsgBox(data.Rows.Count.ToString())
        If data.Rows.Count > 0 Then
            For i = 0 To data.Rows.Count - 1
                HJ.gmes = Trim(TextBox4.Text)
                HJ.gano = Label5.Text
                HJ.galmacen = Trim(TextBox3.Text)
                HJ.gempresa = Label4.Text
                HJ.gmotivos = Trim(TextBox1.Text)
                HJ.gcodigo = data.Rows(i)("linea_3a").ToString()
                'MsgBox(data.Rows(i)("linea_3a").ToString())
                TG.Valorizar_Salidas(HJ)
            Next
            MsgBox("SE VALORIZO LAS SALIDAS DEL ALMACEN SOLICITADO")
        End If







    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Select Case Trim(ComboBox2.Text)
            Case "ENERO" : TextBox4.Text = "01"
            Case "FEBRERO" : TextBox4.Text = "02"
            Case "MARZO" : TextBox4.Text = "03"
            Case "ABRIL" : TextBox4.Text = "04"
            Case "MAYO" : TextBox4.Text = "05"
            Case "JUNIO" : TextBox4.Text = "06"
            Case "JULIO" : TextBox4.Text = "07"
            Case "AGOSTO" : TextBox4.Text = "08"
            Case "SETIEMBRE" : TextBox4.Text = "09"
            Case "OCTUBRE" : TextBox4.Text = "10"
            Case "NOVIEMBRE" : TextBox4.Text = "11"
            Case "DICIEMBRE" : TextBox4.Text = "12"
        End Select
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Motivos_valorizar.Label1.Text = Label4.Text
        Motivos_valorizar.Label2.Text = TextBox3.Text
        Motivos_valorizar.Label3.Text = 2
        Motivos_valorizar.ShowDialog()
    End Sub

    Sub llenar_combo_box2()

        Try

            conn = New SqlDataAdapter(" select talm_15m as almacen, talm_15m + ' | '+ nomb_15m as descrip from custom_vianny.dbo.Mag1500 where ccia ='" + Label4.Text + "'", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "descrip"
            ComboBox1.ValueMember = "almacen"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class