Imports System.Data.SqlClient
Public Class TablaColor
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
    Dim dt As DataTable
    Dim Rsr20 As SqlDataReader
    Dim DT14 As New DataTable
    Private Sub TablaColor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrarCorrelativo()
        MOSTRAR_INFORMACION()
    End Sub
    Public Sub MOSTRAR_INFORMACION()
        abrir()
        DataGridView1.DataSource = ""
        DT14.Clear()
        Dim cmd6 As New SqlDataAdapter("select ccolor_3c as CODIGO,cnomb_3c as COLOR from custom_vianny.dbo.qarc0300 where ccia_3c ='01'", conx)
        cmd6.Fill(DT14)
        DataGridView1.DataSource = DT14
        DataGridView1.Columns(0).Width = 110
        DataGridView1.Columns(1).Width = 380
    End Sub
    Sub mostrarCorrelativo()
        abrir()
        Dim sql10212 As String = "SELECT MAX(CAST( SUBSTRING(ccolor_3c,2,5) AS int)) FROM  custom_vianny.dbo.Qarc0300   Where ccia_3c ='01'"
        Dim cmd10212 As New SqlCommand(sql10212, conx)
        Rsr20 = cmd10212.ExecuteReader()
        If Rsr20.Read() = True Then
            TextBox5.Text = Convert.ToInt64(Rsr20(0)) + 1
        End If
        Select Case TextBox5.Text.Length
            Case "1" : TextBox5.Text = "0000" & TextBox5.Text
            Case "2" : TextBox5.Text = "000" & TextBox5.Text
            Case "3" : TextBox5.Text = "00" & TextBox5.Text
            Case "4" : TextBox5.Text = "0" & TextBox5.Text
            Case "5" : TextBox5.Text = TextBox5.Text

        End Select

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DT14.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "COLOR" & " like '%" & TextBox4.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 110
                DataGridView1.Columns(1).Width = 380
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        buscar()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim cmd20 As New SqlCommand("insert into custom_vianny.dbo.qarc0300 (ccia_3c,ccolor_3c ,version_3c,tiprec_3c,idrecet_3c,cnomb_3c,cdsctlm_3c,csigla_3c   ,cclient_3c,tiptela_3c,color_3c,fecemi_3c,lote_3c,ruclote_3c,relbano_3c,curva_3c,obser_3c,user_3c,fupdate_3c,flag_3c,rapor_3c,sdc_3c) 
								values ('01'   ,@ccolor_3c, ''       ,''       ,''        ,@cnomb_3c,@cdsctlm_3c,@csigla_3c,''        ,''        ,''      , null    ,''     ,''        ,0         ,0       ,''      ,''     ,null      ,0      ,0 ,'' )", conx)
        cmd20.Parameters.AddWithValue("@ccolor_3c", Trim(TextBox1.Text) & Trim(TextBox5.Text))
        cmd20.Parameters.AddWithValue("@cnomb_3c", Mid(TextBox2.Text.ToString().Trim(), 1, 55))
        cmd20.Parameters.AddWithValue("@cdsctlm_3c", Mid(TextBox2.Text.ToString().Trim(), 1, 40))
        cmd20.Parameters.AddWithValue("@csigla_3c", Mid(TextBox2.Text.ToString().Trim(), 1, 20))
        cmd20.ExecuteNonQuery()
        MsgBox(" SE REGISTRO CORRECTAMENTE LA INFORMACION")
        limpiar()
        MOSTRAR_INFORMACION()
        mostrarCorrelativo()
    End Sub
    Sub limpiar()
        TextBox2.Text = ""
        TextBox4.Text = ""
    End Sub
End Class