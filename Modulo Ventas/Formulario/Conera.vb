Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class Conera
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2 As New DataTable
    Public enunciado4 As SqlCommand
    Public respuesta4 As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public enunciado5 As SqlCommand
    Public respuesta5 As SqlDataReader
    Public enunciado6 As SqlCommand
    Public respuesta6 As SqlDataReader
    Public enunciado7 As SqlCommand
    Public respuesta7 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub llenar_combo_box2(ByVal cb As ComboBox, ByVal alm As String)
        Try

            conn = New SqlDataAdapter("select urdido from programa_tenido_detalle where  programa =" + alm, conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "urdido"

            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add(1)
        Dim ui As Integer
        ui = DataGridView1.Rows.Count

        If ui = 1 Then

            DataGridView1.Rows(0).Cells(2).Value = DateTime.Now.ToString("yyyy/MM/dd") & " " & DateTime.Now.ToString("hh:mm:ss")

        Else

            DataGridView1.Rows(ui - 1).Cells(2).Value = DateTime.Now.ToString("yyyy/MM/dd") & " " & DateTime.Now.ToString("hh:mm:ss")

        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila As Integer

        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CODIGO" Then
            Dim fg As New vpartida
            Dim fg1 As New fpartida
            Select Case DataGridView1.Rows(fila).Cells(0).Value.Length
                Case "1" : DataGridView1.Rows(fila).Cells(0).Value = "0000" & "" & DataGridView1.Rows(fila).Cells(0).Value
                Case "2" : DataGridView1.Rows(fila).Cells(0).Value = "000" & "" & DataGridView1.Rows(fila).Cells(0).Value
                Case "3" : DataGridView1.Rows(fila).Cells(0).Value = "00" & "" & DataGridView1.Rows(fila).Cells(0).Value
                Case "4" : DataGridView1.Rows(fila).Cells(0).Value = "0" & "" & DataGridView1.Rows(fila).Cells(0).Value
                Case "5" : DataGridView1.Rows(fila).Cells(0).Value = DataGridView1.Rows(fila).Cells(0).Value
            End Select
            Select Case DataGridView1.Rows(fila).Cells(0).Value.Length
                Case "1" : fg.gcodigo = "0000" & "" & DataGridView1.Rows(fila).Cells(0).Value
                Case "2" : fg.gcodigo = "000" & "" & DataGridView1.Rows(fila).Cells(0).Value
                Case "3" : fg.gcodigo = "00" & "" & DataGridView1.Rows(fila).Cells(0).Value
                Case "4" : fg.gcodigo = "0" & "" & DataGridView1.Rows(fila).Cells(0).Value
                Case "5" : fg.gcodigo = DataGridView1.Rows(fila).Cells(0).Value
            End Select


            DataGridView1.Rows(fila).Cells(1).Value = fg1.buscar_persona(fg)


        End If
    End Sub

    Private Sub Conera_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim func As New fconera
        Me.TextBox12.Text = func.correlativo_conera + 1
        Select Case TextBox12.Text.Length
            Case "1" : TextBox12.Text = "000000000" & "" & TextBox12.Text
            Case "2" : TextBox12.Text = "00000000" & "" & TextBox12.Text
            Case "3" : TextBox12.Text = "0000000" & "" & TextBox12.Text
            Case "4" : TextBox12.Text = "000000" & "" & TextBox12.Text
            Case "5" : TextBox12.Text = "00000" & "" & TextBox12.Text
            Case "6" : TextBox12.Text = "0000" & "" & TextBox12.Text
            Case "7" : TextBox12.Text = "000" & "" & TextBox12.Text
            Case "8" : TextBox12.Text = "00" & "" & TextBox12.Text
            Case "9" : TextBox12.Text = "0" & "" & TextBox12.Text
            Case "10" : TextBox12.Text = TextBox12.Text
        End Select
        TextBox10.Select()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        abrir()
        enunciado2 = New SqlCommand("select abridora from programa_tenido_detalle where urdido ='" + ComboBox1.Text + "'", conx)
        respuesta2 = enunciado2.ExecuteReader
        While respuesta2.Read
            TextBox1.Text = respuesta2.Item("abridora")
        End While
        respuesta2.Close()

        enunciado4 = New SqlCommand("select tenido from programa_tenido_detalle where urdido ='" + Trim(ComboBox1.Text) + "'", conx)
        respuesta4 = enunciado4.ExecuteReader
        While respuesta4.Read
            TextBox11.Text = respuesta4.Item("tenido")
        End While
        respuesta4.Close()
        enunciado5 = New SqlCommand("select articulo from programa_tenido_detalle where urdido ='" + ComboBox1.Text + "'", conx)
        respuesta5 = enunciado5.ExecuteReader
        While respuesta5.Read
            TextBox2.Text = respuesta5.Item("articulo")
        End While
        respuesta5.Close()

        enunciado6 = New SqlCommand("select titulo from programa_tenido_detalle where urdido ='" + Trim(ComboBox1.Text) + "'", conx)
        respuesta6 = enunciado6.ExecuteReader
        While respuesta6.Read
            TextBox3.Text = respuesta6.Item("titulo")
        End While
        respuesta6.Close()

        enunciado7 = New SqlCommand("select hilo from programa_tenido_detalle where urdido ='" + Trim(ComboBox1.Text) + "'", conx)
        respuesta7 = enunciado7.ExecuteReader
        While respuesta7.Read
            TextBox5.Text = respuesta7.Item("hilo")
        End While
        respuesta7.Close()
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox10.Text.Length
                Case "1" : TextBox10.Text = "000000000" & "" & TextBox10.Text
                Case "2" : TextBox10.Text = "00000000" & "" & TextBox10.Text
                Case "3" : TextBox10.Text = "0000000" & "" & TextBox10.Text
                Case "4" : TextBox10.Text = "000000" & "" & TextBox10.Text
                Case "5" : TextBox10.Text = "00000" & "" & TextBox10.Text
                Case "6" : TextBox10.Text = "0000" & "" & TextBox10.Text
                Case "7" : TextBox10.Text = "000" & "" & TextBox10.Text
                Case "8" : TextBox10.Text = "00" & "" & TextBox10.Text
                Case "9" : TextBox10.Text = "0" & "" & TextBox10.Text
                Case "10" : TextBox10.Text = TextBox10.Text
            End Select

            abrir()
            llenar_combo_box2(ComboBox1, TextBox10.Text)

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim jj As New vconera
        Dim jj1 As New vconera
        Dim jj2 As New vconera
        Dim ll As New fconera
        Dim j As Integer
        'eliminar abridora
        jj.gid_conera = TextBox12.Text

        ll.eliminar_conera(jj)

        'insertar cabecera
        jj1.gid_conera = TextBox12.Text
        jj1.gjuego_urdido = ComboBox1.Text
        jj1.gjuego_abridora = TextBox1.Text
        jj1.gjuego_tenido = TextBox11.Text
        jj1.gjuego_conera = TextBox4.Text
        jj1.garticulo = TextBox2.Text
        jj1.gtitulo = TextBox3.Text
        jj1.ghilo = TextBox5.Text
        jj1.glongitud = TextBox7.Text
        jj1.gconos = TextBox8.Text
        jj1.gcuerda = TextBox9.Text
        jj1.ghinicio = DateTimePicker1.Value
        jj1.ghfinal = DateTimePicker2.Value

        j = DataGridView1.Rows.Count
        If ll.insertar_abridora_cabecera(jj1) = True Then
            For i = 0 To j - 1
                jj2.gid_conerad = TextBox12.Text
                jj2.gcodigod = DataGridView1.Rows(i).Cells(0).Value
                jj2.goperariod = DataGridView1.Rows(i).Cells(1).Value
                jj2.gfechad = DataGridView1.Rows(i).Cells(2).Value
                jj2.gminiciod = DataGridView1.Rows(i).Cells(3).Value
                jj2.gmfinald = DataGridView1.Rows(i).Cells(4).Value
                jj2.gconerad = DataGridView1.Rows(i).Cells(5).Value
                ll.insertar_abridora_detalle(jj2)
            Next
            MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            TextBox7.Text = ""
            TextBox1.Text = ""
            TextBox11.Text = ""
            TextBox4.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            DataGridView1.Rows.Clear()
            Dim func As New fconera
            Me.TextBox12.Text = func.correlativo_conera + 1
            Select Case TextBox12.Text.Length
                Case "1" : TextBox12.Text = "000000000" & "" & TextBox12.Text
                Case "2" : TextBox12.Text = "00000000" & "" & TextBox12.Text
                Case "3" : TextBox12.Text = "0000000" & "" & TextBox12.Text
                Case "4" : TextBox12.Text = "000000" & "" & TextBox12.Text
                Case "5" : TextBox12.Text = "00000" & "" & TextBox12.Text
                Case "6" : TextBox12.Text = "0000" & "" & TextBox12.Text
                Case "7" : TextBox12.Text = "000" & "" & TextBox12.Text
                Case "8" : TextBox12.Text = "00" & "" & TextBox12.Text
                Case "9" : TextBox12.Text = "0" & "" & TextBox12.Text
                Case "10" : TextBox12.Text = TextBox12.Text
            End Select
            TextBox10.Select()
        End If
    End Sub
End Class