Imports System.Data.SqlClient
Public Class UPDATE_HOT
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conn As SqlDataAdapter
    Dim ty2 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(ComboBox2.Text) = "INGRESO" Then
            Dim cmd As New SqlCommand("update rama set hora =@hora where  partida =@partida", conx)
            cmd.Parameters.AddWithValue("@hora", Trim(TextBox1.Text))
            cmd.Parameters.AddWithValue("@partida", TextBox2.Text)
            cmd.ExecuteNonQuery()
        Else
            Dim cmd As New SqlCommand("update rama set hora2 =@hora where  partida =@partida", conx)
            cmd.Parameters.AddWithValue("@hora", TextBox1.Text)
            cmd.Parameters.AddWithValue("@partida", TextBox2.Text)
            cmd.ExecuteNonQuery()
        End If
        PASADO_RAMA.ACTUALIZAR()
        Close()
    End Sub
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    'Sub llenar_combo_box2(ByVal cb As ComboBox)
    '    Try

    '        conn = New SqlDataAdapter("select partida from rama where codigo ='" + PASADO_RAMA.TextBox1.Text + "' and flag =1 and pasado >0", conx)
    '        conn.Fill(ty2)
    '        ComboBox1.DataSource = ty2
    '        ComboBox1.DisplayMember = "partida"
    '        ComboBox1.ValueMember = "partida"
    '        'respuesta = enunciado.ExecuteReader
    '        'While respuesta.Read
    '        '    cb.Items.Add(respuesta.Item("nomb_17f"))
    '        'End While
    '        'respuesta.Close()
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub
    Public vad As Boolean
    Private Sub UPDATE_HOT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        'llenar_combo_box2(ComboBox1)
        TextBox1.Text = DateTime.Now.ToString("HH:mm")
        ComboBox2.SelectedIndex = 0
        vad = True
    End Sub

    Private Sub UPDATE_HOT_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        PASADO_RAMA.Label4.Text = 2
    End Sub
End Class