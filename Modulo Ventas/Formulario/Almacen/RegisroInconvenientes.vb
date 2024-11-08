Imports System.Data.SqlClient
Public Class RegisroInconvenientes
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Public comando As SqlCommand
    Dim da As New DataTable
    Dim Rsr2, Rsr3 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub registrar_informacion()
        Dim cmd20 As New SqlCommand("insert into AreaProblema (AreApb,SolApb,ProApb,RucApb,MotApb,RubApb,FecApb,UsuApb,OpApb,TdmApb) values (@AreApb,@SolApb,@ProApb,@RucApb,@MotApb,@RubApb,@FecApb,@UsuApb,@OpApb,@TdmApb)", conx)
        cmd20.Parameters.AddWithValue("@AreApb", Trim(ComboBox1.Text.ToString()))
        cmd20.Parameters.AddWithValue("@SolApb", Trim(ComboBox2.Text.ToString()))
        cmd20.Parameters.AddWithValue("@ProApb", Trim(TextBox2.Text.ToString()))
        cmd20.Parameters.AddWithValue("@RucApb", Trim(TextBox1.Text.ToString()))
        cmd20.Parameters.AddWithValue("@MotApb", Trim(TextBox3.Text.ToString()))
        cmd20.Parameters.AddWithValue("@RubApb", TextBox6.Text.ToString().Trim())
        cmd20.Parameters.AddWithValue("@FecApb", DateTimePicker1.Value)
        cmd20.Parameters.AddWithValue("@UsuApb", Trim(ComboBox5.Text.ToString()))
        cmd20.Parameters.AddWithValue("@OpApb", Trim(TextBox8.Text.ToString()))
        cmd20.Parameters.AddWithValue("@TdmApb", Convert.ToInt32(TextBox9.Text))
        cmd20.ExecuteNonQuery()
        MsgBox("Se Registro la Informacion Correctamnete")
    End Sub
    Sub actualizarRegistros()
        Dim cmd20 As New SqlCommand("update AreaProblema set AreApb=@AreApb and SolApb=@SolApb and ProApb=@ProApb and RucApb =@RucApb and MotApb =@MotApb and RubApb=@RubApb and FecApb=@FecApb and UsuApb=@UsuApb and OpApb=@OpApb and TdmApb=@TdmApb where IdApb =@IdApb", conx)
        cmd20.Parameters.AddWithValue("@AreApb", Trim(ComboBox1.Text.ToString()))
        cmd20.Parameters.AddWithValue("@SolApb", Trim(ComboBox2.Text.ToString()))
        cmd20.Parameters.AddWithValue("@ProApb", Trim(TextBox2.Text.ToString()))
        cmd20.Parameters.AddWithValue("@RucApb", Trim(TextBox1.Text.ToString()))
        cmd20.Parameters.AddWithValue("@MotApb", Trim(TextBox3.Text.ToString()))
        cmd20.Parameters.AddWithValue("@RubApb", TextBox6.Text.ToString().Trim())
        cmd20.Parameters.AddWithValue("@FecApb", DateTimePicker1.Value)
        cmd20.Parameters.AddWithValue("@UsuApb", Trim(ComboBox5.Text.ToString()))
        cmd20.Parameters.AddWithValue("@IdApb", Trim(TextBox7.Text.ToString()))
        cmd20.Parameters.AddWithValue("@OpApb", Trim(TextBox8.Text.ToString()))
        If TextBox9.Text.ToString().Trim().Length = 0 Then
            cmd20.Parameters.AddWithValue("@TdmApb", 0)
        Else
            cmd20.Parameters.AddWithValue("@TdmApb", Convert.ToInt32(TextBox9.Text))
        End If

        cmd20.ExecuteNonQuery()
        MsgBox("Se Actualizo la Informacion Correctamnete")
    End Sub
    Sub Listar_Informacion()
        da.Clear()
        abrir()
        Dim cmd As New SqlDataAdapter("select IdApb As Id,AreApb as Area,SolApb as Solicitado,ProApb as Proveedor,RucApb,MotApb,m.DesMpb as Motivo,RubApb,r.DesRpb as Rubro,FecApb as Fecha,UsuApb as Usuario from AreaProblema a
        left join MotivosProblemas m on a.MotApb = m.CodMpb 
        left join RubroProblemas r on a.RubApb = r.CodRpb", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).Width = 180
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Width = 192
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(8).Width = 192
            DataGridView1.Columns(9).Width = 80
            DataGridView1.Columns(10).Width = 80
        Else
            DataGridView1.DataSource = Nothing
        End If

    End Sub

    Private Sub RegisroInconvenientes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        da.Clear()
        Listar_Informacion()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick

    End Sub
    Private Function validar()
        If TextBox3.Text.ToString().Trim().Length = 0 Then
            MsgBox("Falta Ingresar el Motivo")
            Return False
        End If
        If TextBox6.Text.ToString().Trim().Length = 0 Then
            MsgBox("Falta Ingresar el Rubro")
            Return False
        End If
        If ComboBox1.Text.ToString().Trim().Length = 0 Then
            MsgBox("Falta Ingresar el Area")
            Return False
        End If
        Return True
    End Function

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown

        If e.KeyCode = Keys.F1 Then
            TablaIncidencias.Label2.Text = "1"
            TablaIncidencias.Label1.Text = "MOTIVO"
            TablaIncidencias.ShowDialog()
        End If
    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.F1 Then
            TablaIncidencias.Label2.Text = "2"
            TablaIncidencias.Label1.Text = "RUBRO"
            TablaIncidencias.ShowDialog()
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "22"
            Clientes.Show()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        Select Case Trim(TextBox8.Text).Length
            Case "1" : TextBox8.Text = "01" & "0000000" & TextBox8.Text
            Case "2" : TextBox8.Text = "01" & "000000" & TextBox8.Text
            Case "3" : TextBox8.Text = "01" & "00000" & TextBox8.Text
            Case "4" : TextBox8.Text = "01" & "0000" & TextBox8.Text
            Case "5" : TextBox8.Text = "01" & "000" & TextBox8.Text
            Case "6" : TextBox8.Text = "01" & "00" & TextBox8.Text
            Case "7" : TextBox8.Text = "01" & "0" & TextBox8.Text
            Case "8" : TextBox8.Text = "01" & TextBox8.Text
            Case "9" : TextBox8.Text = "0" & TextBox8.Text
            Case "10" : TextBox8.Text = TextBox8.Text

        End Select
        If TextBox7.Text.ToString().Trim().Length = 0 Then
            If validar() Then
                registrar_informacion()
            End If

        Else
            actualizarRegistros()
        End If
        LIMPIAR()
        Listar_Informacion()
    End Sub
    Sub LIMPIAR()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
    End Sub
End Class