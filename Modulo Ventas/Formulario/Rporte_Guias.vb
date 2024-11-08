
Imports System.Data.SqlClient
Public Class Rporte_Guias
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty As New DataTable
    Dim da As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub llenar_combo_box()
        Try

            conn = New SqlDataAdapter("select rtrim(ltrim(motiv_17f)) AS motiv_17f,rtrim(ltrim(nomb_17f)) AS nomb_17f from custom_vianny.dbo.Fag1700  where ccia_17F ='" + Trim(Label4.Text) + "'", conx)
            conn.Fill(ty)
            ComboBox1.DataSource = ty
            ComboBox1.DisplayMember = "nomb_17f"
            ComboBox1.ValueMember = "motiv_17f"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Rporte_Guias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.DataSource = ""
        da.Clear()
        abrir()
        Dim ruc, serie, motivo As String
        If CheckBox1.Checked = False Then
            ruc = "NULL"
        Else
            ruc = Trim(TextBox1.Text)
        End If
        If CheckBox2.Checked = False Then
            serie = "NULL"
        Else
            serie = Trim(TextBox2.Text)
        End If
        If CheckBox3.Checked = True Then
            motivo = ComboBox1.SelectedValue.ToString
        Else
            motivo = "NULL"
        End If
        Dim cmd As New SqlDataAdapter("EXEC CaeSoft_ReporteListadoGuiasRemision '" + Trim(Label4.Text) + "', 1, '" + DateTimePicker1.Value + "',  '" + DateTimePicker2.Value + "', NULL, NULL, NULL, NULL, " + ruc + ", " + serie + ", " + motivo + ", 1, 1, 1", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        Dim ex As New Exportar
        ex.llenarExcel(DataGridView1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Rppt_guiar.TextBox1.Text = Label4.Text
        Rppt_guiar.TextBox2.Text = 1
        'Rppt_guiar.TextBox3.Text = Replace(DateTimePicker1.Value.ToString("dd-MM-yyyy"), "-", "")
        'Rppt_guiar.TextBox4.Text = Replace(DateTimePicker2.Value.ToString("dd-MM-yyyy"), "-", "")
        Rppt_guiar.TextBox3.Text = DateTimePicker1.Value
        Rppt_guiar.TextBox4.Text = DateTimePicker2.Value
        Rppt_guiar.TextBox5.Text = "0"
        Rppt_guiar.TextBox6.Text = "0"
        Rppt_guiar.TextBox7.Text = "0"
        Rppt_guiar.TextBox8.Text = "0"
        If CheckBox1.Checked = False Then
            Rppt_guiar.TextBox9.Text = "0"
        Else
            Rppt_guiar.TextBox9.Text = Trim(TextBox1.Text)
        End If
        If CheckBox2.Checked = False Then
            Rppt_guiar.TextBox10.Text = "0"
        Else
            Rppt_guiar.TextBox10.Text = Trim(TextBox2.Text)
        End If
        If CheckBox3.Checked = True Then
            Rppt_guiar.TextBox11.Text = ComboBox1.SelectedValue.ToString
        Else
            Rppt_guiar.TextBox11.Text = "0"
        End If

        Rppt_guiar.TextBox12.Text = 1
        Rppt_guiar.TextBox13.Text = 1
        Rppt_guiar.TextBox14.Text = 1

        Rppt_guiar.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.Enabled = True
            TextBox1.Select()
        Else
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox2.Enabled = True
            TextBox2.Select()
        Else
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            ComboBox1.Enabled = True
            ComboBox1.Select()
        Else
            ComboBox1.Enabled = False
        End If
    End Sub
End Class