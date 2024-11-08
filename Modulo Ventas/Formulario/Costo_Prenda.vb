Imports System.Data.SqlClient
Public Class Costo_Prenda
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
    Private Sub Costo_Prenda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Trim(TextBox2.Text) = "01" Then
            TextBox1.Text = "VIANNY"
        Else
            TextBox1.Text = "GRAUS"
        End If
        TextBox3.Enabled = False
        ComboBox1.Enabled = False
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox3.Text).Length
                Case "1" : TextBox3.Text = "01" & "0000000" & TextBox3.Text
                Case "2" : TextBox3.Text = "01" & "000000" & TextBox3.Text
                Case "3" : TextBox3.Text = "01" & "00000" & TextBox3.Text
                Case "4" : TextBox3.Text = "01" & "0000" & TextBox3.Text
                Case "5" : TextBox3.Text = "01" & "000" & TextBox3.Text
                Case "6" : TextBox3.Text = "01" & "00" & TextBox3.Text
                Case "7" : TextBox3.Text = "01" & "0" & TextBox3.Text
                Case "8" : TextBox3.Text = "01" & TextBox3.Text
                Case "9" : TextBox3.Text = "0" & TextBox3.Text
            End Select
            Button1.Select()
        End If
    End Sub
    Dim da, da1 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim DT, DT1 As String
        DataGridView1.DataSource = ""
        da.Clear()
        da1.Clear()
        abrir()

        If CheckBox1.Checked = True Then
            Dim cmd As New SqlDataAdapter("exec costo_pt '" + TextBox3.Text + "','" + TextBox2.Text + "'", conx)
            cmd.Fill(da)

            DataGridView1.DataSource = da
            Dim ex As New Exportar3
            ex.llenarExcel(DataGridView1)
        Else
            If CheckBox2.Checked = True Then
                Dim mes As String
                Select Case ComboBox1.Text
                    Case "ENERO" : mes = "01"
                    Case "FEBRERO" : mes = "02"
                    Case "MARZO" : mes = "03"
                    Case "ABRIL" : mes = "04"
                    Case "MAYO" : mes = "05"
                    Case "JUNIO" : mes = "06"
                    Case "JULIO" : mes = "07"
                    Case "AGOSTO" : mes = "08"
                    Case "SETIEMBRE" : mes = "09"
                    Case "OCTUBRE" : mes = "10"
                    Case "NOVIEMBRE" : mes = "11"
                    Case "DICIEMBRE" : mes = "12"
                End Select
                MsgBox(mes)
                MsgBox(TextBox2.Text)
                MsgBox(TextBox4.Text)
                Dim cmd As New SqlDataAdapter("exec Costos_Pt_Mes '" + mes + "','" + TextBox2.Text + "','" + TextBox4.Text + "'", conx)
                cmd.Fill(da1)

                DataGridView1.DataSource = da1
                Dim ex As New Exportar
                ex.llenarExcel(DataGridView1)
            End If
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox3.Enabled = True
            CheckBox2.Checked = False
        Else
            TextBox3.Enabled = False
            ComboBox1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ComboBox1.Enabled = True
            CheckBox1.Checked = False
        Else
            ComboBox1.Enabled = True
            TextBox3.Enabled = False
        End If
    End Sub
End Class