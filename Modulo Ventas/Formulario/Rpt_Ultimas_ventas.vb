Imports System.Net.Mail
Imports System.Data.SqlClient
Public Class Rpt_Ultimas_ventas
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado4 As SqlCommand
    Public respuesta4 As SqlDataReader
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Sub llenar_combo_box2()
        Try

            conn = New SqlDataAdapter("SELECT codigo_ven,alias_ven FROM  custom_vianny.dbo.Vendedores WHERE  rpm_ven='2'  and admin_ven in (0,2)", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "alias_ven"
            ComboBox1.ValueMember = "codigo_ven"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub Rpt_Ultimas_ventas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box2()
        ComboBox2.SelectedIndex = 0
        TextBox5.Text = "0005"
    End Sub
    Dim da As New DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        da.Clear()
        DataGridView1.DataSource = Nothing
        abrir()
        Dim gmes As String
        Select Case ComboBox2.Text
            Case "ENERO" : gmes = "01"
            Case "FEBRERO" : gmes = "02"
            Case "MARZO" : gmes = "03"
            Case "ABRIL" : gmes = "04"
            Case "MAYO" : gmes = "05"
            Case "JUNIO" : gmes = "06"
            Case "JULIO" : gmes = "07"
            Case "AGOSTO" : gmes = "08"
            Case "SETIEMBRE" : gmes = "09"
            Case "OCTUBRE" : gmes = "10"
            Case "NOVIEMBRE" : gmes = "11"
            Case "DICIEMBRE" : gmes = "12"
            Case "TODOS" : gmes = "13"

        End Select

        If CheckBox1.Checked = True And CheckBox2.Checked = False Then
            Dim cmd6 As New SqlDataAdapter("exec rpt_ultimas_ventas '" + TextBox2.Text + "','" + gmes + "','" + TextBox3.Text + "','" + TextBox5.Text + "',null", conx)
            cmd6.Fill(da)
        Else

            If CheckBox1.Checked = True And CheckBox2.Checked = True Then
                Dim cmd6 As New SqlDataAdapter("exec rpt_ultimas_ventas '" + TextBox2.Text + "','" + gmes + "','" + TextBox3.Text + "','" + TextBox5.Text + "','" + Trim(TextBox4.Text) + "'", conx)
                cmd6.Fill(da)
            Else
                If CheckBox1.Checked = False And CheckBox2.Checked = True Then
                    Dim cmd6 As New SqlDataAdapter("exec rpt_ultimas_ventas '" + TextBox2.Text + "','" + gmes + "','" + TextBox3.Text + "',null,'" + Trim(TextBox4.Text) + "'", conx)
                    cmd6.Fill(da)
                End If
            End If
        End If

        DataGridView1.DataSource = da
        DataGridView1.Columns(2).DefaultCellStyle.Format = "dd/MM/yyyy"
        Dim ext As New Exportar
        ext.llenarExcel(DataGridView1)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox5.Text = ComboBox1.SelectedValue.ToString
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Listar_Clientes.TextBox2.Text = 11
        If CheckBox1.Checked = True Then
            Listar_Clientes.TextBox3.Text = Trim(ComboBox1.Text)


        Else
            Listar_Clientes.TextBox3.Text = "A"

        End If

        Listar_Clientes.ShowDialog()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ComboBox1.Enabled = True
            TextBox4.Text = ""
            TextBox6.Text = ""
        Else
            ComboBox1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox4.Enabled = True
            Button1.Enabled = True
        Else
            TextBox4.Enabled = False
            Button1.Enabled = False
        End If
    End Sub
End Class