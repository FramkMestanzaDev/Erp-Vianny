Imports System.Data.SqlClient
Public Class ProddiariaTej
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim ty As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub llenar_combo_box(ByVal cb As ComboBox)
        Try
            conn = New SqlDataAdapter("SELECT nombre_mq AS NOMBRE, sigla_mq AS SIGLA FROM custom_vianny.DBO.maquinas", conx)
            conn.Fill(ty)
            ComboBox1.DataSource = ty
            ComboBox1.DisplayMember = "NOMBRE"
            ComboBox1.ValueMember = "SIGLA"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Rpt_Rollos.TextBox1.Text = "01"
        Dim fecini, fecfin As String
        fecini = Format(DateTimePicker1.Value, "yyyy-MM-dd")
        fecfin = Format(DateTimePicker2.Value, "yyyy-MM-dd")
        Rpt_Rollos.TextBox2.Text = (Replace(fecini, "-", ""))
        Rpt_Rollos.TextBox3.Text = (Replace(fecfin, "-", ""))
        If CheckBox4.Checked = True Then
            Rpt_Rollos.TextBox4.Text = TextBox2.Text
        Else
            Rpt_Rollos.TextBox4.Text = ""
        End If

        If CheckBox3.Checked = True Then
            Rpt_Rollos.TextBox5.Text = ComboBox1.Text
        Else
            Rpt_Rollos.TextBox5.Text = ""
        End If

        If CheckBox1.Checked = True Then
            Rpt_Rollos.TextBox6.Text = TextBox1.Text
        Else
            Rpt_Rollos.TextBox6.Text = ""
        End If
        Rpt_Rollos.Show()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox1.Enabled = True

        Else
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then

            TextBox2.Enabled = True
        Else
            TextBox2.Enabled = False
        End If
    End Sub
    Dim GH As New DataTable
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim JH As New vtejeduria
        Dim LL As New ftejeduria
        DataGridView1.DataSource = ""
        JH.gccia = "01"
        JH.gfechaini = DateTimePicker1.Value
        JH.gfechafin = DateTimePicker2.Value
        'If CheckBox4.Checked = True Then
        '    JH.gtejedor = TextBox2.Text
        'Else
        '    JH.gtejedor = ""
        'End If

        'If CheckBox3.Checked = True Then
        '    JH.gmaquina = ComboBox1.Text
        'Else
        '    JH.gmaquina = ""
        'End If

        'If CheckBox1.Checked = True Then
        '    JH.gpedidoot = TextBox1.Text
        'Else
        '    JH.gpedidoot = ""
        'End If

        GH = LL.reporte_tejeduria2(JH)
        DataGridView1.DataSource = GH
        Dim k As New Exportar
        k.llenarExcel(DataGridView1)
    End Sub

    Private Sub ProddiariaTej_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box(ComboBox1)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.F1 Then
            Det_Rollo.Label1.Text = "TRABAJADOR"
            Det_Rollo.Label2.Text = 5
            Det_Rollo.Show()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Det_Rollo.Label1.Text = "PO"
            Det_Rollo.Label2.Text = 6
            Det_Rollo.Show()

        End If
    End Sub
End Class