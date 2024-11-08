Imports System.Data.SqlClient
Public Class Seguimiento_Actividades
    Dim dt As New DataTable
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader

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
    Dim Rsr2 As SqlDataReader
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim jj As New fgestionact
        Dim hj As New vgestionact
        hj.gvededor_co = ComboBox2.Text
        abrir()
        Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + ComboBox2.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            hj.gvededor_co = Rsr2(0)
        End If
        Rsr2.Close()
        'Select Case ComboBox2.Text
        '    Case "VPIZARRO" : hj.gvededor_co = "0027"
        '    Case "VSILVERIO" : hj.gvededor_co = "0005"
        '    Case "GBALVIN" : hj.gvededor_co = "0028"
        '    Case "GBEDON" : hj.gvededor_co = "0010"
        '    Case "VINCIO" : hj.gvededor_co = "0022"
        '    Case "DBRAVO" : hj.gvededor_co = "0023"
        '    Case "AMENDO" : hj.gvededor_co = "0026"
        '    Case "GCUEVA" : hj.gvededor_co = "0029"
        '    Case "JSALINAS" : hj.gvededor_co = "0025"
        '    Case "VGRAUS" : hj.gvededor_co = "0007"
        '    Case "WSALINAS" : hj.gvededor_co = "0034"
        'End Select
        hj.gfini = DateTimePicker1.Text
        hj.gffin = DateTimePicker2.Text
        If ComboBox1.Text = "FECH. GESTION" Then
            dt = jj.mostrar_gestion_actividades(hj)
        Else
            If ComboBox1.Text = "FECH. PROX. VISITA" Then
                dt = jj.mostrar_gestion_actividades2(hj)
            End If
        End If


        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(5).Width = 180
            DataGridView1.Columns(7).Width = 180
        Else
            MsgBox("NO EXISTE INFORMACION PARA MOSTRAR")
        End If
    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseHover
        ToolTip1.SetToolTip(PictureBox1, "BUSCAR")
        ToolTip1.ToolTipTitle = "SEGUIMIENTO ACTIVIDADES"
        ToolTip1.ToolTipIcon = ToolTipIcon.Info
    End Sub

    Private Sub Seguimiento_Actividades_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim ext As New Exportar
        ext.llenarExcel(DataGridView1)
    End Sub
End Class