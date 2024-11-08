Imports System.Data.SqlClient
Public Class REQUERIMIENTO
    Public conx As SqlConnection
    Public conn As SqlDataAdapter

    Public comando As SqlCommand
    Dim Rsr2 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim da As New DataTable
    Private Sub REQUERIMIENTO_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("EXEC custom_vianny.DBO.CaeSoft_RequerimientoProgramadoPedidoCrudoDetallado '01' , '" + Trim(Label1.Text) + "'", conx)
        cmd.Fill(da)

        If da.Rows.Count <> 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(2).Width = 150
            DataGridView1.Columns(3).Width = 350
            DataGridView1.Columns(6).Width = 75
            DataGridView1.Columns(7).Width = 75
            DataGridView1.Columns(8).Width = 75
        Else
            DataGridView1.DataSource = ""
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'Dim I As Integer
        'I = CREAR_PARTIDAS.DataGridView1.Rows.Count
        'If I > 0 Then
        CREAR_PARTIDAS.DataGridView1.Rows.Clear()
        'End If
        CREAR_PARTIDAS.DataGridView1.Rows.Add()
        CREAR_PARTIDAS.TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(6).Value
        CREAR_PARTIDAS.TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value
        CREAR_PARTIDAS.DataGridView1.Rows(0).Cells(0).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        CREAR_PARTIDAS.DataGridView1.Rows(0).Cells(1).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        CREAR_PARTIDAS.DataGridView1.Rows(0).Cells(2).Value = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        abrir()
        Dim sql102 As String = "select fich_3,nomb_10 from custom_vianny.dbo.qag0300 q left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3 = c.fich_10 where ncom_3 ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        Dim i5 As Integer
        i5 = 0
        If Rsr2.Read() Then
            CREAR_PARTIDAS.TextBox5.Text = Rsr2(0)
            CREAR_PARTIDAS.TextBox6.Text = Rsr2(1)
            If Trim(Rsr2(0)) = "20508740361" Then
                CREAR_PARTIDAS.ComboBox3.SelectedIndex = 0
            Else
                CREAR_PARTIDAS.ComboBox3.SelectedIndex = 1
            End If
        Else
            CREAR_PARTIDAS.TextBox5.Text = ""
            CREAR_PARTIDAS.TextBox6.Text = ""
        End If
        CREAR_PARTIDAS.DataGridView1.Enabled = True
        CREAR_PARTIDAS.TextBox7.Enabled = True
        Rsr2.Close()

        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class