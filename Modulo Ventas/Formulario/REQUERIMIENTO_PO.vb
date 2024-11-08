Imports System.Data.SqlClient
Public Class REQUERIMIENTO_PO
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("FALTA INGRESAR LA PO")
        Else
            abrir()
            Dim cmd1 As New SqlDataAdapter("SELECT  '' AS ESTADO,ncom_3 AS  OP, descri_3 as PRODUCTO, ISNULL(B.nomb_10,'') AS Nomb_10,cants_3 AS CANT_SOLICITADA,cantp_3 AS CANT_PROGRAMADA,C.color_16b AS COLOR,merchan_3,linea_3 FROM custom_vianny.dbo.QAG0300 A LEFT JOIN custom_vianny.dbo.CAG1000 B  ON '01' = B.CCIA AND A.FICH_3 = B.FICH_10 LEFT JOIN custom_vianny.dbo.qag160b C	ON A.ncom_3 = C.ncom_16b AND A.ccia = C.ccia Where A.CCia='" + Label3.Text + "' AND A.Program_3 ='" + TextBox1.Text + "'", conx)
            Dim da1 As New DataTable
            cmd1.Fill(da1)

            If da1.Rows.Count <> 0 Then
                DataGridView1.DataSource = da1
                DataGridView1.Columns(1).Width = 85
                DataGridView1.Columns(2).Width = 200
                DataGridView1.Columns(3).Width = 200
                DataGridView1.Columns(5).Width = 120
                DataGridView1.Columns(6).Width = 120
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(8).Visible = False
            End If
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Det_Rollo.Label1.Text = "PO"
            Det_Rollo.Label2.Text = 4
            Det_Rollo.Show()

        End If
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Requerimiento_op.TextBox1.Text = Trim(DataGridView1.Rows(Label2.Text).Cells(1).Value)
        Requerimiento_op.TextBox2.Text = TextBox1.Text
        Requerimiento_op.TextBox3.Text = Trim(DataGridView1.Rows(Label2.Text).Cells(4).Value)
        Requerimiento_op.TextBox5.Text = Trim(DataGridView1.Rows(Label2.Text).Cells(2).Value)
        Requerimiento_op.TextBox6.Text = Trim(DataGridView1.Rows(Label2.Text).Cells(3).Value)

        Requerimiento_op.TextBox12.Text = Trim(DataGridView1.Rows(Label2.Text).Cells(5).Value)
        Select Case Trim(DataGridView1.Rows(Label2.Text).Cells(7).Value)
            Case "0013" : Requerimiento_op.TextBox4.Text = "GBEDON"
            Case "0021" : Requerimiento_op.TextBox4.Text = "VINCIO"
            Case "0023" : Requerimiento_op.TextBox4.Text = "DBRAVO"
            Case "0024" : Requerimiento_op.TextBox4.Text = "AMENDO"
            Case "0025" : Requerimiento_op.TextBox4.Text = "VGRAUS"
            Case "0011" : Requerimiento_op.TextBox4.Text = "VPIZARRO"
            Case "0003" : Requerimiento_op.TextBox4.Text = "GBALVIN"
            Case "0001" : Requerimiento_op.TextBox4.Text = "VSILVERIO"
            Case "0014" : Requerimiento_op.TextBox4.Text = "WSALINAS"
            Case "0028" : Requerimiento_op.TextBox4.Text = "ASILVA"
        End Select
        Requerimiento_op.Show()
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim ig As Integer

        ig = DataGridView1.Rows.Count
        If ig < 0 Then
            Label2.Text = 0
        Else
            Label2.Text = e.RowIndex
        End If

        If e.RowIndex = -1 Then
            Label2.Text = 0
        End If
    End Sub

    Private Sub REQUERIMIENTO_PO_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = 0
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Reporte__OP.TextBox1.Text = Label2.Text
        Reporte__OP.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox1_Layout(sender As Object, e As LayoutEventArgs) Handles TextBox1.Layout

    End Sub
End Class