Imports System.Data.SqlClient
Public Class Rpt_factura_oc
    Dim dt As DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
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
    Private Sub Rpt_factura_oc_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fg As New vfactura
        Dim fg1 As New ffactura
        Dim oo As Integer
        Dim JJ As String
        oo = DataGridView1.Rows.Count
        For i = 0 To oo - 1
            fg.gdoc = Trim(DataGridView1.Rows(i).Cells(6).Value)
            fg.gndoc = Trim(DataGridView1.Rows(i).Cells(4).Value) + "-" + Trim(DataGridView1.Rows(i).Cells(5).Value)
            fg.gCLIENTE = Trim(DataGridView1.Rows(i).Cells(2).Value)
            fg.gperiodo = Trim(TextBox2.Text)
            fg.gccia = Trim(Label3.Text)

            JJ = fg1.mostrar_estado_facbol3(fg)
            If JJ = "False" Then
                DataGridView1.Rows(i).Cells(0).Value = fg1.mostrar_estado_facbol(fg)
            Else
                DataGridView1.Rows(i).Cells(0).Value = fg1.mostrar_estado_facbol3(fg)
            End If
        Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim jk As New Exportar
        jk.llenarExcel(DataGridView1)
    End Sub
    Dim bs As New BindingSource()
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox3.Text = ""
        abrir()
        da.Clear()
        DataGridView1.DataSource = Nothing
        Dim ja As Double
        Dim cmd As New SqlDataAdapter("exec FACTURA_OCOMPRA NULL,NULL,'" + Trim(Label3.Text) + "','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "','" + Replace(DateTimePicker2.Value.ToString("yyyy-MM-dd"), "-", "") + "'", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            bs.DataSource = da
            DataGridView1.DataSource = bs
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(1).Width = 85
            DataGridView1.Columns(2).Width = 160
            DataGridView1.Columns(3).Width = 85
            DataGridView1.Columns(4).Width = 65
            DataGridView1.Columns(5).Width = 85
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Width = 70
            DataGridView1.Columns(8).Width = 70
            DataGridView1.Columns(9).Width = 70
            DataGridView1.Columns(10).Width = 70
            DataGridView1.Columns(11).Width = 70
            DataGridView1.Columns(12).Width = 70

            Dim fg As New vfactura
            Dim fg1 As New ffactura
            Dim oo As Integer
            Dim JJ As String
            oo = DataGridView1.Rows.Count
            ja = 0
            For i = 0 To oo - 1
                fg.gdoc = Trim(DataGridView1.Rows(i).Cells(6).Value)
                fg.gndoc = Trim(DataGridView1.Rows(i).Cells(4).Value) + "-" + Trim(DataGridView1.Rows(i).Cells(5).Value)
                fg.gCLIENTE = Trim(DataGridView1.Rows(i).Cells(2).Value)
                fg.gperiodo = Trim(TextBox2.Text)
                fg.gccia = Trim(Label3.Text)

                JJ = fg1.mostrar_estado_facbol3(fg)
                If JJ = "False" Then
                    DataGridView1.Rows(i).Cells(0).Value = fg1.mostrar_estado_facbol(fg)
                Else
                    DataGridView1.Rows(i).Cells(0).Value = fg1.mostrar_estado_facbol3(fg)
                End If

                If Trim(DataGridView1.Rows(i).Cells(0).Value) = "CANCELADO" Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White

                Else
                    If Trim(DataGridView1.Rows(i).Cells(0).Value) = "CANJEADO" Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                        DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                    End If
                End If
                ja = ja + DataGridView1.Rows(i).Cells(12).Value
            Next
            TextBox5.Text = ja.ToString("N2")
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        bs.Filter = String.Format("CLIENTE LIKE '%{0}%'", TextBox3.Text)
        Dim fg As New vfactura
        Dim fg1 As New ffactura
        Dim oo As Integer
        Dim JJ As String
        Dim ja1 As Double
        oo = DataGridView1.Rows.Count
        For i = 0 To oo - 1
            fg.gdoc = Trim(DataGridView1.Rows(i).Cells(6).Value)
            fg.gndoc = Trim(DataGridView1.Rows(i).Cells(4).Value) + "-" + Trim(DataGridView1.Rows(i).Cells(5).Value)
            fg.gCLIENTE = Trim(DataGridView1.Rows(i).Cells(2).Value)
            fg.gperiodo = Trim(TextBox2.Text)
            fg.gccia = Trim(Label3.Text)

            JJ = fg1.mostrar_estado_facbol3(fg)
            If JJ = "False" Then
                DataGridView1.Rows(i).Cells(0).Value = fg1.mostrar_estado_facbol(fg)
            Else
                DataGridView1.Rows(i).Cells(0).Value = fg1.mostrar_estado_facbol3(fg)
            End If
            If Trim(DataGridView1.Rows(i).Cells(0).Value) = "CANCELADO" Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.White

            Else
                If Trim(DataGridView1.Rows(i).Cells(0).Value) = "CANJEADO" Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                End If
            End If
            ja1 = ja1 + DataGridView1.Rows(i).Cells(12).Value
            TextBox5.Text = ja1.ToString("N2")
        Next
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            Dim oo As Integer
            Dim ja1 As Double
            oo = DataGridView1.Rows.Count

            For i = 0 To oo - 1
                If Me.DataGridView1.Rows(i).Cells(0).Value = "CANCELADO" Or Me.DataGridView1.Rows(i).Cells(0).Value = "CANJEADO" Then
                    Dim pos As Integer = i
                    Me.DataGridView1.CurrentCell = Nothing
                    Me.DataGridView1.Rows(pos).Visible = False

                Else
                    ja1 = ja1 + DataGridView1.Rows(i).Cells(12).Value
                End If

            Next

            TextBox5.Text = ja1.ToString("N2")
        Else
            DataGridView1.DataSource = Nothing
            TextBox3.Text = ""
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then
            MsgBox("Esta Funcion esta Bloqueada")
            For Each col As DataGridViewColumn In DataGridView1.Columns

                col.SortMode = DataGridViewColumnSortMode.NotSortable

            Next
        End If

    End Sub
End Class