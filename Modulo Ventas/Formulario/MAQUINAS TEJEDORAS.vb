Imports System.Data.SqlClient
Public Class MAQUINAS_TEJEDORAS
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da, ty, TY2 As New DataTable
    Public conn As SqlDataAdapter

    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub MAQUINAS_TEJEDORAS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MOSTRAR()
        Label11.Text = DataGridView1.Rows(0).Cells(0).Value
        Me.TabControl1.TabPages(1).Enabled = False
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Sub MOSTRAR()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT codigo_mq AS CODIGO,nombre_mq AS NOMBRE,T.dele AS TIPO,T1.dele AS MARCA, sigla_mq AS SIGLA,M.capmin_mq,M.capnom_mq,M.capmax_mq FROM custom_vianny.DBO.maquinas M INNER JOIN custom_vianny.DBO.TAB0100 T ON M.ccia_mq=T.ccia AND M.tipmaq_mq = T.cele AND T.CCia = '01' AND T.CTab = 'TIPMAQ' LEFT JOIN  custom_vianny.DBO.TAB0100 T1  ON M.ccia_mq=T1.ccia AND M.marca_mq = T1.cele AND T1.CCia = '01' AND T1.CTab = 'MARMAQ' ", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(8).Visible = False
        DataGridView1.Columns(9).Visible = False
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 170
        DataGridView1.Columns(4).Width = 170
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Label11.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        'Me.TabControl1.TabPages(2).Enabled = True
        If e.ColumnIndex = 0 Then
            Dim num1, num2 As Integer
            Dim da6 As New DataTable
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            det_maquina.Text = "GALGA"

            det_maquina.Label1.Visible = False
            det_maquina.ComboBox1.Visible = False
            det_maquina.Button1.Visible = False
            det_maquina.Button2.Visible = False
            det_maquina.Button3.Visible = False
            abrir()
            Dim cmd As New SqlDataAdapter("select galga from galga where tejedora ='" + DataGridView1.Rows(e.RowIndex).Cells(2).Value + "'", conx)
            cmd.Fill(da6)
            det_maquina.DataGridView1.DataSource = da6
            det_maquina.DataGridView1.Columns(0).Visible = False
            det_maquina.Show()
        End If
        If e.ColumnIndex = 1 Then
            Dim num1, num2 As Integer
            Dim da5 As New DataTable
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            det_maquina.Text = "ALIMENTADOR"
            det_maquina.Label1.Visible = False
            det_maquina.ComboBox1.Visible = False
            det_maquina.Button1.Visible = False
            det_maquina.Button2.Visible = False
            det_maquina.Button3.Visible = False
            abrir()
            Dim cmd As New SqlDataAdapter("select alimentador from alimentador where tejedora ='" + DataGridView1.Rows(e.RowIndex).Cells(2).Value + "'", conx)
            cmd.Fill(da5)
            det_maquina.DataGridView1.DataSource = da5
            det_maquina.DataGridView1.Columns(0).Visible = False
            det_maquina.Show()


        End If
    End Sub
    Function ELIMINAR(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim respuesta As DialogResult
        Dim i As Integer
        If Label2.Text = "" Then
            MsgBox("NO HA SELECCIONADO NINITEMS A ELIMINAR")
        Else
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
                abrir()
                Dim agregar As String = "delete from custom_vianny.DBO.maquinas where codigo_mq ='" + Trim(Label11.Text) + "'"
                If (Eliminar(agregar)) Then
                    MessageBox.Show("DATO ELIMINADO CORRECTAMENTE")
                    Dim cmd As New SqlDataAdapter("SELECT codigo_mq AS CODIGO,nombre_mq AS NOMBRE,T.dele AS TIPO,T1.dele AS MARCA, sigla_mq AS SIGLA,M.capmin_mq,M.capnom_mq,M.capmax_mq FROM custom_vianny.DBO.maquinas M INNER JOIN custom_vianny.DBO.TAB0100 T ON M.ccia_mq=T.ccia AND M.tipmaq_mq = T.cele AND T.CCia = '01' AND T.CTab = 'TIPMAQ' LEFT JOIN  custom_vianny.DBO.TAB0100 T1  ON M.ccia_mq=T1.ccia AND M.marca_mq = T1.cele AND T1.CCia = '01' AND T1.CTab = 'MARMAQ' ", conx)

                    cmd.Fill(da)
                    DataGridView1.Columns(7).Visible = False
                    DataGridView1.Columns(8).Visible = False
                    DataGridView1.Columns(9).Visible = False
                    DataGridView1.Columns(2).Width = 80
                    DataGridView1.Columns(3).Width = 170
                    DataGridView1.Columns(4).Width = 170
                    Label2.Text = 0
                End If
            End If
        End If
    End Sub
    Sub llenar_combo_box1(ByVal cb As ComboBox)
        Try

            conn = New SqlDataAdapter("SELECT cele,DELE FROM custom_vianny.DBO.TAB0100  Where CCia = '01' AND ctab = 'MARMAQ'", conx)
            conn.Fill(TY2)
            ComboBox2.DataSource = TY2
            ComboBox2.DisplayMember = "DELE"
            ComboBox2.ValueMember = "cele"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box(ByVal cb As ComboBox)
        Try

            conn = New SqlDataAdapter("SELECT CELE,DELE FROM custom_vianny.DBO.TAB0100 A  Where CCia = '01' AND CTab = 'TIPMAQ'", conx)
            conn.Fill(ty)
            ComboBox1.DataSource = ty
            ComboBox1.DisplayMember = "DELE"
            ComboBox1.ValueMember = "CELE"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        det_maquina.Text = "GALGA"
        det_maquina.ComboBox1.Items.Add("18")
        det_maquina.ComboBox1.Items.Add("20")
        det_maquina.ComboBox1.Items.Add("24")
        det_maquina.ComboBox1.Items.Add("28")
        det_maquina.ComboBox1.Items.Add("32")
        det_maquina.ComboBox1.Items.Add("40")
        det_maquina.ComboBox1.SelectedIndex = 0
        det_maquina.Label1.Text = "GALGA"
        det_maquina.Label2.Text = 1
        Dim t As Integer
        t = DataGridView4.Rows.Count
        If t > 0 And DataGridView2.Rows.Count = 0 Then
            det_maquina.DataGridView1.Rows.Add(t)
            For i = 0 To t - 1
                det_maquina.DataGridView1.Rows(i).Cells(0).Value = DataGridView4.Rows(i).Cells(0).Value
            Next
        Else
            If DataGridView2.Rows.Count = 0 Then

            Else
                det_maquina.DataGridView1.Rows.Add(DataGridView2.Rows.Count)
                For i = 0 To DataGridView2.Rows.Count - 1
                    det_maquina.DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(0).Value
                Next
            End If

        End If

        det_maquina.Show()

    End Sub
    Function insertar(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        det_maquina.Text = "ALIMENTADOR"
        det_maquina.ComboBox1.Items.Add("90")
        det_maquina.ComboBox1.Items.Add("96")
        det_maquina.ComboBox1.SelectedIndex = 0
        det_maquina.Label1.Text = "ALIMENTADOR"
        det_maquina.Label2.Text = 2
        Dim t As Integer
        t = DataGridView5.Rows.Count
        If t > 0 And DataGridView3.Rows.Count = 0 Then
            det_maquina.DataGridView1.Rows.Add(t)
            For i = 0 To t - 1
                det_maquina.DataGridView1.Rows(i).Cells(0).Value = DataGridView5.Rows(i).Cells(0).Value
            Next
        Else
            det_maquina.DataGridView1.Rows.Add(DataGridView3.Rows.Count)
            For i = 0 To DataGridView3.Rows.Count - 1
                det_maquina.DataGridView1.Rows(i).Cells(0).Value = DataGridView3.Rows(i).Cells(0).Value
            Next
        End If

        det_maquina.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("EL CAMPO DESCRIPCION O SIGLA ESTA VACIO")
        Else
            abrir()
            Dim cmd1 As New SqlDataAdapter("select *  from custom_vianny.dbo.maquinas where nombre_mq like '" + TextBox2.Text + "%'", conx)
            Dim da1 As New DataTable
            cmd1.Fill(da1)

            If da1.Rows.Count > 0 And Label12.Text = 1 Then
                MsgBox("LA MAQUINATEJEDORA YA ESTA REGISTRADO")
            Else
                Dim eliminar1 As String = "delete from custom_vianny.DBO.maquinas where codigo_mq ='" + Trim(TextBox1.Text) + "'"
                ELIMINAR(eliminar1)
                Dim agregar As String = "insert into custom_vianny.dbo.maquinas (ccia_mq,codigo_mq,tipmaq_mq,marca_mq,nombre_mq,sigla_mq,capmin_mq,capnom_mq,capmax_mq) values('" + "01" + "','" + TextBox1.Text + "','" + ComboBox1.SelectedValue.ToString + "','" + ComboBox2.SelectedValue.ToString + "','" + TextBox2.Text + "','" + TextBox3.Text + "','" + TextBox4.Text + "','" + TextBox5.Text + "','" + TextBox6.Text + "')"
                If (insertar(agregar)) Then

                    Dim k, k3 As Integer
                    k = DataGridView2.Rows.Count
                    k3 = DataGridView3.Rows.Count
                    Dim eliminar10 As String = "delete from galga where tejedora ='" + Trim(TextBox1.Text) + "'"
                    ELIMINAR(eliminar10)
                    For i1 = 0 To k - 1
                        Dim agregar1 As String = "insert into galga(galga,tejedora) values ('" + DataGridView2.Rows(i1).Cells(0).Value + "','" + TextBox1.Text + "')"
                        insertar(agregar1)
                    Next
                    Dim eliminar11 As String = "delete from alimentador where tejedora ='" + Trim(TextBox1.Text) + "'"
                    ELIMINAR(eliminar11)
                    For i2 = 0 To k3 - 1
                        Dim agregar2 As String = "insert into alimentador(alimentador,tejedora) values ('" + DataGridView3.Rows(i2).Cells(0).Value + "','" + TextBox1.Text + "')"
                        insertar(agregar2)
                    Next



                    Me.TabControl1.SelectedIndex = 0
                    Dim cmd5 As New SqlDataAdapter("SELECT codigo_mq AS CODIGO,nombre_mq AS NOMBRE,T.dele AS TIPO,T1.dele AS MARCA, sigla_mq AS SIGLA,M.capmin_mq,M.capnom_mq,M.capmax_mq FROM custom_vianny.DBO.maquinas M INNER JOIN custom_vianny.DBO.TAB0100 T ON M.ccia_mq=T.ccia AND M.tipmaq_mq = T.cele AND T.CCia = '01' AND T.CTab = 'TIPMAQ' LEFT JOIN  custom_vianny.DBO.TAB0100 T1  ON M.ccia_mq=T1.ccia AND M.marca_mq = T1.cele AND T1.CCia = '01' AND T1.CTab = 'MARMAQ' ", conx)
                    Dim dt1 As New DataTable
                    cmd5.Fill(dt1)
                    DataGridView1.DataSource = dt1
                    DataGridView1.Columns(7).Visible = False
                    DataGridView1.Columns(8).Visible = False
                    DataGridView1.Columns(9).Visible = False
                    DataGridView1.Columns(2).Width = 80
                    DataGridView1.Columns(3).Width = 170
                    DataGridView1.Columns(4).Width = 170
                    Label2.Text = 0
                    MessageBox.Show("DATOS INGRESADOS CORRECTAMENTE")
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox4.Text = ""
                    TextBox5.Text = ""
                    TextBox6.Text = ""
                    DataGridView4.DataSource = ""
                    DataGridView5.DataSource = ""
                    DataGridView2.Rows.Clear()
                    DataGridView3.Rows.Clear()
                End If
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        Label12.Text = 5
        Button5.Enabled = True
        Button6.Enabled = True
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click
        DataGridView4.DataSource = ""
        DataGridView5.DataSource = ""
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Me.TabControl1.SelectedIndex = 1
        abrir()
        Dim cmd1 As New SqlCommand("select top 1 codigo_mq as ncom from custom_vianny.dbo.maquinas  order by codigo_mq desc", conx)
        Dim da1 As Integer
        da1 = cmd1.ExecuteScalar

        TextBox1.Text = da1 + 1
        Select Case TextBox1.Text.Length
            Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = TextBox1.Text
        End Select

        llenar_combo_box(ComboBox1)
        llenar_combo_box1(ComboBox2)

        Me.TabControl1.TabPages(1).Enabled = True
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = False
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Me.TabControl1.TabPages(1).Enabled = True
        Dim MiDataSet As New DataSet()
        llenar_combo_box(ComboBox1)
        llenar_combo_box1(ComboBox2)
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(6).Value
        ComboBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        ComboBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value
        TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value
        TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(8).Value
        TextBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(9).Value

        Me.TabControl1.SelectedIndex = 1
        Button1.Enabled = False
        Button2.Enabled = True
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Dim cmd61 As New SqlDataAdapter("SELECT alimentador FROM alimentador where tejedora ='" + DataGridView1.Rows(e.RowIndex).Cells(2).Value + "'", conx)
        Dim dt10 As New DataTable
        cmd61.Fill(dt10)
        DataGridView4.DataSource = dt10

        Dim cmd62 As New SqlDataAdapter("SELECT alimentador FROM alimentador where tejedora ='" + DataGridView1.Rows(e.RowIndex).Cells(2).Value + "'", conx)
        Dim dt11 As New DataTable
        cmd62.Fill(dt11)
        DataGridView5.DataSource = dt11

    End Sub
End Class