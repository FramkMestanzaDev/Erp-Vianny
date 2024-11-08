Imports System.Data.SqlClient
Public Class DEFECTOS_AUDITORIA
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da, ty As New DataTable
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
    Sub MOSTRAR()
        abrir()
        Dim cmd As New SqlDataAdapter("select tdr_codigo AS CODIGO , tdr_descripcion as DESCRIPCION,T.dele AS TIPO_DEFECTO  from custom_vianny.dbo.tdefectos_rollo td inner join custom_vianny.dbo.tab0100 t on td.tdr_tipodefecto = t.cele where ccia ='01'	and ctab='TDTEJ'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 70
        DataGridView1.Columns(1).Width = 320
        DataGridView1.Columns(2).Width = 120
    End Sub
    'Function mostrar_campo(ByVal sql)
    '    abrir()
    '    comando = New SqlCommand(sql, conx)
    '    Dim i As Integer = comando.ExecuteNonQuery
    '    If i > 0 Then
    '        Return True
    '    Else
    '        Return False
    '    End If
    'End Function
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox3.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 70
                DataGridView1.Columns(1).Width = 320
                DataGridView1.Columns(2).Width = 120
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub DEFECTOS_AUDITORIA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MOSTRAR()
        Me.TabControl1.TabPages(1).Enabled = False
        'If Me.TabControl1.SelectedIndex = 1 Then
        '    Button1.Enabled = True
        'End If
        Label5.Text = DataGridView1.Rows(0).Cells(0).Value
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        buscar()
    End Sub



    Sub llenar_combo_box(ByVal cb As ComboBox)
        Try

            conn = New SqlDataAdapter("select cele,dele from  custom_vianny.dbo.tab0100 where ccia ='01' and ctab='TDTEJ'", conx)
            conn.Fill(ty)
            ComboBox1.DataSource = ty
            ComboBox1.DisplayMember = "dele"
            ComboBox1.ValueMember = "cele"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.TabControl1.SelectedIndex = 1
        abrir()
        Dim cmd1 As New SqlCommand("select top 1 tdr_codigo as ncom from custom_vianny.dbo.tdefectos_rollo  order by tdr_codigo desc", conx)
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
        ComboBox1.SelectedIndex = 0
        Me.TabControl1.TabPages(1).Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = False
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox2.Enabled = True
    End Sub

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
                Dim agregar As String = "delete from custom_vianny.dbo.tdefectos_rollo where tdr_codigo ='" + Trim(Label5.Text) + "'"
                If (Eliminar(agregar)) Then
                    MessageBox.Show("DATO ELIMINADO CORRECTAMENTE")
                    Dim cmd As New SqlDataAdapter("select tdr_codigo AS CODIGO , tdr_descripcion as DESCRIPCION,T.dele AS TIPO_DEFECTO  from custom_vianny.dbo.tdefectos_rollo td inner join custom_vianny.dbo.tab0100 t on td.tdr_tipodefecto = t.cele where ccia ='01'	and ctab='TDTEJ'", conx)

                    cmd.Fill(da)
                    DataGridView1.DataSource = da
                    DataGridView1.Columns(0).Width = 70
                    DataGridView1.Columns(1).Width = 320
                    DataGridView1.Columns(2).Width = 120
                    Label2.Text = 0
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Label5.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            MsgBox("EL CAMPO DESCRIPCION ESTA VACIO")
        Else
            abrir()
            Dim cmd1 As New SqlDataAdapter("select *  from custom_vianny.dbo.tdefectos_rollo where tdr_descripcion like '" + TextBox2.Text + "%'", conx)
            Dim da1 As New DataTable
            cmd1.Fill(da1)

            If da1.Rows.Count > 0 Then
                MsgBox("EL DEFECTO DE CALIDAD YA ESTA REGISTRADO")
            Else
                Dim hj As New vtejeduria
                Dim jj As New ftejeduria

                hj.gtdr_codigo = TextBox1.Text
                hj.gtdr_descripcion = TextBox2.Text
                hj.gtdr_usuario = MDIParent1.Label3.Text
                hj.gtdr_fupdate = DateTimePicker1.Value
                hj.gtdr_pc = My.Computer.Name
                hj.gtdr_tipodefecto = ComboBox1.SelectedValue.ToString

                If jj.insertar_defectos(hj) Then
                    MsgBox("DATOS INGRESADO CORRECTAMENTE")
                    Me.TabControl1.SelectedIndex = 0
                    Dim cmd5 As New SqlDataAdapter("select tdr_codigo AS CODIGO , tdr_descripcion as DESCRIPCION,T.dele AS TIPO_DEFECTO  from custom_vianny.dbo.tdefectos_rollo td inner join custom_vianny.dbo.tab0100 t on td.tdr_tipodefecto = t.cele where ccia ='01'	and ctab='TDTEJ'", conx)

                    cmd5.Fill(da)
                    DataGridView1.DataSource = da
                    DataGridView1.Columns(0).Width = 70
                    DataGridView1.Columns(1).Width = 320
                    DataGridView1.Columns(2).Width = 120
                    Label2.Text = 0
                End If
            End If
        End If

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        ComboBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Me.TabControl1.SelectedIndex = 1
        Button1.Enabled = True
        TextBox2.Enabled = False

    End Sub
End Class
