Imports System.Data.SqlClient
Public Class Tabla_Comisiones
    Public conx As SqlConnection
    Public conn As SqlDataAdapter

    Public comando As SqlCommand
    Dim da As New DataTable
    Dim bs As New BindingSource()
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim ty, ty2, ty3 As New DataTable
    Sub llenar_combo_box2()

        Try

            conn = New SqlDataAdapter("SELECT oper_14f AS RUBRO,nomb_14f as DESCRIPCION FROM custom_vianny.DBO.fag1400 WHERE  ccia_14f='" + Label4.Text + "' AND VEN=1", conx)
            conn.Fill(ty2)
            ComboBox3.DataSource = ty2
            ComboBox3.DisplayMember = "DESCRIPCION"
            ComboBox3.ValueMember = "RUBRO"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Tabla_Comisiones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()

        llenar_combo_box2()
        ComboBox1.SelectedIndex = 1
        ComboBox2.SelectedIndex = -1

        correlativo()


    End Sub
    Sub correlativo()
        Dim Rsr1991 As SqlDataReader
        Dim sql1011 As String = "select top 1 correl_tabla from  tabla_comisiones order by correl_tabla desc"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        If Rsr1991.Read() Then

            TextBox3.Text = Rsr1991(0) + 1
        Else
            TextBox3.Text = 1
        End If
        Select Case Trim(TextBox3.Text).Length
            Case "1" : TextBox3.Text = "0000000" & TextBox3.Text
            Case "2" : TextBox3.Text = "000000" & TextBox3.Text
            Case "3" : TextBox3.Text = "00000" & TextBox3.Text
            Case "4" : TextBox3.Text = "0000" & TextBox3.Text
            Case "5" : TextBox3.Text = "000" & TextBox3.Text
            Case "6" : TextBox3.Text = "00" & TextBox3.Text
            Case "7" : TextBox3.Text = "0" & TextBox3.Text
            Case "8" : TextBox3.Text = TextBox3.Text

        End Select
        Rsr1991.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Trim(ComboBox2.Text) = "" Then
            MsgBox("SELECCIONA OTRA OPCION")
        Else
            DataGridView1.Rows.Clear()
            Dim Rsr19912 As SqlDataReader
            Dim i51 As Integer
            Dim sql10112 As String = "select codigo,nombre from  TABLA_RUBROS"
            Dim cmd10112 As New SqlCommand(sql10112, conx)
            Rsr19912 = cmd10112.ExecuteReader()
            i51 = 0

            While Rsr19912.Read()
                DataGridView1.Rows.Add(1)
                DataGridView1.Rows(i51).Cells(0).Value = Rsr19912(1)
                DataGridView1.Rows(i51).Cells(2).Value = 0
                DataGridView1.Rows(i51).Cells(4).Value = 0
                DataGridView1.Rows(i51).Cells(6).Value = Rsr19912(0)
                DataGridView1.Rows(i51).Cells(5).Value = Trim(ComboBox2.Text)
                DataGridView1.Rows(i51).Cells(3).Value = 0
                i51 = i51 + 1
            End While
            Rsr19912.Close()

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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim jk As New ftablacomisiones
        Dim kk As New vtablacomisiones
        Dim u As Integer
        Dim agregar1 As String = "delete from tabla_comisiones where correl_tabla ='" + Trim(TextBox3.Text) + "'"
        ELIMINAR(agregar1)
        u = DataGridView1.Rows.Count

        For i = 0 To u - 1
            kk.gid = Trim(TextBox3.Text)
            kk.gccia = Label4.Text
            kk.gperiodo = Trim(TextBox1.Text)


            kk.grubro = DataGridView1.Rows(i).Cells(0).Value

            kk.gfpagocontado = "1"


            kk.gpreciocontado = DataGridView1.Rows(i).Cells(2).Value
            kk.goperacioncontado = DataGridView1.Rows(i).Cells(3).Value
            kk.gcomisioncontado = DataGridView1.Rows(i).Cells(4).Value
            kk.gfecha = DateTimePicker1.Value.Date
            kk.galmacen = DataGridView1.Rows(i).Cells(5).Value
            kk.grubro_codigo = DataGridView1.Rows(i).Cells(6).Value
            jk.insertar_tabal_comisiones(kk)
        Next
        MsgBox("SE REGISTRO LA INFORMACION SOLICITADA")
        DataGridView1.Rows.Clear()
        correlativo()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
    Dim DT1, DT2 As New DataTable

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)

        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox4.Enabled = True
        Else
            TextBox4.Enabled = False
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        DataGridView1.Rows.Clear()
        TextBox4.Text = ""
        CheckBox1.Checked = False
        correlativo()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub DataGridView1_DefaultValuesNeeded(sender As Object, e As DataGridViewRowEventArgs) Handles DataGridView1.DefaultValuesNeeded

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim ko As New Exportar
        ko.llenarExcel(DataGridView1)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ComboBox2_DoubleClick(sender As Object, e As EventArgs) Handles ComboBox2.DoubleClick


    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox4.Text).Length
                Case "1" : TextBox4.Text = "0000000" & TextBox4.Text
                Case "2" : TextBox4.Text = "000000" & TextBox4.Text
                Case "3" : TextBox4.Text = "00000" & TextBox4.Text
                Case "4" : TextBox4.Text = "0000" & TextBox4.Text
                Case "5" : TextBox4.Text = "000" & TextBox4.Text
                Case "6" : TextBox4.Text = "00" & TextBox4.Text
                Case "7" : TextBox4.Text = "0" & TextBox4.Text
                Case "8" : TextBox4.Text = TextBox4.Text

            End Select

            DataGridView1.Rows.Clear()
            Dim Rsr199126 As SqlDataReader
            Dim i516 As Integer
            Dim sql101126 As String = "select rubro,preciocontado,operacioncontado,comisioncontado,rubro_codigo,almacen from tabla_comisiones where ccia ='" + Label4.Text + "' and correl_tabla ='" + Trim(TextBox4.Text) + "' "
            Dim cmd101126 As New SqlCommand(sql101126, conx)
            Rsr199126 = cmd101126.ExecuteReader()
            i516 = 0

            While Rsr199126.Read()
                DataGridView1.Rows.Add(1)
                DataGridView1.Rows(i516).Cells(0).Value = Rsr199126(0)
                DataGridView1.Rows(i516).Cells(2).Value = Rsr199126(1)
                DataGridView1.Rows(i516).Cells(4).Value = Rsr199126(3)
                DataGridView1.Rows(i516).Cells(6).Value = Rsr199126(4)
                DataGridView1.Rows(i516).Cells(5).Value = Rsr199126(5)
                DataGridView1.Rows(i516).Cells(3).Value = Rsr199126(2)
                i516 = i516 + 1
            End While
            Rsr199126.Close()
            TextBox4.Text = ""
            CheckBox1.Checked = False
        End If

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case Trim(TextBox3.Text).Length
                Case "1" : TextBox3.Text = "0000000" & TextBox3.Text
                Case "2" : TextBox3.Text = "000000" & TextBox3.Text
                Case "3" : TextBox3.Text = "00000" & TextBox3.Text
                Case "4" : TextBox3.Text = "0000" & TextBox3.Text
                Case "5" : TextBox3.Text = "000" & TextBox3.Text
                Case "6" : TextBox3.Text = "00" & TextBox3.Text
                Case "7" : TextBox3.Text = "0" & TextBox3.Text
                Case "8" : TextBox3.Text = TextBox3.Text

            End Select

            DataGridView1.Rows.Clear()
            Dim Rsr199127 As SqlDataReader
            Dim i517 As Integer
            Dim sql101127 As String = "select rubro,preciocontado,operacioncontado,comisioncontado,rubro_codigo,almacen  from tabla_comisiones where ccia ='" + Label4.Text + "' and correl_tabla ='" + Trim(TextBox3.Text) + "' "
            Dim cmd101127 As New SqlCommand(sql101127, conx)
            Rsr199127 = cmd101127.ExecuteReader()
            i517 = 0

            While Rsr199127.Read()
                DataGridView1.Rows.Add(1)
                DataGridView1.Rows(i517).Cells(0).Value = Rsr199127(0)
                DataGridView1.Rows(i517).Cells(2).Value = Rsr199127(1)
                DataGridView1.Rows(i517).Cells(4).Value = Rsr199127(3)
                DataGridView1.Rows(i517).Cells(6).Value = Rsr199127(4)
                DataGridView1.Rows(i517).Cells(5).Value = Rsr199127(5)
                DataGridView1.Rows(i517).Cells(3).Value = Rsr199127(2)
                i517 = i517 + 1
            End While
            Rsr199127.Close()


        End If

    End Sub
End Class