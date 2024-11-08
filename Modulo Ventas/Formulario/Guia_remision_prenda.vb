Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Net.Mail
Public Class Guia_remision_prenda
    Public conx As SqlConnection
    Public enunciado3 As SqlCommand
    Public conn As SqlDataAdapter
    Public respuesta3 As SqlDataReader
    Dim ty, ty2, ty3 As New DataTable
    Public comando As SqlCommand
    Dim Rsr101, Rsr301, Rsr1, Rsr11, Rsr2, Rsr22, Rsr3, Rsr33, Rsr4, Rsr44, Rsr5, Rsr55, Rsr6, Rsr66, Rsr7, Rsr77, Rsr8, Rsr88, Rsr9, Rsr99, Rsr10, Rsr1010, Rsr20, Rsr30, Rsr222, Rsr23, Rsr300, Rsr100, t4 As SqlDataReader
    Dim valorAnterior1 As Integer = 0
    Dim valorAnterior2 As Integer = 0
    Dim valorAnterior3 As Integer = 0
    Dim valorAnterior4 As Integer = 0
    Dim valorAnterior5 As Integer = 0
    Dim valorAnterior6 As Integer = 0
    Dim valorAnterior7 As Integer = 0
    Dim valorAnterior8 As Integer = 0
    Dim valorAnterior9 As Integer = 0
    Dim valorAnterior10 As Integer = 0
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

            conn = New SqlDataAdapter("select rtrim(ltrim(motiv_17f)) AS motiv_17f,rtrim(ltrim(nomb_17f)) AS nomb_17f from custom_vianny.dbo.Fag1700 where ccia_17F ='" + Trim(Label33.Text) + "'", conx)
            conn.Fill(ty)
            ComboBox1.DataSource = ty
            ComboBox1.DisplayMember = "nomb_17f"
            ComboBox1.ValueMember = "motiv_17f"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box2(ByVal cb As ComboBox, ByVal alm As String)
        Try

            conn = New SqlDataAdapter("select dest_21s,rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where  Empr_21s =" + Trim(Label33.Text) + "AND almac_21s=" + alm, conx)
            conn.Fill(ty2)
            ComboBox3.DataSource = ty2
            ComboBox3.DisplayMember = "nomb_21s"
            ComboBox3.ValueMember = "dest_21s"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub correlativos()
        Dim fu As New vguiacabecera
        Dim func As New fguiasistema
        Dim fu1 As New vguiacabecera

        Dim func1 As New vguiacabecera
        Select Case TextBox1.Text.Length

            Case "1" : TextBox1.Text = "000" & "" & TextBox1.Text
            Case "2" : TextBox1.Text = "00" & "" & TextBox1.Text
            Case "3" : TextBox1.Text = "0" & "" & TextBox1.Text
            Case "4" : TextBox1.Text = TextBox1.Text

        End Select
        TextBox1.Text = TextBox1.Text
        fu.gserie = TextBox1.Text
        fu.gccia = Label33.Text
        TextBox2.Text = func.mostrar_guia_correlativo(fu) + 1
        Select Case TextBox2.Text.Length

            Case "1" : TextBox2.Text = "0000000" & "" & TextBox2.Text
            Case "2" : TextBox2.Text = "000000" & "" & TextBox2.Text
            Case "3" : TextBox2.Text = "00000" & "" & TextBox2.Text
            Case "4" : TextBox2.Text = "0000" & "" & TextBox2.Text
            Case "5" : TextBox2.Text = "000" & "" & TextBox2.Text
            Case "6" : TextBox2.Text = "00" & "" & TextBox2.Text
            Case "7" : TextBox2.Text = "0" & "" & TextBox2.Text
            Case "8" : TextBox2.Text = TextBox2.Text
        End Select
        func1.gmes = DateTime.Now.ToString("MM")
        func1.galmacen = Label35.Text
        func1.gano = Label36.Text
        func1.gccia = Label33.Text
        Label3.Text = func.mostrar_correlativo_nsa(func1) + 1
        Select Case Label3.Text.Length

            Case "1" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "000" & "" & Label3.Text
            Case "2" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "00" & "" & Label3.Text
            Case "3" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "0" & "" & Label3.Text
            Case "4" : Label3.Text = "14" & DateTime.Now.ToString("MM") & "" & Label3.Text
        End Select
        TextBox2.Select()
        TextBox2.Enabled = True
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox1.Text.ToString().Trim().Length
                Case 1 : TextBox1.Text = "000" & TextBox1.Text
                Case 2 : TextBox1.Text = "00" & TextBox1.Text
                Case 3 : TextBox1.Text = "0" & TextBox1.Text
                Case 4 : TextBox1.Text = TextBox1.Text
            End Select
            correlativos()

            TextBox2.Select()
        End If
    End Sub

    Private Sub Guia_remision_prenda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box(ComboBox1)
        llenar_combo_box2(ComboBox3, Label35.Text)
        ComboBox2.SelectedItem = 0

        CERRAR()
        correlativos()
        If Label33.Text = "02" Then
            TextBox1.ReadOnly = False
        End If
        TextBox1.Select()
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub
    Sub mostrarCabecera()
        Dim sql107 As String = "select CIA_3,sfactu_3,nfactu_3,fich_3,nombfich_3,fcom_3,ftraslad_3,docref_3,tidocr_3,sfactur_3,nfactur_3,nomb_3,direcc_3,chofer_3,placa_3,ruc_3,motivo_3,tienda_3,ppartida_3,ppllegad_3,destino_3,flag_3,almreg_3,origen_3,usuario_3,fecha_3,trans_3,cerrado_3,brevete_3,glosadoc_3,dnichofer_3,nsalida_3
 from custom_vianny.DBO.Fag0400 Where CIA_3 = '" + Label33.Text.ToString().Trim() + "' And sfactu_3 = '" + TextBox1.Text.ToString().Trim() + "' and nfactu_3 ='" + TextBox2.Text.ToString().Trim() + "' and almreg_3='" + Label35.Text + "'"
        Dim cmd107 As New SqlCommand(sql107, conx)
        Rsr100 = cmd107.ExecuteReader()
        Dim i5 As Integer
        i5 = DataGridView1.Rows.Count
        If Rsr100.Read() = True Then
            TextBox3.Text = Rsr100(12).ToString().Trim()
            TextBox4.Text = Rsr100(26).ToString().Trim()
            TextBox5.Text = Rsr100(11).ToString().Trim()
            TextBox8.Text = Rsr100(3).ToString().Trim()
            TextBox9.Text = Rsr100(4).ToString().Trim()
            TextBox10.Text = Rsr100(19).ToString().Trim()
            TextBox11.Text = Rsr100(15).ToString().Trim()
            TextBox12.Text = Rsr100(11).ToString().Trim()
            TextBox13.Text = Rsr100(12).ToString().Trim()
            TextBox14.Text = Rsr100(14).ToString().Trim()
            TextBox15.Text = Rsr100(28).ToString().Trim()
            TextBox16.Text = Rsr100(13).ToString().Trim()
            TextBox6.Text = Rsr100(29).ToString().Trim()
            TextBox18.Text = Rsr100(8).ToString().Trim()
            TextBox7.Text = Rsr100(9).ToString().Trim()
            TextBox17.Text = Rsr100(10).ToString().Trim()
            Label3.Text = Rsr100(31).ToString().Trim()
            DateTimePicker1.Value = Rsr100(5)
            DateTimePicker2.Value = Rsr100(6)
        End If
        Rsr100.Close()
    End Sub
    Sub mostrarDetalle()
        Dim sql107 As String = "select ccia_3m,sfactu_3m,nfactu_3m,item_3m,talm_3m,linea_3m,sinon_3m,pedido_3m,isnull(cant_3m,0),isnull(cant1_3m,0),isnull(cant2_3m,0),isnull(cant3_3m,0),isnull(cant4_3m,0),isnull(cant5_3m,0),isnull(cant6_3m,0),isnull(cant7_3m,0),isnull(cant8_3m,0),isnull(cant9_3m,0),isnull(cant10_3m,0),ordens_3m,flag_3m,qg.tallador_3,isnull( sn.nomb_sin,a.nomb_17)
 from custom_vianny.DBO.Fam0400 fm
 left join custom_vianny.dbo.qag0300 qg on fm.pedido_3m= qg.ncom_3 and fm.ccia_3m = qg.ccia
 left join custom_vianny.dbo.CAG1700 A on qg.ccia = a.ccia and qg.linea_3 = a.linea_17
 left join custom_vianny.dbo.Cag1700_Sinonimos sn on fm.linea_3m = sn.codigo_sin and fm.sinon_3m = sn.item_sin Where ccia_3m = '" + Label33.Text.ToString().Trim() + "' And sfactu_3m = '" + TextBox1.Text.ToString().Trim() + "' And nfactu_3m = '" + TextBox2.Text.ToString().Trim() + "' And talm_3m = '" + Label35.Text + "' "
        Dim cmd107 As New SqlCommand(sql107, conx)
        Rsr101 = cmd107.ExecuteReader()
        Dim i5, sum As Integer
        i5 = 0
        sum = 0
        While Rsr101.Read() = True
            DataGridView1.Rows().Add()
            DataGridView1.Rows(i5).Cells(0).Value = Rsr101(3).ToString().Trim()
            DataGridView1.Rows(i5).Cells(1).Value = Rsr101(7).ToString().Trim()
            DataGridView1.Rows(i5).Cells(2).Value = Rsr101(6).ToString().Trim()
            DataGridView1.Rows(i5).Cells(3).Value = Rsr101(5).ToString().Trim()
            DataGridView1.Rows(i5).Cells(4).Value = Rsr101(22).ToString().Trim()
            DataGridView1.Rows(i5).Cells(5).Value = Convert.ToInt32(Rsr101(9))
            DataGridView1.Rows(i5).Cells(6).Value = Convert.ToInt32(Rsr101(10))
            DataGridView1.Rows(i5).Cells(7).Value = Convert.ToInt32(Rsr101(11))
            DataGridView1.Rows(i5).Cells(8).Value = Convert.ToInt32(Rsr101(12))
            DataGridView1.Rows(i5).Cells(9).Value = Convert.ToInt32(Rsr101(13))
            DataGridView1.Rows(i5).Cells(10).Value = Convert.ToInt32(Rsr101(14))
            DataGridView1.Rows(i5).Cells(11).Value = Convert.ToInt32(Rsr101(15))
            DataGridView1.Rows(i5).Cells(12).Value = Convert.ToInt32(Rsr101(16))
            DataGridView1.Rows(i5).Cells(13).Value = Convert.ToInt32(Rsr101(17))
            DataGridView1.Rows(i5).Cells(14).Value = Convert.ToInt32(Rsr101(18))
            DataGridView1.Rows(i5).Cells(15).Value = Convert.ToInt32(Rsr101(8))
            DataGridView1.Rows(i5).Cells(16).Value = (Rsr101(21))
            i5 = i5 + 1
            sum = sum + Convert.ToInt32(Rsr101(8))
        End While
        TextBox19.Text = Convert.ToInt32(sum).ToString("N2")

        Rsr101.Close()
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim existe As Integer = 0
            CheckBox1.Visible = True
            TextBox25.Enabled = True

            TextBox2.Enabled = False

            ComboBox2.SelectedIndex = 0
            Select Case TextBox2.Text.ToString().Trim().Length

                Case 1 : TextBox2.Text = "0000000" & TextBox2.Text
                Case 2 : TextBox2.Text = "000000" & TextBox2.Text
                Case 3 : TextBox2.Text = "00000" & TextBox2.Text
                Case 4 : TextBox2.Text = "0000" & TextBox2.Text
                Case 5 : TextBox2.Text = "000" & TextBox2.Text
                Case 6 : TextBox2.Text = "00" & TextBox2.Text
                Case 7 : TextBox2.Text = "0" & TextBox2.Text
                Case 8 : TextBox2.Text = TextBox2.Text
            End Select
            TextBox4.Select()
            Dim sql1079 As String = "select COUNT(nfactu_3) from custom_vianny.DBO.Fag0400 Where CIA_3 = '" + Label33.Text.ToString().Trim() + "' And sfactu_3 = '" + TextBox1.Text.ToString().Trim() + "' and nfactu_3 ='" + TextBox2.Text.ToString().Trim() + "' and almreg_3='" + Label35.Text + "'"
            Dim cmd107 As New SqlCommand(sql1079, conx)
            t4 = cmd107.ExecuteReader()
            If t4.Read() = True Then
                existe = Convert.ToInt32(t4(0))
            End If
            t4.Close()

            If existe > 0 Then

                mostrarCabecera()
                mostrarDetalle()
                Button4.Enabled = True
                Button9.Enabled = True
                TextBox25.Enabled = False
                Button2.Enabled = False
                Button7.Enabled = False
                Button6.Enabled = True
            Else

                APERTURAR()
            End If
        Else
            If e.KeyCode = Keys.F1 Then
                Detalle_guia.TextBox2.Text = TextBox44.Text
                Detalle_guia.Label3.Text = Label35.Text.ToString().Trim()
                Detalle_guia.Label4.Text = Label33.Text.ToString().Trim()
                Detalle_guia.Label7.Text = 1
                Detalle_guia.ShowDialog()
            End If

        End If
    End Sub

    Private Sub Guia_remision_prenda_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "'  and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
        ELIMINAR(agregar)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.F1 Then

            Transportistas.Show()
            Transportistas.TextBox1.Text = 4
        Else
            If e.KeyCode = Keys.Enter Then
                If Trim(TextBox4.Text).Length > 0 Then
                    abrir()
                    enunciado3 = New SqlCommand("select nomb_18 , direcc_18,ruc_18,chofer_18,placa_18  AS PLA,brevete_18 from custom_vianny.dbo.Fag1800 where empr_18 =" + Trim(Label33.Text) + " and trans_18 =" + TextBox4.Text, conx)
                    respuesta3 = enunciado3.ExecuteReader
                    While respuesta3.Read
                        TextBox11.Text = respuesta3.Item("ruc_18")
                        TextBox12.Text = respuesta3.Item("nomb_18")
                        TextBox13.Text = respuesta3.Item("direcc_18")
                        TextBox14.Text = respuesta3.Item("PLA")
                        TextBox15.Text = respuesta3.Item("brevete_18")
                        TextBox16.Text = respuesta3.Item("chofer_18")
                        TextBox5.Text = respuesta3.Item("nomb_18")
                    End While
                    respuesta3.Close()
                    TextBox8.Select()
                Else
                    MsgBox("NO INGRESO NINGUN DATO EN TRANSPORTISTA")
                End If

            End If

        End If
    End Sub
    Private WithEvents editingControl As Control

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim indiceFilaSeleccionada As Integer = DataGridView1.SelectedRows(0).Index
            Dim numeroCorrelativo As Integer = CInt(DataGridView1.Rows(indiceFilaSeleccionada).Cells(0).Value)
            DataGridView1.Rows.Insert(indiceFilaSeleccionada + 1)
            Dim correlativo As Int16 = 0
            Dim correlativo1 As Int16 = 0
            For i As Integer = 0 To DataGridView1.Rows.Count - 1

                correlativo += 1
                Select Case correlativo.ToString().Length
                    Case 1 : DataGridView1.Rows(i).Cells(0).Value = "00" & correlativo
                    Case 2 : DataGridView1.Rows(i).Cells(0).Value = "0" & correlativo
                    Case 3 : DataGridView1.Rows(i).Cells(0).Value = correlativo
                End Select
            Next
        End If
    End Sub

    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit

        If e.ColumnIndex = "5" Then
            If DataGridView1.Rows(e.RowIndex).Cells(5).Value IsNot Nothing Then
                valorAnterior1 = DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString()
            Else
                valorAnterior1 = 0
            End If

        End If

        If e.ColumnIndex = "6" Then
            If DataGridView1.Rows(e.RowIndex).Cells(6).Value IsNot Nothing Then
                valorAnterior2 = DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString()
            Else
                valorAnterior2 = 0
            End If
        End If
        If e.ColumnIndex = "7" Then
            If DataGridView1.Rows(e.RowIndex).Cells(7).Value IsNot Nothing Then
                valorAnterior3 = DataGridView1.Rows(e.RowIndex).Cells(7).Value.ToString()
            Else
                valorAnterior3 = 0
            End If
        End If
        If e.ColumnIndex = "8" Then
            If DataGridView1.Rows(e.RowIndex).Cells(8).Value IsNot Nothing Then
                valorAnterior4 = DataGridView1.Rows(e.RowIndex).Cells(8).Value.ToString()
            Else
                valorAnterior4 = 0
            End If
        End If
        If e.ColumnIndex = "9" Then
            If DataGridView1.Rows(e.RowIndex).Cells(9).Value IsNot Nothing Then
                valorAnterior5 = DataGridView1.Rows(e.RowIndex).Cells(9).Value.ToString()
            Else
                valorAnterior5 = 0
            End If
        End If
        If e.ColumnIndex = "10" Then
            If DataGridView1.Rows(e.RowIndex).Cells(10).Value IsNot Nothing Then
                valorAnterior6 = DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString()
            Else
                valorAnterior6 = 0
            End If
        End If
        If e.ColumnIndex = "11" Then
            If DataGridView1.Rows(e.RowIndex).Cells(11).Value IsNot Nothing Then
                valorAnterior7 = DataGridView1.Rows(e.RowIndex).Cells(11).Value.ToString()
            Else
                valorAnterior7 = 0
            End If
        End If
        If e.ColumnIndex = "12" Then
            If DataGridView1.Rows(e.RowIndex).Cells(12).Value IsNot Nothing Then
                valorAnterior8 = DataGridView1.Rows(e.RowIndex).Cells(12).Value.ToString()
            Else
                valorAnterior8 = 0
            End If
        End If
        If e.ColumnIndex = "13" Then
            If DataGridView1.Rows(e.RowIndex).Cells(13).Value IsNot Nothing Then
                valorAnterior9 = DataGridView1.Rows(e.RowIndex).Cells(13).Value.ToString()
            Else
                valorAnterior9 = 0
            End If
        End If
        If e.ColumnIndex = "14" Then
            If DataGridView1.Rows(e.RowIndex).Cells(14).Value IsNot Nothing Then
                valorAnterior10 = DataGridView1.Rows(e.RowIndex).Cells(14).Value.ToString()
            Else
                valorAnterior10 = 0
            End If
        End If

    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        If editingControl IsNot Nothing Then
            RemoveHandler editingControl.KeyDown, AddressOf EditingControl_KeyDown
        End If

        ' Asignar el nuevo control de edición y agregar el evento KeyDown
        editingControl = e.Control
        AddHandler editingControl.KeyDown, AddressOf EditingControl_KeyDown
    End Sub
    Private Sub EditingControl_KeyDown(sender As Object, e As KeyEventArgs)
        ' Verificar si la tecla presionada es F1
        If e.KeyCode = Keys.F1 Then
            ' Verificar si la celda actual está en la columna "Factor de Consumo"
            If DataGridView1.CurrentCell.ColumnIndex = 2 Then
                ' Cancelar la edición para evitar conflictos
                DataGridView1.EndEdit()
                ' Abrir el formulario deseado
                TablaSinonimos.Label1.Text = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex.ToString).Cells(3).Value.ToString().Trim()
                TablaSinonimos.Label2.Text = Trim(Label33.Text.ToString())
                TablaSinonimos.Label3.Text = DataGridView1.CurrentCell.RowIndex.ToString
                TablaSinonimos.ShowDialog()
            End If

        End If
    End Sub



    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            Dim Cli As New Clientes
            Cli.TextBox3.Text = "39080"
            Cli.Show()
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub



    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        RETROCEDER()
        CERRAR()
        Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "'  and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
        ELIMINAR(agregar)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox43_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox25_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim VAL As Integer
            Select Case Trim(TextBox25.Text).Length
                Case "1" : TextBox25.Text = "01" & "0000000" & TextBox25.Text
                Case "2" : TextBox25.Text = "01" & "000000" & TextBox25.Text
                Case "3" : TextBox25.Text = "01" & "00000" & TextBox25.Text
                Case "4" : TextBox25.Text = "01" & "0000" & TextBox25.Text
                Case "5" : TextBox25.Text = "01" & "000" & TextBox25.Text
                Case "6" : TextBox25.Text = "01" & "00" & TextBox25.Text
                Case "7" : TextBox25.Text = "01" & "0" & TextBox25.Text
                Case "8" : TextBox25.Text = "01" & TextBox25.Text

            End Select

            Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, q.tallador_3 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(TextBox25.Text) + "' and q.ccia ='" + Trim(Label33.Text) + "'
"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            Dim i5 As Integer
            i5 = DataGridView1.Rows.Count

            If Rsr2.Read() = True Then
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i5).Cells(1).Value = Rsr2(0).ToString().Trim()
                DataGridView1.Rows(i5).Cells(2).Value = "000"
                DataGridView1.Rows(i5).Cells(3).Value = Rsr2(1).ToString().Trim()
                DataGridView1.Rows(i5).Cells(4).Value = Rsr2(2).ToString().Trim()
                DataGridView1.Rows(i5).Cells(5).Value = 0
                DataGridView1.Rows(i5).Cells(6).Value = 0
                DataGridView1.Rows(i5).Cells(7).Value = 0
                DataGridView1.Rows(i5).Cells(8).Value = 0
                DataGridView1.Rows(i5).Cells(9).Value = 0
                DataGridView1.Rows(i5).Cells(10).Value = 0
                DataGridView1.Rows(i5).Cells(11).Value = 0
                DataGridView1.Rows(i5).Cells(12).Value = 0
                DataGridView1.Rows(i5).Cells(13).Value = 0
                DataGridView1.Rows(i5).Cells(14).Value = 0
                DataGridView1.Rows(i5).Cells(15).Value = 0
                DataGridView1.Rows(i5).Cells(16).Value = Rsr2(3).ToString().Trim()
                DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                VAL = DataGridView1.Rows(i5).Cells(0).Value
                Select Case VAL.ToString.Length
                    Case "1" : DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                    Case "2" : DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                    Case "3" : DataGridView1.Rows(i5).Cells(0).Value = VAL
                End Select


                Rsr2.Close()


                Dim sql1 As String = "  select  cast( ISNULL(SUM(cant_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_Total,
                 cast( isnull(SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1,
                 cast( isnull(SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_2,
                 cast( isnull(SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_3,
                 cast( isnull(SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_4,
                 cast( isnull(SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_5,
                 cast( isnull(SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_6,
                 cast( isnull(SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_7,
                 cast( isnull(SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_8,
                 cast( isnull(SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_9,
                 cast( isnull(SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_10
                from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label33.Text) + "'  and talm_3b ='" + Trim(Label35.Text) + "' and pedido_3b ='" + Trim(DataGridView1.Rows(i5).Cells(1).Value) + "' and flag_3b ='1'  and cperiodo_3b ='" + Trim(Label36.Text) + "' "
                Dim cmd1 As New SqlCommand(sql1, conx)
                Rsr222 = cmd1.ExecuteReader()
                If Rsr222.Read() = True Then
                    TextBox35.Text = Rsr222(1)
                    TextBox34.Text = Rsr222(2)
                    TextBox33.Text = Rsr222(3)
                    TextBox32.Text = Rsr222(4)
                    TextBox31.Text = Rsr222(5)
                    TextBox30.Text = Rsr222(6)
                    TextBox29.Text = Rsr222(7)
                    TextBox28.Text = Rsr222(8)
                    TextBox27.Text = Rsr222(9)
                    TextBox26.Text = Rsr222(10)

                Else
                    TextBox35.Text = 0
                    TextBox34.Text = 0
                    TextBox33.Text = 0
                    TextBox32.Text = 0
                    TextBox31.Text = 0
                    TextBox30.Text = 0
                    TextBox18.Text = 0
                    TextBox29.Text = 0
                    TextBox28.Text = 0
                    TextBox27.Text = 0
                    TextBox26.Text = 0
                End If
                Rsr222.Close()
                Dim sql10235 As String = "SELECT A.*,
                		  ISNULL(B.Dele,'') AS NTalla1,
                		  ISNULL(C.Dele,'') AS NTalla2,
                		  ISNULL(D.Dele,'') AS NTalla3,
                		  ISNULL(E.Dele,'') AS NTalla4,
                		  ISNULL(F.Dele,'') AS NTalla5,
                		  ISNULL(G.Dele,'') AS NTalla6,
                		  ISNULL(H.Dele,'') AS NTalla7,
                		  ISNULL(I.Dele,'') AS NTalla8,
                		  ISNULL(J.Dele,'') AS NTalla9,
                		  ISNULL(K.Dele,'') AS NTalla10
                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
                 Where A.CCia_tl = '" + Trim(Label33.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(i5).Cells(16).Value) + "'
                ORDER BY  A.CCia_tl, A.Codigo_tl"
                Dim cmd10235 As New SqlCommand(sql10235, conx)
                Rsr23 = cmd10235.ExecuteReader()
                If Rsr23.Read() = True Then
                    DataGridView1.Columns(5).HeaderText = Trim(Rsr23(13).ToString())
                    DataGridView1.Columns(6).HeaderText = Trim(Rsr23(14).ToString())
                    DataGridView1.Columns(7).HeaderText = Trim(Rsr23(15).ToString())
                    DataGridView1.Columns(8).HeaderText = Trim(Rsr23(16).ToString())
                    DataGridView1.Columns(9).HeaderText = Trim(Rsr23(17).ToString())
                    DataGridView1.Columns(10).HeaderText = Trim(Rsr23(18).ToString())
                    DataGridView1.Columns(11).HeaderText = Trim(Rsr23(19).ToString())
                    DataGridView1.Columns(12).HeaderText = Trim(Rsr23(20).ToString())
                    DataGridView1.Columns(13).HeaderText = Trim(Rsr23(21).ToString())
                    DataGridView1.Columns(14).HeaderText = Trim(Rsr23(22).ToString())



                End If
                Rsr23.Close()
                TextBox25.Text = ""

                If (DataGridView1.Columns(5).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(5).ReadOnly = True

                Else
                    DataGridView1.Rows(i5).Cells(5).ReadOnly = False

                End If
                If (DataGridView1.Columns(6).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(6).ReadOnly = True
                Else
                    DataGridView1.Rows(i5).Cells(6).ReadOnly = False
                End If
                If (DataGridView1.Columns(7).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(7).ReadOnly = True
                Else
                    DataGridView1.Rows(i5).Cells(7).ReadOnly = False
                End If
                If (DataGridView1.Columns(8).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(8).ReadOnly = True
                Else
                    DataGridView1.Rows(i5).Cells(8).ReadOnly = False
                End If
                If (DataGridView1.Columns(9).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(9).ReadOnly = True
                Else
                    DataGridView1.Rows(i5).Cells(9).ReadOnly = False
                End If
                If (DataGridView1.Columns(10).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(10).ReadOnly = True
                Else
                    DataGridView1.Rows(i5).Cells(10).ReadOnly = False
                End If
                If (DataGridView1.Columns(11).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(11).ReadOnly = True
                Else
                    DataGridView1.Rows(i5).Cells(11).ReadOnly = False
                End If
                If (DataGridView1.Columns(12).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(12).ReadOnly = True
                Else
                    DataGridView1.Rows(i5).Cells(12).ReadOnly = False
                End If
                If (DataGridView1.Columns(13).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(13).ReadOnly = True
                Else
                    DataGridView1.Rows(i5).Cells(13).ReadOnly = False
                End If
                If (DataGridView1.Columns(14).HeaderText.ToString().Trim()).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(14).ReadOnly = True
                Else
                    DataGridView1.Rows(i5).Cells(14).ReadOnly = False
                End If
                TextBox25.Select()

                DataGridView1.ClearSelection()

                DataGridView1.Rows(i5).Selected = True

                DataGridView1.CurrentCell = DataGridView1.Rows(i5).Cells(1)
            Else
                MsgBox("LA OP INGRESADA NO EXISTE")
            End If
            Rsr2.Close()
            TextBox25.Select()
        End If

    End Sub
    Private Sub generarnsa()
        'cabecera nsa
        Dim cmd17 As New SqlCommand("IF NOT EXISTS( SELECT NCom_3 FROM custom_vianny.DBO.Mag030f WHERE CCia_3 = @CCIA_3 AND CPeriodo_3 = @CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 = @ccom_3 AND NCom_3 = @ncom_3)
        INSERT INTO custom_vianny.DBO.MAG030F (CCIA_3 , CPERIODO_3, talm_3, ccom_3, ncom_3, fcom_3, gloa_3, glob_3,fich_3 , orig_3, nombe_3, adevol_3, transf_3, tdoc_3, grm_3, fase_3, flag_3, origen_3, usuario_3, fecha_3, hora_3, subfase_3, tidoc_3,sfactu_3, nfactu_3, transfer_3, pedorig_3, Auto_3,orides_3,motdes_3) 
                                       VALUES (@CCIA_3,@CPERIODO_3,@TAlm_3,@ccom_3,@ncom_3,@fcom_3,@gloa_3,     '',@fich_3, @orig_3, ''    , 0       ,         0,@TDoc_3,@grm_3,''    ,1      ,@origen_3,@usuario_3,@fecha_3,@hora_3, ''       , '009', '" + TextBox1.Text.ToString().Trim() + "', '" + TextBox2.Text.ToString().Trim() + "', 0, '', 0,@orides_3,@motdes_3)
        				ELSE
        					UPDATE custom_vianny.DBO.MAG030F SET fcom_3  = @fcom_3, 
        					                   gloa_3  = @gloa_3, 
        					                   glob_3  = '', 
        					                   fich_3  = @fich_3, 
        					                   orig_3  = @orig_3, 
        					                   nombe_3 = '', 
        					                   adevol_3  = 0, 
        					                   transf_3  = 0, 
        					                   tdoc_3    =@TDoc_3,
        					                   grm_3     = @grm_3, 
        					                   fase_3    = '',
        					                   flag_3    = 1, 
        					                   origen_3  = @origen_3, 
        					                   usuario_3 = @usuario_3, 
        					                   fecha_3   =@fecha_3, 
        					                   hora_3    = @hora_3, 
        					                   subfase_3 = '', 
        					                   tidoc_3   = '009', 
        					                   sfactu_3  = '" + TextBox1.Text.ToString().Trim() + "', 
        					                   nfactu_3  = '" + TextBox2.Text.ToString().Trim() + "', 
        					                   transfer_3= 0, 
        					                   pedorig_3 = '',
        					                   Auto_3 = 0
        									   WHERE CCIA_3 = @CCIA_3 And CPERIODO_3 =@CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 =@ccom_3 And ncom_3 = @ncom_3", conx)
        cmd17.Parameters.AddWithValue("@CCIA_3", Label33.Text.ToString().Trim())
        cmd17.Parameters.AddWithValue("@CPERIODO_3", Trim(Label36.Text.ToString()))
        cmd17.Parameters.AddWithValue("@TAlm_3", "01")
        cmd17.Parameters.AddWithValue("@ccom_3", "2")
        cmd17.Parameters.AddWithValue("@ncom_3", Label3.Text.ToString().Trim())
        cmd17.Parameters.AddWithValue("@fcom_3", DateTimePicker1.Value)
        cmd17.Parameters.AddWithValue("@gloa_3", "")
        cmd17.Parameters.AddWithValue("@orig_3", ComboBox3.SelectedValue.ToString())
        cmd17.Parameters.AddWithValue("@fich_3", TextBox8.Text.ToString().Trim())
        cmd17.Parameters.AddWithValue("orides_3", "")
        cmd17.Parameters.AddWithValue("motdes_3", "")
        If Trim(ComboBox2.Text) = "INT" Then
            cmd17.Parameters.AddWithValue("@origen_3", "INT")
        Else
            cmd17.Parameters.AddWithValue("@origen_3", "EXT")
        End If
        cmd17.Parameters.AddWithValue("@TDoc_3", "009")
        cmd17.Parameters.AddWithValue("@grm_3", TextBox1.Text.ToString().Trim() & "-" & TextBox2.Text.ToString().Trim())
        cmd17.Parameters.AddWithValue("@usuario_3", Trim(Label37.Text.ToString()))
        cmd17.Parameters.AddWithValue("@fecha_3", DateTimePicker1.Value)
        cmd17.Parameters.AddWithValue("@hora_3", DateTime.Now.ToString("hh:mm:ss"))
        cmd17.ExecuteNonQuery()


        Dim agregar As String = "DELETE custom_vianny.DBO.Map030f Where CCIA_3a = '" + Trim(Label33.Text) + "' And CPERIODO_3a = '" + Trim(Label36.Text) + "' And talm_3a = '01' And ccom_3a = '2' And ncom_3a = '" + Label3.Text.ToString().Trim() + "'"
        Dim agregar1 As String = "DELETE custom_vianny.DBO.Mat030f Where CCIA_3b = '" + Trim(Label33.Text) + "' And CPERIODO_3b = '" + Trim(Label36.Text) + "' And talm_3b = '01' And ccom_3b = '2' And ncom_3b = '" + Label3.Text.ToString().Trim() + "'"
        ELIMINAR(agregar)
        ELIMINAR(agregar1)
        ' fin cabecera
        For i = 0 To DataGridView1.Rows.Count - 1
            Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.DBO.Mat030f (ccia_3b , cperiodo_3b, talm_3b, ccom_3b, ncom_3b, item_3b,linea_3b ,pedido_3b ,cant_3b ,cant1_3b ,cant2_3b ,cant3_3b ,cant4_3b ,cant5_3b ,cant6_3b ,cant7_3b ,cant8_3b ,cant9_3b ,cant10_3b,ordens_3b,flag_3b)
                        VALUES (@ccia_3b,@cperiodo_3b,@talm_3b,@ccom_3b,@ncom_3b,@item_3b,@linea_3b,@pedido_3b,@cant_3b,@cant1_3b,@cant2_3b,@cant3_3b,@cant4_3b,@cant5_3b,@cant6_3b,@cant7_3b,@cant8_3b,@cant9_3b,@cant10_3b,'', 1)", conx)
            cmd15.Parameters.AddWithValue("@ccia_3b", Label33.Text.ToString().Trim())
            cmd15.Parameters.AddWithValue("@cperiodo_3b", Trim(Label36.Text.ToString()))
            cmd15.Parameters.AddWithValue("@talm_3b", "01")
            cmd15.Parameters.AddWithValue("@ccom_3b", 2)
            cmd15.Parameters.AddWithValue("@ncom_3b", Label3.Text.ToString().Trim())
            cmd15.Parameters.AddWithValue("@item_3b", DataGridView1.Rows(i).Cells(0).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@linea_3b", DataGridView1.Rows(i).Cells(3).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@pedido_3b", DataGridView1.Rows(i).Cells(1).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@cant_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(15).Value))
            cmd15.Parameters.AddWithValue("@cant1_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value))
            cmd15.Parameters.AddWithValue("@cant2_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(6).Value))
            cmd15.Parameters.AddWithValue("@cant3_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
            cmd15.Parameters.AddWithValue("@cant4_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(8).Value))
            cmd15.Parameters.AddWithValue("@cant5_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value))
            cmd15.Parameters.AddWithValue("@cant6_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(10).Value))
            cmd15.Parameters.AddWithValue("@cant7_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(11).Value))
            cmd15.Parameters.AddWithValue("@cant8_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(12).Value))
            cmd15.Parameters.AddWithValue("@cant9_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(13).Value))
            cmd15.Parameters.AddWithValue("@cant10_3b", Convert.ToDouble(DataGridView1.Rows(i).Cells(14).Value))
            cmd15.ExecuteNonQuery()

            For y = 5 To 14
                If Convert.ToDouble(DataGridView1.Rows(i).Cells(y).Value) > 0 Then
                    Dim cmd16 As New SqlCommand("INSERT INTO custom_vianny.DBO.MAP030F (CCIA_3A , CPERIODO_3A, talm_3a, ccom_3a, ncom_3a, item_3a,linea_3a ,sinon_3a,linea2_3a, cant_3a, unid_3a, cantu_3a, obser1_3a, obser2_3a, obser3_3a, cantk_3a, cantke_3a, pedido_3a, flag_3a, peso_3a,tpeso_3a,unidk_3a , pieza_3a, color_3a, ncolor_3a, parti_3a, ocorte_3a, ordens_3a, nrollo_3a, kgrollo_3a, lote_3a, maquina_3a, galga_3a,sigla_3a, sitcal_3a, cantc_3a, ancho_3a, densid_3a, pedod_3a, ordtej_3a  , tipapt_3a , combo_3a,ccosto_3a,PreUn1_3a,Tot1_3a  )
                           Values (@CCIA_3A,@CPERIODO_3A,@talm_3a,@ccom_3a,@ncom_3aa,@item_3a,@linea_3a, ''     ,'PT0101' ,@cantid ,@unid_3a,  @cantid,''        , ''       ,''        ,@cantid  , 0.00     ,@pedido_3a, 1      , 0.00000, 0.00   ,@unidk_3a,   ''    , ''      , ''       , ''      , ''       , ''       , ''       , 0.00      , ''     , ''        , 0.00    ,''      , 0        , 0.000   , 0.000   , 0.000    , ''      , ''         , ''        , ''      ,''       ,0.00000  , 0.00)", conx)
                    cmd16.Parameters.AddWithValue("@CCIA_3A", Label33.Text.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@CPERIODO_3A", Trim(Label36.Text.ToString()))
                    cmd16.Parameters.AddWithValue("@talm_3a", "01")
                    cmd16.Parameters.AddWithValue("@ccom_3a", "2")
                    cmd16.Parameters.AddWithValue("@ncom_3aa", Label3.Text.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@item_3a", DataGridView1.Rows(i).Cells(0).Value.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@linea_3a", DataGridView1.Rows(i).Cells(3).Value.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@cantid", DataGridView1.Rows(i).Cells(y).Value.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@unid_3a", "UND")
                    cmd16.Parameters.AddWithValue("@pedido_3a", DataGridView1.Rows(i).Cells(1).Value.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@unidk_3a", "UND")
                    cmd16.ExecuteNonQuery()
                End If
            Next
        Next
    End Sub

    Private Sub generarguia()
        'inicio de cabecera
        Dim agregar As String = "DELETE custom_vianny.DBO.Fag0400 Where CIA_3 = '" + Label33.Text.ToString().Trim() + "' And sfactu_3 = '" + TextBox1.Text.ToString().Trim() + "' And nfactu_3 = '" + TextBox2.Text.ToString().Trim() + "' And almreg_3 = '01' "
        Dim agregar1 As String = "DELETE custom_vianny.DBO.Fam0400 Where ccia_3m = '" + Label33.Text.ToString().Trim() + "' And sfactu_3m = '" + TextBox1.Text.ToString().Trim() + "' And nfactu_3m = '" + TextBox2.Text.ToString().Trim() + "' And talm_3m = '01' "
        Dim agregar2 As String = "DELETE custom_vianny.DBO.Fap0400 Where CIA_3A = '" + Label33.Text.ToString().Trim() + "' And sfactu_3a = '" + TextBox1.Text.ToString().Trim() + "' And nfactu_3a ='" + TextBox2.Text.ToString().Trim() + "'"
        ELIMINAR(agregar2)
        ELIMINAR(agregar)
        ELIMINAR(agregar1)
        Dim cmd166 As New SqlCommand("insert into custom_vianny.DBO.Fag0400(CIA_3,sfactu_3,nfactu_3,fich_3,nombfich_3,fcom_3,ftraslad_3,docref_3,tidocr_3,sfactur_3,nfactur_3,nomb_3,direcc_3,chofer_3,placa_3,ruc_3,motivo_3,tienda_3,ppartida_3,ppllegad_3,destino_3,flag_3,origen_3,usuario_3,fecha_3,hora_3,nsalida_3,situa_3,trans_3,cerrado_3,brevete_3,glosadoc_3,dnichofer_3,afactu_4,almreg_3,refactu_3,ordenc_3,fase_3,subfase_3,fasereg_3,emp_3,almac_3,sfase_3,color_3,ordens_3,habavi_3,direc_3,ttrans_3,certificado_3) 
                                                 values (@CIA_3,@sfactu_3,@nfactu_3,@fich_3,@nombfich_3,@fcom_3,@ftraslad_3,@docref_3,@tidocr_3,@sfactur_3,@nfactur_3,@nomb_3,@direcc_3,@chofer_3,@placa_3,@ruc_3,@motivo_3,@tienda_3,@ppartida_3,@ppllegad_3,@destino_3,@flag_3,@origen_3,@usuario_3,CAST(GETDATE() AS DATE),CAST(GETDATE() AS TIME),@nsalida_3,@situa_3,@trans_3,@cerrado_3,@brevete_3,@glosadoc_3,@dnichofer_3,@afactu_4,@almreg_3,@refactu_3,@ordenc_3,@fase_3,@subfase_3,@fasereg_3,@emp_3,@almac_3,@sfase_3,@color_3,@ordens_3,@habavi_3,@direc_3,@ttrans_3,@certificado_3) ", conx)
        cmd166.Parameters.AddWithValue("@CIA_3", Label33.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@sfactu_3", TextBox1.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@nfactu_3", TextBox2.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@fich_3", TextBox8.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@nombfich_3", TextBox9.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@fcom_3", DateTimePicker1.Value.Date)
        cmd166.Parameters.AddWithValue("@ftraslad_3", DateTimePicker2.Value.Date)
        cmd166.Parameters.AddWithValue("@docref_3", 0)
        cmd166.Parameters.AddWithValue("@tidocr_3", TextBox18.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@sfactur_3", TextBox7.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@nfactur_3", TextBox17.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@nomb_3", TextBox12.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@direcc_3", TextBox13.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@chofer_3", TextBox16.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@placa_3", TextBox14.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@ruc_3", TextBox11.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@motivo_3", ComboBox1.SelectedValue.ToString())
        cmd166.Parameters.AddWithValue("@tienda_3", "01")
        cmd166.Parameters.AddWithValue("@ppartida_3", TextBox3.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@ppllegad_3", TextBox10.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@destino_3", ComboBox3.SelectedValue.ToString())
        cmd166.Parameters.AddWithValue("@flag_3", 1)
        cmd166.Parameters.AddWithValue("@origen_3", ComboBox2.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@usuario_3", "")
        cmd166.Parameters.AddWithValue("@nsalida_3", Label3.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@situa_3", 1)
        cmd166.Parameters.AddWithValue("@trans_3", TextBox4.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@cerrado_3", 1)
        cmd166.Parameters.AddWithValue("@brevete_3", TextBox15.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@glosadoc_3", TextBox6.Text.ToString().Trim())
        cmd166.Parameters.AddWithValue("@dnichofer_3", Mid(TextBox15.Text.ToString().Trim(), 2, 8))
        cmd166.Parameters.AddWithValue("@afactu_4", 0)
        cmd166.Parameters.AddWithValue("@almreg_3", "01")
        cmd166.Parameters.AddWithValue("@refactu_3", "-")
        cmd166.Parameters.AddWithValue("@ordenc_3", "")
        cmd166.Parameters.AddWithValue("@fase_3", "")
        cmd166.Parameters.AddWithValue("@subfase_3", "")
        cmd166.Parameters.AddWithValue("@fasereg_3", "")
        cmd166.Parameters.AddWithValue("@emp_3", "")
        cmd166.Parameters.AddWithValue("@almac_3", "")
        cmd166.Parameters.AddWithValue("@sfase_3", "")
        cmd166.Parameters.AddWithValue("@color_3", "")
        cmd166.Parameters.AddWithValue("@ordens_3", "")
        cmd166.Parameters.AddWithValue("@habavi_3", "")
        cmd166.Parameters.AddWithValue("@direc_3", "")
        cmd166.Parameters.AddWithValue("@ttrans_3", "")
        cmd166.Parameters.AddWithValue("@certificado_3", "")
        cmd166.ExecuteNonQuery()
        'fin de cabecera
        Dim sum As Integer = 0
        Dim it As String
        For i = 0 To DataGridView1.Rows.Count - 1
            'Detalle de guia
            Dim cmd167 As New SqlCommand("INSERT INTO custom_vianny.dbo.Fam0400 (CCia_3m, sfactu_3m, nfactu_3m, talm_3m,Item_3m,linea_3m, Sinon_3m,cant_3m, cant1_3m, cant2_3m, cant3_3m, cant4_3m, cant5_3m, cant6_3m, cant7_3m, cant8_3m, cant9_3m, cant10_3m, pedido_3m, flag_3m, ordens_3m)	
                                  Values (@CCia_3m,@sfactu_3m,@nfactu_3m,@talm_3m,@Item_3m,@linea_3m,@Sinon_3m,@cant_3m,@cant1_3m,@cant2_3m,@cant3_3m,@cant4_3m,@cant5_3m,@cant6_3m,@cant7_3m,@cant8_3m,@cant9_3m,@cant10_3m,@pedido_3m,@flag_3m,@ordens_3m)", conx)
            cmd167.Parameters.AddWithValue("@CCia_3m", Label33.Text.ToString().Trim())
            cmd167.Parameters.AddWithValue("@sfactu_3m", TextBox1.Text.ToString().Trim())
            cmd167.Parameters.AddWithValue("@nfactu_3m", TextBox2.Text.ToString().Trim())
            cmd167.Parameters.AddWithValue("@talm_3m", Label35.Text.ToString().Trim())
            cmd167.Parameters.AddWithValue("@Item_3m", DataGridView1.Rows(i).Cells(0).Value.ToString().Trim())
            cmd167.Parameters.AddWithValue("@linea_3m", DataGridView1.Rows(i).Cells(3).Value.ToString().Trim())
            cmd167.Parameters.AddWithValue("@Sinon_3m", DataGridView1.Rows(i).Cells(2).Value.ToString().Trim())
            cmd167.Parameters.AddWithValue("@cant_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(15).Value))
            cmd167.Parameters.AddWithValue("@cant1_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(5).Value))
            cmd167.Parameters.AddWithValue("@cant2_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(6).Value))
            cmd167.Parameters.AddWithValue("@cant3_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value))
            cmd167.Parameters.AddWithValue("@cant4_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(8).Value))
            cmd167.Parameters.AddWithValue("@cant5_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(9).Value))
            cmd167.Parameters.AddWithValue("@cant6_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(10).Value))
            cmd167.Parameters.AddWithValue("@cant7_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(11).Value))
            cmd167.Parameters.AddWithValue("@cant8_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(12).Value))
            cmd167.Parameters.AddWithValue("@cant9_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(13).Value))
            cmd167.Parameters.AddWithValue("@cant10_3m", Convert.ToDouble(DataGridView1.Rows(i).Cells(14).Value))
            cmd167.Parameters.AddWithValue("@pedido_3m", DataGridView1.Rows(i).Cells(1).Value.ToString().Trim())
            cmd167.Parameters.AddWithValue("@flag_3m", "0")
            cmd167.Parameters.AddWithValue("@unidk_3a", "UND")
            cmd167.Parameters.AddWithValue("@ordens_3m", "")
            cmd167.ExecuteNonQuery()
            '    ' fin detalle de guia

            ''    'inicio de detalle guia general

            For y = 5 To 14
                If Convert.ToDouble(DataGridView1.Rows(i).Cells(y).Value) > 0 Then
                    Dim cmd168 As New SqlCommand("INSERT INTO custom_vianny.dbo.Fap0400 (Cia_3a, sfactu_3a, nfactu_3a, Item_3a,linea_3a, Sinon_3a,cant_3a, unid_3a, cantk_3a, pedido_3a, flag_3a, unidk_3a,ordens_3a,linea2_3a     ,obser2_3a,obser3_3a,obser4_3a,obser5_3a,obser6_3a,obser7_3a,obser8_3a,obser9_3a,obser10_3a,PARTI_3A,secpartida_3a,ocorte_3a,ordtej_3a,ps_3a,lote_3a)
                                                                                 Values (@Cia_3a,@sfactu_3a,@nfactu_3a,@Item_3a,@linea_3a,@Sinon_3a,@cant_3a,@unid_3a,@cantk_3a,@pedido_3a,@flag_3a,@unidk_3a,@ordens_3a,@linea2_3a,''       ,''       ,''       ,''       ,''       ,''       ,''       ,''       ,''        ,''      ,''           ,''       ,''       ,'','')", conx)
                    cmd168.Parameters.AddWithValue("@Cia_3a", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@sfactu_3a", TextBox1.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@nfactu_3a", TextBox2.Text.ToString().Trim())

                    cmd168.Parameters.AddWithValue("@linea_3a", DataGridView1.Rows(i).Cells(3).Value.ToString().Trim())
                    Select Case sum.ToString().Length
                        Case 1 : it = "00" & sum
                        Case 2 : it = "0" & sum
                        Case 3 : it = sum
                    End Select
                    cmd168.Parameters.AddWithValue("@Item_3a", it)
                    cmd168.Parameters.AddWithValue("@Sinon_3a", DataGridView1.Rows(i).Cells(2).Value.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@cant_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(y).Value))
                    cmd168.Parameters.AddWithValue("@unid_3a", "UND")
                    cmd168.Parameters.AddWithValue("@cantk_3a", Convert.ToDouble(DataGridView1.Rows(i).Cells(y).Value))
                    cmd168.Parameters.AddWithValue("@pedido_3a", DataGridView1.Rows(i).Cells(1).Value.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@flag_3a", "1")
                    cmd168.Parameters.AddWithValue("@unidk_3a", "UND")
                    cmd168.Parameters.AddWithValue("@ordens_3a", "")
                    cmd168.Parameters.AddWithValue("@linea2_3a", "")
                    cmd168.ExecuteNonQuery()
                    sum = sum + 1
                End If

            Next
            '    'fin de detalle guia general
        Next



    End Sub
    Sub actualizarop()
        For i = 0 To DataGridView1.Rows.Count - 1
            Dim cmd168 As New SqlCommand("update custom_vianny.dbo.qag0300 set finped_3 ='1' where ncom_3 = @ncom_3 and ccia =@ccia", conx)
            cmd168.Parameters.AddWithValue("@ncom_3", DataGridView1.Rows(i).Cells(1).Value.ToString().Trim())
            cmd168.Parameters.AddWithValue("@ccia", Label33.Text.ToString().Trim())
            cmd168.ExecuteNonQuery()
        Next
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            generarguia()
            generarnsa()
            actualizarop()
            MsgBox("Se Registro la informacion Correctamente")
            limpiar()
            correlativos()
            CERRAR()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim sql102213 As String = "exec Anular_guia '" + Label13.Text + "','" + Label3.Text + "','" + Trim(Label36.Text.ToString()) + "','" + TextBox1.Text.ToString().Trim() + "','" + TextBox2.Text.ToString().Trim() + "'"
            Dim cmd102213 As New SqlCommand(sql102213, conx)
            Rsr301 = cmd102213.ExecuteReader()
            Dim jo As Integer
            Rsr301.Read()
            Rsr301.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MsgBox("Se Anulo la Informacion Correctamente")
        limpiar()
        correlativos()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If RadioButton2.Checked = True Then
            MsgBox("LA GUIA DE REMISION ESTA ANULADA NO SE PUEDE EDITAR")
        Else
            Dim sql102213 As String = "SELECT cierre_3 FROM custom_vianny.dbo.Cie0300 A  Where a.ccia ='" + Label13.Text + "' AND A.talm_3 ='" + ComboBox3.Text + "' AND A.fcie_3='" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
            Dim cmd102213 As New SqlCommand(sql102213, conx)
            Rsr300 = cmd102213.ExecuteReader()
            Dim jo As Integer
            If Rsr300.Read() Then
                jo = Rsr300(0)
            Else
                jo = 0
            End If
            Rsr300.Close()
            Dim respuesta As DialogResult
            respuesta = MessageBox.Show("QUIERES EDITAR ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                If jo = 0 Then

                    DataGridView1.Enabled = True
                    DataGridView1.Columns(7).ReadOnly = False

                    TextBox9.Enabled = True
                    DateTimePicker1.Enabled = True
                    DateTimePicker2.Enabled = True
                    DataGridView1.ReadOnly = False
                    'TextBox12.Enabled = True
                    TextBox25.Enabled = True
                    TextBox4.Enabled = True
                    TextBox10.Enabled = True
                    TextBox20.Enabled = True
                    TextBox5.Enabled = False
                    Button2.Enabled = True
                    Button5.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox3.Enabled = True
                    TextBox16.Enabled = True
                    TextBox26.Enabled = True
                    TextBox17.Enabled = True
                    TextBox6.Enabled = True
                    Button1.Enabled = True
                    Button3.Enabled = True
                    Button7.Enabled = True
                    TextBox8.Enabled = True
                    TextBox43.Text = "2"
                Else
                    MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
                End If

            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        Dim filaSeleccionada As DataGridViewRow = DataGridView1.SelectedRows(0)
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then


            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "'  and codigo_la = '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%' and RIGHT(RTRIM(codigo_la), 3)   ='" + filaSeleccionada.Cells("Item").Value + "'"
            ELIMINAR(agregar)
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        End If
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        DataGridView1.BeginEdit(True)
        DataGridView1.Columns(0).ReadOnly = True

        DataGridView1.Columns(3).ReadOnly = True
        DataGridView1.Columns(4).ReadOnly = True
        If e.ColumnIndex >= 2 And e.ColumnIndex <= 14 Then
            If DataGridView1.Rows(e.RowIndex).Cells(1).Value IsNot Nothing AndAlso
   DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim().Length > 0 AndAlso
   DataGridView1.Rows(e.RowIndex).Cells(16).Value IsNot Nothing AndAlso
   DataGridView1.Rows(e.RowIndex).Cells(16).Value.ToString().Trim().Length > 0 Then

                Dim stock1, stock2, stock3, stock4, stock5, stock6, stock7, stock8, stock9, stock10 As Double
                Dim jo1, jo2, jo3, jo4, jo5, jo6, jo7, jo8, jo9, jo10 As Double
                stock1 = 0
                stock2 = 0
                stock3 = 0
                stock4 = 0
                stock5 = 0
                stock6 = 0
                stock7 = 0
                stock8 = 0
                stock9 = 0
                stock10 = 0
                jo1 = 0
                jo2 = 0
                jo3 = 0
                jo4 = 0
                jo5 = 0
                jo6 = 0
                jo7 = 0
                jo8 = 0
                jo9 = 0
                jo10 = 0
                Dim sql102358 As String = "SELECT A.*,
                        		  ISNULL(B.Dele,'') AS NTalla1,
                        		  ISNULL(C.Dele,'') AS NTalla2,
                        		  ISNULL(D.Dele,'') AS NTalla3,
                        		  ISNULL(E.Dele,'') AS NTalla4,
                        		  ISNULL(F.Dele,'') AS NTalla5,
                        		  ISNULL(G.Dele,'') AS NTalla6,
                        		  ISNULL(H.Dele,'') AS NTalla7,
                        		  ISNULL(I.Dele,'') AS NTalla8,
                        		  ISNULL(J.Dele,'') AS NTalla9,
                        		  ISNULL(K.Dele,'') AS NTalla10
                         FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
                         				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
                         				 LEFT JOIN custom_vianny.dbo.Tab0100 C
                         				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
                         				 LEFT JOIN custom_vianny.dbo.Tab0100 D
                         				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
                         				
                         				 LEFT JOIN custom_vianny.dbo.Tab0100 E
                         				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
                         				 LEFT JOIN custom_vianny.dbo.Tab0100 F
                         				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
                         				 LEFT JOIN custom_vianny.dbo.Tab0100 G
                         				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
                         				 LEFT JOIN custom_vianny.dbo.Tab0100 H
                         				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
                         				 LEFT JOIN custom_vianny.dbo.Tab0100 I
                         				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
                         				 LEFT JOIN custom_vianny.dbo.Tab0100 J
                         				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
                         				 LEFT JOIN custom_vianny.dbo.Tab0100 K
                         				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
                         Where A.CCia_tl = '" + Trim(Label33.Text.ToString()) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(e.RowIndex).Cells(16).Value) + "'
                        ORDER BY  A.CCia_tl, A.Codigo_tl"
                Dim cmd102358 As New SqlCommand(sql102358, conx)
                Rsr20 = cmd102358.ExecuteReader()
                If Rsr20.Read() = True Then
                    DataGridView1.Columns(5).HeaderText = Trim(Rsr20(13).ToString().Trim())
                    DataGridView1.Columns(6).HeaderText = Trim(Rsr20(14).ToString().Trim())
                    DataGridView1.Columns(7).HeaderText = Trim(Rsr20(15).ToString().Trim())
                    DataGridView1.Columns(8).HeaderText = Trim(Rsr20(16).ToString().Trim())
                    DataGridView1.Columns(9).HeaderText = Trim(Rsr20(17).ToString().Trim())
                    DataGridView1.Columns(10).HeaderText = Trim(Rsr20(18).ToString().Trim())
                    DataGridView1.Columns(11).HeaderText = Trim(Rsr20(19).ToString().Trim())
                    DataGridView1.Columns(12).HeaderText = Trim(Rsr20(20).ToString().Trim())
                    DataGridView1.Columns(13).HeaderText = Trim(Rsr20(21).ToString().Trim())
                    DataGridView1.Columns(14).HeaderText = Trim(Rsr20(22).ToString().Trim())
                End If
                Rsr20.Close()
                ' columna 1
                Dim sql1021 As String = "select  cast( isnull(SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1,cast( isnull(SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_2,
            cast( isnull(SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_3,cast( isnull(SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_4,
            cast( isnull(SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_5,cast( isnull(SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_6,
            cast( isnull(SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_7,cast( isnull(SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_8,
            cast( isnull(SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_9,cast( isnull(SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_10
             from custom_vianny.dbo.mat030f  where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr1 = cmd1021.ExecuteReader()
                If Rsr1.Read() Then
                    stock1 = Convert.ToInt32(Rsr1(0))
                    stock2 = Convert.ToInt32(Rsr1(1))
                    stock3 = Convert.ToInt32(Rsr1(2))
                    stock4 = Convert.ToInt32(Rsr1(3))
                    stock5 = Convert.ToInt32(Rsr1(4))
                    stock6 = Convert.ToInt32(Rsr1(5))
                    stock7 = Convert.ToInt32(Rsr1(6))
                    stock8 = Convert.ToInt32(Rsr1(7))
                    stock9 = Convert.ToInt32(Rsr1(8))
                    stock10 = Convert.ToInt32(Rsr1(9))

                End If
                Rsr1.Close()

                Dim sql1022139 As String = " select cast( isnull(SUM( ISNULL(stock1_la,0)),0) AS INT) as stock1,cast( isnull(SUM( ISNULL(stock2_la,0)),0) AS INT) as stock2, cast( isnull(SUM( ISNULL(stock3_la,0)),0) AS INT) as stock3, cast( isnull(SUM( ISNULL(stock4_la,0)),0) AS INT) as stock4, 
             cast( isnull(SUM( ISNULL(stock5_la,0)),0) AS INT) as stock5, cast( isnull(SUM( ISNULL(stock6_la,0)),0) AS INT) as stock6, cast( isnull(SUM( ISNULL(stock7_la,0)),0) AS INT) as stock7, cast( isnull(SUM( ISNULL(stock8_la,0)),0) AS INT) as stock8, 
             cast( isnull(SUM( ISNULL(stock9_la,0)),0) AS INT) as stock9, cast( isnull(SUM( ISNULL(stock10_la,0)),0) AS INT) as stock10
             from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la LIKE '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
                Dim cmd1022139 As New SqlCommand(sql1022139, conx)
                Rsr11 = cmd1022139.ExecuteReader()

                If Rsr11.Read() Then
                    jo1 = Convert.ToInt32(Rsr11(0))
                    jo2 = Convert.ToInt32(Rsr11(1))
                    jo3 = Convert.ToInt32(Rsr11(2))
                    jo4 = Convert.ToInt32(Rsr11(3))
                    jo5 = Convert.ToInt32(Rsr11(4))
                    jo6 = Convert.ToInt32(Rsr11(5))
                    jo7 = Convert.ToInt32(Rsr11(6))
                    jo8 = Convert.ToInt32(Rsr11(7))
                    jo9 = Convert.ToInt32(Rsr11(8))
                    jo10 = Convert.ToInt32(Rsr11(9))
                Else
                    jo1 = 0
                    jo2 = 0
                    jo3 = 0
                    jo4 = 0
                    jo5 = 0
                    jo6 = 0
                    jo7 = 0
                    jo8 = 0
                    jo9 = 0
                    jo10 = 0

                End If
                Rsr11.Close()

                TextBox35.Text = (stock1 - jo1).ToString()
                TextBox34.Text = (stock2 - jo2).ToString()
                TextBox33.Text = (stock3 - jo3).ToString()
                TextBox32.Text = (stock4 - jo4).ToString()
                TextBox31.Text = (stock5 - jo5).ToString()
                TextBox30.Text = (stock6 - jo6).ToString()
                TextBox29.Text = (stock7 - jo7).ToString()
                TextBox28.Text = (stock8 - jo8).ToString()
                TextBox27.Text = (stock9 - jo9).ToString()
                TextBox26.Text = (stock10 - jo10).ToString()

                ' fin columna 1
                '////////////////////////////////////////////////////////////////////////////////////////////////

            End If
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox44.Visible = True
            TextBox45.Visible = True
            TextBox45.Select()
        Else
            TextBox44.Visible = False
            TextBox45.Visible = False
            TextBox45.Select()
        End If
    End Sub

    Private Sub TextBox45_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox45.KeyDown

        Detalle_guia.TextBox2.Text = TextBox44.Text
        Detalle_guia.Label3.Text = Label35.Text.ToString().Trim()
        Detalle_guia.Label4.Text = Label33.Text.ToString().Trim()
        Detalle_guia.Label7.Text = "2"
        Detalle_guia.ShowDialog()
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged


    End Sub
    Sub limpiar()
        TextBox2.Text = ""

        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox20.Text = ""
        TextBox25.Text = ""
        TextBox45.Text = ""
        ComboBox3.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        TextBox44.Visible = False
        TextBox45.Visible = False
        CheckBox1.Visible = False
        CheckBox1.Checked = False
        Button5.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Button2.Enabled = False
        DataGridView1.Rows.Clear()
        TextBox1.Select()
        TextBox35.Text = 0
        TextBox34.Text = 0
        TextBox33.Text = 0
        TextBox32.Text = 0
        TextBox31.Text = 0
        TextBox30.Text = 0
        TextBox29.Text = 0
        TextBox28.Text = 0
        TextBox27.Text = 0
        TextBox26.Text = 0
        TextBox42.Text = 0
        TextBox41.Text = 0
        TextBox40.Text = 0
        TextBox39.Text = 0
        TextBox38.Text = 0
        TextBox37.Text = 0
        TextBox24.Text = 0
        TextBox23.Text = 0
        TextBox22.Text = 0
        TextBox21.Text = 0
    End Sub
    Sub RETROCEDER()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox20.Text = ""
        TextBox25.Text = ""
        TextBox45.Text = ""
        ComboBox3.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        TextBox44.Visible = False
        TextBox45.Visible = False
        CheckBox1.Visible = False
        CheckBox1.Checked = False
        Button5.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Button2.Enabled = False
        DataGridView1.Rows.Clear()
        TextBox1.Select()
        TextBox35.Text = 0
        TextBox34.Text = 0
        TextBox33.Text = 0
        TextBox32.Text = 0
        TextBox31.Text = 0
        TextBox30.Text = 0
        TextBox29.Text = 0
        TextBox28.Text = 0
        TextBox27.Text = 0
        TextBox26.Text = 0
        TextBox42.Text = 0
        TextBox41.Text = 0
        TextBox40.Text = 0
        TextBox39.Text = 0
        TextBox38.Text = 0
        TextBox37.Text = 0
        TextBox24.Text = 0
        TextBox23.Text = 0
        TextBox22.Text = 0
        TextBox21.Text = 0
    End Sub
    Sub CERRAR()
        DateTimePicker1.Enabled = False
        DateTimePicker2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False
        TextBox12.Enabled = False
        TextBox13.Enabled = False
        TextBox14.Enabled = False
        TextBox15.Enabled = False
        TextBox16.Enabled = False
        TextBox20.Enabled = False
        TextBox25.Enabled = False
        ComboBox3.Enabled = False
        ComboBox2.Enabled = False
        ComboBox1.Enabled = False
        Button5.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Button2.Enabled = False
        Button7.Enabled = False
        DataGridView1.ReadOnly = True
    End Sub

    Sub APERTURAR()
        DataGridView1.ReadOnly = False
        DateTimePicker1.Enabled = True
        DateTimePicker2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox8.Enabled = True
        TextBox9.Enabled = True
        TextBox10.Enabled = True
        TextBox11.Enabled = True
        TextBox12.Enabled = True
        TextBox13.Enabled = True
        TextBox14.Enabled = True
        TextBox15.Enabled = True
        TextBox16.Enabled = True
        TextBox20.Enabled = True
        TextBox25.Enabled = True
        ComboBox3.Enabled = True
        ComboBox2.Enabled = True
        ComboBox1.Enabled = True
        Button5.Enabled = True

        Button6.Enabled = True
        Button2.Enabled = True
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
    Dim Rsr212, Rsr333 As SqlDataReader
    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = "1" Then
            If DataGridView1.Rows(e.RowIndex).Cells(1).Value IsNot Nothing AndAlso
   DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim().Length > 0 Then


                abrir()
                Dim VAL As Integer
                Select Case DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim().Length
                    Case "1" : DataGridView1.Rows(e.RowIndex).Cells(1).Value = "01" & "0000000" & DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                    Case "2" : DataGridView1.Rows(e.RowIndex).Cells(1).Value = "01" & "000000" & DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                    Case "3" : DataGridView1.Rows(e.RowIndex).Cells(1).Value = "01" & "00000" & DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                    Case "4" : DataGridView1.Rows(e.RowIndex).Cells(1).Value = "01" & "0000" & DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                    Case "5" : DataGridView1.Rows(e.RowIndex).Cells(1).Value = "01" & "000" & DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                    Case "6" : DataGridView1.Rows(e.RowIndex).Cells(1).Value = "01" & "00" & DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                    Case "7" : DataGridView1.Rows(e.RowIndex).Cells(1).Value = "01" & "0" & DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()
                    Case "8" : DataGridView1.Rows(e.RowIndex).Cells(1).Value = "01" & DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim()

                End Select

                Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, q.tallador_3 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and q.ccia ='" + Trim(Label33.Text) + "'
"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()
                Dim i5 As Integer
                i5 = DataGridView1.Rows.Count

                If Rsr2.Read() = True Then

                    DataGridView1.Rows(e.RowIndex).Cells(1).Value = Rsr2(0).ToString().Trim()
                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = "000"
                    DataGridView1.Rows(e.RowIndex).Cells(3).Value = Rsr2(1).ToString().Trim()
                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = Rsr2(2).ToString().Trim()
                    DataGridView1.Rows(e.RowIndex).Cells(5).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(6).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(7).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(8).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(9).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(10).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(12).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(13).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(14).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(15).Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells(16).Value = Rsr2(3).ToString().Trim()
                    DataGridView1.Rows(e.RowIndex).Cells(0).Value = e.RowIndex + 1
                    VAL = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                    Select Case VAL.ToString.Length
                        Case "1" : DataGridView1.Rows(e.RowIndex).Cells(0).Value = "00" & "" & VAL
                        Case "2" : DataGridView1.Rows(e.RowIndex).Cells(0).Value = "0" & "" & VAL
                        Case "3" : DataGridView1.Rows(e.RowIndex).Cells(0).Value = VAL
                    End Select


                    Rsr2.Close()


                    Dim sql1 As String = "  select  cast( ISNULL(SUM(cant_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_Total,
                 cast( isnull(SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1,
                 cast( isnull(SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_2,
                 cast( isnull(SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_3,
                 cast( isnull(SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_4,
                 cast( isnull(SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_5,
                 cast( isnull(SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_6,
                 cast( isnull(SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_7,
                 cast( isnull(SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_8,
                 cast( isnull(SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_9,
                 cast( isnull(SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_10
                from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label33.Text) + "'  and talm_3b ='" + Trim(Label35.Text) + "' and pedido_3b ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value) + "' and flag_3b ='1'  and cperiodo_3b ='" + Trim(Label36.Text) + "' "
                    Dim cmd1 As New SqlCommand(sql1, conx)
                    Rsr222 = cmd1.ExecuteReader()
                    If Rsr222.Read() = True Then
                        TextBox35.Text = Rsr222(1)
                        TextBox34.Text = Rsr222(2)
                        TextBox33.Text = Rsr222(3)
                        TextBox32.Text = Rsr222(4)
                        TextBox31.Text = Rsr222(5)
                        TextBox30.Text = Rsr222(6)
                        TextBox29.Text = Rsr222(7)
                        TextBox28.Text = Rsr222(8)
                        TextBox27.Text = Rsr222(9)
                        TextBox26.Text = Rsr222(10)

                    Else
                        TextBox35.Text = 0
                        TextBox34.Text = 0
                        TextBox33.Text = 0
                        TextBox32.Text = 0
                        TextBox31.Text = 0
                        TextBox30.Text = 0
                        TextBox18.Text = 0
                        TextBox29.Text = 0
                        TextBox28.Text = 0
                        TextBox27.Text = 0
                        TextBox26.Text = 0
                    End If
                    Rsr222.Close()
                    Dim sql10235 As String = "SELECT A.*,
                		  ISNULL(B.Dele,'') AS NTalla1,
                		  ISNULL(C.Dele,'') AS NTalla2,
                		  ISNULL(D.Dele,'') AS NTalla3,
                		  ISNULL(E.Dele,'') AS NTalla4,
                		  ISNULL(F.Dele,'') AS NTalla5,
                		  ISNULL(G.Dele,'') AS NTalla6,
                		  ISNULL(H.Dele,'') AS NTalla7,
                		  ISNULL(I.Dele,'') AS NTalla8,
                		  ISNULL(J.Dele,'') AS NTalla9,
                		  ISNULL(K.Dele,'') AS NTalla10
                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
                 Where A.CCia_tl = '" + Trim(Label33.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(e.RowIndex).Cells(16).Value) + "'
                ORDER BY  A.CCia_tl, A.Codigo_tl"
                    Dim cmd10235 As New SqlCommand(sql10235, conx)
                    Rsr23 = cmd10235.ExecuteReader()
                    If Rsr23.Read() = True Then
                        DataGridView1.Columns(5).HeaderText = Trim(Rsr23(13).ToString())
                        DataGridView1.Columns(6).HeaderText = Trim(Rsr23(14).ToString())
                        DataGridView1.Columns(7).HeaderText = Trim(Rsr23(15).ToString())
                        DataGridView1.Columns(8).HeaderText = Trim(Rsr23(16).ToString())
                        DataGridView1.Columns(9).HeaderText = Trim(Rsr23(17).ToString())
                        DataGridView1.Columns(10).HeaderText = Trim(Rsr23(18).ToString())
                        DataGridView1.Columns(11).HeaderText = Trim(Rsr23(19).ToString())
                        DataGridView1.Columns(12).HeaderText = Trim(Rsr23(20).ToString())
                        DataGridView1.Columns(13).HeaderText = Trim(Rsr23(21).ToString())
                        DataGridView1.Columns(14).HeaderText = Trim(Rsr23(22).ToString())



                    End If
                    Rsr23.Close()
                    TextBox25.Text = ""

                    If (DataGridView1.Columns(5).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(5).ReadOnly = True

                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(5).ReadOnly = False

                    End If
                    If (DataGridView1.Columns(6).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = True
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = False
                    End If
                    If (DataGridView1.Columns(7).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(7).ReadOnly = True
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(7).ReadOnly = False
                    End If
                    If (DataGridView1.Columns(8).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(8).ReadOnly = True
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(8).ReadOnly = False
                    End If
                    If (DataGridView1.Columns(9).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(9).ReadOnly = True
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(9).ReadOnly = False
                    End If
                    If (DataGridView1.Columns(10).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(10).ReadOnly = True
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(10).ReadOnly = False
                    End If
                    If (DataGridView1.Columns(11).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(11).ReadOnly = True
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(11).ReadOnly = False
                    End If
                    If (DataGridView1.Columns(12).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(12).ReadOnly = True
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(12).ReadOnly = False
                    End If
                    If (DataGridView1.Columns(13).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(13).ReadOnly = True
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(13).ReadOnly = False
                    End If
                    If (DataGridView1.Columns(14).HeaderText.ToString().Trim()).Length = 0 Then
                        DataGridView1.Rows(e.RowIndex).Cells(14).ReadOnly = True
                    Else
                        DataGridView1.Rows(e.RowIndex).Cells(14).ReadOnly = False
                    End If

                    DataGridView1.ClearSelection()

                    DataGridView1.Rows(e.RowIndex).Selected = True

                    DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(1)
                Else
                    MsgBox("LA OP INGRESADA NO EXISTE")
                End If
                Rsr2.Close()
            End If
        End If
            If e.ColumnIndex = "2" Then
            If DataGridView1.Rows(e.RowIndex).Cells(2).Value IsNot Nothing AndAlso DataGridView1.Rows(e.RowIndex).Cells(1).Value IsNot Nothing AndAlso
                DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim().Length > 0 AndAlso DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim().Length > 0 Then
                Dim sin As String = ""
                Select Case DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim().Length
                    Case 1 : sin = "00" & DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim()
                    Case 2 : sin = "0" & DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim()
                    Case 3 : sin = DataGridView1.Rows(e.RowIndex).Cells(2).Value.ToString().Trim()
                    Case 0 : sin = "000"
                End Select

                Dim sql1022139 As String = "SELECT  item_sin as Codigo, nomb_sin as Descripcion FROM custom_vianny.DBO.Cag1700_Sinonimos WHERE codigo_sin ='" + DataGridView1.Rows(e.RowIndex).Cells(3).Value.ToString().Trim() + "' and ccia_sin ='" + Label33.Text.ToString().Trim() + "'  and item_sin ='" + sin + "'"
                Dim cmd1022139 As New SqlCommand(sql1022139, conx)
                Rsr11 = cmd1022139.ExecuteReader()
                If Rsr11.Read() Then
                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = Rsr11(0).ToString().Trim()
                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = Rsr11(1).ToString().Trim()
                Else
                    MsgBox("El Codigo Ingresado No Existe, Vuelva a ingresar uno correcto")
                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = "000"
                End If
                Rsr11.Close()
            Else
                DataGridView1.Rows(e.RowIndex).Cells(2).Value = "000"
            End If


        End If
        If e.ColumnIndex = "5" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(5).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock1_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr333 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr333.Read() Then
                jo = Convert.ToInt32(Rsr333(0))
            Else
                jo = 0
            End If
            Rsr333.Close()
            'TextBox35.Text = (stock - jo).ToString()
            If TextBox43.Text = "2" Then

                stockFinal = (stock - jo) + valorAnterior1

            Else
                stockFinal = (stock - jo)
            End If
            'MsgBox(valorAnterior1.ToString)
            'MsgBox(Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value))
            'MsgBox(stockFinal)
            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value) > stockFinal Then

                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior1 + (Convert.ToInt32(TextBox35.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(5).Value = valorAnterior1
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox35.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(5).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock1_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock1_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock1_la", stockFinal - valorAnterior1)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock1_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value))
                    'End If
                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(5).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If


            End If
            Dim suma5 As Integer = 0
            Dim suma55 As Integer = 0
            For i1 = 5 To 14
                suma5 = suma5 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma5.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "6" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(6).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock, stockFinal As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock2_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox43.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior2

            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value) > stockFinal Then
                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior2 + (Convert.ToInt32(TextBox34.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(6).Value = valorAnterior2
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox34.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(6).Value = 0
                End If
            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock2_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock2_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock2_la", stockFinal - valorAnterior2)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock2_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value))
                    'End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(6).HeaderText & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma6 As Integer = 0
            Dim suma66 As Integer = 0
            For i1 = 5 To 14
                suma6 = suma6 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma6.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "7" Then

            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(7).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock3_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox43.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior3
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value) > stockFinal Then
                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior3 + (Convert.ToInt32(TextBox33.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(7).Value = valorAnterior3
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox33.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(7).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock3_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock3_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock3_la", stockFinal - valorAnterior3)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock3_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value))
                    'End If
                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(7).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma7 As Integer = 0
            Dim suma77 As Integer = 0
            For i1 = 5 To 14
                suma7 = suma7 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma7.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "8" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(8).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock4_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox43.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior4
            Else
                stockFinal = (stock - jo)
            End If
            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value) > stockFinal Then
                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior4 + (Convert.ToInt32(TextBox32.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(8).Value = valorAnterior4
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox32.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(8).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock4_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock4_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock4_la", stockFinal - valorAnterior4)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock4_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value))
                    'End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(8).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma8 As Integer = 0
            Dim suma88 As Integer = 0
            For i1 = 5 To 14
                suma8 = suma8 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma8.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "9" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(9).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock5_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox43.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior5
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value) > stockFinal Then
                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior5 + (Convert.ToInt32(TextBox31.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(9).Value = valorAnterior5
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox31.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(9).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock5_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock5_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock5_la", stockFinal - valorAnterior5)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock5_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value))
                    'End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(9).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma9 As Integer = 0
            Dim suma99 As Integer = 0
            For i1 = 5 To 14
                suma9 = suma9 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma9.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "10" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(10).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock6_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox43.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior6
            Else
                stockFinal = (stock - jo)
            End If
            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value) > stockFinal Then
                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior6 + (Convert.ToInt32(TextBox30.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(10).Value = valorAnterior6
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox30.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(10).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock6_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock6_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock6_la", stockFinal - valorAnterior6)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock6_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value))
                    'End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(10).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma10 As Integer = 0
            Dim suma100 As Integer = 0
            For i1 = 5 To 14
                suma10 = suma10 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma10.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "11" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(11).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock7_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox43.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior7
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value) > stockFinal Then
                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior7 + (Convert.ToInt32(TextBox29.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(11).Value = valorAnterior7
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox29.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock7_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock7_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock7_la", stockFinal - valorAnterior7)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock7_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value))
                    'End If
                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(11).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma11 As Integer = 0
            Dim suma111 As Integer = 0
            For i1 = 5 To 14
                suma11 = suma11 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma11.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "12" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(12).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock, stockFinal As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock8_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox43.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior8
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value) > stockFinal Then
                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior8 + (Convert.ToInt32(TextBox28.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(12).Value = valorAnterior8
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox28.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(12).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock8_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock8_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock8_la", stockFinal - valorAnterior8)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock8_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value))
                    'End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(12).HeaderText & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma12 As Integer = 0
            Dim suma122 As Integer = 0
            For i1 = 5 To 14
                suma12 = suma12 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma12.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "13" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(13).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock, stockFinal As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock9_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            'MsgBox(stock.ToString())
            'MsgBox(jo.ToString())
            If TextBox43.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior9
            Else
                stockFinal = (stock - jo)
            End If
            'MsgBox(stock.ToString())
            'MsgBox(jo.ToString())
            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value) > stockFinal Then
                MsgBox("paso")
                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior9 + (Convert.ToInt32(TextBox27.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(13).Value = valorAnterior9
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox27.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(13).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock9_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock9_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock9_la", stockFinal - valorAnterior9)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock9_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value))
                    'End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(13).HeaderText & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma13 As Integer = 0
            Dim suma133 As Integer = 0
            For i1 = 5 To 14
                suma13 = suma13 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma13.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "14" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(14).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock, stockFinal As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label33.Text.ToString().Trim() + "'  and talm_3b ='" + Label35.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label36.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()
            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock10_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label33.Text.ToString().Trim() + "' and almac_la ='" + Label35.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox43.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior10
            Else
                stockFinal = (stock - jo)
            End If
            'MsgBox(stock.ToString())
            'MsgBox(jo.ToString())
            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(14).Value) > stockFinal Then
                If TextBox43.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior10 + (Convert.ToInt32(TextBox27.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(14).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(14).Value = valorAnterior10
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox27.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(14).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(14).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(14).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock10_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock10_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label33.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", Label35.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    'If TextBox43.Text = "2" Then
                    '    cmd168.Parameters.AddWithValue("@stock10_la", stockFinal - valorAnterior10)
                    'Else
                    cmd168.Parameters.AddWithValue("@stock10_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(14).Value))
                    'End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox1.Text.ToString().Trim() & TextBox2.Text.ToString().Trim() & DataGridView1.Columns(14).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma14 As Integer = 0
            Dim suma144 As Integer = 0
            For i1 = 5 To 14
                suma14 = suma14 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(15).Value = suma14.ToString("N2")
            SumarCantidades()
        End If
    End Sub
    Private Sub SumarCantidades()
        Dim total As Double = 0
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim cantidad As Integer
            'If Integer.TryParse(Convert.ToString(row.Cells("ctotal").Value), cantidad) Then
            cantidad = row.Cells("ctotal").Value
            total += cantidad
            'End If

        Next

        ' Muestra el total en el Label o en el TextBox correspondiente
        TextBox19.Text = total.ToString("N2") ' o TextBoxTotal.Text = total.ToString("N2")
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub

    'Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
    '    If e.KeyCode = Keys.Enter Then
    '        Dim currentRow As Integer = DataGridView1.CurrentCell.RowIndex
    '        Dim currentColumn As Integer = DataGridView1.CurrentCell.ColumnIndex
    '        MsgBox("paso")
    '        If currentColumn = 2 Then
    '            MessageBox.Show("Acción ejecutada en columna 'Sin'")
    '            Dim sin As String = ""
    '            Select Case DataGridView1.Rows(currentRow).Cells(2).Value.ToString().Trim().Length
    '                Case 1 : sin = "00" & DataGridView1.Rows(currentRow).Cells(2).Value.ToString().Trim()
    '                Case 2 : sin = "0" & DataGridView1.Rows(currentRow).Cells(2).Value.ToString().Trim()
    '                Case 3 : sin = DataGridView1.Rows(currentRow).Cells(2).Value.ToString().Trim()
    '                Case 0 : sin = "000"
    '            End Select

    '            Dim sql1022139 As String = "SELECT  item_sin as Codigo, nomb_sin as Descripcion FROM custom_vianny.DBO.Cag1700_Sinonimos WHERE codigo_sin ='" + DataGridView1.Rows(currentRow).Cells(3).Value.ToString().Trim() + "' and ccia_sin ='" + Label33.Text.ToString().Trim() + "'  and item_sin ='" + sin + "'"
    '            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
    '            Rsr11 = cmd1022139.ExecuteReader()
    '            If Rsr11.Read() Then
    '                DataGridView1.Rows(currentRow).Cells(2).Value = Rsr11(0).ToString().Trim()
    '                DataGridView1.Rows(currentRow).Cells(4).Value = Rsr11(1).ToString().Trim()
    '            Else
    '                MsgBox("El Codigo Ingresado No Existe, Vuelva a ingresar uno correcto")
    '                DataGridView1.Rows(currentRow).Cells(2).Value = "000"
    '            End If
    '            Rsr11.Close()
    '        End If
    '    End If

    'End Sub

    Private Sub TextBox43_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub
End Class