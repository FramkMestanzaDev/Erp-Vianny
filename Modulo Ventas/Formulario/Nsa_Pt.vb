Imports System.Data.SqlClient
Public Class Nsa_Pt
    Public conx As SqlConnection
    Public enunciado, enunciado5 As SqlCommand
    Public respuesta, respuesta5 As SqlDataReader
    Public conn As SqlDataAdapter
    Dim ty2 As New DataTable
    Public comando As SqlCommand
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Dim ty, ty3 As New DataTable
    Dim Rsr101, Rsr301, Rsr1, Rsr11, Rsr2, Rsr333, Rsr22, Rsr3, Rsr33, Rsr4, Rsr44, Rsr5, Rsr55, Rsr6, Rsr66, Rsr7, Rsr77, Rsr8, Rsr88, Rsr9, Rsr99, Rsr10, Rsr1010, Rsr20, Rsr30, Rsr222, Rsr23, Rsr300, Rsr100, t4, Rsr212 As SqlDataReader
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
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Select Case ComboBox3.Text
            Case "01" : TextBox2.Text = "PRODUCTO TERMINADO"
            Case "32" : TextBox2.Text = "TIENDA PLANTA 1"
            Case "50" : TextBox2.Text = "PRODUCTO TERMINADO SEGUNDA"
        End Select

        TextBox4.Enabled = True
        TextBox4.ReadOnly = False
        TextBox4.Text = DateTimePicker1.Value.Month
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & TextBox4.Text
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select
        ty2.Clear()
        abrir()
        llenar_combo_box2(ComboBox1, ComboBox3.Text)
        TextBox4.Select()
    End Sub

    Private Sub Nsa_Pt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox22.Enabled = False
        RadioButton1.Checked = True
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub
    Sub CORRELATIVO()
        Dim JU As Integer
        Select Case TextBox4.Text.Length

            Case "1" : TextBox4.Text = "0" & "" & TextBox4.Text
            Case "2" : TextBox4.Text = TextBox4.Text
        End Select
        'Dim dtFecha As Date = DateSerial(Month(Date.Today), 1, 1)

        Dim ano As String
        ano = Convert.ToString(Year(DateTimePicker1.Value))
        'DateTimePicker1.Value = "01/" + TextBox4.Text + "/" + MDIParent1.Label5.Text
        abrir()
        TextBox5.Enabled = True
        'MsgBox("select top 1 ncom_3, substring(ncom_3,3,7) as ncom from custom_vianny.dbo.mag030f  where ccia_3 = '01' and cperiodo_3 =  YEAR(GETDATE()) and talm_3 = " + ComboBox3.Text + " " + "and ncom_3 like" + " " + TextBox4.Text + "% and ccom_3 = '2' order by ncom_3 desc")
        enunciado = New SqlCommand("select top 1 ncom_3, substring(ncom_3,3,7) as ncom from custom_vianny.dbo.mag030f  where ccia_3 =" + Label13.Text + " and cperiodo_3 = '" + Label11.Text + "' and talm_3 = '" + ComboBox3.Text + "' " + "and ncom_3 like" + " '" + TextBox4.Text + "%' and ccom_3 = '2' order by ncom_3 desc", conx)
        respuesta = enunciado.ExecuteReader
        While respuesta.Read
            JU = respuesta.Item("ncom")
        End While
        respuesta.Close()
        TextBox5.Text = JU + 1
        Select Case TextBox5.Text.Length

            Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
            Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
            Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
            Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
            Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
            Case "6" : TextBox5.Text = TextBox5.Text
        End Select
        TextBox5.Select()
        TextBox5.ReadOnly = False
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            CORRELATIVO()
        End If

    End Sub
    Sub llenar_combo_box2(ByVal cb As ComboBox, ByVal alm As String)
        Try

            conn = New SqlDataAdapter("select dest_21s,rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where  Empr_21s =" + Trim(Label13.Text) + " AND almac_21s =" + alm + "order by nomb_21s", conx)
            conn.Fill(ty2)
            ComboBox1.DataSource = ty2
            ComboBox1.DisplayMember = "nomb_21s"
            ComboBox1.ValueMember = "dest_21s"
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        DataGridView1.BeginEdit(True)
        DataGridView1.Columns(0).ReadOnly = True
        DataGridView1.Columns(1).ReadOnly = True

        If e.ColumnIndex >= 1 And e.ColumnIndex <= 14 Then
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
                         Where A.CCia_tl = '" + Trim(Label13.Text.ToString()) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(e.RowIndex).Cells(15).Value) + "'
                        ORDER BY  A.CCia_tl, A.Codigo_tl"
            Dim cmd102358 As New SqlCommand(sql102358, conx)
            Rsr20 = cmd102358.ExecuteReader()
            If Rsr20.Read() = True Then
                DataGridView1.Columns(4).HeaderText = Trim(Rsr20(13).ToString().Trim())
                DataGridView1.Columns(5).HeaderText = Trim(Rsr20(14).ToString().Trim())
                DataGridView1.Columns(6).HeaderText = Trim(Rsr20(15).ToString().Trim())
                DataGridView1.Columns(7).HeaderText = Trim(Rsr20(16).ToString().Trim())
                DataGridView1.Columns(8).HeaderText = Trim(Rsr20(17).ToString().Trim())
                DataGridView1.Columns(9).HeaderText = Trim(Rsr20(18).ToString().Trim())
                DataGridView1.Columns(10).HeaderText = Trim(Rsr20(19).ToString().Trim())
                DataGridView1.Columns(11).HeaderText = Trim(Rsr20(20).ToString().Trim())
                DataGridView1.Columns(12).HeaderText = Trim(Rsr20(21).ToString().Trim())
                DataGridView1.Columns(13).HeaderText = Trim(Rsr20(22).ToString().Trim())
            End If
            Rsr20.Close()
            ' columna 1
            Dim sql1021 As String = "select  cast( isnull(SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1,cast( isnull(SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_2,
            cast( isnull(SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_3,cast( isnull(SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_4,
            cast( isnull(SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_5,cast( isnull(SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_6,
            cast( isnull(SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_7,cast( isnull(SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_8,
            cast( isnull(SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_9,cast( isnull(SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_10
             from custom_vianny.dbo.mat030f  where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
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
             from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la LIKE '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
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

            TextBox11.Text = (stock1 - jo1).ToString()
            TextBox13.Text = (stock2 - jo2).ToString()
            TextBox14.Text = (stock3 - jo3).ToString()
            TextBox15.Text = (stock4 - jo4).ToString()
            TextBox16.Text = (stock5 - jo5).ToString()
            TextBox17.Text = (stock6 - jo6).ToString()
            TextBox18.Text = (stock7 - jo7).ToString()
            TextBox19.Text = (stock8 - jo8).ToString()
            TextBox20.Text = (stock9 - jo9).ToString()
            TextBox21.Text = (stock10 - jo10).ToString()

            ' fin columna 1
            '////////////////////////////////////////////////////////////////////////////////////////////////


        End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
            I18 = DataGridView1.Rows.Count
            For i1 = 0 To I18 - 1
                VAL = i1 + 1
                Select Case VAL.ToString.Length
                    Case "1" : DataGridView1.Rows(i1).Cells(0).Value = "00" & "" & VAL
                    Case "2" : DataGridView1.Rows(i1).Cells(0).Value = "0" & "" & VAL
                    Case "3" : DataGridView1.Rows(i1).Cells(0).Value = VAL
                End Select
            Next
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add()
        Dim I18, VAL As Integer
        I18 = DataGridView1.Rows.Count
        For i1 = 0 To I18 - 1
            VAL = i1 + 1
            Select Case VAL.ToString.Length
                Case "1" : DataGridView1.Rows(i1).Cells(0).Value = "00" & "" & VAL
                Case "2" : DataGridView1.Rows(i1).Cells(0).Value = "0" & "" & VAL
                Case "3" : DataGridView1.Rows(i1).Cells(0).Value = VAL
            End Select
        Next
    End Sub

    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        If e.ColumnIndex = "4" Then
            valorAnterior1 = Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
        End If
        If e.ColumnIndex = "5" Then
            valorAnterior2 = DataGridView1.Rows(e.RowIndex).Cells(5).Value.ToString()
        End If
        If e.ColumnIndex = "6" Then
            valorAnterior3 = DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString()
        End If
        If e.ColumnIndex = "7" Then
            valorAnterior4 = DataGridView1.Rows(e.RowIndex).Cells(7).Value.ToString()
        End If
        If e.ColumnIndex = "8" Then
            valorAnterior5 = DataGridView1.Rows(e.RowIndex).Cells(8).Value.ToString()
        End If
        If e.ColumnIndex = "9" Then
            valorAnterior6 = DataGridView1.Rows(e.RowIndex).Cells(9).Value.ToString()
        End If
        If e.ColumnIndex = "10" Then
            valorAnterior7 = DataGridView1.Rows(e.RowIndex).Cells(10).Value.ToString()
        End If
        If e.ColumnIndex = "11" Then
            valorAnterior8 = DataGridView1.Rows(e.RowIndex).Cells(11).Value.ToString()
        End If
        If e.ColumnIndex = "12" Then
            valorAnterior9 = DataGridView1.Rows(e.RowIndex).Cells(12).Value.ToString()
        End If
        If e.ColumnIndex = "13" Then
            valorAnterior10 = DataGridView1.Rows(e.RowIndex).Cells(13).Value.ToString()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        Button2.Enabled = False
        DataGridView1.Enabled = False
        DataGridView1.Rows.Clear()

        Button5.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        RadioButton1.Checked = True
        limpiar()
        TextBox8.Enabled = False
        TextBox5.ReadOnly = False
        TextBox4.ReadOnly = False
        TextBox1.ReadOnly = False
        TextBox5.Enabled = True
        TextBox4.Enabled = True
        ComboBox3.Enabled = True
        TextBox27.Text = "1"
        TextBox4.Select()
        ty2.Clear()
        CORRELATIVO()
    End Sub


    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub



    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs) Handles TextBox26.TextChanged

    End Sub
    Dim Rsr205, Rsr2333, Rsr24 As SqlDataReader

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox22.Enabled = True
            TextBox22.Select()
        Else
            TextBox22.Enabled = False
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox26_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox26.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim VAL As Integer
            Select Case Trim(TextBox26.Text).Length
                Case "1" : TextBox26.Text = "01" & "0000000" & TextBox26.Text
                Case "2" : TextBox26.Text = "01" & "000000" & TextBox26.Text
                Case "3" : TextBox26.Text = "01" & "00000" & TextBox26.Text
                Case "4" : TextBox26.Text = "01" & "0000" & TextBox26.Text
                Case "5" : TextBox26.Text = "01" & "000" & TextBox26.Text
                Case "6" : TextBox26.Text = "01" & "00" & TextBox26.Text
                Case "7" : TextBox26.Text = "01" & "0" & TextBox26.Text
                Case "8" : TextBox26.Text = "01" & TextBox26.Text

            End Select

            Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, q.tallador_3 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(TextBox26.Text) + "' and q.ccia ='" + Trim(Label13.Text) + "'
"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            Dim i5 As Integer
            i5 = DataGridView1.Rows.Count

            If Rsr2.Read() = True Then
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i5).Cells(1).Value = Rsr2(0)
                DataGridView1.Rows(i5).Cells(2).Value = Rsr2(1)
                DataGridView1.Rows(i5).Cells(3).Value = Rsr2(2)
                DataGridView1.Rows(i5).Cells(4).Value = 0
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
                DataGridView1.Rows(i5).Cells(15).Value = Rsr2(3)
                DataGridView1.Rows(i5).Cells(0).Value = i5 + 1
                VAL = DataGridView1.Rows(i5).Cells(0).Value
                Select Case VAL.ToString.Length
                    Case "1" : DataGridView1.Rows(i5).Cells(0).Value = "00" & "" & VAL
                    Case "2" : DataGridView1.Rows(i5).Cells(0).Value = "0" & "" & VAL
                    Case "3" : DataGridView1.Rows(i5).Cells(0).Value = VAL
                End Select


                'DataGridView1.Rows(i5).Cells(4).Selected = True
                DataGridView1.CurrentCell = DataGridView1.Rows(i5).Cells(4)
                DataGridView1.BeginEdit(True)
                Rsr2.Close()




                Dim sql1 As String = "select ISNULL( SUM(cant_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) AS Cant_Total,
  ISNULL(SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) AS Cant_1,
  SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_2,
  SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_3,
  SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_4,
  SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_5,
  SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_6,
  SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_7,
  SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_8,
  SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_9,
  SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_10
from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label13.Text) + "'  and talm_3b ='" + Trim(ComboBox3.Text) + "' and linea_3b ='" + Trim(DataGridView1.Rows(i5).Cells(2).Value) + "' and flag_3b ='1'  and cperiodo_3b ='" + Trim(Label11.Text) + "' "
                Dim cmd1 As New SqlCommand(sql1, conx)
                Rsr222 = cmd1.ExecuteReader()
                If Rsr222.Read() = True Then
                    TextBox11.Text = Rsr222(1)
                    TextBox13.Text = Rsr222(2)
                    TextBox14.Text = Rsr222(3)
                    TextBox15.Text = Rsr222(4)
                    TextBox16.Text = Rsr222(5)
                    TextBox17.Text = Rsr222(6)
                    TextBox18.Text = Rsr222(7)
                    TextBox19.Text = Rsr222(8)
                    TextBox20.Text = Rsr222(9)
                    TextBox21.Text = Rsr222(10)

                Else
                    TextBox11.Text = 0
                    TextBox13.Text = 0
                    TextBox14.Text = 0
                    TextBox15.Text = 0
                    TextBox16.Text = 0
                    TextBox17.Text = 0
                    TextBox18.Text = 0
                    TextBox19.Text = 0
                    TextBox20.Text = 0
                    TextBox21.Text = 0

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
                 Where A.CCia_tl = '" + Trim(Label13.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(i5).Cells(15).Value) + "'
                ORDER BY  A.CCia_tl, A.Codigo_tl"
                Dim cmd10235 As New SqlCommand(sql10235, conx)
                Rsr23 = cmd10235.ExecuteReader()
                If Rsr23.Read() = True Then
                    DataGridView1.Columns(4).HeaderText = Trim(Rsr23(13))
                    DataGridView1.Columns(5).HeaderText = Trim(Rsr23(14))
                    DataGridView1.Columns(6).HeaderText = Trim(Rsr23(15))
                    DataGridView1.Columns(7).HeaderText = Trim(Rsr23(16))
                    DataGridView1.Columns(8).HeaderText = Trim(Rsr23(17))
                    DataGridView1.Columns(9).HeaderText = Trim(Rsr23(18))
                    DataGridView1.Columns(10).HeaderText = Trim(Rsr23(19))
                    DataGridView1.Columns(11).HeaderText = Trim(Rsr23(20))
                    DataGridView1.Columns(12).HeaderText = Trim(Rsr23(21))
                    DataGridView1.Columns(13).HeaderText = Trim(Rsr23(22))



                End If
                Rsr23.Close()


                TextBox26.Text = ""

                If Trim(DataGridView1.Columns(4).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(4).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(5).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(5).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(6).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(6).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(7).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(7).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(8).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(8).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(9).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(9).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(10).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(10).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(11).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(11).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(12).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(12).ReadOnly = True
                End If
                If Trim(DataGridView1.Columns(13).HeaderText).Length = 0 Then
                    DataGridView1.Rows(i5).Cells(13).ReadOnly = True
                End If
            Else
                MsgBox("LA OP INGRESADA NO EXISTE")
            End If
            Rsr2.Close()
            TextBox26.Select()
        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub
    Dim dt1, dt2, HG As New DataTable
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox26.Enabled = True
            TextBox26.Select()

            ty2.Clear()
            abrir()
            llenar_combo_box2(ComboBox1, ComboBox3.Text)
            Dim ml As New vnia
            Dim mk As New fnsa
            Dim i As Integer
            i = Len(TextBox5.Text)

            Select Case i
                Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
                Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
                Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
                Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
                Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
                Case "6" : TextBox5.Text = TextBox5.Text

            End Select
            ml.gcomprobante = TextBox4.Text & TextBox5.Text

            ml.galmacen1 = ComboBox3.Text
            ml.gano = Label11.Text
            ml.gccia = Label13.Text
            If mk.mostrar_correlativo_nsa1(ml) = True Then
                Dim jk As New fnsa

                Button4.Enabled = True
                Button6.Enabled = True
                DataGridView1.Rows.Clear()
                dt1 = jk.mostrar_cabecera_nsa(ml)
                dt2 = jk.mostrar_detalle_nsa_pt(ml)
                If dt1.Rows.Count <> 0 Then
                    DataGridView3.DataSource = dt1

                    TextBox9.Text = DataGridView3.Rows(0).Cells(0).Value
                    TextBox8.Text = DataGridView3.Rows(0).Cells(4).Value
                    If Convert.ToString(DataGridView3.Rows(0).Cells(6).Value) Is DBNull.Value Then
                        TextBox10.Text = ""
                    Else
                        TextBox10.Text = Convert.ToString(DataGridView3.Rows(0).Cells(6).Value)
                    End If
                    TextBox22.Text = DataGridView3.Rows(0).Cells(10).Value
                    TextBox23.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView3.Rows(0).Cells(11).Value), 1, 2)
                    TextBox24.Text = Microsoft.VisualBasic.Mid(Trim(DataGridView3.Rows(0).Cells(11).Value), 3, 6)
                    DateTimePicker1.Value = DataGridView3.Rows(0).Cells(1).Value

                    TextBox1.Text = DataGridView3.Rows(0).Cells(2).Value
                    If Trim(DataGridView3.Rows(0).Cells(9).Value) = "0" Then
                        CheckBox1.Checked = False
                    Else
                        CheckBox1.Checked = True
                    End If
                    abrir()
                    enunciado2 = New SqlCommand("select rtrim(ltrim(nomb_21s)) as nomb_21s from custom_vianny.dbo.cag210s where  Empr_21s ='" + Trim(Label13.Text) + "' AND almac_21s='" + Trim(ComboBox3.Text) + "'  and dest_21s='" + Trim(DataGridView3.Rows(0).Cells(5).Value) + "'", conx)
                    respuesta2 = enunciado2.ExecuteReader
                    While respuesta2.Read
                        ComboBox1.Text = respuesta2.Item("nomb_21s")
                    End While
                    respuesta2.Close()

                    If DataGridView3.Rows(0).Cells(7).Value = "EXT" Then
                        ComboBox2.SelectedIndex = 0
                    Else
                        ComboBox2.SelectedIndex = 1
                    End If
                    If Trim(DataGridView3.Rows(0).Cells(11).Value) <> "" Then
                        CheckBox1.Checked = True
                        TextBox22.Enabled = False
                    End If
                End If
                If dt2.Rows.Count <> 0 Then
                    Dim nu1 As Integer
                    DataGridView4.DataSource = dt2
                    nu1 = DataGridView4.Rows.Count
                    DataGridView1.Rows.Add(nu1)
                    For i1 = 0 To nu1 - 1

                        DataGridView1.Rows(i1).Cells(0).Value = DataGridView4.Rows(i1).Cells(5).Value
                        DataGridView1.Rows(i1).Cells(1).Value = DataGridView4.Rows(i1).Cells(8).Value
                        DataGridView1.Rows(i1).Cells(2).Value = DataGridView4.Rows(i1).Cells(6).Value

                        DataGridView1.Rows(i1).Cells(4).Value = DataGridView4.Rows(i1).Cells(10).Value
                        DataGridView1.Rows(i1).Cells(5).Value = DataGridView4.Rows(i1).Cells(11).Value
                        DataGridView1.Rows(i1).Cells(6).Value = DataGridView4.Rows(i1).Cells(12).Value
                        DataGridView1.Rows(i1).Cells(7).Value = DataGridView4.Rows(i1).Cells(13).Value
                        DataGridView1.Rows(i1).Cells(8).Value = DataGridView4.Rows(i1).Cells(14).Value
                        DataGridView1.Rows(i1).Cells(9).Value = DataGridView4.Rows(i1).Cells(15).Value
                        DataGridView1.Rows(i1).Cells(10).Value = DataGridView4.Rows(i1).Cells(16).Value
                        DataGridView1.Rows(i1).Cells(11).Value = DataGridView4.Rows(i1).Cells(17).Value
                        DataGridView1.Rows(i1).Cells(12).Value = DataGridView4.Rows(i1).Cells(18).Value
                        DataGridView1.Rows(i1).Cells(13).Value = DataGridView4.Rows(i1).Cells(19).Value
                        DataGridView1.Rows(i1).Cells(14).Value = DataGridView4.Rows(i1).Cells(9).Value

                        Dim sql102 As String = "select ncom_3,linea_3,a.nomb_17, q.tallador_3 from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.CAG1700 A on q.ccia = a.ccia and q.linea_3 = a.linea_17
where ncom_3 ='" + Trim(DataGridView1.Rows(i1).Cells(1).Value) + "' and q.ccia ='" + Trim(Label13.Text) + "'
"
                        Dim cmd102 As New SqlCommand(sql102, conx)
                        Rsr2 = cmd102.ExecuteReader()


                        If Rsr2.Read() = True Then
                            DataGridView1.Rows(i1).Cells(3).Value = Rsr2(2)
                            DataGridView1.Rows(i1).Cells(15).Value = Rsr2(3)
                        End If


                        Rsr2.Close()


                    Next


                    DataGridView1.Columns(0).Frozen = True
                    DataGridView1.Columns(1).Frozen = True
                    DataGridView1.Columns(2).Frozen = True
                End If

                TextBox9.Enabled = False
                TextBox26.Enabled = False
                DataGridView1.Enabled = False
                DateTimePicker1.Enabled = False
                ComboBox1.Enabled = False
                Button1.Enabled = False
                Button2.Enabled = False
                CheckBox1.Enabled = False
                Dim kl, sum12 As Integer

                kl = DataGridView1.Rows.Count

                For ol = 0 To kl - 1
                    sum12 = sum12 + Convert.ToInt32(DataGridView1.Rows(ol).Cells(14).Value)
                Next
                TextBox38.Text = sum12

                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                ComboBox3.Enabled = False
                ComboBox4.Enabled = False
                TextBox8.Enabled = False
                TextBox9.Enabled = False
                TextBox1.Enabled = False
                TextBox1.ReadOnly = False
                Button2.Enabled = False
                DataGridView1.Enabled = False
                Button1.Enabled = False
                Button5.Enabled = False
                Button7.Enabled = Enabled
                Button3.Enabled = False
                TextBox9.Enabled = False
                TextBox1.Enabled = False
                TextBox8.Enabled = False
                TextBox10.Enabled = False
                TextBox12.ReadOnly = False
            Else
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox3.Enabled = False
                ComboBox4.Enabled = True
                TextBox8.Enabled = True
                TextBox9.Enabled = True
                TextBox1.Enabled = True
                TextBox1.ReadOnly = False
                Button2.Enabled = True
                DataGridView1.Enabled = True
                Button1.Enabled = True
                Button5.Enabled = True
                Button3.Enabled = True
                DataGridView1.Rows.Clear()
                TextBox9.Text = ""
                TextBox1.Text = ""
                TextBox8.Text = ""
                TextBox10.Text = ""
                TextBox12.ReadOnly = False
                Button7.Enabled = Enabled
            End If
            ComboBox3.Enabled = False
            ComboBox4.SelectedIndex = -1
            DateTimePicker1.Select()
            CheckBox1.Enabled = True
        End If
    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        pTERMINADO.TextBox1.Text = Label13.Text
        pTERMINADO.TextBox2.Text = Label11.Text
        pTERMINADO.TextBox3.Text = ComboBox3.Text
        pTERMINADO.TextBox4.Text = "2"
        pTERMINADO.TextBox5.Text = Trim(TextBox4.Text) & Trim(TextBox5.Text)
        pTERMINADO.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If RadioButton2.Checked = True Then
            MsgBox("LA NOTA DE INGRESO ESTA ANULADA NO SE PUEDE EDITAR")
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
                    TextBox13.Enabled = True
                    DataGridView1.Enabled = True
                    DataGridView1.Columns(7).ReadOnly = False

                    TextBox9.Enabled = True
                    DateTimePicker1.Enabled = True
                    'TextBox12.Enabled = True
                    TextBox12.ReadOnly = False
                    TextBox4.Enabled = False
                    TextBox20.Enabled = True
                    TextBox5.Enabled = False
                    Button2.Enabled = True
                    Button5.Enabled = True
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    ComboBox4.Enabled = True
                    TextBox16.Enabled = True
                    TextBox26.Enabled = True
                    TextBox17.Enabled = True

                    Button1.Enabled = True
                    Button3.Enabled = True
                    TextBox8.Enabled = True
                    TextBox27.Text = "2"
                Else
                    MsgBox("DIA " & DateTimePicker1.Value & " CERRADO,NO SE PUEDE EDITAR LA INFORMACION - CONSULTE CON EL AREA CONTABLE ")
                End If

            End If
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
    Public Sub limpiar()
        TextBox8.Text = ""
        TextBox10.Text = ""
        TextBox9.Text = ""
        TextBox11.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox26.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        TextBox12.Text = ""
        TextBox38.Text = ""
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating

    End Sub


    Private Sub TextBox27_TextChanged(sender As Object, e As EventArgs) Handles TextBox27.TextChanged

    End Sub

    Dim da As New DataTable

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        MsgBox("ESTA FUNCION NO ESTA HABILITADA")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim camposFaltantes As New List(Of String) ' Lista para almacenar los nombres de los campos faltantes

        ' Validar campos obligatorios
        If String.IsNullOrEmpty(TextBox8.Text) Then
            camposFaltantes.Add(" RUC ")
        End If

        If String.IsNullOrEmpty(TextBox4.Text) Then
            camposFaltantes.Add(" MES ")
        End If

        If String.IsNullOrEmpty(TextBox4.Text) Then
            camposFaltantes.Add(" CORRELATIVO ")
        End If

        'Agregar más validaciones para otros campos obligatorios según sea necesario

        ' Verificar si hay campos faltantes
        If camposFaltantes.Count > 0 Then
            ' Mostrar mensaje de error indicando los campos faltantes
            Dim camposFaltantesStr As String = String.Join(", ", camposFaltantes)
            MessageBox.Show("Falta ingresar información en los siguientes campos: " & camposFaltantesStr, "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            abrir()
            da.Clear()
            Dim cmd17 As New SqlCommand("IF NOT EXISTS( SELECT NCom_3 FROM custom_vianny.DBO.Mag030f WHERE CCia_3 = @CCIA_3 AND CPeriodo_3 = @CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 = @ccom_3 AND NCom_3 = @ncom_3)
        INSERT INTO custom_vianny.DBO.MAG030F (CCIA_3 , CPERIODO_3, talm_3, ccom_3, ncom_3, fcom_3, gloa_3, glob_3,fich_3 , orig_3, nombe_3, adevol_3, transf_3, TDoc_3, grm_3, fase_3, flag_3, origen_3, usuario_3, fecha_3, hora_3, subfase_3, tidoc_3,sfactu_3, nfactu_3, transfer_3, pedorig_3, Auto_3,orides_3,motdes_3) 
                     VALUES (@CCIA_3,@CPERIODO_3,@TAlm_3,@ccom_3,@ncom_3,@fcom_3,@gloa_3,'',@fich_3,@orig_3, '', 0, 0,@TDoc_3,@grm_3,'',1,@origen_3,@usuario_3,@fecha_3,@hora_3, '', '', '', '', 0, '', 0,@orides_3,@motdes_3)
        				ELSE
        					UPDATE custom_vianny.DBO.MAG030F SET fcom_3  = @fcom_3, 
        					                   gloa_3  = @gloa_3, 
        					                   glob_3  = '', 
        					                   fich_3  = @fich_3, 
        					                   orig_3  = @orig_3, 
        					                   nombe_3 = '', 
        					                   adevol_3  = 0, 
        					                   transf_3  = 0, 
        					                   TDoc_3    =@TDoc_3,
        					                   grm_3     = @grm_3, 
        					                   fase_3    = '',
        					                   flag_3    = 1, 
        					                   origen_3  = @origen_3, 
        					                   usuario_3 = @usuario_3, 
        					                   fecha_3   =@fecha_3, 
        					                   hora_3    = @hora_3, 
        					                   subfase_3 = '', 
        					                   tidoc_3   = '', 
        					                   sfactu_3  = '', 
        					                   nfactu_3  = '', 
        					                   transfer_3= 0, 
        					                   pedorig_3 = '',
        					                   Auto_3 = 0
        									   WHERE CCIA_3 = @CCIA_3 And CPERIODO_3 =@CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 =@ccom_3 And ncom_3 = @ncom_3", conx)
            cmd17.Parameters.AddWithValue("@CCIA_3", Trim(Label13.Text))
            cmd17.Parameters.AddWithValue("@CPERIODO_3", Trim(Label11.Text))
            cmd17.Parameters.AddWithValue("@TAlm_3", Trim(ComboBox3.Text))
            cmd17.Parameters.AddWithValue("@ccom_3", "2")
            cmd17.Parameters.AddWithValue("@ncom_3", Trim(TextBox4.Text) & Trim(TextBox5.Text))
            cmd17.Parameters.AddWithValue("@fcom_3", DateTimePicker1.Value.Date)
            cmd17.Parameters.AddWithValue("@gloa_3", Trim(TextBox9.Text))
            cmd17.Parameters.AddWithValue("@orig_3", Trim(ComboBox1.SelectedValue.ToString))
            cmd17.Parameters.AddWithValue("@fich_3", Trim(TextBox8.Text))
            cmd17.Parameters.AddWithValue("orides_3", Trim(TextBox22.Text))
            cmd17.Parameters.AddWithValue("motdes_3", Trim(TextBox23.Text) & Trim(TextBox24.Text))
            If Trim(ComboBox2.Text) = "INTERNA" Then
                cmd17.Parameters.AddWithValue("@origen_3", "INT")
            Else
                cmd17.Parameters.AddWithValue("@origen_3", "EXT")
            End If
            Dim doc As String
            Select Case ComboBox4.Text
                Case "GUIA" : doc = "009"
                Case "FACT" : doc = "001"
                Case "OTRO" : doc = "002"
                Case "" : doc = ""
            End Select
            cmd17.Parameters.AddWithValue("@TDoc_3", Trim(doc))
            cmd17.Parameters.AddWithValue("@grm_3", Trim(TextBox12.Text))
            cmd17.Parameters.AddWithValue("@usuario_3", Trim(TextBox7.Text))
            cmd17.Parameters.AddWithValue("@fecha_3", DateTimePicker1.Value)
            cmd17.Parameters.AddWithValue("@hora_3", DateTime.Now.ToString("hh:mm:ss"))

            cmd17.ExecuteNonQuery()


            If Trim(TextBox27.Text) = "2" Then


                Dim cmd As New SqlDataAdapter("SELECT ccia_3b,cperiodo_3b,linea_3b, pedido_3b,talm_3b,cant_3b,cant1_3b,cant2_3b,cant3_3b,cant4_3b,cant5_3b,cant6_3b,cant7_3b,cant8_3b,cant9_3b,cant10_3b  FROM custom_vianny.DBO.Mat030f A 
		       		   Where A.CCia_3b = '" + Trim(Label13.Text) + "' AND A.CPeriodo_3b = '" + Trim(Label11.Text) + "' AND A.TAlm_3b = '" + Trim(ComboBox3.Text) + "' AND A.CCom_3b = '2' AND A.NCom_3b = '" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "'", conx)
                cmd.Fill(da)
                DataGridView11.DataSource = da


            End If
            Dim agregar As String = "DELETE custom_vianny.DBO.Map030f Where CCIA_3a = '" + Trim(Label13.Text) + "' And CPERIODO_3a = '" + Trim(Label11.Text) + "' And talm_3a = '" + Trim(ComboBox3.Text) + "' And ccom_3a = '2' And ncom_3a = '" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "'"
            Dim agregar1 As String = "DELETE custom_vianny.DBO.Mat030f Where CCIA_3b = '" + Trim(Label13.Text) + "' And CPERIODO_3b = '" + Trim(Label11.Text) + "' And talm_3b = '" + Trim(ComboBox3.Text) + "' And ccom_3b = '2' And ncom_3b = '" + Trim(TextBox4.Text) & Trim(TextBox5.Text) + "'"
            ELIMINAR(agregar)
            ELIMINAR(agregar1)
            Dim p As Integer
            p = DataGridView1.Rows.Count

            For i = 0 To p - 1
                Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.DBO.Mat030f (ccia_3b , cperiodo_3b, talm_3b, ccom_3b, ncom_3b, item_3b,linea_3b ,pedido_3b ,cant_3b ,cant1_3b ,cant2_3b ,cant3_3b ,cant4_3b ,cant5_3b ,cant6_3b ,cant7_3b ,cant8_3b ,cant9_3b ,cant10_3b,ordens_3b,flag_3b)
                        VALUES (@ccia_3b,@cperiodo_3b,@talm_3b,@ccom_3b,@ncom_3b,@item_3b,@linea_3b,@pedido_3b,@cant_3b,@cant1_3b,@cant2_3b,@cant3_3b,@cant4_3b,@cant5_3b,@cant6_3b,@cant7_3b,@cant8_3b,@cant9_3b,@cant10_3b,'', 1)", conx)
                cmd15.Parameters.AddWithValue("@ccia_3b", Trim(Label13.Text))
                cmd15.Parameters.AddWithValue("@cperiodo_3b", Trim(Label11.Text))
                cmd15.Parameters.AddWithValue("@talm_3b", Trim(ComboBox3.Text))
                cmd15.Parameters.AddWithValue("@ccom_3b", 2)
                cmd15.Parameters.AddWithValue("@ncom_3b", Trim(TextBox4.Text) & Trim(TextBox5.Text))
                cmd15.Parameters.AddWithValue("@item_3b", Trim(DataGridView1.Rows(i).Cells(0).Value))
                cmd15.Parameters.AddWithValue("@linea_3b", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd15.Parameters.AddWithValue("@pedido_3b", Trim(DataGridView1.Rows(i).Cells(1).Value))
                cmd15.Parameters.AddWithValue("@cant_3b", DataGridView1.Rows(i).Cells(14).Value)
                cmd15.Parameters.AddWithValue("@cant1_3b", DataGridView1.Rows(i).Cells(4).Value)
                cmd15.Parameters.AddWithValue("@cant2_3b", DataGridView1.Rows(i).Cells(5).Value)
                cmd15.Parameters.AddWithValue("@cant3_3b", DataGridView1.Rows(i).Cells(6).Value)
                cmd15.Parameters.AddWithValue("@cant4_3b", DataGridView1.Rows(i).Cells(7).Value)
                cmd15.Parameters.AddWithValue("@cant5_3b", DataGridView1.Rows(i).Cells(8).Value)
                cmd15.Parameters.AddWithValue("@cant6_3b", DataGridView1.Rows(i).Cells(9).Value)
                cmd15.Parameters.AddWithValue("@cant7_3b", DataGridView1.Rows(i).Cells(10).Value)
                cmd15.Parameters.AddWithValue("@cant8_3b", DataGridView1.Rows(i).Cells(11).Value)
                cmd15.Parameters.AddWithValue("@cant9_3b", DataGridView1.Rows(i).Cells(12).Value)
                cmd15.Parameters.AddWithValue("@cant10_3b", DataGridView1.Rows(i).Cells(13).Value)
                cmd15.ExecuteNonQuery()

                For y = 4 To 13
                    If DataGridView1.Rows(i).Cells(y).Value > 0 Then
                        Dim cmd16 As New SqlCommand("INSERT INTO custom_vianny.DBO.MAP030F (CCIA_3A , CPERIODO_3A, talm_3a, ccom_3a, ncom_3a, item_3a,linea_3a ,sinon_3a,linea2_3a, cant_3a, unid_3a, cantu_3a, obser1_3a, obser2_3a, obser3_3a, cantk_3a, cantke_3a, pedido_3a, flag_3a, peso_3a,tpeso_3a,unidk_3a , pieza_3a, color_3a, ncolor_3a, parti_3a, ocorte_3a, ordens_3a, nrollo_3a, kgrollo_3a, lote_3a, maquina_3a, galga_3a,sigla_3a, sitcal_3a, cantc_3a, ancho_3a, densid_3a, pedod_3a, ordtej_3a  , tipapt_3a , combo_3a,ccosto_3a,PreUn1_3a,Tot1_3a  )
                           Values (@CCIA_3A,@CPERIODO_3A,@talm_3a,@ccom_3a,@ncom_3aa,@item_3a,@linea_3a, ''     ,'PT0101' ,@cantid ,@unid_3a,  @cantid,''        , ''       ,''        ,@cantid  , 0.00     ,@pedido_3a, 1      , 0.00000, 0.00   ,@unidk_3a,   ''    , ''      , ''       , ''      , ''       , ''       , ''       , 0.00      , ''     , ''        , 0.00    ,''      , 0        , 0.000   , 0.000   , 0.000    , ''      , ''         , ''        , ''      ,''       ,0.00000  , 0.00)", conx)
                        cmd16.Parameters.AddWithValue("@CCIA_3A", Trim(Label13.Text))
                        cmd16.Parameters.AddWithValue("@CPERIODO_3A", Trim(Label11.Text))
                        cmd16.Parameters.AddWithValue("@talm_3a", Trim(ComboBox3.Text))
                        cmd16.Parameters.AddWithValue("@ccom_3a", "2")
                        cmd16.Parameters.AddWithValue("@ncom_3aa", Trim(TextBox4.Text) & Trim(TextBox5.Text))
                        cmd16.Parameters.AddWithValue("@item_3a", Trim(DataGridView1.Rows(i).Cells(0).Value))
                        cmd16.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
                        cmd16.Parameters.AddWithValue("@cantid", DataGridView1.Rows(i).Cells(y).Value)
                        cmd16.Parameters.AddWithValue("@unid_3a", "UND")
                        cmd16.Parameters.AddWithValue("@pedido_3a", Trim(DataGridView1.Rows(i).Cells(1).Value))
                        cmd16.Parameters.AddWithValue("@unidk_3a", "UND")
                        cmd16.ExecuteNonQuery()
                    End If

                Next
            Next

            MsgBox("SE REGISTRO LA INFORMACION CORRECTAMENTE")

            Dim respuesta As DialogResult

            respuesta = MessageBox.Show("DESEA IMPRIMIR LA NOTA DE SALIDA?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                pTERMINADO.TextBox1.Text = Label13.Text
                pTERMINADO.TextBox2.Text = Label11.Text
                pTERMINADO.TextBox3.Text = ComboBox3.Text
                pTERMINADO.TextBox4.Text = "2"
                pTERMINADO.TextBox5.Text = Trim(TextBox4.Text) & Trim(TextBox5.Text)
                pTERMINADO.Show()
            End If
            'If CheckBox1.Checked = True Then
            '    registro_nia_automatica()
            'End If
            TextBox27.Text = "1"
            DataGridView11.DataSource = Nothing
            CheckBox1.Checked = False
            TextBox22.Enabled = False
            BLOQUEARC()
            limpiar()
            CORRELATIVO()
        End If
    End Sub
    Private Sub BLOQUEARC()
        TextBox13.Enabled = False
        DataGridView1.Enabled = False
        DataGridView1.Columns(7).ReadOnly = False

        TextBox9.Enabled = False
        DateTimePicker1.Enabled = False
        'TextBox12.Enabled = True
        TextBox12.ReadOnly = False

        TextBox20.Enabled = False

        Button2.Enabled = False
        Button5.Enabled = False
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox4.Enabled = False
        TextBox16.Enabled = False
        TextBox26.Enabled = False
        TextBox17.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button6.Enabled = False
        Button1.Enabled = False
        TextBox8.Enabled = False
    End Sub

    '    Sub registro_nia_automatica()
    '        abrir()
    '        Dim agregar As String = "DELETE custom_vianny.DBO.Map030f Where CCIA_3a = '" + Trim(Label13.Text) + "' And CPERIODO_3a = '" + Trim(Label11.Text) + "' And talm_3a = '" + TextBox22.Text + "' And ccom_3a = '1' And ncom_3a = '" + Trim(TextBox23.Text) & Trim(TextBox24.Text) + "'"
    '        Dim agregar1 As String = "DELETE custom_vianny.DBO.Mat030f Where CCIA_3b = '" + Trim(Label13.Text) + "' And CPERIODO_3b = '" + Trim(Label11.Text) + "' And talm_3b = '" + TextBox22.Text + "' And ccom_3b = '1' And ncom_3b = '" + Trim(TextBox23.Text) & Trim(TextBox24.Text) + "'"
    '        ELIMINAR(agregar)
    '        ELIMINAR(agregar1)



    '        Dim cmd17 As New SqlCommand("IF NOT EXISTS( SELECT NCom_3 FROM custom_vianny.DBO.Mag030f WHERE CCia_3 = @CCIA_3 AND CPeriodo_3 = @CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 = @ccom_3 AND NCom_3 = @ncom_3)
    'INSERT INTO custom_vianny.DBO.MAG030F (CCIA_3 , CPERIODO_3, talm_3, ccom_3, ncom_3, fcom_3, gloa_3, glob_3,fich_3 , orig_3, nombe_3, adevol_3, transf_3, TDoc_3, grm_3, fase_3, flag_3, origen_3, usuario_3, fecha_3, hora_3, subfase_3, tidoc_3,sfactu_3, nfactu_3, transfer_3, pedorig_3, Auto_3) 
    '             VALUES (@CCIA_3,@CPERIODO_3,@TAlm_3,@ccom_3,@ncom_3,@fcom_3,@gloa_3,'',@fich_3,@orig_3, '', 0, 0,@TDoc_3,@grm_3,'',1,@origen_3,@usuario_3,@fecha_3,@hora_3, '', '', '', '', 0, '', 0)
    '				ELSE
    '					UPDATE custom_vianny.DBO.MAG030F SET fcom_3  = @fcom_3, 
    '					                   gloa_3  = @gloa_3, 
    '					                   glob_3  = '', 
    '					                   fich_3  = @fich_3, 
    '					                   orig_3  = @orig_3, 
    '					                   nombe_3 = '', 
    '					                   adevol_3  = 0, 
    '					                   transf_3  = 0, 
    '					                   TDoc_3    =@TDoc_3,
    '					                   grm_3     = @grm_3, 
    '					                   fase_3    = '',
    '					                   flag_3    = 1, 
    '					                   origen_3  = @origen_3, 
    '					                   usuario_3 = @usuario_3, 
    '					                   fecha_3   =@fecha_3, 
    '					                   hora_3    = @hora_3, 
    '					                   subfase_3 = '', 
    '					                   tidoc_3   = '', 
    '					                   sfactu_3  = '', 
    '					                   nfactu_3  = '', 
    '					                   transfer_3= 0, 
    '					                   pedorig_3 = '',
    '					                   Auto_3 = 0
    '									   WHERE CCIA_3 = @CCIA_3 And CPERIODO_3 =@CPERIODO_3 AND TAlm_3 = @TAlm_3 AND CCom_3 =@ccom_3 And ncom_3 = @ncom_3", conx)
    '        cmd17.Parameters.AddWithValue("@CCIA_3", Trim(Label13.Text))
    '        cmd17.Parameters.AddWithValue("@CPERIODO_3", Trim(Label11.Text))
    '        cmd17.Parameters.AddWithValue("@TAlm_3", Trim(TextBox22.Text))
    '        cmd17.Parameters.AddWithValue("@ccom_3", "1")
    '        cmd17.Parameters.AddWithValue("@ncom_3", Trim(TextBox23.Text) & Trim(TextBox24.Text))
    '        cmd17.Parameters.AddWithValue("@fcom_3", DateTimePicker1.Value.Date)
    '        cmd17.Parameters.AddWithValue("@gloa_3", Trim(TextBox9.Text))
    '        cmd17.Parameters.AddWithValue("@orig_3", "0001")
    '        cmd17.Parameters.AddWithValue("@fich_3", Trim(TextBox8.Text))
    '        If Trim(ComboBox2.Text) = "INTERNA" Then
    '            cmd17.Parameters.AddWithValue("@origen_3", "INT")
    '        Else
    '            cmd17.Parameters.AddWithValue("@origen_3", "EXT")
    '        End If
    '        Dim doc As String
    '        Select Case ComboBox4.Text
    '            Case "GUIA" : doc = "009"
    '            Case "FACT" : doc = "001"
    '            Case "OTRO" : doc = "002"
    '            Case "" : doc = ""
    '        End Select
    '        cmd17.Parameters.AddWithValue("@TDoc_3", Trim(doc))
    '        cmd17.Parameters.AddWithValue("@grm_3", Trim(TextBox12.Text))
    '        cmd17.Parameters.AddWithValue("@usuario_3", Trim(TextBox7.Text))
    '        cmd17.Parameters.AddWithValue("@fecha_3", DateTimePicker1.Value.Date)
    '        cmd17.Parameters.AddWithValue("@hora_3", DateTime.Now.ToString("hh:mm:ss"))

    '        cmd17.ExecuteNonQuery()


    '        Dim p As Integer
    '        p = DataGridView1.Rows.Count

    '        For i = 0 To p - 1
    '            Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.DBO.Mat030f (ccia_3b , cperiodo_3b, talm_3b, ccom_3b, ncom_3b, item_3b,linea_3b ,pedido_3b ,cant_3b ,cant1_3b ,cant2_3b ,cant3_3b ,cant4_3b ,cant5_3b ,cant6_3b ,cant7_3b ,cant8_3b ,cant9_3b ,cant10_3b,ordens_3b,flag_3b)
    '             VALUES (@ccia_3b,@cperiodo_3b,@talm_3b,@ccom_3b,@ncom_3b,@item_3b,@linea_3b,@pedido_3b,@cant_3b,@cant1_3b,@cant2_3b,@cant3_3b,@cant4_3b,@cant5_3b,@cant6_3b,@cant7_3b,@cant8_3b,@cant9_3b,@cant10_3b,'', 1)", conx)
    '            cmd15.Parameters.AddWithValue("@ccia_3b", Trim(Label13.Text))
    '            cmd15.Parameters.AddWithValue("@cperiodo_3b", Trim(Label11.Text))
    '            cmd15.Parameters.AddWithValue("@talm_3b", Trim(TextBox22.Text))
    '            cmd15.Parameters.AddWithValue("@ccom_3b", 1)
    '            cmd15.Parameters.AddWithValue("@ncom_3b", Trim(TextBox23.Text) & Trim(TextBox24.Text))
    '            cmd15.Parameters.AddWithValue("@item_3b", Trim(DataGridView1.Rows(i).Cells(0).Value))
    '            cmd15.Parameters.AddWithValue("@linea_3b", Trim(DataGridView1.Rows(i).Cells(2).Value))
    '            cmd15.Parameters.AddWithValue("@pedido_3b", Trim(DataGridView1.Rows(i).Cells(1).Value))
    '            cmd15.Parameters.AddWithValue("@cant_3b", DataGridView1.Rows(i).Cells(14).Value)
    '            cmd15.Parameters.AddWithValue("@cant1_3b", DataGridView1.Rows(i).Cells(4).Value)
    '            cmd15.Parameters.AddWithValue("@cant2_3b", DataGridView1.Rows(i).Cells(5).Value)
    '            cmd15.Parameters.AddWithValue("@cant3_3b", DataGridView1.Rows(i).Cells(6).Value)
    '            cmd15.Parameters.AddWithValue("@cant4_3b", DataGridView1.Rows(i).Cells(7).Value)
    '            cmd15.Parameters.AddWithValue("@cant5_3b", DataGridView1.Rows(i).Cells(8).Value)
    '            cmd15.Parameters.AddWithValue("@cant6_3b", DataGridView1.Rows(i).Cells(9).Value)
    '            cmd15.Parameters.AddWithValue("@cant7_3b", DataGridView1.Rows(i).Cells(10).Value)
    '            cmd15.Parameters.AddWithValue("@cant8_3b", DataGridView1.Rows(i).Cells(11).Value)
    '            cmd15.Parameters.AddWithValue("@cant9_3b", DataGridView1.Rows(i).Cells(12).Value)
    '            cmd15.Parameters.AddWithValue("@cant10_3b", DataGridView1.Rows(i).Cells(13).Value)
    '            cmd15.ExecuteNonQuery()


    '            Dim cmd18 As New SqlCommand("IF LEN(@Codigo_la) > 0
    '						BEGIN
    '						IF @ccom = '1' OR @ccom = '3'
    '								BEGIN
    '								IF NOT EXISTS ( SELECT Codigo_la FROM custom_vianny.DBO.Cag1700_Almac_Lotes WHERE CCia_la = @CCia_la AND Almac_la = @Almac_la AND Codigo_la = @Codigo_la AND Lote_la =@op )
    '									INSERT INTO custom_vianny.DBO.Cag1700_Almac_Lotes (CCia_la, Almac_la, Codigo_la, Lote_la, Stock_la, Stock1_la, Stock2_la, Stock3_la, Stock4_la, Stock5_la, Stock6_la, Stock7_la, Stock8_la, Stock9_la, Stock10_la)
    '															                 VALUES (@CCia_la,@Almac_la,@Codigo_la  ,@op   , @stock  , @stock1  , @stock2  , @stock3  , @stock4  , @stock5  , @stock6  , @stock7  , @stock8  , @stock9  ,@stock10)
    '								ELSE
    '									UPDATE custom_vianny.DBO.Cag1700_Almac_Lotes SET Stock_la = Stock_la +  @stock,
    '																   Stock1_la = Stock1_la + @stock1,
    '																   Stock2_la = Stock2_la + @stock2,
    '																   Stock3_la = Stock3_la + @stock3,
    '																   Stock4_la = Stock4_la +@stock4,
    '																   Stock5_la = Stock5_la + @stock5,
    '																   Stock6_la = Stock6_la + @stock6,
    '																   Stock7_la = Stock7_la + @stock7,
    '																   Stock8_la = Stock8_la + @stock8,
    '																   Stock9_la = Stock9_la + @stock9,
    '																   Stock10_la = Stock10_la + @stock10
    '														   WHERE CCia_la = @CCia_la AND Almac_la = @Almac_la AND Codigo_la = @Codigo_la AND Lote_la = @op
    '                                    UPDATE custom_vianny.DBO.Cag1700_Lotes SET Stock_lt = Stock_lt +  @stock,
    '														     Stock1_lt = Stock1_lt + @stock1,
    '														     Stock2_lt = Stock2_lt + @stock2,
    '														     Stock3_lt = Stock3_lt + @stock3,
    '														     Stock4_lt = Stock4_lt + @stock4,
    '														     Stock5_lt = Stock5_lt + @stock5,
    '														     Stock6_lt = Stock6_lt + @stock6,
    '														     Stock7_lt = Stock7_lt + @stock7,
    '														     Stock8_lt = Stock8_lt + @stock8,
    '														     Stock9_lt = Stock8_lt + @stock9,
    '														     Stock10_lt = Stock10_lt + @stock10
    '													   	     WHERE CCia_lt = @CCia_la AND Codigo_lt = @Codigo_la AND Lote_lt = @op

    '								END
    '						ELSE
    '								BEGIN 
    '								IF NOT EXISTS ( SELECT Codigo_la FROM custom_vianny.DBO.Cag1700_Almac_Lotes WHERE CCia_la = @CCia_la AND Almac_la = @Almac_la AND Codigo_la =@Codigo_la AND Lote_la = @op )
    '									INSERT INTO custom_vianny.DBO.Cag1700_Almac_Lotes (CCia_la, Almac_la, Codigo_la, Lote_la, Stock_la, Stock1_la, Stock2_la, Stock3_la, Stock4_la, Stock5_la, Stock6_la, Stock7_la, Stock8_la, Stock9_la, Stock10_la)
    '															 VALUES (@CCia_la, @Almac_la, @Codigo_la, @op, @stock  , @stock1  , @stock2  , @stock3  , @stock4  , @stock5  , @stock6  , @stock7  , @stock8  , @stock9  ,@stock10)
    '								ELSE
    '									UPDATE custom_vianny.DBO.Cag1700_Almac_Lotes SET Stock_la = Stock_la - @stock,
    '																   Stock1_la = Stock1_la - @stock1,
    '																   Stock2_la = Stock2_la - @stock2,
    '																   Stock3_la = Stock3_la - @stock3,
    '																   Stock4_la = Stock4_la - @stock4,
    '																   Stock5_la = Stock5_la - @stock5,
    '																   Stock6_la = Stock6_la - @stock6,
    '																   Stock7_la = Stock7_la - @stock7,
    '																   Stock8_la = Stock8_la - @stock8,
    '																   Stock9_la = Stock9_la - @stock9,
    '																   Stock10_la = Stock10_la - @stock10
    '														   WHERE CCia_la =@CCia_la AND Almac_la =@Almac_la AND Codigo_la =@Codigo_la AND Lote_la =@op
    '									UPDATE Cag1700_Lotes SET Stock_lt = Stock_lt - @stock,
    '														 Stock1_lt = Stock1_lt - @stock1,
    '														 Stock2_lt = Stock2_lt - @stock2,
    '														 Stock3_lt = Stock3_lt - @stock3,
    '														 Stock4_lt = Stock4_lt - @stock4,
    '														 Stock5_lt = Stock5_lt - @stock5,
    '														 Stock6_lt = Stock6_lt - @stock6,
    '														 Stock7_lt = Stock7_lt - @stock7,
    '														 Stock8_lt = Stock8_lt - @stock8,
    '														 Stock9_lt = Stock8_lt - @stock9,
    '														 Stock10_lt = Stock10_lt - @stock1
    '													   WHERE CCia_lt = @CCia_la AND Codigo_lt = @Codigo_la AND Lote_lt = @op		   

    '								END
    '						END", conx)
    '            cmd18.Parameters.AddWithValue("@Codigo_la", Trim(DataGridView1.Rows(i).Cells(2).Value))
    '            cmd18.Parameters.AddWithValue("@ccom", "1")
    '            cmd18.Parameters.AddWithValue("@CCia_la", Trim(Label13.Text))
    '            cmd18.Parameters.AddWithValue("@Almac_la", Trim(TextBox22.Text))
    '            cmd18.Parameters.AddWithValue("@op", Trim(DataGridView1.Rows(i).Cells(1).Value))
    '            cmd18.Parameters.AddWithValue("@stock", DataGridView1.Rows(i).Cells(14).Value)
    '            cmd18.Parameters.AddWithValue("@stock1", DataGridView1.Rows(i).Cells(4).Value)
    '            cmd18.Parameters.AddWithValue("@stock2", DataGridView1.Rows(i).Cells(5).Value)
    '            cmd18.Parameters.AddWithValue("@stock3", DataGridView1.Rows(i).Cells(6).Value)
    '            cmd18.Parameters.AddWithValue("@stock4", DataGridView1.Rows(i).Cells(7).Value)
    '            cmd18.Parameters.AddWithValue("@stock5", DataGridView1.Rows(i).Cells(8).Value)
    '            cmd18.Parameters.AddWithValue("@stock6", DataGridView1.Rows(i).Cells(9).Value)
    '            cmd18.Parameters.AddWithValue("@stock7", DataGridView1.Rows(i).Cells(10).Value)
    '            cmd18.Parameters.AddWithValue("@stock8", DataGridView1.Rows(i).Cells(11).Value)
    '            cmd18.Parameters.AddWithValue("@stock9", DataGridView1.Rows(i).Cells(12).Value)
    '            cmd18.Parameters.AddWithValue("@stock10", DataGridView1.Rows(i).Cells(13).Value)
    '            cmd18.ExecuteNonQuery()


    '            Dim cmd19 As New SqlCommand("IF LEN(@Codigo_la) > 0
    '					BEGIN
    '					IF @ccom  = '1' OR @ccom = '3'
    '						BEGIN
    '							IF NOT EXISTS ( SELECT Codigo_st FROM custom_vianny.DBO.Cag1700_Stock WHERE CCia_st = @CCia_la AND Almac_st = @Almac_la AND Codigo_st = @Codigo_la)
    '								INSERT INTO  custom_vianny.DBO.Cag1700_Stock (CCia_st,
    '														   Almac_st,
    '														   Codigo_st,
    '														   Stock_st,
    '														   Stock1_st, Stock2_st, Stock3_st, Stock4_st, Stock5_st, Stock6_st, Stock7_st, Stock8_st, Stock9_st, Stock10_st)
    '												  VALUES (@CCia_la,
    '														  @Almac_la,
    '														 @Codigo_la,
    '														  @stock,
    '														  @stock1, @stock2, @stock3, @stock4, @stock5,@stock6, @stock7, @stock8, @stock9,@stock10)
    '							ELSE
    '								UPDATE custom_vianny.DBO.Cag1700_Stock SET Stock_st  = Stock_st + @stock,
    '														 Stock1_st = Stock1_st + @stock1,
    '														 Stock2_st = Stock2_st + @stock2,
    '														 Stock3_st = Stock3_st + @stock3,
    '														 Stock4_st = Stock4_st + @stock4,
    '														 Stock5_st = Stock5_st + @stock5,
    '														 Stock6_st = Stock6_st + @stock6,
    '														 Stock7_st = Stock7_st + @stock7,
    '														 Stock8_st = Stock8_st + @stock8,
    '														 Stock9_st = Stock9_st + @stock9,
    '														 Stock10_st = Stock10_st + @stock10
    '														 WHERE CCia_st = @CCia_la AND Almac_st = @Almac_la AND Codigo_st = @Codigo_la


    '							UPDATE custom_vianny.DBO.Cag1700 SET Stock_17 = Stock_17 + @stock
    '												   WHERE CCia = @CCia_la AND Linea_17 = @Codigo_la
    '						END
    '					ELSE
    '						BEGIN
    '					      IF NOT EXISTS ( SELECT Codigo_st FROM custom_vianny.DBO.Cag1700_Stock WHERE CCia_st =@CCia_la AND Almac_st = @Almac_la AND Codigo_st = @Codigo_la)
    '								INSERT INTO custom_vianny.DBO.Cag1700_Stock (CCia_st,
    '														   Almac_st,
    '														   Codigo_st,
    '														   Stock_st,
    '														   Stock1_st, Stock2_st, Stock3_st, Stock4_st, Stock5_st, Stock6_st, Stock7_st, Stock8_st, Stock9_st, Stock10_st)
    '												  VALUES (@CCia_la,
    '														  @Almac_la,
    '														  @Codigo_la,
    '														  -1 * @stock,
    '														  -1 * @stock1, 
    '														  -1 * @stock2, 
    '														  -1 * @stock3, 
    '														  -1 * @stock4, 
    '														  -1 * @stock5, 
    '														  -1 * @stock6, 
    '														  -1 * @stock7, 
    '														  -1 * @stock8, 
    '														  -1 * @stock9, 
    '														  -1 * @stock10)
    '							ELSE
    '								UPDATE custom_vianny.DBO.Cag1700_Stock SET Stock_st = Stock_st - @stock,
    '														 Stock1_st = Stock1_st - @stock1,
    '														 Stock2_st = Stock2_st - @stock2,
    '														 Stock3_st = Stock3_st - @stock3,
    '														 Stock4_st = Stock4_st - @stock4,
    '														 Stock5_st = Stock5_st - @stock5,
    '														 Stock6_st = Stock6_st - @stock6,
    '														 Stock7_st = Stock7_st - @stock7,
    '														 Stock8_st = Stock8_st - @stock8,
    '														 Stock9_st = Stock9_st - @stock9,
    '														 Stock10_st = Stock10_st - @stock10
    '												     WHERE CCia_st = @CCia_la AND Almac_st = @Almac_la AND Codigo_st = @Codigo_la

    '							UPDATE custom_vianny.DBO.Cag1700 SET Stock_17 = Stock_17 - @stock
    '												   WHERE CCia = @CCia_la AND Linea_17 = @Codigo_la
    '						END
    '					END", conx)
    '            cmd19.Parameters.AddWithValue("@Codigo_la", Trim(DataGridView1.Rows(i).Cells(2).Value))
    '            cmd19.Parameters.AddWithValue("@ccom", "1")
    '            cmd19.Parameters.AddWithValue("@CCia_la", Trim(Label13.Text))
    '            cmd19.Parameters.AddWithValue("@Almac_la", Trim(TextBox22.Text))
    '            cmd19.Parameters.AddWithValue("@stock", DataGridView1.Rows(i).Cells(14).Value)
    '            cmd19.Parameters.AddWithValue("@stock1", DataGridView1.Rows(i).Cells(4).Value)
    '            cmd19.Parameters.AddWithValue("@stock2", DataGridView1.Rows(i).Cells(5).Value)
    '            cmd19.Parameters.AddWithValue("@stock3", DataGridView1.Rows(i).Cells(6).Value)
    '            cmd19.Parameters.AddWithValue("@stock4", DataGridView1.Rows(i).Cells(7).Value)
    '            cmd19.Parameters.AddWithValue("@stock5", DataGridView1.Rows(i).Cells(8).Value)
    '            cmd19.Parameters.AddWithValue("@stock6", DataGridView1.Rows(i).Cells(9).Value)
    '            cmd19.Parameters.AddWithValue("@stock7", DataGridView1.Rows(i).Cells(10).Value)
    '            cmd19.Parameters.AddWithValue("@stock8", DataGridView1.Rows(i).Cells(11).Value)
    '            cmd19.Parameters.AddWithValue("@stock9", DataGridView1.Rows(i).Cells(12).Value)
    '            cmd19.Parameters.AddWithValue("@stock10", DataGridView1.Rows(i).Cells(13).Value)
    '            cmd19.ExecuteNonQuery()


    '            For y = 4 To 13
    '                If DataGridView1.Rows(i).Cells(y).Value > 0 Then
    '                    Dim cmd16 As New SqlCommand("INSERT INTO custom_vianny.DBO.MAP030F (CCIA_3A , CPERIODO_3A, talm_3a, ccom_3a, ncom_3a, item_3a,linea_3a ,sinon_3a,linea2_3a, cant_3a, unid_3a, cantu_3a, obser1_3a, obser2_3a, obser3_3a, cantk_3a, cantke_3a, pedido_3a, flag_3a, peso_3a,tpeso_3a,unidk_3a , pieza_3a, color_3a, ncolor_3a, parti_3a, ocorte_3a, ordens_3a, nrollo_3a, kgrollo_3a, lote_3a, maquina_3a, galga_3a,sigla_3a, sitcal_3a, cantc_3a, ancho_3a, densid_3a, pedod_3a, ordtej_3a  , tipapt_3a , combo_3a,ccosto_3a,PreUn1_3a,Tot1_3a  )
    '                Values (@CCIA_3A,@CPERIODO_3A,@talm_3a,@ccom_3a,@ncom_3aa,@item_3a,@linea_3a, ''     ,'PT0101' ,@cantid ,@unid_3a,  @cantid,''        , ''       ,''        ,@cantid  , 0.00     ,@pedido_3a, 1      , 0.00000, 0.00   ,@unidk_3a,   ''    , ''      , ''       , ''      , ''       , ''       , ''       , 0.00      , ''     , ''        , 0.00    ,''      , 0        , 0.000   , 0.000   , 0.000    , ''      , ''         , ''        , ''      ,''       ,0.00000  , 0.00)", conx)
    '                    cmd16.Parameters.AddWithValue("@CCIA_3A", Trim(Label13.Text))
    '                    cmd16.Parameters.AddWithValue("@CPERIODO_3A", Trim(Label11.Text))
    '                    cmd16.Parameters.AddWithValue("@talm_3a", Trim(TextBox22.Text))
    '                    cmd16.Parameters.AddWithValue("@ccom_3a", "1")
    '                    cmd16.Parameters.AddWithValue("@ncom_3aa", Trim(TextBox23.Text) & Trim(TextBox24.Text))
    '                    cmd16.Parameters.AddWithValue("@item_3a", Trim(DataGridView1.Rows(i).Cells(0).Value))
    '                    cmd16.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
    '                    cmd16.Parameters.AddWithValue("@cantid", DataGridView1.Rows(i).Cells(y).Value)
    '                    cmd16.Parameters.AddWithValue("@unid_3a", "UND")
    '                    cmd16.Parameters.AddWithValue("@pedido_3a", DataGridView1.Rows(i).Cells(1).Value)
    '                    cmd16.Parameters.AddWithValue("@unidk_3a", "UND")
    '                    cmd16.ExecuteNonQuery()
    '                End If

    '            Next
    '        Next

    '        MsgBox("SE REGISTRO LA NIA CORRECTAMENTE")
    '    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = "4" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(4).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock1_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
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
            If TextBox27.Text = "2" Then

                stockFinal = (stock - jo) + valorAnterior1

            Else
                stockFinal = (stock - jo)
            End If
            'MsgBox(valorAnterior1.ToString)
            'MsgBox(Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value))
            'MsgBox(stockFinal)
            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(4).Value) > stockFinal Then

                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior1 + (Convert.ToInt32(TextBox11.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(4).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = valorAnterior1
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox11.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(4).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(4).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock1_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock1_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock1_la", stockFinal - valorAnterior1)
                    Else
                        cmd168.Parameters.AddWithValue("@stock1_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(4).Value))
                    End If
                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(4).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If


            End If
            Dim suma5 As Integer = 0
            Dim suma55 As Integer = 0
            For i1 = 4 To 13
                suma5 = suma5 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma5.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "5" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(5).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock, stockFinal As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock2_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox27.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior2

            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value) > (stock - jo) Then
                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior2 + (Convert.ToInt32(TextBox13.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(5).Value = valorAnterior2
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox13.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(5).Value = 0
                End If
            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock2_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock2_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock2_la", stockFinal - valorAnterior2)
                    Else
                        cmd168.Parameters.AddWithValue("@stock2_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(5).Value))
                    End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(5).HeaderText & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma6 As Integer = 0
            Dim suma66 As Integer = 0
            For i1 = 4 To 13
                suma6 = suma6 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma6.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "6" Then

            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(6).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock3_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox27.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior3
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value) > (stock - jo) Then
                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior3 + (Convert.ToInt32(TextBox14.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(6).Value = valorAnterior3
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox14.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(6).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock3_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock3_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock3_la", stockFinal - valorAnterior3)
                    Else
                        cmd168.Parameters.AddWithValue("@stock3_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(6).Value))
                    End If
                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(6).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma7 As Integer = 0
            Dim suma77 As Integer = 0
            For i1 = 4 To 13
                suma7 = suma7 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma7.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "7" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(7).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock4_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox27.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior4
            Else
                stockFinal = (stock - jo)
            End If
            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value) > (stock - jo) Then
                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior4 + (Convert.ToInt32(TextBox15.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(7).Value = valorAnterior4
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox15.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(7).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock4_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock4_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock4_la", stockFinal - valorAnterior4)
                    Else
                        cmd168.Parameters.AddWithValue("@stock4_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(7).Value))
                    End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(7).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma8 As Integer = 0
            Dim suma88 As Integer = 0
            For i1 = 4 To 13
                suma8 = suma8 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma8.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "8" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(8).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock5_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox27.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior5
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value) > (stock - jo) Then
                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior5 + (Convert.ToInt32(TextBox16.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(8).Value = valorAnterior5
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox16.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(8).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock5_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock5_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock5_la", stockFinal - valorAnterior5)
                    Else
                        cmd168.Parameters.AddWithValue("@stock5_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(8).Value))
                    End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(8).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma9 As Integer = 0
            Dim suma99 As Integer = 0
            For i1 = 4 To 13
                suma9 = suma9 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma9.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "9" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(9).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock6_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox27.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior6
            Else
                stockFinal = (stock - jo)
            End If
            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value) > (stock - jo) Then
                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior6 + (Convert.ToInt32(TextBox17.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(9).Value = valorAnterior6
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox17.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(9).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock6_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock6_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock6_la", stockFinal - valorAnterior6)
                    Else
                        cmd168.Parameters.AddWithValue("@stock6_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(9).Value))
                    End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(9).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma10 As Integer = 0
            Dim suma100 As Integer = 0
            For i1 = 4 To 13
                suma10 = suma10 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma10.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "10" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(10).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock7_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo, stockFinal As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox27.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior7
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value) > (stock - jo) Then
                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior7 + (Convert.ToInt32(TextBox18.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(10).Value = valorAnterior7
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox18.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(10).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock7_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock7_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock7_la", stockFinal - valorAnterior7)
                    Else
                        cmd168.Parameters.AddWithValue("@stock7_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(10).Value))
                    End If
                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(10).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma11 As Integer = 0
            Dim suma111 As Integer = 0
            For i1 = 4 To 13
                suma11 = suma11 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma11.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "11" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(11).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock, stockFinal As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock8_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox27.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior8
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value) > (stock - jo) Then
                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior8 + (Convert.ToInt32(TextBox19.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(11).Value = valorAnterior8
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox19.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(11).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock8_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock8_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock8_la", stockFinal - valorAnterior8)
                    Else
                        cmd168.Parameters.AddWithValue("@stock8_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(11).Value))
                    End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(11).HeaderText & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma12 As Integer = 0
            Dim suma122 As Integer = 0
            For i1 = 4 To 13
                suma12 = suma12 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma12.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "12" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(12).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock, stockFinal As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock8_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox27.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior9
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value) > (stock - jo) Then
                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior9 + (Convert.ToInt32(TextBox30.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(12).Value = valorAnterior9
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox20.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(12).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock8_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock8_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock8_la", stockFinal - valorAnterior9)
                    Else
                        cmd168.Parameters.AddWithValue("@stock8_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(12).Value))
                    End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(12).HeaderText & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma13 As Integer = 0
            Dim suma133 As Integer = 0
            For i1 = 4 To 13
                suma13 = suma13 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma13.ToString("N2")
            SumarCantidades()
        End If
        If e.ColumnIndex = "13" Then
            abrir()
            Dim agregar As String = "delete from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la ='" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(13).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim() + "'"
            ELIMINAR(agregar)
            'Calcual Stock
            Dim stock, stockFinal As Integer
            stock = 0
            Dim sql1021 As String = "select  cast( isnull(SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ),0) as int) AS Cant_1 from custom_vianny.dbo.mat030f where ccia_3b ='" + Label13.Text.ToString().Trim() + "'  and talm_3b ='" + ComboBox3.Text.ToString().Trim() + "' and pedido_3b ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and flag_3b ='1'  and cperiodo_3b ='" + Label11.Text.ToString().Trim() + "'"
            Dim cmd1021 As New SqlCommand(sql1021, conx)
            Rsr212 = cmd1021.ExecuteReader()
            If Rsr212.Read() Then
                stock = Convert.ToInt32(Rsr212(0))
            End If
            Rsr212.Close()

            'sumar del almacen ficticio
            Dim sql1022139 As String = "select cast( isnull(SUM( ISNULL(stock10_la,0)),0) AS INT) from custom_vianny.dbo.cag1700_almac_lotes where ccia_la ='" + Label13.Text.ToString().Trim() + "' and almac_la ='" + ComboBox3.Text.ToString().Trim() + "' and lote_la ='" + DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim() + "' and codigo_la like '" + TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() + "%'"
            Dim cmd1022139 As New SqlCommand(sql1022139, conx)
            Rsr3 = cmd1022139.ExecuteReader()
            Dim jo As Integer
            If Rsr3.Read() Then
                jo = Convert.ToInt32(Rsr3(0))
            Else
                jo = 0
            End If
            Rsr3.Close()
            If TextBox27.Text = "2" Then
                stockFinal = (stock - jo) + valorAnterior10
            Else
                stockFinal = (stock - jo)
            End If

            If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value) > (stock - jo) Then
                If TextBox27.Text = "2" Then
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (valorAnterior10 + (Convert.ToInt32(TextBox21.Text)) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(13).Value = valorAnterior10
                Else
                    MsgBox("La Cantidad Solicitada es Mayor al Stock, La Cantidad Ingresada Excede en: " + (Convert.ToInt32(TextBox21.Text) - Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value)).ToString())
                    DataGridView1.Rows(e.RowIndex).Cells(13).Value = 0
                End If

            Else
                If Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value) <> 0 Then
                    Dim cmd168 As New SqlCommand("insert into custom_vianny.dbo.cag1700_almac_lotes (ccia_la,almac_la,lote_la,stock10_la,codigo_la) values (@ccia_la,@almac_la,@lote_la,@stock10_la,@codigo_la)", conx)
                    cmd168.Parameters.AddWithValue("@ccia_la", Label13.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@almac_la", ComboBox3.Text.ToString().Trim())
                    cmd168.Parameters.AddWithValue("@lote_la", DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString().Trim())
                    If TextBox27.Text = "2" Then
                        cmd168.Parameters.AddWithValue("@stock10_la", stockFinal - valorAnterior10)
                    Else
                        cmd168.Parameters.AddWithValue("@stock10_la", Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(13).Value))
                    End If

                    cmd168.Parameters.AddWithValue("@codigo_la", TextBox4.Text.ToString().Trim() & TextBox5.Text.ToString().Trim() & DataGridView1.Columns(13).HeaderText.ToString().Trim() & DataGridView1.Rows(e.RowIndex).Cells(0).Value.ToString().Trim())
                    cmd168.ExecuteNonQuery()
                End If

            End If
            Dim suma14 As Integer = 0
            Dim suma144 As Integer = 0
            For i1 = 4 To 13
                suma14 = suma14 + Convert.ToInt32(DataGridView1.Rows(e.RowIndex).Cells(i1).Value)
            Next
            DataGridView1.Rows(e.RowIndex).Cells(14).Value = suma14.ToString("N2")
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
        TextBox38.Text = total.ToString("N2") ' o TextBoxTotal.Text = total.ToString("N2")
    End Sub
    Sub typeOnlynumbers(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        If Asc(e.KeyChar) >= 33 And Asc(e.KeyChar) <= 47 Or
            Asc(e.KeyChar) >= 58 Then
            e.Handled = True
        End If

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.F1 Then
            Clientes.TextBox3.Text = "35080"
            Clientes.Show()
        End If
    End Sub

    Private Sub TextBox22_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim sql1023 As String = " SELECT A.nomb_15m FROM custom_vianny.dbo.Mag1500 A  Where a.ccia ='" + Trim(Label13.Text) + "' and a.talm_15m ='" + Trim(TextBox22.Text) + "'"
            Dim cmd1023 As New SqlCommand(sql1023, conx)
            Rsr24 = cmd1023.ExecuteReader()
            If Rsr24.Read() = True Then
                TextBox25.Text = Rsr24(0)

            End If
            Rsr24.Close()

            TextBox23.Text = DateTimePicker1.Value.Month
            Select Case TextBox23.Text.Length
                Case "1" : TextBox23.Text = "0" & TextBox23.Text
                Case "2" : TextBox23.Text = TextBox23.Text

            End Select

            Dim JU As Integer

            Dim ano As String
            ano = Convert.ToString(Year(DateTimePicker1.Value))

            abrir()
            TextBox24.Enabled = True
            enunciado5 = New SqlCommand("select top 1 ncom_3, substring(ncom_3,3,7) as ncom from custom_vianny.dbo.mag030f  where ccia_3 =" + Label13.Text + " and cperiodo_3 = '" + Label11.Text + "' and talm_3 = '" + TextBox22.Text + "' " + "and ncom_3 like" + " '" + TextBox23.Text + "%' and ccom_3 = '1' order by ncom_3 desc", conx)
            respuesta5 = enunciado5.ExecuteReader
            While respuesta5.Read
                JU = respuesta5.Item("ncom")
            End While
            respuesta5.Close()
            TextBox24.Text = JU + 1
            Select Case TextBox24.Text.Length

                Case "1" : TextBox24.Text = "00000" & "" & TextBox24.Text
                Case "2" : TextBox24.Text = "0000" & "" & TextBox24.Text
                Case "3" : TextBox24.Text = "000" & "" & TextBox24.Text
                Case "4" : TextBox24.Text = "00" & "" & TextBox24.Text
                Case "5" : TextBox24.Text = "0" & "" & TextBox24.Text
                Case "6" : TextBox24.Text = TextBox24.Text
            End Select
            TextBox26.Select()

        End If
    End Sub

    Private Sub TextBox26_MouseDown(sender As Object, e As MouseEventArgs) Handles TextBox26.MouseDown

    End Sub



    'Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
    '    '    If e.KeyCode = Keys.Up OrElse e.KeyCode = Keys.Down Then
    '    '        MsgBox("1")

    '    '        If Trim(TextBox27.Text) = "2" Then

    '    '            Dim sql1 As String = "select  SUM(cant_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_Total,
    '    '  SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_1,
    '    '  SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_2,
    '    '  SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_3,
    '    '  SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_4,
    '    '  SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_5,
    '    '  SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_6,
    '    '  SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_7,
    '    '  SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_8,
    '    '  SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_9,
    '    '  SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_10
    '    'from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label13.Text) + "'  and talm_3b ='" + Trim(ComboBox3.Text) + "' and linea_3b ='" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(2).Value) + "' and flag_3b ='1' and cperiodo_3b ='" + Trim(Label11.Text) + "'"
    '    '            Dim cmd1 As New SqlCommand(sql1, conx)
    '    '            Rsr222 = cmd1.ExecuteReader()
    '    '            If Rsr222.Read() = True Then
    '    '                TextBox37.Text = Rsr222(1)
    '    '                TextBox36.Text = Rsr222(2)
    '    '                TextBox35.Text = Rsr222(3)
    '    '                TextBox34.Text = Rsr222(4)
    '    '                TextBox33.Text = Rsr222(5)
    '    '                TextBox32.Text = Rsr222(6)
    '    '                TextBox31.Text = Rsr222(7)
    '    '                TextBox30.Text = Rsr222(8)
    '    '                TextBox29.Text = Rsr222(9)
    '    '                TextBox28.Text = Rsr222(10)

    '    '            Else
    '    '                TextBox37.Text = 0
    '    '                TextBox36.Text = 0
    '    '                TextBox35.Text = 0
    '    '                TextBox34.Text = 0
    '    '                TextBox33.Text = 0
    '    '                TextBox32.Text = 0
    '    '                TextBox31.Text = 0
    '    '                TextBox30.Text = 0
    '    '                TextBox29.Text = 0
    '    '                TextBox28.Text = 0

    '    '            End If
    '    '            Rsr222.Close()
    '    '            Dim sql1023 As String = " select cant1_3b,cant2_3b,cant3_3b,cant4_3b,cant5_3b,cant6_3b,cant7_3b,cant8_3b,cant9_3b,cant10_3b,cant_3b from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label13.Text) + "' and cperiodo_3b ='" + Trim(Label11.Text) + "' and talm_3b ='" + Trim(ComboBox3.Text) + "' and flag_3b ='1' and linea_3b ='" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(2).Value) + "' and ccom_3b ='2'"
    '    '            Dim cmd1023 As New SqlCommand(sql1023, conx)
    '    '            Rsr22 = cmd1023.ExecuteReader()
    '    '            If Rsr22.Read() = True Then
    '    '                TextBox11.Text = Rsr22(0) + TextBox37.Text
    '    '                TextBox13.Text = Rsr22(1) + TextBox36.Text
    '    '                TextBox14.Text = Rsr22(2) + TextBox35.Text
    '    '                TextBox15.Text = Rsr22(3) + TextBox34.Text
    '    '                TextBox16.Text = Rsr22(4) + TextBox33.Text
    '    '                TextBox17.Text = Rsr22(5) + TextBox32.Text
    '    '                TextBox18.Text = Rsr22(6) + TextBox31.Text
    '    '                TextBox19.Text = Rsr22(7) + TextBox30.Text
    '    '                TextBox20.Text = Rsr22(8) + TextBox29.Text
    '    '                TextBox21.Text = Rsr22(9) + TextBox28.Text

    '    '            Else
    '    '                TextBox11.Text = 0
    '    '                TextBox13.Text = 0
    '    '                TextBox14.Text = 0
    '    '                TextBox15.Text = 0
    '    '                TextBox16.Text = 0
    '    '                TextBox17.Text = 0
    '    '                TextBox18.Text = 0
    '    '                TextBox19.Text = 0
    '    '                TextBox20.Text = 0
    '    '                TextBox21.Text = 0

    '    '            End If
    '    '            Rsr22.Close()

    '    '            Dim sql10235 As String = "SELECT A.*,
    '    '                		  ISNULL(B.Dele,'') AS NTalla1,
    '    '                		  ISNULL(C.Dele,'') AS NTalla2,
    '    '                		  ISNULL(D.Dele,'') AS NTalla3,
    '    '                		  ISNULL(E.Dele,'') AS NTalla4,
    '    '                		  ISNULL(F.Dele,'') AS NTalla5,
    '    '                		  ISNULL(G.Dele,'') AS NTalla6,
    '    '                		  ISNULL(H.Dele,'') AS NTalla7,
    '    '                		  ISNULL(I.Dele,'') AS NTalla8,
    '    '                		  ISNULL(J.Dele,'') AS NTalla9,
    '    '                		  ISNULL(K.Dele,'') AS NTalla10
    '    '                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
    '    '                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
    '    '                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
    '    '                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
    '    '                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
    '    '                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
    '    '                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
    '    '                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
    '    '                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
    '    '                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
    '    '                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
    '    '                 Where A.CCia_tl = '" + Trim(Label13.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(15).Value) + "'
    '    '                ORDER BY  A.CCia_tl, A.Codigo_tl"
    '    '            Dim cmd10235 As New SqlCommand(sql10235, conx)
    '    '            Rsr23 = cmd10235.ExecuteReader()
    '    '            If Rsr23.Read() = True Then
    '    '                DataGridView1.Columns(4).HeaderText = Trim(Rsr23(13))
    '    '                DataGridView1.Columns(5).HeaderText = Trim(Rsr23(14))
    '    '                DataGridView1.Columns(6).HeaderText = Trim(Rsr23(15))
    '    '                DataGridView1.Columns(7).HeaderText = Trim(Rsr23(16))
    '    '                DataGridView1.Columns(8).HeaderText = Trim(Rsr23(17))
    '    '                DataGridView1.Columns(9).HeaderText = Trim(Rsr23(18))
    '    '                DataGridView1.Columns(10).HeaderText = Trim(Rsr23(19))
    '    '                DataGridView1.Columns(11).HeaderText = Trim(Rsr23(20))
    '    '                DataGridView1.Columns(12).HeaderText = Trim(Rsr23(21))
    '    '                DataGridView1.Columns(13).HeaderText = Trim(Rsr23(22))



    '    '            End If
    '    '            Rsr23.Close()
    '    '        Else
    '    '            If (DataGridView1.CurrentCell.RowIndex - 1) >= 0 Then



    '    '                Dim sql1 As String = "select  SUM(cant_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_Total,
    '    '  SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_1,
    '    '  SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_2,
    '    '  SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_3,
    '    '  SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_4,
    '    '  SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_5,
    '    '  SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_6,
    '    '  SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_7,
    '    '  SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_8,
    '    '  SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_9,
    '    '  SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_10
    '    'from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label13.Text) + "'  and talm_3b ='" + Trim(ComboBox3.Text) + "' and linea_3b ='" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(2).Value) + "' and flag_3b ='1'  and cperiodo_3b ='" + Trim(Label11.Text) + "' "
    '    '                Dim cmd1 As New SqlCommand(sql1, conx)
    '    '                Rsr222 = cmd1.ExecuteReader()
    '    '                If Rsr222.Read() = True Then
    '    '                    TextBox11.Text = Rsr222(1)
    '    '                    TextBox13.Text = Rsr222(2)
    '    '                    TextBox14.Text = Rsr222(3)
    '    '                    TextBox15.Text = Rsr222(4)
    '    '                    TextBox16.Text = Rsr222(5)
    '    '                    TextBox17.Text = Rsr222(6)
    '    '                    TextBox18.Text = Rsr222(7)
    '    '                    TextBox19.Text = Rsr222(8)
    '    '                    TextBox20.Text = Rsr222(9)
    '    '                    TextBox21.Text = Rsr222(10)

    '    '                Else
    '    '                    TextBox11.Text = 0
    '    '                    TextBox13.Text = 0
    '    '                    TextBox14.Text = 0
    '    '                    TextBox15.Text = 0
    '    '                    TextBox16.Text = 0
    '    '                    TextBox17.Text = 0
    '    '                    TextBox18.Text = 0
    '    '                    TextBox19.Text = 0
    '    '                    TextBox20.Text = 0
    '    '                    TextBox21.Text = 0

    '    '                End If
    '    '                Rsr222.Close()
    '    '                Dim sql10235 As String = "SELECT A.*,
    '    '                		  ISNULL(B.Dele,'') AS NTalla1,
    '    '                		  ISNULL(C.Dele,'') AS NTalla2,
    '    '                		  ISNULL(D.Dele,'') AS NTalla3,
    '    '                		  ISNULL(E.Dele,'') AS NTalla4,
    '    '                		  ISNULL(F.Dele,'') AS NTalla5,
    '    '                		  ISNULL(G.Dele,'') AS NTalla6,
    '    '                		  ISNULL(H.Dele,'') AS NTalla7,
    '    '                		  ISNULL(I.Dele,'') AS NTalla8,
    '    '                		  ISNULL(J.Dele,'') AS NTalla9,
    '    '                		  ISNULL(K.Dele,'') AS NTalla10
    '    '                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
    '    '                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
    '    '                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
    '    '                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
    '    '                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
    '    '                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
    '    '                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
    '    '                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
    '    '                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
    '    '                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
    '    '                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
    '    '                 Where A.CCia_tl = '" + Trim(Label13.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(15).Value) + "'
    '    '                ORDER BY  A.CCia_tl, A.Codigo_tl"
    '    '                Dim cmd10235 As New SqlCommand(sql10235, conx)
    '    '                Rsr23 = cmd10235.ExecuteReader()
    '    '                If Rsr23.Read() = True Then
    '    '                    DataGridView1.Columns(4).HeaderText = Trim(Rsr23(13))
    '    '                    DataGridView1.Columns(5).HeaderText = Trim(Rsr23(14))
    '    '                    DataGridView1.Columns(6).HeaderText = Trim(Rsr23(15))
    '    '                    DataGridView1.Columns(7).HeaderText = Trim(Rsr23(16))
    '    '                    DataGridView1.Columns(8).HeaderText = Trim(Rsr23(17))
    '    '                    DataGridView1.Columns(9).HeaderText = Trim(Rsr23(18))
    '    '                    DataGridView1.Columns(10).HeaderText = Trim(Rsr23(19))
    '    '                    DataGridView1.Columns(11).HeaderText = Trim(Rsr23(20))
    '    '                    DataGridView1.Columns(12).HeaderText = Trim(Rsr23(21))
    '    '                    DataGridView1.Columns(13).HeaderText = Trim(Rsr23(22))



    '    '                End If
    '    '                Rsr23.Close()
    '    '            End If

    '    '        End If
    '    '    End If

    '    '    If e.KeyCode = Keys.Down OrElse e.KeyCode = Keys.Down Then
    '    '        If Trim(TextBox27.Text) = "2" Then
    '    '            MsgBox("2")

    '    '            Dim sql1 As String = "select  SUM(cant_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_Total,
    '    '  SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_1,
    '    '  SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_2,
    '    '  SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_3,
    '    '  SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_4,
    '    '  SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_5,
    '    '  SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_6,
    '    '  SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_7,
    '    '  SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_8,
    '    '  SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_9,
    '    '  SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_10
    '    'from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label13.Text) + "'  and talm_3b ='" + Trim(ComboBox3.Text) + "' and linea_3b ='" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(2).Value) + "' and flag_3b ='1'  and cperiodo_3b ='" + Trim(Label11.Text) + "'"
    '    '            Dim cmd1 As New SqlCommand(sql1, conx)
    '    '            Rsr222 = cmd1.ExecuteReader()
    '    '            If Rsr222.Read() = True Then
    '    '                TextBox37.Text = Rsr222(1)
    '    '                TextBox36.Text = Rsr222(2)
    '    '                TextBox35.Text = Rsr222(3)
    '    '                TextBox34.Text = Rsr222(4)
    '    '                TextBox33.Text = Rsr222(5)
    '    '                TextBox32.Text = Rsr222(6)
    '    '                TextBox31.Text = Rsr222(7)
    '    '                TextBox30.Text = Rsr222(8)
    '    '                TextBox29.Text = Rsr222(9)
    '    '                TextBox28.Text = Rsr222(10)

    '    '            Else
    '    '                TextBox37.Text = 0
    '    '                TextBox36.Text = 0
    '    '                TextBox35.Text = 0
    '    '                TextBox34.Text = 0
    '    '                TextBox33.Text = 0
    '    '                TextBox32.Text = 0
    '    '                TextBox31.Text = 0
    '    '                TextBox30.Text = 0
    '    '                TextBox29.Text = 0
    '    '                TextBox28.Text = 0

    '    '            End If
    '    '            Rsr222.Close()
    '    '            Dim sql1023 As String = " select cant1_3b,cant2_3b,cant3_3b,cant4_3b,cant5_3b,cant6_3b,cant7_3b,cant8_3b,cant9_3b,cant10_3b,cant_3b from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label13.Text) + "' and cperiodo_3b ='" + Trim(Label11.Text) + "' and talm_3b ='" + Trim(ComboBox3.Text) + "' and flag_3b ='1' and linea_3b ='" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(2).Value) + "' and ccom_3b ='2'"
    '    '            Dim cmd1023 As New SqlCommand(sql1023, conx)
    '    '            Rsr22 = cmd1023.ExecuteReader()
    '    '            If Rsr22.Read() = True Then
    '    '                TextBox11.Text = Rsr22(0) + TextBox37.Text
    '    '                TextBox13.Text = Rsr22(1) + TextBox36.Text
    '    '                TextBox14.Text = Rsr22(2) + TextBox35.Text
    '    '                TextBox15.Text = Rsr22(3) + TextBox34.Text
    '    '                TextBox16.Text = Rsr22(4) + TextBox33.Text
    '    '                TextBox17.Text = Rsr22(5) + TextBox32.Text
    '    '                TextBox18.Text = Rsr22(6) + TextBox31.Text
    '    '                TextBox19.Text = Rsr22(7) + TextBox30.Text
    '    '                TextBox20.Text = Rsr22(8) + TextBox29.Text
    '    '                TextBox21.Text = Rsr22(9) + TextBox28.Text

    '    '            Else
    '    '                TextBox11.Text = 0
    '    '                TextBox13.Text = 0
    '    '                TextBox14.Text = 0
    '    '                TextBox15.Text = 0
    '    '                TextBox16.Text = 0
    '    '                TextBox17.Text = 0
    '    '                TextBox18.Text = 0
    '    '                TextBox19.Text = 0
    '    '                TextBox20.Text = 0
    '    '                TextBox21.Text = 0

    '    '            End If
    '    '            Rsr22.Close()

    '    '            Dim sql10235 As String = "SELECT A.*,
    '    '                		  ISNULL(B.Dele,'') AS NTalla1,
    '    '                		  ISNULL(C.Dele,'') AS NTalla2,
    '    '                		  ISNULL(D.Dele,'') AS NTalla3,
    '    '                		  ISNULL(E.Dele,'') AS NTalla4,
    '    '                		  ISNULL(F.Dele,'') AS NTalla5,
    '    '                		  ISNULL(G.Dele,'') AS NTalla6,
    '    '                		  ISNULL(H.Dele,'') AS NTalla7,
    '    '                		  ISNULL(I.Dele,'') AS NTalla8,
    '    '                		  ISNULL(J.Dele,'') AS NTalla9,
    '    '                		  ISNULL(K.Dele,'') AS NTalla10
    '    '                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
    '    '                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
    '    '                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
    '    '                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
    '    '                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
    '    '                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
    '    '                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
    '    '                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
    '    '                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
    '    '                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
    '    '                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
    '    '                 Where A.CCia_tl = '" + Trim(Label13.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(15).Value) + "'
    '    '                ORDER BY  A.CCia_tl, A.Codigo_tl"
    '    '            Dim cmd10235 As New SqlCommand(sql10235, conx)
    '    '            Rsr23 = cmd10235.ExecuteReader()
    '    '            If Rsr23.Read() = True Then
    '    '                DataGridView1.Columns(4).HeaderText = Trim(Rsr23(13))
    '    '                DataGridView1.Columns(5).HeaderText = Trim(Rsr23(14))
    '    '                DataGridView1.Columns(6).HeaderText = Trim(Rsr23(15))
    '    '                DataGridView1.Columns(7).HeaderText = Trim(Rsr23(16))
    '    '                DataGridView1.Columns(8).HeaderText = Trim(Rsr23(17))
    '    '                DataGridView1.Columns(9).HeaderText = Trim(Rsr23(18))
    '    '                DataGridView1.Columns(10).HeaderText = Trim(Rsr23(19))
    '    '                DataGridView1.Columns(11).HeaderText = Trim(Rsr23(20))
    '    '                DataGridView1.Columns(12).HeaderText = Trim(Rsr23(21))
    '    '                DataGridView1.Columns(13).HeaderText = Trim(Rsr23(22))



    '    '            End If
    '    '            Rsr23.Close()
    '    '        Else
    '    '            Dim sql1 As String = "select  SUM(cant_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_Total,
    '    '  SUM(cant1_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_1,
    '    '  SUM(cant2_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_2,
    '    '  SUM(cant3_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_3,
    '    '  SUM(cant4_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_4,
    '    '  SUM(cant5_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_5,
    '    '  SUM(cant6_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_6,
    '    '  SUM(cant7_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_7,
    '    '  SUM(cant8_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_8,
    '    '  SUM(cant9_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_9,
    '    '  SUM(cant10_3b * CASE WHEN ccom_3b IN ('2','4') THEN -1 ELSE 1 END ) AS Cant_10
    '    'from custom_vianny.dbo.mat030f where ccia_3b ='" + Trim(Label13.Text) + "'  and talm_3b ='" + Trim(ComboBox3.Text) + "' and linea_3b ='" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(2).Value) + "' and flag_3b ='1'  and cperiodo_3b ='" + Trim(Label11.Text) + "' "
    '    '            Dim cmd1 As New SqlCommand(sql1, conx)
    '    '            Rsr222 = cmd1.ExecuteReader()
    '    '            If Rsr222.Read() = True Then
    '    '                TextBox11.Text = Rsr222(1)
    '    '                TextBox13.Text = Rsr222(2)
    '    '                TextBox14.Text = Rsr222(3)
    '    '                TextBox15.Text = Rsr222(4)
    '    '                TextBox16.Text = Rsr222(5)
    '    '                TextBox17.Text = Rsr222(6)
    '    '                TextBox18.Text = Rsr222(7)
    '    '                TextBox19.Text = Rsr222(8)
    '    '                TextBox20.Text = Rsr222(9)
    '    '                TextBox21.Text = Rsr222(10)

    '    '            Else
    '    '                TextBox11.Text = 0
    '    '                TextBox13.Text = 0
    '    '                TextBox14.Text = 0
    '    '                TextBox15.Text = 0
    '    '                TextBox16.Text = 0
    '    '                TextBox17.Text = 0
    '    '                TextBox18.Text = 0
    '    '                TextBox19.Text = 0
    '    '                TextBox20.Text = 0
    '    '                TextBox21.Text = 0

    '    '            End If
    '    '            Rsr222.Close()
    '    '            Dim sql10235 As String = "SELECT A.*,
    '    '                		  ISNULL(B.Dele,'') AS NTalla1,
    '    '                		  ISNULL(C.Dele,'') AS NTalla2,
    '    '                		  ISNULL(D.Dele,'') AS NTalla3,
    '    '                		  ISNULL(E.Dele,'') AS NTalla4,
    '    '                		  ISNULL(F.Dele,'') AS NTalla5,
    '    '                		  ISNULL(G.Dele,'') AS NTalla6,
    '    '                		  ISNULL(H.Dele,'') AS NTalla7,
    '    '                		  ISNULL(I.Dele,'') AS NTalla8,
    '    '                		  ISNULL(J.Dele,'') AS NTalla9,
    '    '                		  ISNULL(K.Dele,'') AS NTalla10
    '    '                 FROM custom_vianny.dbo.Tallado A LEFT JOIN custom_vianny.dbo.Tab0100 B
    '    '                 				 ON A.CCia_tl = B.CCia AND 'TALLAS' = B.CTab AND A.Talla1_tl = LEFT(B.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 C
    '    '                 				 ON A.CCia_tl = C.CCia AND 'TALLAS' = C.CTab AND A.Talla2_tl = LEFT(C.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 D
    '    '                 				 ON A.CCia_tl = D.CCia AND 'TALLAS' = D.CTab AND A.Talla3_tl = LEFT(D.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 E
    '    '                 				 ON A.CCia_tl = E.CCia AND 'TALLAS' = E.CTab AND A.Talla4_tl = LEFT(E.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 F
    '    '                 				 ON A.CCia_tl = F.CCia AND 'TALLAS' = F.CTab AND A.Talla5_tl = LEFT(F.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 G
    '    '                 				 ON A.CCia_tl = G.CCia AND 'TALLAS' = G.CTab AND A.Talla6_tl = LEFT(G.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 H
    '    '                 				 ON A.CCia_tl = H.CCia AND 'TALLAS' = H.CTab AND A.Talla7_tl = LEFT(H.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 I
    '    '                 				 ON A.CCia_tl = I.CCia AND 'TALLAS' = I.CTab AND A.Talla8_tl = LEFT(I.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 J
    '    '                 				 ON A.CCia_tl = J.CCia AND 'TALLAS' = J.CTab AND A.Talla9_tl = LEFT(J.Cele,4)
    '    '                 				 LEFT JOIN custom_vianny.dbo.Tab0100 K
    '    '                 				 ON A.CCia_tl = K.CCia AND 'TALLAS' = K.CTab AND A.Talla10_tl = LEFT(K.Cele,4)
    '    '                 Where A.CCia_tl = '" + Trim(Label13.Text) + "' AND A.Codigo_tl = '" + Trim(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(15).Value) + "'
    '    '                ORDER BY  A.CCia_tl, A.Codigo_tl"
    '    '            Dim cmd10235 As New SqlCommand(sql10235, conx)
    '    '            Rsr23 = cmd10235.ExecuteReader()
    '    '            If Rsr23.Read() = True Then
    '    '                DataGridView1.Columns(4).HeaderText = Trim(Rsr23(13))
    '    '                DataGridView1.Columns(5).HeaderText = Trim(Rsr23(14))
    '    '                DataGridView1.Columns(6).HeaderText = Trim(Rsr23(15))
    '    '                DataGridView1.Columns(7).HeaderText = Trim(Rsr23(16))
    '    '                DataGridView1.Columns(8).HeaderText = Trim(Rsr23(17))
    '    '                DataGridView1.Columns(9).HeaderText = Trim(Rsr23(18))
    '    '                DataGridView1.Columns(10).HeaderText = Trim(Rsr23(19))
    '    '                DataGridView1.Columns(11).HeaderText = Trim(Rsr23(20))
    '    '                DataGridView1.Columns(12).HeaderText = Trim(Rsr23(21))
    '    '                DataGridView1.Columns(13).HeaderText = Trim(Rsr23(22))



    '    '            End If
    '    '            Rsr23.Close()
    '    '        End If
    '    '    End If

    'End Sub





End Class