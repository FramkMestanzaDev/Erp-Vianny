Imports System.Data.SqlClient
Public Class Explosion_avios_graus
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Public Sub NumConFrac(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False

        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        NumConFrac(TextBox7, e)
    End Sub
    Dim Rs1, Rs2, Rs3, Rs24 As SqlDataReader
    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()
            abrir()
            Dim u As String
            Dim sql1 As String = "SELECT descri_3,fich_3,program_3,cantp_3,b.color_16b,tallador_3,SUM(t.cantidad)  FROM custom_vianny.dbo.qag0300 Q 
inner join custom_vianny.dbo.qag160b b on q.ccia = b.ccia  and q.ncom_3 = b.ncom_16b
inner join detalle_TENDIDO_CORTE t on t.op = q.ncom_3
WHERE ncom_3 ='" + TextBox7.Text + "' AND q.ccia='" + Label8.Text + "' GROUP BY descri_3,fich_3,program_3,cantp_3,b.color_16b,tallador_3"
            Dim cmd1 As New SqlCommand(sql1, conx)
            Rs1 = cmd1.ExecuteReader()
            If Rs1.Read() = True Then
                TextBox1.Text = Rs1(0)
                TextBox2.Text = Rs1(1)
                TextBox3.Text = Rs1(3)

                TextBox5.Text = Rs1(2)
                TextBox6.Text = Rs1(4)
                TextBox9.Text = Rs1(6)
                u = Rs1(5)
                Rs1.Close()

                Dim sql22 As String = "SELECT 
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
		        Where A.CCia_tl = '" + Trim(Label8.Text) + "' AND A.Codigo_tl = '" + u + "'
		       ORDER BY  A.CCia_tl, A.Codigo_tl"
                Dim cmd22 As New SqlCommand(sql22, conx)
                Rs3 = cmd22.ExecuteReader()
                Rs3.Read()
                TextBox4.Text = Trim(Rs3(0)) + "/" + Trim(Rs3(1)) + "/" + Trim(Rs3(2)) + "/" + Trim(Rs3(3)) + "/" + Trim(Rs3(4)) + "/" + Trim(Rs3(5)) + "/" + Trim(Rs3(6)) + "/" + Trim(Rs3(7)) + "/" + Trim(Rs3(8)) + "/" + Trim(Rs3(9))







                Rs3.Close()
                abrir()
                Dim sql2 As String = "select COUNT(ncom_8) from custom_vianny.dbo.qamc800 where ncom_8 ='" + Trim(TextBox7.Text) + "' AND ccia_8 ='" + Trim(Label8.Text) + "'"
                Dim cmd2 As New SqlCommand(sql2, conx)
                Rs2 = cmd2.ExecuteReader()
                Rs2.Read()

                If Rs2(0) > 0 Then
                    Rs2.Close()
                    abrir()

                    Dim cmd As New SqlDataAdapter("SELECT item_8,t.cele,gene_8,A.linea_17,a.nomb_17,factor_8,udm_8,total,ISNULL(talla,'0')   FROM custom_vianny.dbo.qag0300 q 
                    inner join custom_vianny.dbo.qamc800 c on q.ccia = c.ccia_8 and q.ncom_3 =c.ncom_8  
                    left join custom_vianny.dbo.cag1700 a on c.linea_8 = a.linea_17 and c.ccia_8 =a.ccia 
                    left join custom_vianny.dbo.TAB0100 t on c.area_8 = t.cele
                    left join custom_vianny.dbo.cag1000 g on q.fich_3 = g.ruc_10  and g.ccia ='" + Trim(Label8.Text) + "'
                    where ncom_3='" + TextBox7.Text + "'and q.ccia ='" + Trim(Label8.Text) + "' and  t.ccia='" + Trim(Label8.Text) + "' AND t.CTAB='AREAMC' order by cast(item_8 as int)", conx)
                    Dim da As New DataTable
                    cmd.Fill(da)
                    DataGridView2.DataSource = da

                    DataGridView1.Rows.Add(DataGridView2.Rows.Count)

                    For i = 0 To DataGridView2.Rows.Count - 1
                        DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(0).Value
                        ''DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(1).Value
                        'Dim lsQuery As String = "SELECT cele,dele FROM custom_vianny.dbo.TAB0100 A  Where A.ccia='01' AND A.CTAB='AREAMC'  and dele ='" + Trim(DataGridView2.Rows(i).Cells(1).Value) + "' "
                        'Dim loDataTable As New DataTable
                        'Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
                        'loDataAdapter.Fill(loDataTable)
                        'ET.DataSource = loDataTable

                        'ET.DisplayMember = "dele"
                        'ET.ValueMember = "cele"
                        'ET.DataPropertyName = "dele"
                        'ET.DropDownWidth = 200
                        'Dim cmb As New DataGridViewComboBoxColumn()
                        'cmb.DataSource = loDataTable
                        llenar_combo_box()
                        llenar_combo_box1()
                        llenar_combo_box3()
                        DataGridView1.Rows(i).Cells(1).Value = (Trim(DataGridView2.Rows(i).Cells(1).Value)).ToString
                        DataGridView1.Rows(i).Cells(2).Value = (Trim(DataGridView2.Rows(i).Cells(2).Value)).ToString
                        DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(3).Value
                        DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(4).Value
                        DataGridView1.Rows(i).Cells(7).Value = DataGridView2.Rows(i).Cells(5).Value
                        DataGridView1.Rows(i).Cells(8).Value = DataGridView2.Rows(i).Cells(6).Value
                        DataGridView1.Rows(i).Cells(9).Value = DataGridView2.Rows(i).Cells(7).Value
                        Dim J As String
                        J = "GENERAL"
                        If Trim(DataGridView2.Rows(i).Cells(8).Value) = "0" Then
                            DataGridView1.Rows(i).Cells(6).Value = (J).ToString
                        Else
                            DataGridView1.Rows(i).Cells(6).Value = (Trim(DataGridView2.Rows(i).Cells(8).Value)).ToString
                        End If

                    Next

                End If
                Rs2.Close()
                Button1.Enabled = True
                Button3.Enabled = True
                TextBox8.Enabled = True
            Else
                MsgBox("LA OP INGRESADA NO EXISTE")
                TextBox7.Text = ""
                Rs1.Close()
            End If


        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox3.Text = "" Or TextBox3.Text = "0.00" Then
            MsgBox("NO SE PUEDE AGREGAR INFORMACION CUANDO NO HAY CANTIDAD PROGRAMADA")
        Else
            DataGridView1.Rows.Add(1)
            Dim k As Integer
            k = DataGridView1.Rows.Count

            If k = 1 Then
                DataGridView1.Rows(0).Cells(0).Value = 1
            Else
                For i = 1 To k - 1
                    DataGridView1.Rows(i).Cells(0).Value = DataGridView1.Rows(i - 1).Cells(0).Value + 1
                Next
            End If
            abrir()
            llenar_combo_box()
            llenar_combo_box1()
            llenar_combo_box3()
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        If e.ColumnIndex = 1 Or e.ColumnIndex = 2 Then  ' <<<< Aquí pones el nº de columna del Combo
            Me.DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
            If Not IsNothing(DataGridView1.EditingControl) Then

                Dim cmb As ComboBox = DataGridView1.EditingControl
                cmb.DroppedDown = True

            End If
        Else
            Dim i, fila As Integer

            i = DataGridView1.Rows.Count
            fila = e.RowIndex
            If e.ColumnIndex = 7 Then
                'DataGridView1.Rows(fila).Cells(9).Value = DataGridView1.Rows(fila).Cells(7).Value * TextBox3.Text

            End If
        End If

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 3 Then
            Articulos.Label2.Text = e.RowIndex
            Articulos.Label3.Text = Label10.Text

            Articulos.Show()
        End If


    End Sub





    Private Sub Explosion_avios_graus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox7.Select()


    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila As Integer

        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CONSUMO" Then
            If (Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)).ToString = "" Then
                MsgBox("FALTA INGRESAL TALLA")
            Else
                If (Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)).ToString = "GENERAL" Then
                    DataGridView1.Rows(fila).Cells(9).Value = DataGridView1.Rows(fila).Cells(7).Value * TextBox9.Text
                Else
                    Dim sql1 As String = "select isnull(SUM(cantidad),0) from detalle_TENDIDO_CORTE where op ='" + Trim(TextBox7.Text) + "' and talla ='" + (Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)).ToString + "' group by talla"
                    Dim cmd1 As New SqlCommand(sql1, conx)
                    Rs1 = cmd1.ExecuteReader()
                    If Rs1.Read() Then
                        DataGridView1.Rows(fila).Cells(9).Value = Rs1(0) * DataGridView1.Rows(fila).Cells(7).Value
                    End If

                    Rs1.Close()
                End If

            End If
            'SELECT  isnull(cant_16d,0) FROM custom_vianny.dbo.qag16dv d left join custom_vianny.dbo.qag160c c on d.ccia_16d=c.ccia and d.ncom_16d=c.ncom_16c and c.correl_16c=d.correl_16d
            ' Where  d.ccia_16d = '" + Trim(Label8.Text) + "' And ncom_16d = '" + Trim(TextBox7.Text) + "' And color_16d = '01'  and talla_16c ='" + (Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value)).ToString + "'

        End If
    End Sub
    Dim ty, TY2 As DataTable
    Sub llenar_combo_box()
        Try
            Dim lsQuery As String = "SELECT cele,dele FROM custom_vianny.dbo.TAB0100 A  Where A.ccia='" + Label8.Text + "' AND A.CTAB='AREAMC'  and dele <> 'TODOS'"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            ET.DataSource = loDataTable

            ET.DisplayMember = "dele"
            ET.ValueMember = "cele"
            ET.DropDownWidth = 200
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box3()
        Try

            Dim lsQuery As String = "select 'GENERAL' AS talla_16c
union all
SELECT talla_16c FROM custom_vianny.dbo.qag160d d left join custom_vianny.dbo.qag160c c on d.ccia=c.ccia and d.ncom_16d=c.ncom_16c and c.correl_16c=d.correl_16d
            Where  d.ccia = '" + Trim(Label8.Text) + "' And ncom_16d = '" + TextBox7.Text + "' And color_16d = '01'"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            TALLA.DataSource = loDataTable

            TALLA.DisplayMember = "talla_16c"
            TALLA.ValueMember = "talla_16c"
            TALLA.DropDownWidth = 200
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub llenar_combo_box1()
        Try

            Dim lsQuery As String = "SELECT cele,dele FROM custom_vianny.dbo.TAB0100 A  Where A.ccia='01' AND A.CTAB='ARTGEN'"
            Dim loDataTable As New DataTable
            Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
            loDataAdapter.Fill(loDataTable)
            GEN.DataSource = loDataTable

            GEN.DisplayMember = "dele"
            GEN.ValueMember = "cele"
            GEN.DropDownWidth = 200
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim respuesta As DialogResult


        respuesta = MessageBox.Show("QUIERES IMPRIMIR LA MATRIZ CONSUMO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Rpt_Explocion.TextBox1.Text = TextBox7.Text
            Rpt_Explocion.TextBox2.Text = Label8.Text
            Rpt_Explocion.Show()

        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim U As Integer
        U = DataGridView1.Rows.Count
        If U = 0 Then
            MsgBox("NO HAY DATOS PARA ELIMINAR")
        Else
            Dim respuesta As DialogResult
            Dim I18, VAL As Integer
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
                I18 = DataGridView1.Rows.Count
                For i1 = 0 To I18 - 1
                    VAL = i1 + 1
                    DataGridView1.Rows(i1).Cells(0).Value = i1 + 1
                Next
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()

        Dim agregar As String = "DELETE FROM custom_vianny.dbo.qamc800  WHERE ncom_8 ='" + TextBox7.Text + "' AND ccia_8 ='" + Trim(Label8.Text) + "'"
        ELIMINAR(agregar)
        Dim o As Integer
        o = DataGridView1.Rows.Count
        For i = 0 To o - 1
            abrir()
            Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.dbo.qamc800 ( ccia_8 , item_8 , ncom_8 , linea_8 , factor_8 , udm_8 ,gene_8 , nomb_8 , obser_8 , talcol_8 , area_8 , cierre_8 , cliente_8,total,talla) 
        VALUES 	 ( @ccia , @items , @op , @codigo , @consumo , @unidad ,@gene , @descripgene , '' , 0 , @area , 0 , '',@total,@talla )", conx)
            cmd15.Parameters.AddWithValue("@ccia", Trim(Label8.Text))
            cmd15.Parameters.AddWithValue("@items", DataGridView1.Rows(i).Cells(0).Value)
            cmd15.Parameters.AddWithValue("@op", TextBox7.Text)
            cmd15.Parameters.AddWithValue("@codigo", DataGridView1.Rows(i).Cells(4).Value)
            cmd15.Parameters.AddWithValue("@consumo", DataGridView1.Rows(i).Cells(7).Value)
            cmd15.Parameters.AddWithValue("@unidad", DataGridView1.Rows(i).Cells(8).Value)
            cmd15.Parameters.AddWithValue("@gene", DataGridView1.Rows(i).Cells(2).Value)
            If Trim(DataGridView1.Rows(i).Cells(6).Value) = "GENERAL" Then
                cmd15.Parameters.AddWithValue("@talla", 0)
            Else
                cmd15.Parameters.AddWithValue("@talla", DataGridView1.Rows(i).Cells(6).Value)
            End If

            Dim sql1 As String = "SELECT dele FROM custom_vianny.dbo.TAB0100 A  Where A.ccia='01' AND A.CTAB='ARTGEN' and cele ='" + Trim(DataGridView1.Rows(i).Cells(2).Value) + "' "
            Dim cmd1 As New SqlCommand(sql1, conx)
            Rs1 = cmd1.ExecuteReader()
            Rs1.Read()

            cmd15.Parameters.AddWithValue("@descripgene", Rs1(0))
            Rs1.Close()
            cmd15.Parameters.AddWithValue("@area", DataGridView1.Rows(i).Cells(1).Value)
            cmd15.Parameters.AddWithValue("@total", DataGridView1.Rows(i).Cells(9).Value)
            cmd15.ExecuteNonQuery()
            '

        Next
        MsgBox("LA INFORMACION SE GUARDO CORRECTAMENTE")
        Dim respuesta As DialogResult


        respuesta = MessageBox.Show("QUIERES IMPRIMIR LA MATRIZ CONSUMO?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Rpt_Explocion.TextBox1.Text = TextBox7.Text
            Rpt_Explocion.TextBox2.Text = Label8.Text
            Rpt_Explocion.Show()

        End If

        DataGridView1.Rows.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
    End Sub
    Dim da55 As New DataTable
    Dim Rsr21 As SqlDataReader
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Rows.Clear()
            da55.Clear()
            abrir()
            Dim sql24 As String = "select COUNT(ncom_8) from custom_vianny.dbo.qamc800 where ncom_8 ='" + Trim(TextBox8.Text) + "' AND ccia_8 ='" + Trim(Label8.Text) + "'"
            Dim cmd24 As New SqlCommand(sql24, conx)
            Rs24 = cmd24.ExecuteReader()
            Rs24.Read()

            If Rs24(0) > 0 Then
                Rs24.Close()

                abrir()

                Dim cmd55 As New SqlDataAdapter("SELECT item_8,t.cele,gene_8,A.linea_17,a.nomb_17 ,factor_8,udm_8,total,talla   FROM custom_vianny.dbo.qag0300 q 
                    inner join custom_vianny.dbo.qamc800 c on q.ccia = c.ccia_8 and q.ncom_3 =c.ncom_8  
                    left join custom_vianny.dbo.cag1700 a on c.linea_8 = a.linea_17 and c.ccia_8 =a.ccia 
                    left join custom_vianny.dbo.TAB0100 t on c.area_8 = t.cele
                    left join custom_vianny.dbo.cag1000 g on q.fich_3 = g.ruc_10  and g.ccia ='" + Trim(Label8.Text) + "'
                    where ncom_3='" + TextBox8.Text + "'and q.ccia ='" + Trim(Label8.Text) + "' and  t.ccia='" + Trim(Label8.Text) + "' AND t.CTAB='AREAMC' order by cast(item_8 as int)", conx)

                cmd55.Fill(da55)
                DataGridView2.DataSource = da55

                DataGridView1.Rows.Add(DataGridView2.Rows.Count)

                For i = 0 To DataGridView2.Rows.Count - 1
                    DataGridView1.Rows(i).Cells(0).Value = i + 1
                    ''DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(1).Value
                    'Dim lsQuery As String = "SELECT cele,dele FROM custom_vianny.dbo.TAB0100 A  Where A.ccia='01' AND A.CTAB='AREAMC'  and dele ='" + Trim(DataGridView2.Rows(i).Cells(1).Value) + "' "
                    'Dim loDataTable As New DataTable
                    'Dim loDataAdapter As New SqlDataAdapter(lsQuery, conx)
                    'loDataAdapter.Fill(loDataTable)
                    'ET.DataSource = loDataTable

                    'ET.DisplayMember = "dele"
                    'ET.ValueMember = "cele"
                    'ET.DataPropertyName = "dele"
                    'ET.DropDownWidth = 200
                    'Dim cmb As New DataGridViewComboBoxColumn()
                    'cmb.DataSource = loDataTable
                    llenar_combo_box()
                    llenar_combo_box1()
                    llenar_combo_box3()
                    DataGridView1.Rows(i).Cells(1).Value = (Trim(DataGridView2.Rows(i).Cells(1).Value)).ToString
                    DataGridView1.Rows(i).Cells(2).Value = (Trim(DataGridView2.Rows(i).Cells(2).Value)).ToString
                    DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(3).Value
                    DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(4).Value
                    DataGridView1.Rows(i).Cells(7).Value = DataGridView2.Rows(i).Cells(5).Value
                    DataGridView1.Rows(i).Cells(8).Value = DataGridView2.Rows(i).Cells(6).Value
                    'DataGridView1.Rows(i).Cells(9).Value = DataGridView2.Rows(i).Cells(7).Value
                    Dim J As String
                    J = "GENERAL"
                    'If Trim(DataGridView2.Rows(i).Cells(8).Value) = 0 Then
                    DataGridView1.Rows(i).Cells(6).Value = (J).ToString
                    'Else
                    '    DataGridView1.Rows(i).Cells(6).Value = (Trim(DataGridView2.Rows(i).Cells(8).Value)).ToString
                    'End If
                    Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_ReporteStockFisico '02','" + Trim(Label10.Text) + "','13','" + Trim(Label10.Text) + "0101','" + Trim(Label10.Text) + "1231','" + Trim(DataGridView1.Rows(i).Cells(4).Value) + "','" + Trim(DataGridView1.Rows(i).Cells(4).Value) + "',NULL, 1"
                    Dim cmd1021 As New SqlCommand(sql1021, conx)
                    Rsr21 = cmd1021.ExecuteReader()

                    If Rsr21.Read() Then
                        DataGridView1.Rows(i).Cells(10).Value = Rsr21(10)
                    Else
                        DataGridView1.Rows(i).Cells(10).Value = 0
                    End If

                    Rsr21.Close()

                Next
            Else
                MsgBox("LA OP NO EXISTE O NO TIENE EXPLOSION DE AVIOS")
            End If
        End If
    End Sub

    Private Sub TextBox7_Resize(sender As Object, e As EventArgs) Handles TextBox7.Resize

    End Sub

    Private Sub TextBox7_LostFocus(sender As Object, e As EventArgs) Handles TextBox7.LostFocus

    End Sub

    Private Sub TextBox8_PaddingChanged(sender As Object, e As EventArgs) Handles TextBox8.PaddingChanged

    End Sub
End Class