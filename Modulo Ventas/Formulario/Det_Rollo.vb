Imports System.Data.SqlClient
Public Class Det_Rollo
    Public conx As SqlConnection
    Public comando As SqlCommand
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
    Private Sub buscar()
        Try


            If Label2.Text = 1 Then
                Dim ds As New DataSet
                ds.Tables.Add(dt10.Copy)
                Dim dv As New DataView(ds.Tables(0))
                dv.RowFilter = "NOMBRE" & " like '%" & TextBox1.Text & "%'"
                If dv.Count <> 0 Then
                    DataGridView1.DataSource = dv
                    DataGridView1.Columns(1).Width = 200
                Else
                    DataGridView1.DataSource = Nothing
                End If
            Else
                If Label2.Text = 2 Then
                    Dim ds As New DataSet
                    ds.Tables.Add(DT12.Copy)
                    Dim dv As New DataView(ds.Tables(0))
                    dv.RowFilter = "TRABAJADOR" & " like '%" & TextBox1.Text & "%'"
                    If dv.Count <> 0 Then
                        DataGridView1.DataSource = dv
                        DataGridView1.Columns(1).Width = 310
                    Else
                        DataGridView1.DataSource = Nothing
                    End If
                Else
                    If Label2.Text = 3 Then
                        Dim ds As New DataSet
                        ds.Tables.Add(DT13.Copy)
                        Dim dv As New DataView(ds.Tables(0))
                        dv.RowFilter = "TRABAJADOR" & " like '%" & TextBox1.Text & "%'"
                        If dv.Count <> 0 Then
                            DataGridView1.DataSource = dv
                            DataGridView1.Columns(2).Width = 200
                            DataGridView1.Columns(3).Visible = False
                        Else
                            DataGridView1.DataSource = Nothing
                        End If
                    Else
                        If Label2.Text = 4 Then
                            Dim ds As New DataSet
                            ds.Tables.Add(DT14.Copy)
                            Dim dv As New DataView(ds.Tables(0))
                            dv.RowFilter = "PO" & " like '%" & TextBox1.Text & "%'"
                            If dv.Count <> 0 Then
                                DataGridView1.DataSource = dv
                                DataGridView1.Columns(1).Width = 400
                            Else
                                DataGridView1.DataSource = Nothing
                            End If
                        Else
                            If Label2.Text = 5 Then
                                Dim ds As New DataSet
                                ds.Tables.Add(DT15.Copy)
                                Dim dv As New DataView(ds.Tables(0))
                                dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"
                                If dv.Count <> 0 Then
                                    DataGridView1.DataSource = dv
                                    DataGridView1.Columns(1).Width = 310
                                Else
                                    DataGridView1.DataSource = Nothing
                                End If
                            Else
                                If Label2.Text = 7 Then

                                    Dim ds As New DataSet
                                    ds.Tables.Add(DT11.Copy)
                                    Dim dv As New DataView(ds.Tables(0))
                                    dv.RowFilter = "PO" & " like '%" & TextBox1.Text & "%'"
                                    If dv.Count <> 0 Then
                                        DataGridView1.DataSource = dv
                                        DataGridView1.Columns(1).Width = 400
                                    Else
                                        DataGridView1.DataSource = Nothing
                                    End If
                                Else
                                    If Label2.Text = 6 Then
                                        Dim ds As New DataSet
                                        ds.Tables.Add(DT16.Copy)
                                        Dim dv As New DataView(ds.Tables(0))
                                        dv.RowFilter = "PO" & " like '%" & TextBox1.Text & "%'"
                                        If dv.Count <> 0 Then
                                            DataGridView1.DataSource = dv
                                            DataGridView1.Columns(1).Width = 400
                                        Else
                                            DataGridView1.DataSource = Nothing
                                        End If
                                    Else
                                        If Label2.Text = 44 Then
                                            Dim ds As New DataSet
                                            ds.Tables.Add(DT14.Copy)
                                            Dim dv As New DataView(ds.Tables(0))
                                            dv.RowFilter = "PO" & " like '%" & TextBox1.Text & "%'"
                                            If dv.Count <> 0 Then
                                                DataGridView1.DataSource = dv
                                                DataGridView1.Columns(1).Width = 400
                                            Else
                                                DataGridView1.DataSource = Nothing
                                            End If
                                        End If
                                    End If
                                    End If
                            End If
                        End If
                    End If
                End If
            End If





        Catch ex As Exception


        End Try
    End Sub
    Dim dt10, DT11, DT12, DT13, DT14, DT15, DT16 As New DataTable
    Private Sub Det_Rollo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
        abrir()
        dt10.Clear()

        If Label2.Text = 1 Then
            Dim cmd5 As New SqlDataAdapter("SELECT codigo_mq as CODIGO,nombre_mq AS NOMBRE, sigla_mq AS SIGLA FROM custom_vianny.DBO.maquinas", conx)

            cmd5.Fill(dt10)
            DataGridView1.DataSource = dt10
            DataGridView1.Columns(1).Width = 200
        Else
            If Label2.Text = 2 Then

                Dim cmd6 As New SqlDataAdapter(" Select  id_teje As CODIGO,DNI As DNI,trabajador As TRABAJADOR FROM tejedores", conx)

                cmd6.Fill(DT12)
                DataGridView1.DataSource = DT12
                DataGridView1.Columns(2).Width = 310
            Else
                If Label2.Text = 3 Then

                    Dim cmd6 As New SqlDataAdapter("SELECT id_audit as CODIGO, dni AS DNI,trabajador AS TRABAJADOR,clave as CLAVE FROM auditores", conx)
                    Dim dt10 As New DataTable
                    cmd6.Fill(DT13)
                    DataGridView1.DataSource = DT13
                    DataGridView1.Columns(2).Width = 200
                    DataGridView1.Columns(3).Visible = False
                Else
                    If Label2.Text = 4 Then

                        Dim cmd6 As New SqlDataAdapter("SELECT distinct A.program_3 AS PO, ISNULL(B.nomb_10,'') AS CLIENTE  FROM custom_vianny.dbo.QAG0300 A LEFT JOIN custom_vianny.dbo.CAG1000 B  ON '01' = B.CCIA AND A.FICH_3 = B.FICH_10  Where A.CCia='01' AND YEAR(fcom_3) >= '2021' AND merchan_3 IN (0025,0028,0023,0009) ORDER BY program_3", conx)
                        Dim dt10 As New DataTable
                        cmd6.Fill(DT14)
                        DataGridView1.DataSource = DT14
                        DataGridView1.Columns(1).Width = 400
                    Else
                        If Label2.Text = 5 Then

                            Dim cmd6 As New SqlDataAdapter(" Select distinct  id_teje As CODIGO,DNI As DNI,trabajador As TRABAJADOR FROM tejedores", conx)
                            Dim dt10 As New DataTable
                            cmd6.Fill(DT15)
                            DataGridView1.DataSource = DT15
                            DataGridView1.Columns(2).Width = 310
                        Else
                            If Label2.Text = 6 Then

                                Dim cmd6 As New SqlDataAdapter("SELECT distinct A.program_3 AS PO, ISNULL(B.nomb_10,'') AS CLIENTE  FROM custom_vianny.dbo.QAG0300 A LEFT JOIN custom_vianny.dbo.CAG1000 B  ON '01' = B.CCIA AND A.FICH_3 = B.FICH_10  Where A.CCia='01' AND YEAR(fcom_3) >= '2021' AND merchan_3 IN (0025,0028,0023,0009) ORDER BY program_3", conx)
                                Dim dt10 As New DataTable
                                cmd6.Fill(DT15)
                                DataGridView1.DataSource = DT15
                                DataGridView1.Columns(1).Width = 400
                            Else
                                If Label2.Text = 31 Then

                                    Dim cmd6 As New SqlDataAdapter("SELECT id_audit as CODIGO, dni AS DNI,trabajador AS TRABAJADOR,clave as CLAVE FROM auditores", conx)
                                    Dim dt10 As New DataTable
                                    cmd6.Fill(dt10)
                                    DataGridView1.DataSource = dt10
                                    DataGridView1.Columns(2).Width = 200
                                    DataGridView1.Columns(3).Visible = False
                                Else
                                    If Label2.Text = 7 Then
                                        Dim cmd6 As New SqlDataAdapter("SELECT distinct A.program_3 AS PO, ISNULL(B.nomb_10,'') AS CLIENTE  FROM custom_vianny.dbo.QAG0300 A LEFT JOIN custom_vianny.dbo.CAG1000 B  ON '01' = B.CCIA AND A.FICH_3 = B.FICH_10  Where A.CCia='01' AND YEAR(fcom_3) >= '2021' AND merchan_3 IN (0025,0028,0023,0009) ORDER BY program_3", conx)
                                        Dim dt10 As New DataTable
                                        cmd6.Fill(DT11)
                                        DataGridView1.DataSource = DT11
                                        DataGridView1.Columns(1).Width = 400
                                    Else
                                        If Label2.Text = 44 Then

                                            Dim cmd6 As New SqlDataAdapter("SELECT distinct A.program_3 AS PO, ISNULL(B.nomb_10,'') AS CLIENTE  FROM custom_vianny.dbo.QAG0300 A LEFT JOIN custom_vianny.dbo.CAG1000 B  ON '01' = B.CCIA AND A.FICH_3 = B.FICH_10  Where A.CCia='01' AND YEAR(fcom_3) >= '2021' AND merchan_3 IN (0025,0028,0023,0009) ORDER BY program_3", conx)
                                            Dim dt10 As New DataTable
                                            cmd6.Fill(DT14)
                                            DataGridView1.DataSource = DT14
                                            DataGridView1.Columns(1).Width = 400
                                        Else
                                            If Label2.Text = 45 Then

                                                Dim cmd6 As New SqlDataAdapter("select maqui_22q as CODIGO,desc_22q AS NOMBRE,ubica_22q as UBIC  from custom_vianny.dbo.qaq2200 where ubica_22q ='TEJEDURIA' order by maqui_22q", conx)
                                                Dim dt105 As New DataTable
                                                cmd6.Fill(DT105)
                                                DataGridView1.DataSource = DT105
                                                DataGridView1.Columns(1).Width = 200
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Dim ty, ty1 As New DataTable
    Sub llenar_combo_box(ByVal cb As ComboBox, ByVal tejedo As String)
        Try

            conn = New SqlDataAdapter("SELECT galga FROM galga where tejedora ='" + tejedo + "'", conx)
            conn.Fill(ty)
            pesaod_rollo.ComboBox2.DataSource = ty
            pesaod_rollo.ComboBox2.DisplayMember = "galga"
            pesaod_rollo.ComboBox2.ValueMember = "galga"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box1(ByVal cb As ComboBox, ByVal alimen As String)
        Try

            conn = New SqlDataAdapter("SELECT alimentador FROM alimentador where tejedora ='" + alimen + "'", conx)
            conn.Fill(ty1)
            pesaod_rollo.ComboBox3.DataSource = ty1
            pesaod_rollo.ComboBox3.DisplayMember = "alimentador"
            pesaod_rollo.ComboBox3.ValueMember = "alimentador"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If Label2.Text = 1 Then
            pesaod_rollo.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
            Dim valor As String
            Select Case Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value).Length
                Case "1" : valor = "000" + Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                Case "2" : valor = "00" + Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                Case "3" : valor = "0" + Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                Case "4" : valor = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)

            End Select

            llenar_combo_box(pesaod_rollo.ComboBox2, valor)
            llenar_combo_box1(pesaod_rollo.ComboBox3, valor)
            Me.Close()
        Else
            If Label2.Text = 2 Then
                pesaod_rollo.TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                pesaod_rollo.TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
                Me.Close()
            Else
                If Label2.Text = 3 Then
                    Dim clave, CLAVE1 As String

                    CLAVE1 = DataGridView1.Rows(e.RowIndex).Cells(3).Value
                    clave = InputBox("Introduzca la Clave del Auditor")

                    If Trim(clave) = Trim(CLAVE1) Then
                        Calidad_tela_cruda.TextBox28.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                        Calidad_tela_cruda.TextBox25.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value

                        Me.Close()
                    Else
                        MsgBox("CONTRASEÑA INCORRECTA")
                    End If
                Else
                    If Label2.Text = 4 Then
                        REQUERIMIENTO_PO.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value

                        Me.Close()
                    Else
                        If Label2.Text = 5 Then
                            ProddiariaTej.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                            ProddiariaTej.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
                            Me.Close()
                        Else
                            If Label2.Text = 6 Then
                                ProddiariaTej.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                                Me.Close()
                            Else
                                If Label2.Text = 31 Then
                                    Dim clave, CLAVE1 As String

                                    CLAVE1 = DataGridView1.Rows(e.RowIndex).Cells(3).Value
                                    clave = InputBox("Introduzca la Clave del Auditor")

                                    If Trim(clave) = Trim(CLAVE1) Then
                                        Form4.TextBox23.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                                        Form4.TextBox12.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value

                                        Me.Close()
                                    Else
                                        MsgBox("CONTRASEÑA INCORRECTA")
                                    End If
                                Else
                                    If Label2.Text = 7 Then
                                        EMISION_PARTIDAS.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                                        EMISION_PARTIDAS.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                                        Me.Close()
                                    Else
                                        If Label2.Text = 44 Then
                                            CREAR_PARTIDAS.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                                            CREAR_PARTIDAS.Button4.Enabled = True
                                            Me.Close()
                                        Else
                                            If Label2.Text = 45 Then
                                                Dim o As Integer
                                                o = Label3.Text
                                                Otejeduria.DataGridView1.Rows(o).Cells(1).Value = ""
                                                Otejeduria.DataGridView1.Rows(o).Cells(1).Value = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                                                Otejeduria.DataGridView1.CurrentCell = Otejeduria.DataGridView1.Item(0, o)

                                                Me.Close()
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode = Keys.Enter Then
            'MsgBox("bien")
            If Label2.Text = 1 Then
                Dim pos As Integer
                pos = DataGridView1.CurrentCell.RowIndex
                pesaod_rollo.TextBox2.Text = DataGridView1.Rows(pos).Cells(2).Value
                Dim valor As String
                Select Case Trim(DataGridView1.Rows(pos).Cells(0).Value).Length
                    Case "1" : valor = "000" + Trim(DataGridView1.Rows(pos).Cells(0).Value)
                    Case "2" : valor = "00" + Trim(DataGridView1.Rows(pos).Cells(0).Value)
                    Case "3" : valor = "0" + Trim(DataGridView1.Rows(pos).Cells(0).Value)
                    Case "4" : valor = Trim(DataGridView1.Rows(pos).Cells(0).Value)

                End Select

                llenar_combo_box(pesaod_rollo.ComboBox2, valor)
                llenar_combo_box1(pesaod_rollo.ComboBox3, valor)
                Me.Close()
            Else
                If Label2.Text = 2 Then
                    Dim pos As Integer
                    pos = DataGridView1.CurrentCell.RowIndex
                    pesaod_rollo.TextBox3.Text = DataGridView1.Rows(pos).Cells(1).Value
                    pesaod_rollo.TextBox4.Text = DataGridView1.Rows(pos).Cells(2).Value
                    pesaod_rollo.TextBox6.Select()
                    Me.Close()
                Else
                    If Label2.Text = 3 Then
                        Dim clave, CLAVE1 As String
                        Dim pos As Integer
                        pos = DataGridView1.CurrentCell.RowIndex
                        CLAVE1 = DataGridView1.Rows(pos).Cells(3).Value
                        clave = InputBox("Introduzca la Clave del Auditor")

                        If Trim(clave) = Trim(CLAVE1) Then

                            Calidad_tela_cruda.TextBox28.Text = DataGridView1.Rows(pos).Cells(1).Value
                            Calidad_tela_cruda.TextBox25.Text = DataGridView1.Rows(pos).Cells(2).Value

                            Me.Close()
                        Else
                            MsgBox("CONTRASEÑA INCORRECTA")
                        End If
                    Else
                        If Label2.Text = 4 Then
                            Dim pos As Integer
                            pos = DataGridView1.CurrentCell.RowIndex
                            REQUERIMIENTO_PO.TextBox1.Text = DataGridView1.Rows(pos).Cells(0).Value

                            Me.Close()
                        Else
                            If Label2.Text = 5 Then
                                Dim pos As Integer
                                pos = DataGridView1.CurrentCell.RowIndex
                                ProddiariaTej.TextBox3.Text = Trim(DataGridView1.Rows(pos).Cells(1).Value)
                                ProddiariaTej.TextBox2.Text = Trim(DataGridView1.Rows(pos).Cells(2).Value)
                                Me.Close()
                            Else
                                If Label2.Text = 6 Then
                                    Dim pos As Integer
                                    pos = DataGridView1.CurrentCell.RowIndex
                                    ProddiariaTej.TextBox1.Text = Trim(DataGridView1.Rows(pos).Cells(0).Value)
                                    Me.Close()
                                Else
                                    If Label2.Text = 31 Then
                                        Dim clave, CLAVE1 As String
                                        Dim pos As Integer
                                        pos = DataGridView1.CurrentCell.RowIndex
                                        CLAVE1 = DataGridView1.Rows(pos).Cells(3).Value
                                        clave = InputBox("Introduzca la Clave del Auditor")

                                        If Trim(clave) = Trim(CLAVE1) Then
                                            Form4.TextBox23.Text = DataGridView1.Rows(pos).Cells(1).Value
                                            Form4.TextBox12.Text = DataGridView1.Rows(pos).Cells(2).Value

                                            Me.Close()
                                        Else
                                            MsgBox("CONTRASEÑA INCORRECTA")
                                        End If
                                    Else
                                        If Label2.Text = 7 Then
                                            Dim pos As Integer
                                            pos = DataGridView1.CurrentCell.RowIndex
                                            EMISION_PARTIDAS.TextBox1.Text = DataGridView1.Rows(pos).Cells(0).Value
                                            EMISION_PARTIDAS.TextBox2.Text = DataGridView1.Rows(pos).Cells(1).Value
                                            Me.Close()
                                        Else
                                            If Label2.Text = 44 Then
                                                Dim pos As Integer
                                                pos = DataGridView1.CurrentCell.RowIndex
                                                CREAR_PARTIDAS.TextBox2.Text = DataGridView1.Rows(pos).Cells(0).Value
                                                CREAR_PARTIDAS.Button4.Enabled = True
                                                Me.Close()
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
End Class