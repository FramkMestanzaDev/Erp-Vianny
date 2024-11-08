Imports System.Data.SqlClient
Public Class Form_Cliente
    Dim dt, da, da1, da2 As New DataTable
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
    Private Sub Form_Cliente_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
        TextBox1.Text = ""
        dt.Clear()
        DataGridView1.DataSource = Nothing
        If Label1.Text = "PAIS" Then
            Me.Text = "PAIS"
            Dim dy As New fdiciosionpais
            Dim dy1 As New VPAIS_CO

            dy1.gruc = TextBox2.Text
            dy1.gorden = 2
            dt = dy.buscar_cliente_comerial(dy1)
            If dt.Rows.Count <> 0 Then
                DataGridView1.DataSource = dt
                DataGridView1.Columns(0).Width = 80
                DataGridView1.Columns(1).Width = 220
            End If

        Else
            If Label1.Text = "DIVISION" Then
                Me.Text = "DIVISION"
                Dim dy As New fdiciosionpais
                Dim dy1 As New VPAIS_CO

                dy1.gruc = TextBox2.Text
                dy1.gorden = 2
                dt = dy.buscar_divicion_comerial(dy1)
                If dt.Rows.Count <> 0 Then
                    DataGridView1.DataSource = dt
                    DataGridView1.Columns(0).Width = 80
                    DataGridView1.Columns(1).Width = 220
                End If
            Else
                If Label1.Text = "DIVISIONOD" Then
                    Me.Text = "DIVISION"
                    abrir()
                    Dim cmd As New SqlDataAdapter("SELECT  LEFT( CELE,4) AS CODIGO, dele AS DESCRIPCION FROM custom_vianny.dbo.TAb0100 WHERE ccia = '01' AND CTAB='DIVCLI' ", conx)
                    cmd.Fill(dt)
                    DataGridView1.DataSource = dt
                    If dt.Rows.Count <> 0 Then
                        DataGridView1.DataSource = dt
                        DataGridView1.Columns(0).Width = 80
                        DataGridView1.Columns(1).Width = 220

                    End If
                Else
                    If Label1.Text = "MARCA" Then
                        Me.Text = "MARCA"
                        abrir()
                        Dim cmd As New SqlDataAdapter("SELECT A.cele AS CODIGO, a.dele AS DESCRIPCION,codigo FROM  custom_vianny.dbo.TAB0100 A  Where CCia = '01' AND CTab = 'MARCLI' and  codigo2 ='1' order by cele", conx)
                        cmd.Fill(dt)
                        DataGridView1.DataSource = dt
                        If dt.Rows.Count <> 0 Then
                            DataGridView1.DataSource = dt
                            DataGridView1.Columns(0).Width = 80
                            DataGridView1.Columns(1).Width = 220
                            DataGridView1.Columns(2).Visible = False
                        End If
                    Else
                        If Label1.Text = "PAIS_OD" Then
                            Me.Text = "PAIS"
                            abrir()
                            Dim cmd As New SqlDataAdapter("SELECT LEFT( A.CELE,4) AS CODIGO , A.DELE AS DESCRIPCION FROM custom_vianny.dbo.TAb0100 A WHERE A.ccia = '01' AND  A.CTAB='ZONA '", conx)
                            cmd.Fill(dt)
                            'DataGridView1.DataSource = dt
                            If dt.Rows.Count <> 0 Then
                                DataGridView1.DataSource = dt
                                DataGridView1.Columns(0).Width = 80
                                DataGridView1.Columns(1).Width = 220

                            End If
                        Else
                            Dim dy1 As New vfpago
                            dt = dy1.forma_pago()
                            If dt.Rows.Count <> 0 Then
                                DataGridView1.DataSource = dt
                                DataGridView1.Columns(0).Width = 80
                                DataGridView1.Columns(1).Width = 220
                                DataGridView1.Columns(2).Visible = False
                            End If
                        End If

                    End If

                End If

            End If


        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 80
                DataGridView1.Columns(1).Width = 220
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
    Dim Rsr2 As SqlDataReader
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub

    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then


            With DataGridView1
                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If


                If Label1.Text = "PAIS" Then
                    Op_Manufactura.TextBox4.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                    Op_Manufactura.TextBox24.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value

                Else
                    If Label1.Text = "DIVISION" Then
                        Op_Manufactura.TextBox3.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                        Op_Manufactura.TextBox23.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                    Else
                        If Label1.Text = "F. PAGO" Then

                            If TextBox3.Text = 1 Then
                                Op_Manufactura.TextBox19.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                                Op_Manufactura.TextBox25.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                            Else

                                If TextBox3.Text = 3 Then
                                    Nota_Pedido.TextBox20.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                                    Nota_Pedido.TextBox21.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value


                                Else
                                    Registro_Facturas.TextBox23.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                                    Registro_Facturas.TextBox20.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                                    Registro_Facturas.TextBox24.Text = DataGridView1.Rows(hti.RowIndex).Cells(2).Value
                                    Registro_Facturas.DateTimePicker2.Value = DateAdd(DateInterval.Day, DataGridView1.Rows(hti.RowIndex).Cells(2).Value, Registro_Facturas.DateTimePicker1.Value)
                                End If



                            End If
                        Else
                            If Label1.Text = "MARCA" Then
                                Dim po As String
                                OD.TextBox3.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                                OD.TextBox72.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                                po = Microsoft.VisualBasic.Mid(OD.TextBox81.Text, 1, 2) & Trim(DataGridView1.Rows(hti.RowIndex).Cells(2).Value)
                                OD.TextBox81.Text = ""
                                OD.TextBox81.Text = po

                                Dim corela As Integer
                                corela = 0

                                Dim bpo As String
                                bpo = Trim(OD.TextBox82.Text) & Trim(OD.TextBox81.Text)
                                Dim sql102 As String = "select top 1 program_3, CAST( substring(program_3,7,4) as int) as numero from custom_vianny.dbo.qag0300  where  program_3 like  '" + bpo + "%' and ccia='01' AND SUBSTRING(ncom_3,1,2)<> '02'  order by numero desc"
                                Dim cmd102 As New SqlCommand(sql102, conx)
                                Rsr2 = cmd102.ExecuteReader()
                                If Rsr2.Read() Then
                                    corela = Convert.ToInt32(Rsr2(1)) + 1
                                Else
                                    corela = 1
                                End If
                                Rsr2.Close()

                                Select Case Convert.ToString(corela).Length
                                    Case "1" : OD.TextBox8.Text = "000" & corela
                                    Case "2" : OD.TextBox8.Text = "00" & corela
                                    Case "3" : OD.TextBox8.Text = "0" & corela
                                    Case "4" : OD.TextBox8.Text = corela
                                End Select
                                OD.TextBox4.Select()
                            Else
                                If Label1.Text = "DIVISIONOD" Then
                                    OD.TextBox4.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                                    OD.TextBox73.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                                    OD.TextBox5.Select()
                                Else
                                    If Label1.Text = "PAIS_OD" Then
                                        OD.TextBox5.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                                        OD.TextBox74.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                                        OD.TextBox9.Select()
                                    End If
                                End If

                            End If
                        End If
                    End If


                End If
            End With
            Me.Close()
        End If

    End Sub
End Class