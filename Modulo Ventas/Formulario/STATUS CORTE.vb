Imports System.Data.SqlClient
Public Class STATUS_CORTE
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
    Dim da, da1 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            da.Clear()
            Button2.Visible = True
            abrir()
            Dim cmd As New SqlDataAdapter("select	program_3 as PO, ncom_3 as OP,Q.fich_3 AS RUC ,c.nomb_10 as CLIENTE, descri_3 AS PRODUCTO,estilo_3 AS ESTILO,broker_3,OP.kilos AS CONSUMO,g.nomb_17 as TELA, mruta7_3  AS MOLDE,q.modelista_3 as DNI, C1.nomb_10 as MODELISTA,OP.tela  as t,Q.cantp_3 AS PROGRAMADO,OP.kilos* cantp_3 AS KG,a.dele AS VENDEDOR  from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3=c.ruc_10 AND c.tdoc_10<>'01'
left join custom_vianny.dbo.Consumo_Op_DET OP ON OP.ccia = Q.ccia and op.op = q.ncom_3 AND OP.ccia ='01'
left join custom_vianny.dbo.cag1700 g on op.ccia =q.ccia and op.tela = g.linea_17 
LEFT join custom_vianny.dbo.cag1000 c1 on  q.ccia = c1.ccia and q.modelista_3=c1.ruc_10 and c1.tdoc_10='01' AND Q.modelista_3 <> ''
LEFT join custom_vianny.dbo.TAB0100 A ON a.cele= q.merchan_3 and a.CTAB ='MERCHA' and a.ccia='01' 
WHERE  YEAR(Q.fcom_3) = '" + Trim(TextBox1.Text) + "' AND Q.fcom_3 >='20211001' and q.ccia ='01' and (broker_3 ='0001' or  broker_3 ='0000') and OP.cxplineal >0 AND  LEN( RTRIM(g.nomb_17)) > 0 AND  LEN(RTRIM(CAST(Q.mruta7_3 AS VARCHAR(500))))> 0 AND LEN(RTRIM(modelista_3))>0 and hruta12_3 = 0 ORDER BY q.ncom_3 desc", conx)

            cmd.Fill(da)
            If da.Rows.Count > 0 Then
                DataGridView1.DataSource = da
                DataGridView1.Columns(0).Visible = True
                DataGridView1.Columns(3).Visible = False
                DataGridView1.Columns(11).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(13).Visible = False
                DataGridView1.Columns(4).Width = 200
                DataGridView1.Columns(5).Width = 200
                DataGridView1.Columns(9).Width = 220
                DataGridView1.Columns(10).Width = 220
                DataGridView1.Columns(12).Width = 140
                DataGridView1.Columns(0).Frozen = True
                DataGridView1.Columns(1).Frozen = True
                DataGridView1.Columns(2).Frozen = True
                DataGridView1.Columns(3).Frozen = True
                DataGridView1.Columns(4).Frozen = True
                DataGridView1.Columns(1).DefaultCellStyle.BackColor = Color.DarkSalmon
            Else
                DataGridView1.DataSource = Nothing
            End If

        Else
            If RadioButton2.Checked = True Then
                da1.Clear()
                Button2.Visible = False
                abrir()
                Dim cmd1 As New SqlDataAdapter("select program_3 as PO, ncom_3 as OP ,c.nomb_10 as CLIENTE, descri_3 AS PRODUCTO,estilo_3 AS ESTILO,a.dele as VENDEDOR,po.id_corte AS OCORTE,Q.cantp_3 AS PROGRAMADO, '' AS CORTADO  from custom_vianny.dbo.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and q.fich_3=c.ruc_10 AND c.tdoc_10<>'01'
left join custom_vianny.dbo.Consumo_Op_DET OP ON OP.ccia = Q.ccia and op.op = q.ncom_3 AND OP.ccia ='01'
left join custom_vianny.dbo.cag1700 g on op.ccia =q.ccia and op.tela = g.linea_17 
LEFT join custom_vianny.dbo.TAB0100 A ON a.cele= q.merchan_3 and a.CTAB ='MERCHA' and a.ccia='01' 
LEFT join orden_corte_po po   ON po.op= q.ncom_3 and op.ccia= q.ccia
WHERE  YEAR(Q.fcom_3) = '" + Trim(TextBox1.Text) + "' AND Q.fcom_3 >='20211001' and q.ccia ='01' and (broker_3 ='0001' or  broker_3 ='0000') and OP.cxplineal >0 AND  LEN( RTRIM(g.nomb_17)) > 0 AND  LEN(RTRIM(CAST(Q.mruta7_3 AS VARCHAR(500))))> 0 AND LEN(RTRIM(modelista_3))>0 and hruta12_3 = 1 ORDER BY q.ncom_3 desc", conx)

                cmd1.Fill(da1)
                If da1.Rows.Count > 0 Then
                    DataGridView1.DataSource = da1

                    DataGridView1.Columns(0).Visible = False
                    'DataGridView1.Columns(12).Visible = False
                    'DataGridView1.Columns(8).Visible = False
                    'DataGridView1.Columns(14).Visible = False
                    DataGridView1.Columns(3).Width = 235
                    DataGridView1.Columns(4).Width = 235
                    DataGridView1.Columns(6).Width = 170
                    'DataGridView1.Columns(11).Width = 220
                    'DataGridView1.Columns(13).Width = 140
                    DataGridView1.Columns(0).Frozen = False
                    DataGridView1.Columns(1).Frozen = False
                    DataGridView1.Columns(2).Frozen = False
                    DataGridView1.Columns(3).Frozen = False
                    DataGridView1.Columns(4).Frozen = False
                    DataGridView1.Columns(5).Frozen = False
                    DataGridView1.Columns(2).DefaultCellStyle.BackColor = Color.DarkSalmon
                Else
                    DataGridView1.DataSource = Nothing
                End If
            Else
                If RadioButton3.Checked = True Then

                End If
            End If
        End If
    End Sub
    Dim Rsr20, Rsr21, Rsr23 As SqlDataReader
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If RadioButton2.Checked = True Then
            ORDEN_CORTE.Close()
            ORDEN_CORTE.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value
            abrir()
            Dim i5, i6 As Integer
            i5 = 0
            i6 = 0
            Dim sql1020 As String = "select * from orden_corte_po where id_corte ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) + "'"
            Dim cmd1020 As New SqlCommand(sql1020, conx)
            Rsr20 = cmd1020.ExecuteReader()

            While Rsr20.Read()
                ORDEN_CORTE.DataGridView1.Rows.Add()
                ORDEN_CORTE.DataGridView1.Rows(i5).Cells(0).Value = Rsr20(1)
                ORDEN_CORTE.DataGridView1.Rows(i5).Cells(1).Value = Rsr20(2)
                ORDEN_CORTE.DataGridView1.Rows(i5).Cells(2).Value = Rsr20(3)
                ORDEN_CORTE.DataGridView1.Rows(i5).Cells(3).Value = Rsr20(4)
                ORDEN_CORTE.DataGridView1.Rows(i5).Cells(4).Value = Rsr20(5)
                ORDEN_CORTE.DataGridView1.Rows(i5).Cells(5).Value = Rsr20(6)
                i5 = i5 + 1

            End While

            Rsr20.Close()

            abrir()
            Dim sql10201 As String = "select * from orden_corte_tela where id_corte ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) + "'"
            Dim cmd10201 As New SqlCommand(sql10201, conx)
            Rsr21 = cmd10201.ExecuteReader()

            While Rsr21.Read()
                ORDEN_CORTE.DataGridView2.Rows.Add()
                ORDEN_CORTE.DataGridView2.Rows(i6).Cells(0).Value = Rsr21(1)
                ORDEN_CORTE.DataGridView2.Rows(i6).Cells(1).Value = Rsr21(2)
                ORDEN_CORTE.DataGridView2.Rows(i6).Cells(2).Value = Rsr21(3)
                i6 = i6 + 1
            End While

            Rsr21.Close()

            Dim sql102012 As String = "select * from orden_corte where id_corte ='" + Trim(DataGridView1.Rows(e.RowIndex).Cells(7).Value) + "'"
            Dim cmd102012 As New SqlCommand(sql102012, conx)
            Rsr23 = cmd102012.ExecuteReader()

            If Rsr23.Read() = True Then
                ORDEN_CORTE.TextBox3.Text = Rsr23(1)
                ORDEN_CORTE.TextBox5.Text = Rsr23(2)
                ORDEN_CORTE.TextBox2.Text = Rsr23(3)
                ORDEN_CORTE.TextBox6.Text = Rsr23(4)
                ORDEN_CORTE.TextBox4.Text = Rsr23(5)
                ORDEN_CORTE.TextBox7.Text = Rsr23(6)
                ORDEN_CORTE.TextBox8.Text = Rsr23(7)
                ORDEN_CORTE.DateTimePicker1.Value = Rsr23(8)
                ORDEN_CORTE.Label8.Text = Rsr23(10)
                ORDEN_CORTE.Label12.Text = Rsr23(9)
            End If



            Rsr23.Close()
            ORDEN_CORTE.Label5.Text = 1
            ORDEN_CORTE.Button1.Enabled = False
            ORDEN_CORTE.ShowDialog()
        End If

    End Sub

    Sub cargar()
        MsgBox("eliminar")
        DataGridView1.Rows(0).Cells(0).Value = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim o As Integer
        ORDEN_CORTE.Close()
        o = DataGridView1.Rows.Count
        Dim k, L1, l2 As Int16
        k = 0
        l2 = 0
        For pl = 0 To o - 1

            If DataGridView1.Rows(pl).Cells(0).Value = True Then
                If l2 = 0 Then
                    TextBox2.Text = DataGridView1.Rows(pl).Cells(1).Value
                End If

                If Trim(DataGridView1.Rows(pl).Cells(1).Value) <> TextBox2.Text Then
                    k = k + 1
                End If
                l2 = l2 + 1
            End If

        Next

        If k = 0 Then
            For i = 0 To o - 1
                If DataGridView1.Rows(i).Cells(0).Value = True Then
                    ORDEN_CORTE.TextBox2.Text = DataGridView1.Rows(i).Cells(1).Value
                    ORDEN_CORTE.TextBox4.Text = DataGridView1.Rows(i).Cells(4).Value
                    ORDEN_CORTE.TextBox7.Text = DataGridView1.Rows(i).Cells(12).Value
                    ORDEN_CORTE.TextBox8.Text = DataGridView1.Rows(i).Cells(10).Value
                    ORDEN_CORTE.TextBox6.Text = DataGridView1.Rows(i).Cells(16).Value
                    ORDEN_CORTE.DataGridView1.Rows.Add()
                    ORDEN_CORTE.DataGridView1.Rows(i).Cells(0).Value = DataGridView1.Rows(i).Cells(2).Value
                    ORDEN_CORTE.DataGridView1.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(5).Value
                    ORDEN_CORTE.DataGridView1.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(6).Value
                    ORDEN_CORTE.DataGridView1.Rows(i).Cells(3).Value = DataGridView1.Rows(i).Cells(14).Value
                    ORDEN_CORTE.DataGridView1.Rows(i).Cells(4).Value = DataGridView1.Rows(i).Cells(8).Value
                    ORDEN_CORTE.DataGridView1.Rows(i).Cells(5).Value = DataGridView1.Rows(i).Cells(15).Value
                    'ORDEN_CORTE.DataGridView2.Rows(i).Cells(0).Value =

                End If
            Next
            Dim j As Integer
            Dim suma As Double
            j = 0
            For j = 0 To o - 1
                If DataGridView1.Rows(j).Cells(0).Value = True Then
                    Dim p As Integer
                    p = ORDEN_CORTE.DataGridView2.Rows.Count
                    If p = 0 Then
                        ORDEN_CORTE.DataGridView2.Rows.Add()

                        ORDEN_CORTE.DataGridView2.Rows(j).Cells(0).Value = DataGridView1.Rows(j).Cells(13).Value
                        ORDEN_CORTE.DataGridView2.Rows(j).Cells(1).Value = DataGridView1.Rows(j).Cells(9).Value
                        ORDEN_CORTE.DataGridView2.Rows(j).Cells(2).Value = DataGridView1.Rows(j).Cells(15).Value
                    Else
                        For l = 0 To p - 1
                            If Trim(DataGridView1.Rows(j).Cells(13).Value) = Trim(ORDEN_CORTE.DataGridView2.Rows(l).Cells(0).Value) Then
                                ORDEN_CORTE.DataGridView2.Rows(l).Cells(2).Value = ORDEN_CORTE.DataGridView2.Rows(l).Cells(2).Value + DataGridView1.Rows(j).Cells(15).Value


                            Else

                                ORDEN_CORTE.DataGridView2.Rows.Add()
                                ORDEN_CORTE.DataGridView2.Rows(l + 1).Cells(0).Value = DataGridView1.Rows(j).Cells(13).Value
                                ORDEN_CORTE.DataGridView2.Rows(l + 1).Cells(1).Value = DataGridView1.Rows(j).Cells(9).Value
                            End If
                        Next
                    End If
                End If

            Next
            ORDEN_CORTE.Label8.Text = TextBox1.Text

            ORDEN_CORTE.Label12.Text = Label2.Text
            ORDEN_CORTE.Show()
        Else
            MsgBox("SOLO SE PERMITE AGREGRAR OP QUE PERTENECEN A LA MISMA PO")
        End If







        'Dim datosenlista As Boolean
        'datosenlista = False

        'Dim valorstring As String
        'Dim fila As New DataGridViewRow
        'fila = DataGridView1.SelectedRows(13)
        'valorstring = fila.Cells(0).Value.ToString

        'For Each filab As DataGridViewRow In DataGridView1.Rows

        '    If filab.Cells(0).Value.ToString.Equals(valorstring) Then
        '        MsgBox("repetido")
        '        datosenlista = True
        '        Exit For
        '    End If

        'Next
        'If datosenlista <> True Then
        '    ORDEN_CORTE.DataGridView2.Rows.Add()
        '    ORDEN_CORTE.DataGridView2.Rows(j).Cells(0).Value = DataGridView1.Rows(j).Cells(8).Value
        'End If





        'If Convert.ToString(Trim(DataGridView1.Rows(j).Cells(13).Value)).Equals(Trim(DataGridView1.CurrentRow.Cells(j + 1).Value)) Then
        '    ORDEN_CORTE.DataGridView2.Rows.Add()
        '    suma = suma + DataGridView1.Rows(j).Cells(8).Value
        '    MsgBox(Row.Cells("t").Value)
        '    MsgBox("bien")
        'Else
        '    MsgBox(Row.Cells("t").Value)
        '    MsgBox("mal")
        '    ORDEN_CORTE.DataGridView2.Rows.Add()
        '    ORDEN_CORTE.DataGridView2.Rows(j).Cells(0).Value = DataGridView1.Rows(j).Cells(8).Value
        'End If
        'ORDEN_CORTE.DataGridView2.Rows(j).Cells(0).Value = suma
        ''j = j + 1

    End Sub
    Private Function EstadoCheck(ByVal filaposicion As Integer) As Boolean
        Return Me.DataGridView1.Rows(filaposicion).Cells("A").Value
    End Function

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class