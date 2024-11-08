Imports System.Data.SqlClient
Public Class Listar_Clientes
    Dim dt As New DataTable
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim Rsr2, Rsr3 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Listar_Clientes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hu As New fcliente
        Dim jk As New vcliente
        Dim va11, va21 As String
        va11 = ""
        va21 = ""
        dt.Clear()

        If Trim(TextBox3.Text) <> "A" Then
            abrir()

            Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + Trim(TextBox3.Text) + "' AND ccia_ven ='01' "
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()



            If Rsr2.Read() = True Then
                va11 = Rsr2(0)
            End If
            Rsr2.Close()
            jk.gVendedor = va11

            dt = hu.mostrar_cliente(jk)
        Else
            abrir()
            Dim cmd6 As New SqlDataAdapter("select COD_CLI as Ruc,CASE WHEN t_doc= 2 then r_social else replace(apaterno,' ','')+ '  ' + replace(amaterno,' ','')+ ','+ replace(nombres,' ','') end as Cliente, codigo as Codigo,
           D_fiscal,email,telefono,Vendedor,sigla_10  from cliente C LEFT JOIN custom_vianny.dbo.cag1000 G  ON C.COD_CLI = G.fich_10 AND G.ccia ='01' ", conx)
            cmd6.Fill(dt)

        End If

        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(2).Width = 405
            If TextBox2.Text = 6 Then
                DataGridView1.Columns(0).Visible = False
                Button1.Visible = False
            Else
                DataGridView1.Columns(0).Visible = True
                Button1.Visible = True
            End If
        End If




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, conta As Integer
        i = DataGridView1.Rows.Count
        a = 0
        conta = 0
        If TextBox2.Text = 1 Then
            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                    conta = conta + 1
                End If
            Next
            If conta > 1 Then
                MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
            Else
                For a = 0 To i - 1
                    If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                        Nota_Pedido.TextBox2.Text = DataGridView1.Rows(a).Cells(1).Value
                        Nota_Pedido.TextBox3.Text = DataGridView1.Rows(a).Cells(2).Value
                        Nota_Pedido.TextBox4.Text = DataGridView1.Rows(a).Cells(4).Value
                        Nota_Pedido.TextBox5.Text = DataGridView1.Rows(a).Cells(6).Value
                        Nota_Pedido.TextBox6.Text = DataGridView1.Rows(a).Cells(5).Value
                        Nota_Pedido.TextBox8.Text = DataGridView1.Rows(a).Cells(7).Value


                        Nota_Pedido.TextBox11.Text = DataGridView1.Rows(a).Cells(3).Value
                    End If
                Next
            End If
            TextBox2.Text = 1
            Close()
        Else
            If TextBox2.Text = 2 Then
                For a = 0 To i - 1
                    If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                        conta = conta + 1
                    End If
                Next
                If conta > 1 Then
                    MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                Else
                    For a = 0 To i - 1
                        If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                            Orden_Produccion.TextBox9.Text = DataGridView1.Rows(a).Cells(1).Value
                            Orden_Produccion.TextBox2.Text = DataGridView1.Rows(a).Cells(2).Value
                        End If
                    Next
                End If

                Close()
            Else
                If TextBox2.Text = 5 Then
                    For a = 0 To i - 1
                        If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                            conta = conta + 1
                        End If
                    Next
                    If conta > 1 Then
                        MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                    Else
                        For a = 0 To i - 1
                            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                Op_Manufactura.TextBox22.Text = DataGridView1.Rows(a).Cells(1).Value
                                Op_Manufactura.TextBox2.Text = DataGridView1.Rows(a).Cells(2).Value
                            End If
                        Next
                    End If

                    Close()
                Else
                    If TextBox2.Text = 19 Then
                        For a = 0 To i - 1
                            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                conta = conta + 1
                            End If
                        Next
                        If conta > 1 Then
                            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                        Else
                            For a = 0 To i - 1
                                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                    GESTION_ACTIVIDADES.TextBox2.Text = DataGridView1.Rows(a).Cells(1).Value
                                    GESTION_ACTIVIDADES.TextBox4.Text = DataGridView1.Rows(a).Cells(2).Value
                                End If
                            Next
                        End If

                        Close()
                    Else
                        If TextBox2.Text = "800" Then
                            For a = 0 To i - 1
                                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                    conta = conta + 1
                                End If
                            Next
                            If conta > 1 Then
                                MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                            Else
                                For a = 0 To i - 1
                                    If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                        Hoja_Reclamos.TextBox2.Text = DataGridView1.Rows(a).Cells(1).Value
                                        Hoja_Reclamos.TextBox3.Text = DataGridView1.Rows(a).Cells(2).Value
                                    End If
                                Next
                            End If

                            Close()
                        Else
                            If TextBox2.Text = 11 Then
                                For a = 0 To i - 1
                                    If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                        conta = conta + 1
                                    End If
                                Next
                                If conta > 1 Then
                                    MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                Else
                                    For a = 0 To i - 1
                                        If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                            Rpt_Ultimas_ventas.TextBox4.Text = DataGridView1.Rows(a).Cells(1).Value
                                            Rpt_Ultimas_ventas.TextBox6.Text = DataGridView1.Rows(a).Cells(2).Value

                                        End If
                                    Next

                                End If
                                TextBox2.Text = 1
                                Close()
                            Else
                                If TextBox2.Text = 6 Then
                                    For a = 0 To i - 1
                                        If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                            conta = conta + 1
                                        End If
                                    Next
                                    If conta > 1 Then
                                        MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
                                    Else

                                        For a = 0 To i - 1
                                            If Me.DataGridView1.Rows(a).Cells(8).Value.ToString.Trim.Length = 0 Then
                                                MsgBox("SE TIENE QUE AGREGAR LA ABREVITURA AL CLIENTE, PARA QUE SE GENERE EL CORRELATIVO DE LA PM DE FORMA AUTOMATICA")
                                            Else
                                                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                                                    Dim corela As Integer
                                                    corela = 0
                                                    Dim po As String
                                                    OD.TextBox1.Text = DataGridView1.Rows(a).Cells(2).Value
                                                    OD.TextBox58.Text = DataGridView1.Rows(a).Cells(1).Value
                                                    If Trim(OD.TextBox81.Text).Length > 0 Then
                                                        MsgBox(Microsoft.VisualBasic.Mid(OD.TextBox81.Text, 3, 3))
                                                        MsgBox(Trim(DataGridView1.Rows(a).Cells(8).Value))
                                                        po = Trim(DataGridView1.Rows(a).Cells(8).Value) & Microsoft.VisualBasic.Mid(OD.TextBox81.Text, 3, 3)
                                                        OD.TextBox81.Text = ""
                                                        OD.TextBox81.Text = po
                                                        If Trim(OD.TextBox81.Text.ToString()).Length = 5 Then
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
                                                        End If

                                                    Else
                                                        OD.TextBox81.Text = Trim(DataGridView1.Rows(a).Cells(8).Value)
                                                    End If

                                                End If

                                            End If


                                        Next

                                    End If
                                    TextBox2.Text = 1
                                    Close()
                                End If
                            End If
                        End If
                    End If
                    End If
            End If
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(4).Visible = False
                DataGridView1.Columns(5).Visible = False
                DataGridView1.Columns(6).Visible = False
                DataGridView1.Columns(7).Visible = False
                DataGridView1.Columns(2).Width = 405
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub



    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        If TextBox2.Text = 6 Then
            If e.Button = Windows.Forms.MouseButtons.Right Then
                With DataGridView1

                    Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                    If hti.Type = DataGridViewHitTestType.Cell Then
                        .CurrentCell =
                        .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                    End If

                    Dim po As String
                    OD.TextBox1.Text = DataGridView1.Rows(hti.RowIndex).Cells(2).Value
                    OD.TextBox58.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                    If Trim(OD.TextBox81.Text).Length > 0 Then

                        po = Trim(DataGridView1.Rows(hti.RowIndex).Cells(8).Value) & Microsoft.VisualBasic.Mid(OD.TextBox81.Text, 3, 3)
                        OD.TextBox81.Text = ""
                        OD.TextBox81.Text = po
                    Else
                        OD.TextBox81.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(8).Value)
                    End If

                End With
                Me.Close()
            End If
        End If
    End Sub
End Class