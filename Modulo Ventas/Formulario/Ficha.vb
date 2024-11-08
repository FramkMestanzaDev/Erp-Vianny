Imports System.Data.SqlClient
Public Class Ficha
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public Property Padre1 As Codificacion_Prodcutos
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim dt As DataTable
    Private Sub Ficha_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
        Dim f As New fproductos
        Dim fq As New vproducot
        fq.gcia = Label2.Text
        dt = f.buscar_colores(fq)

        If dt.Rows.Count <> 0 Then
            DataGridView1.DataSource = dt
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).Width = 250

        End If

        'abrir()
        'Dim cmd As New SqlDataAdapter("select ccolor_3c as Codigo,cnomb_3c as Color from custom_vianny.dbo.qarc0300 where ccia_3c ='" + Label2.Text + "'", conx)

        'Dim da As New DataTable

        'cmd.Fill(da)
        'DataGridView1.DataSource = da
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "Color" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 80
                DataGridView1.Columns(1).Width = 250
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub





    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'If TextBox2.Text = "2" Then
        '    Op_Manufactura.TextBox17.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        '    Op_Manufactura.colorprenda.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value

        'Else
        '    If TextBox2.Text = "3" Then
        '        Requerimiento_op.DataGridView1.Rows(0).Cells(10).Value = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        '        Requerimiento_op.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        '    Else
        '        If TextBox2.Text = "4" Then
        '            Codificacion_Prodcutos.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        '            Codificacion_Prodcutos.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        '            Dim fc As New fcodigo
        '            Dim kk As New vlogica

        '            kk.gCCia = Label2.Text
        '            kk.gFamilia = Codificacion_Prodcutos.ComboBox1.SelectedValue.ToString
        '            kk.gSubFamilia = Codificacion_Prodcutos.ComboBox2.SelectedValue.ToString
        '            kk.gOrigen = "N"
        '            kk.gColor = DataGridView1.Rows(e.RowIndex).Cells(1).Value

        '            Codificacion_Prodcutos.TextBox3.Text = fc.mostrar_correlativo(kk)
        '            Codificacion_Prodcutos.TextBox7.Text = Codificacion_Prodcutos.ComboBox1.SelectedValue.ToString + Codificacion_Prodcutos.ComboBox2.SelectedValue.ToString + "N" + Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value) + Trim(fc.mostrar_correlativo(kk))
        '        Else
        '            If TextBox2.Text = "5" Then
        '                OD.TextBox18.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        '                OD.TextBox6.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        '            Else
        '                Orden_Produccion.TextBox5.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        '                Orden_Produccion.TextBox8.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        '            End If

        '        End If


        '    End If




        'End If
        'Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_MouseClick(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            With DataGridView1

                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If
                If TextBox2.Text.ToString().Trim() = "2" Then
                    Op_Manufactura.TextBox18.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                    Op_Manufactura.colortela.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value

                Else
                    If TextBox2.Text.ToString().Trim() = "3" Then
                        Requerimiento_op.DataGridView1.Rows(0).Cells(10).Value = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                        Requerimiento_op.DataGridView1.Rows(0).Cells(3).Value = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                    Else
                        If TextBox2.Text.ToString().Trim() = "4" Then
                            Codificacion_Prodcutos.TextBox1.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                            Codificacion_Prodcutos.TextBox2.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                            Dim fc As New fcodigo
                            Dim kk As New vlogica

                            kk.gCCia = Label2.Text
                            kk.gFamilia = Codificacion_Prodcutos.TextBox13.Text.ToString().Trim()
                            kk.gSubFamilia = Codificacion_Prodcutos.TextBox14.Text.ToString().Trim()
                            kk.gOrigen = "N"
                            kk.gColor = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                            Codificacion_Prodcutos.TextBox3.Text = fc.mostrar_correlativo(kk)

                            Codificacion_Prodcutos.TextBox7.Text = Codificacion_Prodcutos.TextBox13.Text.ToString().Trim() + Codificacion_Prodcutos.TextBox14.Text.ToString().Trim() + "N" + Trim(DataGridView1.Rows(hti.RowIndex).Cells(0).Value) + Trim(fc.mostrar_correlativo(kk))
                            Codificacion_Prodcutos.TextBox3.Select()
                        Else
                            If TextBox2.Text.ToString().Trim() = "5" Then
                                OD.TextBox18.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                                OD.TextBox6.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                            Else
                                Orden_Produccion.TextBox5.Text = DataGridView1.Rows(hti.RowIndex).Cells(0).Value
                                Orden_Produccion.TextBox8.Text = DataGridView1.Rows(hti.RowIndex).Cells(1).Value
                            End If

                        End If


                    End If




                End If




            End With
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If TextBox2.Text.ToString().Trim() = "4" Then
            Padre1.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            Padre1.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            Dim fc As New fcodigo
            Dim kk As New vlogica
            kk.gCCia = Label2.Text
            kk.gFamilia = Padre1.TextBox13.Text.ToString().Trim()
            kk.gSubFamilia = Padre1.TextBox14.Text.ToString().Trim()
            kk.gOrigen = "N"
            kk.gColor = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            Padre1.TextBox3.Text = fc.mostrar_correlativo(kk)
            Padre1.TextBox7.Text = Padre1.TextBox13.Text.ToString().Trim() + Padre1.TextBox14.Text.ToString().Trim() + "N" + Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value) + Trim(fc.mostrar_correlativo(kk))

            Padre1.TextBox3.Select()
            Padre1.TextBox3.Focus()

        End If
        Me.Close()
    End Sub
End Class