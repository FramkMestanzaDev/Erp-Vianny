Imports System.Data.SqlClient
Public Class Detalle_ruta
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

    Dim da As New DataTable
    Sub mostrar_informacion()
        abrir()
        da.Clear()
        DataGridView1.DataSource = Nothing
        If Label2.Text = 1 Then
            Dim cmd As New SqlDataAdapter("SELECT fich_10 as RUC,nomb_10 AS EMPRESA FROM custom_vianny.dbo.cag1000 WHERE Url_10='1'", conx)
            cmd.Fill(da)
            DataGridView1.DataSource = da
            DataGridView1.Columns(1).Width = 400

        Else
            If Label2.Text = 2 Then
                Dim cmd As New SqlDataAdapter("SELECT fich_10 as RUC,nomb_10 AS EMPRESA FROM custom_vianny.dbo.cag1000 WHERE Url_10='2'  and ccia ='01'", conx)
                cmd.Fill(da)
                DataGridView1.DataSource = da

                DataGridView1.Columns(1).Width = 400
            Else
                If Label2.Text = 3 Then

                    Dim col As New DataGridViewTextBoxColumn
                    col.Name = "MODELISTAS"
                    DataGridView1.Columns.Add(col)
                    DataGridView1.Rows.Add(3)
                    DataGridView1.Columns(0).Width = 400
                    DataGridView1.Rows(0).Cells(0).Value = "DAGOBERTO"
                    DataGridView1.Rows(1).Cells(0).Value = "BEDER BLASS"
                    DataGridView1.Rows(2).Cells(0).Value = "SANDRA ACOSTA"
                End If

            End If
        End If

    End Sub

    Private Sub Detalle_ruta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrar_informacion()
    End Sub


    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        'If Label2.Text = 1 Then
        '    Od_Udp.TextBox84.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)

        'Else
        '    If Label2.Text = 2 Then
        '        Od_Udp.TextBox85.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        '    Else
        '        Od_Udp.TextBox80.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
        '        DataGridView1.Rows.Clear()
        '        DataGridView1.Columns.Clear()
        '    End If



        'End If

        'Close()
    End Sub



    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            With DataGridView1

                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                ' Obtenemos la parte del control a las que
                ' pertenecen las coordenadas.
                '
                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If
                If Label2.Text = 1 Then
                    Od_Udp.TextBox84.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)

                Else
                    If Label2.Text = 2 Then
                        Od_Udp.TextBox85.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)
                    Else
                        Od_Udp.ComboBox6.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(0).Value)
                        DataGridView1.Rows.Clear()
                        DataGridView1.Columns.Clear()
                    End If



                End If


            End With
            Close()
        End If
    End Sub
End Class