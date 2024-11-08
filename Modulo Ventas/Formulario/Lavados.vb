Imports System.Data.SqlClient
Public Class Lavados
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
    Dim da, da1, da2 As New DataTable

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub
    Sub actualizarinformacion()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT A.cele as CODIGO,a.dele AS DESCRIPCION  FROM custom_vianny.dbo.TAB0100 A  Where A.CCia = '01' AND A.CTab = 'UDPLAV'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 190
        DataGridView1.Columns(0).ReadOnly = True
        DataGridView1.Columns(1).ReadOnly = True
        Dim cmd1 As New SqlDataAdapter("SELECT A.cele as CODIGO,a.dele AS DESCRIPCION FROM custom_vianny.dbo.TAB0100 A  Where A.CCia = '01' AND A.CTab = 'UDPCOL'", conx)
        cmd1.Fill(da1)
        DataGridView2.DataSource = da1
        DataGridView2.Columns(0).Width = 60
        DataGridView2.Columns(1).Width = 190
        DataGridView2.Columns(0).ReadOnly = True
        DataGridView2.Columns(1).ReadOnly = True
        Dim cmd2 As New SqlDataAdapter("SELECT A.cele as CODIGO,a.dele AS DESCRIPCION FROM custom_vianny.dbo.TAB0100 A  Where A.CCia = '01' AND A.CTab = 'UDPEFE'", conx)
        cmd2.Fill(da2)
        DataGridView3.DataSource = da2
        DataGridView3.Columns(0).Width = 60
        DataGridView3.Columns(1).Width = 190
        DataGridView3.Columns(0).ReadOnly = True
        DataGridView3.Columns(1).ReadOnly = True
    End Sub
    Private Sub Lavados_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        actualizarinformacion()

    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

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
                If Trim(TextBox4.Text).Length = 0 Then
                    TextBox4.Text = Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)
                Else
                    TextBox4.Text = Trim(TextBox4.Text) & " + " & Trim(DataGridView1.Rows(hti.RowIndex).Cells(1).Value)
                End If

            End With
        End If
    End Sub

    Private Sub DataGridView2_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView2.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            With DataGridView2

                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                ' Obtenemos la parte del control a las que
                ' pertenecen las coordenadas.
                '
                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If
                If Trim(TextBox5.Text).Length = 0 Then
                    TextBox5.Text = Trim(DataGridView2.Rows(hti.RowIndex).Cells(1).Value)
                Else
                    TextBox5.Text = Trim(TextBox5.Text) & " + " & Trim(DataGridView2.Rows(hti.RowIndex).Cells(1).Value)
                End If
            End With
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox4.Text = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox5.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox6.Text = ""
    End Sub

    Private Sub DataGridView3_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView3.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            With DataGridView3

                Dim hti As DataGridView.HitTestInfo = .HitTest(e.X, e.Y)

                ' Obtenemos la parte del control a las que
                ' pertenecen las coordenadas.
                '
                If hti.Type = DataGridViewHitTestType.Cell Then
                    .CurrentCell =
                    .Rows(hti.RowIndex).Cells(hti.ColumnIndex)
                End If
                If Trim(TextBox6.Text).Length = 0 Then
                    TextBox6.Text = Trim(DataGridView3.Rows(hti.RowIndex).Cells(1).Value)
                Else
                    TextBox6.Text = Trim(TextBox6.Text) & " + " & Trim(DataGridView3.Rows(hti.RowIndex).Cells(1).Value)
                End If
            End With
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim Resul As String
        Dim re, re1, re2 As Integer
        If Trim(TextBox4.Text) <> "" Then
            re = InStr(Trim(TextBox4.Text), Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value))
            re1 = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value).Length
            re2 = Trim(TextBox4.Text).Length
            Resul = Microsoft.VisualBasic.Mid(Trim(TextBox4.Text), 1, re - 4) & Microsoft.VisualBasic.Mid(Trim(TextBox4.Text), re + re1, re2)
            TextBox4.Text = ""
            TextBox4.Text = Resul
        End If

    End Sub

    Private Sub Lavados_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Trim(Label7.Text) = "1" Then
            OD.TextBox76.Text = Trim(TextBox4.Text)
            OD.TextBox96.Text = Trim(TextBox5.Text)
            OD.TextBox97.Text = Trim(TextBox6.Text)

        Else
            Od_Udp.TextBox35.Text = Trim(TextBox4.Text)
            Od_Udp.TextBox36.Text = Trim(TextBox5.Text)
            Od_Udp.TextBox37.Text = Trim(TextBox6.Text)
        End If
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox7.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 60
                DataGridView1.Columns(1).Width = 190
                DataGridView1.Columns(0).ReadOnly = True
                DataGridView1.Columns(1).ReadOnly = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        buscar()
    End Sub
    Private Sub buscar1()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da1.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox8.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView2.DataSource = dv
                DataGridView2.Columns(0).Width = 60
                DataGridView2.Columns(1).Width = 190
                DataGridView2.Columns(0).ReadOnly = True
                DataGridView2.Columns(1).ReadOnly = True
            Else

                DataGridView2.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        buscar1()
    End Sub
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da2.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox9.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView3.DataSource = dv
                DataGridView3.Columns(0).Width = 60
                DataGridView3.Columns(1).Width = 190
                DataGridView3.Columns(0).ReadOnly = True
                DataGridView3.Columns(1).ReadOnly = True
            Else

                DataGridView3.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        buscar2()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim tabla As New TablaCLE
        tabla.Label3.Text = "1"
        tabla.Label4.Text = "BUSCAR LAVADO"
        tabla.Label5.Text = "01"
        tabla.Text = "LAVADO"
        tabla.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim ta As New TablaCLE
        ta.Label3.Text = "2"
        ta.Label5.Text = "01"
        ta.Label4.Text = "BUSCAR COLOR"
        ta.Text = "COLOR"
        ta.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim tab As New TablaCLE
        tab.Label3.Text = "3"
        tab.Label5.Text = "01"
        tab.Label4.Text = "BUSCAR EFECTO"
        tab.Text = "EFECTO"
        tab.ShowDialog()
    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        Dim Resul As String
        Dim re, re1, re2 As Integer
        If Trim(TextBox5.Text) <> "" Then
            re = InStr(Trim(TextBox5.Text), Trim(DataGridView2.Rows(e.RowIndex).Cells(1).Value))
            re1 = Trim(DataGridView2.Rows(e.RowIndex).Cells(1).Value).Length
            re2 = Trim(TextBox5.Text).Length
            Resul = Microsoft.VisualBasic.Mid(Trim(TextBox5.Text), 1, re - 4) & Microsoft.VisualBasic.Mid(Trim(TextBox5.Text), re + re1, re2)
            TextBox5.Text = ""
            TextBox5.Text = Resul
        End If

    End Sub

    Private Sub DataGridView3_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellDoubleClick
        Dim Resul As String
        Dim re, re1, re2 As Integer
        If Trim(TextBox6.Text) <> "" Then
            re = InStr(Trim(TextBox6.Text), Trim(DataGridView3.Rows(e.RowIndex).Cells(1).Value))
            re1 = Trim(DataGridView3.Rows(e.RowIndex).Cells(1).Value).Length
            re2 = Trim(TextBox6.Text).Length
            Resul = Microsoft.VisualBasic.Mid(Trim(TextBox6.Text), 1, re - 4) & Microsoft.VisualBasic.Mid(Trim(TextBox6.Text), re + re1, re2)
            TextBox6.Text = ""
            TextBox6.Text = Resul
        End If

    End Sub
End Class