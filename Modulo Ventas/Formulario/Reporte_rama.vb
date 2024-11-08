Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Data.SqlClient
Public Class Reporte_rama
    Dim dt As New DataTable
    Dim dt1 As New DataTable
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
    'Private Sub buscar()
    '    Try
    '        Dim ds As New DataSet
    '        ds.Tables.Add(dt.Copy)
    '        Dim dv As New DataView(ds.Tables(0))


    '        dv.RowFilter = "PROGRAMA" & " like '%" & TextBox1.Text & "%'"

    '        If dv.Count <> 0 Then

    '            DataGridView1.DataSource = dv
    '            DataGridView1.Columns(5).Visible = False
    '            DataGridView1.Columns(0).ReadOnly = True
    '            DataGridView1.Columns(1).ReadOnly = True
    '            DataGridView1.Columns(2).ReadOnly = True
    '            DataGridView1.Columns(3).ReadOnly = True
    '            DataGridView1.Columns(4).ReadOnly = True
    '            DataGridView1.Columns(1).Width = 70
    '        Else

    '            DataGridView1.DataSource = Nothing
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message)

    '    End Try
    'End Sub
    Private Sub Reporte_rama_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim hj As New frama
        'dt = hj.reporte_ram()
        'DataGridView1.DataSource = dt
        'DataGridView1.Columns(5).Visible = False
        'DataGridView1.Columns(0).ReadOnly = True
        'DataGridView1.Columns(1).ReadOnly = True
        'DataGridView1.Columns(2).ReadOnly = True
        'DataGridView1.Columns(3).ReadOnly = True
        'DataGridView1.Columns(4).ReadOnly = True
        'DataGridView1.Columns(1).Width = 70
        abrir()
        Dim sql3 As String = "select top 1 codigo  from rama order by codigo desc"
        Dim cmd3 As New SqlCommand(sql3, conx)
        RS3 = cmd3.ExecuteReader()
        RS3.Read()
        TextBox1.Text = RS3(0)

        RS3.Close()
    End Sub




    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'buscar()
    End Sub
    Dim RS4, RS3, RS5 As SqlDataReader

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim ext As New Exportar
        ext.llenarExcel(DataGridView2)
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim kk As New frama
            Dim ff As New vrama
            Dim SUM As Double
            Dim SUM1, SUM2, SUM3 As Integer
            TextBox6.Text = ""
            Select Case TextBox1.Text.Length

                Case "1" : TextBox1.Text = "000000000" & "" & TextBox1.Text
                Case "2" : TextBox1.Text = "00000000" & "" & TextBox1.Text
                Case "3" : TextBox1.Text = "0000000" & "" & TextBox1.Text
                Case "4" : TextBox1.Text = "000000" & "" & TextBox1.Text
                Case "5" : TextBox1.Text = "00000" & "" & TextBox1.Text
                Case "6" : TextBox1.Text = "0000" & "" & TextBox1.Text
                Case "7" : TextBox1.Text = "000" & "" & TextBox1.Text
                Case "8" : TextBox1.Text = "00" & "" & TextBox1.Text
                Case "9" : TextBox1.Text = "0" & "" & TextBox1.Text
                Case "10" : TextBox1.Text = TextBox1.Text
            End Select
            ff.gcodigo = Trim(TextBox1.Text)
            dt1 = kk.seguimiento_rama(ff)
            DataGridView2.DataSource = dt1
            DataGridView2.Columns(0).Width = 80
            DataGridView2.Columns(1).Width = 65
            DataGridView2.Columns(2).Width = 65
            DataGridView2.Columns(3).Width = 90
            DataGridView2.Columns(4).Width = 90
            DataGridView2.Columns(5).Width = 350
            DataGridView2.Columns(6).Width = 80
            DataGridView2.Columns(7).Width = 80
            DataGridView2.Columns(8).Width = 70
            DataGridView2.Columns(9).Width = 70
            DataGridView2.Columns(10).Width = 130
            DataGridView2.Columns(0).Frozen = True
            DataGridView2.Columns(1).Frozen = True
            DataGridView2.Columns(2).Frozen = True
            DataGridView2.Columns(3).Frozen = True
            Dim i7 As Integer
            i7 = DataGridView2.Rows.Count
            For I1 = 0 To i7 - 1
                If Trim(DataGridView2.Rows(I1).Cells(4).Value) = "TERMINADO" Then
                    DataGridView2.Rows(I1).DefaultCellStyle.BackColor = Color.DarkRed
                    DataGridView2.Rows(I1).DefaultCellStyle.ForeColor = Color.White
                    SUM = SUM + DataGridView2.Rows(I1).Cells(1).Value
                    SUM1 = SUM1 + 1
                    'Label19.Text = Format(TimeValue(DataGridView2.Rows(e.RowIndex).Cells(3).Value) - TimeValue(DataGridView2.Rows(e.RowIndex).Cells(2).Value), "hh:mm")

                Else
                    If Trim(DataGridView2.Rows(I1).Cells(4).Value) = "PASANDO" Then
                        DataGridView2.Rows(I1).DefaultCellStyle.BackColor = Color.DarkKhaki
                        DataGridView2.Rows(I1).DefaultCellStyle.ForeColor = Color.White
                        SUM2 = SUM2 + DataGridView2.Rows(I1).Cells(1).Value
                        SUM3 = SUM3 + 1
                    End If
                End If

            Next
            For I1 = 0 To i7 - 1
                If Trim(DataGridView2.Rows(I1).Cells(4).Value) = "TERMINADO" Then
                    'Dim tim1, tim2, TIM3, TIM4, HORA, MINUTO As Integer
                    'Dim RESPUETA, RESPUESTA2, RESPUESTA3 As Double



                End If
                Exit For
            Next
            abrir()
            Dim sql2 As String = "SELECT COUNT(codigo),SUM(CAST (rollos AS NUMERIC(7,0))),SUM(peso) FROM rama where codigo ='" + Trim(TextBox1.Text) + "'"
            Dim cmd2 As New SqlCommand(sql2, conx)
            RS4 = cmd2.ExecuteReader()
            RS4.Read()
            TextBox2.Text = RS4(0)
            TextBox3.Text = RS4(1)
            TextBox4.Text = RS4(2)
            RS4.Close()
            TextBox5.Text = SUM1

            TextBox7.Text = SUM
            TextBox8.Text = SUM3
            TextBox9.Text = SUM2
            ProgressBar1.Value = 0
            Dim gh As Double
            gh = (TextBox5.Text * 100) / TextBox2.Text
            ProgressBar1.Increment(gh)
            Label12.Text = (TextBox5.Text * 100) / TextBox2.Text
            abrir()
            Dim sql4 As String = "select MOTIVO + COMENTARIO from comentario_rama where PROGRAMA ='" + Trim(TextBox1.Text) + "'
"
            Dim cmd4 As New SqlCommand(sql4, conx)
            RS5 = cmd4.ExecuteReader()

            Do While RS5.Read()
                TextBox6.Text = Trim(TextBox6.Text) + " , " + Trim(RS5(0))
            Loop

        End If

    End Sub
End Class