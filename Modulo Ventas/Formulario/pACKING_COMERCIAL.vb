Imports System.Data.SqlClient
Public Class pACKING_COMERCIAL
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Dim Rsr2, Rsr20, Rsr21, Rsr22 As SqlDataReader
    Private Sub buscar2()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim k As Integer
            Dim jk As Double

            dv.RowFilter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(1).Width = 70
                DataGridView1.Columns(2).Width = 550
                DataGridView1.Columns(3).Width = 70
                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 70
                DataGridView1.Columns(6).Width = 80
                DataGridView1.Columns(7).Visible = False
                Dim U As Integer
                U = DataGridView1.Rows.Count
                For i = 0 To U - 1
                    If DataGridView1.Rows(i).Cells(6).Value = "INGRESADO" Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Azure
                    Else
                        If DataGridView1.Rows(i).Cells(6).Value = "DESPACHADO" Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Beige
                        End If

                    End If
                Next

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar2()
    End Sub

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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        DataGridView1.DataSource = ""
        da.Clear()
        Dim calidad As String
        If RadioButton1.Checked = True Then
            calidad = "PRIMERA"
        Else
            If RadioButton2.Checked = True Then
                calidad = "SEGUNDA"
            Else
                calidad = "TERCERA"
            End If

        End If
        Dim estad As String
        estad = ""

        Dim cmd As New SqlDataAdapter("SELECT partida AS PARTIDA,des_articulo AS PRODUCTO,numero_rollos AS CALIDAD,ROLLOS AS ROLLOS,  total- (ROLLOS*0.3) AS KGS, CASE WHEN seleccionado =0 THEN 'PENDIENTE' WHEN seleccionado =1 THEN 'INGRESADO' ELSE 'DESPACHADO' END AS ESTADO,orden  FROM cab_ingresop WHERE numero_rollos ='" + Trim(calidad) + "'   ORDER BY PARTIDA", conx)
        cmd.Fill(da)
        If da.Rows.Count <> 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(2).Width = 550
            DataGridView1.Columns(3).Width = 70
            DataGridView1.Columns(4).Width = 70
            DataGridView1.Columns(5).Width = 70
            DataGridView1.Columns(6).Width = 80
            DataGridView1.Columns(7).Visible = False
            Dim U As Integer
            U = DataGridView1.Rows.Count
            For i = 0 To U - 1
                If DataGridView1.Rows(i).Cells(6).Value = "INGRESADO" Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Azure
                Else
                    If DataGridView1.Rows(i).Cells(6).Value = "DESPACHADO" Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Beige
                    End If

                End If
            Next
        Else
            DataGridView1.DataSource = ""
        End If

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 0 Then
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            If Trim(DataGridView1.Rows(num1).Cells(7).Value) = "" Then

                Packing_nacional.TextBox1.Text = DataGridView1.Rows(num1).Cells(1).Value
                Packing_nacional.TextBox2.Text = DataGridView1.Rows(num1).Cells(3).Value
                Packing_nacional.Show()
            Else
                Packing_exportacion.TextBox1.Text = DataGridView1.Rows(num1).Cells(1).Value
                Packing_exportacion.TextBox2.Text = DataGridView1.Rows(num1).Cells(3).Value
                Packing_exportacion.Show()
            End If

        End If
    End Sub
End Class