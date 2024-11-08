Imports System.Data.SqlClient
Public Class Almacenes_Precios
    Public conx As SqlConnection
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.DataSource = Nothing
        importarExcel(DataGridView1)

        Dim i1 As Integer
        Dim total, cantidad As Double
        cantidad = 0
        total = 0
        i1 = DataGridView1.Rows.Count

        For i = 0 To i1 - 1
            'If Trim(DataGridView1.Rows(i).Cells(5).Value) <> DBNull.Value Then
            cantidad = cantidad + DataGridView1.Rows(i).Cells(6).Value
                total = total + DataGridView1.Rows(i).Cells(8).Value
            'End If

        Next
        TextBox1.Text = Convert.ToString(cantidad)
        TextBox2.Text = Convert.ToString(total)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim o As Integer
        o = DataGridView1.Rows.Count
        abrir()
        If CheckBox1.Checked = False Then
            For i = 0 To o - 1

                Dim cmd17 As New SqlCommand("update  custom_vianny.dbo.map030f set preun1_3a=@preun1_3a , tot1_3a =@tot1_3a where CCIA_3A =@CCIA_3A and CPERIODO_3A =@CPERIODO_3A and talm_3a =@talm_3a and ccom_3a =@ccom_3a and ncom_3a =@ncom_3a and linea_3a =@linea_3a and cantk_3a =@cantk_3a ", conx)
                cmd17.Parameters.AddWithValue("@preun1_3a", DataGridView1.Rows(i).Cells(7).Value)
                cmd17.Parameters.AddWithValue("@tot1_3a", DataGridView1.Rows(i).Cells(8).Value)
                cmd17.Parameters.AddWithValue("@CCIA_3A", Trim(DataGridView1.Rows(i).Cells(0).Value))
                cmd17.Parameters.AddWithValue("@CPERIODO_3A", Trim(DataGridView1.Rows(i).Cells(1).Value))
                cmd17.Parameters.AddWithValue("@talm_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd17.Parameters.AddWithValue("@ccom_3a", Trim(DataGridView1.Rows(i).Cells(3).Value))
                cmd17.Parameters.AddWithValue("@ncom_3a", Trim(DataGridView1.Rows(i).Cells(4).Value))
                cmd17.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(5).Value))
                cmd17.Parameters.AddWithValue("@cantk_3a", DataGridView1.Rows(i).Cells(6).Value)
                cmd17.ExecuteNonQuery()
            Next
        Else
            For i = 0 To o - 1

                Dim cmd17 As New SqlCommand("update  custom_vianny.dbo.map030f set preun1_3a=@preun1_3a , tot1_3a =@tot1_3a where CCIA_3A =@CCIA_3A and CPERIODO_3A =@CPERIODO_3A and talm_3a =@talm_3a and ccom_3a =@ccom_3a and ncom_3a =@ncom_3a and linea_3a =@linea_3a and cantk_3a =@cantk_3a and lote_3a =@lote_3a ", conx)
                cmd17.Parameters.AddWithValue("@preun1_3a", DataGridView1.Rows(i).Cells(7).Value)
                cmd17.Parameters.AddWithValue("@tot1_3a", DataGridView1.Rows(i).Cells(8).Value)
                cmd17.Parameters.AddWithValue("@CCIA_3A", Trim(DataGridView1.Rows(i).Cells(0).Value))
                cmd17.Parameters.AddWithValue("@CPERIODO_3A", Trim(DataGridView1.Rows(i).Cells(1).Value))
                cmd17.Parameters.AddWithValue("@talm_3a", Trim(DataGridView1.Rows(i).Cells(2).Value))
                cmd17.Parameters.AddWithValue("@ccom_3a", Trim(DataGridView1.Rows(i).Cells(3).Value))
                cmd17.Parameters.AddWithValue("@ncom_3a", Trim(DataGridView1.Rows(i).Cells(4).Value))
                cmd17.Parameters.AddWithValue("@linea_3a", Trim(DataGridView1.Rows(i).Cells(5).Value))
                cmd17.Parameters.AddWithValue("@cantk_3a", DataGridView1.Rows(i).Cells(6).Value)
                cmd17.Parameters.AddWithValue("@lote_3a", DataGridView1.Rows(i).Cells(9).Value)
                cmd17.ExecuteNonQuery()
            Next
        End If


        MsgBox("LA INFORMACION DE REGISTRO CORRECTAMENTE")
        DataGridView1.DataSource = Nothing
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub Almacenes_Precios_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class