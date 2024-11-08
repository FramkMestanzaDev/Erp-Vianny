Imports System.Data.SqlClient
Public Class SegundoÑPasado_rama
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
    Dim dy As New DataTable
    Private Sub buscar()
        Try

            Dim ds As New DataSet

            ds.Tables.Add(dy.Copy)

            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then
                Dim jk As Integer
                DataGridView1.DataSource = dv
                DataGridView1.Columns(2).Width = 400
                DataGridView1.Columns(4).Visible = False
                Dim u As Integer

                u = DataGridView1.Rows.Count
                For i = 0 To u - 1
                    If DataGridView1.Rows(i).Cells(4).Value = 2 Then
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                    Else
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                    End If
                Next
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim u As Integer
        u = DataGridView1.Rows.Count
        For i = 0 To u - 1
            If DataGridView1.Rows(i).Cells(0).Value = True Then
                abrir()
                Dim cmd15 As New SqlCommand("update validar_partida set estado =1 where partida = @partida ", conx)
                cmd15.Parameters.AddWithValue("@partida", Trim(DataGridView1.Rows(i).Cells(1).Value))
                cmd15.ExecuteNonQuery()
            End If
        Next
        motrar()
    End Sub
    Sub motrar()
        abrir()
        Dim cmd As New SqlDataAdapter("select partida AS PARTIDA,articulo AS ARTICULO,kgsobtenidos AS KGS,estado AS ESTADO from  validar_partida", conx)
        cmd.Fill(dy)

        If dy.Rows.Count > 0 Then
            'DataGridView1.Rows.Clear()
            DataGridView1.DataSource = dy
            DataGridView1.Columns(2).Width = 400
            DataGridView1.Columns(4).Visible = False
            Dim u As Integer

            u = DataGridView1.Rows.Count
            For i = 0 To u - 1
                If DataGridView1.Rows(i).Cells(4).Value = 2 Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Wheat
                Else
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                End If
            Next


        Else
            DataGridView1.Rows.Clear()
        End If
    End Sub
    Private Sub SegundoÑPasado_rama_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        motrar()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class