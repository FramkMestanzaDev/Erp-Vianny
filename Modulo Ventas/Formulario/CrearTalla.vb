Imports System.Data.SqlClient
Public Class CrearTalla
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim Rsr20 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim DT14 As New DataTable
    Private Sub CrearTalla_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MOSTRAR_INFORMACION()
        mostrarCorrelativo()
    End Sub
    Sub mostrarCorrelativo()
        Dim sql10212 As String = " SELECT max(cele) FROM custom_vianny.DBO.TAB0100 Where ccia='01' AND CTAB='TALLAS'"
        Dim cmd10212 As New SqlCommand(sql10212, conx)
        Rsr20 = cmd10212.ExecuteReader()
        If Rsr20.Read() = True Then
            TextBox1.Text = Convert.ToInt64(Rsr20(0)) + 1
        End If

    End Sub
    Function ELIMINAR(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Sub MOSTRAR_INFORMACION()
        abrir()
        DataGridView1.DataSource = ""
        DT14.Clear()
        Dim cmd6 As New SqlDataAdapter(" SELECT cele,dele FROM custom_vianny.DBO.TAB0100   Where ccia='01' AND CTAB='TALLAS'", conx)
        cmd6.Fill(DT14)
        DataGridView1.DataSource = DT14
        DataGridView1.Columns(0).Width = 110
        DataGridView1.Columns(1).Width = 340
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim agregar As String = "DELETE FROM custom_vianny.DBO.TAB0100 Where ccia='01' AND CTAB='TALLAS' AND cele ='" + Trim(TextBox1.Text) + "'"
        ELIMINAR(agregar)
        Dim cmd20 As New SqlCommand("insert into custom_vianny.DBO.TAB0100 (ccia,ctab,cele,dele,nele,codigo,codigo2) values ('01','TALLAS',@cele,@dele,0.0,'','')", conx)
        cmd20.Parameters.AddWithValue("@cele", Trim(TextBox1.Text))
        cmd20.Parameters.AddWithValue("@dele", Trim(TextBox2.Text))
        cmd20.ExecuteNonQuery()
        MsgBox(" SE REGISTRO CORRECTAMENTE LA INFORMACION")
        limpiar()
        MOSTRAR_INFORMACION()
        mostrarCorrelativo()
    End Sub
    Sub limpiar()
        TextBox1.Text = ""
        TextBox2.Text = ""

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer
            ds.Tables.Add(DT14.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer
            dv.RowFilter = "dele" & " like '%" & TextBox2.Text & "%'"
            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(5).Width = 400
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 85
                DataGridView1.Columns(4).Width = 85
                DataGridView1.Columns(6).Width = 65
                DataGridView1.Columns(7).Width = 65
                DataGridView1.Columns(8).Width = 85
                DataGridView1.Columns(9).Width = 130
                DataGridView1.Columns(10).Visible = False
                DataGridView1.Columns(11).Visible = False
                For i = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Rows(i).Cells(11).Value.ToString.Trim = "1" Then
                        DataGridView1.Rows(i).Cells(1).Value = True
                    End If
                Next
            Else
                DataGridView1.DataSource = Nothing
            End If
        Catch ex As Exception


        End Try
    End Sub
End Class