Imports System.Data.SqlClient
Public Class MCoFamilia
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
    Dim DT14 As New DataTable
    Dim Rsr20 As SqlDataReader
    Private Sub MCoFamilia_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.SelectedIndex = 0
        TextBox4.Text = "PPL"
        MOSTRAR_INFORMACION(Trim(TextBox4.Text))
    End Sub
    Public Sub MOSTRAR_INFORMACION(familia)
        abrir()
        DataGridView1.DataSource = ""
        DT14.Clear()
        Dim cmd6 As New SqlDataAdapter("select  codigo_sfam as CODIGO,descrip_sfam AS DESCRIPCION,siglas_sfam AS ABREVIATURA  from custom_vianny.dbo.subfamilias where ccia_sfam ='01' and talm_sfam ='01' and familia_sfam='" + familia + "'", conx)
        cmd6.Fill(DT14)
        DataGridView1.DataSource = DT14
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns(1).Width = 250
        DataGridView1.Columns(2).Width = 250
    End Sub

    Sub BUSCAR()
        Dim ds As New DataSet
        ds.Tables.Add(DT14.Copy)
        Dim dv As New DataView(ds.Tables(0))
        dv.RowFilter = "DESCRIPCION" & " like '%" & TextBox5.Text & "%'"
        If dv.Count <> 0 Then
            DataGridView1.DataSource = dv
            DataGridView1.Columns(0).Width = 60
            DataGridView1.Columns(1).Width = 250
            DataGridView1.Columns(2).Width = 250
        Else
            DataGridView1.DataSource = Nothing
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        BUSCAR()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "PRENDA TELA PLANA" Then
            TextBox4.Text = "PPL"
            MOSTRAR_INFORMACION(Trim(TextBox4.Text))
        Else
            TextBox4.Text = "PTP"
            MOSTRAR_INFORMACION(Trim(TextBox4.Text))
        End If
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
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
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Dim sql10212 As String = "select   COUNT( codigo_sfam)   from custom_vianny.dbo.subfamilias where ccia_sfam ='01' and talm_sfam ='01' and familia_sfam='" + Trim(TextBox4.Text) + "' and codigo_sfam ='" + Trim(TextBox1.Text) + "'"
        Dim cmd10212 As New SqlCommand(sql10212, conx)
        Rsr20 = cmd10212.ExecuteReader()
        Rsr20.Read()
        If Rsr20(0) = 0 Then
            Rsr20.Close()
            Dim agregar As String = "delete from custom_vianny.dbo.subfamilias where ccia_sfam ='01' and talm_sfam ='01' and familia_sfam='" + Trim(TextBox4.Text) + "' and codigo_sfam ='" + Trim(TextBox1.Text) + "' "
            ELIMINAR(agregar)
            Dim cmd20 As New SqlCommand("insert into custom_vianny.dbo.subfamilias (ccia_sfam,talm_sfam,familia_sfam,codigo_sfam,descrip_sfam,siglas_sfam) values ('01','01',@familia_sfam,@codigo_sfam,@descrip_sfam,@siglas_sfam)", conx)
            cmd20.Parameters.AddWithValue("@familia_sfam", Trim(TextBox4.Text))
            cmd20.Parameters.AddWithValue("@codigo_sfam", Trim(TextBox1.Text))
            cmd20.Parameters.AddWithValue("@descrip_sfam", Trim(TextBox2.Text))
            cmd20.Parameters.AddWithValue("@siglas_sfam", Trim(TextBox3.Text))
            cmd20.ExecuteNonQuery()
            MsgBox(" SE REGISTRO CORRECTAMENTE LA INFORMACION")
            limpiar()
            TextBox4.Text = "PPL"
            MOSTRAR_INFORMACION(Trim(TextBox4.Text))

        Else
            Rsr20.Close()
            MsgBox("EL CODIGO YA EXISTE, INGRESADO UNO DIFERENTE")
        End If

    End Sub
    Sub limpiar()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        limpiar()
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
        TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
    End Sub
End Class