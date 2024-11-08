Imports System.Data.SqlClient
Public Class Editar_Producto
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim DY As New DataTable
    Public conn As SqlDataAdapter
    Dim ty, ty2, ty3 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Editar_Producto_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box()
    End Sub
    Sub llenar_combo_box()

        Try

            conn = New SqlDataAdapter(" SELECT unid_16m as unidad, nomb_16m as nombre FROM custom_vianny.dbo.MAG1600", conx)
            conn.Fill(ty)
            ComboBox6.DataSource = ty
            ComboBox6.DisplayMember = "unidad"
            ComboBox6.ValueMember = "nombre"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        abrir()
        Dim cmd15 As New SqlCommand("UPDATE custom_vianny.dbo.cag1700 SET nomb_17 =@NOMBRE, nombs_17=@NOMBRE, unid_17 =@UNIDAD WHERE linea_17 ='" + Trim(TextBox1.Text) + "'  AND ccia ='01' ", conx)
        cmd15.Parameters.AddWithValue("@NOMBRE", Trim(TextBox2.Text))
        cmd15.Parameters.AddWithValue("@UNIDAD", Trim(ComboBox6.Text))

        cmd15.ExecuteNonQuery()
        MsgBox("SE ACTUALIZO LO SOLICITADO")
        Codificacion_Prodcutos.TextBox4.Text = ""
        Codificacion_Prodcutos.TextBox9.Text = ""
        Codificacion_Prodcutos.ACTUALIZAR()
        Dim hj2 As New flog
        Dim hj1 As New vlog
        hj1.gmodulo = "CODIGO"
        hj1.gcuser = MDIParent2.Label6.Text
        hj1.gaccion = 2
        hj1.gpc = My.Computer.Name
        hj1.gfecha = DateTimePicker1.Value
        hj1.gdetalle = TextBox1.Text.ToString().Trim()
        hj1.gccia = Label4.Text
        hj2.insertar_log(hj1)
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Codigo_barra.Close()
        Codigo_barra.TextBox1.Text = TextBox1.Text
        Codigo_barra.ShowDialog()
    End Sub
End Class