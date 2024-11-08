Imports System.Data.SqlClient
Public Class Tabla_Vendedores
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim da As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub MOSTRAR()
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT a.codigo_ven  AS CODIGO, A.alias_ven AS ALIAS,A.nombres_ven  AS NOMBRES ,A.apepat_ven AS AP_PATERNO,A.apemat_ven AS AP_MATENO,ccia_ven,case when  rpm_ven = '1' then 'VEND ACTIVO' ELSE 'INACTIVO' END AS ESTADO FROM custom_vianny.dbo.Vendedores A where ccia_ven ='01' order by codigo_ven", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Width = 70
        DataGridView1.Columns(1).Width = 140
        DataGridView1.Columns(2).Width = 140
        DataGridView1.Columns(3).Width = 140
        DataGridView1.Columns(4).Width = 140
        DataGridView1.Columns(5).Visible = False
        DataGridView1.Columns(6).Width = 120
    End Sub
    Private Sub Tabla_Vendedores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MOSTRAR()
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(2).Value)
        TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(3).Value)
        TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value)
        TextBox4.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
        If Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value) = "01" Then
            ComboBox1.SelectedIndex = 0
        Else
            ComboBox1.SelectedIndex = 1
        End If

        If Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value) = "INACTIVO" Then
            RadioButton1.Checked = False
            RadioButton2.Checked = True
        Else
            RadioButton1.Checked = True
            RadioButton2.Checked = False
        End If
        Label9.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
        Button1.Enabled = False

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        Button3.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        Button1.Enabled = True
        Label11.Text = 2
    End Sub
    Dim Rsr2 As SqlDataReader
    Sub CORRELATIVO()
        abrir()
        Dim sql102 As String = "SELECT MAX(A.Codigo_ven) AS Ultimo 
				   FROM custom_vianny.dbo.Vendedores A"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        Rsr2.Read()
        TextBox5.Text = Rsr2(0) + 1
        Rsr2.Close()
        Select Case TextBox5.Text.ToString.Length

            Case "1" : TextBox5.Text = "000" & TextBox5.Text
            Case "2" : TextBox5.Text = "00" & TextBox5.Text
            Case "3" : TextBox5.Text = "0" & TextBox5.Text
            Case "4" : TextBox5.Text = TextBox5.Text
        End Select
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
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox6.Enabled = True
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        Button1.Enabled = True
        RadioButton1.Checked = True
        Label11.Text = 1
        CORRELATIVO()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        abrir()
        If Label11.Text = 1 Then
            Dim cmd15 As New SqlCommand("INSERT INTO custom_vianny.dbo.Vendedores (Codigo_ven,
										   CCia_ven,
										   ficha_ven,
										   alias_ven,
										   comision_ven,
										   apepat_ven,
										   apemat_ven,
										   nombres_ven,
										   email_ven,
										   fono_ven,
										   rpm_ven,
										   admin_ven,
										   clave_ven)
								VALUES   ( @Codigo_ven,
										   @CCia_ven,
										   '',
										   @alias_ven,
										   0.00,
										   @apepat_ven,
										   @apemat_ven,
										   @nombres_ven,
										   '',
										   '',
										   @ESTADO,
										   0,
										  @clave_ven)", conx)
            cmd15.Parameters.AddWithValue("@Codigo_ven", Trim(TextBox5.Text))
            cmd15.Parameters.AddWithValue("@CCia_ven", Trim(Label9.Text))
            cmd15.Parameters.AddWithValue("@alias_ven", Trim(TextBox4.Text))
            cmd15.Parameters.AddWithValue("@apepat_ven", Trim(TextBox2.Text))
            cmd15.Parameters.AddWithValue("@apemat_ven", Trim(TextBox3.Text))
            cmd15.Parameters.AddWithValue("@nombres_ven", Trim(TextBox1.Text))
            cmd15.Parameters.AddWithValue("@clave_ven", Trim(TextBox6.Text))
            If RadioButton1.Checked = True Then
                cmd15.Parameters.AddWithValue("@ESTADO", "1")
            Else
                cmd15.Parameters.AddWithValue("@ESTADO", "2")
            End If

            cmd15.ExecuteNonQuery()
            MsgBox("SE AGREGO CORRECTAMENTE AL NUEVO VENDEDOR")
        Else
            Dim cmd15 As New SqlCommand("UPDATE  custom_vianny.dbo.Vendedores SET  CCia_ven     = @CCia_ven,
										   ficha_ven    = '',
										   alias_ven    = @alias_ven,
										   comision_ven = 0.00,
										   apepat_ven   = @apepat_ven,
										   apemat_ven   = @apemat_ven,
										   nombres_ven  = @nombres_ven,
										   email_ven    = '',
										   fono_ven     = '',
										   rpm_ven      = @ESTADO,
										   admin_ven    = 0,
										   clave_ven    = @clave_ven
									WHERE  Codigo_ven = @Codigo_ven", conx)
            cmd15.Parameters.AddWithValue("@Codigo_ven", Trim(TextBox5.Text))
            cmd15.Parameters.AddWithValue("@CCia_ven", Trim(Label9.Text))
            cmd15.Parameters.AddWithValue("@alias_ven", Trim(TextBox4.Text))
            cmd15.Parameters.AddWithValue("@apepat_ven", Trim(TextBox2.Text))
            cmd15.Parameters.AddWithValue("@apemat_ven", Trim(TextBox3.Text))
            cmd15.Parameters.AddWithValue("@nombres_ven", Trim(TextBox1.Text))
            cmd15.Parameters.AddWithValue("@clave_ven", Trim(TextBox6.Text))
            If RadioButton1.Checked = True Then
                cmd15.Parameters.AddWithValue("@ESTADO", "1")
            Else
                cmd15.Parameters.AddWithValue("@ESTADO", "2")
            End If

            cmd15.ExecuteNonQuery()
            MsgBox("SE ACTUALIZO CORRECTAMENTE LA INFORMACION")
        End If

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        CORRELATIVO()
        MOSTRAR()
        Button1.Enabled = False

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox6.Enabled = False
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "CONSORCIO TEXTIL VIANNY" Then
            Label9.Text = "01"
        Else
            Label9.Text = "02"
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        abrir()
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("SE ELIMINAR TODA LA INFORMACIO DE LA ORDEN DE TEJIDO, PARTIDA Y ROLLOS AGREGADOS ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim agregar1 As String = "DELETE FROM  Vendedores WHERE codigo_ven ='" + Trim(TextBox5.Text) + "'"
            ELIMINAR(agregar1)
            MsgBox("SE ELIMINO LA INFORMACION CORRECTAMENTE")
            MOSTRAR()
            Button3.Enabled = True
        End If

    End Sub
End Class