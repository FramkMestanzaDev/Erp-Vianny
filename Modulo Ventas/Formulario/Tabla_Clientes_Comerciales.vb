Imports System.Data.SqlClient
Public Class Tabla_Clientes_Comerciales
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
        If Trim(ComboBox1.Text) = "VSILVERIO" Then
            Dim cmd As New SqlDataAdapter("select  COD_CLI AS RUC, r_social as CLIENTE,D_fiscal AS DIRECCION, email AS EMAIL,telefono AS TELEFONO,contacto AS CONTACTO,sigla_10 
 AS ABREVIATURA from cliente c left join custom_vianny.dbo.Vendedores v on c.Vendedor= v.codigo_ven LEFT JOIN custom_vianny.dbo.cag1000 B ON C.COD_CLI = B.fich_10 AND B.ccia ='01' WHERE Vendedor ='0005'", conx)
            cmd.Fill(da)
            DataGridView1.DataSource = da
            DataGridView1.Columns(2).Width = 210
            DataGridView1.Columns(1).Width = 90
            DataGridView1.Columns(3).Width = 160
            DataGridView1.Columns(4).Width = 140
            DataGridView1.Columns(5).Width = 90
            DataGridView1.Columns(6).Width = 165
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).ReadOnly = True
            DataGridView1.Columns(6).ReadOnly = True
        Else
            If Trim(ComboBox1.Text) = "GBALVIN" Then
                Dim cmd As New SqlDataAdapter("select  COD_CLI AS RUC, r_social as CLIENTE,D_fiscal AS DIRECCION, email AS EMAIL,telefono AS TELEFONO,contacto AS CONTACTO,sigla_10 
 AS ABREVIATURA from cliente c left join custom_vianny.dbo.Vendedores v on c.Vendedor= v.codigo_ven LEFT JOIN custom_vianny.dbo.cag1000 B ON C.COD_CLI = B.fich_10 AND B.ccia ='01' WHERE Vendedor ='0028'", conx)
                cmd.Fill(da)
                DataGridView1.DataSource = da
                DataGridView1.Columns(2).Width = 210
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(3).Width = 160
                DataGridView1.Columns(4).Width = 170
                DataGridView1.Columns(5).Width = 90
                DataGridView1.Columns(6).Width = 165
                DataGridView1.Columns(1).ReadOnly = True
                DataGridView1.Columns(2).ReadOnly = True
                DataGridView1.Columns(3).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
            Else
                If Trim(ComboBox1.Text) = "VGRAUS" Then
                    Dim cmd As New SqlDataAdapter("select  COD_CLI AS RUC, r_social as CLIENTE,D_fiscal AS DIRECCION, email AS EMAIL,telefono AS TELEFONO,contacto AS CONTACTO,sigla_10
 AS ABREVIATURA from cliente c left join custom_vianny.dbo.Vendedores v on c.Vendedor= v.codigo_ven LEFT JOIN custom_vianny.dbo.cag1000 B ON C.COD_CLI = B.fich_10 AND B.ccia ='01' WHERE Vendedor ='0007'", conx)
                    cmd.Fill(da)
                    DataGridView1.DataSource = da
                    DataGridView1.Columns(2).Width = 210
                    DataGridView1.Columns(1).Width = 90
                    DataGridView1.Columns(3).Width = 160
                    DataGridView1.Columns(4).Width = 140
                    DataGridView1.Columns(5).Width = 90
                    DataGridView1.Columns(6).Width = 165
                    DataGridView1.Columns(1).ReadOnly = True
                    DataGridView1.Columns(2).ReadOnly = True
                    DataGridView1.Columns(3).ReadOnly = True
                    DataGridView1.Columns(4).ReadOnly = True
                    DataGridView1.Columns(5).ReadOnly = True
                    DataGridView1.Columns(6).ReadOnly = True
                Else
                    If Trim(ComboBox1.Text) = "VVARGAS" Then
                        Dim cmd As New SqlDataAdapter("select  COD_CLI AS RUC, r_social as CLIENTE,D_fiscal AS DIRECCION, email AS EMAIL,telefono AS TELEFONO,contacto AS CONTACTO,sigla_10
 AS ABREVIATURA from cliente c left join custom_vianny.dbo.Vendedores v on c.Vendedor= v.codigo_ven LEFT JOIN custom_vianny.dbo.cag1000 B ON C.COD_CLI = B.fich_10 AND B.ccia ='01' WHERE Vendedor ='0041'", conx)
                        cmd.Fill(da)
                        DataGridView1.DataSource = da
                        DataGridView1.Columns(2).Width = 210
                        DataGridView1.Columns(1).Width = 90
                        DataGridView1.Columns(3).Width = 160
                        DataGridView1.Columns(4).Width = 140
                        DataGridView1.Columns(5).Width = 90
                        DataGridView1.Columns(6).Width = 165
                        DataGridView1.Columns(1).ReadOnly = True
                        DataGridView1.Columns(2).ReadOnly = True
                        DataGridView1.Columns(3).ReadOnly = True
                        DataGridView1.Columns(4).ReadOnly = True
                        DataGridView1.Columns(5).ReadOnly = True
                        DataGridView1.Columns(6).ReadOnly = True
                    Else
                        If Trim(ComboBox1.Text) = "JRAMIREZ" Then
                            Dim cmd As New SqlDataAdapter("select  COD_CLI AS RUC, r_social as CLIENTE,D_fiscal AS DIRECCION, email AS EMAIL,telefono AS TELEFONO,contacto AS CONTACTO,sigla_10
 AS ABREVIATURA from cliente c left join custom_vianny.dbo.Vendedores v on c.Vendedor= v.codigo_ven LEFT JOIN custom_vianny.dbo.cag1000 B ON C.COD_CLI = B.fich_10 AND B.ccia ='01' WHERE Vendedor ='0043'", conx)
                            cmd.Fill(da)
                            DataGridView1.DataSource = da
                            DataGridView1.Columns(2).Width = 210
                            DataGridView1.Columns(1).Width = 90
                            DataGridView1.Columns(3).Width = 160
                            DataGridView1.Columns(4).Width = 140
                            DataGridView1.Columns(5).Width = 90
                            DataGridView1.Columns(6).Width = 165
                            DataGridView1.Columns(1).ReadOnly = True
                            DataGridView1.Columns(2).ReadOnly = True
                            DataGridView1.Columns(3).ReadOnly = True
                            DataGridView1.Columns(4).ReadOnly = True
                            DataGridView1.Columns(5).ReadOnly = True
                            DataGridView1.Columns(6).ReadOnly = True
                        Else
                            If Trim(ComboBox1.Text) = "DBRAVO" Then
                                Dim cmd As New SqlDataAdapter("select  COD_CLI AS RUC, r_social as CLIENTE,D_fiscal AS DIRECCION, email AS EMAIL,telefono AS TELEFONO,contacto AS CONTACTO,sigla_10
 AS ABREVIATURA from cliente c left join custom_vianny.dbo.Vendedores v on c.Vendedor= v.codigo_ven LEFT JOIN custom_vianny.dbo.cag1000 B ON C.COD_CLI = B.fich_10 AND B.ccia ='01' WHERE Vendedor ='0023'", conx)
                                cmd.Fill(da)
                                DataGridView1.DataSource = da
                                DataGridView1.Columns(2).Width = 210
                                DataGridView1.Columns(1).Width = 90
                                DataGridView1.Columns(3).Width = 160
                                DataGridView1.Columns(4).Width = 140
                                DataGridView1.Columns(5).Width = 90
                                DataGridView1.Columns(6).Width = 165
                                DataGridView1.Columns(1).ReadOnly = True
                                DataGridView1.Columns(2).ReadOnly = True
                                DataGridView1.Columns(3).ReadOnly = True
                                DataGridView1.Columns(4).ReadOnly = True
                                DataGridView1.Columns(5).ReadOnly = True
                                DataGridView1.Columns(6).ReadOnly = True


                            End If
                        End If
                        End If
                End If
                    End If
        End If
    End Sub
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        mostrar_informacion()
    End Sub

    Private Sub Tabla_Clientes_Comerciales_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            Dim K As Integer

            'ds.Tables(0).Columns.Add("SELECCIONAR")
            'DataGridView1.DataSource = ds.Tables(0)
            'Dim F As BindingSource = New BindingSource
            'F.DataSource = DataGridView1.DataSource
            'F.Filter = "PARTIDA" & " like '%" & TextBox1.Text & "%'"
            'DataGridView1.DataSource = F
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))
            Dim jk As Integer

            dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                jk = DataGridView1.Rows.Count
                DataGridView1.DataSource = da
                DataGridView1.Columns(2).Width = 210
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(3).Width = 160
                DataGridView1.Columns(4).Width = 140
                DataGridView1.Columns(5).Width = 90
                DataGridView1.Columns(6).Width = 165
                DataGridView1.Columns(1).ReadOnly = True
                DataGridView1.Columns(2).ReadOnly = True
                DataGridView1.Columns(3).ReadOnly = True
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(5).ReadOnly = True
                DataGridView1.Columns(6).ReadOnly = True
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception


        End Try
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString
            Registro_Cli_Comer.TextBox1.Text = Trim(DataGridView1.Rows(num1).Cells(1).Value)
            Registro_Cli_Comer.TextBox2.Text = Trim(DataGridView1.Rows(num1).Cells(2).Value)
            Registro_Cli_Comer.TextBox3.Text = Trim(DataGridView1.Rows(num1).Cells(3).Value)
            Registro_Cli_Comer.TextBox4.Text = Trim(DataGridView1.Rows(num1).Cells(6).Value)
            Registro_Cli_Comer.TextBox6.Text = Trim(DataGridView1.Rows(num1).Cells(4).Value)
            Registro_Cli_Comer.TextBox7.Text = Trim(DataGridView1.Rows(num1).Cells(5).Value)
            Registro_Cli_Comer.ShowDialog()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class