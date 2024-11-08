Imports System.Data.SqlClient
Public Class BUSCAR_VENTA
    Dim dt As New DataTable
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
    Private Sub BUSCAR_VENTA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Label1.Text = 1 Then
            Dim hi As New fventasn
            Dim jk As New vventas_n
            jk.gperiodo = TextBox2.Text
            jk.gTIPO = TextBox3.Text
            dt = hi.BUSCAR_DETALLE_LIBRE(jk)
            If dt.Rows.Count > 0 Then
                DataGridView1.DataSource = dt
                DataGridView1.Columns(0).Width = 80
                DataGridView1.Columns(1).Width = 100
                DataGridView1.Columns(2).Width = 250
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 80
                DataGridView1.Columns(5).Width = 50
                DataGridView1.Columns(6).Width = 200
            End If
        Else
            If Label1.Text = 2 Then
                abrir()
                Dim cmd As New SqlDataAdapter("SELECT ncom_3 AS REGISTRO,sfactu_3+ '-'+nfactu_3 AS FACTURA,nomb_3 AS CLIENTE , CASE WHEN  mone_3 = 1 THEN 'SOLES' ELSE 'DOLARES' END AS MOENDAS,CASE WHEN mone_3 = 1 THEN tot1_3 ELSE tot2_3 END AS VENTA,gloa_3 as GLOSA  
FROM custom_vianny.DBO.fag0300 where tienda_3 ='" + Trim(Registro_Facturas.TextBox10.Text) + "' and SUBSTRING(ncom_3,1,2) = '" + Trim(Registro_Facturas.TextBox2.Text) + "' and cperiodo_3='" + Trim(Registro_Facturas.Label27.Text) + "' and ccia_3 ='" + Trim(Registro_Facturas.Label29.Text) + "'", conx)

                Dim da As New DataTable

                cmd.Fill(da)
                DataGridView1.DataSource = da
                DataGridView1.Columns(0).Width = 80
                DataGridView1.Columns(1).Width = 100
                DataGridView1.Columns(2).Width = 250
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 80
                DataGridView1.Columns(5).Width = 320
            Else
                If Label1.Text = 3 Then
                    abrir()
                    Dim cmd As New SqlDataAdapter("SELECT ncom_3 AS REGISTRO,sfactu_3+ '-'+nfactu_3 AS FACTURA,nomb_3 AS CLIENTE , CASE WHEN  mone_3 = 1 THEN 'SOLES' ELSE 'DOLARES' END AS MOENDAS,CASE WHEN mone_3 = 1 THEN tot1_3 ELSE tot2_3 END AS VENTA,gloa_3 as GLOSA  
FROM custom_vianny.DBO.fag0300 where tienda_3 ='" + Trim(Registro_Facturas.TextBox10.Text) + "'  and cperiodo_3='" + Trim(Registro_Facturas.Label27.Text) + "' and ccia_3 ='" + Trim(Registro_Facturas.Label29.Text) + "'", conx)

                    Dim da As New DataTable

                    cmd.Fill(da)
                    DataGridView1.DataSource = da
                    DataGridView1.Columns(0).Width = 80
                    DataGridView1.Columns(1).Width = 100
                    DataGridView1.Columns(2).Width = 250
                    DataGridView1.Columns(3).Width = 80
                    DataGridView1.Columns(4).Width = 80
                    DataGridView1.Columns(5).Width = 320
                End If
            End If
        End If
    End Sub
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(dt.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "CLIENTE" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv
                DataGridView1.Columns(0).Width = 80
                DataGridView1.Columns(1).Width = 90
                DataGridView1.Columns(2).Width = 200
                DataGridView1.Columns(3).Width = 80
                DataGridView1.Columns(4).Width = 80
                DataGridView1.Columns(5).Width = 50
                DataGridView1.Columns(5).Width = 150
            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub


    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If Label1.Text = 1 Then
            Registro_Ventas.TextBox4.Text = Mid(DataGridView1.CurrentRow.Cells(0).Value, 1, 2)
            Registro_Ventas.TextBox3.Text = Mid(DataGridView1.CurrentRow.Cells(0).Value, 3, 6)
            Registro_Ventas.TextBox3.Select()
            Me.Close()
        Else
            If Label1.Text = 2 Then
                Registro_Facturas.TextBox2.Text = Mid(DataGridView1.CurrentRow.Cells(0).Value, 1, 2)
                Registro_Facturas.TextBox1.Text = Mid(DataGridView1.CurrentRow.Cells(0).Value, 3, 6)
                Registro_Ventas.TextBox1.Select()
                Registro_Ventas.TextBox1.Focus()
                SendKeys.Send("{ENTER}")
                Me.Close()
            Else
                If Label1.Text = 2 Then
                    Registro_Facturas.TextBox2.Text = Mid(DataGridView1.CurrentRow.Cells(0).Value, 1, 2)
                    Registro_Facturas.TextBox1.Text = Mid(DataGridView1.CurrentRow.Cells(0).Value, 3, 6)
                    Registro_Ventas.TextBox1.Select()

                    Registro_Ventas.TextBox1.Focus()
                    SendKeys.Send("{ENTER}")
                    Me.Close()
                End If
            End If
        End If
    End Sub
End Class