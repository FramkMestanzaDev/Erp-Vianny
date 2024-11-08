Imports System.Data.SqlClient
Public Class DetalleMatrizAvios
    Public conx As SqlConnection
    Dim da As New DataTable
    Dim Rsr2 As SqlDataReader
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
    Dim rs1, Rsr22 As SqlDataReader
    Private Sub DetalleMatrizAvios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        DataGridView1.Rows.Clear()
        Dim sql10112 As String = "select  COUNT(ncom_8c) from custom_vianny.dbo.qamc800ct where ccia_8c ='" + Label1.Text + "' and ncom_8c ='" + Label2.Text + "' and item_8c ='" + Label3.Text + "'"
        Dim cmd10112 As New SqlCommand(sql10112, conx)
        rs1 = cmd10112.ExecuteReader()
        If rs1.Read() = True And rs1(0) = 0 Then
            rs1.Close()
            Dim sql102 As String = "EXEC custom_vianny.DBO.CaeSoft_GeneraDetalleDetalleMatrizConsumoAvios '" + Label1.Text + "','" + Label2.Text + "','" + Label3.Text + "'"
            Dim cmd102 As New SqlCommand(sql102, conx)
            Rsr2 = cmd102.ExecuteReader()
            Dim i5 As Integer
            i5 = 0
            While Rsr2.Read()
                DataGridView1.Rows.Add()
                DataGridView1.Rows(i5).Cells(0).Value = Rsr2(3)
                DataGridView1.Rows(i5).Cells(1).Value = Rsr2(5)
                DataGridView1.Rows(i5).Cells(7).Value = Rsr2(2)
                DataGridView1.Rows(i5).Cells(8).Value = Rsr2(4)
                DataGridView1.Rows(i5).Cells(9).Value = Rsr2(1)
                i5 = i5 + 1
            End While
            Rsr2.Close()
        Else
            rs1.Close()
            Dim sql1 As String = "select c.*,A.CORREL_16B , CONVERT(CHAR(70) , A.color_16B),q.CORREL_16C , CONVERT(CHAR(70) , q.Talla_16C),f.nomb_17 from custom_vianny.dbo.qamc800ct c 
                                 inner join custom_vianny.dbo.qag160b A on  A.ccia + A.ncom_16b = c.ccia_8c + c.ncom_8c and a.CORREL_16B=c.color_8c
                                 inner join custom_vianny.dbo.qag160c q on q.ccia + q.ncom_16c =  c.ccia_8c+ c.ncom_8c and c.talla_8c= q.correl_16c
                                 left join custom_vianny.dbo.cag1700 f on c.ccia_8c = f.ccia and c.linea_8c = f.linea_17
                                 where ccia_8c ='" + Label1.Text + "' and ncom_8c ='" + Label2.Text + "' and item_8c ='" + Label3.Text + "' order by correl_16c"
            Dim cmd1 As New SqlCommand(sql1, conx)
            Rsr22 = cmd1.ExecuteReader()
            Dim i5 As Integer
            i5 = 0
            While Rsr22.Read()
                DataGridView1.Rows.Add()

                DataGridView1.Rows(i5).Cells(0).Value = Rsr22(13)
                DataGridView1.Rows(i5).Cells(1).Value = Rsr22(15)
                DataGridView1.Rows(i5).Cells(3).Value = Rsr22(3)
                DataGridView1.Rows(i5).Cells(4).Value = Rsr22(16)
                DataGridView1.Rows(i5).Cells(6).Value = Rsr22(10)
                DataGridView1.Rows(i5).Cells(7).Value = Rsr22(12)
                DataGridView1.Rows(i5).Cells(8).Value = Rsr22(14)
                DataGridView1.Rows(i5).Cells(9).Value = Rsr22(1)
                i5 = i5 + 1
            End While
            Rsr22.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim I18, VAL As Integer
        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add(1)
        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(0).Value
        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(7).Value
        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = Label3.Text
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
        ' Ponemos la celda en modo de edición
        DataGridView1.BeginEdit(True)
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "A" Then
            Dim fechaActual As Date = Today

            pproductos.CCIA.Text = Label1.Text
            pproductos.ALMACEN.Text = "13"
            pproductos.ANO.Text = DateTime.Now.Year
            pproductos.FECHA.Text = Replace(fechaActual.ToString("yyyy-MM-dd"), "-", "")
            pproductos.TextBox1.Text = "13"
            pproductos.Label3.Text = 13
            pproductos.Label5.Text = e.RowIndex
            pproductos.TextBox3.Text = ""
            pproductos.TextBox2.Text = ""
            pproductos.ShowDialog()
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Obs" Then
            ObsMatriz.TextBox1.Text = ""
            ObsMatriz.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(6).Value.ToString().Trim()
            ObsMatriz.Label1.Text = e.RowIndex
            ObsMatriz.Label2.Text = 2
            ObsMatriz.ShowDialog()
        End If

    End Sub
    Private WithEvents editingControl As Control
    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        If editingControl IsNot Nothing Then
            RemoveHandler editingControl.KeyDown, AddressOf EditingControl_KeyDown
        End If

        ' Asignar el nuevo control de edición y agregar el evento KeyDown
        editingControl = e.Control
        AddHandler editingControl.KeyDown, AddressOf EditingControl_KeyDown
    End Sub
    Dim sumaa, sumat As Integer
    Function validarArticulo()
        sumaa = 0
        For i = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(i).Cells(3).Value Is Nothing OrElse String.IsNullOrEmpty(DataGridView1.Rows(i).Cells(3).Value.ToString()) Then
                sumaa = sumaa + 1
            End If
        Next
        Return sumaa
    End Function
    Function validarTalla()
        sumat = 0
        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(1).Value Is Nothing OrElse String.IsNullOrEmpty(DataGridView1.Rows(i).Cells(1).Value.ToString()) Then
                sumat = sumat + 1
            End If
        Next
        Return sumat
    End Function
    Private Sub DetalleMatrizAvios_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        validarArticulo()
        validarTalla()

        If validarTalla() > 0 Or validarArticulo() > 0 Then
            MsgBox("EXISTE UN CAMPO EN LA TALLA O ARTICULO QUE FALTA AGREGAR INFORMACION, AGREGE O ELIMINE EL CAMPO PARA CONTINUAR O PRESIONE LIMPIAR PARA BORRAR TODAS LAS FILAS")
            e.Cancel = True
            Else
            Dim itm As String = ""
            Dim agregar As String = "delete from custom_vianny.dbo.qamc800ct where ncom_8c ='" + Trim(Label2.Text) + "' and ccia_8c ='" + Trim(Label1.Text) + "' and item_8c ='" + Trim(Label3.Text) + "'"
                ELIMINAR(agregar)
                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim cmd16 As New SqlCommand("insert into   custom_vianny.dbo.qamc800ct (ccia_8c,item_8c,ncom_8c,linea_8c,color_8c,talla_8c,factor_8c,udm_8c,gene_8c,nomb_8c,obser_8c,cliente_8c) 
        	 values (@ccia_8c,@item_8c,@ncom_8c,@linea_8c,@color_8c,@talla_8c,@factor_8c,@udm_8c,@gene_8c,@nomb_8c,@obser_8c,@cliente_8c)", conx)
                    cmd16.Parameters.AddWithValue("@ccia_8c", Trim(Label1.Text))
                    Select Case Label3.Text.ToString().Trim().Length
                    Case 1 : itm = "0" & Label3.Text.ToString().Trim()
                    Case 2 : itm = Label3.Text.ToString().Trim()
                End Select
                cmd16.Parameters.AddWithValue("@item_8c", Trim(itm))
                cmd16.Parameters.AddWithValue("@ncom_8c", Trim(Label2.Text))
                    cmd16.Parameters.AddWithValue("@linea_8c", DataGridView1.Rows(i).Cells(3).Value.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@color_8c", DataGridView1.Rows(i).Cells(7).Value.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@talla_8c", DataGridView1.Rows(i).Cells(8).Value.ToString().Trim())
                    cmd16.Parameters.AddWithValue("@factor_8c", 0)
                    cmd16.Parameters.AddWithValue("@udm_8c", "")
                    cmd16.Parameters.AddWithValue("@gene_8c", "")
                    cmd16.Parameters.AddWithValue("@nomb_8c", "")
                    If IsDBNull(DataGridView1.Rows(i).Cells(6).Value) Or DataGridView1.Rows(i).Cells(6).Value Is Nothing Then
                        cmd16.Parameters.AddWithValue("@obser_8c", "")
                    Else
                        cmd16.Parameters.AddWithValue("@obser_8c", DataGridView1.Rows(i).Cells(6).Value.ToString().Trim())
                    End If

                    cmd16.Parameters.AddWithValue("@cliente_8c", "")
                    cmd16.ExecuteNonQuery()
                Next

            End If





    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DataGridView1.Rows.Clear()
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
    Private Sub EditingControl_KeyDown(sender As Object, e As KeyEventArgs)
        ' Verificar si la tecla presionada es F1
        If e.KeyCode = Keys.F1 Then
            ' Verificar si la celda actual está en la columna "Factor de Consumo"
            If DataGridView1.CurrentCell.ColumnIndex = 1 Then
                ' Cancelar la edición para evitar conflictos
                DataGridView1.EndEdit()
                ' Abrir el formulario deseado
                TallaMatriz.Label2.Text = Label1.Text
                TallaMatriz.Label3.Text = Label2.Text
                TallaMatriz.Label4.Text = DataGridView1.CurrentCell.RowIndex.ToString
                TallaMatriz.ShowDialog()
            End If

        End If
    End Sub
End Class