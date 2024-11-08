Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class ModuloComercial
    Public comando As SqlCommand
    Public conx As SqlConnection
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
    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
        ' Cree una nueva instancia del formulario secundario.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Conviértalo en un elemento secundario de este formulario MDI antes de mostrarlo.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Ventana " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs)
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: agregue código aquí para abrir el archivo.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: agregue código aquí para guardar el contenido actual del formulario en un archivo.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Utilice My.Computer.Clipboard para insertar el texto o las imágenes seleccionadas en el Portapapeles
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Utilice My.Computer.Clipboard para insertar el texto o las imágenes seleccionadas en el Portapapeles
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Utilice My.Computer.Clipboard.GetText() o My.Computer.Clipboard.GetData para recuperar la información del Portapapeles.
    End Sub



    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Cierre todos los formularios secundarios del principal.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

    End Sub
    Dim Rsr2 As SqlDataReader

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Dim imageList As New ImageList()
        imageList.Images.Add(Global.Modulo_Ventas.Resources._3d_applications_folder_20538)
        imageList.Images.Add(Global.Modulo_Ventas.Resources.bindersbluedocuments_carpetas_azu_10985)
        TreeView1.ImageList = imageList
        e.Node.ImageIndex = 0
        e.Node.SelectedImageIndex = 1
        Select Case e.Node.Text
            Case "Orden Desarrollo"

                OD.TextBox15.Text = Label6.Text
                OD.TextBox7.Text = Label9.Text
                OD.Label23.Text = Label10.Text
                OD.Show(Me)
            Case "Orden Produccion"

                Op_Manufactura.TextBox10.Text = Label6.Text
                Op_Manufactura.Label33.Text = Label10.Text
                Op_Manufactura.Show(Me)
            Case "Programacion Pedidos"

                Programa_Tejeduria.Label3.Text = 3
                Programa_Tejeduria.Show(Me)
            Case "Comision"

                Comisiones_ven.TextBox3.Text = Label6.Text
                Comisiones_ven.Label7.Text = Label8.Text
                Comisiones_ven.Label8.Text = Label10.Text
                Comisiones_ven.Label13.Text = Label9.Text
                Comisiones_ven.Label15.Text = 0
                Comisiones_ven.Show(Me)
            Case "Calidad Partidas"

                Calidad_tela__acabada.Label3.Text = Label8.Text
                Calidad_tela__acabada.Show(Me)
            Case "Hoja Reclamo"

                abrir()
                Dim sql102 As String = "SELECT codigo_ven FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + Label6.Text + "'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()
                If Rsr2.Read() = True Then
                    Hoja_Reclamos.TextBox14.Text = Rsr2(0)
                End If
                Rsr2.Close()
                Hoja_Reclamos.TextBox9.Text = Label6.Text
                Hoja_Reclamos.Label26.Text = Label10.Text
                Hoja_Reclamos.Label27.Text = Label8.Text
                Hoja_Reclamos.Show(Me)
            Case "Informacion Ventas"

                Rpt_Ventas_ven.TextBox1.Text = Label8.Text
                Rpt_Ventas_ven.TextBox2.Text = Label6.Text
                Rpt_Ventas_ven.TextBox3.Text = Label9.Text
                Rpt_Ventas_ven.Label5.Text = Label10.Text
                abrir()
                Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + Label6.Text + "' and ccia_ven ='01'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()
                If Rsr2.Read() = True Then
                    Rpt_Ventas_ven.Label6.Text = Rsr2(0)
                End If
                Rsr2.Close()
                Rpt_Ventas_ven.Show(Me)
            Case "Listado de OD x PM"

                OdxPm.TextBox1.Text = Label9.Text
                OdxPm.Label3.Text = Label10.Text
                OdxPm.Label4.Text = "1"
                OdxPm.Label2.Text = "PM"
                OdxPm.Show(Me)
            Case "Listado de OP x PO"

                OdxPm.TextBox1.Text = Label9.Text
                OdxPm.Label3.Text = Label10.Text
                OdxPm.Label4.Text = "2"
                OdxPm.Label2.Text = "PO"
                OdxPm.Show(Me)
            Case "Stock Fisico"

                Stock_fisico_comerci.Label5.Text = Label8.Text
                Stock_fisico_comerci.Label6.Text = Label10.Text
                Stock_fisico_comerci.Show(Me)

            Case "Nota de Pedido"

                Select Case Label4.Text.ToString.Trim
                    Case "COMERCIAL TEXTIL" : Nota_Pedido.TextBox18.Text = 1
                    Case "COMERCIAL MANUFACTURA" : Nota_Pedido.TextBox18.Text = 2

                End Select
                abrir()
                Dim sql102 As String = "SELECT codigo_ven FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + Label6.Text + "' "
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()
                If Rsr2.Read() = True Then
                    Nota_Pedido.TextBox8.Text = Rsr2(0)
                    Nota_Pedido.TextBox22.Text = Rsr2(0)
                End If

                Rsr2.Close()


                Nota_Pedido.TextBox9.Text = Label6.Text
                Nota_Pedido.Label25.Text = Label8.Text
                Nota_Pedido.Label26.Text = Label10.Text
                Nota_Pedido.Label27.Text = Label9.Text
                Nota_Pedido.Label28.Text = Label7.Text

                Nota_Pedido.Show(Me)

            Case "Status de Od"

                LisOd.Label3.Text = Label10.Text
                LisOd.TextBox1.Text = Label6.Text
                LisOd.Show(Me)
            Case "Lista de Op"

                LisOp.Label3.Text = Label10.Text
                LisOp.TextBox1.Text = Label6.Text
                LisOp.Show(Me)


            Case "Reporte Orden Compra"

                Orden_Compra_Ac.Label3.Text = Label10.Text
                Orden_Compra_Ac.Show(Me)
            Case "Crear Familia"
                MCoFamilia.Show(Me)
            Case "Crear Talla"
                CrearTalla.Show(Me)
            Case "Crear Padron Tallas"
                PadronTalla.Show(Me)
            Case "Crear Marca"
                TablaMarca.Show(Me)
            Case "Crear Color"
                TablaColor.Show(Me)
            Case "Orden Compra"
                Ocs.Label16.Text = Label10.Text
                Ocs.Label17.Text = Label8.Text
                Ocs.Show(Me)
            Case "Abreviatura Clientes"
                Tabla_Clientes_Comerciales.Show(Me)
            Case "Status Op"
                StatusProduccionPP.Label2.Text = Label10.Text
                StatusProduccionPP.Label5.Text = Label8.Text
                StatusProduccionPP.TextBox1.Text = Label9.Text
                StatusProduccionPP.Show(Me)
            Case "Hoja Cotizacion"
                ListaHojaCotizacion.Label3.Text = Label7.Text
                ListaHojaCotizacion.Show(Me)

        End Select
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class
