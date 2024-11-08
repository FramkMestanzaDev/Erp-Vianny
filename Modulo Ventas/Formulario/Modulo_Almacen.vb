Imports System.Windows.Forms
Imports System.Data.SqlClient
Public Class Modulo_Almacen
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
            Case "Stock Fisico"
                Stock_Fisico.Label5.Text = Label8.Text
                Stock_Fisico.Label6.Text = Label10.Text
                Stock_Fisico.Show(Me)
            Case "Nota de Pedido Comercial"
                validar_Almacen.TextBox3.Text = 2
                validar_Almacen.Label6.Text = Label8.Text
                validar_Almacen.Label7.Text = Label10.Text
                validar_Almacen.Label8.Text = Label9.Text
                validar_Almacen.Label9.Text = Label6.Text
                validar_Almacen.Show(Me)
            Case "Nota Ingreso T"
                NotaIngreso.TextBox7.Text = Label6.Text
                NotaIngreso.Label11.Text = Label8.Text
                NotaIngreso.Label14.Text = Label10.Text
                NotaIngreso.Show(Me)
            Case "Nota Salida T"
                Nsalida.TextBox7.Text = Label6.Text
                Nsalida.Label10.Text = Label8.Text
                Nsalida.Label13.Text = Label10.Text
                Nsalida.Show(Me)
            Case "Guia Remision T"
                Almacen_guia_nuevo.Label1.Text = Label8.Text
                Almacen_guia_nuevo.Label2.Text = Label10.Text
                Almacen_guia_nuevo.Label3.Text = Label6.Text
                Almacen_guia_nuevo.Label4.Text = "1"
                Almacen_guia_nuevo.Label6.Text = "0"
                Almacen_guia_nuevo.Show(Me)
            Case "Nota Ingreso H"
                NiaHilo.TextBox7.Text = Label6.Text
                NiaHilo.Label11.Text = Label8.Text
                NiaHilo.Label13.Text = Label10.Text
                NiaHilo.TextBox14.Text = 2
                NiaHilo.TextBox16.Text = 2
                NiaHilo.Show(Me)
            Case "Nota Salida H"
                NsaHilo.TextBox7.Text = Label6.Text
                NsaHilo.Label10.Text = Label8.Text
                NsaHilo.Label13.Text = Label10.Text
                NsaHilo.TextBox12.Text = 2
                NsaHilo.TextBox18.Text = 2
                NsaHilo.Show(Me)
            Case "Guia Remision H"
                Almacen_guia_nuevo.Label1.Text = Label8.Text
                Almacen_guia_nuevo.Label2.Text = Label10.Text
                Almacen_guia_nuevo.Label3.Text = Label6.Text
                Almacen_guia_nuevo.Label4.Text = "2"
                Almacen_guia_nuevo.Show(Me)
            Case "Nota Ingreso S"
                NiaHilo.TextBox7.Text = Label6.Text
                NiaHilo.Label11.Text = Label8.Text
                NiaHilo.Label13.Text = Label10.Text
                NiaHilo.TextBox14.Text = 1
                NiaHilo.TextBox16.Text = 3
                NiaHilo.Show(Me)
            Case "Nota Salida S"
                NsaHilo.TextBox7.Text = Label6.Text
                NsaHilo.Label10.Text = Label8.Text
                NsaHilo.Label13.Text = Label10.Text
                NsaHilo.TextBox18.Text = 3
                NsaHilo.TextBox12.Text = 1
                NsaHilo.Show(Me)
            Case "Guia Remision S"
                Almacen_guia_nuevo.Label1.Text = Label8.Text
                Almacen_guia_nuevo.Label2.Text = Label10.Text
                Almacen_guia_nuevo.Label3.Text = Label10.Text
                Almacen_guia_nuevo.Label4.Text = "3"
                Almacen_guia_nuevo.Show(Me)
            Case "Nota Ingreso R"
                NiaHilo.TextBox7.Text = Label6.Text
                NiaHilo.Label11.Text = Label8.Text
                NiaHilo.Label13.Text = Label10.Text
                NiaHilo.TextBox14.Text = 3
                NiaHilo.TextBox16.Text = 4
                NiaHilo.Show(Me)
            Case "Nota Salida R"
                NsaHilo.TextBox7.Text = Label6.Text
                NsaHilo.Label10.Text = Label8.Text
                NsaHilo.Label13.Text = Label10.Text
                NsaHilo.TextBox12.Text = 3
                NsaHilo.TextBox18.Text = 4
                NsaHilo.Show(Me)
            Case "Guia Remision R"
                Almacen_guia_nuevo.Label1.Text = Label8.Text
                Almacen_guia_nuevo.Label2.Text = Label10.Text
                Almacen_guia_nuevo.Label3.Text = Label6.Text
                Almacen_guia_nuevo.Label4.Text = "4"
                Almacen_guia_nuevo.Show(Me)
            Case "Nota Ingreso A"
                NiaHilo.TextBox7.Text = Label6.Text
                NiaHilo.Label11.Text = Label8.Text
                NiaHilo.Label13.Text = Label10.Text
                NiaHilo.TextBox16.Text = 5
                NiaHilo.Show(Me)
            Case "Nota Salida A"
                NsaHilo.TextBox7.Text = Label6.Text
                NsaHilo.Label10.Text = Label8.Text
                NsaHilo.Label13.Text = Label10.Text
                NsaHilo.TextBox12.Text = 4
                NsaHilo.TextBox18.Text = 5
                NsaHilo.Show(Me)
            Case "Guia Remision A"
                Guia_hilo.Label23.Text = "13"
                Guia_hilo.Label25.Text = "  ALMACEN AVIOS"
                Guia_hilo.TextBox19.Text = Label6.Text
                Guia_hilo.Label29.Text = Label8.Text
                Guia_hilo.TextBox1.Text = "T005"
                Guia_hilo.Label30.Text = Label10.Text
                Guia_hilo.Show(Me)
            Case "Nota Ingreso PT"
                Nia_Pt.TextBox7.Text = Label6.Text
                Nia_Pt.Label11.Text = Label8.Text
                Nia_Pt.Label13.Text = Label10.Text
                Nia_Pt.Show(Me)
            Case "Nota Salida PT"
                Nsa_Pt.TextBox7.Text = Label6.Text
                Nsa_Pt.Label11.Text = Label8.Text
                Nsa_Pt.Label13.Text = Label10.Text
                Nsa_Pt.Show(Me)
            Case "Guia Remision PT"
                Guia_remision_prenda.Label33.Text = Label10.Text
                Guia_remision_prenda.Label35.Text = "01"
                Guia_remision_prenda.Label36.Text = Label8.Text
                Guia_remision_prenda.Show(Me)
            Case "Packing Tela"
                Planta_packing.Show(Me)
            Case "Parte Ingreso"
                Dim kl1 As New Reporte_Ingreso_Salida
                Reporte_Ingreso_Salida.TextBox4.Text = Label8.Text
                Reporte_Ingreso_Salida.TextBox5.Text = 1
                Reporte_Ingreso_Salida.Text = "REPORTE PARTE INGRESO"
                Reporte_Ingreso_Salida.Label4.Text = Label10.Text
                Reporte_Ingreso_Salida.Show(Me)
            Case "Parte Salida"
                Dim kl As New Reporte_Ingreso_Salida
                Reporte_Ingreso_Salida.TextBox4.Text = Label8.Text
                Reporte_Ingreso_Salida.TextBox5.Text = 2
                Reporte_Ingreso_Salida.Text = "REPORTE PARTE SALIDA"
                Reporte_Ingreso_Salida.Label4.Text = Label10.Text
                Reporte_Ingreso_Salida.Show(Me)
            Case "Calidad Partidas"
                Calidad_tela__acabada.Label3.Text = Label8.Text
                Calidad_tela__acabada.Show(Me)
            Case "Packing Tela"
                buscar_Packing.Show(Me)
            Case "Sticker Calidad"
                Dim jh As New Rpt_Rama
                Dim clave As String


                clave = InputBox("Introduzca la Partida")

                abrir()
                '-------------------------
                Dim sql102 As String = "select substring(ncom_4,1,10),substring(ncom_4,11,3) from custom_vianny.dbo.matg040f where partidA_4 ='" + Trim(clave) + "'"
                Dim cmd102 As New SqlCommand(sql102, conx)
                Rsr2 = cmd102.ExecuteReader()

                If Rsr2.Read() Then
                    Impr_calidad.TextBox1.Text = Rsr2(0)
                    Impr_calidad.TextBox2.Text = Rsr2(1)

                    Impr_calidad.Show(Me)
                    Rsr2.Close()
                Else
                    MsgBox("LA PARTIDA NO EXISTE")
                End If
            Case "Movimientos Almacen"
                Rpt_ing_sal.Label7.Text = Label10.Text
                Rpt_ing_sal.Show(Me)
            Case "Produccion Corte"
                Produccion_diaria_cortee.Label3.Text = Label9.Text
                Produccion_diaria_cortee.Show(Me)
            Case "Requerimiento Tela"
                Consumo_tela_Consolidado.Show(Me)
            Case "Orden Compra"
                Ocs.Label16.Text = Label10.Text
                Ocs.Label17.Text = Label8.Text
                Ocs.Show(Me)
            Case "Orden Servicio"
                os.Label16.Text = Label10.Text
                os.Label17.Text = Label8.Text
                os.Show(Me)
            Case "Rpt Orden Compra"
                Orden_Compra_Ac.Label3.Text = Label10.Text
                Orden_Compra_Ac.Show(Me)
            Case "Rpt Orden Servicio"
                ORDE_SERVISO_AC.Label3.Text = Label10.Text
                ORDE_SERVISO_AC.Show(Me)
            Case "Matriz Avios-Almacen"
                FormatoMa.Label1.Text = Label10.Text
                FormatoMa.Label2.Text = Label8.Text
                FormatoMa.Label3.Text = Label6.Text
                FormatoMa.Show(Me)
            Case "Almacen de Solicitudes A-P"
                modulo_solicitud.Label1.Text = Label6.Text
                modulo_solicitud.Label2.Text = Label8.Text
                modulo_solicitud.Label3.Text = Label10.Text
                modulo_solicitud.Show(Me)
            Case "Registro Ventas"
                almacen_nuevo.Label3.Text = Label8.Text
                almacen_nuevo.Label2.Text = Label10.Text
                almacen_nuevo.Show(Me)
            Case "Almacen Requerimiento Tela"
                ReqTelaProduccion.Label5.Text = Label10.Text
                ReqTelaProduccion.Label8.Text = Label8.Text
                ReqTelaProduccion.Label9.Text = Label6.Text
                ReqTelaProduccion.Label10.Text = Label7.Text
                ReqTelaProduccion.Show(Me)
            Case "Almacen Orden Compra"
                AlmacenOrdenCompra.Label3.Text = Label10.Text
                AlmacenOrdenCompra.Label4.Text = Label8.Text
                AlmacenOrdenCompra.Label5.Text = Label6.Text
                AlmacenOrdenCompra.Show(Me)
            Case "Requerimiento Avios"
                RequerimientoAvios.Label5.Text = Label10.Text
                RequerimientoAvios.Label8.Text = Label8.Text
                RequerimientoAvios.Label9.Text = Label6.Text
                RequerimientoAvios.Label10.Text = Label7.Text
                RequerimientoAvios.Label17.Text = Label6.Text
                RequerimientoAvios.Show(Me)
            Case "Registro de Incidencias"
                RegisroInconvenientes.Show(Me)
            Case "Hoja Cotizacion"
                AlmacenCotizacion.Label3.Text = Label7.Text
                AlmacenCotizacion.Show(Me)
            Case "Status Op"
                Pre_produccion.TextBox1.Text = Label8.Text
                Pre_produccion.Show(Me)
            Case "Status Produccion"
                Form7.TextBox4.Text = Label10.Text
                Form7.TextBox1.Text = Label9.Text
                Form7.Show(Me)
            Case "Matriz Avios (Monitor)"
                MatrizAviosMonitor.Label3.Text = Label10.Text
                MatrizAviosMonitor.Show(Me)
            Case "Explosion Logistica Avios"
                ExplosionAvios.Label3.Text = Label10.Text
                ExplosionAvios.Show(Me)
        End Select

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class
