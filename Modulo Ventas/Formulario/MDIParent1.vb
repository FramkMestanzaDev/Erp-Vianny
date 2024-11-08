Imports System.Data.SqlClient
Public Class MDIParent1
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

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Select Case Label4.Text.ToString.Trim
            Case "COMERCIAL TEXTIL" : Nota_Pedido.TextBox18.Text = 1
            Case "COMERCIAL MANUFACTURA" : Nota_Pedido.TextBox18.Text = 2

        End Select
        abrir()
        Dim sql102 As String = "SELECT codigo_ven FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + Label3.Text + "' "
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            Nota_Pedido.TextBox8.Text = Rsr2(0)
            Nota_Pedido.TextBox22.Text = Rsr2(0)
        End If

        Rsr2.Close()


        Nota_Pedido.TextBox9.Text = Label3.Text
        Nota_Pedido.Label25.Text = Label5.Text
        Nota_Pedido.Label26.Text = Label9.Text
        Nota_Pedido.Label27.Text = Label7.Text
        Nota_Pedido.Label28.Text = Label4.Text

        Nota_Pedido.Show()
    End Sub

    Private Sub AprobacionPedidosToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Aprobar_notapedido.Show()
    End Sub

    Private Sub AprobacionPedidosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Aprobacion_comercial.Show()
    End Sub

    Private Sub ValidarAlmacenToolStripMenuItem_Click(sender As Object, e As EventArgs)
        validar_Almacen.Show()
    End Sub

    Private Sub HojaDeRutaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Transporte.Show()
    End Sub

    Private Sub ClientesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClientesToolStripMenuItem.Click
        Form_Clientes.Show()
    End Sub

    Private Sub UsuariosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form1.Show()
    End Sub

    Private Sub DireccionDespacjToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Direccion_Despacho.Show()
    End Sub

    Private Sub MDIParent1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If LoginForm1.ComboBox1.Text = "JEFATURA COMERCIAL" Or LoginForm1.ComboBox1.Text = "ADMINISTRADOR" Then
            AprobacionPedidosToolStripMenuItem.Enabled = True

            ToolStripButton3.Enabled = True


        End If
    End Sub

    Private Sub COMERCIALToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles COMERCIALToolStripMenuItem.Click

    End Sub

    Private Sub MDIParent1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        LoginForm1.Close()
    End Sub

    Private Sub PackingTelaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Ingresar_Packing.Show()
    End Sub

    Private Sub PackingHiloToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub COBRANZASToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ALMACENToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ALMACENToolStripMenuItem.Click

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        ListaHojaCotizacion.Label3.Text = Label4.Text
        ListaHojaCotizacion.ShowDialog()
    End Sub

    Private Sub COTABToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Contabilidad.Show()
    End Sub

    Private Sub TESORERIAToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub RegistroDocumentosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'Registro_Almacen.Show()
        Nsalida.TextBox7.Text = Label3.Text
        Nsalida.Label10.Text = Label5.Text
        Nsalida.Show()
    End Sub

    Private Sub RegistroDocumentosToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Reporte_Documentos.Show()
    End Sub

    Private Sub TotalDocumentosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_Facturas.Show()
    End Sub

    Private Sub CotizacionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_Cotizacion.TextBox4.Text = 1
        Reporte_Cotizacion.Show()
    End Sub

    Private Sub MenuStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip.ItemClicked

    End Sub

    Private Sub REGISTRODOCUMENTOToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Tesoreria.Show()
    End Sub

    Private Sub COTIZAIONToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_Cotizacion.TextBox4.Text = 2
        Reporte_Cotizacion.Show()
    End Sub

    Private Sub AnalisisCotizacionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Analisis_Cotizaciones.Show()
    End Sub

    Private Sub AprobarCotizacionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Aprobar_Cotizacion.TextBox5.Text = 1
        Aprobar_Cotizacion.Show()
    End Sub

    Private Sub EstadoCotizacionesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Imprimir_Teñido.Show()
    End Sub

    Private Sub CotizacionesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Teñido.Show()
    End Sub

    Private Sub CancelacionDocumentoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Cancelacion_documento.Show()
    End Sub

    Private Sub TesoreriaToolStripMenuItem_Click_1(sender As Object, e As EventArgs)
        Reporte_Tesoreria.Show()
    End Sub

    Private Sub ReporteVentasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Ventas.Label33.Text = Label5.Text
        Ventas.Show()
    End Sub

    Private Sub RegistroVentaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Opcion_venta.Show()
    End Sub

    Private Sub CancelarFacturasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Cancelar_Factura.Show()
    End Sub

    Private Sub VENDEDORToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VENDEDORToolStripMenuItem.Click

    End Sub

    Private Sub REPORTESGERENCIAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles REPORTESGERENCIAToolStripMenuItem.Click

    End Sub

    Private Sub MorosidadClienteToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Morosidad_Clientes.Show()
    End Sub

    Private Sub ReporteCobranzasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_Cobranzas.Show()
    End Sub

    Private Sub GuiaRemisionToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub BuscarYCancelarVentasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form3.TextBox4.Text = 1
        Form3.Show()
    End Sub

    Private Sub VentasSinFacturaToolStripMenuItem_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub EstadoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Estado_Cliente.Label5.Text = Label9.Text
        Estado_Cliente.Show()
    End Sub

    Private Sub VentasNToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form3.TextBox4.Text = 2
        Form3.Show()
    End Sub

    Private Sub ORDENCOSTURAToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Orden_Costura.Show()
    End Sub

    Private Sub ANALISISCONFECCIONToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Confeccion.Show()
    End Sub

    Private Sub PRODUCCIONDIARIACONFECCIONToolStripMenuItem_Click(sender As Object, e As EventArgs)
        produc_diaria_confec.Show()
    End Sub

    Private Sub ANALISISDECONFECCIONToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Analisis_Confeccion.Show()
    End Sub

    Private Sub ReporteVentasNToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_ventasn.Show()
    End Sub

    Private Sub RegistroVentaToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Opcion_venta.Show()
    End Sub

    Private Sub CancelarVentasSFacturaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form3.TextBox4.Text = 1
        Form3.TextBox9.Text = Label5.Text
        Form3.Label9.Text = Label9.Text
        Form3.Show()
    End Sub

    Private Sub ReporteVentasNToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Reporte_ventasn.Show()
    End Sub

    Private Sub AsignacionDeCosturaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_Costura.Show()
    End Sub

    Private Sub ComisionesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Comisiones_Vendedores.Label8.Text = Label5.Text
        Comisiones_Vendedores.Show()
    End Sub

    Private Sub ClientesToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        'Form_Clientes.TextBox17.Text = Label3.Text
        Form_Clientes.Show()
    End Sub

    Private Sub RegistroFacturaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Almacen.Label1.Text = Label5.Text
        Almacen.Label2.Text = Label9.Text
        Almacen.Show()
    End Sub

    Private Sub ReporteComisionesToolStripMenuItem_Click(sender As Object, e As EventArgs)

        Comisiones_Vendedores.Label8.Text = Label5.Text
        Comisiones_Vendedores.Label12.Text = Label9.Text
        Comisiones_Vendedores.Label13.Text = Label7.Text
        Comisiones_Vendedores.Show()
    End Sub

    Private Sub TesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TesToolStripMenuItem.Click

    End Sub

    Private Sub ToolStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip.ItemClicked

    End Sub

    Private Sub NotaIngresoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        NotaIngreso.TextBox7.Text = Label3.Text
        NotaIngreso.Label11.Text = Label5.Text
        NotaIngreso.Show()
    End Sub

    Private Sub RegistroVentaLibreToolStripMenuItem_Click(sender As Object, e As EventArgs)
        If Label9.Text = "01" Then
            Opcion_venta.Label1.Text = Label5.Text
            Opcion_venta.Label2.Text = Label9.Text
            Opcion_venta.Show()
        Else
            VENTA_GRAUS.Label1.Text = Label5.Text
            VENTA_GRAUS.Label2.Text = Label9.Text
            VENTA_GRAUS.Show()
        End If

    End Sub

    Private Sub EnlazarGuiaFacturaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Guia_Factura.Label5.Text = Label9.Text
        Guia_Factura.Show()
    End Sub

    Private Sub ReporteVentasNToolStripMenuItem_Click_1(sender As Object, e As EventArgs)
        Reporte_conta.Show()
    End Sub

    Private Sub ReporteVentasNToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        Reporte_conta.Label2.Text = Label5.Text
        Reporte_conta.Show()
    End Sub

    Private Sub CalidadPartidasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form4.Show()
    End Sub

    Private Sub ComisionesGeneralToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Rpt_Comision.Label3.Text = Label5.Text
        Rpt_Comision.Show()
    End Sub

    Private Sub StockFisicoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StockFisicoToolStripMenuItem.Click
        Stock_Fisico.Label5.Text = Label5.Text
        Stock_Fisico.Label6.Text = Label9.Text
        Stock_Fisico.Show()
    End Sub

    Private Sub ReporteToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Rpt_Comision.Label3.Text = Label5.Text
        Rpt_Comision.Label4.Text = Label9.Text
        Rpt_Comision.Show()
    End Sub

    Private Sub VentasMensualesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Ventas.Label33.Text = Label5.Text
        Ventas.Label36.Text = Label9.Text
        Ventas.Label35.Text = Label7.Text
        Ventas.Show()
    End Sub

    Private Sub ReporteVentasLToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_conta.Label2.Text = Label5.Text
        Reporte_conta.Show()
    End Sub

    Private Sub StockFisicoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles StockFisicoToolStripMenuItem1.Click
        Stock_fisico_comerci.Label5.Text = Label5.Text
        Stock_fisico_comerci.Label6.Text = Label9.Text
        Stock_fisico_comerci.Show()
    End Sub

    Private Sub ReporteClientesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteClientesToolStripMenuItem.Click
        Reporte_Cliente.ComboBox1.Text = Label3.Text
        Reporte_Cliente.Label4.Text = Label4.Text
        Reporte_Cliente.Show()
    End Sub

    Private Sub CrToolStripTextBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Tipo_Venta.Label1.Text = Label4.Text
        Tipo_Venta.Label2.Text = Label3.Text
        Tipo_Venta.Label3.Text = Label9.Text
        Tipo_Venta.Show()
    End Sub

    Private Sub AprobarNotaPedidoToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PedidosComercialToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PedidosComercialToolStripMenuItem.Click
        validar_Almacen.TextBox3.Text = 2
        validar_Almacen.Label6.Text = Label5.Text
        validar_Almacen.Label7.Text = Label9.Text
        validar_Almacen.Label8.Text = Label7.Text
        validar_Almacen.Label9.Text = Label3.Text
        validar_Almacen.Show()
    End Sub

    Private Sub SEguimientoPedidoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SEguimientoPedidoToolStripMenuItem.Click
        Detalle_Pedidos.TextBox4.Text = Label3.Text
        Detalle_Pedidos.Label3.Text = Label9.Text
        Detalle_Pedidos.Label4.Text = Label5.Text
        Detalle_Pedidos.Show()
    End Sub

    Private Sub ValidarDespachoPedidosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        validar_Despacho.Show()
    End Sub

    Private Sub ComisionesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ComisionesToolStripMenuItem1.Click
        Comisiones_ven.TextBox3.Text = Label3.Text
        Comisiones_ven.Label7.Text = Label5.Text
        Comisiones_ven.Label8.Text = Label9.Text
        Comisiones_ven.Label13.Text = Label7.Text
        Comisiones_ven.Label15.Text = 0
        Comisiones_ven.Show()
    End Sub

    Private Sub AnalisiVentasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Error_ventas.Label4.Text = Label5.Text
        Error_ventas.Label5.Text = Label9.Text
        Error_ventas.Show()
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        OD.TextBox15.Text = Label3.Text
        OD.TextBox7.Text = Label7.Text
        OD.Label23.Text = Label9.Text
        OD.Show()
    End Sub

    Private Sub GestionActividadesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        GESTION_ACTIVIDADES.TextBox3.Text = Label3.Text
        GESTION_ACTIVIDADES.Show()
    End Sub

    Private Sub SeguimientoVendedoresToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub SeguimientoActividadesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Seguimiento_Actividades.ComboBox2.Items.Clear()
        Seguimiento_Actividades.ComboBox2.Items.Add(Label3.Text)
        Seguimiento_Actividades.Show()
    End Sub

    Private Sub ReporteTotalCobranzasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_Total_Cobranzas.Show()
    End Sub

    Private Sub EstadoCuentaClienteToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Estado_Cliente.Label5.Text = Label9.Text
        Estado_Cliente.Show()
    End Sub

    Private Sub AnalisisVentaAnualXProductoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Analisis_Productos.Show()
    End Sub

    Private Sub ReporteClientesXVendedorToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Vendedor_Cliente.Show()
    End Sub

    Private Sub ReporteVentasToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Rpt_ventas.Show()
    End Sub

    Private Sub ProgramaRamaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        rama.Show()
    End Sub

    Private Sub AcabadoDeTelaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        PASADO_RAMA.Show()
    End Sub

    Private Sub SEguimientoRamaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_rama.Show()
    End Sub

    Private Sub PackingTelaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PackingTelaToolStripMenuItem1.Click
        Ingresar_Packing.Label7.Text = "LIMA"
        Ingresar_Packing.Show()

    End Sub

    Private Sub NotaSalidaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NotaSalidaToolStripMenuItem.Click
        Nsalida.TextBox7.Text = Label3.Text
        Nsalida.Label10.Text = Label5.Text
        Nsalida.Label13.Text = Label9.Text
        Nsalida.Show()
    End Sub

    Private Sub NotaIngresoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NotaIngresoToolStripMenuItem1.Click
        NotaIngreso.TextBox7.Text = Label3.Text
        NotaIngreso.Label11.Text = Label5.Text
        NotaIngreso.Label14.Text = Label9.Text
        NotaIngreso.Show()
    End Sub

    Private Sub GuiaRemisionToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles GuiaRemisionToolStripMenuItem1.Click

        Almacen_guia_nuevo.Label1.Text = Label5.Text
        Almacen_guia_nuevo.Label2.Text = Label9.Text
        Almacen_guia_nuevo.Label3.Text = Label3.Text
        Almacen_guia_nuevo.Label4.Text = "1"
        Almacen_guia_nuevo.Show()


    End Sub

    Private Sub PackingHiloToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PackingHiloToolStripMenuItem1.Click
        Packing_hilo.Show()
    End Sub

    Private Sub NotaIngresoToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles NotaIngresoToolStripMenuItem.Click
        NiaHilo.TextBox7.Text = Label3.Text
        NiaHilo.Label11.Text = Label5.Text
        NiaHilo.Label13.Text = Label9.Text
        NiaHilo.TextBox14.Text = 2
        NiaHilo.TextBox16.Text = 2
        NiaHilo.Show()
    End Sub

    Private Sub NotaSalidaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NotaSalidaToolStripMenuItem1.Click
        NsaHilo.TextBox7.Text = Label3.Text
        NsaHilo.Label10.Text = Label5.Text
        NsaHilo.Label13.Text = Label9.Text
        NsaHilo.TextBox12.Text = 2
        NsaHilo.TextBox18.Text = 2

        NsaHilo.Show()
    End Sub

    Private Sub GuiaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GuiaToolStripMenuItem.Click

        'Almacen_guia_chil.TextBox1.Text = Label3.Text
        'Almacen_guia_chil.Label1.Text = Label5.Text
        'Almacen_guia_chil.Label2.Text = Label9.Text

        'Almacen_guia_chil.Show()

        Almacen_guia_nuevo.Label1.Text = Label5.Text
        Almacen_guia_nuevo.Label2.Text = Label9.Text
        Almacen_guia_nuevo.Label3.Text = Label3.Text
        Almacen_guia_nuevo.Label4.Text = "2"
        Almacen_guia_nuevo.Show()
    End Sub

    Private Sub AnalisisVentasToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ProgramaRamaToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        rama.Show()
    End Sub

    Private Sub AcabadoRamaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        PASADO_RAMA.Show()
    End Sub

    Private Sub SeguimientoRamaToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Reporte_rama.Show()
    End Sub

    Private Sub HojaTeñidoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Teñido.Show()
    End Sub

    Private Sub GraficaTeñidoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Imprimir_Teñido.Show()
    End Sub

    Private Sub VentasAcumuladasToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ReporteVentasToolStripMenuItem2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComisionesVendedorToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Comisiones_Vendedores.Label8.Text = Label5.Text
        Comisiones_Vendedores.Label12.Text = Label9.Text
        Comisiones_Vendedores.Label13.Text = Label7.Text
        Comisiones_Vendedores.Show()
    End Sub

    Private Sub ComisionGenrealToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Rpt_Comision.Label3.Text = Label5.Text
        Rpt_Comision.Label4.Text = Label9.Text
        Rpt_Comision.Show()
    End Sub

    Private Sub VentasLibresToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Form3.TextBox4.Text = 2

        Form3.TextBox9.Text = Label5.Text
        Form3.Label9.Text = Label9.Text
        Form3.Show()
    End Sub

    Private Sub GestionCobranzasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_Cobranzas.Label17.Text = Label5.Text
        Reporte_Cobranzas.Label18.Text = Label9.Text
        Reporte_Cobranzas.Show()
    End Sub

    Private Sub SeguimientoVendedoresToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Seguimiento_Actividades.Show()
    End Sub

    Private Sub ReporteClienteXVendedorToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Vendedor_Cliente.Show()
    End Sub

    Private Sub AnalisisVentaAnualXProductoToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Analisis_Productos.Label5.Text = Label9.Text
        Analisis_Productos.Show()
    End Sub

    Private Sub ProgramacioOpToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ProgramacionPedidosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Programa_Tejeduria.Label3.Text = 2
        Programa_Tejeduria.Show()
    End Sub

    Private Sub TejeduriaToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub VentasPorFechaToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub AbridoraToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Abridora.Show()
    End Sub

    Private Sub UrdidoraToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Urdido.Show()
    End Sub

    Private Sub TeñidoCuerdasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        TEÑIDO_DE_CUERDAS.Show()
    End Sub

    Private Sub ProgramaTeñidoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Orden_Trabajo_Indigo.Show()
    End Sub

    Private Sub NotaSalidaToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles NotaSalidaToolStripMenuItem2.Click
        NsaHilo.TextBox7.Text = Label3.Text
        NsaHilo.Label10.Text = Label5.Text
        NsaHilo.Label13.Text = Label9.Text
        NsaHilo.TextBox18.Text = 3
        NsaHilo.TextBox12.Text = 1
        NsaHilo.Show()
    End Sub

    Private Sub ProcesoServicioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProcesoServicioToolStripMenuItem.Click

    End Sub

    Private Sub NotaIngresoToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles NotaIngresoToolStripMenuItem2.Click
        NiaHilo.TextBox7.Text = Label3.Text
        NiaHilo.Label11.Text = Label5.Text
        NiaHilo.Label13.Text = Label9.Text
        NiaHilo.TextBox14.Text = 1
        NiaHilo.TextBox16.Text = 3
        NiaHilo.Show()
    End Sub

    Private Sub PedidosComercialToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        validar_Almacen.TextBox3.Text = 1
        validar_Almacen.Label6.Text = Label5.Text
        validar_Almacen.Label7.Text = Label9.Text

        validar_Almacen.Show()
    End Sub

    Private Sub GuiaRemisionToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles GuiaRemisionToolStripMenuItem.Click
        'Almacen_guia_chil.TextBox1.Text = Label3.Text
        'Almacen_guia_chil.Label1.Text = Label5.Text
        'Almacen_guia_chil.Label2.Text = Label9.Text
        'Almacen_guia_chil.Show()

        Almacen_guia_nuevo.Label1.Text = Label5.Text
        Almacen_guia_nuevo.Label2.Text = Label9.Text
        Almacen_guia_nuevo.Label3.Text = Label3.Text
        Almacen_guia_nuevo.Label4.Text = "3"
        Almacen_guia_nuevo.Show()
    End Sub

    Private Sub NotaIngresoToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles NotaIngresoToolStripMenuItem3.Click
        NiaHilo.TextBox7.Text = Label3.Text
        NiaHilo.Label11.Text = Label5.Text
        NiaHilo.Label13.Text = Label9.Text
        NiaHilo.TextBox14.Text = 3
        NiaHilo.TextBox16.Text = 4
        NiaHilo.Show()
    End Sub

    Private Sub NotaSalidaToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles NotaSalidaToolStripMenuItem3.Click
        NsaHilo.TextBox7.Text = Label3.Text
        NsaHilo.Label10.Text = Label5.Text
        NsaHilo.Label13.Text = Label9.Text
        NsaHilo.TextBox12.Text = 3
        NsaHilo.TextBox18.Text = 4
        NsaHilo.Show()
    End Sub

    Private Sub ConeraToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Conera.Show()
    End Sub

    Private Sub ReporteProformasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_Proforma.TextBox2.Text = Label5.Text
        Reporte_Proforma.Label4.Text = Label9.Text
        Reporte_Proforma.Show()
    End Sub

    Private Sub ReporteDespachosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Reporte_despacho.TextBox1.Text = Label5.Text
        Reporte_despacho.Label5.Text = Label9.Text
        Reporte_despacho.Show()
    End Sub

    Private Sub LibrosElectronicosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        PDT_LIBROS.TextBox1.Text = Label5.Text
        PDT_LIBROS.Show()
    End Sub

    Private Sub DespachosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DespachosToolStripMenuItem.Click
        Reporte_despacho.TextBox1.Text = Label5.Text
        Reporte_despacho.Show()
    End Sub

    Private Sub ParteIngresoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ParteIngresoToolStripMenuItem.Click
        'Parte_Ingreso_Salida.Text = "PARTE INGRESO"
        'Parte_Ingreso_Salida.Show()
        Dim kl1 As New Reporte_Ingreso_Salida
        Reporte_Ingreso_Salida.TextBox4.Text = Label5.Text
        Reporte_Ingreso_Salida.TextBox5.Text = 1
        Reporte_Ingreso_Salida.Text = "REPORTE PARTE INGRESO"
        Reporte_Ingreso_Salida.Label4.Text = Label9.Text
        Reporte_Ingreso_Salida.Show()
    End Sub

    Private Sub ParteSalidaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ParteSalidaToolStripMenuItem.Click
        Dim kl As New Reporte_Ingreso_Salida
        Reporte_Ingreso_Salida.TextBox4.Text = Label5.Text
        Reporte_Ingreso_Salida.TextBox5.Text = 2
        Reporte_Ingreso_Salida.Text = "REPORTE PARTE SALIDA"
        Reporte_Ingreso_Salida.Label4.Text = Label9.Text
        Reporte_Ingreso_Salida.Show()
    End Sub

    Private Sub EnlazarGuiaFacturaToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Guia_Factura.Label5.Text = Label9.Text
        Guia_Factura.Show()
    End Sub

    Private Sub ActualizarRubroTipoVentaTiendaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Actualizar_tienda.TextBox1.Text = Label5.Text
        Actualizar_tienda.Show()
    End Sub

    Private Sub AvancePedidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AvancePedidosToolStripMenuItem.Click
        Programa_Tejeduria.Label3.Text = 3
        Programa_Tejeduria.Show()
    End Sub

    Private Sub GuiaRemisionToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles GuiaRemisionToolStripMenuItem2.Click

        Almacen_guia_nuevo.Label1.Text = Label5.Text
        Almacen_guia_nuevo.Label2.Text = Label9.Text
        Almacen_guia_nuevo.Label3.Text = Label3.Text
        Almacen_guia_nuevo.Label4.Text = "4"
        Almacen_guia_nuevo.Show()
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Registro_Celulares.Show()
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub FinalizarDespachoToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub AprobacionPedidosToolStripMenuItem_Click_1(sender As Object, e As EventArgs)
        Aprobar_notapedido.Show()
    End Sub

    Private Sub TablaDeComisionesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Tabla_Comisiones.TextBox1.Text = Label5.Text
        Tabla_Comisiones.TextBox2.Text = Label7.Text
        Tabla_Comisiones.Label4.Text = Label9.Text

        Tabla_Comisiones.Show()
    End Sub

    Private Sub ActualizarFechaToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub AsignarRubroACOdigoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ADICIONAR_RUBRO_A_CODIGO.Label1.Text = Label9.Text
        ADICIONAR_RUBRO_A_CODIGO.Label2.Text = Label5.Text
        ADICIONAR_RUBRO_A_CODIGO.Show()
    End Sub

    Private Sub PedidosComercialToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        validar_Almacen.TextBox3.Text = 1
        validar_Almacen.Label6.Text = Label5.Text
        validar_Almacen.Label7.Text = Label9.Text

        validar_Almacen.Show()
    End Sub

    Private Sub MarcaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Importar_Marcaciones.Show()
    End Sub

    Private Sub ConFacturaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConFacturaToolStripMenuItem.Click

        almacen_nuevo.Label3.Text = Label5.Text
        almacen_nuevo.Label2.Text = Label9.Text
        almacen_nuevo.Label4.Text = "1"
        almacen_nuevo.Show()
    End Sub

    Private Sub VentaLibToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VentaLibToolStripMenuItem.Click


        Registro_Ventas.Label26.Text = Label9.Text
        Registro_Ventas.Label17.Text = Label5.Text
        Registro_Ventas.Show()

    End Sub

    Private Sub CancelarVentasLibresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancelarVentasLibresToolStripMenuItem.Click
        Form3.TextBox4.Text = 1
        Form3.TextBox9.Text = Label5.Text
        Form3.Label9.Text = Label9.Text
        Form3.Show()
    End Sub

    Private Sub ReporteVentasLibresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteVentasLibresToolStripMenuItem.Click
        Reporte_ventasn.Show()
    End Sub

    Private Sub EstadoCuentaClienteToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles EstadoCuentaClienteToolStripMenuItem1.Click
        Estado_Cliente.Label5.Text = Label9.Text
        Estado_Cliente.Show()
    End Sub

    Private Sub ReporteTotalCobranzasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ReporteTotalCobranzasToolStripMenuItem1.Click
        Reporte_Total_Cobranzas.Show()
    End Sub

    Private Sub PedidosComercialToolStripMenuItem4_Click(sender As Object, e As EventArgs)
        validar_Almacen.TextBox3.Text = 1
        validar_Almacen.Label6.Text = Label5.Text
        validar_Almacen.Label7.Text = Label9.Text

        validar_Almacen.Show()
    End Sub

    Private Sub ReporteDeVentasLibresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteDeVentasLibresToolStripMenuItem.Click
        Reporte_Proforma.TextBox2.Text = Label5.Text
        Reporte_Proforma.Label4.Text = Label9.Text
        Reporte_Proforma.Show()
    End Sub

    Private Sub ReporteDEspachosToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ReporteDEspachosToolStripMenuItem1.Click
        Reporte_despacho.TextBox1.Text = Label5.Text
        Reporte_despacho.Label5.Text = Label9.Text
        Reporte_despacho.Show()
    End Sub

    Private Sub AprobacionDePedidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AprobacionDePedidosToolStripMenuItem.Click
        Aprobar_notapedido.Label3.Text = Label9.Text
        Aprobar_notapedido.Label5.Text = Label7.Text
        Aprobar_notapedido.Label6.Text = Label3.Text
        Aprobar_notapedido.Show()
    End Sub

    Private Sub LibrosELectronicosToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles LibrosELectronicosToolStripMenuItem1.Click
        PDT_LIBROS.TextBox1.Text = Label5.Text
        PDT_LIBROS.Show()
    End Sub

    Private Sub EToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EToolStripMenuItem.Click
        Guia_Factura.Label5.Text = Label9.Text
        Guia_Factura.Show()
    End Sub

    Private Sub ActualizarRubroTipoVentaTiendaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ActualizarRubroTipoVentaTiendaToolStripMenuItem1.Click
        Actualizar_tienda.TextBox1.Text = Label5.Text
        Actualizar_tienda.Show()
    End Sub

    Private Sub RptVentasLibresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RptVentasLibresToolStripMenuItem.Click
        Reporte_conta.Show()

    End Sub

    Private Sub PedidosComercialToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles PedidosComercialToolStripMenuItem3.Click
        validar_Almacen.TextBox3.Text = 1
        validar_Almacen.Label6.Text = Label5.Text
        validar_Almacen.Label7.Text = Label9.Text

        validar_Almacen.Show()
    End Sub

    Private Sub PorVendedorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PorVendedorToolStripMenuItem.Click
        'Comisiones_Vendedores.Label8.Text = Label5.Text
        'Comisiones_Vendedores.Label12.Text = Label9.Text
        'Comisiones_Vendedores.Label13.Text = Label7.Text
        'Comisiones_Vendedores.Show()
        Comisiones_ven.TextBox3.Text = Label3.Text
        Comisiones_ven.Label7.Text = Label5.Text
        Comisiones_ven.Label8.Text = Label9.Text
        Comisiones_ven.Label13.Text = Label7.Text
        Comisiones_ven.Label15.Text = 1
        Comisiones_ven.Show()
    End Sub
    Private Sub TodosLosVendedoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TodosLosVendedoresToolStripMenuItem.Click
        Rpt_Comision.Label3.Text = Label5.Text
        Rpt_Comision.Label4.Text = Label9.Text
        Rpt_Comision.Show()
    End Sub

    Private Sub VentasMensualesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles VentasMensualesToolStripMenuItem1.Click
        Ventas.Label33.Text = Label5.Text
        Ventas.Label36.Text = Label9.Text
        Ventas.Label35.Text = Label7.Text
        Ventas.Show()
    End Sub

    Private Sub VentasLibresToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles VentasLibresToolStripMenuItem1.Click
        Reporte_conta.Label2.Text = Label5.Text
        Reporte_conta.Show()
    End Sub

    Private Sub AprobacionPedidosToolStripMenuItem_Click_2(sender As Object, e As EventArgs) Handles AprobacionPedidosToolStripMenuItem.Click



    End Sub

    Private Sub AprobacionPedidosToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles AprobacionPedidosToolStripMenuItem1.Click
        If Label3.Text = "DBRAVO" Or Label3.Text = "LGONZALEZ" Then
            Aprobacion_comercial.Label4.Text = Label9.Text
            Aprobacion_comercial.Label5.Text = Label7.Text
            Aprobacion_comercial.Label7.Text = Label3.Text
            Aprobacion_comercial.Show()
        Else
            MsgBox("NO TIENE PERMISO PARA INGRESAR A ESTE FORMULARIO")
        End If
    End Sub

    Private Sub ReporteVentasToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ReporteVentasToolStripMenuItem.Click
        Ventas.Label33.Text = Label5.Text
        Ventas.Label36.Text = Label9.Text
        Ventas.Label35.Text = Label7.Text
        Ventas.Show()
    End Sub

    Private Sub RptComisionesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RptComisionesToolStripMenuItem.Click
        Comisiones_Vendedores.Label8.Text = Label5.Text
        Comisiones_Vendedores.Label12.Text = Label9.Text
        Comisiones_Vendedores.Label13.Text = Label7.Text
        Comisiones_Vendedores.Show()
    End Sub

    Private Sub ReporteAnalisisVentasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteAnalisisVentasToolStripMenuItem.Click
        Rpt_ventas.Label2.Text = Label9.Text
        Rpt_ventas.Label4.Text = Label7.Text
        Rpt_ventas.Label7.Text = Label5.Text
        Rpt_ventas.Show()
    End Sub

    Private Sub ConeraToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ConeraToolStripMenuItem1.Click
        Conera.Show()
    End Sub

    Private Sub AbridoraToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AbridoraToolStripMenuItem1.Click
        Abridora.Show()
    End Sub

    Private Sub TeñidoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TeñidoToolStripMenuItem.Click
        TEÑIDO_DE_CUERDAS.Show()
    End Sub

    Private Sub UrdidoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UrdidoToolStripMenuItem.Click
        Urdido.Show()
    End Sub

    Private Sub ProgramaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProgramaToolStripMenuItem.Click
        Orden_Trabajo_Indigo.Show()
    End Sub

    Private Sub HojaTeñidoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HojaTeñidoToolStripMenuItem1.Click
        Teñido.Show()
    End Sub

    Private Sub TeñidoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TeñidoToolStripMenuItem1.Click
        Teñido.Show()
    End Sub

    Private Sub AsignarRubroToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AsignarRubroToolStripMenuItem.Click
        ADICIONAR_RUBRO_A_CODIGO.Label1.Text = Label9.Text
        ADICIONAR_RUBRO_A_CODIGO.Label2.Text = Label5.Text
        ADICIONAR_RUBRO_A_CODIGO.Show()
    End Sub

    Private Sub TablaComicionesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TablaComicionesToolStripMenuItem.Click
        Tabla_Comisiones.TextBox1.Text = Label5.Text
        Tabla_Comisiones.TextBox2.Text = Label7.Text
        Tabla_Comisiones.Label4.Text = Label9.Text

        Tabla_Comisiones.Show()
    End Sub

    Private Sub EstadoCuentaClienteToolStripMenuItem_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub SeguimientoVendedoresToolStripMenuItem_Click_1(sender As Object, e As EventArgs)
        Seguimiento_Actividades.Show()
    End Sub

    Private Sub ReporteClienteVendedorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteClienteVendedorToolStripMenuItem.Click
        Vendedor_Cliente.Show()
    End Sub

    Private Sub MuestraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MuestraToolStripMenuItem.Click
        Codificacion_Prodcutos.Label13.Text = Label9.Text
        Codificacion_Prodcutos.Label14.Text = 1
        Codificacion_Prodcutos.Show()
    End Sub

    Private Sub TejedoresToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TejedoresToolStripMenuItem1.Click
        TEJEDORES.Show()
    End Sub

    Private Sub DefectosTejeduriaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DefectosTejeduriaToolStripMenuItem1.Click
        DEFECTOS_AUDITORIA.Label6.Text = Label3.Text
        DEFECTOS_AUDITORIA.Show()
    End Sub

    Private Sub AuditoresToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AuditoresToolStripMenuItem1.Click
        AUDITORES.Show()
    End Sub

    Private Sub MaquinasTejedorasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles MaquinasTejedorasToolStripMenuItem1.Click
        MAQUINAS_TEJEDORAS.Show()
    End Sub

    Private Sub PesadoRollosToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PesadoRollosToolStripMenuItem1.Click
        pesaod_rollo.Label28.Text = Label9.Text
        pesaod_rollo.Show()
    End Sub

    Private Sub CalidadTelaCrudaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalidadTelaCrudaToolStripMenuItem.Click
        Calidad_tela_cruda.Label28.Text = Label9.Text
        Calidad_tela_cruda.Show()
    End Sub

    Private Sub RequerimietoPorPOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RequerimietoPorPOToolStripMenuItem.Click

    End Sub

    Private Sub CalidadPartidasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CalidadPartidasToolStripMenuItem1.Click
        Form4.Show()
    End Sub

    Private Sub SeguimientoRamaToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles SeguimientoRamaToolStripMenuItem.Click
        Reporte_rama.Show()
    End Sub

    Private Sub AcabadoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AcabadoToolStripMenuItem.Click
        PASADO_RAMA.Show()
    End Sub

    Private Sub ProgramaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ProgramaToolStripMenuItem1.Click
        rama.Label5.Text = Label3.Text
        rama.Label8.Text = Label9.Text
        rama.Show()
    End Sub

    Private Sub ProduccionToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ProduccionToolStripMenuItem1.Click
        Codificacion_Prodcutos.Label13.Text = Label9.Text
        Codificacion_Prodcutos.Label14.Text = 2
        Codificacion_Prodcutos.Show()
    End Sub

    Private Sub CreacionCodigosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreacionCodigosToolStripMenuItem.Click

    End Sub

    Private Sub ValidarPartidaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ValidarPartidaToolStripMenuItem.Click
        Validar_partida.Label10.Text = Label9.Text
        Validar_partida.Show()
    End Sub

    Private Sub ProduccionDiariaTejeduriaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProduccionDiariaTejeduriaToolStripMenuItem.Click
        ProddiariaTej.Show()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub CalidadCrudoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        REPORTE_CALIDADCC.Show()
    End Sub

    Private Sub MatrizAviosToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub NotaIngresoToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles NotaIngresoToolStripMenuItem4.Click
        NiaHilo.TextBox7.Text = Label3.Text
        NiaHilo.Label11.Text = Label5.Text
        NiaHilo.Label13.Text = Label9.Text
        NiaHilo.TextBox16.Text = 5
        NiaHilo.Show()
    End Sub

    Private Sub NotaSalidaToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles NotaSalidaToolStripMenuItem4.Click
        NsaHilo.TextBox7.Text = Label3.Text
        NsaHilo.Label10.Text = Label5.Text
        NsaHilo.Label13.Text = Label9.Text
        NsaHilo.TextBox12.Text = 4
        NsaHilo.TextBox18.Text = 5
        NsaHilo.Show()
    End Sub

    Private Sub ProgramacionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Planemaiento_OP.Label2.Text = Label9.Text
        Planemaiento_OP.Show()
    End Sub

    Private Sub GuiaRemisionToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles GuiaRemisionToolStripMenuItem3.Click

        Guia_hilo.Label23.Text = "13"
        Guia_hilo.Label25.Text = "  ALMACEN AVIOS"
        Guia_hilo.TextBox19.Text = Label3.Text
        Guia_hilo.Label29.Text = Label5.Text
        Guia_hilo.TextBox1.Text = "T005"
        Guia_hilo.Label30.Text = Label9.Text
        Guia_hilo.Show()
    End Sub

    Private Sub ValidacionDespachoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ValidacionDespachoToolStripMenuItem.Click

    End Sub

    Private Sub ValidarDespachoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ValidarDespachoToolStripMenuItem.Click
        validar_Despacho.Label1.Text = Label9.Text
        validar_Despacho.Label2.Text = Label5.Text
        validar_Despacho.Show()
    End Sub

    Private Sub OrdenCorteToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ReporteRamaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteRamaToolStripMenuItem.Click
        Reporte_rama.Show()
    End Sub

    Private Sub ProgramacionToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ProgramacionToolStripMenuItem1.Click
        Despacho_Pedido.Label4.Text = Label7.Text
        Despacho_Pedido.Label5.Text = Label9.Text
        Despacho_Pedido.Show()
    End Sub

    Private Sub ProduccionDiariaCorteToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Prod_corteeee.Label4.Text = Label7.Text
        Prod_corteeee.Label5.Text = Label9.Text
        Prod_corteeee.Show()
    End Sub

    Private Sub HabilitarPartidaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HabilitarPartidaToolStripMenuItem.Click
        SegundoÑPasado_rama.Show()
    End Sub

    Private Sub GenerarParteIngresoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerarParteIngresoToolStripMenuItem.Click
        Generara_ingreso.Label2.Text = Label5.Text
        Generara_ingreso.Label3.Text = Label9.Text
        Generara_ingreso.Show()
    End Sub

    Private Sub PartidasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PartidasToolStripMenuItem.Click
        Calidad_tela__acabada.Label3.Text = Label5.Text
        Calidad_tela__acabada.Show()
    End Sub

    Private Sub ExportarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportarToolStripMenuItem.Click
        Rpt_ventas.Label2.Text = Label9.Text
        Rpt_ventas.Label4.Text = Label7.Text
        Rpt_ventas.Label7.Text = Label5.Text
        Rpt_ventas.Show()
    End Sub

    Private Sub RequerimientoPorPOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RequerimientoPorPOToolStripMenuItem.Click
        REQUERIMIENTO_PO.Label3.Text = Label9.Text
        REQUERIMIENTO_PO.Show()
    End Sub

    Private Sub EmisionOtPartidasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EmisionOtPartidasToolStripMenuItem.Click
        EMISION_PARTIDAS.Label2.Text = Label9.Text
        EMISION_PARTIDAS.Label3.Text = Label5.Text
        EMISION_PARTIDAS.Label4.Text = Label3.Text
        EMISION_PARTIDAS.Show()
    End Sub

    Private Sub VentasToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles VentasToolStripMenuItem3.Click
        Rpt_ventas.Label2.Text = Label9.Text
        Rpt_ventas.Label4.Text = Label7.Text
        Rpt_ventas.Label7.Text = Label5.Text
        Rpt_ventas.Show()
    End Sub

    Private Sub CalidadPartidasToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles CalidadPartidasToolStripMenuItem.Click
        Calidad_tela__acabada.Label3.Text = Label5.Text
        Calidad_tela__acabada.Show()
    End Sub

    Private Sub ActualizarPesoRolloToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PartidasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PartidasToolStripMenuItem1.Click
        CREAR_PARTIDAS.Label11.Text = Label5.Text
        CREAR_PARTIDAS.Label12.Text = Label9.Text
        CREAR_PARTIDAS.Show()
    End Sub

    Private Sub PackingTelaToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles PackingTelaToolStripMenuItem.Click
        buscar_Packing.Label1.Text = Label3.Text
        buscar_Packing.Label2.Text = Label5.Text
        buscar_Packing.Label4.Text = Label9.Text
        buscar_Packing.Show()
    End Sub
    Dim Rsr2 As SqlDataReader
    Private Sub ROLLOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ROLLOToolStripMenuItem.Click
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

            Impr_calidad.Show()
            Rsr2.Close()
        Else
            MsgBox("LA PARTIDA NO EXISTE")
        End If


    End Sub

    Private Sub ProgramaRamaToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ProgramaRamaToolStripMenuItem.Click
        Dim jh As New Rpt_Rama
        Dim clave As String


        clave = InputBox("Introduzca el Programa de Rama")
        Select Case clave.Length
            Case "1" : clave = "000000000" & "" & clave
            Case "2" : clave = "00000000" & "" & clave
            Case "3" : clave = "0000000" & "" & clave
            Case "4" : clave = "000000" & "" & clave
            Case "5" : clave = "00000" & "" & clave
            Case "6" : clave = "0000" & "" & clave
            Case "7" : clave = "000" & "" & clave
            Case "8" : clave = "00" & "" & clave
            Case "9" : clave = "0" & "" & clave
            Case "10" : clave = clave
        End Select
        jh.TextBox1.Text = clave
        jh.Show()
    End Sub

    Private Sub ActualizarPesoRollosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActualizarPesoRollosToolStripMenuItem.Click
        Actualizar_Fecha.Label2.Text = Label9.Text
        Actualizar_Fecha.Show()
    End Sub

    Private Sub TelaAacabadaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TelaAacabadaToolStripMenuItem.Click

    End Sub

    Private Sub CalidadPartidasToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles CalidadPartidasToolStripMenuItem2.Click
        Calidad_tela__acabada.Show()
    End Sub

    Private Sub ProgramacionDePedidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProgramacionDePedidosToolStripMenuItem.Click
        Programa_Tejeduria.Label3.Text = 1
        Programa_Tejeduria.Label4.Text = Label9.Text
        Programa_Tejeduria.Show()
    End Sub

    Private Sub PackingTelaToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        pACKING_COMERCIAL.Show()
    End Sub

    Private Sub RequerimientoPedidosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RequerimientoPedidosToolStripMenuItem.Click
        Requerimiento__comercial.Show()
    End Sub

    Private Sub StockFisicoToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles StockFisicoToolStripMenuItem2.Click
        Stock_Fisico.Label5.Text = Label5.Text
        Stock_Fisico.Label6.Text = Label9.Text
        Stock_Fisico.Show()
    End Sub

    Private Sub RequerimientoComercialToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RequerimientoComercialToolStripMenuItem.Click
        buscar_requerimitno.Show()
    End Sub

    Private Sub RequerimientoComercialToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RequerimientoComercialToolStripMenuItem1.Click
        buscar_requerimitno.Show()
    End Sub



    Private Sub PesadoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        jalar_peso.Show()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ParteSalidaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ParteSalidaToolStripMenuItem1.Click
        Rpt_ing_sal.Show()
    End Sub

    Private Sub TablaDeRubrosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TablaDeRubrosToolStripMenuItem.Click
        Rubros_Tesoreria.Show()
    End Sub

    Private Sub PesoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        jalar_peso.Show()
    End Sub

    Private Sub DataBancosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataBancosToolStripMenuItem.Click
        Data_Bancos.TextBox1.Text = Label7.Text
        Data_Bancos.Label7.Text = Label9.Text
        Data_Bancos.TextBox4.Text = Label5.Text
        Data_Bancos.Show()
    End Sub

    Private Sub MantenimientoVehiculosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MantenimientoVehiculosToolStripMenuItem.Click

    End Sub

    Private Sub TablaRubroToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TablaRubroToolStripMenuItem.Click
        rubros_manteminiento.Show()
    End Sub

    Private Sub RegistroIncidentesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegistroIncidentesToolStripMenuItem.Click
        Mantenimiento_carro.TextBox10.Text = Label5.Text
        Mantenimiento_carro.Show()
    End Sub

    Private Sub TablaVehiiculosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TablaVehiiculosToolStripMenuItem.Click
        vehuculo.Show()
    End Sub

    Private Sub FichaTecnicaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FichaTecnicaToolStripMenuItem.Click
        FICHA_ALMACE.Show()
    End Sub

    Private Sub MovimientoAlmacenesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MovimientoAlmacenesToolStripMenuItem.Click
        Rpt_ing_sal.Label7.Text = Label9.Text
        Rpt_ing_sal.Show()
    End Sub

    Private Sub PLANEAMIENTOOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PLANEAMIENTOOPToolStripMenuItem.Click
        Planemaiento_OP.Label2.Text = Label9.Text
        Planemaiento_OP.Show()
    End Sub

    Private Sub MatrizAviosToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles MatrizAviosToolStripMenuItem1.Click
        Explosion_avios_graus.Label8.Text = Label9.Text
        Explosion_avios_graus.Label10.Text = Label5.Text
        Explosion_avios_graus.Show()
    End Sub

    Private Sub OrdenCorteToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles OrdenCorteToolStripMenuItem1.Click
        corte.Label8.Text = Label9.Text
        corte.Show()
    End Sub

    Private Sub ProduccionDiariaCorteToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PlaneaminetoOpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PlaneaminetoOpToolStripMenuItem.Click
        Planemaiento_OP.Label2.Text = Label9.Text
        Planemaiento_OP.Show()
    End Sub

    'Private Sub ASIGNACIONTELAOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASIGNACIONTELAOPToolStripMenuItem.Click
    '    Asignacion_op_tela.Show()
    'End Sub

    Private Sub ASIGNACIONDEMOLDEOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASIGNACIONDEMOLDEOPToolStripMenuItem.Click
        Asignacion_Molde.Label9.Text = Label9.Text
        Asignacion_Molde.Label13.Text = Label3.Text
        Asignacion_Molde.Show()
    End Sub

    Private Sub FICHATECNICAToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FICHATECNICAToolStripMenuItem1.Click
        FICHA_ALMACE.Show()
    End Sub

    Private Sub PREPRODUCCIONToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PREPRODUCCIONToolStripMenuItem.Click
        Pre_produccion.TextBox1.Text = Label5.Text
        Pre_produccion.Show()
    End Sub

    Private Sub STATUSCORTEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles STATUSCORTEToolStripMenuItem.Click

    End Sub

    Private Sub CONSUMOOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CONSUMOOPToolStripMenuItem.Click
        FICHA_CONSUMO.Label8.Text = Label9.Text
        FICHA_CONSUMO.Show()
    End Sub

    Private Sub TablaVendedoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TablaVendedoresToolStripMenuItem.Click
        Tabla_Vendedores.Show()
    End Sub

    Private Sub ALMACENToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ALMACENToolStripMenuItem2.Click
        STATUS_CORTE.TextBox1.Text = Label5.Text
        STATUS_CORTE.Label2.Text = Label9.Text
        STATUS_CORTE.Show()
    End Sub

    Private Sub STATUSCONFECCIONToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles STATUSCONFECCIONToolStripMenuItem.Click

    End Sub

    Private Sub FechaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FechaToolStripMenuItem.Click
        Rpt_Flujo_Caja_Fecha.TextBox1.Text = Label7.Text
        Rpt_Flujo_Caja_Fecha.TextBox2.Text = Label5.Text
        Rpt_Flujo_Caja_Fecha.Label5.Text = Label9.Text
        Rpt_Flujo_Caja_Fecha.Show()
    End Sub

    Private Sub ReporteVentasMensualToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteVentasMensualToolStripMenuItem.Click
        Ventas.Label33.Text = Label5.Text
        Ventas.Label36.Text = Label9.Text
        Ventas.Label35.Text = Label7.Text
        Ventas.Show()
    End Sub

    Private Sub ComisionesToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ComisionesToolStripMenuItem.Click
        Comisiones_Vendedores.Label8.Text = Label5.Text
        Comisiones_Vendedores.Label12.Text = Label9.Text
        Comisiones_Vendedores.Label13.Text = Label7.Text
        Comisiones_Vendedores.Show()
    End Sub

    Private Sub AnalisisVentasToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles AnalisisVentasToolStripMenuItem.Click
        Rpt_ventas.Label2.Text = Label9.Text
        Rpt_ventas.Label4.Text = Label7.Text
        Rpt_ventas.Label7.Text = Label5.Text
        Rpt_ventas.Show()
    End Sub

    Private Sub EstadoCuentaClienteToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles EstadoCuentaClienteToolStripMenuItem2.Click
        Estado_Cliente.Label5.Text = Label9.Text
        Estado_Cliente.Show()
    End Sub

    Private Sub VentasXClienteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VentasXClienteToolStripMenuItem.Click
        Ventas_Cliente.Show()
    End Sub

    Private Sub MovimientosAlmacenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MovimientosAlmacenToolStripMenuItem.Click
        Rpt_ing_sal.Show()
    End Sub

    Private Sub ReporteDespachosToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ReporteDespachosToolStripMenuItem.Click
        Despacho_Pedido.Label4.Text = Label7.Text
        Despacho_Pedido.Label5.Text = Label9.Text
        Despacho_Pedido.Show()
    End Sub

    Private Sub PackingTelaToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles PackingTelaToolStripMenuItem3.Click
        buscar_Packing.Show()
    End Sub

    Private Sub StockFisicoToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles StockFisicoToolStripMenuItem3.Click
        Stock_Fisico.Label5.Text = Label5.Text
        Stock_Fisico.Label6.Text = Label9.Text
        Stock_Fisico.Show()
    End Sub

    Private Sub FechaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FechaToolStripMenuItem1.Click
        Rpt_Flujo_Caja_Fecha.TextBox1.Text = Label7.Text
        Rpt_Flujo_Caja_Fecha.TextBox2.Text = Label5.Text
        Rpt_Flujo_Caja_Fecha.Label5.Text = Label9.Text
        Rpt_Flujo_Caja_Fecha.Show()
    End Sub

    Private Sub RequerimientoComercialToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles RequerimientoComercialToolStripMenuItem2.Click
        buscar_requerimitno.Show()
    End Sub

    Private Sub ProgramacionPedidosToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ProgramacionPedidosToolStripMenuItem.Click
        Programa_Tejeduria.Label3.Text = 1
        Programa_Tejeduria.Label4.Text = Label9.Text
        Programa_Tejeduria.Show()
    End Sub

    Private Sub ProduccionTejeduriaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProduccionTejeduriaToolStripMenuItem.Click
        ProddiariaTej.Show()
    End Sub

    Private Sub ProgramaRamaToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles ProgramaRamaToolStripMenuItem1.Click
        rama.Show()
    End Sub

    Private Sub SeguimientoRamaToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles SeguimientoRamaToolStripMenuItem1.Click
        Reporte_rama.Show()
    End Sub

    Private Sub CalidadPartidasToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles CalidadPartidasToolStripMenuItem3.Click
        Calidad_tela__acabada.Label3.Text = Label9.Text
        Calidad_tela__acabada.Show()
    End Sub

    Private Sub NIAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem.Click
        Nia_Produccion.TextBox1.Text = "0400"
        Nia_Produccion.TextBox13.Text = "0400"
        Nia_Produccion.TextBox14.Text = "0421"
        Nia_Produccion.TextBox16.Text = "CONFECCION"
        Nia_Produccion.TextBox11.Text = "CONFECCION"
        Nia_Produccion.TextBox2.Text = "AREA CONFECCION"
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub NIAToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem1.Click
        Nia_Produccion.TextBox1.Text = "0600"
        Nia_Produccion.TextBox13.Text = "0600"
        Nia_Produccion.TextBox14.Text = "0601"
        Nia_Produccion.TextBox16.Text = "LAVANDERIA"
        Nia_Produccion.TextBox11.Text = "LAVANDERIA"
        Nia_Produccion.TextBox2.Text = "AREA LAVANDERIA"
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub NIAToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem2.Click
        Nia_Produccion.TextBox1.Text = "0500"
        Nia_Produccion.TextBox13.Text = "0500"
        Nia_Produccion.TextBox14.Text = "0522"
        Nia_Produccion.TextBox16.Text = "ACABADO"
        Nia_Produccion.TextBox11.Text = "ACABADO"
        Nia_Produccion.TextBox2.Text = "AREA ACABADO"
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub NSAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0400"
        nsa_produccion.TextBox13.Text = "0400"
        nsa_produccion.TextBox14.Text = "0421"
        nsa_produccion.TextBox16.Text = "CONFECCION"
        nsa_produccion.TextBox11.Text = "CONFECCION"
        nsa_produccion.TextBox2.Text = "AREA CONFECCION"
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.ShowDialog()
    End Sub

    Private Sub NSAToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem1.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0600"
        nsa_produccion.TextBox13.Text = "0600"
        nsa_produccion.TextBox14.Text = "0601"
        nsa_produccion.TextBox16.Text = "LAVANDERIA"
        nsa_produccion.TextBox11.Text = "LAVANDERIA"
        nsa_produccion.TextBox2.Text = "AREA LAVANDERIA"
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Show()
    End Sub

    Private Sub NSAToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem2.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0500"
        nsa_produccion.TextBox13.Text = "0500"
        nsa_produccion.TextBox14.Text = "0522"
        nsa_produccion.TextBox16.Text = "ACABADO"
        nsa_produccion.TextBox11.Text = "ACABADO"
        nsa_produccion.TextBox2.Text = "AREA ACABADO"
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Show()
    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem4.Click
        guia_talleres.Label23.Text = "0400"
        guia_talleres.Label25.Text = "CONFECCION"
        guia_talleres.TextBox24.Text = "0421"
        guia_talleres.TextBox23.Text = "SERVICIO CONFECCION"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()
    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem5.Click
        guia_talleres.Label23.Text = "0600"
        guia_talleres.Label25.Text = "LAVANDERIA"
        guia_talleres.TextBox24.Text = "0601"
        guia_talleres.TextBox23.Text = "SERVICIO LAVANDERIA"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()
    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem6.Click
        guia_talleres.Label23.Text = "0500"
        guia_talleres.Label25.Text = "ACABADO"
        guia_talleres.TextBox24.Text = "0522"
        guia_talleres.TextBox23.Text = "SERVICIO ACABADO"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()
    End Sub

    Private Sub NIAToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem3.Click
        Nia_Produccion.TextBox1.Text = "0300"
        Nia_Produccion.TextBox13.Text = "0300"
        Nia_Produccion.TextBox14.Text = "0301"
        Nia_Produccion.TextBox16.Text = "CORTE"
        Nia_Produccion.TextBox11.Text = "CORTE"
        Nia_Produccion.TextBox2.Text = "AREA CORTE"
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub ReporteVentasToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles ReporteVentasToolStripMenuItem1.Click
        Rpt_Ventas_ven.TextBox1.Text = Label5.Text
        Rpt_Ventas_ven.TextBox2.Text = Label3.Text
        Rpt_Ventas_ven.TextBox3.Text = Label7.Text
        Rpt_Ventas_ven.Label5.Text = Label9.Text
        abrir()
        Dim sql102 As String = "SELECT codigo_ven  FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + Label3.Text + "' and ccia_ven ='01'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            Rpt_Ventas_ven.Label6.Text = Rsr2(0)
        End If
        Rsr2.Close()
        Rpt_Ventas_ven.Show()
    End Sub

    Private Sub NSAToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem3.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0300"
        nsa_produccion.TextBox13.Text = "0300"
        nsa_produccion.TextBox14.Text = "0301"
        nsa_produccion.TextBox16.Text = "CORTE"
        nsa_produccion.TextBox11.Text = "SERVICIO CORTE"
        nsa_produccion.TextBox2.Text = "AREA CORTE"
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.ShowDialog()
    End Sub

    Private Sub RegistroClientesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegistroClientesToolStripMenuItem.Click
        'Form_Clientes.TextBox17.Text = Label3.Text
        Form_Clientes.Show()
    End Sub

    Private Sub ReporteProveedoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteProveedoresToolStripMenuItem.Click

    End Sub

    Private Sub DataProveedoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataProveedoresToolStripMenuItem.Click
        abrir()
        Dim Rsr1991 As SqlDataReader
        Dim cco As String
        Dim sql1011 As String = "select COUNT(fuen_3a) from  custom_vianny.dbo.cac3p00 where ncom_3a in (01000001,01000002) and  ccia_3A ='" + Label9.Text + "' and cperiodo_3A ='" + Label5.Text + "' AND ccom_3A ='500'  and fuen_3a= '1'"
        Dim cmd1011 As New SqlCommand(sql1011, conx)
        Rsr1991 = cmd1011.ExecuteReader()
        Rsr1991.Read()
        cco = Rsr1991(0)
        Rsr1991.Close()

        If cco > 0 Then
            proveedores.TextBox1.Text = Label5.Text
            proveedores.TextBox2.Text = Label7.Text
            proveedores.Label1.Text = Label9.Text
            proveedores.Show()
        Else
            Dim respuesta As DialogResult

            respuesta = MessageBox.Show("SE DETECTO UNA ACTUALIZACION EN LOS ASIENTOS DE APERTURA, DESEA ACTUALIZAR CON LA INFORMACION GUARDADA DE DOCUMENTOS RECIBIDOS?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                Form_regis.Show()
            Else
                proveedores.TextBox1.Text = Label5.Text
                proveedores.TextBox2.Text = Label7.Text
                proveedores.Label1.Text = Label9.Text
                proveedores.Show()
            End If

        End If


    End Sub

    Private Sub ReporteProveedoresToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ReporteProveedoresToolStripMenuItem1.Click
        Rpt_cuentaxPagar.TextBox1.Text = Label5.Text
        Rpt_cuentaxPagar.TextBox2.Text = Label7.Text
        Rpt_cuentaxPagar.Label5.Text = Label9.Text
        Rpt_cuentaxPagar.Show()
    End Sub

    Private Sub DataClientesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataClientesToolStripMenuItem.Click
        Form_cuentarxcobrar.TextBox1.Text = Label5.Text
        Form_cuentarxcobrar.TextBox2.Text = Label7.Text
        Form_cuentarxcobrar.Label1.Text = Label9.Text
        Form_cuentarxcobrar.Show()
    End Sub

    Private Sub CompraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompraToolStripMenuItem.Click

        Ocs.Label16.Text = Label9.Text
        Ocs.Label17.Text = Label5.Text
        Ocs.Show()
    End Sub

    Private Sub ToolStripLabel1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub NIAToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem4.Click
        Nia_Produccion.TextBox1.Text = "0700"
        Nia_Produccion.TextBox13.Text = "0700"
        Nia_Produccion.TextBox14.Text = "0701"
        Nia_Produccion.TextBox16.Text = "APLICACIONES Y PIEZAS"
        Nia_Produccion.TextBox11.Text = "APLICACIONES Y PIEZAS"
        Nia_Produccion.TextBox2.Text = "AREA APLICACIONES"
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub NSAToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem4.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0700"
        nsa_produccion.TextBox13.Text = "0700"
        nsa_produccion.TextBox14.Text = "0701"
        nsa_produccion.TextBox16.Text = "APLICACIONES Y PIEZAS"
        nsa_produccion.TextBox11.Text = "APLICACIONES Y PIEZAS"
        nsa_produccion.TextBox2.Text = "AREA APLICACIONES"
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.ShowDialog()
    End Sub

    Private Sub RUBROSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RUBROSToolStripMenuItem.Click
        'rubro2.Label1.Text = Label9.Text
        'rubro2.Label3.Text = Label5.Text
        'rubro2.Show()
        tabla_rubros.Show()
    End Sub

    Private Sub CalidadTelaCrudaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CalidadTelaCrudaToolStripMenuItem1.Click
        REPORTE_CALIDADCC.Show()
    End Sub

    Private Sub CodificacionDeProductosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CodificacionDeProductosToolStripMenuItem.Click
        Codificacion_Prodcutos.Label13.Text = Label9.Text
        Codificacion_Prodcutos.Label14.Text = "2"
        Codificacion_Prodcutos.Show()
    End Sub

    Private Sub MermaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MermaToolStripMenuItem.Click
        merma.Label4.Text = Label9.Text
        merma.Label5.Text = Label5.Text
        merma.TextBox2.Text = Label5.Text
        merma.Show()
    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem7.Click
        guia_talleres.Label23.Text = "0700"
        guia_talleres.Label25.Text = "APLICACIONES Y PIEZAS"
        guia_talleres.TextBox24.Text = "0701"
        guia_talleres.TextBox23.Text = "APLICACIONES Y PIEZAS"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()

    End Sub

    Private Sub ALMACENToolStripMenuItem5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub AnalisisMovimientoProduccionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalisisMovimientoProduccionToolStripMenuItem.Click
        Analisis_movimientos.Label4.Text = Label9.Text
        Analisis_movimientos.Show()
    End Sub

    Private Sub APLICACIONESYPIEZASToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles APLICACIONESYPIEZASToolStripMenuItem.Click

    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem8.Click
        guia_talleres.Label23.Text = "0300"
        guia_talleres.Label25.Text = "CORTE"
        guia_talleres.TextBox24.Text = "0301"
        guia_talleres.TextBox23.Text = "SERVICIO CORTE"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()
    End Sub

    Private Sub ParteIngresoSalidaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ParteIngresoSalidaToolStripMenuItem.Click
        Ingreso_salida_Produ.Label6.Text = Label9.Text
        Ingreso_salida_Produ.Show()
    End Sub

    Private Sub FacturasEliminadasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FacturasEliminadasToolStripMenuItem.Click
        Otras_Facturas.TextBox1.Text = Label5.Text
        Otras_Facturas.Label1.Text = Label9.Text
        Otras_Facturas.Label4.Text = 1

        Otras_Facturas.TextBox2.Text = Label7.Text
        Otras_Facturas.Show()
    End Sub

    Private Sub FacturasExportacionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FacturasExportacionToolStripMenuItem.Click
        Otras_Facturas.TextBox1.Text = Label5.Text
        Otras_Facturas.Label1.Text = Label9.Text
        Otras_Facturas.Label4.Text = 2

        Otras_Facturas.TextBox2.Text = Label7.Text
        Otras_Facturas.Show()
    End Sub

    Private Sub LiquidacionServiciosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LiquidacionServiciosToolStripMenuItem.Click
        'Liquidacion.Label3.Text = Label5.Text
        'Liquidacion.Label4.Text = Label9.Text
        'Liquidacion.Show()
        Liquidacion_Produccion.Label2.Text = Label5.Text
        Liquidacion_Produccion.Label3.Text = Label9.Text
        Liquidacion_Produccion.Show()
    End Sub

    Private Sub StatusStrip_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub ORDENSERVICIOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ORDENSERVICIOToolStripMenuItem.Click

    End Sub

    Private Sub ReporteFacturasOrdenCOmpraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteFacturasOrdenCOmpraToolStripMenuItem.Click
        Rpt_factura_oc.TextBox1.Text = Label7.Text
        Rpt_factura_oc.TextBox2.Text = Label5.Text
        Rpt_factura_oc.Label3.Text = Label9.Text
        Rpt_factura_oc.Show()
    End Sub

    Private Sub OrdenServicioToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        os.Label16.Text = Label9.Text
        os.Label17.Text = Label5.Text
        os.Show()
    End Sub

    Private Sub ServicioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ServicioToolStripMenuItem.Click
        os.Label16.Text = Label9.Text
        os.Label17.Text = Label5.Text
        os.Show()
    End Sub

    Private Sub CONFECCIONToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CONFECCIONToolStripMenuItem.Click
        If Label9.Text = "01" Then
            Formato_solic_avios.TextBox1.Text = "CV"
        Else
            Formato_solic_avios.TextBox1.Text = "CG"
        End If

        Formato_solic_avios.Label7.Text = Label9.Text
        Formato_solic_avios.Label15.Text = Label5.Text
        Formato_solic_avios.Label16.Text = "CONFECCION"
        Formato_solic_avios.Label17.Text = Label3.Text
        Formato_solic_avios.CheckBox1.Checked = True
        Formato_solic_avios.Show()
    End Sub

    Private Sub ACABADOSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ACABADOSToolStripMenuItem.Click
        If Label9.Text = "01" Then
            Formato_solic_avios.TextBox1.Text = "AV"
        Else
            Formato_solic_avios.TextBox1.Text = "AG"
        End If
        Formato_solic_avios.Label7.Text = Label9.Text
        Formato_solic_avios.Label15.Text = Label5.Text
        Formato_solic_avios.Label16.Text = "ACABADOS"
        Formato_solic_avios.Label17.Text = Label3.Text
        Formato_solic_avios.CheckBox2.Checked = True
        Formato_solic_avios.Show()
    End Sub

    Private Sub DespachosProductosTerminadosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DespachosProductosTerminadosToolStripMenuItem.Click
        Rpt_pt.Show()
    End Sub

    Private Sub ProduccionDiariaCorteToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ProduccionDiariaCorteToolStripMenuItem.Click
        Produccion_diaria_cortee.Label3.Text = Label9.Text
        Produccion_diaria_cortee.Show()
    End Sub

    Private Sub ReporteOrdenServicioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteOrdenServicioToolStripMenuItem.Click
        ORDE_SERVISO_AC.Label3.Text = Label9.Text
        ORDE_SERVISO_AC.Show()
    End Sub

    Private Sub ReportePlanillaPaqueteoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportePlanillaPaqueteoToolStripMenuItem.Click
        PLANILLA_PAQUETEO.Label2.Text = Label9.Text
        PLANILLA_PAQUETEO.TextBox1.Select()
        PLANILLA_PAQUETEO.Show()
    End Sub

    Private Sub ORDENToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ORDENToolStripMenuItem.Click
        Ord_Corte.Label44.Text = Label9.Text
        Ord_Corte.Show()
    End Sub

    Private Sub RequerimientoAviosProduccionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RequerimientoAviosProduccionToolStripMenuItem.Click

    End Sub

    Private Sub ProcesoSuministrosYRepuestosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProcesoSuministrosYRepuestosToolStripMenuItem.Click

    End Sub

    Private Sub ProduccionDiariaDeCorteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProduccionDiariaDeCorteToolStripMenuItem.Click
        Produccion_diaria_cortee.Label3.Text = Label9.Text
        Produccion_diaria_cortee.Show()
    End Sub

    Private Sub AnalisisOpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalisisOpToolStripMenuItem.Click
        Analisis_Op.TextBox2.Text = Label7.Text
        Analisis_Op.TextBox3.Text = Label9.Text
        Analisis_Op.Show()
    End Sub

    Private Sub REQUERIMIENTODEAVIOSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles REQUERIMIENTODEAVIOSToolStripMenuItem.Click

    End Sub

    Private Sub StatusDeProduccionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatusDeProduccionToolStripMenuItem.Click
        Form7.TextBox4.Text = Label9.Text
        Form7.TextBox1.Text = Label7.Text
        Form7.Show()
    End Sub

    Private Sub ValorizarAlmacenesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ValorizarAlmacenesToolStripMenuItem.Click
        'Almacenes_Precios.Show()
    End Sub

    Private Sub StatusProduccionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatusProduccionToolStripMenuItem.Click
        Form7.TextBox4.Text = Label9.Text
        Form7.TextBox1.Text = Label7.Text
        Form7.Show()
    End Sub

    Private Sub PRODUCCIONDIARIACORTEToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        Prod_corteeee.Label4.Text = Label7.Text
        Prod_corteeee.Label5.Text = Label9.Text
        Prod_corteeee.Show()
    End Sub

    Private Sub ORDERSERVICIOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ORDERSERVICIOToolStripMenuItem.Click
        'os.Label16.Text = Label9.Text
        'os.Label17.Text = Label5.Text
        'os.Show()
    End Sub

    Private Sub NIAToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem5.Click
        Nia_Produccion.TextBox1.Text = "0300"
        Nia_Produccion.TextBox13.Text = "0300"
        Nia_Produccion.TextBox14.Text = "0301"
        Nia_Produccion.TextBox16.Text = "CORTE"
        Nia_Produccion.TextBox11.Text = "CORTE"
        Nia_Produccion.TextBox2.Text = "AREA CORTE"
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub NSAToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem5.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0300"
        nsa_produccion.TextBox13.Text = "0300"
        nsa_produccion.TextBox14.Text = "0301"
        nsa_produccion.TextBox16.Text = "CORTE"
        nsa_produccion.TextBox11.Text = "SERVICIO CORTE"
        nsa_produccion.TextBox2.Text = "AREA CORTE"
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.ShowDialog()
    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem9.Click
        guia_talleres.Label23.Text = "0300"
        guia_talleres.Label25.Text = "CORTE"
        guia_talleres.TextBox24.Text = "0301"
        guia_talleres.TextBox23.Text = "SERVICIO CORTE"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()
    End Sub

    Private Sub NIAToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem6.Click
        Nia_Produccion.TextBox1.Text = "0700"
        Nia_Produccion.TextBox13.Text = "0700"
        Nia_Produccion.TextBox14.Text = "0701"
        Nia_Produccion.TextBox16.Text = "APLICACIONES Y PIEZAS"
        Nia_Produccion.TextBox11.Text = "APLICACIONES Y PIEZAS"
        Nia_Produccion.TextBox2.Text = "AREA APLICACIONES"
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub NSAToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem6.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0700"
        nsa_produccion.TextBox13.Text = "0700"
        nsa_produccion.TextBox14.Text = "0701"
        nsa_produccion.TextBox16.Text = "APLICACIONES Y PIEZAS"
        nsa_produccion.TextBox11.Text = "APLICACIONES Y PIEZAS"
        nsa_produccion.TextBox2.Text = "AREA APLICACIONES"
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.ShowDialog()
    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem10_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem10.Click
        guia_talleres.Label23.Text = "0700"
        guia_talleres.Label25.Text = "APLICACIONES Y PIEZAS"
        guia_talleres.TextBox24.Text = "0701"
        guia_talleres.TextBox23.Text = "APLICACIONES Y PIEZAS"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()
    End Sub

    Private Sub NIAToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem7.Click
        Nia_Produccion.TextBox1.Text = "0400"
        Nia_Produccion.TextBox13.Text = "0400"
        Nia_Produccion.TextBox14.Text = "0421"
        Nia_Produccion.TextBox16.Text = "CONFECCION"
        Nia_Produccion.TextBox11.Text = "CONFECCION"
        Nia_Produccion.TextBox2.Text = "AREA CONFECCION"
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub NSAToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem7.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0400"
        nsa_produccion.TextBox13.Text = "0400"
        nsa_produccion.TextBox14.Text = "0421"
        nsa_produccion.TextBox16.Text = "CONFECCION"
        nsa_produccion.TextBox11.Text = "CONFECCION"
        nsa_produccion.TextBox2.Text = "AREA CONFECCION"
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.ShowDialog()
    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem11_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem11.Click
        guia_talleres.Label23.Text = "0400"
        guia_talleres.Label25.Text = "CONFECCION"
        guia_talleres.TextBox24.Text = "0421"
        guia_talleres.TextBox23.Text = "SERVICIO CONFECCION"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()
    End Sub

    Private Sub NIAToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem8.Click
        Nia_Produccion.TextBox1.Text = "0600"
        Nia_Produccion.TextBox13.Text = "0600"
        Nia_Produccion.TextBox14.Text = "0601"
        Nia_Produccion.TextBox16.Text = "LAVANDERIA"
        Nia_Produccion.TextBox11.Text = "LAVANDERIA"
        Nia_Produccion.TextBox2.Text = "AREA LAVANDERIA"
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub NSAToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem8.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0600"
        nsa_produccion.TextBox13.Text = "0600"
        nsa_produccion.TextBox14.Text = "0601"
        nsa_produccion.TextBox16.Text = "LAVANDERIA"
        nsa_produccion.TextBox11.Text = "LAVANDERIA"
        nsa_produccion.TextBox2.Text = "AREA LAVANDERIA"
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Show()
    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem12_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem12.Click
        guia_talleres.Label23.Text = "0600"
        guia_talleres.Label25.Text = "LAVANDERIA"
        guia_talleres.TextBox24.Text = "0601"
        guia_talleres.TextBox23.Text = "SERVICIO LAVANDERIA"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()
    End Sub

    Private Sub GUIAREMISIONToolStripMenuItem13_Click(sender As Object, e As EventArgs) Handles GUIAREMISIONToolStripMenuItem13.Click
        guia_talleres.Label23.Text = "0500"
        guia_talleres.Label25.Text = "ACABADO"
        guia_talleres.TextBox24.Text = "0522"
        guia_talleres.TextBox23.Text = "SERVICIO ACABADO"
        guia_talleres.Label30.Text = Label9.Text
        guia_talleres.Label29.Text = Label5.Text
        guia_talleres.Show()
    End Sub

    Private Sub NIAToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles NIAToolStripMenuItem9.Click
        Nia_Produccion.TextBox1.Text = "0500"
        Nia_Produccion.TextBox13.Text = "0500"
        Nia_Produccion.TextBox14.Text = "0522"
        Nia_Produccion.TextBox16.Text = "ACABADO"
        Nia_Produccion.TextBox11.Text = "ACABADO"
        Nia_Produccion.TextBox2.Text = "AREA ACABADO"
        Nia_Produccion.Label4.Text = Label9.Text
        Nia_Produccion.Label11.Text = Label5.Text
        Nia_Produccion.Show()
    End Sub

    Private Sub NSAToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles NSAToolStripMenuItem9.Click
        nsa_produccion.Close()
        nsa_produccion.TextBox1.Text = "0500"
        nsa_produccion.TextBox13.Text = "0500"
        nsa_produccion.TextBox14.Text = "0522"
        nsa_produccion.TextBox16.Text = "ACABADO"
        nsa_produccion.TextBox11.Text = "ACABADO"
        nsa_produccion.TextBox2.Text = "AREA ACABADO"
        nsa_produccion.Label4.Text = Label9.Text
        nsa_produccion.Label11.Text = Label5.Text
        nsa_produccion.Show()
    End Sub

    Private Sub PARTEINGRESOSALIDAToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PARTEINGRESOSALIDAToolStripMenuItem1.Click
        Ingreso_salida_Produ.Label6.Text = Label9.Text
        Ingreso_salida_Produ.Show()
    End Sub

    Private Sub ANALISISMOVIMIENTOPRODUCCIONToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ANALISISMOVIMIENTOPRODUCCIONToolStripMenuItem1.Click
        Analisis_movimientos.Label4.Text = Label9.Text
        Analisis_movimientos.Show()
    End Sub

    Private Sub REPORTEORDENSERVICIOToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles REPORTEORDENSERVICIOToolStripMenuItem1.Click
        ORDE_SERVISO_AC.Label3.Text = Label9.Text
        ORDE_SERVISO_AC.Show()
    End Sub

    Private Sub STATUSPRODUCCIONToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles STATUSPRODUCCIONToolStripMenuItem1.Click
        Form7.TextBox4.Text = Label9.Text
        Form7.TextBox1.Text = Label7.Text
        Form7.Show()
    End Sub

    Private Sub AnalisisDeOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalisisDeOPToolStripMenuItem.Click
        Analisis_Op.TextBox2.Text = Label7.Text
        Analisis_Op.TextBox3.Text = Label9.Text
        Analisis_Op.Show()
    End Sub

    Private Sub TablasComisionesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TablasComisionesToolStripMenuItem.Click
        Tabla_Comisiones.TextBox1.Text = Label5.Text
        Tabla_Comisiones.TextBox2.Text = Label7.Text
        Tabla_Comisiones.Label4.Text = Label9.Text
        Tabla_Comisiones.Show()
    End Sub

    Private Sub HojaReclamoCalidadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HojaReclamoCalidadToolStripMenuItem.Click
        Buzin_reclamos.Show()
    End Sub

    Private Sub ToolStripButton6_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Op_Manufactura.TextBox10.Text = Label3.Text
        Op_Manufactura.Label33.Text = Label9.Text
        Op_Manufactura.Show()
    End Sub

    Private Sub HojaReclamoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HojaReclamoToolStripMenuItem.Click
        abrir()
        Dim sql102 As String = "SELECT codigo_ven FROM  custom_vianny.dbo.Vendedores WHERE alias_ven ='" + Label3.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        If Rsr2.Read() = True Then
            Hoja_Reclamos.TextBox14.Text = Rsr2(0)
        End If

        Rsr2.Close()
        Hoja_Reclamos.TextBox9.Text = Label3.Text
        Hoja_Reclamos.Label26.Text = Label9.Text
        Hoja_Reclamos.Label27.Text = Label5.Text
        Hoja_Reclamos.Show()
    End Sub

    Private Sub StockFisicoToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles StockFisicoToolStripMenuItem4.Click
        Stock_Fisico.Label5.Text = Label5.Text
        Stock_Fisico.Label6.Text = Label9.Text
        Stock_Fisico.Show()
    End Sub

    Private Sub MoviemientoAlmacenTelaAviosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoviemientoAlmacenTelaAviosToolStripMenuItem.Click
        Rpt_ing_sal.Label7.Text = Label9.Text
        Rpt_ing_sal.Show()
    End Sub

    Private Sub SeguimientoPedidoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SeguimientoPedidoToolStripMenuItem1.Click
        Seguimiento_Despachos.Label3.Text = Label9.Text
        Seguimiento_Despachos.Label4.Text = Label5.Text
        Seguimiento_Despachos.Show()
    End Sub

    Private Sub ReporteSeguimientoReclamosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteSeguimientoReclamosToolStripMenuItem.Click
        Seguimiento_reclamo.Show()
    End Sub

    Private Sub ReporteTelaAcabadaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteTelaAcabadaToolStripMenuItem.Click
        Reporte_calidad_aca.Show()
    End Sub

    Private Sub AprobarNotaPedidoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AprobarNotaPedidoToolStripMenuItem1.Click
        Aprobacion_comercial.Label4.Text = Label9.Text
        Aprobacion_comercial.Label5.Text = Label7.Text
        Aprobacion_comercial.Label7.Text = Label3.Text
        Aprobacion_comercial.Show()
    End Sub

    Private Sub AprobarNotaPedidoCobranzasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AprobarNotaPedidoCobranzasToolStripMenuItem.Click
        Aprobar_notapedido.Label3.Text = Label9.Text
        Aprobar_notapedido.Label5.Text = Label7.Text
        Aprobar_notapedido.Label6.Text = Label3.Text
        Aprobar_notapedido.Show()
    End Sub

    Private Sub ValorizarIngresosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ValorizarIngresosToolStripMenuItem.Click
        Almacenes_Precios.Show()
    End Sub

    Private Sub ValorizarSalidasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ValorizarSalidasToolStripMenuItem.Click

        Valorizr_Salidas.Label4.Text = Label9.Text
        Valorizr_Salidas.Label5.Text = Label5.Text
        Valorizr_Salidas.Show()
    End Sub

    Private Sub COMPRAToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles COMPRAToolStripMenuItem1.Click
        Ocs.Label16.Text = Label9.Text
        Ocs.Label17.Text = Label5.Text
        Ocs.Show()
    End Sub

    Private Sub SERVICIOToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SERVICIOToolStripMenuItem1.Click
        os.Label16.Text = Label9.Text
        os.Label17.Text = Label5.Text
        os.Show()
    End Sub

    Private Sub ReporteOrdenCompraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteOrdenCompraToolStripMenuItem.Click
        Orden_Compra_Ac.Label3.Text = Label9.Text
        Orden_Compra_Ac.Show()
    End Sub

    Private Sub FICHATECNICADISEÑOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FICHATECNICADISEÑOToolStripMenuItem.Click
        Ficha_udp.Label10.Text = Label9.Text

        Ficha_udp.Show()
    End Sub

    Private Sub OrdenCompraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrdenCompraToolStripMenuItem.Click
        Ocs.Label16.Text = Label9.Text
        Ocs.Label17.Text = Label5.Text
        Ocs.Show()
    End Sub

    Private Sub OrdenServicioToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles OrdenServicioToolStripMenuItem1.Click
        os.Label16.Text = Label9.Text
        os.Label17.Text = Label5.Text
        os.Show()
    End Sub

    Private Sub RequerimientoPedisodToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RequerimientoPedisodToolStripMenuItem.Click
        Requerimiento__comercial.Show()
    End Sub

    Private Sub INGRESARREQUERIMIENTOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles INGRESARREQUERIMIENTOToolStripMenuItem.Click
        Solicitud_telas.Show()
    End Sub

    Private Sub MARKETINGToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MARKETINGToolStripMenuItem.Click

    End Sub

    Private Sub ReporteDeUltimasVentasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteDeUltimasVentasToolStripMenuItem.Click
        Rpt_Ultimas_ventas.TextBox1.Text = Label7.Text
        Rpt_Ultimas_ventas.TextBox2.Text = Label9.Text
        Rpt_Ultimas_ventas.TextBox3.Text = Label5.Text
        Rpt_Ultimas_ventas.Show()
    End Sub

    Private Sub ReporteClientesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ReporteClientesToolStripMenuItem1.Click
        Vendedor_Cliente.Show()
    End Sub

    Private Sub ReporteFrecuenciaVentasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteFrecuenciaVentasToolStripMenuItem.Click
        Form9.TextBox1.Text = Label7.Text
        Form9.TextBox2.Text = Label9.Text
        Form9.TextBox3.Text = Label5.Text
        Form9.Show()
    End Sub

    Private Sub MatrizAviosAlamcenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MatrizAviosAlamcenToolStripMenuItem.Click
        FormatoMa.Label1.Text = Label9.Text
        FormatoMa.Label2.Text = Label5.Text
        FormatoMa.Label3.Text = Label3.Text

        FormatoMa.ShowDialog()
    End Sub

    Private Sub ReporteOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteOPToolStripMenuItem.Click
        Rppt_op.Label3.Text = Label9.Text
        Rppt_op.TextBox2.Text = Label7.Text
        Rppt_op.Show()
    End Sub

    Private Sub AlmacenDeSolicitudesDeAviosProduccionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlmacenDeSolicitudesDeAviosProduccionToolStripMenuItem.Click
        modulo_solicitud.Label1.Text = Label3.Text
        modulo_solicitud.Label2.Text = Label5.Text
        modulo_solicitud.Label3.Text = Label9.Text
        modulo_solicitud.Show()
    End Sub

    Private Sub ProductosDeLineaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductosDeLineaToolStripMenuItem.Click
        ProductosLinea.Label1.Text = Label5.Text
        ProductosLinea.Label2.Text = Label9.Text
        ProductosLinea.Show()
    End Sub

    Private Sub CONFECCIONToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles CONFECCIONToolStripMenuItem2.Click
        Asignacion_Confeccion.Show()
    End Sub

    Private Sub ACABADOSToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ACABADOSToolStripMenuItem2.Click
        Asignacion_Acabados.Show()
    End Sub

    Private Sub COMPRAToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles COMPRAToolStripMenuItem2.Click
        Ocs.Label16.Text = Label9.Text
        Ocs.Label17.Text = Label5.Text
        Ocs.Show()
    End Sub

    Private Sub SERIVCIOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SERIVCIOToolStripMenuItem.Click
        os.Label16.Text = Label9.Text
        os.Label17.Text = Label5.Text
        os.Show()
    End Sub

    Private Sub ReporteOpToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ReporteOpToolStripMenuItem1.Click
        Rppt_op.Label3.Text = Label9.Text
        Rppt_op.TextBox2.Text = Label7.Text
        Rppt_op.Show()
    End Sub

    Private Sub REPORTEORDENCOMPRAToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles REPORTEORDENCOMPRAToolStripMenuItem1.Click
        Orden_Compra_Ac.Label3.Text = Label9.Text
        Orden_Compra_Ac.Show()
    End Sub

    Private Sub ANALISIADEOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ANALISIADEOPToolStripMenuItem.Click
        Analisis_Op.TextBox2.Text = Label7.Text
        Analisis_Op.TextBox3.Text = Label9.Text
        Analisis_Op.Show()
    End Sub

    Private Sub NotaSalidaToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles NotaSalidaToolStripMenuItem5.Click
        Nsa_Pt.TextBox7.Text = Label3.Text
        Nsa_Pt.Label11.Text = Label5.Text
        Nsa_Pt.Label13.Text = Label9.Text
        Nsa_Pt.Show()
    End Sub

    Private Sub NotaIngresoToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles NotaIngresoToolStripMenuItem5.Click
        Nia_Pt.TextBox7.Text = Label3.Text
        Nia_Pt.Label11.Text = Label5.Text
        Nia_Pt.Label13.Text = Label9.Text
        Nia_Pt.Show()
    End Sub

    Private Sub LIQUIDACIONPRENDASToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LIQUIDACIONPRENDASToolStripMenuItem.Click
        Liqui_Produccion.Label15.Text = Label9.Text
        Liqui_Produccion.Show()
    End Sub

    Private Sub LiquidacionServicioTejeduriaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LiquidacionServicioTejeduriaToolStripMenuItem.Click
        Liq_servicios_T.Label3.Text = Label5.Text
        Liq_servicios_T.ShowDialog()
    End Sub

    Private Sub AprobarOrdenCompraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AprobarOrdenCompraToolStripMenuItem.Click
        Apro_Oc.Label7.Text = Label9.Text
        Apro_Oc.Show()
    End Sub

    Private Sub AprobarOrdenServicioToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AprobarOrdenServicioToolStripMenuItem.Click
        Aprob_os.Label7.Text = Label9.Text
        Aprob_os.Show()
    End Sub

    Private Sub STATUSACABADOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles STATUSACABADOToolStripMenuItem.Click

    End Sub

    Private Sub ODToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub GuiaRemisionToolStripMenuItem14_Click(sender As Object, e As EventArgs) Handles GuiaRemisionToolStripMenuItem14.Click
        Guia_remision_prenda.Label33.Text = Label9.Text
        Guia_remision_prenda.Label35.Text = "01"
        Guia_remision_prenda.Label36.Text = Label5.Text
        Guia_remision_prenda.Label37.Text = Label3.Text
        Guia_remision_prenda.Label23.Text = "01"
        Guia_remision_prenda.Label25.Text = "ALMACEN DE PT"
        Guia_remision_prenda.Show()
    End Sub

    Private Sub ReporteGuiasRemisionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteGuiasRemisionToolStripMenuItem.Click
        Rporte_Guias.Label4.Text = Label9.Text
        Rporte_Guias.Show()
    End Sub

    Private Sub RequerimientoDeTelaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RequerimientoDeTelaToolStripMenuItem1.Click
        Consumo_tela_Consolidado.Show()
    End Sub

    Private Sub AnalisisDeDespachosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalisisDeDespachosToolStripMenuItem.Click
        Reporte_ptt.TextBox1.Text = Label9.Text
        Reporte_ptt.TextBox2.Text = Label5.Text
        Reporte_ptt.Show()
    End Sub

    Private Sub CostosDeTelaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CostosDeTelaToolStripMenuItem.Click
        Costos_Tela.Label3.Text = Label5.Text
        Costos_Tela.Label4.Text = Label9.Text
        Costos_Tela.Show()
    End Sub

    Private Sub SagaFalabellaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SagaFalabellaToolStripMenuItem.Click

    End Sub

    Private Sub CdToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CdToolStripMenuItem.Click
        FormLpn.Show(Me)
    End Sub

    Private Sub TiendasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TiendasToolStripMenuItem.Click
        Packing_Saga.Show()
    End Sub

    Private Sub CostosDePrendaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CostosDePrendaToolStripMenuItem.Click
        Costo_Prenda.TextBox2.Text = Label9.Text
        Costo_Prenda.TextBox4.Text = Label5.Text
        Costo_Prenda.Show()
    End Sub

    Private Sub GenerarLetraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerarLetraToolStripMenuItem.Click

    End Sub

    Private Sub ActualizarLetraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActualizarLetraToolStripMenuItem.Click
        Form10.Show()
    End Sub

    Private Sub GenerarLetraToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles GenerarLetraToolStripMenuItem1.Click
        Letra.TextBox6.Text = Mid(Label5.Text, 3, 2)
        Letra.Label21.Text = 1
        Letra.ShowDialog()
    End Sub

    Private Sub ReporteDeLetrasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteDeLetrasToolStripMenuItem.Click
        Rpt_letra.Show()
    End Sub

    Private Sub TablaServiciosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TablaServiciosToolStripMenuItem.Click
        SERVICIOS.Show()
    End Sub

    Private Sub TablaClienteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TablaClienteToolStripMenuItem.Click
        Tabla_Clientes_Comerciales.Show()
    End Sub

    Private Sub TablaMarcaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TablaMarcaToolStripMenuItem.Click
        Listado_Marca.Show()
    End Sub

    Private Sub STATUSLAVANDERIAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles STATUSLAVANDERIAToolStripMenuItem.Click

    End Sub

    Private Sub StatusMuestrasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatusMuestrasToolStripMenuItem.Click
        Gand_Udp.Show()
    End Sub

    Private Sub CIERREPROCESOSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CIERREPROCESOSToolStripMenuItem.Click
        Cierre_Proc_Od.Show()
    End Sub

    Private Sub STATUSODToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles STATUSODToolStripMenuItem.Click

    End Sub

    Private Sub PRODUCCIONToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PRODUCCIONToolStripMenuItem.Click

    End Sub

    Private Sub CIERREPROCESOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CIERREPROCESOToolStripMenuItem.Click
        Cierre_Proc_Od.Show()
    End Sub

    Private Sub FICHATECNICAToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles FICHATECNICAToolStripMenuItem2.Click
        FICHA_ALMACE.Show()
    End Sub

    'Private Sub ASIGNACIONDETELAOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASIGNACIONDETELAOPToolStripMenuItem.Click
    '    Asignacion_op_tela.Show()
    'End Sub

    Private Sub CONSUMODEODToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CONSUMODEODToolStripMenuItem.Click
        FICHA_CONSUMO.Label8.Text = Label9.Text
        FICHA_CONSUMO.Show()
    End Sub

    Private Sub ASIGNACIONDEMOLDEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASIGNACIONDEMOLDEToolStripMenuItem.Click
        Asignacion_Molde.Label9.Text = Label9.Text
        Asignacion_Molde.Label13.Text = Label3.Text
        Asignacion_Molde.Show()
    End Sub

    Private Sub AsientoCostosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AsientoCostosToolStripMenuItem.Click
        Importar_Asiento.Label6.Text = Label5.Text
        Importar_Asiento.Label7.Text = Label9.Text
        Importar_Asiento.Show()
    End Sub

    Private Sub MemorandumToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MemorandumToolStripMenuItem.Click
        Memorándum.Show()
    End Sub

    Private Sub PlanillaMovilidadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PlanillaMovilidadToolStripMenuItem.Click
        Planilla_Movilidas.Show()
    End Sub

    Private Sub ConsultarFechasDocumentosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsultarFechasDocumentosToolStripMenuItem.Click
        Fecha__Factura.Label1.Text = Label9.Text
        Fecha__Factura.Show()
    End Sub

    Private Sub ASIGNACIONDETELAOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASIGNACIONDETELAOPToolStripMenuItem.Click

    End Sub

    Private Sub BUZONDESOLICITUDESToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BUZONDESOLICITUDESToolStripMenuItem.Click

    End Sub

    Private Sub ListadoDeOpPorPOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListadoDeOpPorPOToolStripMenuItem.Click

    End Sub

    Private Sub CotizacionToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles CotizacionToolStripMenuItem.Click
        HojaContizacionAlm.Show()
    End Sub

    Private Sub ASistenciaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASistenciaToolStripMenuItem.Click
        'FormatoAsistencia.TextBox1.Text = Label5.Text
        FormatoAsistencia.Show()
    End Sub

    Private Sub AbrirCerrarOpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AbrirCerrarOpToolStripMenuItem.Click
        AperturarOp.Label4.Text = Label9.Text
        AperturarOp.Show()
    End Sub

    Private Sub ReporteSalidasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReporteSalidasToolStripMenuItem.Click
        ReporteSalidas.Label4.Text = Label9.Text
        ReporteSalidas.Show()
    End Sub

    Private Sub SinonimoProductoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SinonimoProductoToolStripMenuItem.Click
        FormSinonimo.Label3.Text = Label9.Text
        FormSinonimo.Show(Me)
    End Sub

    Private Sub SeguimientoProduccionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeguimientoProduccionToolStripMenuItem.Click

        StatusProduccionPP.Label2.Text = Label9.Text
        StatusProduccionPP.TextBox1.Text = Label7.Text
        StatusProduccionPP.Show(Me)
    End Sub

    Private Sub HojaCotizacionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HojaCotizacionToolStripMenuItem.Click
        ListaHojaCotizacion.Label3.Text = Label4.Text
        ListaHojaCotizacion.ShowDialog()
    End Sub

    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        StatusProduccionPP.Label2.Text = Label9.Text
        StatusProduccionPP.Label5.Text = Label5.Text
        StatusProduccionPP.TextBox1.Text = Label7.Text
        StatusProduccionPP.Show(Me)
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        Cierre_Proc_Od.Show(Me)
    End Sub

    Private Sub CargarInformacionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CargarInformacionToolStripMenuItem.Click
        Packing_Saga.Show()
    End Sub

    Private Sub ImprimirLpnRojoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImprimirLpnRojoToolStripMenuItem.Click
        FormLpnRipley.Label3.Text = 1
        FormLpnRipley.Show(Me)
    End Sub

    Private Sub ImprimirLnpBlancoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImprimirLnpBlancoToolStripMenuItem.Click
        FormLpnRipley.Label3.Text = 2
        FormLpnRipley.Show(Me)
    End Sub
End Class
