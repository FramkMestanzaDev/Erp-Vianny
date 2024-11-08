Imports System.Data.SqlClient
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Public Class MDIParent2
    Dim da As New DataTable
    Public conx As SqlConnection
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


    Dim Rsr2, Rsr3 As SqlDataReader
    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Dim imageList As New ImageList()
        imageList.Images.Add(Global.Modulo_Ventas.Resources._3d_applications_folder_20538)
        imageList.Images.Add(Global.Modulo_Ventas.Resources.bindersbluedocuments_carpetas_azu_10985)
        TreeView1.ImageList = imageList
        e.Node.ImageIndex = 0
        e.Node.SelectedImageIndex = 1
        Select Case e.Node.Text
            Case "Status Od"

                Cierre_Proc_Od.Show(Me)

            Case "Ficha Tecnica"

                FICHA_ALMACE.Show(Me)
            Case "Asignacion Molde"

                Asignacion_Molde.Label9.Text = Label10.Text
                Asignacion_Molde.Label13.Text = Label6.Text
                Asignacion_Molde.Show(Me)
            Case "Consumo Op"

                ListaOpConsumo.Label3.Text = Label10.Text
                ListaOpConsumo.Label4.Text = Label8.Text
                ListaOpConsumo.Label5.Text = Label6.Text
                ListaOpConsumo.Show(Me)
            Case "Listado de PM x OD"

                OdxPm.TextBox1.Text = Label9.Text
                OdxPm.Label3.Text = Label10.Text
                OdxPm.Label4.Text = "1"
                OdxPm.Label2.Text = "PM"
                OdxPm.Show(Me)
            Case "Planeamiento Udp"

                abrir()
                Dim cmd As New SqlDataAdapter("SELECT CONVERT(varchar,fcom_3,23) AS 'FECHA INICIO',CONVERT(varchar,FCome1_3,23) AS 'FECHA ENTREGA', program_3 AS PM,SUBSTRING(ncom_3 ,1,9) AS OD,SUBSTRING( ncom_3,10,1) AS VERSION,c.nomb_10 AS CLIENTE, marcacli_3
,descri_3 AS DESCRIPCION,estilo_3 as 'ESTILO CLIENTE', cants_3 AS SOLICITADO,q.otros1_3 AS COLOR,q.alterno_3 as LAVADO ,case when merchan_3 = '0001' then 'VSILVERIO' WHEN merchan_3 = '0003' then 'GBALVIN' WHEN merchan_3 = '0031' then 'JPOZZI' WHEN merchan_3 = '0033' then 'HQUISPE' ELSE 'VVARGAR' END AS COMERCIAL,
telaprinc_3 AS TELA,(SELECT CASE WHEN    MAX(EtaRut)= '01' THEN 'ANALISIS PRENDA' WHEN   MAX(EtaRut)= '02' THEN 'CREACION MOLDE' WHEN   MAX(EtaRut)= '03' THEN 'CORTE' WHEN   MAX(EtaRut)= '04' THEN 'APLICACIONES' 
WHEN    MAX(EtaRut)= '05' THEN 'CONFECCION' WHEN    MAX(EtaRut)= '06' THEN 'LAVADO' WHEN    MAX(EtaRut)= '07' THEN 'ACABADOS' WHEN    MAX(EtaRut)= '08' THEN 'ENTREGA COMERCIAL' else 'PENDIENTE'  END AS UBICACION FROM Ruta_Udp WHERE OdRut =Q.ncom_3 AND CiRut ='1')AS UBICACION,'' AS PROYECCION
FROM custom_vianny.DBO.qag0300 q
left join custom_vianny.dbo.cag1000 c on q.ccia = c.ccia and (q.fich_3 = c.fich_10 or q.fich_3 =c.ruc_10)
where ncom_3 like '03%' AND C.ccia ='01' and finped_3 = 0 AND flag_3 ='1' ORDER BY ncom_3", conx)
                cmd.Fill(da)
                DataGridView1.DataSource = da
                Dim ext As New ExportarSoloData
                ext.llenarExcel(DataGridView1)
            Case "Planeamiento Op"

                Planemaiento_OP.Label2.Text = Label10.Text
                Planemaiento_OP.Show(Me)
            Case "Nota Ingreso C."

                Nia_Produccion.TextBox1.Text = "0300"
                Nia_Produccion.TextBox13.Text = "0300"
                Nia_Produccion.TextBox14.Text = "0301"
                Nia_Produccion.TextBox16.Text = "CORTE"
                Nia_Produccion.TextBox11.Text = "CORTE"
                Nia_Produccion.TextBox2.Text = "AREA CORTE"
                Nia_Produccion.Label11.Text = Label8.Text
                Nia_Produccion.Label4.Text = Label10.Text
                Nia_Produccion.Show(Me)
            Case "Nota Salida C."

                nsa_produccion.Close()
                nsa_produccion.TextBox1.Text = "0300"
                nsa_produccion.TextBox13.Text = "0300"
                nsa_produccion.TextBox14.Text = "0301"
                nsa_produccion.TextBox16.Text = "CORTE"
                nsa_produccion.TextBox11.Text = "SERVICIO CORTE"
                nsa_produccion.TextBox2.Text = "AREA CORTE"
                nsa_produccion.Label11.Text = Label8.Text
                nsa_produccion.Label4.Text = Label10.Text
                nsa_produccion.Show(Me)
            Case "Guia Remision C."

                guia_talleres.Label23.Text = "0300"
                guia_talleres.Label25.Text = "CORTE"
                guia_talleres.TextBox24.Text = "0301"
                guia_talleres.TextBox23.Text = "SERVICIO CORTE"
                guia_talleres.Label30.Text = Label10.Text
                guia_talleres.Label29.Text = Label8.Text
                guia_talleres.Show(Me)
            Case "Nota Ingreso Ap."

                Nia_Produccion.TextBox1.Text = "0700"
                Nia_Produccion.TextBox13.Text = "0700"
                Nia_Produccion.TextBox14.Text = "0701"
                Nia_Produccion.TextBox16.Text = "APLICACIONES Y PIEZAS"
                Nia_Produccion.TextBox11.Text = "APLICACIONES Y PIEZAS"
                Nia_Produccion.TextBox2.Text = "AREA APLICACIONES"
                Nia_Produccion.Label11.Text = Label8.Text
                Nia_Produccion.Label4.Text = Label10.Text
                Nia_Produccion.Show(Me)
            Case "Nota Salida Ap."

                nsa_produccion.Close()
                nsa_produccion.TextBox1.Text = "0700"
                nsa_produccion.TextBox13.Text = "0700"
                nsa_produccion.TextBox14.Text = "0701"
                nsa_produccion.TextBox16.Text = "APLICACIONES Y PIEZAS"
                nsa_produccion.TextBox11.Text = "APLICACIONES Y PIEZAS"
                nsa_produccion.TextBox2.Text = "AREA APLICACIONES"
                nsa_produccion.Label11.Text = Label8.Text
                nsa_produccion.Label4.Text = Label10.Text
                nsa_produccion.Show(Me)
            Case "Guia Remision Ap."

                guia_talleres.Label23.Text = "0700"
                guia_talleres.Label25.Text = "APLICACIONES Y PIEZAS"
                guia_talleres.TextBox24.Text = "0701"
                guia_talleres.TextBox23.Text = "APLICACIONES Y PIEZAS"
                guia_talleres.Label30.Text = Label10.Text
                guia_talleres.Label29.Text = Label8.Text
                guia_talleres.Show(Me)
            Case "Nota Ingreso Co."

                Nia_Produccion.TextBox1.Text = "0400"
                Nia_Produccion.TextBox13.Text = "0400"
                Nia_Produccion.TextBox14.Text = "0421"
                Nia_Produccion.TextBox16.Text = "CONFECCION"
                Nia_Produccion.TextBox11.Text = "CONFECCION"
                Nia_Produccion.TextBox2.Text = "AREA CONFECCION"
                Nia_Produccion.Label11.Text = Label8.Text
                Nia_Produccion.Label4.Text = Label10.Text
                Nia_Produccion.Show(Me)
            Case "Nota Salida Co."

                nsa_produccion.Close()
                nsa_produccion.TextBox1.Text = "0400"
                nsa_produccion.TextBox13.Text = "0400"
                nsa_produccion.TextBox14.Text = "0421"
                nsa_produccion.TextBox16.Text = "CONFECCION"
                nsa_produccion.TextBox11.Text = "CONFECCION"
                nsa_produccion.TextBox2.Text = "AREA CONFECCION"
                nsa_produccion.Label11.Text = Label8.Text
                nsa_produccion.Label4.Text = Label10.Text
                nsa_produccion.Show(Me)
            Case "Guia Remision Co."

                guia_talleres.Label23.Text = "0400"
                guia_talleres.Label25.Text = "CONFECCION"
                guia_talleres.TextBox24.Text = "0421"
                guia_talleres.TextBox23.Text = "SERVICIO CONFECCION"
                guia_talleres.Label30.Text = Label10.Text
                guia_talleres.Label29.Text = Label8.Text
                guia_talleres.Show(Me)
            Case "Nota Ingreso L."

                Nia_Produccion.TextBox1.Text = "0600"
                Nia_Produccion.TextBox13.Text = "0600"
                Nia_Produccion.TextBox14.Text = "0601"
                Nia_Produccion.TextBox16.Text = "LAVANDERIA"
                Nia_Produccion.TextBox11.Text = "LAVANDERIA"
                Nia_Produccion.TextBox2.Text = "AREA LAVANDERIA"
                Nia_Produccion.Label4.Text = Label10.Text
                Nia_Produccion.Label11.Text = Label8.Text
                Nia_Produccion.Show(Me)
            Case "Nota Salida L."

                nsa_produccion.Close()
                nsa_produccion.TextBox1.Text = "0600"
                nsa_produccion.TextBox13.Text = "0600"
                nsa_produccion.TextBox14.Text = "0601"
                nsa_produccion.TextBox16.Text = "LAVANDERIA"
                nsa_produccion.TextBox11.Text = "LAVANDERIA"
                nsa_produccion.TextBox2.Text = "AREA LAVANDERIA"
                nsa_produccion.Label4.Text = Label10.Text
                nsa_produccion.Label11.Text = Label8.Text
                nsa_produccion.Show(Me)
            Case "Guia Remision L."

                guia_talleres.Label23.Text = "0600"
                guia_talleres.Label25.Text = "LAVANDERIA"
                guia_talleres.TextBox24.Text = "0601"
                guia_talleres.TextBox23.Text = "SERVICIO LAVANDERIA"
                guia_talleres.Label30.Text = Label10.Text
                guia_talleres.Label29.Text = Label8.Text
                guia_talleres.Show(Me)
            Case "Nota Ingreso A."

                Nia_Produccion.TextBox1.Text = "0500"
                Nia_Produccion.TextBox13.Text = "0500"
                Nia_Produccion.TextBox14.Text = "0522"
                Nia_Produccion.TextBox16.Text = "ACABADO"
                Nia_Produccion.TextBox11.Text = "ACABADO"
                Nia_Produccion.TextBox2.Text = "AREA ACABADO"
                Nia_Produccion.Label4.Text = Label10.Text
                Nia_Produccion.Label11.Text = Label8.Text
                Nia_Produccion.Show(Me)
            Case "Nota Salida A."

                nsa_produccion.Close()
                nsa_produccion.TextBox1.Text = "0500"
                nsa_produccion.TextBox13.Text = "0500"
                nsa_produccion.TextBox14.Text = "0522"
                nsa_produccion.TextBox16.Text = "ACABADO"
                nsa_produccion.TextBox11.Text = "ACABADO"
                nsa_produccion.TextBox2.Text = "AREA ACABADO"
                nsa_produccion.Label4.Text = Label10.Text
                nsa_produccion.Label11.Text = Label8.Text
                nsa_produccion.Show(Me)
            Case "Guia Remision A."

                guia_talleres.Label23.Text = "0500"
                guia_talleres.Label25.Text = "ACABADO"
                guia_talleres.TextBox24.Text = "0522"
                guia_talleres.TextBox23.Text = "SERVICIO ACABADO"
                guia_talleres.Label30.Text = Label10.Text
                guia_talleres.Label29.Text = Label8.Text
                guia_talleres.Show(Me)
            Case "Asig. Confeccion"
                Asignacion_Confeccion.Show(Me)
            Case "Asig. Acabados"

                Asignacion_Acabados.Show(Me)
            Case "Liquidacion de Prendas"

                Liqui_Produccion.Label15.Text = Label10.Text
                Liqui_Produccion.Show(Me)
            Case "Solicitudes"

                Solicitud_telas.Show(Me)
            Case "Op"

                Rppt_op.Label3.Text = Label10.Text
                Rppt_op.TextBox2.Text = Label9.Text
                Rppt_op.Show(Me)
            Case "Analisis Movimiento Pro."
                Analisis_movimientos.Label4.Text = Label10.Text
                Analisis_movimientos.Show(Me)
            Case "Produccion Diaria Corte"
                Produccion_diaria_cortee.Label3.Text = Label10.Text
                Produccion_diaria_cortee.Show(Me)
            Case "Planilla Paqueteo"
                PLANILLA_PAQUETEO.Label2.Text = Label10.Text
                PLANILLA_PAQUETEO.TextBox1.Select()
                PLANILLA_PAQUETEO.Show(Me)
            Case "Status Produccion"
                Form7.TextBox4.Text = Label10.Text
                Form7.TextBox1.Text = Label9.Text
                Form7.Show(Me)
            Case "Analisis Op"
                Analisis_Op.TextBox2.Text = Label9.Text
                Analisis_Op.TextBox3.Text = Label10.Text
                Analisis_Op.Show(Me)
            Case "Movimiento Tela-Avios"
                Rpt_ing_sal.Label7.Text = Label10.Text
                Rpt_ing_sal.Show(Me)
            Case "Stock Fisico"
                Stock_Fisico.Label5.Text = Label8.Text
                Stock_Fisico.Label6.Text = Label10.Text
                Stock_Fisico.Show(Me)
            Case "Rpt Orden Compra"
                Orden_Compra_Ac.Label3.Text = Label10.Text
                Orden_Compra_Ac.Show(Me)
            Case "Rpt Orden Servicio"
                ORDE_SERVISO_AC.Label3.Text = Label10.Text
                ORDE_SERVISO_AC.Show(Me)
            Case "Parte Ingreso-Salida"
                Ingreso_salida_Produ.Label6.Text = Label10.Text
                Ingreso_salida_Produ.Show(Me)
            Case "Orden Servicio"
                os.Label16.Text = Label10.Text
                os.Label17.Text = Label8.Text
                os.Show(Me)
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
            Case "Hoja Cotizacion"
                HojaContizacionAlm.Show(Me)
            Case "Consumo de Tela"
                Consumo_tela_Consolidado.Show(Me)
            Case "Actualizar Fecha Despacho"
                Actualizar_FechaDespacho.Label9.Text = Label10.Text
                Actualizar_FechaDespacho.Show(Me)
            Case "Status Produccion Graf."
                StatusProduccionPP.TextBox1.Text = Label9.Text
                StatusProduccionPP.Label2.Text = Label10.Text
                StatusProduccionPP.Show(Me)
            Case "Listado de PO x OP"
                OdxPm.TextBox1.Text = Label9.Text
                OdxPm.Label3.Text = Label10.Text
                OdxPm.Label4.Text = "2"
                OdxPm.Label2.Text = "PO"
                OdxPm.Show(Me)
            Case "Matriz Avios Op"
                Matriz_Avios.Label10.Text = Label10.Text
                Matriz_Avios.Label11.Text = Label8.Text
                Matriz_Avios.Show(Me)
            Case "Reporte Matriz Po"

                ReporteMatPo.Label2.Text = Label10.Text
                ReporteMatPo.Show(Me)
            Case "Reporte Orden Compra"
                Orden_Compra_Ac.Label3.Text = Label10.Text
                Orden_Compra_Ac.Show(Me)
            Case "Reporte Estado Od"
                ReporteOD.TextBox1.Text = Label10.Text
                ReporteOD.Show(Me)
            Case "Requerimiento Tela"
                ReqTelaProduccion.Label5.Text = Label10.Text
                ReqTelaProduccion.Label8.Text = Label8.Text
                ReqTelaProduccion.Label9.Text = Label6.Text
                ReqTelaProduccion.Label10.Text = Label7.Text
                ReqTelaProduccion.limpiar()
                ReqTelaProduccion.Show(Me)
            Case "Requerimiento Tela Udp"
                ReqTelaProduccion.Label5.Text = Label10.Text
                ReqTelaProduccion.Label8.Text = Label8.Text
                ReqTelaProduccion.Label9.Text = Label6.Text
                ReqTelaProduccion.Label10.Text = Label7.Text
                ReqTelaProduccion.limpiar()
                ReqTelaProduccion.Show(Me)
            Case "Matriz Consumo Od"
                MatrizAviosOd.Label10.Text = Label10.Text
                MatrizAviosOd.Label11.Text = Label8.Text
                MatrizAviosOd.Show(Me)
            Case "Reporte Consumo Tela"
                ConsultarConsumo.Label3.Text = Label10.Text
                ConsultarConsumo.TextBox1.Text = Label9.Text
                ConsultarConsumo.Show(Me)
            Case "Lista Requerimiento Tela"
                TelaUdp.Show(Me)
            Case "Requerimiento Avios"
                RequerimientoAvios.Label5.Text = Label10.Text
                RequerimientoAvios.Label8.Text = Label8.Text
                RequerimientoAvios.Label9.Text = Label6.Text
                RequerimientoAvios.Label10.Text = Label7.Text
                RequerimientoAvios.Label17.Text = Label6.Text
                RequerimientoAvios.Show(Me)
            Case "Hoja Cotizacion P."
                AlmacenCotizacion.Label3.Text = Label7.Text
                AlmacenCotizacion.Show(Me)
                'HojaContizacionAlm.Label3.Text = Label7.Text
                'HojaContizacionAlm.Show(Me)
            Case "Hoja Cotizacion U."
                AlmacenCotizacion.Label3.Text = Label7.Text
                AlmacenCotizacion.Show(Me)
                'HojaContizacionAlm.Label3.Text = Label7.Text
                'HojaContizacionAlm.Show(Me)
            Case "Guia Remision PT"
                Guia_remision_prenda.Label33.Text = Label10.Text
                Guia_remision_prenda.Label35.Text = "01"
                Guia_remision_prenda.Label36.Text = Label8.Text
                Guia_remision_prenda.Label37.Text = Label6.Text
                Guia_remision_prenda.Label23.Text = "01"
                Guia_remision_prenda.Label25.Text = "ALMACEN PT"
                Guia_remision_prenda.Show(Me)
            Case "Sinonimo Productos"
                FormSinonimo.Label3.Text = Label10.Text
                FormSinonimo.Show(Me)
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
            Case "Codificacion Producto"
                almacen_nuevo.Label2.Text = Label10.Text
                almacen_nuevo.Label3.Text = Label8.Text
                almacen_nuevo.Label4.Text = "1"
                almacen_nuevo.Show(Me)

            Case "Liquidacion Prendas"
                RptLiquidacionPrendas.textBox1.Text = Label9.Text
                RptLiquidacionPrendas.textBox3.Text = Label10.Text
                RptLiquidacionPrendas.textBox2.Text = Label8.Text
                RptLiquidacionPrendas.Show(Me)
            Case "Status Op"
                Pre_produccion.TextBox1.Text = Label8.Text
                Pre_produccion.Label5.Text = Label10.Text
                Pre_produccion.Show(Me)
            Case "Status Encogimiento"
                FormEncogimiento.Label5.Text = Label10.Text
                FormEncogimiento.TextBox1.Text = Label8.Text
                FormEncogimiento.Show(Me)
            Case "Status Op Prod."
                StatusProduccionPP.Label2.Text = Label10.Text
                StatusProduccionPP.Label5.Text = Label8.Text
                StatusProduccionPP.TextBox1.Text = Label9.Text
                StatusProduccionPP.Show(Me)
            Case "Reporte Op Producidas"
                RptOpPrecoso.Show(Me)
        End Select
    End Sub

    Private consultarBaseDatos As Threading.Thread
    Private Sub MDIParent2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim sql As String = "Exec ActualizarCierreOp"
        Dim cmd1 As New SqlCommand(sql, conx)
        Rsr3 = cmd1.ExecuteReader()
        Rsr3.Read()
        Rsr3.Close()
        consultarBaseDatos = New Thread(New ThreadStart(AddressOf LeerNotificacion))
        consultarBaseDatos.Start()
    End Sub
    Dim id As Integer
    Public Sub LeerNotificacion()
        Dim usuariom As String = ""
        Select Case Label6.Text
            Case "MVARA" : usuariom = "APrNot"
            Case "ACORDOVA" : usuariom = "JPrNot"
            Case "BCHAVEZ" : usuariom = "Pr1Not"
            Case "CLIZARRAGA" : usuariom = "Pr2Not"
            Case "ANALISTA" : usuariom = "Pr3Not"
            Case "BBLAS" : usuariom = "An1Not"
            Case "SACOSTA" : usuariom = "An2Not"
            Case "GMIRANDA" : usuariom = "ComNot"
            Case "WROSALES" : usuariom = "JudNot"
        End Select

        While True
            ' Consulta a nuestra base de datos
            Dim consultaNotificacion As String = "select  FecNot,DetNot,ModNot,IdNot from  Notificaciones where " + usuariom + " = 0"
            Dim adapter As New SqlDataAdapter(consultaNotificacion, conx)
            Dim ds As New DataSet("DetNot")
            adapter.FillSchema(ds, SchemaType.Source, "DetNot")
            adapter.Fill(ds, "DetNot")

            If ds.Tables("DetNot").Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables("DetNot").Rows
                    id = 0
                    id = Convert.ToInt32(row("IdNot"))
                    Dim mensaje As String = row("DetNot").ToString()
                    NotifyIcon1.ShowBalloonTip(1000, "Nueva notificación", mensaje, ToolTipIcon.Info)

                    ' Puedes agregar una pausa entre notificaciones si es necesario
                    Threading.Thread.Sleep(6000) ' Espera 2 segundos entre notificaciones
                Next
            End If

            Threading.Thread.Sleep(10000)
        End While
    End Sub


    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub NotifyIcon1_BalloonTipClosed(sender As Object, e As EventArgs) Handles NotifyIcon1.BalloonTipClosed

    End Sub

    Private Sub NotifyIcon1_Click(sender As Object, e As EventArgs) Handles NotifyIcon1.Click

    End Sub

    Private Sub NotifyIcon1_BalloonTipClicked(sender As Object, e As EventArgs) Handles NotifyIcon1.BalloonTipClicked
        Dim usuariom As String = ""
        Select Case Label6.Text
            Case "MVARA" : usuariom = "APrNot"
            Case "ACORDOVA" : usuariom = "JPrNot"
            Case "BCHAVEZ" : usuariom = "Pr1Not"
            Case "CLIZARRAGA" : usuariom = "Pr2Not"
            Case "ANALISTA" : usuariom = "Pr3Not"
            Case "BBLAS" : usuariom = "An1Not"
            Case "SACOSTA" : usuariom = "An2Not"
            Case "GMIRANDA" : usuariom = "ComNot"
            Case "WROSALES" : usuariom = "JudNot"
        End Select
        Dim cmd168 As New SqlCommand("UPDATE Notificaciones SET  " + usuariom + " = 1 where IdNot  = @IdNot", conx)
        cmd168.Parameters.AddWithValue("@IdNot", id)
        cmd168.ExecuteNonQuery()
        MsgBox("Notificación marcada como leída")
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

    End Sub
End Class
