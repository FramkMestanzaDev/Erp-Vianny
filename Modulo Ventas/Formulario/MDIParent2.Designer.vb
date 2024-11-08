<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MDIParent2
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub


    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Hoja Cotizacion P.")
        Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Status Op Prod.")
        Dim TreeNode3 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Planeamiento Op")
        Dim TreeNode4 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Requerimiento Tela")
        Dim TreeNode5 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Requerimiento Avios")
        Dim TreeNode6 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Ingreso C.")
        Dim TreeNode7 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Salida C.")
        Dim TreeNode8 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Guia Remision C.")
        Dim TreeNode9 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Corte", New System.Windows.Forms.TreeNode() {TreeNode6, TreeNode7, TreeNode8})
        Dim TreeNode10 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Ingreso Ap.")
        Dim TreeNode11 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Salida Ap.")
        Dim TreeNode12 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Guia Remision Ap.")
        Dim TreeNode13 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Aplicaciones y Piezas", New System.Windows.Forms.TreeNode() {TreeNode10, TreeNode11, TreeNode12})
        Dim TreeNode14 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Ingreso Co.")
        Dim TreeNode15 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Salida Co.")
        Dim TreeNode16 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Guia Remision Co.")
        Dim TreeNode17 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Confeccion", New System.Windows.Forms.TreeNode() {TreeNode14, TreeNode15, TreeNode16})
        Dim TreeNode18 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Ingreso L.")
        Dim TreeNode19 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Salida L.")
        Dim TreeNode20 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Guia Remision L.")
        Dim TreeNode21 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Lavanderia", New System.Windows.Forms.TreeNode() {TreeNode18, TreeNode19, TreeNode20})
        Dim TreeNode22 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Ingreso A.")
        Dim TreeNode23 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Salida A.")
        Dim TreeNode24 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Guia Remision A.")
        Dim TreeNode25 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Acabado", New System.Windows.Forms.TreeNode() {TreeNode22, TreeNode23, TreeNode24})
        Dim TreeNode26 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Sinonimo Productos")
        Dim TreeNode27 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Ingreso PT")
        Dim TreeNode28 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nota Salida PT")
        Dim TreeNode29 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Guia Remision PT")
        Dim TreeNode30 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Prendas", New System.Windows.Forms.TreeNode() {TreeNode26, TreeNode27, TreeNode28, TreeNode29})
        Dim TreeNode31 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Asig. Confeccion")
        Dim TreeNode32 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Asig. Acabados")
        Dim TreeNode33 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Asignacion ", New System.Windows.Forms.TreeNode() {TreeNode31, TreeNode32})
        Dim TreeNode34 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Liquidacion de Prendas")
        Dim TreeNode35 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Orden Servicio")
        Dim TreeNode36 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Sticker Calidad")
        Dim TreeNode37 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Actualizar Fecha Despacho")
        Dim TreeNode38 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Informacion", New System.Windows.Forms.TreeNode() {TreeNode1, TreeNode2, TreeNode3, TreeNode4, TreeNode5, TreeNode9, TreeNode13, TreeNode17, TreeNode21, TreeNode25, TreeNode30, TreeNode33, TreeNode34, TreeNode35, TreeNode36, TreeNode37})
        Dim TreeNode39 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Op")
        Dim TreeNode40 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Analisis Movimiento Pro.")
        Dim TreeNode41 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Produccion Diaria Corte")
        Dim TreeNode42 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Planilla Paqueteo")
        Dim TreeNode43 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Status Produccion")
        Dim TreeNode44 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Analisis Op")
        Dim TreeNode45 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Movimiento Tela-Avios")
        Dim TreeNode46 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Stock Fisico")
        Dim TreeNode47 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Rpt Orden Compra")
        Dim TreeNode48 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Rpt Orden Servicio")
        Dim TreeNode49 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Parte Ingreso-Salida")
        Dim TreeNode50 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Status Produccion Graf.")
        Dim TreeNode51 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Liquidacion Prendas")
        Dim TreeNode52 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Listado de PO x OP")
        Dim TreeNode53 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Reportes", New System.Windows.Forms.TreeNode() {TreeNode39, TreeNode40, TreeNode41, TreeNode42, TreeNode43, TreeNode44, TreeNode45, TreeNode46, TreeNode47, TreeNode48, TreeNode49, TreeNode50, TreeNode51, TreeNode52})
        Dim TreeNode54 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Produccion", New System.Windows.Forms.TreeNode() {TreeNode38, TreeNode53})
        Dim TreeNode55 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Codificacion Producto")
        Dim TreeNode56 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Hoja Cotizacion U.")
        Dim TreeNode57 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Status Od")
        Dim TreeNode58 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Status Op")
        Dim TreeNode59 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Status Encogimiento")
        Dim TreeNode60 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Ficha Tecnica")
        Dim TreeNode61 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Matriz Avios Op")
        Dim TreeNode62 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Asignacion Molde")
        Dim TreeNode63 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Consumo Op")
        Dim TreeNode64 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Orden Produccion")
        Dim TreeNode65 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Requerimiento Tela Udp")
        Dim TreeNode66 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Informacion", New System.Windows.Forms.TreeNode() {TreeNode55, TreeNode56, TreeNode57, TreeNode58, TreeNode59, TreeNode60, TreeNode61, TreeNode62, TreeNode63, TreeNode64, TreeNode65})
        Dim TreeNode67 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Listado de PM x OD")
        Dim TreeNode68 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Listado de PO x OP")
        Dim TreeNode69 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Planeamiento Udp")
        Dim TreeNode70 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Consumo de Tela")
        Dim TreeNode71 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Reporte Matriz Po")
        Dim TreeNode72 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Reporte Orden Compra")
        Dim TreeNode73 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Reporte Estado Od")
        Dim TreeNode74 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Reporte Consumo Tela")
        Dim TreeNode75 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Lista Requerimiento Tela")
        Dim TreeNode76 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Reporte Op Producidas")
        Dim TreeNode77 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Reportes", New System.Windows.Forms.TreeNode() {TreeNode67, TreeNode68, TreeNode69, TreeNode70, TreeNode71, TreeNode72, TreeNode73, TreeNode74, TreeNode75, TreeNode76})
        Dim TreeNode78 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Udp", New System.Windows.Forms.TreeNode() {TreeNode66, TreeNode77})
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDIParent2))
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.Controls.Add(Me.DataGridView1)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.TreeView1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(279, 748)
        Me.Panel1.TabIndex = 9
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(241, 716)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(14, 10)
        Me.DataGridView1.TabIndex = 12
        Me.DataGridView1.Visible = False
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Rockwell", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(176, 716)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(54, 16)
        Me.Label10.TabIndex = 11
        Me.Label10.Text = "Label10"
        Me.Label10.Visible = False
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Rockwell", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(96, 716)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(47, 16)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "Label9"
        '
        'Label8
        '
        Me.Label8.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Rockwell", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(96, 690)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(47, 16)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Label8"
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Rockwell", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(96, 665)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(47, 16)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Label7"
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Rockwell", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(96, 636)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(47, 16)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Label6"
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Rockwell", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Maroon
        Me.Label5.Location = New System.Drawing.Point(17, 690)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(71, 16)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "PERIODO :"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Rockwell", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Maroon
        Me.Label4.Location = New System.Drawing.Point(14, 716)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 16)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "EMPRESA :"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Rockwell", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Maroon
        Me.Label3.Location = New System.Drawing.Point(31, 665)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "GRUPO :"
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Rockwell", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Maroon
        Me.Label2.Location = New System.Drawing.Point(19, 636)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "USUARIO :"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Info
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(250, 45)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Rockwell", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(4, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(239, 19)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "MODULO UDP - PRODUCCION"
        '
        'TreeView1
        '
        Me.TreeView1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TreeView1.BackColor = System.Drawing.SystemColors.HighlightText
        Me.TreeView1.Font = New System.Drawing.Font("Rockwell", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeView1.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.TreeView1.Location = New System.Drawing.Point(12, 76)
        Me.TreeView1.Name = "TreeView1"
        TreeNode1.Name = "Hoja Cotizacion P."
        TreeNode1.Text = "Hoja Cotizacion P."
        TreeNode2.Name = "Status Op Prod."
        TreeNode2.Text = "Status Op Prod."
        TreeNode3.Name = "Nodo0"
        TreeNode3.Text = "Planeamiento Op"
        TreeNode4.Name = "Requerimiento Tela"
        TreeNode4.Text = "Requerimiento Tela"
        TreeNode5.Name = "Requerimiento Avios"
        TreeNode5.Text = "Requerimiento Avios"
        TreeNode6.Name = "Nodo0"
        TreeNode6.Text = "Nota Ingreso C."
        TreeNode7.Name = "Nodo1"
        TreeNode7.Text = "Nota Salida C."
        TreeNode8.Name = "Nodo2"
        TreeNode8.Text = "Guia Remision C."
        TreeNode9.Name = "Nodo3"
        TreeNode9.Text = "Corte"
        TreeNode10.Name = "Nodo3"
        TreeNode10.Text = "Nota Ingreso Ap."
        TreeNode11.Name = "Nodo4"
        TreeNode11.Text = "Nota Salida Ap."
        TreeNode12.Name = "Nodo5"
        TreeNode12.Text = "Guia Remision Ap."
        TreeNode13.Name = "cobranzaxpo"
        TreeNode13.Text = "Aplicaciones y Piezas"
        TreeNode14.Name = "Nodo6"
        TreeNode14.Text = "Nota Ingreso Co."
        TreeNode15.Name = "Nodo7"
        TreeNode15.Text = "Nota Salida Co."
        TreeNode16.Name = "Nodo8"
        TreeNode16.Text = "Guia Remision Co."
        TreeNode17.Name = "cobranzaxmes"
        TreeNode17.Text = "Confeccion"
        TreeNode18.Name = "Nodo9"
        TreeNode18.Text = "Nota Ingreso L."
        TreeNode19.Name = "Nodo10"
        TreeNode19.Text = "Nota Salida L."
        TreeNode20.Name = "Nodo11"
        TreeNode20.Text = "Guia Remision L."
        TreeNode21.Name = "Nodo0"
        TreeNode21.Text = "Lavanderia"
        TreeNode22.Name = "Nodo12"
        TreeNode22.Text = "Nota Ingreso A."
        TreeNode23.Name = "Nodo13"
        TreeNode23.Text = "Nota Salida A."
        TreeNode24.Name = "Nodo14"
        TreeNode24.Text = "Guia Remision A."
        TreeNode25.Name = "Nodo0"
        TreeNode25.Text = "Acabado"
        TreeNode26.Name = "Sinonimo Productos"
        TreeNode26.Text = "Sinonimo Productos"
        TreeNode27.Name = "Nota Ingreso PT"
        TreeNode27.Text = "Nota Ingreso PT"
        TreeNode28.Name = "Nota Salida PT"
        TreeNode28.Text = "Nota Salida PT"
        TreeNode29.Name = "Guia Remision PT"
        TreeNode29.Text = "Guia Remision PT"
        TreeNode30.Name = "Prendas"
        TreeNode30.Text = "Prendas"
        TreeNode31.Name = "Nodo15"
        TreeNode31.Text = "Asig. Confeccion"
        TreeNode32.Name = "Nodo16"
        TreeNode32.Text = "Asig. Acabados"
        TreeNode33.Name = "Nodo3"
        TreeNode33.Text = "Asignacion "
        TreeNode34.Name = "Nodo2"
        TreeNode34.Text = "Liquidacion de Prendas"
        TreeNode35.Name = "Nodo0"
        TreeNode35.Text = "Orden Servicio"
        TreeNode36.Name = "Nodo1"
        TreeNode36.Text = "Sticker Calidad"
        TreeNode37.Name = "Actualizar Fecha Despacho"
        TreeNode37.Text = "Actualizar Fecha Despacho"
        TreeNode38.Name = "Nodo4"
        TreeNode38.Text = "Informacion"
        TreeNode39.Name = "Nodo20"
        TreeNode39.Text = "Op"
        TreeNode40.Name = "Nodo22"
        TreeNode40.Text = "Analisis Movimiento Pro."
        TreeNode41.Name = "Nodo23"
        TreeNode41.Text = "Produccion Diaria Corte"
        TreeNode42.Name = "Nodo24"
        TreeNode42.Text = "Planilla Paqueteo"
        TreeNode43.Name = "Nodo25"
        TreeNode43.Text = "Status Produccion"
        TreeNode44.Name = "Nodo26"
        TreeNode44.Text = "Analisis Op"
        TreeNode45.Name = "Nodo27"
        TreeNode45.Text = "Movimiento Tela-Avios"
        TreeNode46.Name = "Nodo28"
        TreeNode46.Text = "Stock Fisico"
        TreeNode47.Name = "Nodo29"
        TreeNode47.Text = "Rpt Orden Compra"
        TreeNode48.Name = "Nodo30"
        TreeNode48.Text = "Rpt Orden Servicio"
        TreeNode49.Name = "Nodo32"
        TreeNode49.Text = "Parte Ingreso-Salida"
        TreeNode50.Name = "Status Produccion Graf."
        TreeNode50.Text = "Status Produccion Graf."
        TreeNode51.Name = "Liquidacion Prendas"
        TreeNode51.Text = "Liquidacion Prendas"
        TreeNode52.Name = "Listado de PO x OP"
        TreeNode52.Text = "Listado de PO x OP"
        TreeNode53.Name = "Nodo5"
        TreeNode53.Text = "Reportes"
        TreeNode54.BackColor = System.Drawing.Color.White
        TreeNode54.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        TreeNode54.Name = "Nodo2"
        TreeNode54.SelectedImageIndex = -2
        TreeNode54.Text = "Produccion"
        TreeNode55.Name = "Codificacion Producto"
        TreeNode55.Text = "Codificacion Producto"
        TreeNode56.Name = "Hoja Cotizacion U."
        TreeNode56.Text = "Hoja Cotizacion U."
        TreeNode57.Name = "Status Od"
        TreeNode57.Text = "Status Od"
        TreeNode58.Name = "Status Op"
        TreeNode58.Text = "Status Op"
        TreeNode59.Name = "Status Encogimiento"
        TreeNode59.Text = "Status Encogimiento"
        TreeNode60.Name = "Nodo2"
        TreeNode60.Text = "Ficha Tecnica"
        TreeNode61.Name = "Matriz Avios Op"
        TreeNode61.Text = "Matriz Avios Op"
        TreeNode62.Name = "Nodo5"
        TreeNode62.Text = "Asignacion Molde"
        TreeNode63.Name = "Nodo3"
        TreeNode63.Text = "Consumo Op"
        TreeNode64.Name = "Nodo7"
        TreeNode64.Text = "Orden Produccion"
        TreeNode65.Name = "Requerimiento Tela"
        TreeNode65.Text = "Requerimiento Tela Udp"
        TreeNode66.Name = "Nodo8"
        TreeNode66.Text = "Informacion"
        TreeNode67.Name = "Nodo10"
        TreeNode67.Text = "Listado de PM x OD"
        TreeNode68.Name = "Listado de PO x OP"
        TreeNode68.Text = "Listado de PO x OP"
        TreeNode69.Name = "Nodo11"
        TreeNode69.Text = "Planeamiento Udp"
        TreeNode70.Name = "Consumo de Tela"
        TreeNode70.Text = "Consumo de Tela"
        TreeNode71.Name = "Reporte Matriz Po"
        TreeNode71.Text = "Reporte Matriz Po"
        TreeNode72.Name = "Reporte Orden Compra"
        TreeNode72.Text = "Reporte Orden Compra"
        TreeNode73.Name = "Reporte Estado Od"
        TreeNode73.Text = "Reporte Estado Od"
        TreeNode74.Name = "Reporte Consumo Tela"
        TreeNode74.Text = "Reporte Consumo Tela"
        TreeNode75.Name = "Lista Requerimiento Tela"
        TreeNode75.Text = "Lista Requerimiento Tela"
        TreeNode76.Name = "Reporte Op Producidas"
        TreeNode76.Text = "Reporte Op Producidas"
        TreeNode77.Name = "Nodo9"
        TreeNode77.Text = "Reportes"
        TreeNode78.Name = "Udp"
        TreeNode78.Text = "Udp"
        Me.TreeView1.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode54, TreeNode78})
        Me.TreeView1.Size = New System.Drawing.Size(263, 543)
        Me.TreeView1.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.AliceBlue
        Me.Panel2.BackgroundImage = Global.Modulo_Ventas.Resources.LOGO_DG_PNG1
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(279, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(873, 748)
        Me.Panel2.TabIndex = 10
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "NotifyIcon1"
        Me.NotifyIcon1.Visible = True
        '
        'MDIParent2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(1152, 748)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.IsMdiContainer = True
        Me.Name = "MDIParent2"
        Me.Text = "UDP - PRODUCCION CONSORCIO TEXTIL VIANNY"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents Panel1 As Panel
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TreeView1 As TreeView
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents NotifyIcon1 As NotifyIcon
End Class
