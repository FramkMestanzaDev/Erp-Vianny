<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Despachos_adicionales
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.CHOFER = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.VEHICULO = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.OBSERVACION = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ESTIBA = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.PEDIDO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DES = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CLIENTE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.FINH = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.FECH = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CHO = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.ve = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.OBSE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ESTIB = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.PEDI = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DIR = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CLI = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.u = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.Button3 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.SeaShell
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft JhengHei UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Button1.Location = New System.Drawing.Point(6, 80)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(92, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "AGREGAR"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CHOFER, Me.VEHICULO, Me.OBSERVACION, Me.ESTIBA, Me.PEDIDO, Me.DES, Me.CLIENTE})
        Me.DataGridView1.Location = New System.Drawing.Point(6, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 5
        Me.DataGridView1.Size = New System.Drawing.Size(984, 52)
        Me.DataGridView1.TabIndex = 1
        '
        'CHOFER
        '
        Me.CHOFER.DisplayStyleForCurrentCellOnly = True
        Me.CHOFER.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.CHOFER.HeaderText = "CHOFER"
        Me.CHOFER.Name = "CHOFER"
        Me.CHOFER.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CHOFER.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.CHOFER.Width = 150
        '
        'VEHICULO
        '
        Me.VEHICULO.DisplayStyleForCurrentCellOnly = True
        Me.VEHICULO.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.VEHICULO.HeaderText = "VEHICULO"
        Me.VEHICULO.Name = "VEHICULO"
        Me.VEHICULO.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.VEHICULO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.VEHICULO.Width = 80
        '
        'OBSERVACION
        '
        Me.OBSERVACION.HeaderText = "OBSERVACION"
        Me.OBSERVACION.Name = "OBSERVACION"
        Me.OBSERVACION.Width = 150
        '
        'ESTIBA
        '
        Me.ESTIBA.DisplayStyleForCurrentCellOnly = True
        Me.ESTIBA.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ESTIBA.HeaderText = "ESTIBA"
        Me.ESTIBA.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9"})
        Me.ESTIBA.Name = "ESTIBA"
        Me.ESTIBA.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ESTIBA.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ESTIBA.Width = 55
        '
        'PEDIDO
        '
        Me.PEDIDO.HeaderText = "PEDIDO"
        Me.PEDIDO.Name = "PEDIDO"
        Me.PEDIDO.Width = 80
        '
        'DES
        '
        Me.DES.HeaderText = "DIR. DESPACHO"
        Me.DES.Name = "DES"
        Me.DES.Width = 200
        '
        'CLIENTE
        '
        Me.CLIENTE.HeaderText = "CLIENTE"
        Me.CLIENTE.Name = "CLIENTE"
        Me.CLIENTE.Width = 235
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Beige
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Image = Global.Modulo_Ventas.Resources.guardar22
        Me.Button2.Location = New System.Drawing.Point(6, 618)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(54, 44)
        Me.Button2.TabIndex = 2
        Me.Button2.UseVisualStyleBackColor = False
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.FINH, Me.FECH, Me.CHO, Me.ve, Me.OBSE, Me.ESTIB, Me.PEDI, Me.DIR, Me.CLI, Me.u})
        Me.DataGridView2.Location = New System.Drawing.Point(6, 109)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowHeadersWidth = 5
        Me.DataGridView2.Size = New System.Drawing.Size(984, 503)
        Me.DataGridView2.TabIndex = 3
        '
        'FINH
        '
        Me.FINH.HeaderText = "FIN"
        Me.FINH.Name = "FINH"
        Me.FINH.Width = 30
        '
        'FECH
        '
        Me.FECH.HeaderText = "FECHA"
        Me.FECH.Name = "FECH"
        Me.FECH.Width = 80
        '
        'CHO
        '
        Me.CHO.DisplayStyleForCurrentCellOnly = True
        Me.CHO.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.CHO.HeaderText = "CHOFER"
        Me.CHO.Name = "CHO"
        Me.CHO.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CHO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.CHO.Width = 90
        '
        've
        '
        Me.ve.DisplayStyleForCurrentCellOnly = True
        Me.ve.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ve.HeaderText = "VEHICULO"
        Me.ve.Name = "ve"
        Me.ve.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ve.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ve.Width = 80
        '
        'OBSE
        '
        Me.OBSE.HeaderText = "OBSERVACION"
        Me.OBSE.Name = "OBSE"
        Me.OBSE.Width = 180
        '
        'ESTIB
        '
        Me.ESTIB.DisplayStyleForCurrentCellOnly = True
        Me.ESTIB.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ESTIB.HeaderText = "ESTIBA"
        Me.ESTIB.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9"})
        Me.ESTIB.Name = "ESTIB"
        Me.ESTIB.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ESTIB.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ESTIB.Width = 70
        '
        'PEDI
        '
        Me.PEDI.HeaderText = "PEDIDO"
        Me.PEDI.Name = "PEDI"
        Me.PEDI.Width = 80
        '
        'DIR
        '
        Me.DIR.HeaderText = "DIR DESPACHO"
        Me.DIR.Name = "DIR"
        Me.DIR.Width = 160
        '
        'CLI
        '
        Me.CLI.HeaderText = "CLIENTE"
        Me.CLI.Name = "CLI"
        Me.CLI.Width = 180
        '
        'u
        '
        Me.u.HeaderText = "id"
        Me.u.Name = "u"
        Me.u.Visible = False
        '
        'DataGridView3
        '
        Me.DataGridView3.AllowUserToAddRows = False
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(923, 93)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.Size = New System.Drawing.Size(10, 10)
        Me.DataGridView3.TabIndex = 4
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.Ivory
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Image = Global.Modulo_Ventas.Resources.cerrar2
        Me.Button3.Location = New System.Drawing.Point(66, 618)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(54, 44)
        Me.Button3.TabIndex = 5
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Despachos_adicionales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(999, 672)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.DataGridView3)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Despachos_adicionales"
        Me.Text = "DESPACHOS ADICIONALES"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button2 As Button
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents CHOFER As DataGridViewComboBoxColumn
    Friend WithEvents VEHICULO As DataGridViewComboBoxColumn
    Friend WithEvents OBSERVACION As DataGridViewTextBoxColumn
    Friend WithEvents ESTIBA As DataGridViewComboBoxColumn
    Friend WithEvents PEDIDO As DataGridViewTextBoxColumn
    Friend WithEvents DES As DataGridViewTextBoxColumn
    Friend WithEvents CLIENTE As DataGridViewTextBoxColumn
    Friend WithEvents DataGridView3 As DataGridView
    Friend WithEvents FINH As DataGridViewCheckBoxColumn
    Friend WithEvents FECH As DataGridViewTextBoxColumn
    Friend WithEvents CHO As DataGridViewComboBoxColumn
    Friend WithEvents ve As DataGridViewComboBoxColumn
    Friend WithEvents OBSE As DataGridViewTextBoxColumn
    Friend WithEvents ESTIB As DataGridViewComboBoxColumn
    Friend WithEvents PEDI As DataGridViewTextBoxColumn
    Friend WithEvents DIR As DataGridViewTextBoxColumn
    Friend WithEvents CLI As DataGridViewTextBoxColumn
    Friend WithEvents u As DataGridViewTextBoxColumn
    Friend WithEvents Button3 As Button
End Class
