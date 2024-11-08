<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TELAS_OP
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.A = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.CODIGO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DESCRIPCION = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NOMBRE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.D = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.G = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SISTEMAS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BW = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CLIENTE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LM1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LM2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LM3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.A, Me.CODIGO, Me.DESCRIPCION, Me.NOMBRE, Me.D, Me.G, Me.SISTEMAS, Me.BW, Me.BD, Me.BT, Me.AA, Me.AD, Me.AT, Me.CLIENTE, Me.LM1, Me.LM2, Me.LM3})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 37)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 4
        Me.DataGridView1.Size = New System.Drawing.Size(966, 196)
        Me.DataGridView1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(327, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(127, 25)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "TELAS OP :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Maroon
        Me.Label2.Location = New System.Drawing.Point(451, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(83, 25)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Label2"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FloralWhite
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft JhengHei UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(12, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "AGREGAR"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'A
        '
        Me.A.Frozen = True
        Me.A.HeaderText = "A"
        Me.A.Name = "A"
        Me.A.Width = 25
        '
        'CODIGO
        '
        Me.CODIGO.Frozen = True
        Me.CODIGO.HeaderText = "CODIGO"
        Me.CODIGO.Name = "CODIGO"
        Me.CODIGO.ReadOnly = True
        Me.CODIGO.Width = 145
        '
        'DESCRIPCION
        '
        Me.DESCRIPCION.Frozen = True
        Me.DESCRIPCION.HeaderText = "DESCRIPCION"
        Me.DESCRIPCION.Name = "DESCRIPCION"
        Me.DESCRIPCION.ReadOnly = True
        Me.DESCRIPCION.Width = 250
        '
        'NOMBRE
        '
        Me.NOMBRE.Frozen = True
        Me.NOMBRE.HeaderText = "NONBRE COMERCIAL"
        Me.NOMBRE.Name = "NOMBRE"
        Me.NOMBRE.ReadOnly = True
        Me.NOMBRE.Width = 150
        '
        'D
        '
        Me.D.HeaderText = "D"
        Me.D.Name = "D"
        Me.D.ReadOnly = True
        Me.D.Width = 60
        '
        'G
        '
        Me.G.HeaderText = "G"
        Me.G.Name = "G"
        Me.G.ReadOnly = True
        Me.G.Width = 60
        '
        'SISTEMAS
        '
        Me.SISTEMAS.HeaderText = "SISTEMAS"
        Me.SISTEMAS.Name = "SISTEMAS"
        Me.SISTEMAS.ReadOnly = True
        Me.SISTEMAS.Width = 70
        '
        'BW
        '
        Me.BW.HeaderText = "B/W ANCHO"
        Me.BW.Name = "BW"
        Me.BW.ReadOnly = True
        Me.BW.Width = 75
        '
        'BD
        '
        Me.BD.HeaderText = "B/W DENSIDAD"
        Me.BD.Name = "BD"
        Me.BD.ReadOnly = True
        Me.BD.Width = 75
        '
        'BT
        '
        Me.BT.HeaderText = "B/W TIPO"
        Me.BT.Name = "BT"
        Me.BT.ReadOnly = True
        Me.BT.Width = 75
        '
        'AA
        '
        Me.AA.HeaderText = "A/W ANCHO"
        Me.AA.Name = "AA"
        Me.AA.ReadOnly = True
        Me.AA.Width = 75
        '
        'AD
        '
        Me.AD.HeaderText = "A/W DENSIDAD"
        Me.AD.Name = "AD"
        Me.AD.ReadOnly = True
        Me.AD.Width = 75
        '
        'AT
        '
        Me.AT.HeaderText = "A/W TIPO"
        Me.AT.Name = "AT"
        Me.AT.ReadOnly = True
        Me.AT.Width = 75
        '
        'CLIENTE
        '
        Me.CLIENTE.HeaderText = "CLIENTE"
        Me.CLIENTE.Name = "CLIENTE"
        Me.CLIENTE.ReadOnly = True
        Me.CLIENTE.Width = 130
        '
        'LM1
        '
        Me.LM1.HeaderText = "LM-1 CLARO"
        Me.LM1.Name = "LM1"
        Me.LM1.ReadOnly = True
        Me.LM1.Width = 75
        '
        'LM2
        '
        Me.LM2.HeaderText = "LM-2 MEDIO"
        Me.LM2.Name = "LM2"
        Me.LM2.ReadOnly = True
        Me.LM2.Width = 75
        '
        'LM3
        '
        Me.LM3.HeaderText = "LM-1 OSCURO"
        Me.LM3.Name = "LM3"
        Me.LM3.ReadOnly = True
        Me.LM3.Width = 75
        '
        'TELAS_OP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(990, 239)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "TELAS_OP"
        Me.Text = "TELAS OP"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents A As DataGridViewCheckBoxColumn
    Friend WithEvents CODIGO As DataGridViewTextBoxColumn
    Friend WithEvents DESCRIPCION As DataGridViewTextBoxColumn
    Friend WithEvents NOMBRE As DataGridViewTextBoxColumn
    Friend WithEvents D As DataGridViewTextBoxColumn
    Friend WithEvents G As DataGridViewTextBoxColumn
    Friend WithEvents SISTEMAS As DataGridViewTextBoxColumn
    Friend WithEvents BW As DataGridViewTextBoxColumn
    Friend WithEvents BD As DataGridViewTextBoxColumn
    Friend WithEvents BT As DataGridViewTextBoxColumn
    Friend WithEvents AA As DataGridViewTextBoxColumn
    Friend WithEvents AD As DataGridViewTextBoxColumn
    Friend WithEvents AT As DataGridViewTextBoxColumn
    Friend WithEvents CLIENTE As DataGridViewTextBoxColumn
    Friend WithEvents LM1 As DataGridViewTextBoxColumn
    Friend WithEvents LM2 As DataGridViewTextBoxColumn
    Friend WithEvents LM3 As DataGridViewTextBoxColumn
End Class
