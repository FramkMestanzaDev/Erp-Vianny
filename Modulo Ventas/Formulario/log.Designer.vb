<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class log
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
        Me.Label3 = New System.Windows.Forms.Label()
        Me.FH = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Modulo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Usu = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Pc = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.accion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Acci = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.detalle = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.FH, Me.Modulo, Me.Usu, Me.Pc, Me.accion, Me.Acci, Me.detalle})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 27)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidth = 4
        Me.DataGridView1.Size = New System.Drawing.Size(847, 387)
        Me.DataGridView1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(163, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        Me.Label1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(429, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Label2"
        Me.Label2.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(675, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Label3"
        Me.Label3.Visible = False
        '
        'FH
        '
        Me.FH.HeaderText = "Fecha y Hora"
        Me.FH.Name = "FH"
        Me.FH.ReadOnly = True
        Me.FH.Width = 130
        '
        'Modulo
        '
        Me.Modulo.HeaderText = "Modulo"
        Me.Modulo.Name = "Modulo"
        Me.Modulo.ReadOnly = True
        Me.Modulo.Width = 140
        '
        'Usu
        '
        Me.Usu.HeaderText = "Usuario"
        Me.Usu.Name = "Usu"
        Me.Usu.ReadOnly = True
        '
        'Pc
        '
        Me.Pc.HeaderText = "Pc"
        Me.Pc.Name = "Pc"
        Me.Pc.ReadOnly = True
        Me.Pc.Width = 120
        '
        'accion
        '
        Me.accion.HeaderText = "Ac"
        Me.accion.Name = "accion"
        Me.accion.ReadOnly = True
        Me.accion.Width = 30
        '
        'Acci
        '
        Me.Acci.HeaderText = "Accion"
        Me.Acci.Name = "Acci"
        Me.Acci.ReadOnly = True
        Me.Acci.Width = 170
        '
        'detalle
        '
        Me.detalle.HeaderText = "Detalle"
        Me.detalle.Name = "detalle"
        Me.detalle.ReadOnly = True
        Me.detalle.Width = 140
        '
        'log
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(871, 424)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "log"
        Me.Text = "REGISTRO DE EVENTOS"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents FH As DataGridViewTextBoxColumn
    Friend WithEvents Modulo As DataGridViewTextBoxColumn
    Friend WithEvents Usu As DataGridViewTextBoxColumn
    Friend WithEvents Pc As DataGridViewTextBoxColumn
    Friend WithEvents accion As DataGridViewTextBoxColumn
    Friend WithEvents Acci As DataGridViewTextBoxColumn
    Friend WithEvents detalle As DataGridViewTextBoxColumn
End Class
