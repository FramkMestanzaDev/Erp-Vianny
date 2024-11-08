<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Fecha__Factura
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
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Cuenta = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RUC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RSOCIAL = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DOC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DOCUMENTO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FECHA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Monto = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FechaR = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CLient = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Honeydew
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Location = New System.Drawing.Point(12, 40)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(119, 23)
        Me.Button2.TabIndex = 53
        Me.Button2.Text = "ELIMINAR DATA"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Location = New System.Drawing.Point(137, 40)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(133, 23)
        Me.Button1.TabIndex = 52
        Me.Button1.Text = "IMPORTAR EXCEL"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Cuenta, Me.RUC, Me.RSOCIAL, Me.DOC, Me.DOCUMENTO, Me.FECHA, Me.Monto, Me.FechaR, Me.CLient})
        Me.DataGridView1.Location = New System.Drawing.Point(9, 69)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 4
        Me.DataGridView1.Size = New System.Drawing.Size(712, 624)
        Me.DataGridView1.TabIndex = 54
        '
        'Cuenta
        '
        Me.Cuenta.HeaderText = "Cuenta"
        Me.Cuenta.Name = "Cuenta"
        '
        'RUC
        '
        Me.RUC.HeaderText = "RUC"
        Me.RUC.Name = "RUC"
        '
        'RSOCIAL
        '
        Me.RSOCIAL.HeaderText = "RSOCIAL"
        Me.RSOCIAL.Name = "RSOCIAL"
        '
        'DOC
        '
        Me.DOC.HeaderText = "TDOC"
        Me.DOC.Name = "DOC"
        '
        'DOCUMENTO
        '
        Me.DOCUMENTO.HeaderText = "DOCUMENTO"
        Me.DOCUMENTO.Name = "DOCUMENTO"
        '
        'FECHA
        '
        Me.FECHA.HeaderText = "FECHA"
        Me.FECHA.Name = "FECHA"
        '
        'Monto
        '
        Me.Monto.HeaderText = "Monto"
        Me.Monto.Name = "Monto"
        '
        'FechaR
        '
        Me.FechaR.HeaderText = "FechaReal"
        Me.FechaR.Name = "FechaR"
        '
        'CLient
        '
        Me.CLient.HeaderText = "Cliente"
        Me.CLient.Name = "CLient"
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button3.Location = New System.Drawing.Point(647, 40)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(74, 23)
        Me.Button3.TabIndex = 55
        Me.Button3.Text = "Registrar "
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.Lavender
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button4.Image = Global.Modulo_Ventas.Resources.EXCELL2
        Me.Button4.Location = New System.Drawing.Point(596, 28)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(45, 35)
        Me.Button4.TabIndex = 56
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(309, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 57
        Me.Label1.Text = "Label1"
        Me.Label1.Visible = False
        '
        'Fecha__Factura
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 705)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Fecha__Factura"
        Me.Text = "Mostrar Fecha Factura"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Cuenta As DataGridViewTextBoxColumn
    Friend WithEvents RUC As DataGridViewTextBoxColumn
    Friend WithEvents RSOCIAL As DataGridViewTextBoxColumn
    Friend WithEvents DOC As DataGridViewTextBoxColumn
    Friend WithEvents DOCUMENTO As DataGridViewTextBoxColumn
    Friend WithEvents FECHA As DataGridViewTextBoxColumn
    Friend WithEvents Monto As DataGridViewTextBoxColumn
    Friend WithEvents FechaR As DataGridViewTextBoxColumn
    Friend WithEvents CLient As DataGridViewTextBoxColumn
    Friend WithEvents Label1 As Label
End Class
