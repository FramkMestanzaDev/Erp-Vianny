<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GandProduccion
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
        Me.Corte = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FCorte = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Aplicaciones = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FAplicaciones = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Confeccion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FConfeccion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Lavanderia = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FLavanderia = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Acabados = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FAcabados = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Despacho = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FDespacho = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Corte, Me.FCorte, Me.Aplicaciones, Me.FAplicaciones, Me.Confeccion, Me.FConfeccion, Me.Lavanderia, Me.FLavanderia, Me.Acabados, Me.FAcabados, Me.Despacho, Me.FDespacho})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 45)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 4
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(1260, 529)
        Me.DataGridView1.TabIndex = 0
        '
        'Corte
        '
        Me.Corte.HeaderText = "Corte"
        Me.Corte.Name = "Corte"
        '
        'FCorte
        '
        Me.FCorte.HeaderText = "FCorte"
        Me.FCorte.Name = "FCorte"
        Me.FCorte.Width = 80
        '
        'Aplicaciones
        '
        Me.Aplicaciones.HeaderText = "Aplicaciones"
        Me.Aplicaciones.Name = "Aplicaciones"
        '
        'FAplicaciones
        '
        Me.FAplicaciones.HeaderText = "FAplicaciones"
        Me.FAplicaciones.Name = "FAplicaciones"
        Me.FAplicaciones.Width = 80
        '
        'Confeccion
        '
        Me.Confeccion.HeaderText = "Confeccion"
        Me.Confeccion.Name = "Confeccion"
        Me.Confeccion.Width = 80
        '
        'FConfeccion
        '
        Me.FConfeccion.HeaderText = "FConfeccion"
        Me.FConfeccion.Name = "FConfeccion"
        '
        'Lavanderia
        '
        Me.Lavanderia.HeaderText = "Lavanderia"
        Me.Lavanderia.Name = "Lavanderia"
        Me.Lavanderia.Width = 80
        '
        'FLavanderia
        '
        Me.FLavanderia.HeaderText = "FLavanderia"
        Me.FLavanderia.Name = "FLavanderia"
        '
        'Acabados
        '
        Me.Acabados.HeaderText = "Acabados"
        Me.Acabados.Name = "Acabados"
        Me.Acabados.Width = 80
        '
        'FAcabados
        '
        Me.FAcabados.HeaderText = "FAcabados"
        Me.FAcabados.Name = "FAcabados"
        '
        'Despacho
        '
        Me.Despacho.HeaderText = "Despacho"
        Me.Despacho.Name = "Despacho"
        Me.Despacho.Width = 80
        '
        'FDespacho
        '
        Me.FDespacho.HeaderText = "FDespacho"
        Me.FDespacho.Name = "FDespacho"
        '
        'GandProduccion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1309, 586)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "GandProduccion"
        Me.Text = "Gand Produccion"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Corte As DataGridViewTextBoxColumn
    Friend WithEvents FCorte As DataGridViewTextBoxColumn
    Friend WithEvents Aplicaciones As DataGridViewTextBoxColumn
    Friend WithEvents FAplicaciones As DataGridViewTextBoxColumn
    Friend WithEvents Confeccion As DataGridViewTextBoxColumn
    Friend WithEvents FConfeccion As DataGridViewTextBoxColumn
    Friend WithEvents Lavanderia As DataGridViewTextBoxColumn
    Friend WithEvents FLavanderia As DataGridViewTextBoxColumn
    Friend WithEvents Acabados As DataGridViewTextBoxColumn
    Friend WithEvents FAcabados As DataGridViewTextBoxColumn
    Friend WithEvents Despacho As DataGridViewTextBoxColumn
    Friend WithEvents FDespacho As DataGridViewTextBoxColumn
End Class
