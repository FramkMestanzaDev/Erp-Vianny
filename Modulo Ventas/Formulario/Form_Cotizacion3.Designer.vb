<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Cotizacion3
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.det = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CON = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.UP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COS_T = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CU = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.SystemColors.Info
        Me.Button3.Font = New System.Drawing.Font("Microsoft JhengHei UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Button3.Location = New System.Drawing.Point(12, 245)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(96, 24)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Guardar"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.det, Me.CON, Me.UP, Me.COS_T, Me.CU, Me.CT})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 12)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 4
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(715, 227)
        Me.DataGridView1.TabIndex = 5
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(146, 248)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(47, 20)
        Me.TextBox1.TabIndex = 6
        Me.TextBox1.Visible = False
        '
        'det
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Info
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft JhengHei UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.det.DefaultCellStyle = DataGridViewCellStyle1
        Me.det.HeaderText = "Descripcion"
        Me.det.Name = "det"
        Me.det.ReadOnly = True
        '
        'CON
        '
        Me.CON.HeaderText = "CONSUMO"
        Me.CON.Name = "CON"
        '
        'UP
        '
        Me.UP.HeaderText = "UNIDAD"
        Me.UP.Name = "UP"
        '
        'COS_T
        '
        Me.COS_T.HeaderText = "CONSUMO_T"
        Me.COS_T.Name = "COS_T"
        Me.COS_T.ReadOnly = True
        '
        'CU
        '
        Me.CU.HeaderText = "C.UNITARIO"
        Me.CU.Name = "CU"
        '
        'CT
        '
        Me.CT.HeaderText = "COS_TOTAL"
        Me.CT.Name = "CT"
        Me.CT.ReadOnly = True
        '
        'Form_Cotizacion3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(736, 281)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button3)
        Me.Name = "Form_Cotizacion3"
        Me.Text = "Otros"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button3 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents det As DataGridViewTextBoxColumn
    Friend WithEvents CON As DataGridViewTextBoxColumn
    Friend WithEvents UP As DataGridViewTextBoxColumn
    Friend WithEvents COS_T As DataGridViewTextBoxColumn
    Friend WithEvents CU As DataGridViewTextBoxColumn
    Friend WithEvents CT As DataGridViewTextBoxColumn
End Class
