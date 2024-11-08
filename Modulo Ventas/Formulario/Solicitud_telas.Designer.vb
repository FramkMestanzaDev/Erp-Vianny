<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Solicitud_telas
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.PO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CODIGO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TELA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PROGRA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CANS_L = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CONS_KGS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CANTID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PO, Me.OP, Me.CODIGO, Me.TELA, Me.PROGRA, Me.CANS_L, Me.CONS_KGS, Me.CANTID})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 60)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 5
        Me.DataGridView1.Size = New System.Drawing.Size(968, 594)
        Me.DataGridView1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft JhengHei UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(12, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(27, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "PO"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft JhengHei UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(45, 24)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(133, 24)
        Me.TextBox1.TabIndex = 2
        '
        'PO
        '
        Me.PO.HeaderText = "PO"
        Me.PO.Name = "PO"
        Me.PO.ReadOnly = True
        Me.PO.Width = 90
        '
        'OP
        '
        Me.OP.HeaderText = "OP"
        Me.OP.Name = "OP"
        Me.OP.ReadOnly = True
        Me.OP.Width = 90
        '
        'CODIGO
        '
        Me.CODIGO.HeaderText = "CODIGO"
        Me.CODIGO.Name = "CODIGO"
        Me.CODIGO.ReadOnly = True
        Me.CODIGO.Visible = False
        '
        'TELA
        '
        Me.TELA.HeaderText = "TELA"
        Me.TELA.Name = "TELA"
        Me.TELA.ReadOnly = True
        Me.TELA.Width = 350
        '
        'PROGRA
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.Azure
        Me.PROGRA.DefaultCellStyle = DataGridViewCellStyle5
        Me.PROGRA.HeaderText = "PROGRA"
        Me.PROGRA.Name = "PROGRA"
        Me.PROGRA.ReadOnly = True
        '
        'CANS_L
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        Me.CANS_L.DefaultCellStyle = DataGridViewCellStyle6
        Me.CANS_L.HeaderText = "CONS_LINEAL"
        Me.CANS_L.Name = "CANS_L"
        Me.CANS_L.ReadOnly = True
        Me.CANS_L.Width = 90
        '
        'CONS_KGS
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        Me.CONS_KGS.DefaultCellStyle = DataGridViewCellStyle7
        Me.CONS_KGS.HeaderText = "CONS_KGS"
        Me.CONS_KGS.Name = "CONS_KGS"
        Me.CONS_KGS.ReadOnly = True
        Me.CONS_KGS.Width = 90
        '
        'CANTID
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        DataGridViewCellStyle8.BackColor = System.Drawing.Color.Lavender
        Me.CANTID.DefaultCellStyle = DataGridViewCellStyle8
        Me.CANTID.HeaderText = "CANT_SOLICITADA"
        Me.CANTID.Name = "CANTID"
        Me.CANTID.ReadOnly = True
        Me.CANTID.Width = 120
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Honeydew
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft JhengHei UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(885, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(95, 42)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "GENERAR SOLICITUD"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Solicitud_telas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(988, 665)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "Solicitud_telas"
        Me.Text = "SOLICITUD DE TELA"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents PO As DataGridViewTextBoxColumn
    Friend WithEvents OP As DataGridViewTextBoxColumn
    Friend WithEvents CODIGO As DataGridViewTextBoxColumn
    Friend WithEvents TELA As DataGridViewTextBoxColumn
    Friend WithEvents PROGRA As DataGridViewTextBoxColumn
    Friend WithEvents CANS_L As DataGridViewTextBoxColumn
    Friend WithEvents CONS_KGS As DataGridViewTextBoxColumn
    Friend WithEvents CANTID As DataGridViewTextBoxColumn
    Friend WithEvents Button1 As Button
End Class
