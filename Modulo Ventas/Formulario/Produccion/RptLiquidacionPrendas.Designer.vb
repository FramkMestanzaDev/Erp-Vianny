<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RptLiquidacionPrendas
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
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.comboBox1 = New System.Windows.Forms.ComboBox()
        Me.textBox3 = New System.Windows.Forms.TextBox()
        Me.textBox2 = New System.Windows.Forms.TextBox()
        Me.textBox1 = New System.Windows.Forms.TextBox()
        Me.label3 = New System.Windows.Forms.Label()
        Me.label2 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.button1 = New System.Windows.Forms.Button()
        Me.groupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.button1)
        Me.groupBox1.Controls.Add(Me.comboBox1)
        Me.groupBox1.Controls.Add(Me.textBox3)
        Me.groupBox1.Controls.Add(Me.textBox2)
        Me.groupBox1.Controls.Add(Me.textBox1)
        Me.groupBox1.Controls.Add(Me.label3)
        Me.groupBox1.Controls.Add(Me.label2)
        Me.groupBox1.Controls.Add(Me.label1)
        Me.groupBox1.Location = New System.Drawing.Point(12, 12)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(376, 114)
        Me.groupBox1.TabIndex = 1
        Me.groupBox1.TabStop = False
        '
        'comboBox1
        '
        Me.comboBox1.BackColor = System.Drawing.Color.MintCream
        Me.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.comboBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.comboBox1.Font = New System.Drawing.Font("Rockwell", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comboBox1.FormattingEnabled = True
        Me.comboBox1.Items.AddRange(New Object() {"ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"})
        Me.comboBox1.Location = New System.Drawing.Point(102, 76)
        Me.comboBox1.Name = "comboBox1"
        Me.comboBox1.Size = New System.Drawing.Size(134, 25)
        Me.comboBox1.TabIndex = 6
        '
        'textBox3
        '
        Me.textBox3.BackColor = System.Drawing.SystemColors.Info
        Me.textBox3.Font = New System.Drawing.Font("Rockwell", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox3.Location = New System.Drawing.Point(242, 13)
        Me.textBox3.Name = "textBox3"
        Me.textBox3.ReadOnly = True
        Me.textBox3.Size = New System.Drawing.Size(37, 25)
        Me.textBox3.TabIndex = 5
        '
        'textBox2
        '
        Me.textBox2.BackColor = System.Drawing.SystemColors.Info
        Me.textBox2.Font = New System.Drawing.Font("Rockwell", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox2.Location = New System.Drawing.Point(102, 48)
        Me.textBox2.Name = "textBox2"
        Me.textBox2.ReadOnly = True
        Me.textBox2.Size = New System.Drawing.Size(134, 25)
        Me.textBox2.TabIndex = 4
        '
        'textBox1
        '
        Me.textBox1.BackColor = System.Drawing.SystemColors.Info
        Me.textBox1.Font = New System.Drawing.Font("Rockwell", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox1.Location = New System.Drawing.Point(102, 13)
        Me.textBox1.Name = "textBox1"
        Me.textBox1.ReadOnly = True
        Me.textBox1.Size = New System.Drawing.Size(134, 25)
        Me.textBox1.TabIndex = 3
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Font = New System.Drawing.Font("Rockwell", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label3.Location = New System.Drawing.Point(21, 76)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(38, 17)
        Me.label3.TabIndex = 2
        Me.label3.Text = "MES"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Font = New System.Drawing.Font("Rockwell", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label2.Location = New System.Drawing.Point(21, 48)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(75, 17)
        Me.label2.TabIndex = 1
        Me.label2.Text = "PERIODO"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Font = New System.Drawing.Font("Rockwell", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label1.Location = New System.Drawing.Point(21, 16)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(75, 17)
        Me.label1.TabIndex = 0
        Me.label1.Text = "EMPRESA"
        '
        'button1
        '
        Me.button1.Image = Global.Modulo_Ventas.Resources.buscar_1_
        Me.button1.Location = New System.Drawing.Point(268, 48)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(102, 54)
        Me.button1.TabIndex = 8
        Me.button1.UseVisualStyleBackColor = True
        '
        'RptLiquidacionPrendas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(397, 133)
        Me.Controls.Add(Me.groupBox1)
        Me.Name = "RptLiquidacionPrendas"
        Me.Text = "RptLiquidacionPrendas"
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents groupBox1 As GroupBox
    Private WithEvents button1 As Button
    Private WithEvents comboBox1 As ComboBox
    Friend WithEvents textBox1 As TextBox
    Private WithEvents label3 As Label
    Private WithEvents label2 As Label
    Private WithEvents label1 As Label
    Protected Friend WithEvents textBox3 As TextBox
    Protected Friend WithEvents textBox2 As TextBox
End Class
