Public Class Planilla_Movilidas
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add(1)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        '' Verificar si la celda actual es la que deseas controlar (por ejemplo, la columna 1)
        'If DataGridView1.CurrentCell.ColumnIndex = 0 Then
        '    Dim textBox As TextBox = DirectCast(e.Control, TextBox)
        '    AddHandler textBox.KeyPress, AddressOf TextBox_KeyPress
        'Else
        '    If DataGridView1.CurrentCell.ColumnIndex = 4 Then
        '        Dim textBox1 As TextBox = DirectCast(e.Control, TextBox)
        '        AddHandler textBox1.KeyPress, AddressOf TextBox1_KeyPress
        '    End If
        'End If


    End Sub

    Private Sub TextBox_KeyPress(sender As Object, e As KeyPressEventArgs)
        '' Verificar si la tecla presionada es un número o una tecla de control (por ejemplo, Backspace)
        'If Not Char.IsDigit(e.KeyChar) Then
        '    ' Cancelar la pulsación si no es un número ni una tecla de control

        '    e.Handled = True
        'End If


    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs)
        '' Obtener el carácter actual y el texto actual en el TextBox
        'Dim currentChar As Char = e.KeyChar
        'Dim currentText As String = DirectCast(sender, TextBox).Text

        '' Verificar si el carácter es un número, un punto decimal o una tecla de control
        'If Not Char.IsDigit(currentChar) AndAlso currentChar <> "." Then
        '    ' Cancelar la pulsación si no es un número, un punto decimal ni una tecla de control
        '    e.Handled = True
        'End If

        '' Verificar si ya hay un punto decimal en el texto
        'If currentChar = "." AndAlso currentText.Contains(".") Then
        '    ' Cancelar la pulsación si ya hay un punto decimal en el texto
        '    e.Handled = True
        'End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            TRABAJADORES.Label2.Text = 5
            TRABAJADORES.ShowDialog()
        End If
    End Sub

    Private Sub Planilla_Movilidas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Select Case Month(DateTimePicker1.Value)
            Case "01" : ComboBox2.SelectedIndex = 0
            Case "02" : ComboBox2.SelectedIndex = 1
            Case "03" : ComboBox2.SelectedIndex = 2
            Case "04" : ComboBox2.SelectedIndex = 3
            Case "05" : ComboBox2.SelectedIndex = 4
            Case "06" : ComboBox2.SelectedIndex = 5
            Case "07" : ComboBox2.SelectedIndex = 6
            Case "08" : ComboBox2.SelectedIndex = 7
            Case "09" : ComboBox2.SelectedIndex = 8
            Case "10" : ComboBox2.SelectedIndex = 9
        End Select

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.Text
            Case "LOGISTICA TEXTIL (COMPRAS Y ALMACEN)"
                TextBox3.Text = "9213"
                TextBox4.Text = "92599101"
            Case "SERVICIOS GENERALES"
                TextBox3.Text = "9216"
                TextBox4.Text = "92599101"
            Case "CALIDAD TEXTIL"
                TextBox3.Text = "9230"
                TextBox4.Text = "92599101"
            Case "LOGISTICA MANUFACTURA"
                TextBox3.Text = "9310"
                TextBox4.Text = "93599101"
            Case "LOGISTICA TEXTIL"
                TextBox3.Text = "9350"
                TextBox4.Text = "93599101"
            Case "LOGISTICA"
                TextBox3.Text = "9360"
                TextBox4.Text = "93599101"
            Case "LOGISTICA INSTITUCIONES PUBLICAS"
                TextBox3.Text = "9380"
                TextBox4.Text = "93599101"
            Case "PRODUCCION MANUFACTURA"
                TextBox3.Text = "9400"
                TextBox4.Text = "94599101"
            Case "SERVICIOS EXTERNOS - MANUFACTURA"
                TextBox3.Text = "9420"
                TextBox4.Text = "94599101"
            Case "PRODUCCION - UDP"
                TextBox3.Text = "9430"
                TextBox4.Text = "94599101"
            Case "SERVICIOS EXTERNOS - INSTITUCIONES PUBLICAS"
                TextBox3.Text = "9425"
                TextBox4.Text = "94599101"
            Case "ADMINISTRACION - CONTABILIDAD - RRHH - IT - GG - RRHH -SG"
                TextBox3.Text = "9610"
                TextBox4.Text = "96599101"
            Case "COMERCIAL MANUFACTURA"
                TextBox3.Text = "9400"
                TextBox4.Text = "97599101"
            Case "COMERCIAL TEXTIL"
                TextBox3.Text = "9712"
                TextBox4.Text = "97599101"
            Case "COMERCIAL INSTITUCIONES PUBLICAS"
                TextBox3.Text = "9718"
                TextBox4.Text = "97599101"
        End Select
    End Sub
End Class