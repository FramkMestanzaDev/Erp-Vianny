Imports System.Data.SqlClient
Public Class Form_Cotizacion
    Public Property Padre As CotizacionUdp
    Public conx As SqlConnection
    Dim Rsr3, Rsr4 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim P As New Producto_Cotizacion

        'P.ShowDialog()
        Dim p As New pproductos
        p.Padre = Me
        p.CCIA.Text = "01"
        If TextBox3.Text = "1" Then
            p.ALMACEN.Text = "08"
        Else
            p.ALMACEN.Text = "13"
        End If
        p.ANO.Text = Year(DateTimePicker1.Value)
        p.FECHA.Text = Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "")
        p.TextBox1.Text = Trim(TextBox1.Text)
        p.Label3.Text = "16"
        p.Label5.Text = "0"
        p.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim NU, i17, num2, nu2 As Integer


        If TextBox2.Text = "" Then
            MsgBox("FALTA INGRESAR EL ARTICULO")
        Else
            DataGridView1.Rows.Add(1)
            NU = DataGridView1.Rows.Count

            If NU = 1 Then
                DataGridView1.Rows(0).Cells(0).Value = TextBox1.Text
                DataGridView1.Rows(0).Cells(1).Value = TextBox2.Text
                DataGridView1.Rows(0).Cells(8).Value = NU + 1
                DataGridView1.Rows(0).Cells(2).Value = 1
                'DataGridView1.Rows(0).Cells(6).Value = 0
                DataGridView1.Rows(0).Cells(4).Value = TextBox4.Text & "/PRD"
                DataGridView1.Rows(0).Cells(3).Value = 0
                DataGridView1.Rows(0).Cells(5).Value = 0
                DataGridView1.Rows(0).Cells(6).Value = 0
                DataGridView1.Rows(0).Cells(7).Value = 0
                abrir()
                Dim sql102213 As String = "EXEC Sp_ListadoUltimoPrecios '01',NULL,'" + TextBox1.Text.ToString().Trim() + "','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
                Dim cmd102213 As New SqlCommand(sql102213, conx)
                Rsr3 = cmd102213.ExecuteReader()
                If Rsr3.Read() = True Then
                    DataGridView1.Rows(0).Cells(6).Value = Trim(Rsr3(0).ToString())
                End If
                Rsr3.Close()

            Else
                nu2 = DataGridView1.Rows.Count
                For i = 1 To nu2 - 1


                    If DataGridView1.Rows(i).Cells(8).Value < DataGridView1.Rows(i - 1).Cells(8).Value Then
                        num2 = DataGridView1.Rows(i - 1).Cells(8).Value + 1
                        i17 = i

                    End If


                Next


                If i17 = "0" Then
                    MsgBox("SOLO SE PERMITE 5 REGISTROS")
                Else
                    DataGridView1.Rows(i17).Cells(0).Value = TextBox1.Text
                    DataGridView1.Rows(i17).Cells(1).Value = TextBox2.Text
                    DataGridView1.Rows(i17).Cells(8).Value = num2
                    DataGridView1.Rows(i17).Cells(2).Value = 1
                    'DataGridView1.Rows(0).Cells(6).Value = 0
                    DataGridView1.Rows(i17).Cells(4).Value = TextBox4.Text & "/PRD"
                    DataGridView1.Rows(i17).Cells(3).Value = 0
                    DataGridView1.Rows(i17).Cells(5).Value = 0
                    DataGridView1.Rows(i17).Cells(6).Value = 0
                    DataGridView1.Rows(i17).Cells(7).Value = 0
                    abrir()

                    Dim sql102213 As String = "EXEC Sp_ListadoUltimoPrecios '01',NULL,'" + TextBox1.Text.ToString().Trim() + "','" + Replace(DateTimePicker1.Value.ToString("yyyy-MM-dd"), "-", "") + "'"
                    Dim cmd102213 As New SqlCommand(sql102213, conx)
                    Rsr4 = cmd102213.ExecuteReader()
                    If Rsr4.Read() = True Then
                        DataGridView1.Rows(i17).Cells(6).Value = Trim(Rsr4(0).ToString())
                    End If

                    Rsr4.Close()


                End If


            End If
        End If
        REGISTRAR_INFORMACION()
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
    Sub REGISTRAR_INFORMACION()
        If TextBox3.Text.ToString().Trim() = "1" Then
            Me.Text = "TELA"
            Dim va, la As Integer
            Dim sum As Double
            va = DataGridView1.Rows.Count
            Padre.DataGridView1.Rows.Clear()
            If va > 0 Then
                Padre.DataGridView1.Rows.Add(va)
                For i = 0 To va - 1
                    Padre.DataGridView1.Rows(i).Cells(0).Value = DataGridView1.Rows(i).Cells(0).Value
                    Padre.DataGridView1.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(1).Value
                    Padre.DataGridView1.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(2).Value
                    Padre.DataGridView1.Rows(i).Cells(3).Value = DataGridView1.Rows(i).Cells(3).Value
                    Padre.DataGridView1.Rows(i).Cells(4).Value = DataGridView1.Rows(i).Cells(4).Value
                    Padre.DataGridView1.Rows(i).Cells(5).Value = DataGridView1.Rows(i).Cells(5).Value
                    Padre.DataGridView1.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(6).Value
                    Padre.DataGridView1.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(7).Value
                    Padre.DataGridView1.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(8).Value
                Next
            End If
            la = DataGridView1.Rows.Count
            For p = 0 To la - 1
                sum = sum + CDbl(DataGridView1.Rows(p).Cells(7).Value)
            Next
            Padre.TextBox7.Text = Math.Round(sum, 3)

        Else
            If TextBox3.Text.ToString().Trim() = "2" Then
                Me.Text = "AVIOS COSTURA"
                Dim va, va2, la As Integer
                Dim sum As Double
                va = DataGridView1.Rows.Count
                va2 = Padre.DataGridView2.Rows.Count
                Padre.DataGridView2.Rows.Clear()

                If va > 0 Then

                    Padre.DataGridView2.Rows.Add(va)
                    For i = 0 To va - 1
                        Padre.DataGridView2.Rows(i).Cells(0).Value = DataGridView1.Rows(i).Cells(0).Value
                        Padre.DataGridView2.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(1).Value
                        Padre.DataGridView2.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(2).Value
                        Padre.DataGridView2.Rows(i).Cells(3).Value = DataGridView1.Rows(i).Cells(3).Value
                        Padre.DataGridView2.Rows(i).Cells(4).Value = DataGridView1.Rows(i).Cells(4).Value
                        Padre.DataGridView2.Rows(i).Cells(5).Value = DataGridView1.Rows(i).Cells(5).Value
                        Padre.DataGridView2.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(6).Value
                        Padre.DataGridView2.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(7).Value
                        Padre.DataGridView2.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(8).Value
                    Next
                End If
                la = DataGridView1.Rows.Count
                For p = 0 To la - 1
                    sum = sum + CDbl(DataGridView1.Rows(p).Cells(7).Value)
                Next
                Padre.TextBox19.Text = Math.Round(sum, 3)

            Else
                Me.Text = "AVIOS ACABADOS EMBALAJE"
                Dim va, la As Integer
                Dim sum As Double
                va = DataGridView1.Rows.Count
                Padre.DataGridView3.Rows.Clear()
                If va > 0 Then
                    Padre.DataGridView3.Rows.Add(va)
                    For i = 0 To va - 1
                        Padre.DataGridView3.Rows(i).Cells(0).Value = DataGridView1.Rows(i).Cells(0).Value
                        Padre.DataGridView3.Rows(i).Cells(1).Value = DataGridView1.Rows(i).Cells(1).Value
                        Padre.DataGridView3.Rows(i).Cells(2).Value = DataGridView1.Rows(i).Cells(2).Value
                        Padre.DataGridView3.Rows(i).Cells(3).Value = DataGridView1.Rows(i).Cells(3).Value
                        Padre.DataGridView3.Rows(i).Cells(4).Value = DataGridView1.Rows(i).Cells(4).Value
                        Padre.DataGridView3.Rows(i).Cells(5).Value = DataGridView1.Rows(i).Cells(5).Value
                        Padre.DataGridView3.Rows(i).Cells(6).Value = DataGridView1.Rows(i).Cells(6).Value
                        Padre.DataGridView3.Rows(i).Cells(7).Value = DataGridView1.Rows(i).Cells(7).Value
                        Padre.DataGridView3.Rows(i).Cells(8).Value = DataGridView1.Rows(i).Cells(8).Value
                    Next
                End If
                la = DataGridView1.Rows.Count
                For p = 0 To la - 1
                    sum = sum + CDbl(DataGridView1.Rows(p).Cells(7).Value)
                Next
                Padre.TextBox17.Text = Math.Round(sum, 3)

            End If
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, fila As Integer
        'Dim cant9 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "CONSUMO" Then
            DataGridView1.Rows(fila).Cells(5).Value = (Val(DataGridView1.Rows(fila).Cells(3).Value) * (1 + Val(DataGridView1.Rows(fila).Cells(2).Value) / 100))
            Me.DataGridView1.Columns(5).DefaultCellStyle.Format = "n3"
            DataGridView1.Rows(fila).Cells(7).Value = Val(DataGridView1.Rows(fila).Cells(5).Value) * (Val(DataGridView1.Rows(fila).Cells(6).Value))
            Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n3"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "% MERMA" Then
            DataGridView1.Rows(fila).Cells(5).Value = (Val(DataGridView1.Rows(fila).Cells(3).Value) * (1 + Val(DataGridView1.Rows(fila).Cells(2).Value) / 100))
            Me.DataGridView1.Columns(5).DefaultCellStyle.Format = "n3"
            DataGridView1.Rows(fila).Cells(7).Value = Val(DataGridView1.Rows(fila).Cells(5).Value) * (Val(DataGridView1.Rows(fila).Cells(6).Value))
            Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n3"
        End If
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "C.UNITARIO" Then
            DataGridView1.Rows(fila).Cells(7).Value = Val(DataGridView1.Rows(fila).Cells(5).Value) * (Val(DataGridView1.Rows(fila).Cells(6).Value))
            Me.DataGridView1.Columns(7).DefaultCellStyle.Format = "n3"

        End If

    End Sub

    Private Sub Form_Cotizacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Select()
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click

        Dim respuesta As DialogResult
        If DataGridView1.Rows.Count > 0 Then
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then
                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
                REGISTRAR_INFORMACION()
            End If
        Else
            MsgBox("El Datagrid esta vacio, tiene que tener informacion para eliminar un Items")
        End If
    End Sub

    Private Sub Form_Cotizacion_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        REGISTRAR_INFORMACION()
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.C Then
            ' Verifica si hay una celda seleccionada
            If DataGridView1.CurrentCell IsNot Nothing Then
                ' Obtiene el valor de la celda seleccionada
                Dim cellValue As String = DataGridView1.CurrentCell.Value.ToString()
                ' Copia el valor de la celda al portapapeles
                MsgBox(cellValue.ToString())
                Clipboard.SetText(cellValue)
                ' Previene la copia predeterminada de toda la fila
                e.Handled = True
            End If
        End If
    End Sub
    Dim rs77 As SqlDataReader
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            abrir()
            Dim sql10112 As String = "Select nomb_17 from custom_vianny.dbo.cag1700 where ccia ='01' and linea_17 ='" + TextBox1.Text.ToString().Trim() + "'"
            Dim cmd10112 As New SqlCommand(sql10112, conx)
            rs77 = cmd10112.ExecuteReader()
            If rs77.Read() = True Then
                If rs77(0).ToString().Trim().Length() > 0 Then
                    TextBox2.Text = rs77(0).ToString().Trim()
                End If


            Else
                    MsgBox("El codigo Ingresado no Existe")
            End If
            rs77.Close()
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            ' Obtener la celda que se hizo clic
            Dim cell As DataGridViewCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)

            ' Activar el modo de edición de la celda
            DataGridView1.BeginEdit(True)
        End If
        If e.ColumnIndex = 0 Then
            Dim cellValue As String = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString()
            Clipboard.SetText(cellValue)
            ' Opcional: Mostrar un mensaje de confirmación
            'MessageBox.Show("El valor de la celda " & DataGridView1.Columns(e.ColumnIndex).HeaderText.ToString().Trim() & " ha sido copiado al portapapeles: " & cellValue, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class