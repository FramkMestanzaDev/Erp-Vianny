Imports System.Data.SqlClient
Public Class Packing_huabipa
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public enunciado1 As SqlCommand
    Public respuesta1 As SqlDataReader
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataReader
    Public enunciado3 As SqlCommand
    Public respuesta3 As SqlDataReader
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter

    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim jh As New DataTable
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.Text
            Case "PRIMERA" : TextBox2.Text = "10"
            Case "SEGUNDA" : TextBox2.Text = "51"
            Case "TERCERA" : TextBox2.Text = "54"

        End Select
        Select Case ComboBox1.Text

            Case "PRIMERA" : TextBox5.Text = "68"
            Case "SEGUNDA" : TextBox5.Text = "38"
            Case "TERCERA" : TextBox5.Text = "39"
        End Select

        DataGridView1.Rows.Clear()
        DataGridView2.DataSource = ""
        Dim vg2 As New vpackingtela
        Dim vg3 As New fingresopac
        Dim i As Integer

        vg2.gpartida = TextBox1.Text
        vg2.gnumero_rollos = ComboBox1.Text

        vg2.gAL = Trim(TextBox5.Text)
        jh = vg3.imprimir_packing2(vg2)

        If jh.Rows.Count > 0 Then
            Dim suma10 As Double
            Dim suma2 As Integer

            DataGridView2.DataSource = jh
            Dim UI As Integer
            UI = DataGridView2.Rows.Count
            DataGridView1.Rows.Add(UI)
            For i = 0 To UI - 1
                DataGridView1.Rows(i).Cells(0).Value = DataGridView2.Rows(i).Cells(16).Value
                DataGridView1.Rows(i).Cells(1).Value = DataGridView2.Rows(i).Cells(17).Value
                DataGridView1.Rows(i).Cells(5).Value = DataGridView2.Rows(i).Cells(18).Value
                DataGridView1.Rows(i).Cells(3).Value = DataGridView2.Rows(i).Cells(19).Value
                DataGridView1.Rows(i).Cells(4).Value = DataGridView2.Rows(i).Cells(20).Value
                DataGridView1.Rows(i).Cells(7).Value = DataGridView2.Rows(i).Cells(27).Value
                DataGridView1.Rows(i).Cells(8).Value = i + 1
                suma10 = suma10 + DataGridView1.Rows(i).Cells(7).Value
                suma2 = suma2 + 1

            Next

            TextBox8.Text = suma10
            TextBox11.Text = suma2



            ComboBox1.Text = DataGridView2.Rows(0).Cells(5).Value

            PictureBox1.Enabled = True

        Else
            TextBox5.Enabled = True

        End If
        'i = DataGridView1.Rows.Count
        'For a9 = 0 To i - 1
        '    cant9 = Val(DataGridView1.Rows(a9).Cells(1).Value)
        '    sum9 = cant9 + Val(sum9)
        'Next


        'TextBox8.Text = sum9.ToString("N2")
        TextBox5.Select()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim Rsr23 As SqlDataReader
            abrir()
            Dim sql1023 As String = "SELECT COUNT(des_articulo) FROM cab_ingresop WHERE partida ='" + TextBox1.Text + "' "
            Dim cmd1023 As New SqlCommand(sql1023, conx)
            Rsr23 = cmd1023.ExecuteReader()
            Rsr23.Read()

            If Rsr23(0) = 0 Then
                MsgBox("LA PARTIDA SOLICITADA NO TIENE PACKING GENERADO EN CHILCA")
                Rsr23.Close()
            Else
                Rsr23.Close()
                ComboBox1.Enabled = True
                Button3.Enabled = True
                ComboBox1.Select()
                Dim Rsr234 As SqlDataReader
                abrir()

                Dim sql10234 As String = "SELECT DISTINCT des_articulo FROM cab_ingresop WHERE partida ='" + TextBox1.Text + "'"
                Dim cmd10234 As New SqlCommand(sql10234, conx)
                Rsr234 = cmd10234.ExecuteReader()
                Rsr234.Read()
                TextBox3.Text = Rsr234(0)
                Rsr234.Close()
            End If


        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DataGridView1.Rows.Clear()
        Dim Rsr2 As SqlDataReader

        abrir()
        Dim sql102 As String = "SELECT nrollo,peso,posicion,almacen,articulo,partida,id_cabe,GUIA,peso_neto FROM det_ingresop WHERE partida ='" + TextBox1.Text + "' AND almac ='" + TextBox2.Text + "'"
        Dim cmd102 As New SqlCommand(sql102, conx)
        Rsr2 = cmd102.ExecuteReader()
        Dim i5 As Integer
        i5 = 0

        While Rsr2.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i5).Cells(0).Value = Rsr2(0)
            DataGridView1.Rows(i5).Cells(1).Value = Rsr2(1)

            DataGridView1.Rows(i5).Cells(3).Value = Rsr2(3)
            DataGridView1.Rows(i5).Cells(4).Value = Rsr2(4)
            DataGridView1.Rows(i5).Cells(5).Value = Rsr2(2)

            DataGridView1.Rows(i5).Cells(7).Value = Rsr2(8)
            DataGridView1.Rows(i5).Cells(8).Value = i5 + 1
            i5 = i5 + 1
            'DataGridView2.Rows(i5).Cells(6).Value = (Rsr2(12) * DataGridView1.Rows(0).Cells(10).Value) / 100
        End While
        Dim I, sum8 As Integer
        Dim cant9, sum9 As Double
        I = DataGridView1.Rows.Count

        For I1 = 0 To I - 1

            cant9 = Val(DataGridView1.Rows(I1).Cells(7).Value)
            sum9 = cant9 + Val(sum9)
            sum8 = sum8 + 1
        Next


        TextBox11.Text = sum8
        TextBox8.Text = sum9
        TextBox1.Enabled = False
        Rsr2.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim bh As New vpackingtela
        Dim bh5 As New vpackingtela
        Dim ch As New vpackingtela
        Dim bh1 As New fingresopac
        Dim as1, k As Integer
        Dim jh As New Double
        bh.gpartida = TextBox1.Text
        bh.garticulo = TextBox3.Text
        bh.gcodigo_trab = ""
        bh.gnombre_trab = TextBox4.Text
        bh.gnumero_rollos = ComboBox1.Text
        bh.gunidad = "KGS"
        bh.gorden = ""
        bh.gdesidad = 0
        bh.gestado = 1
        k = DataGridView1.Rows.Count
        For o = 0 To k - 1
            jh = jh + DataGridView1.Rows(o).Cells(1).Value
        Next
        bh.garticulo3 = DataGridView1.Rows(0).Cells(4).Value
        bh.gseleccionado = 0
        bh.gtotal = jh
        bh.gROOLO = k
        bh.galmac = TextBox5.Text

        bh5.gpartida = TextBox1.Text
        bh5.gnumero_rollos = ComboBox1.Text

        bh1.eliminar_packing(bh5)
        If bh1.insertar_packing_tela(bh) = True Then
            as1 = DataGridView1.Rows.Count
            For i = 0 To as1 - 1
                ch.gnrollo = DataGridView1.Rows(i).Cells(0).Value
                ch.gpeso = DataGridView1.Rows(i).Cells(1).Value
                If Convert.ToString(DataGridView1.Rows(i).Cells(5).Value) = "" Then
                    ch.gposicion = ""

                Else
                    ch.gposicion = DataGridView1.Rows(i).Cells(5).Value
                End If
                If Convert.ToString(DataGridView1.Rows(i).Cells(3).Value) = "" Then
                    ch.galmacen = ""
                Else
                    ch.galmacen = DataGridView1.Rows(i).Cells(3).Value
                End If



                ch.garticulo2 = DataGridView1.Rows(i).Cells(4).Value
                ch.gpartida = TextBox1.Text
                ch.gestado1 = 1
                ch.gseleccionado = 0
                ch.gidcabe = ComboBox1.Text
                ch.galmac = TextBox5.Text
                ch.gpeso_neto = DataGridView1.Rows(i).Cells(7).Value
                If Trim(Convert.ToString(DataGridView1.Rows(i).Cells(2).Value)) = "" Then
                    ch.gubic_art = ""
                Else
                    ch.gubic_art = DataGridView1.Rows(i).Cells(2).Value
                End If

                bh1.insertar_packing_tela_detalle(ch)

            Next
            MsgBox("SE REGISTRO CON EXITO")
            TextBox1.Text = ""
            TextBox2.Text = ""

            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox8.Text = ""
            TextBox11.Text = ""

            DataGridView1.Rows.Clear()

            ComboBox1.SelectedIndex = -1
            ComboBox1.Enabled = False
            TextBox1.Enabled = True
        Else
            MsgBox("LA PARTIDA YA ESTA REGSTRADA")
        End If

    End Sub

    Private Sub Packing_huabipa_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox4.Text = MDIParent1.Label3.Text
        TextBox1.Select()
    End Sub
    Function ELIMINAR(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        abrir()
        Dim respuesta As DialogResult
        respuesta = MessageBox.Show("SE ELIMINAR TODA LA INFORMACIO DEL PACKING?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            Dim agregar As String = " DELETE FROM cab_ingresop WHERE partida ='" + Trim(TextBox1.Text) + "' and numero_rollos ='" + Trim(ComboBox1.Text) + "' AND almac = '" + TextBox5.Text + "'"
            Dim agregar1 As String = " DELETE FROM det_ingresop where partida='" + Trim(TextBox1.Text) + "' and id_cabe ='" + Trim(ComboBox1.Text) + "' AND almac = '" + TextBox5.Text + "'"
            ELIMINAR(agregar)
            Eliminar(agregar1)
            MsgBox("SE ELIMINO LA INFORMACION CORRECTAMENTE")
            Me.Close()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult

        respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (respuesta = Windows.Forms.DialogResult.Yes) Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
            Dim i, a9 As Integer
            Dim cant9, sum9 As Double
            i = DataGridView1.Rows.Count
            For a9 = 0 To i - 1
                cant9 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                sum9 = cant9 + Val(sum9)
                DataGridView1.Rows(a9).Cells(8).Value = a9 + 1
            Next
            TextBox11.Text = i
            TextBox8.Text = sum9.ToString("N2")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DataGridView1.Rows.Clear()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox11.Text = ""
        TextBox8.Text = ""
        TextBox1.Enabled = True
        Button3.Enabled = False
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        rPT_PACKIN_HIANCHIPA.TextBox1.Text = TextBox1.Text
        rPT_PACKIN_HIANCHIPA.TextBox2.Text = ComboBox1.Text
        rPT_PACKIN_HIANCHIPA.TextBox3.Text = TextBox5.Text
        rPT_PACKIN_HIANCHIPA.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim i, a9, fila, i2 As Integer
        Dim cant9, sum9 As Double
        i = DataGridView1.Rows.Count
        fila = e.RowIndex
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "KGS  NETO" Then
            i2 = DataGridView1.Rows.Count
            For a9 = 0 To i - 1
                cant9 = Val(DataGridView1.Rows(a9).Cells(7).Value)
                sum9 = cant9 + Val(sum9)

            Next
            DataGridView1.Rows(e.RowIndex).Cells(1).Value = Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(7).Value) + 0.3
            Dim u As Integer
            u = e.RowIndex
            TextBox8.Text = sum9.ToString("N2")
            TextBox11.Text = i2
        End If
    End Sub
End Class