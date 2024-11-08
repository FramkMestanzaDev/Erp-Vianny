
Public Class almacen_guia

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Guia_remision.Close()

        If Label2.Text = 1 Then

            If RadioButton1.Checked = True Then
                Guia_remision.Label23.Text = 37
                Guia_remision.Label25.Text = "  PRIMERA PLANTA 2"
                Guia_remision.TextBox19.Text = TextBox1.Text
                Guia_remision.Label29.Text = Label1.Text
                Guia_remision.Label30.Text = Label2.Text

                Guia_remision.Show()


            Else
                If RadioButton2.Checked = True Then
                    Guia_remision.Label23.Text = 38
                    Guia_remision.Label25.Text = "  SEGUNDA PLANTA 2"
                    Guia_remision.TextBox19.Text = TextBox1.Text
                    Guia_remision.Label29.Text = Label1.Text
                    Guia_remision.Label30.Text = Label2.Text

                    Guia_remision.Show()
                Else
                    If RadioButton3.Checked = True Then
                        Guia_remision.Label23.Text = 61
                        Guia_remision.Label25.Text = " MUESTRAS TEXTIL"
                        Guia_remision.TextBox19.Text = TextBox1.Text
                        Guia_remision.Label29.Text = Label1.Text
                        Guia_remision.Label30.Text = Label2.Text

                        Guia_remision.Show()
                    Else
                        If RadioButton4.Checked = True Then
                            Guia_remision.Label23.Text = 10
                            Guia_remision.Label25.Text = " TELA PUNTO PRIMERA CHILCA"
                            Guia_remision.TextBox19.Text = TextBox1.Text
                            Guia_remision.TextBox1.Text = "0014"
                            Guia_remision.TextBox3.Text = "MZ.E LOTE, PARCELACION RUSTICA HUERTOS DE ORO DE SAN HILARION Y PAMPA Y HOYADAS DE CALANGUILLO - CHILCA - CAÑETE -LIMA"
                            Guia_remision.Label29.Text = Label1.Text
                            Guia_remision.Label30.Text = Label2.Text

                            Guia_remision.Show()
                        Else
                            If RadioButton5.Checked = True Then
                                Guia_remision.Label23.Text = 51
                                Guia_remision.Label25.Text = " TELA PUNTO SEGUNDA CHILCA"
                                Guia_remision.TextBox19.Text = TextBox1.Text
                                Guia_remision.TextBox1.Text = "0014"
                                Guia_remision.TextBox3.Text = "MZ.E LOTE, PARCELACION RUSTICA HUERTOS DE ORO DE SAN HILARION Y PAMPA Y HOYADAS DE CALANGUILLO - CHILCA - CAÑETE -LIMA"
                                Guia_remision.Label29.Text = Label1.Text
                                Guia_remision.Label30.Text = Label2.Text

                                Guia_remision.Show()
                            Else
                                If RadioButton6.Checked = True Then
                                    Guia_remision.Label23.Text = 54
                                    Guia_remision.Label25.Text = " TELA PUNTO TERCERA CHILCA"
                                    Guia_remision.TextBox19.Text = TextBox1.Text
                                    Guia_remision.TextBox1.Text = "0014"
                                    Guia_remision.TextBox3.Text = "MZ.E LOTE, PARCELACION RUSTICA HUERTOS DE ORO DE SAN HILARION Y PAMPA Y HOYADAS DE CALANGUILLO - CHILCA - CAÑETE -LIMA"
                                    Guia_remision.Label29.Text = Label1.Text
                                    Guia_remision.Label30.Text = Label2.Text

                                    Guia_remision.Show()
                                Else
                                    If RadioButton7.Checked = True Then
                                        Guia_remision.Label23.Text = 68
                                        Guia_remision.Label25.Text = " ALMACEN HUACHIPA"
                                        Guia_remision.TextBox19.Text = TextBox1.Text
                                        Guia_remision.TextBox1.Text = "0018"
                                        Guia_remision.TextBox3.Text = "AV. LAS TORRES MZ F LOTE 08 HUACHIPA - CHOSICA"
                                        Guia_remision.Label29.Text = Label1.Text
                                        Guia_remision.Label30.Text = Label2.Text

                                        Guia_remision.Show()
                                    Else
                                        If RadioButton8.Checked = True Then
                                            Guia_remision.Label23.Text = "06"
                                            Guia_remision.Label25.Text = "  ALMACEN TELA CRUDA"
                                            Guia_remision.TextBox1.Text = "0014"
                                            Guia_remision.TextBox19.Text = TextBox1.Text
                                            Guia_remision.Label29.Text = Label1.Text
                                            Guia_remision.Label30.Text = Label2.Text

                                            Guia_remision.Show()
                                        Else
                                            If RadioButton9.Checked = True Then
                                                Guia_remision.Label23.Text = 66
                                                Guia_remision.Label25.Text = " MERMA"
                                                Guia_remision.TextBox19.Text = TextBox1.Text
                                                Guia_remision.TextBox1.Text = "0014"
                                                Guia_remision.TextBox3.Text = "MZ.E LOTE, PARCELACION RUSTICA HUERTOS DE ORO DE SAN HILARION Y PAMPA Y HOYADAS DE CALANGUILLO - CHILCA - CAÑETE -LIMA"
                                                Guia_remision.Label29.Text = Label1.Text
                                                Guia_remision.Label30.Text = Label2.Text

                                                Guia_remision.Show()
                                            Else
                                                If RadioButton10.Checked = True Then
                                                    Guia_remision.Label23.Text = 39
                                                    Guia_remision.Label25.Text = "  TERCERA PLANTA 2"
                                                    Guia_remision.TextBox19.Text = TextBox1.Text
                                                    Guia_remision.Label29.Text = Label1.Text
                                                    Guia_remision.Label30.Text = Label2.Text

                                                    Guia_remision.Show()
                                                Else
                                                    If RadioButton11.Checked = True Then
                                                        Guia_remision.Label23.Text = "08"
                                                        Guia_remision.Label25.Text = "  TELA PLANA"
                                                        Guia_remision.TextBox19.Text = TextBox1.Text
                                                        Guia_remision.Label29.Text = Label1.Text
                                                        Guia_remision.Label30.Text = Label2.Text

                                                        Guia_remision.Show()
                                                    Else
                                                        If RadioButton12.Checked = True Then
                                                            Guia_remision.Label23.Text = 24
                                                            Guia_remision.Label25.Text = "  TELA PLANA MANUFACTURA"
                                                            Guia_remision.TextBox19.Text = TextBox1.Text
                                                            Guia_remision.Label29.Text = Label1.Text
                                                            Guia_remision.Label30.Text = Label2.Text

                                                            Guia_remision.Show()
                                                        End If
                                                    End If
                                                End If
                                                End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If RadioButton1.Checked = True Then
                Guia_remision.Label23.Text = "08"
                Guia_remision.Label25.Text = " TELA PLANA GRAUS"
                Guia_remision.TextBox19.Text = TextBox1.Text
                Guia_remision.TextBox1.Text = "0001"

                Guia_remision.Label29.Text = Label1.Text
                Guia_remision.Label30.Text = Label2.Text
                Guia_remision.Show()
            Else
                If RadioButton2.Checked = True Then
                    Guia_remision.Label23.Text = 10
                    Guia_remision.Label25.Text = " TELA PUNTO GRAUS"
                    Guia_remision.TextBox19.Text = TextBox1.Text
                    Guia_remision.TextBox1.Text = "0001"

                    Guia_remision.Label29.Text = Label1.Text
                    Guia_remision.Label30.Text = Label2.Text
                    Guia_remision.Show()
                Else
                    If RadioButton3.Checked = True Then
                        Guia_remision.Label23.Text = 67
                        Guia_remision.Label25.Text = " HILO - TELA GRAUS"
                        Guia_remision.TextBox19.Text = TextBox1.Text
                        Guia_remision.TextBox1.Text = "0013"

                        Guia_remision.Label29.Text = Label1.Text
                        Guia_remision.Label30.Text = Label2.Text
                        Guia_remision.Show()
                    Else
                        If RadioButton7.Checked = True Then
                            Guia_remision.Label23.Text = 68
                            Guia_remision.Label25.Text = " ALMACEN HUACHIPA"
                            Guia_remision.TextBox19.Text = TextBox1.Text
                            Guia_remision.TextBox1.Text = "0001"

                            Guia_remision.Label29.Text = Label1.Text
                            Guia_remision.Label30.Text = Label2.Text
                            Guia_remision.Show()
                        Else
                            If RadioButton11.Checked = True Then
                                Guia_remision.Label23.Text = "08"
                                Guia_remision.Label25.Text = " ALMACEN TELA PLANA"
                                Guia_remision.TextBox19.Text = TextBox1.Text
                                Guia_remision.TextBox1.Text = "0001"

                                Guia_remision.Label29.Text = Label1.Text
                                Guia_remision.Label30.Text = Label2.Text
                                Guia_remision.Show()
                            Else
                                If RadioButton12.Checked = True Then
                                    Guia_remision.Label23.Text = "24"
                                    Guia_remision.Label25.Text = " ALMACEN TELA MANUFACTURA"
                                    Guia_remision.TextBox19.Text = TextBox1.Text
                                    Guia_remision.TextBox1.Text = "0001"

                                    Guia_remision.Label29.Text = Label1.Text
                                    Guia_remision.Label30.Text = Label2.Text
                                    Guia_remision.Show()
                                End If
                            End If
                        End If
                    End If

                End If
            End If
        End If

    End Sub

    Private Sub almacen_guia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Label2.Text = "02" Then
            RadioButton1.Text = "08 - TELA PLANA GRAUS"
            RadioButton2.Text = "10 - TELA PUNTO GRAUS"
            RadioButton3.Text = "67 - HILO - TELA GRAUS"
            RadioButton7.Text = "68 -ALMACEN HUACHIPA GRAUS"
            RadioButton4.Visible = False
            RadioButton5.Visible = False
            RadioButton6.Visible = False
        End If
    End Sub
End Class