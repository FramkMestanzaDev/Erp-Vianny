Public Class Almacen_guia_chil
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Guia_hilo.Close()

        If Label2.Text = "01" Then
            If RadioButton1.Checked = True Then
                Guia_hilo.Label23.Text = "03"
                Guia_hilo.Label25.Text = "  HILADO OPERATIVO"
                Guia_hilo.TextBox19.Text = TextBox1.Text
                Guia_hilo.Label29.Text = Label1.Text
                Guia_hilo.Label30.Text = Label2.Text
                Guia_hilo.TextBox1.Text = "0015"
                Guia_hilo.Show()

            Else
                If RadioButton2.Checked = True Then
                    Guia_hilo.Label23.Text = 42
                    Guia_hilo.Label25.Text = "  HILADO TEÑIDO"
                    Guia_hilo.TextBox19.Text = TextBox1.Text
                    Guia_hilo.Label29.Text = Label1.Text
                    Guia_hilo.Label30.Text = Label2.Text
                    Guia_hilo.TextBox1.Text = "0015"
                    Guia_hilo.Show()
                Else
                    If RadioButton3.Checked = True Then
                        Guia_hilo.Label23.Text = 59
                        Guia_hilo.Label25.Text = " SERVICIO TERCERO"
                        Guia_hilo.TextBox19.Text = TextBox1.Text
                        Guia_hilo.Label29.Text = Label1.Text
                        Guia_hilo.TextBox1.Text = "0013"
                        Guia_hilo.Label30.Text = Label2.Text
                        Guia_hilo.Show()
                    Else
                        If RadioButton4.Checked = True Then
                            Guia_hilo.Label23.Text = 43
                            Guia_hilo.Label25.Text = "  HILADO OPERATIVO"
                            Guia_hilo.TextBox19.Text = TextBox1.Text
                            Guia_hilo.Label29.Text = Label1.Text
                            Guia_hilo.TextBox1.Text = "0013"
                            Guia_hilo.Label30.Text = Label2.Text
                            Guia_hilo.Show()
                        Else
                            If RadioButton5.Checked = True Then
                                Guia_hilo.Label23.Text = 44
                                Guia_hilo.Label25.Text = " ECONOMATOS Y RESPUESTOS CHILCA "
                                Guia_hilo.TextBox19.Text = TextBox1.Text
                                Guia_hilo.Label29.Text = Label1.Text
                                Guia_hilo.TextBox1.Text = "0013"
                                Guia_hilo.Label30.Text = Label2.Text
                                Guia_hilo.Show()
                            Else
                                If RadioButton6.Checked = True Then
                                    Guia_hilo.Label23.Text = 41
                                    Guia_hilo.Label25.Text = " INSUMOS QUIMICOS CHILCA"
                                    Guia_hilo.TextBox19.Text = TextBox1.Text
                                    Guia_hilo.Label29.Text = Label1.Text
                                    Guia_hilo.Label30.Text = Label2.Text
                                    Guia_hilo.TextBox1.Text = "0013"
                                    Guia_hilo.Show()
                                End If
                            End If
                        End If

                    End If
                End If
            End If
        Else
            If RadioButton1.Checked = True Then
                Guia_hilo.Label23.Text = "05"
                Guia_hilo.Label25.Text = " RESPUESTOS Y SUMINISTROS"
                Guia_hilo.TextBox19.Text = TextBox1.Text
                Guia_hilo.Label29.Text = Label1.Text
                Guia_hilo.Label30.Text = Label2.Text
                Guia_hilo.TextBox1.Text = "0001"
                Guia_hilo.Show()
            Else
                If RadioButton2.Checked = True Then
                    Guia_hilo.Label23.Text = "14"
                    Guia_hilo.Label25.Text = " ECONOMATOS"
                    Guia_hilo.TextBox19.Text = TextBox1.Text
                    Guia_hilo.Label29.Text = Label1.Text
                    Guia_hilo.Label30.Text = Label2.Text
                    Guia_hilo.TextBox1.Text = "0001"
                    Guia_hilo.Show()
                End If
            End If
        End If

    End Sub

    Private Sub Almacen_guia_chil_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Label2.Text = "02" Then
            RadioButton1.Text = "05 - RESPUESTOS Y SUMINISTROS"
            RadioButton2.Text = "14 - ECONOMATOS"
            RadioButton3.Visible = False
            RadioButton6.Visible = False
            RadioButton5.Visible = False
            RadioButton4.Visible = False
        End If
    End Sub
End Class