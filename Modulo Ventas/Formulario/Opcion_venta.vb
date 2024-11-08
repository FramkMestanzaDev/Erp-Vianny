Public Class Opcion_venta
    Private Sub Opcion_venta_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            Registro_Ventas.TextBox16.Text = "0001"
            Registro_Ventas.TextBox17.Text = "VENTA - MANUFACTURA"
            Registro_Ventas.TextBox13.Text = "VENTA - MANUFACTURA"
            Registro_Ventas.TextBox24.Text = RadioButton1.Text
            Registro_Ventas.Label26.Text = Label2.Text
            Registro_Ventas.Show()
        Else
            If RadioButton2.Checked = True Then
                Registro_Ventas.TextBox16.Text = "0104"
                Registro_Ventas.TextBox17.Text = "TELA MEDICA - NOTEX SMS"
                Registro_Ventas.TextBox13.Text = "TELA MEDICA - NOTEX SMS"
                Registro_Ventas.TextBox24.Text = RadioButton2.Text
                Registro_Ventas.Label26.Text = Label2.Text
                Registro_Ventas.Show()
            Else
                If RadioButton3.Checked = True Then
                    Registro_Ventas.TextBox16.Text = "0107"
                    Registro_Ventas.TextBox17.Text = "TELA MEDICA - NOTX S"
                    Registro_Ventas.TextBox13.Text = "TELA MEDICA - NOTX S"
                    Registro_Ventas.TextBox24.Text = RadioButton3.Text
                    Registro_Ventas.Label26.Text = Label2.Text
                    Registro_Ventas.Show()
                Else
                    If RadioButton4.Checked = True Then
                        Registro_Ventas.TextBox16.Text = "0103"
                        Registro_Ventas.TextBox17.Text = "INDUMENTARIA MEDICA - NOTEX SMS"
                        Registro_Ventas.TextBox13.Text = "INDUMENTARIA MEDICA - NOTEX SMS"
                        Registro_Ventas.TextBox24.Text = RadioButton4.Text
                        Registro_Ventas.Label26.Text = Label2.Text
                        Registro_Ventas.Show()
                    Else
                        If RadioButton5.Checked = True Then
                            Registro_Ventas.TextBox16.Text = "0106"
                            Registro_Ventas.TextBox17.Text = "INDUMENTARIA MEDICA - NOTEX S"
                            Registro_Ventas.TextBox13.Text = "INDUMENTARIA MEDICA - NOTEX S"
                            Registro_Ventas.TextBox24.Text = RadioButton5.Text
                            Registro_Ventas.Label26.Text = Label2.Text
                            Registro_Ventas.Show()
                        Else
                            If RadioButton6.Checked = True Then
                                Registro_Ventas.TextBox16.Text = "0108"
                                Registro_Ventas.TextBox17.Text = "ACUARIO O PARIS INDIGO"
                                Registro_Ventas.TextBox13.Text = "ACUARIO O PARIS INDIGO"
                                Registro_Ventas.TextBox24.Text = RadioButton6.Text
                                Registro_Ventas.Label26.Text = Label2.Text
                                Registro_Ventas.Show()
                            Else
                                If RadioButton7.Checked = True Then
                                    Registro_Ventas.TextBox16.Text = "0109"
                                    Registro_Ventas.TextBox17.Text = "TELA PLANA"
                                    Registro_Ventas.TextBox13.Text = "TELA PLANA"
                                    Registro_Ventas.TextBox24.Text = RadioButton7.Text
                                    Registro_Ventas.Label26.Text = Label2.Text
                                    Registro_Ventas.Show()
                                Else
                                    If RadioButton8.Checked = True Then
                                        Registro_Ventas.TextBox16.Text = "0023"
                                        Registro_Ventas.TextBox17.Text = "SERVICIO RAMA"
                                        Registro_Ventas.TextBox13.Text = "SERVICIO RAMA"
                                        Registro_Ventas.TextBox24.Text = RadioButton8.Text
                                        Registro_Ventas.Label26.Text = Label2.Text
                                        Registro_Ventas.Show()
                                    Else
                                        If RadioButton11.Checked = True Then
                                            Registro_Ventas.TextBox16.Text = "0099"
                                            Registro_Ventas.TextBox17.Text = "HILADO"
                                            Registro_Ventas.TextBox13.Text = "HILADO"
                                            Registro_Ventas.TextBox24.Text = RadioButton11.Text
                                            Registro_Ventas.Label26.Text = Label2.Text
                                            Registro_Ventas.Show()
                                        Else
                                            If RadioButton9.Checked = True Then
                                                Registro_Ventas.TextBox16.Text = "0110"
                                                Registro_Ventas.TextBox17.Text = "SERVICIO PERCHADO"
                                                Registro_Ventas.TextBox13.Text = "SERVICIO PERCHADO"
                                                Registro_Ventas.TextBox24.Text = RadioButton9.Text
                                                Registro_Ventas.Label26.Text = Label2.Text
                                                Registro_Ventas.Show()
                                            Else
                                                If RadioButton10.Checked = True Then
                                                    Registro_Ventas.TextBox16.Text = "0007"
                                                    Registro_Ventas.TextBox17.Text = "OTROS SERVICIO"
                                                    Registro_Ventas.TextBox13.Text = "OTROS SERVICIO"
                                                    Registro_Ventas.TextBox24.Text = RadioButton10.Text
                                                    Registro_Ventas.Label26.Text = Label2.Text
                                                    Registro_Ventas.Show()
                                                Else
                                                    If RadioButton12.Checked = True Then
                                                        Registro_Ventas.TextBox16.Text = "0005"
                                                        Registro_Ventas.TextBox17.Text = "ACUARIO APT"
                                                        Registro_Ventas.TextBox13.Text = "ACUARIO APT"
                                                        Registro_Ventas.TextBox24.Text = RadioButton12.Text
                                                        Registro_Ventas.Label26.Text = Label2.Text
                                                        Registro_Ventas.Show()
                                                    Else
                                                        If RadioButton13.Checked = True Then
                                                            Registro_Ventas.TextBox16.Text = "0114"
                                                            Registro_Ventas.TextBox17.Text = "PARIS APT"
                                                            Registro_Ventas.TextBox13.Text = "PARIS APT"
                                                            Registro_Ventas.TextBox24.Text = RadioButton13.Text
                                                            Registro_Ventas.Label26.Text = Label2.Text
                                                            Registro_Ventas.Show()
                                                        Else
                                                            If RadioButton14.Checked = True Then
                                                                Registro_Ventas.TextBox16.Text = "0112"
                                                                Registro_Ventas.TextBox17.Text = "JERSEY O PIQUE INDIGO"
                                                                Registro_Ventas.TextBox13.Text = "JERSEY O PIQUE INDIGO"
                                                                Registro_Ventas.TextBox24.Text = RadioButton14.Text
                                                                Registro_Ventas.Label26.Text = Label2.Text
                                                                Registro_Ventas.Show()
                                                            Else
                                                                If RadioButton15.Checked = True Then
                                                                    Registro_Ventas.TextBox16.Text = "0113"
                                                                    Registro_Ventas.TextBox17.Text = "RIB"
                                                                    Registro_Ventas.TextBox13.Text = "RIB"
                                                                    Registro_Ventas.TextBox24.Text = RadioButton15.Text
                                                                    Registro_Ventas.Label26.Text = Label2.Text
                                                                    Registro_Ventas.Show()
                                                                Else
                                                                    If RadioButton16.Checked = True Then
                                                                        Registro_Ventas.TextBox16.Text = "0115"
                                                                        Registro_Ventas.TextBox17.Text = "ACUARIO INDIGO CARDADO"
                                                                        Registro_Ventas.TextBox13.Text = "ACUARIO INDIGO CARDADO"
                                                                        Registro_Ventas.TextBox24.Text = RadioButton16.Text
                                                                        Registro_Ventas.Label26.Text = Label2.Text
                                                                        Registro_Ventas.Show()
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
                    End If
                End If
            End If
        End If

        '-----------------------------------------------------
        Registro_Ventas.Label17.Text = Label1.Text

    End Sub

    Private Sub RadioButton14_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton14.CheckedChanged

    End Sub
End Class