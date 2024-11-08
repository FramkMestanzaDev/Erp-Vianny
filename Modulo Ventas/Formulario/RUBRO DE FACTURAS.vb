Public Class RUBRO_DE_FACTURAS
    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = True Then
            Registro_Facturas.Label4.Text = "0001"
            Registro_Facturas.TextBox13.Text = RadioButton1.Text
            Registro_Facturas.Label27.Text = Label3.Text
            Registro_Facturas.TextBox13.Text = RadioButton1.Text
        Else
            If RadioButton2.Checked = True Then
                Registro_Facturas.Label4.Text = "0104"
                Registro_Facturas.TextBox13.Text = RadioButton2.Text
                Registro_Facturas.Label27.Text = Label3.Text
                Registro_Facturas.TextBox13.Text = RadioButton2.Text
            Else
                If RadioButton3.Checked = True Then
                    Registro_Facturas.Label4.Text = "0107"
                    Registro_Facturas.TextBox13.Text = RadioButton3.Text
                    Registro_Facturas.Label27.Text = Label3.Text
                    Registro_Facturas.TextBox13.Text = RadioButton3.Text
                Else
                    If RadioButton4.Checked = True Then
                        Registro_Facturas.Label4.Text = "0103"
                        Registro_Facturas.TextBox13.Text = RadioButton4.Text
                        Registro_Facturas.Label27.Text = Label3.Text
                        Registro_Facturas.TextBox13.Text = RadioButton4.Text
                    Else
                        If RadioButton5.Checked = True Then
                            Registro_Facturas.Label4.Text = "0106"
                            Registro_Facturas.TextBox13.Text = RadioButton5.Text
                            Registro_Facturas.Label27.Text = Label3.Text
                            Registro_Facturas.TextBox13.Text = RadioButton5.Text
                        Else
                            If RadioButton6.Checked = True Then
                                Registro_Facturas.Label4.Text = "0108"
                                Registro_Facturas.TextBox13.Text = RadioButton6.Text
                                Registro_Facturas.Label27.Text = Label3.Text
                                Registro_Facturas.TextBox13.Text = RadioButton6.Text
                            Else
                                If RadioButton7.Checked = True Then
                                    Registro_Facturas.Label4.Text = "0109"
                                    Registro_Facturas.TextBox13.Text = RadioButton7.Text
                                    Registro_Facturas.Label27.Text = Label3.Text
                                    Registro_Facturas.TextBox13.Text = RadioButton7.Text
                                Else
                                    If RadioButton8.Checked = True Then
                                        Registro_Facturas.Label4.Text = "0099"
                                        Registro_Facturas.TextBox13.Text = RadioButton8.Text
                                        Registro_Facturas.Label27.Text = Label3.Text
                                        Registro_Facturas.TextBox13.Text = RadioButton8.Text
                                    Else
                                        If RadioButton11.Checked = True Then
                                            Registro_Facturas.Label4.Text = "0023"
                                            Registro_Facturas.TextBox13.Text = RadioButton11.Text
                                            Registro_Facturas.Label27.Text = Label3.Text
                                            Registro_Facturas.TextBox13.Text = RadioButton11.Text
                                        Else
                                            If RadioButton9.Checked = True Then
                                                Registro_Facturas.Label4.Text = "0110"
                                                Registro_Facturas.TextBox13.Text = RadioButton9.Text
                                                Registro_Facturas.Label27.Text = Label3.Text
                                                Registro_Facturas.TextBox13.Text = RadioButton9.Text
                                            Else
                                                If RadioButton10.Checked = True Then
                                                    Registro_Facturas.Label4.Text = "0007"
                                                    Registro_Facturas.TextBox13.Text = RadioButton10.Text
                                                    Registro_Facturas.Label27.Text = Label3.Text
                                                    Registro_Facturas.TextBox13.Text = RadioButton10.Text
                                                Else
                                                    If RadioButton12.Checked = True Then
                                                        Registro_Facturas.Label4.Text = "0005"
                                                        Registro_Facturas.TextBox13.Text = RadioButton12.Text
                                                        Registro_Facturas.Label27.Text = Label3.Text
                                                        Registro_Facturas.TextBox13.Text = RadioButton12.Text
                                                    Else
                                                        If RadioButton13.Checked = True Then
                                                            Registro_Facturas.Label4.Text = "0114"
                                                            Registro_Facturas.TextBox13.Text = RadioButton13.Text
                                                            Registro_Facturas.Label27.Text = Label3.Text
                                                            Registro_Facturas.TextBox13.Text = RadioButton13.Text
                                                        Else
                                                            If RadioButton14.Checked = True Then
                                                                Registro_Facturas.Label4.Text = "0112"
                                                                Registro_Facturas.TextBox13.Text = RadioButton14.Text
                                                                Registro_Facturas.Label27.Text = Label3.Text
                                                                Registro_Facturas.TextBox13.Text = RadioButton14.Text
                                                            Else
                                                                If RadioButton15.Checked = True Then
                                                                    Registro_Facturas.Label4.Text = "0113"
                                                                    Registro_Facturas.TextBox13.Text = RadioButton15.Text
                                                                    Registro_Facturas.Label27.Text = Label3.Text
                                                                    Registro_Facturas.TextBox13.Text = RadioButton15.Text
                                                                Else
                                                                    If RadioButton16.Checked = True Then
                                                                        Registro_Facturas.Label4.Text = "0115"
                                                                        Registro_Facturas.TextBox13.Text = RadioButton16.Text
                                                                        Registro_Facturas.Label27.Text = Label3.Text
                                                                        Registro_Facturas.TextBox13.Text = RadioButton16.Text
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
        Registro_Facturas.TextBox8.Text = "ALMACEN" & "" & Label1.Text & "-  REGISTRO DE VENTA  "
        Registro_Facturas.Label5.Text = Label1.Text
        Registro_Facturas.Label29.Text = Label4.Text
        Registro_Facturas.TextBox10.Text = Label1.Text
        Registro_Facturas.Show()
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

    End Sub

    Private Sub RUBRO_DE_FACTURAS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class