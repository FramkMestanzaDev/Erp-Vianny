Public Class Almacen
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Label2.Text = "01" Then

            If RadioButton1.Checked = True Then
                'RUBRO_DE_FACTURAS.Label1.Text = "37"
                'RUBRO_DE_FACTURAS.Label2.Text = RadioButton1.Text
                'RUBRO_DE_FACTURAS.Label3.Text = Label1.Text
                'RUBRO_DE_FACTURAS.Label4.Text = Label2.Text
                'RUBRO_DE_FACTURAS.Show()

                rubro2.Label4.Text = "37"

                rubro2.Label3.Text = Label1.Text
                rubro2.Label1.Text = Label2.Text
                rubro2.Show()
            Else
                If RadioButton2.Checked = True Then
                    RUBRO_DE_FACTURAS.Label1.Text = "38"
                    RUBRO_DE_FACTURAS.Label2.Text = RadioButton2.Text
                    RUBRO_DE_FACTURAS.Label3.Text = Label1.Text
                    RUBRO_DE_FACTURAS.Label4.Text = Label2.Text
                    RUBRO_DE_FACTURAS.Show()
                Else
                    If RadioButton3.Checked = True Then
                        Registro_Facturas.TextBox8.Text = "ALMACEN" & " " & "61" & "-  MUESTRAS "
                        Registro_Facturas.Label5.Text = "61"
                        Registro_Facturas.TextBox10.Text = "61"
                        Registro_Facturas.TextBox16.Text = "32"
                        Registro_Facturas.TextBox17.Text = "FACTURAS/BOLETAS SIN VALOR COMERCIAL"
                        Registro_Facturas.TextBox23.Text = "MUESTRA SIN VALOR COMERCIAL"
                        Registro_Facturas.TextBox20.Text = "69"
                        Registro_Facturas.TextBox13.Text = "MUESTRA SIN VALOR COMERCIAL"
                        Registro_Facturas.ComboBox6.Text = " "
                        RUBRO_DE_FACTURAS.Label3.Text = Label1.Text
                        Registro_Facturas.Label29.Text = Label2.Text
                        Registro_Facturas.Label27.Text = Label1.Text
                        Registro_Facturas.TextBox34.Text = 0
                        Registro_Facturas.Button6.Enabled = False
                        Registro_Facturas.Show()
                    Else
                        If RadioButton4.Checked = True Then
                            RUBRO_DE_FACTURAS.Label1.Text = "10"
                            RUBRO_DE_FACTURAS.Label2.Text = RadioButton4.Text
                            RUBRO_DE_FACTURAS.Label3.Text = Label1.Text
                            RUBRO_DE_FACTURAS.Label4.Text = Label2.Text
                            RUBRO_DE_FACTURAS.Show()
                        Else
                            If RadioButton5.Checked = True Then
                                RUBRO_DE_FACTURAS.Label1.Text = "51"
                                RUBRO_DE_FACTURAS.Label2.Text = RadioButton5.Text
                                RUBRO_DE_FACTURAS.Label3.Text = Label1.Text
                                RUBRO_DE_FACTURAS.Label4.Text = Label2.Text
                                RUBRO_DE_FACTURAS.Show()
                            Else
                                If RadioButton6.Checked = True Then
                                    RUBRO_DE_FACTURAS.Label1.Text = "42"
                                    RUBRO_DE_FACTURAS.Label2.Text = RadioButton5.Text
                                    RUBRO_DE_FACTURAS.Label3.Text = Label1.Text
                                    RUBRO_DE_FACTURAS.Label4.Text = Label2.Text
                                    RUBRO_DE_FACTURAS.Show()
                                Else
                                    If RadioButton7.Checked = True Then
                                        Registro_Facturas.TextBox8.Text = "ALMACEN" & " " & "59" & "-  SERVICIOS "
                                        Registro_Facturas.TextBox13.Text = "SERVICIO"
                                        Registro_Facturas.Label5.Text = "59"
                                        Registro_Facturas.TextBox10.Text = "59"
                                        Registro_Facturas.Label29.Text = Label2.Text
                                        Registro_Facturas.Label27.Text = Label1.Text
                                        Registro_Facturas.TextBox21.Text = "F013"
                                        Registro_Facturas.Label4.Text = "0007"
                                        Registro_Facturas.Show()
                                    Else
                                        If RadioButton8.Checked = True Then
                                            RUBRO_DE_FACTURAS.Label1.Text = "68"
                                            RUBRO_DE_FACTURAS.Label2.Text = RadioButton5.Text
                                            RUBRO_DE_FACTURAS.Label3.Text = Label1.Text
                                            RUBRO_DE_FACTURAS.Label4.Text = Label2.Text
                                            RUBRO_DE_FACTURAS.Show()
                                        Else
                                            If RadioButton9.Checked = True Then
                                                Registro_Facturas.TextBox8.Text = "ALMACEN" & " " & "03" & "-  HILADO CRUDO "
                                                Registro_Facturas.TextBox13.Text = "HILADO CRUDO"
                                                Registro_Facturas.Label5.Text = "03"
                                                Registro_Facturas.TextBox10.Text = "03"
                                                Registro_Facturas.Label29.Text = Label2.Text
                                                Registro_Facturas.Label27.Text = Label1.Text
                                                Registro_Facturas.TextBox21.Text = "F015"
                                                Registro_Facturas.Label4.Text = "0116"
                                                Registro_Facturas.Show()
                                            Else

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
                RUBRO_GRAUS.Label1.Text = "67"
                RUBRO_GRAUS.Label2.Text = RadioButton1.Text
                RUBRO_GRAUS.Label3.Text = Label1.Text
                RUBRO_GRAUS.Label4.Text = Label2.Text
                RUBRO_GRAUS.Show()
            End If
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

    End Sub

    Private Sub Almacen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Label2.Text = "02" Then
            RadioButton2.Visible = False
            RadioButton3.Visible = False
            RadioButton4.Visible = False
            RadioButton5.Visible = False
            RadioButton6.Visible = False
        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged

    End Sub
End Class