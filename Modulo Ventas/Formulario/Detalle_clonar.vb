Imports System.Data.SqlClient
Public Class Detalle_clonar

    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim Rsr2, Rsr3, Rsr212, R25 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Detalle_clonar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
    End Sub

    Dim ll, FTY As DataTable
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim gh As New fop
            Dim kk As New vop
            Dim lp As Integer
            kk.gncom_3 = Trim(TextBox1.Text) & "1"
            kk.gcia = "01"
            ll = gh.buscar_op(kk)

            OD.DataGridView3.DataSource = ll
            lp = OD.DataGridView3.Rows.Count

            If lp <> 0 Then
                For i = 0 To lp - 1

                    If Trim(OD.TextBox78.Text) = Trim(OD.DataGridView3.Rows(0).Cells(44).Value) Or Trim(MDIParent1.Label4.Text) = "ADMINISTRADOR" Then

                        OD.TextBox58.Text = OD.DataGridView3.Rows(0).Cells(7).Value
                        OD.TextBox1.Text = OD.DataGridView3.Rows(0).Cells(8).Value

                        OD.TextBox81.Text = Microsoft.VisualBasic.Mid(Trim(OD.DataGridView3.Rows(0).Cells(4).Value), 2, 5)
                        OD.TextBox4.Text = OD.DataGridView3.Rows(0).Cells(10).Value
                        OD.TextBox5.Text = OD.DataGridView3.Rows(0).Cells(11).Value
                        OD.TextBox9.Text = OD.DataGridView3.Rows(0).Cells(81).Value

                        OD.TextBox78.Text = OD.DataGridView3.Rows(0).Cells(44).Value
                        Select Case OD.TextBox78.Text
                            Case "0001" : OD.TextBox15.Text = "VSILVERIO"
                            Case "0003" : OD.TextBox15.Text = "GBALVIN"
                            Case "0025" : OD.TextBox15.Text = "VGRAUS"
                        End Select
                        OD.TextBox10.Text = OD.DataGridView3.Rows(0).Cells(9).Value
                        OD.TextBox70.Text = OD.DataGridView3.Rows(0).Cells(45).Value
                        OD.TextBox71.Text = OD.DataGridView3.Rows(0).Cells(84).Value
                        OD.TextBox73.Text = OD.DataGridView3.Rows(0).Cells(79).Value
                        OD.TextBox74.Text = OD.DataGridView3.Rows(0).Cells(80).Value
                        OD.TextBox75.Text = OD.DataGridView3.Rows(0).Cells(82).Value
                        OD.TextBox76.Text = OD.DataGridView3.Rows(0).Cells(83).Value
                        OD.TextBox19.Text = OD.DataGridView3.Rows(0).Cells(24).Value


                        OD.TextBox12.Text = Trim(OD.DataGridView3.Rows(0).Cells(17).Value)
                        OD.TextBox39.Text = Trim(OD.DataGridView3.Rows(0).Cells(17).Value)

                        OD.TextBox11.Text = OD.DataGridView3.Rows(0).Cells(53).Value
                        OD.TextBox79.Text = OD.DataGridView3.Rows(0).Cells(43).Value
                        OD.TextBox16.Text = OD.DataGridView3.Rows(0).Cells(52).Value

                        OD.TextBox13.Text = Trim(OD.DataGridView3.Rows(0).Cells(51).Value)

                        OD.TextBox3.Text = Trim(OD.DataGridView3.Rows(0).Cells(26).Value)
                        OD.TextBox72.Text = Trim(OD.DataGridView3.Rows(0).Cells(25).Value)
                        OD.TextBox21.Text = Trim(OD.DataGridView3.Rows(0).Cells(42).Value)
                        OD.TextBox18.Text = OD.DataGridView3.Rows(0).Cells(85).Value

                        OD.TextBox59.Text = Trim(OD.DataGridView3.Rows(0).Cells(49).Value)
                        OD.TextBox6.Text = Trim(OD.DataGridView3.Rows(0).Cells(48).Value)
                        If Trim(OD.DataGridView3.Rows(0).Cells(54).Value) = "N" Then
                            OD.ComboBox1.SelectedIndex = 0
                        Else
                            OD.ComboBox1.SelectedIndex = 1
                        End If
                        'Se genera el correlativo de od


                        Dim corela As Integer
                        corela = 0
                        abrir()
                        Dim bpo As String
                        bpo = "M" & Trim(Microsoft.VisualBasic.Mid(Trim(OD.DataGridView3.Rows(0).Cells(4).Value), 2, 5))
                        Dim sql102 As String = "select top 1 program_3, substring(program_3,7,4) as ncom from custom_vianny.dbo.qag0300  where  program_3 like  '" + bpo + "%' and ccia='01'  order by ncom_3 desc"
                        Dim cmd102 As New SqlCommand(sql102, conx)
                        R25 = cmd102.ExecuteReader()
                        If R25.Read() Then
                            corela = Convert.ToInt32(R25(1)) + 1
                        Else
                            corela = 1
                        End If
                        R25.close
                        Select Case Convert.ToString(corela).Length
                            Case "1" : OD.TextBox8.Text = "000" & corela
                            Case "2" : OD.TextBox8.Text = "00" & corela
                            Case "3" : OD.TextBox8.Text = "0" & corela
                            Case "4" : OD.TextBox8.Text = corela



                        End Select
                        ' OD.TextBox8.Text = Microsoft.VisualBasic.Mid(Trim(OD.DataGridView3.Rows(0).Cells(4).Value), 7, 4)

                        ' fon de correaltivo

                        Dim ab As New vtallas
                        Dim ab1 As New Padron_tallas
                        ab.gcodigo = Trim(OD.TextBox13.Text)
                        ab.gccia = "01"
                        FTY = ab1.mostrar_padrom_tallas_SELECCIONADO(ab)
                        OD.DataGridView5.DataSource = FTY
                        OD.DataGridView4.DataSource = FTY
                        OD.TextBox20.Text = Trim(OD.DataGridView5.Rows(0).Cells(13).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(14).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(15).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(16).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(17).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(18).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(19).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(20).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(21).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(22).Value) & "/" & Trim(OD.DataGridView5.Rows(0).Cells(23).Value)
                    End If
                Next
            End If
            Me.Close()
        Else
            If e.KeyCode = Keys.F1 Then
                ListaOD.Label4.Text = 1
                ListaOD.ShowDialog()
            End If
        End If

    End Sub
End Class