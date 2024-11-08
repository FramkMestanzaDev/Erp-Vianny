Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class Transportistas
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty, ty2 As New DataTable
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
        Dim i, a, conta As Integer
        i = DataGridView1.Rows.Count
        a = 0
        conta = 0
        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                conta = conta + 1
            End If
        Next
        If conta > 1 Then
            MsgBox("SOLO DEBE SELECCIONAR UNA OPCION")
        Else
            For a = 0 To i - 1
                If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                    If TextBox1.Text = 1 Then
                        Guia_remision.TextBox15.Text = DataGridView1.Rows(a).Cells(5).Value
                        Guia_remision.TextBox16.Text = DataGridView1.Rows(a).Cells(3).Value
                        Guia_remision.TextBox4.Text = DataGridView1.Rows(a).Cells(1).Value
                        Guia_remision.TextBox5.Text = DataGridView1.Rows(a).Cells(2).Value
                        Guia_remision.TextBox14.Text = DataGridView1.Rows(a).Cells(4).Value

                    Else
                        If TextBox1.Text = 2 Then
                            Guia_hilo.TextBox15.Text = DataGridView1.Rows(a).Cells(5).Value
                            Guia_hilo.TextBox16.Text = DataGridView1.Rows(a).Cells(3).Value
                            Guia_hilo.TextBox4.Text = DataGridView1.Rows(a).Cells(1).Value
                            Guia_hilo.TextBox5.Text = DataGridView1.Rows(a).Cells(2).Value
                            Guia_hilo.TextBox14.Text = DataGridView1.Rows(a).Cells(4).Value
                        Else
                            If TextBox1.Text = 3 Then
                                guia_talleres.TextBox15.Text = DataGridView1.Rows(a).Cells(5).Value
                                guia_talleres.TextBox16.Text = DataGridView1.Rows(a).Cells(3).Value
                                guia_talleres.TextBox4.Text = DataGridView1.Rows(a).Cells(1).Value
                                guia_talleres.TextBox5.Text = DataGridView1.Rows(a).Cells(2).Value
                                guia_talleres.TextBox14.Text = DataGridView1.Rows(a).Cells(4).Value
                            Else
                                If TextBox1.Text = 4 Then
                                    Guia_remision_prenda.TextBox15.Text = DataGridView1.Rows(a).Cells(5).Value
                                    Guia_remision_prenda.TextBox16.Text = DataGridView1.Rows(a).Cells(3).Value
                                    Guia_remision_prenda.TextBox4.Text = DataGridView1.Rows(a).Cells(1).Value
                                    Guia_remision_prenda.TextBox5.Text = DataGridView1.Rows(a).Cells(2).Value
                                    Guia_remision_prenda.TextBox14.Text = DataGridView1.Rows(a).Cells(4).Value
                                End If
                            End If
                        End If
                    End If



                End If
            Next
        End If

        Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Transportistas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()

        Try

            conn = New SqlDataAdapter("select trans_18 as CODI ,nomb_18 as EMPRESA,chofer_18 AS CHOFER,placa_18 AS PLACA,brevete_18 AS BREVETE,marca_18 AS MARCA from custom_vianny.dbo.Fag1800 where empr_18 ='01' ", conx)
            conn.Fill(ty)
            DataGridView1.DataSource = ty
            DataGridView1.Columns(1).Visible = True
            DataGridView1.Columns(1).Width = 50
            DataGridView1.Columns(2).Width = 200
            DataGridView1.Columns(3).Width = 200
            'respuesta = enunciado.ExecuteReader
            'While respuesta.Read
            '    cb.Items.Add(respuesta.Item("nomb_17f"))
            'End While
            'respuesta.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class