Imports System.Data.SqlClient
Public Class OBSERVACION
    Public conx As SqlConnection
    Public conn As SqlDataAdapter
    Dim ty2, TY3 As New DataTable
    Public comando As SqlCommand
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub OBSERVACION_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Label2.Text = 0 Then

            Dim dh As New vop
            Dim jk As New fop
            dh.gncom_3 = TextBox2.Text
            dh.gcia = Label1.Text
            TextBox1.Text = jk.buscar_observacion(dh)

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Label2.Text = 0 Then
            Dim dh As New vop
            Dim jk As New fop

            dh.gncom_3 = TextBox2.Text
            dh.goobservacion2 = TextBox1.Text
            dh.gcia = Label1.Text
            jk.actualizar_op_obser(dh)
            MsgBox("Se Agrego la Informacio a la Op")
            TextBox1.Text = ""
            TextBox2.Text = ""
            Programa_Tejeduria.mostrar23()
            Me.Close()
        Else
            If Label2.Text = 3 Then
                guia_talleres.DataGridView1.Rows(Label3.Text).Cells(11).Value = Trim(TextBox1.Text)
                MsgBox("SE GUARDO LA INFORMACION SOLICITADA")
                Me.Close()
                TextBox1.Text = ""
            Else

                abrir()
                    Dim cmd18 As New SqlCommand("update facturas_eliminadas set observacion =@observacion where id_feliminadas =@id", conx)
                    cmd18.Parameters.AddWithValue("@observacion", Trim(TextBox1.Text))
                    cmd18.Parameters.AddWithValue("@id", Trim(Label4.Text))
                    cmd18.ExecuteNonQuery()
                    Otras_Facturas.DataGridView1.Rows(Label3.Text).Cells(0).Value = Trim(TextBox1.Text)
                    MsgBox("SE GUARDO LA INFORMACION SOLICITADA")
                    Me.Close()
                    TextBox1.Text = ""

            End If

        End If


    End Sub
End Class