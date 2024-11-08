Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class TRABAJADORES
    Public enunciado2 As SqlCommand
    Public respuesta2 As SqlDataAdapter
    Public conx As SqlConnection
    Dim da As New DataTable
    Private Sub buscar()
        Try
            Dim ds As New DataSet
            ds.Tables.Add(da.Copy)
            Dim dv As New DataView(ds.Tables(0))


            dv.RowFilter = "TRABAJADOR" & " like '%" & TextBox1.Text & "%'"

            If dv.Count <> 0 Then

                DataGridView1.DataSource = dv

                DataGridView1.Columns(1).Width = 400

            Else

                DataGridView1.DataSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub TRABAJADORES_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        da.Clear()
        abrir()
        'enunciado2 = New SqlCommand("SELECT ruc_10 AS DNI,nomb_10 AS TRABAJADOR FROM custom_vianny.dbo.cag1000 WHERE tdeta_10 ='TR' AND activo_10='1' AND tpla_10='C'", conx)
        'respuesta2 = enunciado2.ExecuteNonQuery
        'While respuesta2.Read
        '    DataGridView1.DataSource = respuesta2.Item("ruc_10")
        'End While
        'respuesta2.Close()
        Dim cmd As New SqlDataAdapter("SELECT ruc_10 AS DNI,nomb_10 AS TRABAJADOR FROM custom_vianny.dbo.cag1000 WHERE tdeta_10 ='TR' AND activo_10='1' ", conx)



        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(1).Width = 400
        TextBox1.Select()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If Label2.Text = 1 Then
            TEJEDORES.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
            TEJEDORES.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
        Else
            If Label2.Text = 3 Then
                Memorándum.TextBox3.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
            Else
                If Label2.Text = 4 Then
                    Memorándum.TextBox5.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                Else
                    If Label2.Text = 5 Then
                        Planilla_Movilidas.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                        Planilla_Movilidas.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                    Else
                        If Label2.Text = 6 Then
                            Liqui_Produccion.TextBox11.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                            Liqui_Produccion.TextBox12.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                        Else
                            If Label2.Text = 7 Then
                                Liqui_Produccion.TextBox13.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                                Liqui_Produccion.TextBox14.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                            Else
                                AUDITORES.TextBox1.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                                AUDITORES.TextBox2.Text = Trim(DataGridView1.Rows(e.RowIndex).Cells(1).Value)
                            End If

                        End If

                    End If

                End If

            End If

        End If

        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        buscar()
    End Sub
End Class