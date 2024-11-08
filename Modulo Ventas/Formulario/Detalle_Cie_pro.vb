Imports System.Data.SqlClient
Public Class Detalle_Cie_pro
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

    Dim da As New DataTable
    Sub mostrar_informacion()
        abrir()
        da.Clear()
        DataGridView1.DataSource = Nothing
        Dim cmd As New SqlDataAdapter("SELECT  EtaRut,NomRut as PROCESO,fechRut as FEcha, AsiRut as ASIGNACION, ObsRut as OBSERVACION,CiRut FROM  Ruta_Udp WHERE OdRut ='" + Trim(TextBox1.Text) + "'", conx)
        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(2).Width = 180
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).Width = 240
        DataGridView1.Columns(5).Width = 240
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(6).Visible = False
        Dim p1 As Integer
        p1 = DataGridView1.Rows.Count

        If p1 > 0 Then
            For i1 = 0 To p1 - 1

                If Trim(DataGridView1.Rows(i1).Cells(6).Value) = "1" Then

                    DataGridView1.Rows(i1).Cells(0).Value = True
                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.Black
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.White
                Else

                    DataGridView1.Rows(i1).DefaultCellStyle.BackColor = Color.White
                    DataGridView1.Rows(i1).DefaultCellStyle.ForeColor = Color.Black
                End If
            Next

        End If

    End Sub

    Private Sub Detalle_Cie_pro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = ""
        da.Clear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim p As Integer
        p = DataGridView1.Rows.Count
        If p > 0 Then
            For i = 0 To p - 1
                If DataGridView1.Rows(i).Cells(0).Value = True Then
                    Dim cmd20 As New SqlCommand("UPDATE Ruta_Udp SET CiRut =1 where OdRut =@OdRut and EtaRut =@EtaRut", conx)
                    cmd20.Parameters.AddWithValue("@OdRut", TextBox1.Text)
                    cmd20.Parameters.AddWithValue("@EtaRut", Trim(DataGridView1.Rows(i).Cells(1).Value))
                    cmd20.ExecuteNonQuery()
                End If
            Next
            MsgBox("SE CERRO EL PROCESO")
            mostrar_informacion()
        Else
            MsgBox("NO HAY INFORMACION PARA CERRAR")
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        mostrar_informacion()
    End Sub
End Class