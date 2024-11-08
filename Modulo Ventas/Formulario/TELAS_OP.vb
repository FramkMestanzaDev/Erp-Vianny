Imports System.Data.SqlClient
Public Class TELAS_OP
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

    Dim Rsr21 As SqlDataReader
    Private Sub TELAS_OP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        DataGridView1.Rows.Clear()
        Dim sql1021 As String = "EXEC custom_vianny.dbo.CaeSoft_GetallTelaOP '01',NULL,1,'" + Trim(Label2.Text) + "'"
        Dim cmd1021 As New SqlCommand(sql1021, conx)
        Rsr21 = cmd1021.ExecuteReader()
        Dim i51 As Integer
        i51 = 0
        While Rsr21.Read
            DataGridView1.Rows.Add()

            DataGridView1.Rows(i51).Cells(1).Value = Rsr21(0)
            DataGridView1.Rows(i51).Cells(2).Value = Rsr21(4)
            DataGridView1.Rows(i51).Cells(3).Value = Rsr21(27)
            DataGridView1.Rows(i51).Cells(4).Value = Rsr21(5)
            DataGridView1.Rows(i51).Cells(5).Value = Rsr21(6)
            DataGridView1.Rows(i51).Cells(6).Value = Rsr21(7)
            DataGridView1.Rows(i51).Cells(7).Value = Rsr21(13)
            DataGridView1.Rows(i51).Cells(8).Value = Rsr21(12)
            DataGridView1.Rows(i51).Cells(9).Value = Rsr21(14)
            DataGridView1.Rows(i51).Cells(10).Value = Rsr21(16)
            DataGridView1.Rows(i51).Cells(11).Value = Rsr21(15)
            DataGridView1.Rows(i51).Cells(12).Value = Rsr21(17)
            DataGridView1.Rows(i51).Cells(13).Value = Rsr21(2)
            DataGridView1.Rows(i51).Cells(14).Value = Rsr21(19)
            DataGridView1.Rows(i51).Cells(15).Value = Rsr21(22)
            DataGridView1.Rows(i51).Cells(16).Value = Rsr21(25)
            i51 = i51 + 1
        End While

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i, a, I1 As Integer
        i = DataGridView1.Rows.Count

        For a = 0 To i - 1
            If Me.DataGridView1.Rows(a).Cells(0).Value = True Then
                I1 = FICHA_CONSUMO.DataGridView1.Rows.Count
                FICHA_CONSUMO.DataGridView1.Rows.Add()
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(0).Value = I1 + 1
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(1).Value = Trim(DataGridView1.Rows(a).Cells(1).Value)

                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(3).Value = (DataGridView1.Rows(a).Cells(7).Value) / 2
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(6).Value = (DataGridView1.Rows(a).Cells(8).Value)
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(4).Value = (DataGridView1.Rows(a).Cells(7).Value)
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(7).Value = 0
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(8).Value = 0
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(9).Value = 0
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(5).Value = 0
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(10).Value = 0
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(11).Value = 0
                FICHA_CONSUMO.DataGridView1.Rows(I1).Cells(12).Value = 0
            End If
        Next
        FICHA_CONSUMO.DataGridView1.ReadOnly = False
        Close()
    End Sub
End Class