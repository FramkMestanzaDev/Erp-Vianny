Imports System.Data.SqlClient
Public Class Solicitud_telas
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
    Dim Rsr21, Rsr211, Rsr213 As SqlDataReader
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            Det_Rollo.Label1.Text = "PO"
            Det_Rollo.Label2.Text = 4
            Det_Rollo.Show()

        Else
            If e.KeyCode = Keys.Enter Then
                abrir()
                Dim i5 As Integer
                i5 = 0
                Dim sql10213 As String = "select q.program_3 as PO,q.ncom_3 AS OP, rt.tela AS CODIGO,c.nomb_17 AS TELA,q.cantp_3 AS PROGRA, rt.cxplineal AS CONS_LINEAL,rt.kilos AS CONS_KG ,case when c.talm_17 ='08' then q.cantp_3*cxplineal else q.program_3*kilos end  AS CANT_SOLICITADA from custom_vianny.dbo.qag0300 q 
left join custom_vianny.dbo.Consumo_Op_Det rt on q.ccia = rt.ccia and q.ncom_3 =rt.op
inner join custom_vianny.dbo.cag1700 c on rt.tela = c.linea_17 and c.ccia = rt.ccia
where program_3 ='NTSFI00001' and q.ccia ='01'  order by ncom_3"
                Dim cmd10213 As New SqlCommand(sql10213, conx)
                Rsr213 = cmd10213.ExecuteReader()
                While Rsr213.Read()
                    DataGridView1.Rows.Add()
                    DataGridView1.Rows(i5).Cells(0).Value = Rsr213(0)
                    DataGridView1.Rows(i5).Cells(1).Value = Rsr213(1)
                    DataGridView1.Rows(i5).Cells(2).Value = Rsr213(2)
                    DataGridView1.Rows(i5).Cells(3).Value = Rsr213(3)
                    DataGridView1.Rows(i5).Cells(4).Value = Rsr213(4)
                    DataGridView1.Rows(i5).Cells(5).Value = Rsr213(5)
                    DataGridView1.Rows(i5).Cells(6).Value = Rsr213(6)
                    DataGridView1.Rows(i5).Cells(7).Value = Rsr213(7)

                    i5 = i5 + 1

                End While
                Rsr213.Close()
            End If
        End If
    End Sub
End Class