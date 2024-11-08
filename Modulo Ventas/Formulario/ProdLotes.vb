Imports System.Data.SqlClient
Public Class ProdLotes
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim Rsr21 As SqlDataReader
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Dim da1 As New DataTable
    Private Sub ProdLotes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrar()

    End Sub
    Sub mostrar()
        da1.Rows.Clear()

        abrir()
        Dim TELA As String
        If Label5.Text = "TELA PLANA" Then
            TELA = "08"
        Else
            TELA = "10"
        End If

        Dim cmd As New SqlDataAdapter("EXEC CaeSoft_ReporteStockFisico_38 '" + Label1.Text.ToString.Trim + "','" + TELA + "','" + Label4.Text.ToString.Trim + "',NULL, 2,'" + Label2.Text.ToString.Trim + "'", conx)
        cmd.Fill(da1)
        If da1.Rows.Count > 0 Then
            DataGridView1.DataSource = da1
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).ReadOnly = True
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(8).ReadOnly = True
            DataGridView1.Columns(9).ReadOnly = True
            DataGridView1.Columns(10).ReadOnly = True
            DataGridView1.Columns(11).ReadOnly = True
            DataGridView1.Columns(5).Width = 145
            DataGridView1.Columns(7).Width = 80
            DataGridView1.Columns(9).Width = 350

            For i = 0 To DataGridView1.Rows.Count - 1

                Dim sql1021 As String = "select isnull(CanDtm,0) from DetalleTelaManufactura where OpDtm ='" + Label7.Text.ToString.Trim + "' and CopDtm ='" + DataGridView1.Rows(i).Cells(5).Value.ToString.Trim + "' and LoPDtm ='" + DataGridView1.Rows(i).Cells(7).Value.ToString.Trim + "'"
                Dim cmd1021 As New SqlCommand(sql1021, conx)
                Rsr21 = cmd1021.ExecuteReader()
                If Rsr21.Read() Then
                    DataGridView1.Rows(i).Cells(6).Value = Rsr21(0)

                End If
                Rsr21.Close()

            Next
        End If
    End Sub
    Function ELIMINAR(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim suma, suma1 As Double

        abrir()
        suma = 0
        For p = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(p).Cells(0).Value = True Then
                suma = suma + Convert.ToDouble(DataGridView1.Rows(p).Cells(6).Value)
                suma1 = suma1 + Convert.ToDouble(DataGridView1.Rows(p).Cells(6).Value)
            End If
        Next
        'If suma > Convert.ToDouble(TextBox1.Text) Then
        '    MsgBox("LA CANTIDAD INGRESADA EXCEDE A LO QUE SE SOLICITA UDP PARA DESPCHAR")
        'Else
        INSERTAR(1)
            MsgBox("Se Registro la Informacion Correctamnete")
            DetTelaProduccion.MOSTRAR()
            Me.Close()
        'End If



    End Sub
    Sub INSERTAR(ESTADO)
        Dim num As Integer
        num = DataGridView1.Rows.Count
        For i = 0 To num - 1
            If DataGridView1.Rows(i).Cells(0).Value = True Then
                Dim agregar As String = "DELETE DetalleTelaManufactura where  OpDtm ='" + Label7.Text.ToString + "' and CopDtm ='" + Label4.Text.ToString + "' and LoPDtm = '" + DataGridView1.Rows(i).Cells(7).Value.ToString.Trim + "'"
                ELIMINAR(agregar)
                Dim cmd20 As New SqlCommand("insert into DetalleTelaManufactura (ItmDtm,OpDtm,CopDtm,NomDtm,DetDtm,CanDtm,FamDtm,EstDtm,LoPDtm,CalDtm) values (@ItmDtm,@OpDtm,@CopDtm,@NomDtm,@DetDtm,@CanDtm,@FamDtm,@EstDtm,@LoPDtm,@CalDtm)", conx)
                cmd20.Parameters.AddWithValue("@ItmDtm", i)
                cmd20.Parameters.AddWithValue("@OpDtm", Label7.Text.ToString)
                cmd20.Parameters.AddWithValue("@CopDtm", Label4.Text.ToString)
                cmd20.Parameters.AddWithValue("@NomDtm", Trim(DataGridView1.Rows(i).Cells(9).Value.ToString))
                cmd20.Parameters.AddWithValue("@DetDtm", Label10.Text.ToString)
                cmd20.Parameters.AddWithValue("@CanDtm", Convert.ToDouble(DataGridView1.Rows(i).Cells(6).Value))
                cmd20.Parameters.AddWithValue("@FamDtm", Trim(DataGridView1.Rows(i).Cells(8).Value.ToString))
                cmd20.Parameters.AddWithValue("@EstDtm", ESTADO)
                cmd20.Parameters.AddWithValue("@LoPDtm", Trim(DataGridView1.Rows(i).Cells(7).Value.ToString))
                cmd20.Parameters.AddWithValue("@CalDtm", Trim(DataGridView1.Rows(i).Cells(12).Value.ToString))
                cmd20.ExecuteNonQuery()
            End If
        Next
    End Sub


    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim isCellChecked As Boolean = Convert.ToBoolean(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
        If isCellChecked Then
            DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = False
            'DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(6)
            DataGridView1.BeginEdit(True)
            DataGridView1.Rows(e.RowIndex).Cells(6).Selected = True
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If DataGridView1.Columns(e.ColumnIndex).HeaderText = "Despacho" Then
            If Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(6).Value) <= Convert.ToDouble(DataGridView1.Rows(e.RowIndex).Cells(11).Value) Then
            Else
                MsgBox("LA CANTIDAD SOLICITADA ES MAYOR AL STOCK")
                DataGridView1.Rows(e.RowIndex).Cells(6).Value = 0
                DataGridView1.Rows(e.RowIndex).Cells(0).Value = False
            End If
        End If
    End Sub
End Class