Imports System.Data.SqlClient

Public Class RequerimientoPo
    Public conx As SqlConnection
    Public comando As SqlCommand
    Dim da As New DataTable
    Public Event PasarItems(ByVal items As List(Of Item))
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub RequerimientoPo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mostrarinformacion()
        Me.Activate()  ' Trae el formulario al frente
        Me.Focus()
    End Sub
    Sub mostrarinformacion()
        abrir()
        Dim cmd As New SqlDataAdapter("EXEC custom_vianny.dbo.CaeSoft_StatusRequerimientoLogisticoxPedido '" + Label1.Text + "' , '" + Label3.Text + "', '" + Label4.Text + "', '1'", conx)
        cmd.Fill(da)
        If da.Rows.Count > 0 Then
            DataGridView1.DataSource = da
            DataGridView1.Columns(1).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(2).Width = 85
            DataGridView1.Columns(3).Width = 148
            DataGridView1.Columns(4).Width = 500
            DataGridView1.Columns(6).Width = 40
            DataGridView1.Columns(7).Width = 90
            DataGridView1.Columns(9).Width = 90
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).ReadOnly = True
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(8).ReadOnly = True
            DataGridView1.Columns(9).ReadOnly = True
        Else
            DataGridView1.DataSource = Nothing
        End If

    End Sub

    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        ' Verifica cuando el CheckBox cambia y se confirma el cambio
        If TypeOf DataGridView1.CurrentCell Is DataGridViewCheckBoxCell Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        ' Verifica si la columna es la del CheckBox
        If e.ColumnIndex = 0 Then ' Suponiendo que la columna 0 es el CheckBox
            ' Verifica si el índice de la fila es válido
            If e.RowIndex >= 0 AndAlso e.RowIndex < DataGridView1.Rows.Count Then
                Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

                ' Verifica si el CheckBox está marcado
                If Convert.ToBoolean(row.Cells(0).Value) = True Then
                    row.DefaultCellStyle.BackColor = Color.Yellow ' Cambia a color amarillo
                Else
                    row.DefaultCellStyle.BackColor = Color.White ' Restaura al color original
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim selectedItems As New List(Of Item)
        ' Recorre las filas del DataGridView
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Verifica si la fila está seleccionada (por ejemplo, si hay un checkbox para seleccionar)
            If Convert.ToBoolean(row.Cells("S").Value) = True Then
                ' Crear un objeto Item y agregarlo a la lista
                Dim item As New Item With {
                    .Po = row.Cells("Pedido").Value.ToString(),
                    .Codigo = row.Cells("Linea").Value.ToString(),
                    .Descripcion = row.Cells("Nombre").Value.ToString(),
                    .um = row.Cells("UM").Value.ToString(),
                    .Cantidad = Math.Abs(Convert.ToDouble(row.Cells("CantPend").Value))
                }
                selectedItems.Add(item)
            End If
        Next
        'MsgBox(selectedItems.Count)
        ' Disparar el evento PasarItems y pasar la lista de ítems seleccionados
        RaiseEvent PasarItems(selectedItems)
        ' Cerrar el formulario secundario
        Me.Close()
    End Sub

End Class