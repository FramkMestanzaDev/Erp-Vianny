Imports System.Data.SqlClient
Public Class TEJEDORES
    Public conx As SqlConnection
    Public comando As SqlCommand
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            TRABAJADORES.Label2.Text = 1
            TRABAJADORES.Show()
        End If
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
    Function insertar(ByVal sql)
        abrir()
        comando = New SqlCommand(sql, conx)
        Dim i As Integer = comando.ExecuteNonQuery
        If i > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
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
        abrir()
        Dim cmd1 As New SqlDataAdapter("SELECT * FROM tejedores WHERE dni ='" + TextBox1.Text + "'", conx)
        Dim da1 As New DataTable
        cmd1.Fill(da1)

        If da1.Rows.Count > 0 Then
            MsgBox("EL TRABAJADOR YA ESTA REGISTRADO")
        Else
            Dim agregar As String = "INSERT INTO tejedores (dni,trabajador) VALUES('" + TextBox1.Text + "','" + TextBox2.Text + "')"
            If (insertar(agregar)) Then
                MessageBox.Show("DATOS INGRESADOS CORRECTAMENTE")
                abrir()
                Dim cmd As New SqlDataAdapter("SELECT id_teje,DNI AS DNI,trabajador AS TRABAJADOR FROM tejedores", conx)

                Dim da As New DataTable

                cmd.Fill(da)
                DataGridView1.DataSource = da
                DataGridView1.Columns(0).Visible = False
                DataGridView1.Columns(2).Width = 350
            End If

        End If

        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub TEJEDORES_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT id_teje,DNI AS DNI,trabajador AS TRABAJADOR FROM tejedores", conx)

        Dim da As New DataTable

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(2).Width = 350
        Label2.Text = DataGridView1.Rows(0).Cells(0).Value
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim i As Integer
        If Label2.Text = "" Then
            MsgBox("NO HA SELECCIONADO NINITEMS A ELIMINAR")
        Else
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then

                DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
                abrir()
                Dim agregar As String = "delete from tejedores where  id_teje ='" + Trim(Label2.Text) + "'"
                If (ELIMINAR(agregar)) Then
                    MessageBox.Show("DATO ELIMINADO CORRECTAMENTE")
                    Dim cmd As New SqlDataAdapter("SELECT id_teje,DNI AS DNI,trabajador AS TRABAJADOR FROM tejedores", conx)
                    Dim da As New DataTable
                    cmd.Fill(da)
                    DataGridView1.DataSource = da
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(2).Width = 350
                    Label2.Text = 0
                End If
            End If
        End If

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim j As Integer
        j = DataGridView1.Rows.Count



        If e.RowIndex = -1 Then
            DataGridView1.Rows(0).Selected = True
            Label2.Text = DataGridView1.Rows(0).Cells(0).Value
        Else
            If j > 0 Then
                Label2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class