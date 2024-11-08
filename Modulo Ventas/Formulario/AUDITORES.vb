Imports System.Data.SqlClient
Public Class AUDITORES
    Public conx As SqlConnection
    Public comando As SqlCommand


    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.F1 Then
            TRABAJADORES.Label2.Text = 2
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
        If TextBox1.Text = "" Or TextBox3.Text = "" Then
            MsgBox("FATA INGRESAR EL TRABAJADOR O LA CLAVE")
        Else
            abrir()
            Dim cmd1 As New SqlDataAdapter("SELECT * FROM auditores WHERE dni ='" + TextBox1.Text + "'", conx)
            Dim da1 As New DataTable
            cmd1.Fill(da1)

            If da1.Rows.Count > 0 Then
                MsgBox("EL TRABAJADOR YA ESTA REGISTRADO")
            Else
                Dim agregar As String = "INSERT INTO auditores (dni,trabajador,clave) VALUES('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "')"
                If (insertar(agregar)) Then
                    MessageBox.Show("DATOS INGRESADOS CORRECTAMENTE")
                    abrir()
                    Dim cmd As New SqlDataAdapter("SELECT id_audit,DNI AS DNI,trabajador AS TRABAJADOR,clave  as CLAVE FROM auditores", conx)
                    Dim da As New DataTable
                    cmd.Fill(da)
                    DataGridView1.DataSource = da
                    DataGridView1.Columns(0).Visible = False
                    DataGridView1.Columns(2).Width = 270
                End If

            End If

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim respuesta As DialogResult
        Dim i As Integer
        If Label2.Text = "" Then
            MsgBox("NO HA SELECCIONADO NINGUN ITEMS A ELIMINAR")
        Else
            respuesta = MessageBox.Show("DESEA ELIMINAR LA FILA ?", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If (respuesta = Windows.Forms.DialogResult.Yes) Then


                abrir()
                Dim agregar As String = "delete from auditores where  id_audit ='" + Trim(Label2.Text) + "'"
                If (ELIMINAR(agregar)) Then
                    MessageBox.Show("DATO ELIMINADO CORRECTAMENTE")
                    Dim cmd As New SqlDataAdapter("SELECT id_audit,DNI AS DNI,trabajador,clave  as CLAVE AS TRABAJADOR FROM auditores", conx)
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

    Private Sub AUDITORES_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        Dim cmd As New SqlDataAdapter("SELECT id_audit,DNI AS DNI,trabajador AS TRABAJADOR,clave  as CLAVE FROM auditores", conx)

        Dim da As New DataTable

        cmd.Fill(da)
        DataGridView1.DataSource = da
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(2).Width = 270
        If da.Rows.Count > 0 Then
            Label2.Text = DataGridView1.Rows(0).Cells(0).Value
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class