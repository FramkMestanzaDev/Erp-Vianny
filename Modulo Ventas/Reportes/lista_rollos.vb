Imports System.Data.SqlClient
Public Class lista_rollos
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TEJIDO.TextBox1.Text = Label3.Text
        TEJIDO.TextBox2.Text = ComboBox1.Text
        TEJIDO.TextBox3.Text = ComboBox2.Text
        TEJIDO.TextBox4.Text = Label4.Text
        TEJIDO.Show()
    End Sub
    Sub llenar_combo_box3()
        Try

            conn = New SqlDataAdapter("select rollo_3r from custom_vianny.dbo.marg0001 where partida_3r ='" + Label3.Text + "'", conx)
            conn.Fill(TY3)
            ComboBox1.DataSource = TY3
            ComboBox1.DisplayMember = "rollo_3r"
            ComboBox1.ValueMember = "rollo_3r"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub llenar_combo_box4()
        Try

            conn = New SqlDataAdapter("select rollo_3r from custom_vianny.dbo.marg0001 where partida_3r ='" + Label3.Text + "'", conx)
            conn.Fill(ty2)
            ComboBox2.DataSource = ty2
            ComboBox2.DisplayMember = "rollo_3r"
            ComboBox2.ValueMember = "rollo_3r"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        abrir()
        ty2.Clear()
        llenar_combo_box4()
    End Sub

    Private Sub lista_rollos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box3()
    End Sub
End Class