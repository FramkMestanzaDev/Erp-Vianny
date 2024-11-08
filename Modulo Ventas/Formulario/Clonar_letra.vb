Imports System.Data.SqlClient
Public Class Clonar_letra
    Public conx As SqlConnection
    Public comando As SqlCommand
    Public conn As SqlDataAdapter
    Dim dt10 As New DataTable
    Sub abrir()
        Try
            conx = New SqlConnection("Data Source=172.21.0.1 ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")
            'cnn = New SqlConnection("Data Source=DESKTOP-RPGDGBV ;Initial Catalog=Comercial_Textil;User ID=sa;Password=Vi@Gr@Tex2005%")

            conx.Open()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub


    Private Sub Clonar_letra_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        dt10.Clear()
        DataGridView1.DataSource = ""


        Dim cmd5 As New SqlDataAdapter("SELECT NumLet as LETRA,DoRLet AS DOCUMENTO,FegLet AS FECHA,LugLet,MonLet,ImpLet,ImLLet,NomAce AS ACEPTANTE,DomAce,LocAce,RucAce,TelAce, NomAva,DomAav,LocAva,RucAva,TelAva FROM Letra l
        left join Aval  a on l.NumLet = a.idAva left join Aceptante  ac on l.NumLet = ac.idAce ORDER BY NumLet DESC", conx)
        Dim ds As New DataSet()
        cmd5.Fill(dt10)

        bs.DataSource = dt10
        If dt10.Rows.Count > 0 Then

            DataGridView1.DataSource = bs
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(1).Width = 180
            DataGridView1.Columns(7).Width = 250
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            DataGridView1.Columns(14).Visible = False
            DataGridView1.Columns(15).Visible = False
            DataGridView1.Columns(16).Visible = False

        End If


    End Sub


    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick

        If Label3.Text = 1 Then
            Form10.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Else
            If Label3.Text = 2 Then
                Form10.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            Else
                If Label3.Text = 3 Then
                    Letra.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
                    Letra.TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value
                    If Trim(DataGridView1.Rows(e.RowIndex).Cells(4).Value) = "S/" Then
                        Letra.ComboBox1.SelectedIndex = 0
                    Else
                        Letra.ComboBox1.SelectedIndex = 1
                    End If


                    Letra.TextBox8.Text = DataGridView1.Rows(e.RowIndex).Cells(7).Value
                    Letra.TextBox9.Text = DataGridView1.Rows(e.RowIndex).Cells(8).Value
                    Letra.TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(9).Value
                    Letra.TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(10).Value
                    Letra.TextBox12.Text = DataGridView1.Rows(e.RowIndex).Cells(11).Value

                    Letra.TextBox17.Text = DataGridView1.Rows(e.RowIndex).Cells(12).Value
                    Letra.TextBox16.Text = DataGridView1.Rows(e.RowIndex).Cells(13).Value
                    Letra.TextBox15.Text = DataGridView1.Rows(e.RowIndex).Cells(14).Value
                    Letra.TextBox14.Text = DataGridView1.Rows(e.RowIndex).Cells(15).Value
                    Letra.TextBox13.Text = DataGridView1.Rows(e.RowIndex).Cells(16).Value
                Else
                    If Label3.Text = 4 Then
                        Rpt_letra.TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                    Else
                        If Label3.Text = 5 Then
                            Rpt_letra.TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                        End If
                    End If
                End If
            End If
        End If

        Me.Close()
    End Sub

    Dim bs As New BindingSource()
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text <> String.Empty Then
            Dim filtro As String = String.Format("convert(LETRA, 'System.String') Like '%{0}%' ", TextBox2.Text)
            bs.Filter = filtro
        Else
            ' Si el TextBox está vacío, muestra todos los datos
            bs.Filter = ""
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text <> String.Empty Then
            Dim filtro As String = String.Format("convert(ACEPTANTE, 'System.String') Like '%{0}%' ", TextBox1.Text)
            bs.Filter = filtro
        Else
            ' Si el TextBox está vacío, muestra todos los datos
            bs.Filter = ""
        End If
    End Sub
End Class