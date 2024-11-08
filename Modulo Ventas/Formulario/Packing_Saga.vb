Imports System.Data.SqlClient
Public Class Packing_Saga
    Public conx As SqlConnection
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
        Dim i1 As Integer
        DataGridView1.DataSource = Nothing
        importarExcel2(DataGridView1)
        Dim cajas, cantidad As Double
        cantidad = 0
        cajas = 0
        i1 = DataGridView1.Rows.Count

        For i = 0 To i1 - 1
            cantidad = cantidad + Convert.ToInt32(DataGridView1.Rows(i).Cells(6).Value)
        Next
        TextBox1.Text = Convert.ToString(cantidad)
        TextBox2.Text = DataGridView1.Rows.Count.ToString()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim can As Integer
        Dim CAJ As Integer
        can = DataGridView1.Rows.Count
        CAJ = 0
        For i = 0 To can - 1
            CAJ = CAJ + 1
            DataGridView1.Rows(i).Cells(8).Value = CAJ
        Next

    End Sub
    Sub Registrar_informacion()
        dim i as Integer
        for i=0 to DataGridView1.Rows.Count -1
            Dim cmd15 As New SqlCommand("INSERT INTO Packing (OrdPac,CitPac,FacPac,EanPac,SkuPac,DesPac,CanPac ,TiePac,SucPac ,LpnPac ,CajPac ,PoPac ,MarPac ,LinPac ,DptPac,FecPac,LotPac,CliPac) 
         values (@OrdPac,@CitPac,@FacPac,@EanPac,@SkuPac,@DesPac,@CanPac ,@TiePac,@SucPac ,@LpnPac ,@CajPac ,@PoPac ,@MarPac ,@LinPac ,@DptPac,@FecPac,@LotPac,@CliPac)", conx)
            cmd15.Parameters.AddWithValue("@OrdPac", DataGridView1.Rows(i).Cells(0).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@CitPac", DataGridView1.Rows(i).Cells(1).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@FacPac", DataGridView1.Rows(i).Cells(2).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@EanPac", DataGridView1.Rows(i).Cells(3).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@SkuPac", DataGridView1.Rows(i).Cells(4).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@DesPac", DataGridView1.Rows(i).Cells(5).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@CanPac", Convert.ToInt32(DataGridView1.Rows(i).Cells(6).Value.ToString().Trim()))
            cmd15.Parameters.AddWithValue("@TiePac", DataGridView1.Rows(i).Cells(7).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@SucPac", DataGridView1.Rows(i).Cells(8).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@LpnPac", DataGridView1.Rows(i).Cells(9).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@CajPac", DataGridView1.Rows(i).Cells(10).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@PoPac", DataGridView1.Rows(i).Cells(11).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@MarPac", DataGridView1.Rows(i).Cells(12).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@LinPac", DataGridView1.Rows(i).Cells(13).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@DptPac", DataGridView1.Rows(i).Cells(14).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@FecPac", Convert.ToDateTime(DataGridView1.Rows(i).Cells(15).Value.ToString().Trim()))
            cmd15.Parameters.AddWithValue("@LotPac", DataGridView1.Rows(i).Cells(16).Value.ToString().Trim())
            cmd15.Parameters.AddWithValue("@CliPac", DataGridView1.Rows(i).Cells(17).Value.ToString().Trim())

            cmd15.ExecuteNonQuery()


        Next

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Try
            abrir
            Registrar_informacion()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting


    End Sub



    Private Sub Packing_Saga_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class