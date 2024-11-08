Imports System.Runtime.InteropServices
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Public Class Importar_Asiento
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
        ' Configura el diálogo para abrir un archivo Excel
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "Archivos Excel|*.xls;*.xlsx|Todos los archivos|*.*"
        openFileDialog1.Title = "Seleccionar archivo Excel"

        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Obtiene la ruta del archivo seleccionado
            Dim archivoExcel As String = openFileDialog1.FileName

            ' Inicializa una aplicación Excel
            Dim excelApp As New Application()
            Dim excelLibro As Workbook = excelApp.Workbooks.Open(archivoExcel)
            Dim excelHoja As Worksheet = CType(excelLibro.Sheets(1), Worksheet)

            ' Obtiene el rango de celdas con datos
            Dim rango As Range = excelHoja.UsedRange

            ' Obtiene el número de filas y columnas en el rango
            Dim numRows As Integer = rango.Rows.Count
            Dim numCols As Integer = rango.Columns.Count

            ' Limpia el DataGridView antes de cargar nuevos datos
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            ' Agrega columnas al DataGridView
            For j As Integer = 1 To numCols

                DataGridView1.Columns.Add(rango.Cells(1, j).Value, rango.Cells(1, j).Value)


            Next

            ' Agrega filas al DataGridView
            For i As Integer = 2 To numRows
                Dim fila As New List(Of Object)()
                For j As Integer = 1 To numCols
                    fila.Add(rango.Cells(i, j).Value)
                Next
                DataGridView1.Rows.Add(fila.ToArray())
            Next

            ' Cierra y libera los recursos de Excel
            excelLibro.Close()
            excelApp.Quit()
            Marshal.ReleaseComObject(excelLibro)
            Marshal.ReleaseComObject(excelApp)

            Dim p As Integer
            Dim debe, haber, diferencia As Double
            debe = 0
            haber = 0
            diferencia = 0
            p = DataGridView1.Rows.Count

            For i = 0 To p - 1
                If Trim(DataGridView1.Rows(i).Cells(5).Value) = "D" Then
                    debe = debe + CDbl(DataGridView1.Rows(i).Cells(6).Value)
                Else
                    If Trim(DataGridView1.Rows(i).Cells(5).Value) = "H" Then
                        haber = haber + CDbl(DataGridView1.Rows(i).Cells(6).Value)
                    End If
                End If
            Next
            diferencia = haber - debe
            If RadioButton1.Checked = True Then
                TextBox8.Text = Math.Round(debe, 2)
                TextBox9.Text = Math.Round(haber, 2)
                TextBox10.Text = Math.Round(diferencia, 2)

                TextBox11.Text = Math.Round(debe / CDbl(TextBox6.Text), 2)
                TextBox12.Text = Math.Round(haber / CDbl(TextBox6.Text), 2)
                TextBox13.Text = Math.Round(diferencia / CDbl(TextBox6.Text), 2)

            Else
                If RadioButton2.Checked = True Then
                    TextBox11.Text = Math.Round(debe, 2)
                    TextBox12.Text = Math.Round(haber, 2)
                    TextBox13.Text = Math.Round(diferencia, 2)

                    TextBox8.Text = Math.Round(debe * CDbl(TextBox6.Text), 2)
                    TextBox9.Text = Math.Round(haber * CDbl(TextBox6.Text), 2)
                    TextBox10.Text = Math.Round(diferencia * CDbl(TextBox6.Text), 2)
                End If
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
    End Sub
    Dim Rsr24, Rsr2278 As SqlDataReader
    Private Sub Importar_Asiento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        correlativo()
        RadioButton1.Checked = True
        RadioButton3.Checked = True
        tipo_cambio()
    End Sub
    Dim Rsr22, Rsr227 As SqlDataReader
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim fechaActual As DateTime
        Dim fechaComoString As String
        fechaActual = DateTime.Now
        fechaComoString = fechaActual.ToString("yyyy-MM-dd")
        Dim cmd20 As New SqlCommand("INSERT INTO custom_vianny.dbo.Cac3g00  (  CCIA_3 , CPERIODO_3 , CCOM_3 , NCOM_3     , TCOM_3 , MONE_3 ,FCOM_3     , TCAM_3 , GLOSA_3                                            , GLOSB_3 , TOTD1_3   , TOTD2_3  ,TOTH1_3   , TOTH2_3  , CUENB_3 , GIRADO_3 , NOMBCP_3 , FCOMA_3 ,NROA_3 , TMOV_3 , DIFCAM_3 , TASIEN_3 , FLAG_3 , FCOME_3    ,USUARIO_3  , FUPDATE_3 ) 
                                                                  VALUES    (  @CCIA_3,@CPERIODO_3 ,@CCOM_3 ,@NCOM_3     , 2      , @MONE_3,@FCOM_3    , @TCAM_3, @GLOSA_3                                           , ''      , @TOTD1_3  , @TOTD2_3  ,@TOTH1_3 , @TOTH2_3 , ''      , ''       , ''       , NULL    ,''     , ''     , 0        , 3        , 1      ,@FCOME_3    ,@USUARIO_3 , @FUPDATE_3 )", conx)

        cmd20.Parameters.AddWithValue("@CCIA_3", Label7.Text)
        cmd20.Parameters.AddWithValue("@CPERIODO_3", Label6.Text)
        cmd20.Parameters.AddWithValue("@CCOM_3", Trim(TextBox2.Text))
        cmd20.Parameters.AddWithValue("@NCOM_3", Trim(TextBox4.Text) & Trim(TextBox5.Text))
        If RadioButton1.Checked = True Then
            cmd20.Parameters.AddWithValue("@MONE_3", 1)
        Else
            If RadioButton2.Checked = True Then
                cmd20.Parameters.AddWithValue("@MONE_3", 2)
            End If
        End If
        cmd20.Parameters.AddWithValue("@FCOM_3", DateTimePicker1.Value.Date)
        cmd20.Parameters.AddWithValue("@TCAM_3", CDbl(TextBox6.Text))
        cmd20.Parameters.AddWithValue("@GLOSA_3", Trim(TextBox7.Text))
        cmd20.Parameters.AddWithValue("@FCOME_3", DateTimePicker1.Value)
        cmd20.Parameters.AddWithValue("@TOTD1_3", CDbl(TextBox8.Text))
        cmd20.Parameters.AddWithValue("@TOTD2_3", CDbl(TextBox11.Text))
        cmd20.Parameters.AddWithValue("@TOTH1_3", CDbl(TextBox9.Text))
        cmd20.Parameters.AddWithValue("@TOTH2_3", CDbl(TextBox12.Text))
        cmd20.Parameters.AddWithValue("@USUARIO_3", Mid(Trim(MDIParent1.Label3.Text), 1, 8))
        cmd20.Parameters.AddWithValue("@FUPDATE_3", "")
        cmd20.ExecuteNonQuery()

        Dim Ic As Integer
        Ic = DataGridView1.Rows.Count

        For i = 0 To Ic - 1
            Dim cmd21 As New SqlCommand("INSERT INTO custom_vianny.dbo.Cac3p00 ( CCIA_3A ,CPERIODO_3A , CCOM_3A , NCOM_3A   , CUEN_3A    ,CUEND_3A , CUENC_3A , CODH_3A , FICH_3A, CDOC_3A , NDOC_3A    , CREF_3A ,NREF_3A    , LINEA_3A , PPTOF_3A , NOMB_3A , IMP1_3A   , IMP2_3A  , FREF_3A    ,FLAGD_3A , PROY_3A , PRINCI_3A , FUEN_3A , PARTI_3A , TEFE_3A , SitLet_3a,Banco_3a,NroUnico_3a,Endoso_3a,Aval_3a,ocompra_3a,Ordens_3a,tcam_3a ,glosa_3a ,FEmision_3a,FLAG_3A )
                                                                         VALUES(@CCIA_3A ,@CPERIODO_3A,@CCOM_3A ,@NCOM_3A   ,@CUEN_3A    ,''       , @CUENC_3A,@CODH_3A ,@FICH_3A,@CDOC_3A ,@NDOC_3A    ,@CREF_3A ,@NREF_3A   , ''       , ''       , ''      ,@IMP1_3A   , @IMP2_3A , NULL       ,0        , ''      , ''        , ''      , ''       , ''      , ''       ,'  '    ,''         ,''       ,''     ,''        ,''       ,@tcam_3a ,''      ,NULL       ,1 )", conx)
            cmd21.Parameters.AddWithValue("@CCIA_3A", Trim(Label7.Text))
            cmd21.Parameters.AddWithValue("@CPERIODO_3A", Trim(Label6.Text))
            cmd21.Parameters.AddWithValue("@CCOM_3A", Trim(TextBox2.Text))
            cmd21.Parameters.AddWithValue("@NCOM_3A", Trim(TextBox4.Text) & Trim(TextBox5.Text))
            cmd21.Parameters.AddWithValue("@CUEN_3A", Trim(DataGridView1.Rows(i).Cells(4).Value))
            cmd21.Parameters.AddWithValue("@CODH_3A", Trim(DataGridView1.Rows(i).Cells(5).Value))
            cmd21.Parameters.AddWithValue("@CDOC_3A", Trim(TextBox2.Text))
            cmd21.Parameters.AddWithValue("@NDOC_3A", Trim(TextBox4.Text) & Trim(TextBox5.Text))
            cmd21.Parameters.AddWithValue("@CREF_3A", Trim(TextBox2.Text))
            cmd21.Parameters.AddWithValue("@NREF_3A", Trim(TextBox4.Text) & Trim(TextBox5.Text))
            If Trim(DataGridView1.Rows(i).Cells(10).Value) = "1" Then
                cmd21.Parameters.AddWithValue("@IMP1_3A", CDbl(DataGridView1.Rows(i).Cells(6).Value))
                cmd21.Parameters.AddWithValue("@IMP2_3A", CDbl(DataGridView1.Rows(i).Cells(6).Value) / CDbl(TextBox6.Text))
            Else
                cmd21.Parameters.AddWithValue("@IMP1_3A", CDbl(DataGridView1.Rows(i).Cells(6).Value) * CDbl(TextBox6.Text))
                cmd21.Parameters.AddWithValue("@IMP2_3A", CDbl(DataGridView1.Rows(i).Cells(6).Value))
            End If


            Dim VALOR, VALOR1 As String
            VALOR = 0
            VALOR1 = 0
            cmd21.Parameters.AddWithValue("@tcam_3a", CDbl(TextBox6.Text))
            Dim sql102157 As String = "SELECT nivd_2 FROM custom_vianny.dbo.CAG0200 Where ccia_2='" + Trim(Label7.Text) + "' AND cuen_2 ='" + Trim(DataGridView1.Rows(i).Cells(4).Value) + "'"
            Dim cmd102157 As New SqlCommand(sql102157, conx)
            Rsr227 = cmd102157.ExecuteReader()

            If Rsr227.Read() Then
                VALOR = Rsr227(0)
            End If
            Rsr227.Close()

            If VALOR = "1" Then

                cmd21.Parameters.AddWithValue("@FICH_3A", Trim(DataGridView1.Rows(i).Cells(8).Value))
            Else
                cmd21.Parameters.AddWithValue("@FICH_3A", "")
            End If

            Dim sql1021578 As String = "SELECT ccosto_2 FROM custom_vianny.dbo.CAG0200 Where ccia_2='" + Trim(Label7.Text) + "' AND cuen_2 ='" + Trim(DataGridView1.Rows(i).Cells(4).Value) + "'"
            Dim cmd1021578 As New SqlCommand(sql1021578, conx)
            Rsr2278 = cmd1021578.ExecuteReader()
            If Rsr2278.Read() Then
                VALOR1 = Rsr2278(0)
            End If
            Rsr2278.Close()

            If VALOR1 = "1" Then

                cmd21.Parameters.AddWithValue("@CUENC_3A", Trim(DataGridView1.Rows(i).Cells(7).Value))
            Else
                cmd21.Parameters.AddWithValue("@CUENC_3A", "")
            End If

            cmd21.ExecuteNonQuery()
        Next


        MsgBox("SE REGISTRO LA INFORMACION INFORMACION CORRECTAMENTE")
        correlativo()
        LIMPIAR()
    End Sub
    Sub LIMPIAR()
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        RadioButton1.Checked = True
        RadioButton3.Checked = True
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        tipo_cambio()
    End Sub
    Sub tipo_cambio()
        Dim func As New ftcambio
        Dim dts As New vtcambio
        dts.gfecha = DateTimePicker1.Text
        TextBox6.Text = func.mostrar_tipo_cambio(dts)
    End Sub
    Sub correlativo()
        abrir()
        TextBox4.Text = DateTimePicker1.Value.Month
        Select Case TextBox4.Text.Length
            Case "1" : TextBox4.Text = "0" & "" & TextBox4.Text
            Case "2" : TextBox4.Text = TextBox4.Text

        End Select
        Dim sql1023 As String = " EXEC CaeSoft_GetAllMaximoMovimientoContable '" + Label7.Text + "','" + Label6.Text + "','" + Trim(TextBox2.Text) + "','" + Trim(TextBox4.Text) + "'"
            Dim cmd1023 As New SqlCommand(sql1023, conx)
        Rsr24 = cmd1023.ExecuteReader()
        If Rsr24.Read() = True Then
            TextBox5.Text = Rsr24(0)
        Else
            TextBox5.Text = 1

        End If
        Select Case TextBox5.Text.Length
            Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
            Case "1" : TextBox5.Text = "00000" & "" & TextBox5.Text
            Case "2" : TextBox5.Text = "0000" & "" & TextBox5.Text
            Case "3" : TextBox5.Text = "000" & "" & TextBox5.Text
            Case "4" : TextBox5.Text = "00" & "" & TextBox5.Text
            Case "5" : TextBox5.Text = "0" & "" & TextBox5.Text
            Case "6" : TextBox5.Text = TextBox5.Text
        End Select
        Rsr24.Close()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case TextBox4.Text.Length
                Case "1" : TextBox4.Text = "0" & TextBox4.Text
                Case "2" : TextBox4.Text = TextBox4.Text

            End Select
            TextBox5.Select()
            correlativo()
        End If
    End Sub
End Class