Imports System.Data.SqlClient
Public Class ROLLOS_ADICIONALES
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
        GroupBox2.Visible = True
        Button1.Enabled = False
        DataGridView1.Enabled = False
        DataGridView2.Enabled = False
        abrir()
        Dim Rsr21, RTW As SqlDataReader
        Dim sql1 As String = "SELECT tela_4a,C.nomb_17,ancho_4a,galga_4a,diamet_4a,color_4,E.cnomb_3c FROM custom_vianny.dbo.matp040f P 
	 INNER JOIN  custom_vianny.dbo.matg040f G ON P.ncom_4a =G.ncom_4 AND P.ccia_4a =G.ccia_4 
	 LEFT JOIN custom_vianny.dbo.cag1700 C ON P.tela_4a = C.linea_17 AND P.ccia_4a = C.ccia 
	 LEFT  JOIN custom_vianny.dbo.Qarc0300 E ON G.color_4 = E.ccolor_3c  AND G.ccia_4 =E.ccia_3c
	 WHERE ncom_4a ='" + TextBox1.Text + TextBox2.Text + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        Rsr21 = cmd1.ExecuteReader()
        Rsr21.Read()
        Label12.Text = Rsr21(0)
        TextBox3.Text = Rsr21(1)
        TextBox4.Text = Rsr21(5)
        TextBox5.Text = Rsr21(6)
        TextBox6.Text = Rsr21(2)
        TextBox7.Text = Rsr21(3)
        TextBox8.Text = Rsr21(4)
        Rsr21.Close()

        Dim sql16 As String = " EXEC custom_vianny.dbo.spu_generanumerorollo '01', '" + Trim(Label11.Text) + "'"
        Dim cmd16 As New SqlCommand(sql16, conx)
        RTW = cmd16.ExecuteReader()
        RTW.Read()
        TextBox10.Text = RTW(0)
        RTW.Close()

    End Sub
    Dim ty As New DataTable
    Sub llenar_combo_box()
        Try
            conn = New SqlDataAdapter("select  sigla_mq  from custom_vianny.dbo.maquinas", conx)
            conn.Fill(ty)
            ComboBox2.DataSource = ty
            ComboBox2.DisplayMember = "sigla_mq"
            ComboBox2.ValueMember = "sigla_mq"

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub NumConFrac(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar = "." And Not CajaTexto.Text.IndexOf(".") Then
            e.Handled = True
        ElseIf e.KeyChar = "." Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        GroupBox2.Visible = False
        Button1.Enabled = True
        DataGridView1.Enabled = True
        DataGridView2.Enabled = True

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub
    Sub CARGAR()
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        abrir()
        Dim Rsr19 As SqlDataReader
        Dim sql101 As String = "exec custom_vianny.DBO.spu_consultarollosOrdTejido '01','" + TextBox1.Text + "','" + TextBox2.Text + "',NULL,NULL,NULL"
        Dim cmd101 As New SqlCommand(sql101, conx)
        Rsr19 = cmd101.ExecuteReader()
        Dim i5 As Integer
        i5 = 0
        While Rsr19.Read()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i5).Cells(1).Value = Rsr19(1)
            DataGridView1.Rows(i5).Cells(2).Value = Rsr19(30)
            DataGridView1.Rows(i5).Cells(3).Value = Rsr19(6)
            DataGridView1.Rows(i5).Cells(4).Value = Rsr19(8)
            DataGridView1.Rows(i5).Cells(5).Value = Rsr19(5)
            DataGridView1.Rows(i5).Cells(6).Value = Rsr19(17)
            DataGridView1.Rows(i5).Cells(7).Value = Rsr19(25)
            DataGridView1.Rows(i5).Cells(8).Value = Rsr19(26)
            DataGridView1.Rows(i5).Cells(9).Value = Rsr19(13)
            DataGridView1.Rows(i5).Cells(10).Value = Rsr19(12)
            DataGridView1.Rows(i5).Cells(11).Value = Rsr19(14)
            i5 = i5 + 1
        End While
        Rsr19.Close()
        Dim Rsr20 As SqlDataReader
        Dim sql1010 As String = "exec custom_vianny.DBO.spu_consultaresumenrollosOrdTejido '01','" + TextBox1.Text + "','" + TextBox2.Text + "',NULL"
        Dim cmd1010 As New SqlCommand(sql1010, conx)
        Rsr20 = cmd1010.ExecuteReader()
        Dim i51 As Integer
        i51 = 0
        While Rsr20.Read()
            DataGridView2.Rows.Add()
            DataGridView2.Rows(i51).Cells(0).Value = Rsr20(0)
            DataGridView2.Rows(i51).Cells(1).Value = Rsr20(1)
            DataGridView2.Rows(i51).Cells(2).Value = Rsr20(2)

            i51 = i51 + 1
        End While
        Rsr20.Close()
    End Sub
    Private Sub ROLLOS_ADICIONALES_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        abrir()
        llenar_combo_box()
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        CARGAR()
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        NumConFrac(TextBox9, e)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox9.Text = "" Then
            MsgBox("FALTA INGRESAR EL PESO")
        Else
            Dim r As String
            If ComboBox1.Text = "DIURNO" Then
                r = 1
            Else
                r = 2
            End If
            Dim cmd15 As New SqlCommand("IF NOT EXISTS( SELECT rollo_3r FROM custom_vianny.dbo.marg0001 Where ccia_3r  = '01'  AND rollo_3r= '" + Trim(TextBox10.Text) + "' )
               INSERT INTO custom_vianny.dbo.marg0001 ( ccia_3r , prvtej_3r                          , rollo_3r                  , rollom_3r , ncom_3r , pedido_3r                 , galga_3r                      , lote_3r                                              ,linea_3r                    ,maquina_3r                       , cantk_3r                        , prvhilo_3r                         , fcom_3r                 , fproduc_3r             , partida_3r                                           , turno_3r    , grm_3r , situa_3r , flag_3r , pedmov_3r , cantkmv_3r                     , ancho_3r                     , diam_3r                        , usokar_3r , ordens_3r , ccolor_3r                     , calidad_3r , ubica_3r ,ordtej_3r                                 , fecgen_3r , tejedor_3r , tiprol_3r , solmue_3r , prvtin_3r                         , almacen_3r              , periodo_3r              , voucher_3r              ,zarmado_3r ,letradiv_3r             , mtsafec_3a , kgsafec_3a , calrol_3r )
                                              values ( '01'   ,   '" + Trim(Label13.Text) + "'      , '" + TextBox10.Text + "'  , ''        ,''       , '" + TextBox1.Text + "'   ,'" + Trim(TextBox8.Text) + "'  ,'" + Trim(DataGridView1.Rows(0).Cells(4).Value) + "'  ,'" + Trim(Label12.Text) + "', '" + Trim(ComboBox2.Text) + "'  , '" + Trim(TextBox9.Text) + "'   , '" + Trim(Label13.Text) + "'       , '" + DBNull.Value + "'  , '" + DBNull.Value + "' ,'" + Trim(DataGridView1.Rows(0).Cells(6).Value) + "'  ,'" + r + "'  ,''     ,0         , 1        ,''        , '" + Trim(TextBox9.Text) + "'   ,'" + Trim(TextBox6.Text) + "' ,'" + Trim(TextBox7.Text) + "'   ,0          , ''       ,'" + Trim(TextBox4.Text) + "'   , ''         ,''        , '" + TextBox1.Text + TextBox2.Text + "'  ,getdate()  ,''          , '01'      , ''        ,'" + Trim(Label14.Text) + "'       , '" + DBNull.Value + "'  , '" + DBNull.Value + "'  , '" + DBNull.Value + "'  , 0         ,'" + DBNull.Value + "'  , 0.00     ,0.00      , ''  )

           ELSE
                Update custom_vianny.dbo.marg0001 SET
                prvtej_3r ='" + Trim(Label13.Text) + "',
                rollom_3r ='',          
                pedido_3r =  '" + TextBox1.Text + "' , 
                galga_3r = '" + Trim(TextBox8.Text) + "'  , 
                lote_3r='" + Trim(DataGridView1.Rows(0).Cells(4).Value) + "'  ,
                linea_3r ='" + Trim(Label12.Text) + "'  ,
                maquina_3r ='" + Trim(ComboBox2.Text) + "' , 
                cantk_3r ='" + Trim(TextBox9.Text) + "'  , 
                prvhilo_3r ='" + Trim(Label13.Text) + "'  ,
                fcom_3r ='" + DBNull.Value + "' ,
                fproduc_3r = '" + DBNull.Value + "' ,
                partida_3r='" + Trim(DataGridView1.Rows(0).Cells(6).Value) + "'   , 
                turno_3r ='" + r + "'  , 
                grm_3r ='' ,
                situa_3r=0  ,
                flag_3r=1,
                pedmov_3r= ''  , 
                cantkmv_3r='" + Trim(TextBox9.Text) + "'  ,
                ancho_3r='" + Trim(TextBox6.Text) + "'   ,
                diam_3r='" + Trim(TextBox7.Text) + "' ,
                usokar_3r = 0  ,
                ordens_3r =''  ,
                ccolor_3r ='" + Trim(TextBox4.Text) + "'  , 
                calidad_3r=''  ,
                ubica_3r=''  ,
                ordtej_3r='" + TextBox1.Text + TextBox2.Text + "' ,
                fecgen_3r = getdate(),
                tejedor_3r ='' ,
                tiprol_3r ='01',
                solmue_3r='' ,
                prvtin_3r='" + Trim(Label14.Text) + "'  , 
                almacen_3r ='" + DBNull.Value + "'  , 
                periodo_3r = '" + DBNull.Value + "'  ,
                voucher_3r ='" + DBNull.Value + "'
                Where ccia_3r = '01'  AND  rollo_3r = '" + Trim(TextBox10.Text) + "'", conx)


            cmd15.ExecuteNonQuery()
            MsgBox("SE AGREGO EL ROLLO SOLICITADO")

            TextBox9.Text = ""
            TextBox10.Text = ""
            GroupBox2.Visible = False
            Button1.Enabled = True
            CARGAR()
        End If

    End Sub



    Private Sub TextBox9_LostFocus(sender As Object, e As EventArgs) Handles TextBox9.LostFocus
        Dim NumAuxiliar As Double
        NumAuxiliar = TextBox9.Text

        TextBox9.Text = FormatNumber(NumAuxiliar, 2)
    End Sub
End Class