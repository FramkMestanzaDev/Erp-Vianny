Imports System.Data.SqlClient
Public Class Validar_partida
    Public conx As SqlConnection
    Public comando As SqlCommand

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim gh As New vvalidarpartida
        Dim kk As New fcliente
        gh.gpartida = TextBox1.Text
        gh.garticulo = TextBox8.Text
        gh.gcantrollos = TextBox2.Text
        gh.gkgstejidos = TextBox3.Text
        gh.gmerma = TextBox4.Text
        gh.grollospesados = TextBox5.Text
        gh.gkgsobtenidos = TextBox6.Text
        gh.gobservacion = TextBox7.Text
        gh.gempresa = Label10.Text
        abrir()
        Dim agregar As String = "DELETE FROM validar_partida WHERE partida ='" + Trim(TextBox1.Text) + "' and  empresa='" + Label10.Text + "'"
        ELIMINAR(agregar)

        If kk.insertar_validar_partida(gh) Then
            MessageBox.Show("DATOS INGRESADOS CORRECTAMENTE")
            Dim cmd15 As New SqlCommand("update custom_vianny.dbo.marg0001 set flag_3r ='4' where partida_3r =@partida and ccia_3r=ccia", conx)
            cmd15.Parameters.AddWithValue("@partida", TextBox1.Text)
            cmd15.Parameters.AddWithValue("@ccia", Label10.Text)
            cmd15.ExecuteNonQuery()
            limpiar()
        Else
            MsgBox("NO SE GUARDO LA INFORMACION SOLICTADO")

        End If
        Dim JK As New conexion
        JK.desconectar()


    End Sub
    Public Sub NumConFrac1(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
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
    Public Sub NumConFrac2(ByVal CajaTexto As Windows.Forms.TextBox, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False

        Else
            e.Handled = True
        End If
    End Sub
    Sub limpiar()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = "0.00"
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox4.Text = "0.00"
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


    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then

            abrir()
            Dim Rs, rs1, rs2, Rs12 As SqlDataReader

            If CheckBox1.Checked = True Then
                Dim sql12 As String = " SELECT A.kgreq_3p AS KILOS,0 AS ROLLOS,B.nomb_17,'0.00' AS MERMA
		       		  FROM custom_vianny.dbo.qanp300 A  LEFT JOIN custom_vianny.dbo.Cag1700 B
		       					ON '01' = B.CCia AND A.Linea_3p = B.Linea_17
		       		   Where CCia_3p = '" + Label10.Text + "' AND Regis_3p = '" + TextBox1.Text + "' 
		       		   Order By A.regis_3p"
                Dim cmd12 As New SqlCommand(sql12, conx)
                Rs12 = cmd12.ExecuteReader()

                Rs12.Read()
                TextBox3.Text = Rs12(0) 'aca me pone el primer campo del select
                TextBox2.Text = Rs12(1)
                TextBox8.Text = Rs12(2)
                TextBox4.Text = Rs12(3)
                Rs12.Close()
                EJECUTAR()
            Else
                Try
                    Dim sql As String = "select SUM(cantkmv_3r),count(rollo_3r),C.nomb_17,SUM(kgsafec_3a) from custom_vianny.dbo.marg0001 M INNER JOIN custom_vianny.dbo.cag1700 C ON M.ccia_3r =C.ccia AND  M.linea_3r = C.linea_17 where partida_3r ='" + TextBox1.Text + "' AND flag_3r>0 and M.ccia_3r ='" + Label10.Text + "'  GROUP BY  C.nomb_17"
                    Dim cmd As New SqlCommand(sql, conx)
                    Rs = cmd.ExecuteReader()

                    Rs.Read()
                    TextBox3.Text = Rs(0) 'aca me pone el primer campo del select
                    TextBox2.Text = Rs(1)
                    TextBox8.Text = Rs(2)
                    TextBox4.Text = Rs(3)
                    Rs.Close()
                    EJECUTAR()
                Catch ex As Exception
                    MsgBox("LA PARTIDA NO ESTA REGISTRADA COMO PRODUCCION DE TELA")
                End Try

            End If






        End If

    End Sub
    Sub EJECUTAR()
        Dim respuesta As DialogResult
        Dim rs1, rs2 As SqlDataReader
        Dim sql1 As String = "select COUNT(flag_3r) from custom_vianny.dbo.marg0001 where partida_3r ='" + TextBox1.Text + "' and flag_3r in (4,3) and ccia_3r='" + Label10.Text + "'"
        Dim cmd1 As New SqlCommand(sql1, conx)
        rs1 = cmd1.ExecuteReader()
        rs1.Read()
        If rs1(0) = Trim(TextBox2.Text) Then
            TextBox5.ReadOnly = False
            TextBox6.ReadOnly = False
            TextBox7.ReadOnly = False
        Else
            MsgBox("LA CANTIDAD DE ROLLOS CREADOS NO COINCIDE CON LOS ROLLOS AUDITADOS, FALTA AUDITAR ROLLOS")

        End If
        rs1.Close()
        Dim sql2 As String = "SELECT rollospesados,kgsobtenidos,observacion,estado FROM validar_partida where partida='" + TextBox1.Text + "' and empresa ='" + Label10.Text + "' "
        Dim cmd2 As New SqlCommand(sql2, conx)
        rs2 = cmd2.ExecuteReader()

        If rs2.Read() = True Then

            If rs2(3) = 1 Then
                respuesta = MessageBox.Show("LA PARTIDA YA ESTA VALIDADA? QUIERES EDITARLA", "SALIR DEL PORGRAMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If (respuesta = Windows.Forms.DialogResult.Yes) Then
                    TextBox5.Text = rs2(0)
                    TextBox6.Text = rs2(1)
                    TextBox7.Text = rs2(2)
                    TextBox5.Enabled = True
                    TextBox6.Enabled = True
                    TextBox7.Enabled = True
                    Button1.Enabled = True
                Else
                    TextBox5.Text = rs2(0)
                    TextBox6.Text = rs2(1)
                    TextBox7.Text = rs2(2)
                    TextBox5.Enabled = False
                    TextBox6.Enabled = False
                    TextBox7.Enabled = False
                End If
            Else
                If rs2(3) = 2 Then
                    MsgBox("LA PARTIDA YA FUE AGREGADA EL PROGRAMA DE RAMA,NO SE PUEDE EDITAR")
                    TextBox5.Text = rs2(0)
                    TextBox6.Text = rs2(1)
                    TextBox7.Text = rs2(2)
                    TextBox5.Enabled = False
                    TextBox6.Enabled = False
                    TextBox7.Enabled = False
                    Button1.Enabled = False
                End If



            End If
        End If



        rs2.Close()
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        limpiar()
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        NumConFrac2(TextBox5, e)
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        NumConFrac1(TextBox6, e)
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
    Private Sub Validar_partida_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress

    End Sub

    Private Sub TextBox1_MouseCaptureChanged(sender As Object, e As EventArgs) Handles TextBox1.MouseCaptureChanged

    End Sub
End Class