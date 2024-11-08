Public Class ADICIONAR_RUBRO_A_CODIGO
    Dim gk As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim TG As New FSTOCK_PAR
        Dim HJ As New VSTOCK_PAR
        Dim k As Integer
        Dim jk As Double
        Try


            jk = 0
            HJ.gCCIA = Label1.Text
            Select Case ComboBox1.Text
                Case "37 -- PRIMERA PTA2" : HJ.gALMACEN = "37"
                Case "38 -- SEGUNDA PT2" : HJ.gALMACEN = "38"
                Case "61 -- MUESTRAS" : HJ.gALMACEN = "61"
                Case "10 -- PRIMERA CHILCA" : HJ.gALMACEN = "10"
                Case "51 -- SEGUNDA CHILCA" : HJ.gALMACEN = "51"
                Case "54 -- TERCERA CHILCA" : HJ.gALMACEN = "54"
                Case "03 -- HILADO OPERATIVO" : HJ.gALMACEN = "03"
                Case "42 -- HILADO TEÑIDO CHILCA" : HJ.gALMACEN = "42"
                Case "59 -- SERVICIOS TERCEROS" : HJ.gALMACEN = "59"
                Case "67 -- ALMACEN HILO - TELA G" : HJ.gALMACEN = "67"
                Case "68 -- ALMACEN HUACHIPA" : HJ.gALMACEN = "68"
            End Select


            HJ.gMODAL = 1
            HJ.gano = Label2.Text
            gk = TG.CaeSoft_ReporteStockFisico(HJ)
            DataGridView1.DataSource = gk



            DataGridView1.Columns(3).Width = 165
            DataGridView1.Columns(4).Width = 430
            DataGridView1.Columns(8).Width = 65


            DataGridView1.Columns(5).Visible = False
            DataGridView1.Columns(6).Visible = False
            DataGridView1.Columns(7).Visible = False
            DataGridView1.Columns(9).Visible = False
            DataGridView1.Columns(10).Visible = False
            DataGridView1.Columns(11).Visible = False
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(13).Visible = False
            Dim u As Integer

            u = DataGridView1.Rows.Count

            For i = 0 To u - 1
                Dim ji As New FSTOCK_PAR
                Dim jj As New VSTOCK_PAR

                jj.gcodigo = DataGridView1.Rows(i).Cells(3).Value
                jj.gCCIA = Label1.Text
                DataGridView1.Rows(i).Cells(1).Value = ji.buscar_rubro(jj)
                Select Case DataGridView1.Rows(i).Cells(1).Value
                    Case "0114" : DataGridView1.Rows(i).Cells(2).Value = "PARIS APT"
                    Case "0108" : DataGridView1.Rows(i).Cells(2).Value = "TELA ACUARIO Y PARIS INDIGO"
                    Case "0104" : DataGridView1.Rows(i).Cells(2).Value = "TELA NOTEX SMS"
                    Case "0107" : DataGridView1.Rows(i).Cells(2).Value = "TELA NOTEX S"
                    Case "0103" : DataGridView1.Rows(i).Cells(2).Value = "INDUMENTARIA MEDICA SMS"
                    Case "0106" : DataGridView1.Rows(i).Cells(2).Value = "INDUMENTARIA MEDICA S"
                    Case "0001" : DataGridView1.Rows(i).Cells(2).Value = "VENTA MANUFACTURA"
                    Case "0099" : DataGridView1.Rows(i).Cells(2).Value = "HILADO TEÑIDO"
                    Case "0023" : DataGridView1.Rows(i).Cells(2).Value = "SERVICIO DE RAMA"
                    Case "0110" : DataGridView1.Rows(i).Cells(2).Value = "SERVICIO DE PERCHADO"
                    Case "0005" : DataGridView1.Rows(i).Cells(2).Value = "ACUARIO APT"
                    Case "0112" : DataGridView1.Rows(i).Cells(2).Value = "JERSEY-PIQUE-FRENCH TERRY-RIB 100% INDIGO "
                    Case "0113" : DataGridView1.Rows(i).Cells(2).Value = "RIB CRUDO"
                    Case "0116" : DataGridView1.Rows(i).Cells(2).Value = "HILADO CRUDO 24/1"
                    Case "0117" : DataGridView1.Rows(i).Cells(2).Value = "JULIOS APT"
                    Case "0118" : DataGridView1.Rows(i).Cells(2).Value = "FRENCH TERRY 60%  INDIGO"
                    Case "0038" : DataGridView1.Rows(i).Cells(2).Value = "SPANDEX GRAUS"
                    Case "0119" : DataGridView1.Rows(i).Cells(2).Value = "TELA NOTEX S"
                End Select
            Next

        Catch ex As Exception
            MsgBox("NO EXISTE STOCK PARA ESTE PERIODO")
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then

            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)'
            Dim num1, num2 As Integer
            num1 = e.RowIndex.ToString
            num2 = e.ColumnIndex.ToString

            Rubro.Label1.Text = num1
            Rubro.Label2.Text = num2
            Rubro.Show()


        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ko As New VSTOCK_PAR
        Dim kl As New FSTOCK_PAR
        Dim k As Integer
        k = DataGridView1.Rows.Count
        For i = 0 To k - 1
            If Me.DataGridView1.Rows(i).Cells(0).Value = True Then
                ko.gcodigo = DataGridView1.Rows(i).Cells(3).Value
                ko.gCCIA = Label1.Text
                ko.grubro = DataGridView1.Rows(i).Cells(1).Value
                kl.actualiar_rubro(ko)
            End If

        Next
        MsgBox("SE ACTUALIZARON LOS RUBROS A LOS CODIGOS SOLICITADOS")
        DataGridView1.DataSource = ""
    End Sub

    Private Sub ADICIONAR_RUBRO_A_CODIGO_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class