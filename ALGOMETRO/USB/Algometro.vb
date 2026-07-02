Option Strict Off
Option Explicit On
Imports System.IO
Imports System.IO.Ports
Imports System.Windows.Forms

Public Class Form1
    Inherits System.Windows.Forms.Form

    'Private SupervisorioProc As System.Diagnostics.Process

    Private Const VendorID As Short = 1121 '0461h
    Private Const ProductID As Short = 26

    Private Const BufferInSize As Short = 64
    Private Const BufferOutSize As Short = 64

    Dim fileReader As System.IO.StreamReader
    Dim stringReader As String

    Dim MAXIMO(4) As Decimal

    Dim FORCA(4) As ULong
    Dim PALAVRA_FORCA As String
    Dim AMOSTRAGEM As ULong = 0
    Dim FORCA_CALIBRADA(4) As Decimal

    Dim TRANSD_ZERO As Boolean = False
    Dim ZERO_GRADE(4) As Decimal

    Dim PONTO_ZERO(4) As Decimal
    Dim PONTO_ZERO_CALIBRADO(4) As ULong
    Dim PONTO_GANHO(4) As Decimal
    Dim PESO_CAL As ULong = 50

    Dim CALIBRACAO As Boolean = False

    Dim conta_wd As ULong = 0




    'Threading.Thread.Sleep(4000)

    ' byte - 0 a 255
    ' decimal - (+/-7.9228162514264337593543950335E+28)
    ' double - +/- 1,79769313486231570E+308 a 4,94065645841246544E-324
    ' integer - -2.147.483.648 a 2.147.483.647
    ' Uinteger - 0 a 4.294.967.295
    ' short -  -32.768 a 32.767
    ' long - -9.223.372.036.854.775.808 a 9.223.372.036.854.775.807
    ' Sbyte - -128 a 127
    ' Ulong - 0 a 18.446.744.073.709.551.615
    ' Ushort - 0 a 65.535

    Dim FATOR_FILTRO As Byte = 50 '3 'quantidade de pontos do filtro
    Dim BufferIn(BufferInSize) As Byte
    Dim BufferOut(BufferOutSize) As Byte
    Dim INUTIL As ULong
    Dim escreve_arquivo As Boolean
    Dim gravar As Boolean
    Dim nome_arquivo As String
    Dim data_hoje As String
    Dim texto_saida As String
    Dim data As String
    'Dim contador_tempo As ULong = 0
    Dim ch5, ch4, ch3, ch2, ch1, ch0 As Byte
    Dim conta_pontos As ULong = 0

    Dim ULTIMO As Byte = 0

    Dim VALOR_SEGUNDO As ULong = 0


    Dim palavra3 As String
    Dim contagem_bbb As ULong = 0
    Dim segundo_inicial As Int64 = 0

    Dim leu_modelo As Boolean = False
    Dim controle As ULong = 0
    Dim reconecta As ULong = 0
    Dim watchdog As ULong = 0
    Dim estado As String



    'conjunto de pontos que compoe a media
    'FILTRO_MEDIA_TV, FILTRO_MEDIA_O2, FILTRO_MEDIA_CO2, FILTRO_MEDIA_RR

    'leitura dos pontos para fazer a media
    'media_TV, media_O2, media_CO2,media_RR

    'indice do contador de media de 0 a quantidade de pontos
    'INDEX_MEDIA_FORCA

    Dim FILTRO_MEDIA_FORCA As Decimal
    Dim INDEX_MEDIA_FORCA As Long
    Dim MEDIA_FILTRO_FORCA(100) As Decimal

    Dim LIGADO As Boolean = False
    Dim leu_filtro As Boolean = False

    Dim terminou As Boolean = False
    Dim INICIOU As Boolean = False



    Private Sub ATUALIZA_VISUAL(dgv As DataGridView, valorText As String)
        dgv.Rows.Add()
        Dim idx As Integer = dgv.RowCount - 2
        dgv.Rows(idx).Cells(0).Value = idx + 1
        dgv.Rows(idx).Cells(1).Value = valorText
        dgv.FirstDisplayedScrollingRowIndex = dgv.RowCount - 1
        contagem_bbb += 1
    End Sub


    Private Sub datagrid_dados()
        DataGridView1.ColumnCount = 2
        DataGridView1.Columns(0).Name = "Indice"
        DataGridView1.Columns(0).Width = 50
        DataGridView1.Columns(1).Name = "Força Máxima"
        DataGridView1.Columns(1).Width = 100

        DataGridView2.ColumnCount = 2
        DataGridView2.Columns(0).Name = "Indice"
        DataGridView2.Columns(0).Width = 50
        DataGridView2.Columns(1).Name = "Força Máxima"
        DataGridView2.Columns(1).Width = 100

        DataGridView3.ColumnCount = 2
        DataGridView3.Columns(0).Name = "Indice"
        DataGridView3.Columns(0).Width = 50
        DataGridView3.Columns(1).Name = "Força Máxima"
        DataGridView3.Columns(1).Width = 100

        DataGridView4.ColumnCount = 2
        DataGridView4.Columns(0).Name = "Indice"
        DataGridView4.Columns(0).Width = 50
        DataGridView4.Columns(1).Name = "Força Máxima"
        DataGridView4.Columns(1).Width = 100
    End Sub
    Private Sub Form1_closed(sender As Object, e As EventArgs) Handles MyBase.FormClosed
        BufferOut(0) = 0
        BufferOut(1) = Asc("P")
        hidWriteEx(VendorID, ProductID, BufferOut(0))
    End Sub
    Public Sub INICIALIZAR()
        segundo_inicial = VALOR_SEGUNDO 'Int(Environment.TickCount / 1000)
        'System.Threading.Thread.Sleep(600) ' TEMPORARIO: removido para não travar UI
        data_hoje = Date.Now
        data = Format(Date.Now, "ddMMyyyy_HHmmss")
        nome_arquivo = "C:\ALGOMETRO\Resultados\" & TB_Arquivo.Text & data & ".csv"
        texto_saida = "Data Atual;" & "Contagem;" + "Força Maxima"

        With My.Computer.FileSystem
            .WriteAllText(nome_arquivo, "Data do Teste: ;" & data_hoje & vbCrLf, True)
            .WriteAllText(nome_arquivo, "Inicio do Teste: ;" & TB_Horario.Text & vbCrLf, True)
            .WriteAllText(nome_arquivo, "Identificação do Animal: ;" & TB_Nome.Text & "  Peso (Kg):  " & Tbx_massa.Text & vbCrLf, True)
            .WriteAllText(nome_arquivo, "Avaliador: ;" & TB_Avaliador.Text & vbCrLf, True)
            .WriteAllText(nome_arquivo, "Observações: ;" & TB_Obs.Text & vbCrLf, True)
            .WriteAllText(nome_arquivo, vbCrLf, True)
            .WriteAllText(nome_arquivo, texto_saida & vbCrLf, True)
        End With
        DataGridView1.RowCount = 1
        contagem_bbb = 0
    End Sub

    Protected Overrides ReadOnly Property CreateParams() As CreateParams 'rotina para desabilitar o "X" do formulário
        Get
            Dim param As CreateParams = MyBase.CreateParams
            param.ClassStyle = param.ClassStyle Or &H200
            Return param
        End Get
    End Property

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        'If MessageBox.Show("Deseja sair da aplicação?", "My Application", MessageBoxButtons.YesNo) = DialogResult.No Then
        e.Cancel = True
        'End If

    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BTN_INICIAR.Enabled = False
        ConnectToHID(Me.Handle.ToInt32)
        escreve_arquivo = False
        gravar = False
        datagrid_dados()
        Label26.Text = Date.Now
        executa_filtro()
        fileReader = My.Computer.FileSystem.OpenTextFileReader("C:\Algometro\Config.txt")
        TB_Arquivo.Text = fileReader.ReadLine()
        TB_Horario.Text = fileReader.ReadLine()
        TB_Nome.Text = fileReader.ReadLine()
        Tbx_massa.Text = fileReader.ReadLine()
        TB_Avaliador.Text = fileReader.ReadLine()
        TB_Obs.Text = fileReader.ReadLine()
        fileReader.Close()
        'LIGADO = True
        INICIOU = True
        executa_filtro()
        terminou = False
    End Sub

    Private Sub executa_filtro()
        Dim inutil_filtro As Long
        INDEX_MEDIA_FORCA = 0
        FILTRO_MEDIA_FORCA = 0
        For inutil_filtro = 0 To FATOR_FILTRO - 1 '9
            MEDIA_FILTRO_FORCA(inutil_filtro) = FORCA(1)
            FILTRO_MEDIA_FORCA = FILTRO_MEDIA_FORCA + MEDIA_FILTRO_FORCA(inutil_filtro)
        Next inutil_filtro
    End Sub
    Public Sub OnPlugged(ByVal pHandle As Integer)
        Dim vid, pid As Integer
        vid = hidGetVendorID(pHandle)
        pid = hidGetProductID(pHandle)

        If vid = VendorID And pid = ProductID Then
            Me.Text = "ALGOMETRO - Conectado"
            Label34.Text = "ALGOMETRO - Conectado"
            Label1.Text = vid
            Label2.Text = pid
            Timer4.Enabled = True
        End If

    End Sub

    Public Sub OnUnplugged(ByVal pHandle As Integer)
        If hidGetVendorID(pHandle) = VendorID And hidGetProductID(pHandle) = ProductID Then
            hidSetReadNotify(hidGetHandle(VendorID, ProductID), False)
            Me.Text = "ALGOMETRO - Desconectado"
            Label34.Text = "ALGOMETRO - Desconectado"
            Label1.Text = "N/C"
            Label2.Text = "N/C"
            TextBox3.Text = ""
            TextBox15.Text = ""
            TextBox19.Text = ""
            TextBox23.Text = ""

            TextBox5.Text = ""
            TextBox11.Text = ""
            TextBox17.Text = ""
            TextBox21.Text = ""

            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox16.Text = ""
            TextBox20.Text = ""


            TextBox12.Text = ""
            TextBox26.Text = ""
            TextBox29.Text = ""
            TextBox32.Text = ""

        End If
        leu_filtro = False
    End Sub

    Public Sub OnChanged()
        Dim pHandle As Integer
        pHandle = hidGetHandle(VendorID, ProductID)
        hidSetReadNotify(hidGetHandle(VendorID, ProductID), True)
    End Sub

    Public Sub OnRead(ByVal pHandle As Integer)
        Dim MOSTRA_FORCA(4) As Decimal

        If hidRead(pHandle, BufferIn(0)) Then
            palavra3 = ""
            For I = 1 To 64
                palavra3 = palavra3 + Chr(BufferIn(I))
            Next
            watchdog = 0
        End If
        If Mid(palavra3, 1, 2) = "03" Then
            TextBox2.Text = Mid(palavra3, 17, 7)
            TextBox8.Text = Mid(palavra3, 27, 8)
            TextBox9.Text = Mid(palavra3, 1, 2)
        Else
            ' TEMPORARIO: loop intercalado 4 canais (ch1@byte3, ch2@byte5, ch3@byte7, ch4@byte9, passo 8)
            For ch As Integer = 1 To 4
                Dim soma As Double = 0
                Dim count As Integer = 0
                Dim pos As Integer = 1 + ch * 2
                Do While pos + 1 <= 64
                    soma += (Asc(Mid(palavra3, pos, 1)) * 256 + Asc(Mid(palavra3, pos + 1, 1))) / 1.26
                    pos += 8
                    count += 1
                Loop
                If count > 0 Then FORCA(ch) = soma / count
            Next
        End If

        conta_pontos = conta_pontos + 1
        AMOSTRAGEM = AMOSTRAGEM + 1
        FORCA(1) = leitura_filtro(0, False)
        MOSTRA_FORCA(1) = FORCA(1) / 1000
        MOSTRA_FORCA(2) = FORCA(2) / 1000
        MOSTRA_FORCA(3) = FORCA(3) / 1000
        MOSTRA_FORCA(4) = FORCA(4) / 1000
        TextBox6.Text = Format(MOSTRA_FORCA(1), "###0.000")
        TextBox7.Text = Format(MOSTRA_FORCA(2), "###0.000")
        TextBox16.Text = Format(MOSTRA_FORCA(3), "###0.000")
        TextBox20.Text = Format(MOSTRA_FORCA(4), "###0.000")



        TextBox3.Text = Asc(Mid(palavra3, 1, 1))
        TextBox15.Text = Asc(Mid(palavra3, 1, 1))
        TextBox19.Text = Asc(Mid(palavra3, 1, 1))
        TextBox23.Text = Asc(Mid(palavra3, 1, 1))

        TextBox5.Text = Asc(Mid(palavra3, 2, 1))
        TextBox11.Text = Asc(Mid(palavra3, 2, 1))
        TextBox17.Text = Asc(Mid(palavra3, 2, 1))
        TextBox21.Text = Asc(Mid(palavra3, 2, 1))

        TextBox12.Text = Format(MOSTRA_FORCA(1), "###0.00")
        TextBox26.Text = Format(MOSTRA_FORCA(2), "###0.00")
        TextBox29.Text = Format(MOSTRA_FORCA(3), "###0.00")
        TextBox32.Text = Format(MOSTRA_FORCA(4), "###0.00")


        FORCA_CALIBRADA(1) = (1000 * (MOSTRA_FORCA(1) - PONTO_ZERO(1))) / 0.96
        FORCA_CALIBRADA(2) = (1000 * (MOSTRA_FORCA(2) - PONTO_ZERO(2))) / 0.96
        FORCA_CALIBRADA(3) = (1000 * (MOSTRA_FORCA(3) - PONTO_ZERO(3))) / 0.96
        FORCA_CALIBRADA(4) = (1000 * (MOSTRA_FORCA(4) - PONTO_ZERO(4))) / 0.96
        If FORCA_CALIBRADA(1) < 10 Then
            FORCA_CALIBRADA(1) = 0
        End If
        If FORCA_CALIBRADA(2) < 10 Then
            FORCA_CALIBRADA(2) = 0
        End If
        If FORCA_CALIBRADA(3) < 10 Then
            FORCA_CALIBRADA(3) = 0
        End If
        If FORCA_CALIBRADA(4) < 10 Then
            FORCA_CALIBRADA(4) = 0
        End If

        TextBox10.Text = Format(FORCA_CALIBRADA(1), "###0")
        TextBox24.Text = Format(FORCA_CALIBRADA(2), "###0")
        TextBox27.Text = Format(FORCA_CALIBRADA(3), "###0")
        TextBox30.Text = Format(FORCA_CALIBRADA(4), "###0")
        GaugeCanal1.GaugeValue = CSng(FORCA_CALIBRADA(1))
        GaugeCanal2.GaugeValue = CSng(FORCA_CALIBRADA(2))
        GaugeCanal3.GaugeValue = CSng(FORCA_CALIBRADA(3))
        GaugeCanal4.GaugeValue = CSng(FORCA_CALIBRADA(4))


        If CALIBRACAO = True Then
            If Mid(palavra3, 1, 1) = "K" Then
                TextBox4.BackColor = Color.Black
                Timer2.Enabled = True
                ATUALIZA_VISUAL(DataGridView1, Button2.Text)
                ATUALIZA_VISUAL(DataGridView2, Button10.Text)
                ATUALIZA_VISUAL(DataGridView3, Button11.Text)
                ATUALIZA_VISUAL(DataGridView4, Button12.Text)
                texto_saida = Date.Now & ";" & contagem_bbb & ";" &
                              Button2.Text & ";" & Button10.Text & ";" &
                              Button11.Text & ";" & Button12.Text
                With My.Computer.FileSystem
                    .WriteAllText(nome_arquivo, texto_saida & vbCrLf, True)
                End With
                MAXIMO(1) = 0 : MAXIMO(2) = 0 : MAXIMO(3) = 0 : MAXIMO(4) = 0
                Button2.Text = Format(PONTO_ZERO(1) * 1000, "###0") + " gf"
                Button10.Text = Format(PONTO_ZERO(2) * 1000, "###0") + " gf"
                Button11.Text = Format(PONTO_ZERO(3) * 1000, "###0") + " gf"
                Button12.Text = Format(PONTO_ZERO(4) * 1000, "###0") + " gf"
                'System.Threading.Thread.Sleep(300) ' TEMPORARIO: removido para não travar UI
            End If
        End If
        If Mid(palavra3, 1, 1) = "R" Then
            TextBox4.BackColor = Color.Red
            Timer2.Enabled = True
            CALIBRACAO = True
            PONTO_ZERO(1) = FORCA(1) / 1000
            TextBox13.Text = Format(PONTO_ZERO(1), "###0.00")
            PONTO_ZERO_CALIBRADO(1) = 0
            TabControl1.SelectedIndex = 0
            BTN_INICIAR.Enabled = CALIBRACAO
            'ATUALIZA_VISUAL()
            MAXIMO(1) = 0
            Button2.Text = Format(1000 * (MAXIMO(1) + PONTO_ZERO(1)), "###0") + " gf"
            Timer5.Enabled = False
            BTN_INICIAR.Text = "INICIAR"
            TextBox3.Text = ""
            TextBox15.Text = ""
            TextBox19.Text = ""
            TextBox23.Text = ""

            TextBox5.Text = ""
            TextBox11.Text = ""
            TextBox17.Text = ""
            TextBox21.Text = ""

            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox16.Text = ""
            TextBox20.Text = ""

            'LIGADO = False
        End If

        If FORCA_CALIBRADA(1) > MAXIMO(1) Then
            MAXIMO(1) = FORCA_CALIBRADA(1)
        End If

        If FORCA_CALIBRADA(2) > MAXIMO(2) Then
            MAXIMO(2) = FORCA_CALIBRADA(2)
        End If

        If FORCA_CALIBRADA(3) > MAXIMO(3) Then
            MAXIMO(3) = FORCA_CALIBRADA(3)
        End If

        If FORCA_CALIBRADA(4) > MAXIMO(4) Then
            MAXIMO(4) = FORCA_CALIBRADA(4)
        End If

        If MAXIMO(1) < 30 Then MAXIMO(1) = 0
        If MAXIMO(2) < 30 Then MAXIMO(2) = 0
        If MAXIMO(3) < 30 Then MAXIMO(3) = 0
        If MAXIMO(4) < 30 Then MAXIMO(4) = 0
        Button2.Text = Format(MAXIMO(1), "###0") + " gf"
        Button10.Text = Format(MAXIMO(2), "###0") + " gf"
        Button11.Text = Format(MAXIMO(3), "###0") + " gf"
        Button12.Text = Format(MAXIMO(4), "###0") + " gf"
        TextBox1.Text = Format(FORCA_CALIBRADA(1), "###0")
        TextBox14.Text = Format(FORCA_CALIBRADA(2), "###0")
        TextBox18.Text = Format(FORCA_CALIBRADA(3), "###0")
        TextBox22.Text = Format(FORCA_CALIBRADA(4), "###0")

        conta_pontos = 0
    End Sub

    Private Sub BTN_INICIAR_Click(sender As Object, e As EventArgs) Handles BTN_INICIAR.Click
        With My.Computer.FileSystem
            .DeleteFile("C:\Algometro\Config.txt")
            .WriteAllText("C:\Algometro\config.txt", TB_Arquivo.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", TB_Horario.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", TB_Nome.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", Tbx_massa.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", TB_Avaliador.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", TB_Obs.Text & vbCrLf, True)
        End With

        LIGADO = Not (LIGADO)
        INICIOU = True
        executa_filtro()
        terminou = False

        If LIGADO = True Then
            INICIALIZAR()
            'System.Threading.Thread.Sleep(600) ' TEMPORARIO: removido para não travar UI
            data_hoje = Date.Now
            data = Format(Date.Now, "ddMMyyyy_HHmmss")
            BTN_INICIAR.Text = "PARAR"
            DataGridView1.RowCount = 1
            DataGridView2.RowCount = 1
            DataGridView3.RowCount = 1
            DataGridView4.RowCount = 1
            contagem_bbb = 0
            If CALIBRACAO = True Then
                Button1.Visible = True
                Button2.Visible = True

                Button10.Visible = True
                Button14.Visible = True

                Button13.Visible = True
                Button11.Visible = True

                Button15.Visible = True
                Button12.Visible = True

                DataGridView1.Visible = True
                DataGridView2.Visible = True
                DataGridView3.Visible = True
                DataGridView4.Visible = True
                MAXIMO(1) = 0
                MAXIMO(2) = 0
                MAXIMO(3) = 0
                MAXIMO(4) = 0
            End If

        Else
            BTN_INICIAR.Text = "INICIAR"

            Button1.Visible = False
            Button2.Visible = False

            Button10.Visible = False
            Button14.Visible = False

            Button13.Visible = False
            Button11.Visible = False

            Button15.Visible = False
            Button12.Visible = False

            DataGridView1.Visible = False
            DataGridView2.Visible = False
            DataGridView3.Visible = False
            DataGridView4.Visible = False
        End If
fim_inicio:
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'rotina responsavel por enviar o Watchdog
        '        Label26.Text = (Date.Now)
        '        BufferOut(0) = 0
        '        BufferOut(1) = Asc("W")
        '        hidWriteEx(VendorID, ProductID, BufferOut(0))
        Label6.Text = AMOSTRAGEM
        'If (AMOSTRAGEM < 100) And (LIGADO = True) Then
        'BufferOut(0) = 0
        'BufferOut(1) = Asc("P")
        'hidWriteEx(VendorID, ProductID, BufferOut(0))
        'System.Threading.Thread.Sleep(600)
        'BufferOut(0) = 0
        'BufferOut(1) = Asc("Z")
        'hidWriteEx(VendorID, ProductID, BufferOut(0))
        'System.Threading.Thread.Sleep(600)
        'BufferOut(0) = 0
        'BufferOut(1) = Asc("A")
        'BufferOut(2) = 1
        'BufferOut(3) = 1
        'hidWriteEx(VendorID, ProductID, BufferOut(0))
        'End If
        AMOSTRAGEM = 0

    End Sub


    Function leitura_filtro(fonte As Byte, inicializacao As Boolean) As Integer
        Dim temporario As ULong
        Select Case fonte
            'conjunto de pontos que compoe a media
            'FILTRO_MEDIA_TV, FILTRO_MEDIA_O2, FILTRO_MEDIA_CO2, FILTRO_MEDIA_RR

            'leitura dos pontos para fazer a media
            'media_TV, media_O2, media_CO2,media_RR

            'indice do contador de media de 0 a quantidade de pontos
            'INDEX_MEDIA_TV, INDEX_MEDIA_O2, INDEX_MEDIA_CO2, INDEX_MEDIA_RR

            Case 0
                If inicializacao = False Then
                    '                    FILTRO_MEDIA_TV = FILTRO_MEDIA_TV - MEDIA_TV(INDEX_MEDIA_TV)
                    FILTRO_MEDIA_FORCA = FILTRO_MEDIA_FORCA - MEDIA_FILTRO_FORCA(INDEX_MEDIA_FORCA)
                End If
                MEDIA_FILTRO_FORCA(INDEX_MEDIA_FORCA) = FORCA(1)
                FILTRO_MEDIA_FORCA = FILTRO_MEDIA_FORCA + MEDIA_FILTRO_FORCA(INDEX_MEDIA_FORCA)
                INDEX_MEDIA_FORCA = INDEX_MEDIA_FORCA + 1
                If INDEX_MEDIA_FORCA = FATOR_FILTRO Then
                    INDEX_MEDIA_FORCA = 0
                End If
                temporario = FILTRO_MEDIA_FORCA / FATOR_FILTRO

        End Select
        leitura_filtro = temporario
        Exit Function
    End Function


    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        BufferOut(0) = 0
        BufferOut(1) = Asc("P")
        hidWriteEx(VendorID, ProductID, BufferOut(0))
        'Threading.Thread.Sleep(1000) ' TEMPORARIO: removido para não travar UI
        BufferOut(0) = 0
        BufferOut(1) = Asc("Z")
        hidWriteEx(VendorID, ProductID, BufferOut(0))
        'Threading.Thread.Sleep(1000) ' TEMPORARIO: removido para não travar UI
        BufferOut(0) = 0
        BufferOut(1) = Asc("A")
        BufferOut(2) = 1
        BufferOut(3) = 1
        hidWriteEx(VendorID, ProductID, BufferOut(0))
        Timer4.Enabled = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        BufferOut(0) = 0
        BufferOut(1) = Asc("Z")
        hidWriteEx(VendorID, ProductID, BufferOut(0))
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        BufferOut(0) = 0
        BufferOut(1) = Asc("P")
        hidWriteEx(VendorID, ProductID, BufferOut(0))
        End
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MAXIMO(1) = (FORCA(1) - 1000 * PONTO_ZERO(1)) / 1000
        Button2.Text = Format(MAXIMO(1), "###0") + " gf"
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        MAXIMO(2) = (FORCA(2) - 1000 * PONTO_ZERO(2)) / 1000
        Button10.Text = Format(MAXIMO(2), "###0") + " gf"
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        MAXIMO(3) = (FORCA(3) - 1000 * PONTO_ZERO(3)) / 1000
        Button11.Text = Format(MAXIMO(3), "###0") + " gf"
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        MAXIMO(4) = (FORCA(4) - 1000 * PONTO_ZERO(4)) / 1000
        Button12.Text = Format(MAXIMO(4), "###0") + " gf"
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 2 Then
            Button1.Visible = False
            Button2.Visible = False
            DataGridView1.Visible = False
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        TextBox4.BackColor = Color.White
        Timer2.Enabled = False
    End Sub

    Private Sub Tmr_WD_Tick(sender As Object, e As EventArgs) Handles Tmr_WD.Tick
        Label26.Text = (Date.Now)
        BufferOut(0) = 0
        BufferOut(1) = Asc("W")
        hidWriteEx(VendorID, ProductID, BufferOut(0))
        conta_wd = conta_wd + 1
        TextBox33.Text = conta_wd
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        CALIBRACAO = True
        PONTO_ZERO(2) = FORCA(2) / 1000
        TextBox25.Text = Format(PONTO_ZERO(2), "###0.00")
        PONTO_ZERO_CALIBRADO(2) = 0
        TabControl1.SelectedIndex = 0
        BTN_INICIAR.Enabled = CALIBRACAO
        MAXIMO(2) = 0
        Button10.Text = Format(1000 * (MAXIMO(2) + PONTO_ZERO(2)), "###0") + " gf"
        Timer5.Enabled = False
        BTN_INICIAR.Text = "INICIAR"
        TextBox3.Text = ""
        TextBox15.Text = ""
        TextBox19.Text = ""
        TextBox23.Text = ""

        TextBox5.Text = ""
        TextBox11.Text = ""
        TextBox17.Text = ""
        TextBox21.Text = ""

        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox16.Text = ""
        TextBox20.Text = ""
        LIGADO = False

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ATUALIZA_VISUAL(DataGridView1, Button2.Text)
        texto_saida = Date.Now & ";" & contagem_bbb & ";" & Button2.Text
        With My.Computer.FileSystem
            .WriteAllText(nome_arquivo, texto_saida & vbCrLf, True)
        End With
        MAXIMO(1) = 0
        Button2.Text = Format(1000 * (MAXIMO(1) + PONTO_ZERO(1) - ZERO_GRADE(1)), "###0") + " gf"
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ATUALIZA_VISUAL(DataGridView2, Button10.Text)
        texto_saida = Date.Now & ";" & contagem_bbb & ";" & Button10.Text
        With My.Computer.FileSystem
            .WriteAllText(nome_arquivo, texto_saida & vbCrLf, True)
        End With
        MAXIMO(2) = 0
        Button10.Text = Format(1000 * (MAXIMO(2) + PONTO_ZERO(2) - ZERO_GRADE(2)), "###0") + " gf"
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        ATUALIZA_VISUAL(DataGridView3, Button11.Text)
        texto_saida = Date.Now & ";" & contagem_bbb & ";" & Button11.Text
        With My.Computer.FileSystem
            .WriteAllText(nome_arquivo, texto_saida & vbCrLf, True)
        End With
        MAXIMO(3) = 0
        Button11.Text = Format(1000 * (MAXIMO(3) + PONTO_ZERO(3) - ZERO_GRADE(3)), "###0") + " gf"
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        ATUALIZA_VISUAL(DataGridView4, Button12.Text)
        texto_saida = Date.Now & ";" & contagem_bbb & ";" & Button12.Text
        With My.Computer.FileSystem
            .WriteAllText(nome_arquivo, texto_saida & vbCrLf, True)
        End With
        MAXIMO(4) = 0
        Button12.Text = Format(1000 * (MAXIMO(4) + PONTO_ZERO(4) - ZERO_GRADE(4)), "###0") + " gf"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        With My.Computer.FileSystem
            .DeleteFile("C:\Algometro\Config.txt")
            .WriteAllText("C:\Algometro\config.txt", TB_Arquivo.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", TB_Horario.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", TB_Nome.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", Tbx_massa.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", TB_Avaliador.Text & vbCrLf, True)
            .WriteAllText("C:\Algometro\config.txt", TB_Obs.Text & vbCrLf, True)
        End With
        INICIALIZAR()
        'System.Threading.Thread.Sleep(600) ' TEMPORARIO: removido para não travar UI
        data_hoje = Date.Now
        data = Format(Date.Now, "ddMMyyyy_HHmmss")
        BTN_INICIAR.Text = "PARAR"
        DataGridView1.RowCount = 1
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        CALIBRACAO = True
        PONTO_ZERO(3) = FORCA(3) / 1000
        TextBox28.Text = Format(PONTO_ZERO(3), "###0.00")
        PONTO_ZERO_CALIBRADO(3) = 0
        TabControl1.SelectedIndex = 0
        BTN_INICIAR.Enabled = CALIBRACAO
        MAXIMO(3) = 0
        Button11.Text = Format(1000 * (MAXIMO(3) + PONTO_ZERO(3)), "###0") + " gf"
        Timer5.Enabled = False
        BTN_INICIAR.Text = "INICIAR"
        LIGADO = False
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        CALIBRACAO = True
        PONTO_ZERO(4) = FORCA(4) / 1000
        TextBox31.Text = Format(PONTO_ZERO(4), "###0.00")
        PONTO_ZERO_CALIBRADO(4) = 0
        TabControl1.SelectedIndex = 0
        BTN_INICIAR.Enabled = CALIBRACAO
        MAXIMO(4) = 0
        Button12.Text = Format(1000 * (MAXIMO(4) + PONTO_ZERO(4)), "###0") + " gf"
        Timer5.Enabled = False
        BTN_INICIAR.Text = "INICIAR"
        LIGADO = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CALIBRACAO = True
        PONTO_ZERO(1) = FORCA(1) / 1000
        TextBox13.Text = Format(PONTO_ZERO(1), "###0.00")
        PONTO_ZERO_CALIBRADO(1) = 0
        TabControl1.SelectedIndex = 0
        BTN_INICIAR.Enabled = CALIBRACAO
        MAXIMO(1) = 0
        Button2.Text = Format(1000 * (MAXIMO(1) + PONTO_ZERO(1)), "###0") + " gf"
        Timer5.Enabled = False
        BTN_INICIAR.Text = "INICIAR"
        LIGADO = False
    End Sub
End Class