Imports System.Management
Imports System.Net
Imports System.IO
Imports System.IO.File

Public Interface IfrmMaisEscola
    Sub EsconderPasta()
    Sub gravaArquivoini()
    Sub Historico(ByVal Tipo As String)
    Sub PesquisarBoletins(ByVal ApenasListar As Boolean)
    Sub StartDetection()
    Sub ChecaBoletim()
    Sub BoletimAvancado()
    Function pegaHTML(ByVal URL2 As String) As String
End Interface
Public Class frmMaisProfessor
    Implements IfrmMaisEscola

    Dim Professor_Usuario, Professor_Senha

    Dim meuanexo As String, meuarquivo As String, minhapasta As String
    Dim testaDrive
    Dim EnviadoFINAL = 0
    Dim linha = ""
    Dim contador = 0
    Dim NotaTexto As String
    Dim MinhaDisciplina, MinhaBimestre, MinhaTurma, MinhaData
    Dim MeuLocal = System.Reflection.Assembly.GetExecutingAssembly.Location.ToString

    Private WithEvents m_MediaConnectWatcher As New ManagementEventWatcher
    Private Delegate Sub IncrementProgressDelegate()

    Public Function CarregaDisciplina(TurmaEscolhida)

        Dim ndisciplina = 0
        Dim ArqDisciplinas = tAno.Value & "\professor.txt"

        '//CARREGA AS DISCIPLINAS...
        Dim objStream2 As New System.IO.FileStream(ArqDisciplinas, IO.FileMode.Open)
        Dim Arq2 As New System.IO.StreamReader(objStream2)
        Dim linhaTexto2 As String = Arq2.ReadLine

        Dim matriz() As String

        cbDisciplinas.Items.Clear()
        cbDisciplinas2.Items.Clear()
        cbDisciplinas3.Items.Clear()

        While Not linhaTexto2 Is Nothing

            matriz = linhaTexto2.Split(";")

            If matriz(0) = TurmaEscolhida Then

                cbDisciplinas.Items.Add(matriz(1).Trim)
                cbDisciplinas2.Items.Add(matriz(1).Trim)
                cbDisciplinas3.Items.Add(matriz(1).Trim)

                iDisciplinas(ndisciplina) = matriz(1).Trim

                ndisciplina = ndisciplina + 1
            End If

            linhaTexto2 = Arq2.ReadLine

        End While
        Arq2.Close()
        lDisciplinas.Text = ndisciplina

    End Function

    Public Sub AjustaJanela(Expandir As Boolean)

        If Expandir = True Then
            '796; 405
            Me.Size = New System.Drawing.Size(796, 405)
        Else
            '580; 405
            Me.Size = New System.Drawing.Size(580, 405)
        End If
    End Sub

    Public Sub EsconderPasta() Implements IfrmMaisEscola.EsconderPasta

        Try
            Dim myprocessD As System.Diagnostics.Process = New System.Diagnostics.Process()
            myprocessD.StartInfo.FileName = "attrib"
            myprocessD.StartInfo.Arguments = String.Format("+H {0}", ano)
            myprocessD.Start()
        Catch ex As Exception
        End Try

    End Sub

    Public Sub gravaArquivoini() Implements IfrmMaisEscola.gravaArquivoini

        Dim nome_arquivo_ini As String = nomeArquivoINI()

        WritePrivateProfileString("UltimoBoletim", "Disciplina", MinhaDisciplina, nome_arquivo_ini)
        WritePrivateProfileString("UltimoBoletim", "Bimestre", MinhaBimestre, nome_arquivo_ini)
        WritePrivateProfileString("UltimoBoletim", "Turma", MinhaTurma, nome_arquivo_ini)
        WritePrivateProfileString("UltimoBoletim", "Data", MinhaData, nome_arquivo_ini)

    End Sub

    Public Sub Historico(ByVal Tipo As String) Implements IfrmMaisEscola.Historico

        If Tipo = "Gravar" Then

            ' PROGRAMA...
            lb_Turma.Text = cbTurmas.Text
            lb_Bimestre.Text = testeAF & "o bimestre"
            lb_Data.Text = Format(Date.Now, "dd/MM/yyyy - HH:mm:ss")
            lb_Disciplina.Text = cbDisciplinas.Text

            MinhaTurma = lb_Turma.Text
            MinhaBimestre = lb_Bimestre.Text
            MinhaData = lb_Data.Text
            MinhaDisciplina = lb_Disciplina.Text

            gravaArquivoini()

        ElseIf Tipo = "Leitura" Then

            Dim nome_arquivo_ini As String = nomeArquivoINI()

            ' ARQUIVO...
            Try

                lb_Disciplina.Text = LeArquivoINI(nome_arquivo_ini, "UltimoBoletim", "Disciplina", MinhaDisciplina)
                lb_Bimestre.Text = LeArquivoINI(nome_arquivo_ini, "UltimoBoletim", "Bimestre", MinhaBimestre)
                lb_Turma.Text = LeArquivoINI(nome_arquivo_ini, "UltimoBoletim", "Turma", MinhaTurma)
                lb_Data.Text = LeArquivoINI(nome_arquivo_ini, "UltimoBoletim", "Data", MinhaData)

            Catch ex As Exception
            End Try

            If lb_Disciplina.Text = "" Then
                lb_Disciplina.Text = "Nenhuma disciplina."
                lb_Bimestre.Text = "Nenhum bimestre."
                lb_Turma.Text = "Nenhuma turma."
                lb_Data.Text = "Nenhuma data."
            End If

        End If

    End Sub

    Public Sub PesquisarBoletins(ByVal ApenasListar As Boolean) Implements IfrmMaisEscola.PesquisarBoletins

        Array.Clear(EnviaDiretorios, 0, 9999)
        Array.Clear(EnviaArquivos, 0, 9999)
        EnvioContadorArquivos = 0
        EnvioContadorDiretorios = 0
        contador = 0
        ano = tAno4.Value

        MeuLocal = MeuLocal.Replace("file:\", "")

        ' VERIFICAR BOLETINS CADASTRADOS PELO PROFESSOR
        For Each T As String In cbTurmas.Items
            For Each D As String In cbDisciplinas.Items

                ' não deixar com Ç ÁÉÍÓÚ ÃÕ
                If tBimestre4.Value = "1" Then
                    TestarCaminho1 = String.Format("{0}\{1}\{2}\1.txt", ano, T, D)
                ElseIf tBimestre4.Value = "2" Then
                    TestarCaminho2 = String.Format("{0}\{1}\{2}\2.txt", ano, T, D)
                    TestarCaminho2AF = String.Format("{0}\{1}\{2}\2AF.txt", ano, T, D)
                ElseIf tBimestre4.Value = "3" Then
                    TestarCaminho3 = String.Format("{0}\{1}\{2}\3.txt", ano, T, D)
                ElseIf tBimestre4.Value = "4" Then
                    TestarCaminho4 = String.Format("{0}\{1}\{2}\4.txt", ano, T, D)
                    TestarCaminho4AF = String.Format("{0}\{1}\{2}\4AF.txt", ano, T, D)
                End If

                ' listar...
                If My.Computer.FileSystem.FileExists(TestarCaminho1) = True Then
                    linha = String.Format("{0} - CADASTRADO!;", TestarCaminho1)
                    txtImportacao.AppendText(linha & vbCrLf)
                    contador = contador + 1

                    EnvioContadorArquivos = EnvioContadorArquivos + 1
                    EnviaArquivos(EnvioContadorArquivos) = TestarCaminho1

                End If

                If My.Computer.FileSystem.FileExists(TestarCaminho2) = True Then
                    linha = String.Format("{0} - CADASTRADO!;", TestarCaminho2)
                    txtImportacao.AppendText(linha & vbCrLf)
                    contador = contador + 1

                    EnvioContadorArquivos = EnvioContadorArquivos + 1
                    EnviaArquivos(EnvioContadorArquivos) = TestarCaminho2


                End If

                If My.Computer.FileSystem.FileExists(TestarCaminho3) = True Then
                    linha = String.Format("{0} - CADASTRADO!;", TestarCaminho3)
                    txtImportacao.AppendText(linha & vbCrLf)
                    contador = contador + 1

                    EnvioContadorArquivos = EnvioContadorArquivos + 1
                    EnviaArquivos(EnvioContadorArquivos) = TestarCaminho3

                End If

                If My.Computer.FileSystem.FileExists(TestarCaminho4) = True Then
                    linha = String.Format("{0} - CADASTRADO!;", TestarCaminho4)
                    txtImportacao.AppendText(linha & vbCrLf)
                    contador = contador + 1

                    EnvioContadorArquivos = EnvioContadorArquivos + 1
                    EnviaArquivos(EnvioContadorArquivos) = TestarCaminho4

                End If

                If My.Computer.FileSystem.FileExists(TestarCaminho2AF) = True Then
                    linha = String.Format("{0} - CADASTRADO!;", TestarCaminho2AF)
                    txtImportacao.AppendText(linha & vbCrLf)
                    contador = contador + 1

                    EnvioContadorArquivos = EnvioContadorArquivos + 1
                    EnviaArquivos(EnvioContadorArquivos) = TestarCaminho2AF

                End If

                If My.Computer.FileSystem.FileExists(TestarCaminho4AF) = True Then
                    linha = String.Format("{0} - CADASTRADO!;", TestarCaminho4AF)
                    txtImportacao.AppendText(linha & vbCrLf)
                    contador = contador + 1

                    EnvioContadorArquivos = EnvioContadorArquivos + 1
                    EnviaArquivos(EnvioContadorArquivos) = TestarCaminho4AF

                End If

            Next
        Next

    End Sub

    Public Sub StartDetection() Implements IfrmMaisEscola.StartDetection
        Dim query2 As New WqlEventQuery("SELECT * FROM __InstanceOperationEvent WITHIN 1 " _
        & "WHERE TargetInstance ISA 'Win32_DiskDrive'")
        m_MediaConnectWatcher = New ManagementEventWatcher() With {.Query = query2}
        m_MediaConnectWatcher.Options.Timeout = New TimeSpan(1000)
        m_MediaConnectWatcher.Start()
    End Sub

    Private Sub Arrived(ByVal Sender As Object, ByVal E As System.Management.EventArrivedEventArgs) Handles m_MediaConnectWatcher.EventArrived
        Dim mbo, obj As ManagementBaseObject
        mbo = CType(E.NewEvent, ManagementBaseObject)
        obj = CType(mbo("TargetInstance"), ManagementBaseObject)
        'MsgBox(mbo.ClassPath.ClassName)
        Select Case mbo.ClassPath.ClassName
            Case "__InstanceDeletionEvent"
                If obj("InterfaceType").ToString.ToUpper = "USB" Then
                    'MsgBox(obj("Caption") & " has been unplugged") 
                    'Me.NotifyIcon1.Visible = True
                    Me.Enabled = False
                    Me.Visible = True
                    MsgBox("O Pen-drive foi RETIRADO indevidamente! Favor informar à secretaria da escola!", MsgBoxStyle.Information, NomePrograma)
                    End
                    'MsgBox(obj("Caption").ToString & " has been unplugged")
                Else
                    ' MsgBox(obj("InterfaceType"))
                End If
            Case Else
                'MsgBox("nope: " & obj("Caption").ToString)
        End Select
    End Sub

    Public Sub ChecaBoletim() Implements IfrmMaisEscola.ChecaBoletim

        Dim AnoVigente = tAno.Value

        Array.Clear(rST, 0, 100)
        riST = 1
        nroaluno = 1

        If tBimestre.Value = 2 And cbAF.Checked = True Then
            testeAF = tBimestre.Value & "AF"
        ElseIf tBimestre.Value = 4 And cbAF.Checked = True Then
            testeAF = tBimestre.Value & "AF"
        Else
            testeAF = tBimestre.Value
        End If

        If cbDisciplinas.Text = "" Or cbTurmas.Text = "" Then
            Exit Sub
        Else

            'If tBimestre.Value = 1 Then
            '    tbMedia.Enabled = True
            '    tbFaltas.Enabled = True
            '    tbAC.Enabled = False
            '    tbSN.Enabled = True
            'ElseIf (tBimestre.Value > 1) Then
            tbMedia.Enabled = True
            tbFaltas.Enabled = True
            tbAC.Enabled = True
            tbSN.Enabled = True
            '  End If

            If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\{3}.txt", AnoVigente, cbTurmas.Text, cbDisciplinas.Text, testeAF))) Then
                btCadastrar.Enabled = False
                tbStatus.Text = "Boletim já cadastrado!"
                tbStatus.ForeColor = Color.Red
                tbMedia.Text = ""
                tbFaltas.Text = "0"
                tbAC.Text = "0"
                tbSN.Text = "N"
                tbMedia.Enabled = False
                tbFaltas.Enabled = False
                tbAC.Enabled = False
                tbSN.Enabled = False
                travaBoletim = 1
                Exit Sub
            Else
                travaBoletim = 0
                btCadastrar.Enabled = True
                tbStatus.Text = "Cadastro em andamento..."
                tbStatus.ForeColor = Color.Blue
                tbMedia.Text = ""
                tbFaltas.Text = "0"
                tbAC.Text = "0"
                tbSN.Text = "N"
                tbMedia.Enabled = True
                tbFaltas.Enabled = True
                tbAC.Enabled = True
                tbSN.Enabled = True
            End If
        End If

    End Sub

    Public Sub BoletimAvancado() Implements IfrmMaisEscola.BoletimAvancado

        '// CONSULTA SE É AVANCADO O BOLETIM MOVEL...
        Try
            Dim AnoVigente = tAno.Value

            If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{1}a.txt", AnoVigente, cbTurmas.Text))) Then

                Dim lineST As String
                Dim arquivo1 As String = String.Format("{0}\{1}\{1}a.txt", AnoVigente, cbTurmas.Text)
                Dim ArqST As New System.IO.StreamReader(arquivo1)
                lineST = ArqST.ReadLine
                Dim SomaEvasaoEscolar = 0


                While Not lineST Is Nothing

                    Dim matriz As String() = lineST.Split(";")

                    'Posição da Evasão Escolar...
                    rST(riST) = matriz(1)

                    If Not IsNumeric(rST(riST)) Then
                        SomaEvasaoEscolar = SomaEvasaoEscolar + 1
                    End If

                    'Próximo registro...
                    riST = riST + 1
                    lineST = ArqST.ReadLine

                End While
                ArqST.Close()

                nroEvasaoTotal.Text = SomaEvasaoEscolar

                ' COM EVASAO ESCOLAR...
                If rST(lbCodigo.Text) <> "0" Then

                    NotaTexto = rST(lbCodigo.Text)
                    tbMedia.Text = NotaTexto
                    tbFaltas.Text = "0"
                    tbAC.Text = "0"

                    If cbAF.Checked = False Then
                        tbSN.Text = "N"
                    Else
                        tbSN.Text = "3"
                    End If
                    tbMedia.Enabled = False
                    tbFaltas.Enabled = False
                    tbAC.Enabled = False
                    tbSN.Enabled = False
                Else
                    ' SEM EVASAO ESCOLAR...

                    tbMedia.Text = ""
                    tbFaltas.Text = "0"
                    tbAC.Text = "0"
                    If cbAF.Checked = False Then
                        tbSN.Text = "N"
                    Else
                        tbSN.Text = "1"
                    End If

                    tbMedia.Enabled = True
                    tbFaltas.Enabled = True
                    tbSN.Enabled = True

                    'If tBimestre.Value = 1 Then
                    '    tbAC.Enabled = False
                    ' Else
                    tbAC.Enabled = True
                    '   End If

                    tbMedia.Focus()

                End If
                End If

        Catch ex As Exception
        End Try

    End Sub

    Public Function pegaHTML(ByVal URL2 As String) As String Implements IfrmMaisEscola.pegaHTML
        ' Retorna o HTML da URL informada
        Dim objWC As New System.Net.WebClient
        Return New System.Text.UTF8Encoding().GetString(objWC.DownloadData(URL2))

    End Function

    Private Sub frmPrincipal_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        '  SuporteRapido.Kill()

        If Not (cbDisciplinas.Enabled = False Or cbTurmas.Enabled = False) Then
            If MsgBox("Você deseja sair do programa?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Environment.Exit(0)
            Else
                e.Cancel = True
            End If
        Else
            If MsgBox("Se clicar em Sim, você perderá o boletim digitado. Tem certeza?", MsgBoxStyle.YesNo, "Deseja sair do programa?") = MsgBoxResult.Yes Then
                Environment.Exit(0)
            Else
                e.Cancel = True
            End If
        End If

    End Sub

    Private Sub frmPrincipal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ''#################################### VERIFICA O PROCESSO SE JÁ ESTÁ ATIVADO #################################
        Dim s() As Process     ' Gera um array de processos
        s = Process.GetProcessesByName(NomePrograma)  'Recupera todos os processos com o nome IEXPLORE
        If s.Length > 1 Then  ' Se o tamanho do array for > 0 quer dizer que o processo está ativo
            MsgBox("Este programa já foi aberto!", MsgBoxStyle.Information, NomePrograma)
            End
        End If
        ''############################################## FIM DE VERIFICACAO ###########################################

        AjustaJanela(False)

        ' SE RETIRAR O PEN-DRIVE ELE EXECUTA
        StartDetection()
        '

        'PUXA O BIMESTRE ATUAL................
        Dim atualBimestre = Format(Date.Now, "MM")
        If atualBimestre <= "05" Then
            tBimestre.Value = "1"
            tBimestre2.Value = "1"
            tBimestre3.Value = "1"
            tBimestre4.Value = "1"

        ElseIf atualBimestre <= "07" Then
            tBimestre.Value = "2"
            tBimestre2.Value = "2"
            tBimestre3.Value = "2"
            tBimestre4.Value = "2"

        ElseIf atualBimestre <= "10" Then
            tBimestre.Value = "3"
            tBimestre2.Value = "3"
            tBimestre3.Value = "3"
            tBimestre4.Value = "3"

        ElseIf atualBimestre <= "12" Then
            tBimestre.Value = "4"
            tBimestre2.Value = "4"
            tBimestre3.Value = "4"
            tBimestre4.Value = "4"
        End If

        'DEIXAR COMO PADRAO O ANO
        AnoVigente = Format(Date.Now, "yyyy")
        tAno.Value = AnoVigente
        tAno2.Value = AnoVigente
        tAno3.Value = AnoVigente
        tAno4.Value = AnoVigente

        ' Pega historico do ultimo boletim...
        Historico("Leitura")

        'Tipo de Nota
        Sistema_Tipo = LeArquivoINI(nome_arquivo_ini, "Sistema", "Escola", Escolar)

        Dim eBoletim = 0
        Dim VERSAO As Object = String.Format("{0}", My.Application.Info.Version.ToString)
        Me.Text = "Mais Professor! - v." & VERSAO
        tbStatus.Enabled = False

        ' obtem o caminho no VB .NET
        testaDrive = System.Reflection.Assembly.GetExecutingAssembly.Location.ToString
        testaDrive = Mid(testaDrive, 1, 3)
        tbStatus.Text = "Carregando Turmas..."
        mydrive = testaDrive

        If ((My.Computer.FileSystem.FileExists(tAno.Value & "\professor.txt") = False)) And ((My.Computer.FileSystem.FileExists(tAno.Value & "\turmas.txt") = False) Or (My.Computer.FileSystem.FileExists(tAno.Value & "\disciplinas.txt") = False)) Then

            MsgBox("Lista de turmas/disciplinas não foram encontradas!", MsgBoxStyle.Information)
            tbStatus.Text = "Listas não encontradas!"
            pNao.Visible = True
            pSim.Visible = False
            tbStatus.Enabled = False

        Else

            Try
                Dim nro = 0

                ArqTurmas = tAno.Value & "\turmas.txt"
                ArqDisciplinas = tAno.Value & "\disciplinas.txt"
                Dim ArqProfessor = tAno.Value & "\professor.txt"

                tbStatus.Text = "Carregando Turmas..."

                ' Se for turmas.txt e disciplinas.txt...
                If (My.Computer.FileSystem.FileExists(tAno.Value & "\turmas.txt") = True) Or (My.Computer.FileSystem.FileExists(tAno.Value & "\disciplinas.txt") = True) Then

                    '//CARREGA AS TURMAS...
                    Dim objStream As New System.IO.FileStream(ArqTurmas, IO.FileMode.Open)
                    Dim Arq As New System.IO.StreamReader(objStream)
                    Dim linhaTexto As String = Arq.ReadLine

                    cbTurmas.Items.Clear()
                    cbTurmas2.Items.Clear()
                    cbTurmas3.Items.Clear()

                    While Not linhaTexto Is Nothing
                        cbTurmas.Items.Add(linhaTexto)
                        cbTurmas2.Items.Add(linhaTexto)
                        cbTurmas3.Items.Add(linhaTexto)

                        iTurmas(nturma) = linhaTexto
                        linhaTexto = Arq.ReadLine
                        nturma = nturma + 1
                    End While

                    lTurmas.Text = nturma
                    Arq.Close()

                    tbStatus.Text = "Carregando Disciplinas..."

                    '//CARREGA AS DISCIPLINAS...
                    Dim objStream2 As New System.IO.FileStream(ArqDisciplinas, IO.FileMode.Open)
                    Dim Arq2 As New System.IO.StreamReader(objStream2)
                    Dim linhaTexto2 As String = Arq2.ReadLine

                    cbDisciplinas.Items.Clear()
                    cbDisciplinas2.Items.Clear()
                    cbDisciplinas3.Items.Clear()

                    While Not linhaTexto2 Is Nothing

                        cbDisciplinas.Items.Add(linhaTexto2)
                        cbDisciplinas2.Items.Add(linhaTexto2)
                        cbDisciplinas3.Items.Add(linhaTexto2)

                        iDisciplinas(ndisciplina) = linhaTexto2
                        linhaTexto2 = Arq2.ReadLine
                        ndisciplina = ndisciplina + 1


                    End While

                    lDisciplinas.Text = ndisciplina
                    Arq2.Close()

                Else
                    '..... PROFESSOR.TXT

                    '//CARREGA AS TURMAS...
                    Dim objStream As New System.IO.FileStream(ArqProfessor, IO.FileMode.Open)
                    Dim Arq As New System.IO.StreamReader(objStream)
                    Dim linhaTexto As String = Arq.ReadLine
                    Dim matriz() As String


                    cbTurmas.Items.Clear()
                    cbTurmas2.Items.Clear()
                    cbTurmas3.Items.Clear()

                    While Not linhaTexto Is Nothing

                        matriz = linhaTexto.Split(";")

                        cbTurmas.Items.Add(matriz(0).Trim)
                        cbTurmas2.Items.Add(matriz(0).Trim)
                        cbTurmas3.Items.Add(matriz(0).Trim)

                        iTurmas(nturma) = matriz(0).Trim
                        linhaTexto = Arq.ReadLine
                        nturma = nturma + 1
                    End While

                    lTurmas.Text = nturma
                    Arq.Close()

                End If

                If ((My.Computer.FileSystem.FileExists(tAno.Value & "\professor.txt") = True) Or (My.Computer.FileSystem.FileExists(tAno.Value & "\turmas.txt") = True) And (My.Computer.FileSystem.FileExists(tAno.Value & "\disciplinas.txt") = True)) Then
                    'MsgBox("Lista de turmas e disciplinas foram encontradas", MsgBoxStyle.Information)
                    tbStatus.Text = "Listas encontradas!"
                    tbStatus.ForeColor = Color.Blue
                    pSim.Visible = True
                    pNao.Visible = False
                    tbStatus.Enabled = True
                    eBoletim = 1

                    tbStatus.Text = "Finalizado!"
                Else
                    'MsgBox("Lista de turmas e disciplinas não foram encontradas", MsgBoxStyle.Information)
                    tbStatus.Text = "Listas não encontradas!"
                    tbStatus.ForeColor = Color.Red
                    pNao.Visible = True
                    pSim.Visible = False
                    tbStatus.Enabled = False
                    eBoletim = 0

                    tbStatus.Text = "Finalizado!"
                End If
                tbStatus.Text = "Carregando Boletins..."

                '// ENCONTRAR BOLETINS NO DRIVE
                If eBoletim = 1 Then
                    Dim k As Integer
                    Dim j As Integer
                    Dim myturma
                    Dim mydisciplinas
                    Dim myano
                    Dim nboletim = 0

                    'Checa pela Turma
                    For k = 0 To nturma
                        ' Checa pela Disciplina
                        myturma = iTurmas(k)

                        For j = 0 To ndisciplina

                            mydisciplinas = iDisciplinas(j)
                            myano = tAno.Value

                            '// 1º Bimestre
                            If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\1.txt", myano, myturma, mydisciplinas)) = True) Then
                                nboletim = nboletim + 1

                                '// 2º Bimestre
                            ElseIf (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\2.txt", myano, myturma, mydisciplinas)) = True) Then
                                nboletim = nboletim + 1

                                '// 2ºAF Bimestre
                            ElseIf (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\2AF.txt", myano, myturma, mydisciplinas)) = True) Then
                                nboletim = nboletim + 1

                                '// 3º Bimestre
                            ElseIf (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\3.txt", myano, myturma, mydisciplinas)) = True) Then
                                nboletim = nboletim + 1

                                '// 4º Bimestre
                            ElseIf (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\4.txt", myano, myturma, mydisciplinas)) = True) Then
                                nboletim = nboletim + 1

                                '// 4ºAF Bimestre
                            ElseIf (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\4AF.txt", myano, myturma, mydisciplinas)) = True) Then
                                nboletim = nboletim + 1
                            End If
                        Next

                    Next

                    lBoletins.Text = nboletim

                    If nboletim = 0 Then
                        lBoletins.Text = "00"
                    End If

                End If

            Catch ex As Exception
                tbStatus.Text = "Disco inválido!"
                tbStatus.ForeColor = Color.Red
                tbStatus.Text = "Finalizado!"
            End Try

        End If
        tbStatus.Text = "Finalizado!"

    End Sub

    Private Sub tBimestre_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tBimestre.ValueChanged

        AjustaJanela(False)

        dgBoletim.Visible = False
        exibeBoletim = 0

        Array.Clear(rST, 0, 100)
        riST = 1
        nroaluno = 1

        ChecaBoletim()

        If travaBoletim = 0 Then
            BoletimAvancado()
        End If

        If (tBimestre.Value = 1) Or (tBimestre.Value = 3) Then
            cbAF.Enabled = False
        Else
            cbAF.Enabled = True
        End If


    End Sub

    Private Sub tbMedia_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbMedia.KeyPress

        If Sistema_Tipo = "" Or Sistema_Tipo = "Estadual" Then

            If e.KeyChar = Convert.ToChar(Keys.Return) Or (e.KeyChar = ","c) Then
                e.Handled = True
                tbFaltas.Focus()

                If tbFaltas.Enabled = False Then
                    tbSN.Focus()
                End If
            End If
        Else

            If e.KeyChar = Convert.ToChar(Keys.Return) Then
                e.Handled = True
                tbFaltas.Focus()

                If tbFaltas.Enabled = False Then
                    tbSN.Focus()
                End If
            End If

            If e.KeyChar = "," Then
                e.Handled = True
                SendKeys.Send(".")
            End If
        End If

    End Sub

    Private Sub tbMedia_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbMedia.LostFocus

        Try
            Dim nota As Integer = CInt(tbMedia.Text)

            If nota < 5 Then
                tbMedia.ForeColor = Color.Red
            Else
                tbMedia.ForeColor = Color.Blue
            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub tbMedia_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMedia.TextChanged

        If tbMedia.Text <> "" Then
            btFinalizar.Enabled = False
            btCadastrar.Enabled = True
        Else
            btFinalizar.Enabled = True
            btCadastrar.Enabled = False
        End If

    End Sub

    Private Sub tbFaltas_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbFaltas.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            If tbAC.Enabled = True Then
                tbAC.Focus()
            Else
                tbSN.Focus()
            End If
        End If

        If e.KeyChar = ","c Then
            e.Handled = True
            If tbAC.Enabled = True Then
                tbAC.Focus()
            Else
                tbSN.Focus()
            End If
        End If

    End Sub

    Private Sub tbSN_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbSN.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            btCadastrar.Focus()
        End If
    End Sub

    Private Sub tbAC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbAC.KeyPress
        If e.KeyChar = Convert.ToChar(Keys.Return) Then
            e.Handled = True
            tbSN.Focus()

        End If

        If e.KeyChar = ","c Then
            e.Handled = True
            tbSN.Focus()
        End If

    End Sub

    Private Sub tbAC_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbAC.LostFocus
        If tbFaltas.Text <> "" Then

            If tbAC.Text = "" Then
                MsgBox("Digite as aulas compensadas!", MsgBoxStyle.Information)
                tbAC.Focus()
                Exit Sub
            End If

        End If

    End Sub

    Private Sub cbDisciplinas_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbDisciplinas.KeyPress

        e.Handled = True

    End Sub

    Private Sub cbTurmas_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbTurmas.KeyPress

        e.Handled = True

    End Sub

    Private Sub tBimestre3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tBimestre3.ValueChanged

        If (tBimestre3.Value = 1) Or (tBimestre3.Value = 3) Then
            cbAF3.Enabled = False
        Else
            cbAF3.Enabled = True
        End If

        If (tBimestre3.Value = 2 And cbAF3.Checked = True) Or (tBimestre3.Value = 4 And cbAF3.Checked = True) Then
            tBimestre3.Enabled = False

        Else
            tBimestre3.Enabled = True

        End If

        If (My.Computer.FileSystem.FileExists("" & tAno3.Value & "\" & cbTurmas3.Text & "\" & cbDisciplinas3.Text & "\" & tBimestre3.Value & ".txt")) Then
            tbStatus3.Text = "Boletim encontrado!"
            tbStatus3.ForeColor = Color.Blue
            btExcluir.Enabled = True
            Exit Sub
        Else
            tbStatus3.Text = "Boletim não encontrado!"
            tbStatus3.ForeColor = Color.Red
            btExcluir.Enabled = False
        End If

    End Sub

    Private Sub cbDisciplinas3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbDisciplinas3.KeyPress
        e.Handled = True

    End Sub

    Private Sub cbDisciplinas3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDisciplinas3.SelectedIndexChanged

        AjustaJanela(False)

        dgBoletim.Visible = False
        exibeBoletim = 0

        If (My.Computer.FileSystem.FileExists("" & tAno3.Value & "\" & cbTurmas3.Text & "\" & cbDisciplinas3.Text & "\" & tBimestre3.Value & ".txt")) Then
            tbStatus3.Text = "Boletim encontrado!"
            tbStatus3.ForeColor = Color.Blue
            btExcluir.Enabled = True
            Exit Sub
        Else
            tbStatus3.Text = "Boletim não encontrado!"
            tbStatus3.ForeColor = Color.Red
            btExcluir.Enabled = False
        End If

    End Sub

    Private Sub tAno3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tAno3.ValueChanged

        If (My.Computer.FileSystem.FileExists("" & tAno3.Value & "\" & cbTurmas3.Text & "\" & cbDisciplinas3.Text & "\" & tBimestre3.Value & ".txt")) Then
            tbStatus3.Text = "Boletim encontrado!"
            tbStatus3.ForeColor = Color.Blue
            btExcluir.Enabled = True
            Exit Sub
        Else
            tbStatus3.Text = "Boletim não encontrado!"
            tbStatus3.ForeColor = Color.Red
            btExcluir.Enabled = False
        End If

    End Sub

    Private Sub tBimestre2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tBimestre2.ValueChanged

        AjustaJanela(False)

        dgBoletim.Visible = False
        exibeBoletim = 0

        If (tBimestre2.Value = 1) Or (tBimestre2.Value = 3) Then
            cbAF2.Enabled = False
        Else
            cbAF2.Enabled = True
        End If

        If (tBimestre2.Value = 2 And cbAF2.Checked = True) Or (tBimestre2.Value = 4 And cbAF2.Checked = True) Then
            tBimestre2.Enabled = False

        Else
            tBimestre2.Enabled = True

        End If


        If (My.Computer.FileSystem.FileExists("" & tAno2.Value & "\" & cbTurmas2.Text & "\" & cbDisciplinas2.Text & "\" & tBimestre2.Value & ".txt")) Then
            tbStatus2.Text = "Boletim encontrado!"
            tbStatus2.ForeColor = Color.Blue
            btConsultar.Enabled = True
            tbMedia2.Text = ""
            tbFaltas2.Text = ""
            tbAC2.Text = ""
            tbSN2.Text = ""

            Exit Sub
        Else
            tbStatus2.Text = "Boletim não encontrado!"
            tbStatus2.ForeColor = Color.Red
            btConsultar.Enabled = False
            tbMedia2.Text = ""
            tbFaltas2.Text = ""
            tbAC2.Text = ""
            tbSN2.Text = ""

        End If

    End Sub

    Private Sub cbDisciplinas2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbDisciplinas2.KeyPress
        e.Handled = True
    End Sub

    Private Sub cbTurmas2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbTurmas2.KeyPress
        e.Handled = True
    End Sub

    Private Sub nAluno_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nAluno.ValueChanged

        If (My.Computer.FileSystem.FileExists("" & tAno.Value & "\" & cbTurmas2.Text & "\" & cbTurmas2.Text & "a.txt")) Then

            'If tBimestre2.Value = 1 Then
            '    tbMedia2.Text = EvasaoEscolar(rAV(nAluno.Value))
            '    tbFaltas2.Text = rFT(nAluno.Value)
            '    tbAC2.Text = rAC(nAluno.Value)
            '    tbSN2.Text = rPR(nAluno.Value)
            '    tbAC2.Enabled = False

            '   ElseIf (tBimestre2.Value > 1) Then

            tbMedia2.Text = EvasaoEscolar(rAV(nAluno.Value))
            tbFaltas2.Text = rFT(nAluno.Value)
            tbAC2.Text = rAC(nAluno.Value)
            tbSN2.Text = rPR(nAluno.Value)

            '  End If

            'SE NAO EXISTIR O BOLETIM AVANCADO................
        Else

            'If tBimestre2.Value = 1 Then

            '    tbMedia2.Text = EvasaoEscolar(rAV(nAluno.Value))
            '    tbFaltas2.Text = rFT(nAluno.Value)
            '    tbSN2.Text = rPR(nAluno.Value)

            'ElseIf (tBimestre2.Value > 1) Then

            tbMedia2.Text = EvasaoEscolar(rAV(nAluno.Value))
            tbFaltas2.Text = rFT(nAluno.Value)
            tbAC2.Text = rAC(nAluno.Value)
            tbSN2.Text = rPR(nAluno.Value)

            '   End If
        End If

        Try

            Dim Aluno As Integer = nAluno.Value

            If Aluno > atual Then
                dgBoletim.Rows(Aluno - 2).Selected = False
                dgBoletim.Rows(Aluno - 1).Selected = True
                atual = Aluno
            Else
                dgBoletim.Rows(Aluno - 1).Selected = True
                dgBoletim.Rows(Aluno).Selected = False
                atual = Aluno
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub tbMedia2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbMedia2.KeyPress

        If Sistema_Tipo = "" Or Sistema_Tipo = "Estadual" Then
            If e.KeyChar = ","c Then
                e.Handled = True
                tbFaltas2.Focus()
            End If
        Else
            If e.KeyChar = "," Then
                e.Handled = True
                SendKeys.Send(".")
            End If
        End If

    End Sub

    Private Sub tbFaltas2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbFaltas2.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = True
            tbAC2.Focus()
        End If
    End Sub

    Private Sub tbFaltas_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbFaltas.LostFocus
        If tbFaltas.Text = "" Then
            MsgBox("Digite a falta!", MsgBoxStyle.Information)
            tbFaltas.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub cbAv_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAF.CheckedChanged

        'INICIO PARA CHECAR / ZERAR

        Array.Clear(rST, 0, 100)

        Array.Clear(AV, 0, 100)
        Array.Clear(FT, 0, 100)
        Array.Clear(AC, 0, 100)
        Array.Clear(PR, 0, 100)

        riST = 1
        nroaluno = 1

        ChecaBoletim()
        BoletimAvancado()

        If tBimestre.Value = 1 Or tBimestre.Value = 3 Then
            MsgBox("Atenção! Avaliação Final somente para 2º e 4º bimestre!", MsgBoxStyle.Information, "Atenção")
            tBimestre.Enabled = True
            'cbAF.Checked = False

        Else
            tQtdade.Value = 0
            pQtdade.Value = 0
            tBimestre.Enabled = False
            tbFaltas.Enabled = False
            tbFaltas.Text = "0"
            tbAC.Enabled = False
            tbAC.Text = "0"
            MsgBox("1 - Aprovado (na disciplina)" & vbCrLf & "3 - Retido por Freqüência Insuficiente (na disciplina)" & vbCrLf & "4 - Retido por Rendimento Insuficiente (na disciplina)", MsgBoxStyle.Information, "Informações de Situação Final")

        End If


        If cbAF.Checked = False Then
            tBimestre.Enabled = True
            tbFaltas.Enabled = True
            tbAC.Enabled = True
            lbPR.Text = "PR"

            tQtdade.Minimum = "1"
            tQtdade.Value = "0"
            pQtdade.Minimum = "1"
            pQtdade.Value = "0"
            tQtdade.Enabled = True
            pQtdade.Enabled = True
            tbFaltas.Enabled = True
            tbAC.Enabled = True

        Else
            tbFaltas.Enabled = False
            tbAC.Enabled = False
            lbPR.Text = "ST"

            tQtdade.Minimum = "0"
            tQtdade.Value = "0"
            pQtdade.Minimum = "0"
            pQtdade.Value = "0"
            tQtdade.Enabled = False
            pQtdade.Enabled = False
            tbFaltas.Enabled = False
            tbAC.Enabled = False


        End If

        '// VERIFICA SE JA EXISTE O 2AF OU 4AF
        If cbAF.Checked = False Then

            If (My.Computer.FileSystem.FileExists("" & tAno.Value & "\" & cbTurmas.Text & "\" & cbDisciplinas.Text & "\" & tBimestre.Value & ".txt")) Then
                btCadastrar.Enabled = False
                tbStatus.Text = "Boletim já cadastrado!"
                tbStatus.ForeColor = Color.Red
                tbMedia.Enabled = False
                tbFaltas.Enabled = False
                tbAC.Enabled = False
                tbSN.Enabled = False
                Exit Sub
            Else
                btCadastrar.Enabled = True
                tbStatus.Text = "Cadastro em andamento..."
                tbStatus.ForeColor = Color.Blue
            End If


            'If tBimestre.Value = 1 Then
            '    tbMedia.Enabled = True
            '    tbFaltas.Enabled = True
            '    tbAC.Enabled = False
            '    tbSN.Enabled = True
            '   ElseIf (tBimestre.Value > 1) Then
            tbMedia.Enabled = True
            tbFaltas.Enabled = True
            tbAC.Enabled = True
            tbSN.Enabled = True
            'Else
            '    btCadastrar.Enabled = False
            '    tbMedia.Enabled = False
            '    tbFaltas.Enabled = False
            '    tbAC.Enabled = False
            '    tbSN.Enabled = False
            '    tbStatus.Text = "Bimestre inválido!"
            '    tbStatus.ForeColor = Color.Red
            'End If

        Else
            If (My.Computer.FileSystem.FileExists("" & tAno.Value & "\" & cbTurmas.Text & "\" & cbDisciplinas.Text & "\" & tBimestre.Value & "AF.txt")) Then
                btCadastrar.Enabled = False
                tbStatus.Text = "Boletim já cadastrado!"
                tbStatus.ForeColor = Color.Red
                tbMedia.Enabled = False
                tbFaltas.Enabled = False
                tbAC.Enabled = False
                tbSN.Enabled = False
                Exit Sub
            Else
                btCadastrar.Enabled = True
                tbStatus.Text = "Cadastro em andamento..."
                tbStatus.ForeColor = Color.Blue
            End If


            'If tBimestre.Value = 1 Then
            '    tbMedia.Enabled = True
            '    tbFaltas.Enabled = True
            '    tbAC.Enabled = False
            '    tbSN.Enabled = True
            'ElseIf (tBimestre.Value > 1) Then
            tbMedia.Enabled = True
            tbFaltas.Enabled = True
            tbAC.Enabled = True
            tbSN.Enabled = True
            'Else
            '    btCadastrar.Enabled = False
            '    tbMedia.Enabled = False
            '    tbFaltas.Enabled = False
            '    tbAC.Enabled = False
            '    tbSN.Enabled = False
            '    tbStatus.Text = "Bimestre inválido!"
            '    tbStatus.ForeColor = Color.Red
            'End If
        End If

    End Sub

    Private Sub cbAv2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAF2.CheckedChanged

        If (tBimestre2.Value = 1 Or tBimestre2.Value = 3) Then
            MsgBox("Atenção! Avaliação Final somente para 2º e 4º bimestre!", MsgBoxStyle.Information, "Atenção")
            tBimestre2.Enabled = True
        ElseIf (cbAF2.Checked = True) Then
            tBimestre2.Enabled = False
            MsgBox("1 - Aprovado (na disciplina)" & vbCrLf & "3 - Retido por Freqüência Insuficiente (na disciplina)" & vbCrLf & "4 - Retido por Rendimento Insuficiente (na disciplina)", MsgBoxStyle.Information, "Informações de Situação Final")
        End If

        If cbAF2.Checked = False Then
            tBimestre2.Enabled = True
            tbFaltas2.Enabled = True
            lbPR2.Text = "PR"
            tbAC2.Enabled = True
        Else
            lbPR2.Text = "ST"
            tbAC2.Enabled = False
            tbFaltas2.Enabled = False
            tQtdade2.Value = 0
            pQtdade2.Value = 0
        End If

        '// CHECA SE É 2AF OU 4AF...
        If cbAF2.Checked = False Then

            If (My.Computer.FileSystem.FileExists("" & tAno2.Value & "\" & cbTurmas2.Text & "\" & cbDisciplinas2.Text & "\" & tBimestre2.Value & ".txt")) = True Then
                tbStatus2.Text = "Boletim encontrado!"
                tbStatus2.ForeColor = Color.Blue
                btConsultar.Enabled = True
                Exit Sub
            Else
                tbStatus2.Text = "Boletim não encontrado!"
                tbStatus2.ForeColor = Color.Red
                btConsultar.Enabled = False
            End If

        Else

            If (My.Computer.FileSystem.FileExists("" & tAno2.Value & "\" & cbTurmas2.Text & "\" & cbDisciplinas2.Text & "\" & tBimestre2.Value & "AF.txt")) = True Then
                tbStatus2.Text = "Boletim encontrado!"
                tbStatus2.ForeColor = Color.Blue
                btConsultar.Enabled = True
                Exit Sub
            Else
                tbStatus2.Text = "Boletim não encontrado!"
                tbStatus2.ForeColor = Color.Red
                btConsultar.Enabled = False
            End If

        End If

    End Sub

    Private Sub cbAF3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAF3.CheckedChanged

        '... Se nao for 2 ou 3 Bimestre...
        If tBimestre3.Value = 1 Or tBimestre3.Value = 3 Then
            MsgBox("Atenção! Avaliação Final somente para 2º e 4º bimestre!", MsgBoxStyle.Information, "Atenção")
            tBimestre3.Enabled = True

        Else

            If tBimestre3.Value = 2 And cbAF3.Checked = True Then
                testeAF3 = "2AF"
            ElseIf tBimestre3.Value = 4 And cbAF3.Checked = True Then
                testeAF3 = "4AF"
            Else
                testeAF3 = tBimestre3.Value
            End If

            '// CHECA SE É 2AF OU 4AF
            If cbAF3.Checked = False Then
                If (My.Computer.FileSystem.FileExists("" & tAno3.Value & "\" & cbTurmas3.Text & "\" & cbDisciplinas3.Text & "\" & tBimestre3.Value & ".txt")) Then
                    tbStatus.Text = "Boletim encontrado!"
                    tbStatus.ForeColor = Color.Blue
                    btExcluir.Enabled = True
                    Exit Sub
                Else
                    tbStatus.Text = "Boletim não encontrado!"
                    tbStatus.ForeColor = Color.Red
                    btExcluir.Enabled = False
                End If
            Else
                If (My.Computer.FileSystem.FileExists("" & tAno3.Value & "\" & cbTurmas3.Text & "\" & cbDisciplinas3.Text & "\" & tBimestre3.Value & "AF.txt")) Then
                    tbStatus.Text = "Boletim encontrado!"
                    tbStatus.ForeColor = Color.Blue
                    btExcluir.Enabled = True
                    Exit Sub
                Else
                    tbStatus.Text = "Boletim não encontrado!"
                    tbStatus.ForeColor = Color.Red
                    btExcluir.Enabled = False
                End If
            End If

            If (tBimestre3.Value = 2 And cbAF3.Checked = True) Or (tBimestre3.Value = 4 And cbAF3.Checked = True) Then
                tBimestre3.Enabled = False
            Else
                tBimestre3.Enabled = True
            End If

        End If



    End Sub

    Private Sub tbAC2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbAC2.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = True
            tbSN2.Focus()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCadastrar.Click

        If tbStatus.Text <> "Boletim já cadastrado!" Then

            If tbMedia.Text <> "" Then

                'If (cbAF.Checked = False) And (tQtdade.Value = 0 Or pQtdade.Value = 0) Then
                '    MsgBox("Necessário digitar as aulas dadas e previstas!", MsgBoxStyle.Information)
                '    Exit Sub
                'End If

                Try
                    'Não deixa passar se professor digitar nota maior que 10...
                    Dim testaNota = tbMedia.Text
                    testaNota = testaNota.Substring(0, 2)
                    testaNota = CInt(testaNota)
                    If testaNota > 10 Then
                        MsgBox("A nota está incorreta!", MsgBoxStyle.Information)
                        Exit Sub
                    End If
                Catch ex As Exception
                End Try

                If tbFaltas.Text > tQtdade.Value Then
                    MsgBox("A Falta não pode ser maior que Aulas Dadas!", MsgBoxStyle.Information)
                    Exit Sub
                End If

                Dim testeSN
                If tBimestre.Value = 2 And cbAF.Checked = True Then
                    testeAF = tBimestre.Value & "AF"
                ElseIf tBimestre.Value = 4 And cbAF.Checked = True Then
                    testeAF = tBimestre.Value & "AF"
                Else
                    testeAF = tBimestre.Value
                End If
                ' Fim do teste AF

                If Not (Sistema_Tipo = "Particular" Or Sistema_Tipo = "Municipal") Then
                    If (My.Computer.FileSystem.FileExists(tAno.Value & "\ev.txt") = True) And (tbMedia.Enabled = True) Then
                        MaximoEvasao = 11
                    Else
                        MaximoEvasao = 31
                    End If
                Else
                    MaximoEvasao = 31.0
                End If

                If IsNumeric(testeMedia) And testeMedia < 11 Then
                    testeMedia = (EvasaoEscolar(tbMedia.Text))
                Else
                    testeMedia = tbMedia.Text
                End If


                ' Segue o cadastro da Avaliação!
                If MaximoEvasao = 31 Then
                    If Not IsNumeric(testeMedia) Then
                        testeMedia = (EvasaoEscolar(tbMedia.Text))
                        tbMedia.Text = testeMedia
                    End If

                Else
                    testeMedia = tbMedia.Text
                    tbMedia.Text = testeMedia

                    If Not IsNumeric(testeMedia) Then
                        MsgBox("Evasão Escolar não permitida!", MsgBoxStyle.Information, NomePrograma)
                        Exit Sub
                    End If
                End If

                ' Fim do cadastro da Avaliação...
                If (cbDisciplinas.Text <> "") Then
                    If (cbTurmas.Text <> "") Then

                        '// CHECA SE É AVALIACAO FINAL 2AF OU 4AF
                        If cbAF.Checked = True Then
                            If tbSN.Text = "1" Or tbSN.Text = "3" Or tbSN.Text = "4" Then
                                testeSN = "0"
                            Else
                                testeSN = "1"
                            End If
                        Else
                            If tbSN.Text = "N" Or tbSN.Text = "S" Then
                                testeSN = "0"
                            Else
                                testeSN = "1"
                            End If
                        End If

                        If testeSN = "0" Then

                            If Sistema_Tipo = "Particular" Or Sistema_Tipo = "Municipal" Then
                                testeMedia = testeMedia.Replace(".", ",")
                                Try
                                    testeMedia = CDbl(testeMedia)
                                Catch ex As Exception
                                    testeMedia = EvasaoEscolar(testeMedia)
                                End Try

                            End If

                            If (testeMedia < MaximoEvasao) Then

                                If (IsNumeric(tbFaltas.Text) = True) Then

                                    'Aceita a nota "quebrada"...
                                    '
                                    If Sistema_Tipo = "Particular" Or Sistema_Tipo = "Municipal" Then
                                        testeMedia = testeMedia.Replace(",", ".")
                                        testeMedia = CStr(testeMedia)
                                    End If

                                    '//VERIFICA SE EXISTE O AC
                                    If tbAC.Enabled = True Then
                                        If (IsNumeric(tbAC.Text) = False) Or (tbMedia.Text = "") Or (tbFaltas.Text = "") Or (tbAC.Text = "") Or (tbSN.Text = "") Then
                                            MsgBox("É necessário preencher as notas!", MsgBoxStyle.Information)
                                            Exit Sub
                                        End If
                                    ElseIf (tbAC.Enabled = False) Then
                                        If (tbMedia.Text = "") Or (tbFaltas.Text = "") Or (tbSN.Text = "") Then
                                            MsgBox("É necessário preencher as notas!", MsgBoxStyle.Information)
                                            Exit Sub
                                        End If
                                    End If

                                    '//VERIFICA SE EXISTE O ARQUIVO
                                    If trava = 0 Then
                                        If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\{3}.txt", tAno.Value, cbTurmas.Text, cbDisciplinas.Text, testeAF))) Then
                                            trava = 0
                                        Else
                                            trava = 1
                                        End If
                                    End If

                                    If trava = 1 Then
                                        '//SE NAO EXISTIR, SEGUE ABAIXO AS INSTRUCOES...
                                        cbTurmas.Enabled = False
                                        cbDisciplinas.Enabled = False
                                        tBimestre.Enabled = False
                                        tAno.Enabled = False
                                        tQtdade.Enabled = False
                                        pQtdade.Enabled = False
                                        cbAF.Enabled = False
                                        btFinalizar.Enabled = True

                                        AV(nroaluno) = testeMedia
                                        FT(nroaluno) = tbFaltas.Text
                                        AC(nroaluno) = tbAC.Text
                                        PR(nroaluno) = tbSN.Text

                                        If cbAF.Checked = False Then
                                            tbMedia.Text = ""
                                            tbFaltas.Text = "0"
                                            tbAC.Text = "0"

                                        Else
                                            tbMedia.Text = ""
                                            tbFaltas.Enabled = False
                                            tbFaltas.Text = "0"
                                            tbAC.Enabled = False
                                            tbAC.Text = "0"

                                        End If

                                        nroaluno = nroaluno + 1
                                        lbCodigo.Text = nroaluno
                                        tbMedia.Focus()

                                    Else
                                        '//SE EXISTIR, EXIBE A MENSAGEM A SEGUIR
                                        MsgBox("Desculpe, este boletim já foi cadastrado!", MsgBoxStyle.Information, "Boletim Móvel")
                                        trava = 0
                                        Exit Sub
                                    End If
                                Else
                                    MsgBox("Verifique os campos digitados!", MsgBoxStyle.Information, "Boletim Móvel")
                                End If
                            Else
                                MsgBox("Verifique os campos digitados!", MsgBoxStyle.Information, "Boletim Móvel")
                            End If
                        Else
                            MsgBox("Verifique os campos digitados!", MsgBoxStyle.Information, "Boletim Móvel")
                            tbSN.Focus()
                        End If
                    Else
                        MsgBox("Favor, escolher a turma e disciplina!", MsgBoxStyle.Information, "Atenção!")
                    End If
                Else
                    MsgBox("Favor, escolher a turma e disciplina!", MsgBoxStyle.Information, "Atenção!")
                End If

                '// CONSULTA SE É AVANCADO O BOLETIM MOVEL
                Try
                    Dim AnoVigente = tAno.Value
                    If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{1}a.txt", AnoVigente, cbTurmas.Text))) Then
                        If rST(nroaluno) Is Nothing Then
                            btFinalizar.Enabled = True
                            MsgBox("Todos os alunos foram cadastrados!", MsgBoxStyle.Information, NomePrograma)
                        Else
                            btFinalizar.Enabled = False
                        End If
                    End If
                Catch ex As Exception
                End Try

                If (tbMedia.Enabled = False) And (tbAC.Enabled = False) And (tbFaltas.Enabled = False) And (tbSN.Enabled = False) Then
                    btCadastrar.Focus()
                End If
            Else
                MsgBox("Verificar os campos!", MsgBoxStyle.Information, NomePrograma)
                tbMedia.Focus()
            End If

                btVoltar.Enabled = True

        End If


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btFinalizar.Click

        btFinalizar.Enabled = False

        Try
            '// Cria diretórios...
            ano = tAno.Value
            turma = String.Format("{0}\{1}", tAno.Value, cbTurmas.Text)
            disciplina = String.Format("{0}\{1}\{2}", tAno.Value, cbTurmas.Text, cbDisciplinas.Text)

            My.Computer.FileSystem.CreateDirectory(ano)
            My.Computer.FileSystem.CreateDirectory(turma)
            My.Computer.FileSystem.CreateDirectory(disciplina)

            '// Arquivo do Boletim...
            Dim ArqBoletim = String.Format("{0}\{1}\{2}\{3}.txt", tAno.Value, cbTurmas.Text, cbDisciplinas.Text, testeAF)
            Dim objStreamE As New System.IO.FileStream(ArqBoletim, IO.FileMode.OpenOrCreate)
            Dim ArqME As New System.IO.StreamWriter(objStreamE)

            ArqME.WriteLine(cbDisciplinas.Text)
            ArqME.WriteLine(cbTurmas.Text)
            ArqME.WriteLine(tAno.Value)
            ArqME.WriteLine(testeAF)
            ArqME.WriteLine(tQtdade.Value)
            ArqME.WriteLine(pQtdade.Text)
            ArqME.Close()
            '// Fim - Arquivo do Boletim...

            '// ABAIXO AV-AVALIACAO, FT-FALTAS, AC-COMPENSACAO, PR-RECUPERACAO
            Dim ArqAV = String.Format("{0}\{1}\{2}\{3}boletim.txt", tAno.Value, cbTurmas.Text, cbDisciplinas.Text, testeAF)
            Dim objStreamM1 As New System.IO.FileStream(ArqAV, IO.FileMode.OpenOrCreate)
            Dim ArqM1 As New System.IO.StreamWriter(objStreamM1)
            Dim i

            For i = 1 To nroaluno - 1
                Dim NovoFormato = String.Format("{0};{1};{2};{3};{4};", i, AV(i), FT(i), AC(i), PR(i))
                ArqM1.WriteLine(NovoFormato)
            Next
            ArqM1.Close()

        Catch ex As Exception
            MsgBox(String.Format("Não foi possível gravar!{0}Erro: {1}", vbCrLf, ex.Message), MsgBoxStyle.Information, NomePrograma)
        End Try

        tbStatus.Text = "Boletim cadastrado!"
        tbStatus.ForeColor = Color.Blue

        Dim soma As Object = lBoletins.Text
        soma = soma + 1
        lBoletins.Text = soma

        cbTurmas.Enabled = True
        cbDisciplinas.Enabled = True
        tBimestre.Enabled = True
        tAno.Enabled = True
        tQtdade.Enabled = True
        pQtdade.Enabled = True
        btFinalizar.Enabled = False
        btCadastrar.Enabled = False

        lbCodigo.Text = 1
        tbMedia.Text = ""
        tbSN.Text = "N"
        tbFaltas.Text = "0"
        tbAC.Text = "0"
        tQtdade.Value = 0
        pQtdade.Value = 0

        tbMedia.Enabled = False
        tbFaltas.Enabled = False
        tbAC.Enabled = False
        tbSN.Enabled = False
        btVoltar.Enabled = False
        cbAF.Enabled = True

        EsconderPasta()

        Historico("Gravar")

    End Sub
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btConsultar.Click

        AjustaJanela(False)

        dgBoletim.Visible = False
        exibeBoletim = 0

        If tBimestre2.Value = 2 And cbAF2.Checked = True Then
            testeAF2 = tBimestre2.Value & "AF"
            tbFaltas2.Enabled = False
            tbAC2.Enabled = False
        ElseIf tBimestre2.Value = 4 And cbAF2.Checked = True Then
            testeAF2 = tBimestre2.Value & "AF"
            tbFaltas2.Enabled = False
            tbAC2.Enabled = False
        Else
            testeAF2 = tBimestre2.Value
            tbFaltas2.Enabled = True
            ' If tBimestre2.Value > 1 Then
            tbAC2.Enabled = True
            '   Else
            '   tbAC2.Enabled = False
            '    End If
        End If

        riAV = 1
        riFT = 1
        riAC = 1
        riBL = 1
        riPR = 1

        If (cbDisciplinas2.Text = "") Or (cbTurmas2.Text = "") Then
            MsgBox("Para consultar é necessário preencher os campos!", MsgBoxStyle.Information, "Boletim Móvel")
            Exit Sub
        Else
            nAluno.Enabled = True

            '// Configurações do boletim...
            Dim arquivo = String.Format("{0}\{1}\{2}\{3}.txt", tAno2.Value, cbTurmas2.Text, cbDisciplinas2.Text, testeAF2)
            Dim objStreamER As New System.IO.StreamReader(arquivo)
            Dim line As String = objStreamER.ReadLine

            While line <> Nothing
                rBL(riBL) = line
                riBL = riBL + 1
                line = objStreamER.ReadLine
            End While
            objStreamER.Close()

            ' Boletim...
            Dim arquivo1 = String.Format("{0}\{1}\{2}\{3}boletim.txt", tAno2.Value, cbTurmas2.Text, cbDisciplinas2.Text, testeAF2)
            Dim ArqM1R As New System.IO.StreamReader(arquivo1)
            Dim line1 As String = ArqM1R.ReadLine

            While line1 <> Nothing

                Dim matriz() As String = line1.Split(";")
                matriz = line1.Split(";")

                ' Média Aluno
                rAV(riAV) = matriz(1).Trim
                ' Faltas Aluno
                rFT(riAV) = matriz(2).Trim
                ' AC Aluno
                rAC(riAV) = matriz(3).Trim
                ' Situação Aluno
                rPR(riAV) = matriz(4).Trim
                '
                'Próxima linha
                riAV = riAV + 1
                line1 = ArqM1R.ReadLine

            End While
            ArqM1R.Close()

            '/// Libera para consulta
            nAluno.Enabled = True
            nAluno.Maximum = riAV - 1
            tbMedia2.Enabled = True

            tbSN2.Enabled = True
            btConsultar.Enabled = True
            btAlterar.Enabled = True
            btClasse.Enabled = True

            '/// Joga os valores na tela
            'total
            tQtdade2.Value = rBL(5)
            ' previstas
            pQtdade2.Value = rBL(6)

            '// JOGA OS DADOS DE AVALIACAO

            tbMedia2.Text = EvasaoEscolar(rAV(nAluno.Value))
            tbFaltas2.Text = rFT(nAluno.Value)
            tbAC2.Text = rAC(nAluno.Value)
            tbSN2.Text = rPR(nAluno.Value)
            nromax = riAV

        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAlterar.Click

        Try
            'Não deixa passar se professor digitar nota maior que 10...
            Dim testaNota = tbMedia2.Text
            testaNota = testaNota.Substring(0, 2)
            testaNota = CInt(testaNota)
            If testaNota > 10 Then
                MsgBox("A nota está incorreta!", MsgBoxStyle.Information)
                Exit Sub
            End If
        Catch ex As Exception
        End Try

        If tbFaltas2.Text > tQtdade2.Value Then
            MsgBox("A Falta não pode ser maior que Aulas Dadas!", MsgBoxStyle.Information)
            Exit Sub
        End If

        Dim testeSN2
        testeMedia2 = tbMedia2.Text

        '// CHECA SE É AVALIACAO FINAL 2AF OU 4AF
        If cbAF2.Checked = True Then
            If tbSN2.Text = "1" Or tbSN2.Text = "3" Or tbSN2.Text = "4" Then
                testeSN2 = "0"
            Else
                testeSN2 = "1"
            End If
        Else
            If tbSN2.Text = "N" Or tbSN2.Text = "S" Then
                testeSN2 = "0"
            Else
                testeSN2 = "1"
            End If
        End If

        If testeSN2 = "0" Then
            If Sistema_Tipo = "Particular" Or Sistema_Tipo = "Municipal" Then
                testeMedia2 = testeMedia2.Replace(".", ",")
                Try
                    testeMedia2 = CDbl(testeMedia2)
                Catch ex As Exception
                    testeMedia2 = EvasaoEscolar(testeMedia2)
                End Try
            End If

            If Not IsNumeric(testeMedia2) Then
                testeMedia2 = (EvasaoEscolar(tbMedia2.Text))
                tbMedia2.Text = testeMedia2
            End If

            ' VALOR MAXIMO PRA NOTA
            If testeMedia2 < MaximoEvasao Then
                If (IsNumeric(tbFaltas2.Text) = True) Then

                    If Sistema_Tipo = "Particular" Or Sistema_Tipo = "Municipal" Then
                        testeMedia2 = CStr(testeMedia2)
                        testeMedia2 = testeMedia2.Replace(",", ".")

                    End If

                    '//VERIFICA SE EXISTE O AC
                    If tbAC2.Enabled = True Then
                        If (IsNumeric(tbFaltas2.Text) = False) Or (tbMedia2.Text = "") Or (tbFaltas2.Text = "") Or (tbAC2.Text = "") Or (tbSN2.Text = "") Or (IsNumeric(tbAC2.Text) = False) Then
                            MsgBox("É necessário preencher as notas!", MsgBoxStyle.Information)
                            Exit Sub
                        End If
                    ElseIf (tbAC2.Enabled = False) Then
                        If (tbMedia2.Text = "") Or (tbFaltas2.Text = "") Or (tbSN2.Text = "") Then
                            MsgBox("É necessário preencher as notas!", MsgBoxStyle.Information)
                            Exit Sub
                        End If
                    End If

                    Dim aluno = nAluno.Value
                    rAV(aluno) = testeMedia2
                    rFT(aluno) = tbFaltas2.Text
                    If tBimestre.Value > 1 Then
                        rAC(aluno) = tbAC2.Text
                    End If
                    rPR(aluno) = tbSN2.Text

                    '// Arquivo do Boletim
                    Dim ArqBoletim = String.Format("{0}\{1}\{2}\{3}.txt", tAno2.Value, cbTurmas2.Text, cbDisciplinas2.Text, testeAF2)

                    Dim objStreamE As New System.IO.FileStream(ArqBoletim, IO.FileMode.Create)
                    Dim ArqME As New System.IO.StreamWriter(objStreamE)

                    ArqME.WriteLine(cbDisciplinas2.Text)
                    ArqME.WriteLine(cbTurmas2.Text)
                    ArqME.WriteLine(tAno2.Value)
                    ArqME.WriteLine(testeAF2)
                    ArqME.WriteLine(tQtdade2.Value)
                    ArqME.WriteLine(pQtdade2.Text)
                    ArqME.Close()

                    '// Boletim
                    Dim objStreamM1 As New System.IO.FileStream(String.Format("{0}\{1}\{2}\{3}boletim.txt", tAno2.Value, cbTurmas2.Text, cbDisciplinas2.Text, testeAF2), IO.FileMode.Create)
                    Dim ArqM1 As New System.IO.StreamWriter(objStreamM1)
                    Dim i

                    For i = 1 To nromax - 1

                        'Verifica se Qtdade Total e Previstas são maiores.
                        '
                        If rFT(i) > tQtdade2.Value Then
                            rFT(i) = tQtdade2.Value
                        End If

                        Dim Alteracao = String.Format("{0};{1};{2};{3};{4};", i, rAV(i), rFT(i), rAC(i), rPR(i))
                        ArqM1.WriteLine(Alteracao)
                    Next
                    ArqM1.Close()

                    tbStatus2.Text = "Nota alterada!"
                    tbStatus2.ForeColor = Color.Blue
                    cbTurmas2.Enabled = True
                    cbDisciplinas2.Enabled = True
                    tBimestre2.Enabled = True
                    tAno2.Enabled = True
                    tQtdade2.Enabled = True
                    pQtdade2.Enabled = True

                Else
                    MsgBox("Verifique os campos digitados!", MsgBoxStyle.Information, "Boletim Móvel")
                End If
            Else
                MsgBox("Verifique os campos digitados!", MsgBoxStyle.Information, "Boletim Móvel")
                tbMedia2.Focus()
            End If
        Else
            MsgBox("Verifique os campos digitados!", MsgBoxStyle.Information, "Boletim Móvel")
            tbSN2.Focus()
        End If

        '###################EXIBIR O BOLETIM DO LADO...
        Dim mediaAVT = 0
        Dim mediaFTT = 0
        somaAV = 0
        somaFT = 0

        Try

            If exibeBoletim = 0 Then

                AjustaJanela(True)
                dgBoletim.Visible = True

                Dim dt As New DataTable
                dt.Columns.Add("Nro")
                dt.Columns.Add("AV")
                dt.Columns.Add("FT")

                If tBimestre2.Value > 1 Then
                    dt.Columns.Add("AC")
                End If

                dt.Columns.Add("PR")

                If tBimestre2.Value > 1 Then
                    For i As Integer = 1 To nAluno.Maximum
                        Dim media As String = EvasaoEscolar(rAV(i))
                        '// JOGA OS DADOS DE AVALIACAO
                        somaAV = somaAV + rAV(i)
                        somaFT = somaFT + rFT(i)
                        dt.Rows.Add(New Object() {i, media, rFT(i), rAC(i), rPR(i)})
                    Next
                Else
                    For i As Integer = 1 To nAluno.Maximum
                        Dim media As String = EvasaoEscolar(rAV(i))
                        '// JOGA OS DADOS DE AVALIACAO
                        somaAV = somaAV + rAV(i)
                        somaFT = somaFT + rFT(i)
                        dt.Rows.Add(New Object() {i, media, rFT(i), rPR(i)})
                    Next
                End If
                dgBoletim.DataSource = dt

                If tBimestre2.Value > 1 Then
                    dgBoletim.Columns(0).Width = "26"
                    dgBoletim.Columns(1).Width = "25"
                    dgBoletim.Columns(2).Width = "24"
                    dgBoletim.Columns(3).Width = "24"
                    dgBoletim.Columns(4).Width = "20"

                    dgBoletim.Columns(0).DefaultCellStyle.BackColor = Color.LightBlue

                    dgBoletim.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgBoletim.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgBoletim.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgBoletim.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgBoletim.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                Else
                    dgBoletim.Columns(0).Width = "26"
                    dgBoletim.Columns(1).Width = "25"
                    dgBoletim.Columns(2).Width = "24"
                    dgBoletim.Columns(3).Width = "20"

                    dgBoletim.Columns(0).DefaultCellStyle.BackColor = Color.LightBlue

                    dgBoletim.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgBoletim.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgBoletim.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    dgBoletim.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                End If

                dgBoletim.Columns(0).DefaultCellStyle.Font = New Font("Verdana", 8, FontStyle.Bold)
                dgBoletim.Columns(1).DefaultCellStyle.Font = New Font("Verdana", 8, FontStyle.Bold)
                dgBoletim.DefaultCellStyle.Font = New Font("Verdana", 8, FontStyle.Regular)

                exibeBoletim = 1

            Else

                AjustaJanela(False)

                dgBoletim.Visible = False
                exibeBoletim = 0

            End If

        Catch ex As Exception

        End Try

        Try
            For Each Linha As DataGridViewRow In Me.dgBoletim.Rows
                'Verifica se a célula do teu DataGridView tem o valor < "5"
                Try
                    If Linha.Cells(1).Value < 5 Then
                        Linha.Cells(1).Style.ForeColor = Color.Red
                        Linha.Cells(1).Style.Font = New Font("Verdana", 8, FontStyle.Bold)

                    Else
                        Linha.Cells(1).Style.ForeColor = Color.Blue
                        Linha.Cells(1).Style.Font = New Font("Verdana", 8, FontStyle.Bold)

                    End If
                Catch ex As Exception
                End Try
            Next

            mediaAVT = somaAV / CInt(nAluno.Maximum)
            mediaFTT = somaFT / CInt(nAluno.Maximum)

            If mediaAVT < 5 Then
                mediaAV.ForeColor = Color.Red
            Else
                mediaAV.ForeColor = Color.Blue
            End If

            Dim iFaltas = 0

            iFaltas = (tQtdade2.Value * 0.25)

            If mediaFTT > iFaltas Then
                mediaFT.ForeColor = Color.Red
            Else
                mediaFT.ForeColor = Color.Blue
            End If

            mediaAV.Text = mediaAVT
            mediaFT.Text = mediaFTT

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExcluir.Click

        Dim erro = 0

        Try

            If cbAF3.Checked = True Then

                '// Arquivo do Boletim
                My.Computer.FileSystem.DeleteFile(String.Format("{0}\{1}\{2}\{3}AF.txt", tAno3.Value, cbTurmas3.Text, cbDisciplinas3.Text, tBimestre3.Value))
                '// Boletim
                My.Computer.FileSystem.DeleteFile(String.Format("{0}\{1}\{2}\{3}AFboletim.txt", tAno3.Value, cbTurmas3.Text, cbDisciplinas3.Text, tBimestre3.Value))
            Else
                '// Arquivo do Boletim
                My.Computer.FileSystem.DeleteFile(String.Format("{0}\{1}\{2}\{3}.txt", tAno3.Value, cbTurmas3.Text, cbDisciplinas3.Text, tBimestre3.Value))
                '// Boletim
                My.Computer.FileSystem.DeleteFile(String.Format("{0}\{1}\{2}\{3}boletim.txt", tAno3.Value, cbTurmas3.Text, cbDisciplinas3.Text, tBimestre3.Value))
            End If

        Catch ex As IndexOutOfRangeException
            MsgBox("Alguns dos arquivos a serem excluídos não foram encontrados!", MsgBoxStyle.Information, "Atenção!")
            erro = 1
        End Try

        If erro = 0 Then

            cbDisciplinas3.Text = ""
            cbTurmas3.Text = ""
            btExcluir.Enabled = False
            tbStatus3.Text = "Boletim excluído!"

            Dim soma
            soma = lBoletins.Text
            soma = soma - 1
            lBoletins.Text = soma
            tbStatus3.ForeColor = Color.Blue

        End If
    End Sub

    Private Function ExecuteFile(ByVal FileName As String) As Boolean

        Dim myProcess As New Process
        myProcess.StartInfo.FileName = FileName
        myProcess.StartInfo.UseShellExecute = True

        myProcess.StartInfo.RedirectStandardOutput = False
        myProcess.Start()
        myProcess.Dispose()

    End Function

    Private Sub cbTurmas_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbTurmas.TextChanged

        If (My.Computer.FileSystem.FileExists(tAno.Value & "\professor.txt") = True) Then

            cbDisciplinas.Text = ""
            cbDisciplinas2.Text = ""
            cbDisciplinas3.Text = ""
            CarregaDisciplina(cbTurmas.Text)

        End If

        If cbTurmas.Text <> "" And cbDisciplinas.Text <> "" Then

            AjustaJanela(False)

            dgBoletim.Visible = False

            exibeBoletim = 0
            Array.Clear(rST, 0, 100)
            Array.Clear(AV, 0, 100)
            Array.Clear(FT, 0, 100)
            Array.Clear(AC, 0, 100)
            Array.Clear(PR, 0, 100)

            riST = 1
            nroaluno = 1

            ChecaBoletim()
            BoletimAvancado()

            If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\{3}.txt", tAno.Value, cbTurmas.Text, cbDisciplinas.Text, testeAF))) Then
                btCadastrar.Enabled = False
                tbStatus.Text = "Boletim já cadastrado!"
                tbFaltas.Text = "0"
                tbAC.Text = "0"
                tbSN.Text = "N"
                tbMedia.Enabled = False
                tbFaltas.Enabled = False
                tbAC.Enabled = False
                tbSN.Enabled = False
                btCadastrar.Enabled = False
                btFinalizar.Enabled = False

                travaBoletim = 1
                Exit Sub
            End If
        End If

    End Sub

    Private Sub tAno_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tAno.ValueChanged

        ChecaBoletim()

        If travaBoletim = 0 Then
            BoletimAvancado()
        End If

    End Sub

    Private Sub lbCodigo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbCodigo.TextChanged
        Try

            If travaBoletim = 0 Then
                Dim AnoVigente = CInt(tAno.Value)

                If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{1}a.txt", AnoVigente, cbTurmas.Text))) Then
                    If rST(lbCodigo.Text) <> "0" Then

                        tbMedia.Text = rST(lbCodigo.Text)
                        tbFaltas.Text = "0"
                        tbAC.Text = "0"

                        If cbAF.Checked = False Then
                            tbSN.Text = "N"
                        Else
                            tbSN.Text = "3"
                        End If

                        tbMedia.Enabled = False
                        tbFaltas.Enabled = False
                        tbAC.Enabled = False
                        tbSN.Enabled = False
                    Else

                        tbMedia.Text = ""
                        tbFaltas.Text = "0"
                        tbAC.Text = "0"
                        If cbAF.Checked = False Then
                            tbSN.Text = "N"
                        Else
                            tbSN.Text = "1"
                        End If

                        tbMedia.Enabled = True
                        tbFaltas.Enabled = True
                        tbSN.Enabled = True

                        'If tBimestre.Value = 1 Then
                        '    tbAC.Enabled = False
                        'Else
                        tbAC.Enabled = True
                        'End If

                        tbMedia.Focus()

                    End If

                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btBoletim_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btClasse.Click
        Try

            AjustaJanela(False)
            dgBoletim.Visible = True

            Dim mediaAVT = 0
            Dim mediaFTT = 0
            somaAV = 0
            somaFT = 0
            Dim somaEvasao = 0

            AjustaJanela(True)

            Dim dt As New DataTable
            dt.Columns.Add("Nro")
            dt.Columns.Add("AV")
            dt.Columns.Add("FT")

            '   If tBimestre2.Value > 1 Then
            dt.Columns.Add("AC")
            '   End If
            dt.Columns.Add("PR")

            '    If tBimestre2.Value > 1 Then
            For i As Integer = 1 To nAluno.Maximum
                Dim media As String = rAV(i)
                '// JOGA OS DADOS DE AVALIACAO
                If media < 11 Then
                    somaAV = somaAV + rAV(i)
                Else
                    somaEvasao = somaEvasao + 1
                End If

                somaFT = somaFT + rFT(i)
                dt.Rows.Add(New Object() {i, EvasaoEscolar(media), rFT(i), rAC(i), rPR(i)})
            Next
            'Else
            '    For i As Integer = 1 To nAluno.Maximum
            '        Dim media As String = rAV(i)

            '        '// JOGA OS DADOS DE AVALIACAO
            '        If media < 11 Then
            '            somaAV = somaAV + rAV(i)
            '        Else
            '            somaEvasao = somaEvasao + 1
            '        End If

            '        somaFT = somaFT + rFT(i)
            '        dt.Rows.Add(New Object() {i, EvasaoEscolar(media), rFT(i), rPR(i)})
            '    Next
            'End If
            dgBoletim.DataSource = dt

            '    If tBimestre2.Value > 1 Then
            dgBoletim.Columns(0).Width = "26"
            dgBoletim.Columns(1).Width = "30"
            dgBoletim.Columns(2).Width = "24"
            dgBoletim.Columns(3).Width = "24"
            dgBoletim.Columns(4).Width = "20"

            dgBoletim.Columns(0).DefaultCellStyle.BackColor = Color.LightBlue

            dgBoletim.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgBoletim.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgBoletim.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgBoletim.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgBoletim.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'Else
            '    dgBoletim.Columns(0).Width = "26"
            '    dgBoletim.Columns(1).Width = "30"
            '    dgBoletim.Columns(2).Width = "24"
            '    dgBoletim.Columns(3).Width = "20"

            '    dgBoletim.Columns(0).DefaultCellStyle.BackColor = Color.LightBlue

            '    dgBoletim.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '    dgBoletim.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '    dgBoletim.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '    dgBoletim.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'End If

            dgBoletim.Columns(0).DefaultCellStyle.Font = New Font("Verdana", 8, FontStyle.Bold)
            dgBoletim.Columns(1).DefaultCellStyle.Font = New Font("Verdana", 8, FontStyle.Bold)
            dgBoletim.DefaultCellStyle.Font = New Font("Verdana", 8, FontStyle.Regular)
            exibeBoletim = 1

            For Each Linha As DataGridViewRow In Me.dgBoletim.Rows
                ' Verifica se a célula do teu DataGridView tem o valor < "5"
                '
                Try
                    If Linha.Cells(1).Value < 5 Then
                        Linha.Cells(1).Style.ForeColor = Color.Red
                        Linha.Cells(1).Style.Font = New Font("Verdana", 8, FontStyle.Bold)
                    Else
                        Linha.Cells(1).Style.ForeColor = Color.Blue
                        Linha.Cells(1).Style.Font = New Font("Verdana", 8, FontStyle.Bold)
                    End If
                Catch ex As Exception
                End Try
            Next

            mediaAVT = somaAV / nAluno.Maximum
            mediaFTT = somaFT / nAluno.Maximum

            If mediaAVT < 5 Then
                mediaAV.ForeColor = Color.Red
            Else
                mediaAV.ForeColor = Color.Blue
            End If

            Dim iFaltas = (tQtdade2.Value * 0.25)

            If mediaFTT > iFaltas Then
                mediaFT.ForeColor = Color.Red
            Else
                mediaFT.ForeColor = Color.Blue
            End If

            mediaAV.Text = mediaAVT
            mediaFT.Text = mediaFTT
            nroEvasaoTotal.Text = somaEvasao

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, NomePrograma)
        End Try

    End Sub


    Private Sub tAno2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tAno2.ValueChanged

        ChecaBoletim()

        If travaBoletim = 0 Then
            BoletimAvancado()
        End If
    End Sub

    Private Sub Button1_Click_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btVoltar.Click
        BimestreAtual = testeAF
        AlunoAtual = lbCodigo.Text

        If cbAF.Checked = True Then
            AFAtual = True
        Else
            AFAtual = False
        End If

        frmVoltar.Show()
    End Sub

    Private Sub btFerramentas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btFerramentas.Click

        EnviadoFINAL = 0
        EnviadoFTP = True

        linha = ""
        contador = 0
        txtImportacao.Text = ""

        PesquisarBoletins(True)
        txtImportacao.AppendText(vbCrLf)
        linha = "*** Finalizado busca de boletins  ***"
        txtImportacao.AppendText(linha & vbCrLf)
        txtImportacao.AppendText(vbCrLf)
        If contador = 0 Then
            linha = "Resultado: Nenhum boletim encontrado!  ***"
        ElseIf contador = 1 Then
            linha = "Resultado: 1 boletim encontrado!  ***"
        ElseIf contador > 1 Then
            linha = String.Format("Resultado: {0} boletins encontrados!  ***", contador)
        End If
        txtImportacao.AppendText(linha & vbCrLf)

    End Sub

    Private Sub prTransfere_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

        If EnviadoFTP = True And EnviadoFINAL = 2 Then
            MsgBox("Notas enviadas com sucesso!", MsgBoxStyle.Information, "Sucesso!")
        ElseIf EnviadoFTP = False Then
            MsgBox("Notas não enviadas!", MsgBoxStyle.Information, "Mensagem de erro")
        End If

    End Sub

    Private Sub tbSN_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbSN.LostFocus
        If tbSN.Text = "" Then
            MsgBox("Digite a recuperação paralela!", MsgBoxStyle.Information)
            tbSN.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub tbFaltas2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbFaltas2.LostFocus
        If tbFaltas2.Text = "" Then
            MsgBox("Digite a falta!", MsgBoxStyle.Information)
            tbFaltas2.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub tbAC2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbAC2.LostFocus

        If tbFaltas2.Text <> "" Then
            If tbAC2.Text = "" Then
                MsgBox("Digite a compensação de aulas!", MsgBoxStyle.Information)
                tbAC2.Focus()
                Exit Sub
            End If
        End If

    End Sub

    Private Sub cbTurmas2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbTurmas2.TextChanged

        If (My.Computer.FileSystem.FileExists(tAno2.Value & "\professor.txt") = True) Then
            cbDisciplinas.Text = ""
            cbDisciplinas2.Text = ""
            cbDisciplinas3.Text = ""
            CarregaDisciplina(cbTurmas2.Text)

        End If

        AjustaJanela(False)

        dgBoletim.Visible = False
        exibeBoletim = 0
        Array.Clear(rBL, 0, 999)
        Array.Clear(rAV, 0, 999)
        Array.Clear(rFT, 0, 999)
        Array.Clear(rAC, 0, 999)
        Array.Clear(rPR, 0, 999)

        nAluno.Value = 1
        tbMedia2.Text = ""
        tbFaltas2.Text = ""
        tbAC2.Text = ""
        tbSN2.Text = ""
        nroEvasaoTotal.Text = "00"

        tbMedia2.Enabled = False
        tbFaltas2.Enabled = False
        tbAC2.Enabled = False
        tbSN2.Enabled = False
        btClasse.Enabled = False
        btAlterar.Enabled = False
        dgBoletim.Visible = False

        If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\{3}.txt", tAno2.Value, cbTurmas2.Text, cbDisciplinas2.Text, tBimestre2.Value))) Then
            tbStatus2.Text = "Boletim encontrado!"
            tbStatus2.ForeColor = Color.Blue
            btConsultar.Enabled = True
            tbMedia.Text = ""
            tbFaltas.Text = "0"
            tbAC.Text = "0"
            tbSN.Text = "N"

            Exit Sub
        Else
            tbStatus2.Text = "Boletim não encontrado!"
            tbStatus2.ForeColor = Color.Red
            btConsultar.Enabled = False
            tbMedia.Text = ""
            tbFaltas.Text = "0"
            tbAC.Text = "0"
            tbSN.Text = "N"

        End If

    End Sub

    Private Sub tabProfessor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabProfessor.Click
        AjustaJanela(False)

        dgBoletim.Visible = False
        exibeBoletim = 0
    End Sub

    Private Sub cbDisciplinas_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbDisciplinas.TextChanged

        If cbTurmas.Text <> "" And cbDisciplinas.Text <> "" Then

            Array.Clear(rST, 0, 100)

            Array.Clear(AV, 0, 100)
            Array.Clear(FT, 0, 100)
            Array.Clear(AC, 0, 100)
            Array.Clear(PR, 0, 100)

            riST = 1
            nroaluno = 1

            AjustaJanela(False)

            dgBoletim.Visible = False
            exibeBoletim = 0

            ChecaBoletim()
            BoletimAvancado()

            If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\{3}.txt", tAno.Value, cbTurmas.Text, cbDisciplinas.Text, testeAF))) Then
                btCadastrar.Enabled = False
                tbStatus.Text = "Boletim já cadastrado!"
                tbFaltas.Text = "0"
                tbAC.Text = "0"
                tbSN.Text = "N"
                tbMedia.Enabled = False
                tbFaltas.Enabled = False
                tbAC.Enabled = False
                tbSN.Enabled = False
                btFinalizar.Enabled = False

                travaBoletim = 1
                Exit Sub
            End If

        End If

    End Sub

    Private Sub cbDisciplinas2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbDisciplinas2.TextChanged

        AjustaJanela(False)

        dgBoletim.Visible = False
        exibeBoletim = 0

        Array.Clear(rBL, 0, 999)
        Array.Clear(rAV, 0, 999)
        Array.Clear(rFT, 0, 999)
        Array.Clear(rAC, 0, 999)
        Array.Clear(rPR, 0, 999)

        nAluno.Value = 1
        tbMedia2.Text = ""
        tbFaltas2.Text = ""
        tbAC2.Text = ""
        tbSN2.Text = ""
        nroEvasaoTotal.Text = "00"

        tbMedia2.Enabled = False
        tbFaltas2.Enabled = False
        tbAC2.Enabled = False
        tbSN2.Enabled = False
        dgBoletim.Visible = False

        If (My.Computer.FileSystem.FileExists("" & tAno2.Value & "\" & cbTurmas2.Text & "\" & cbDisciplinas2.Text & "\" & tBimestre2.Value & ".txt")) Then
            tbStatus2.Text = "Boletim encontrado!"
            tbStatus2.ForeColor = Color.Blue
            btConsultar.Enabled = True
            tbMedia.Text = ""
            tbFaltas.Text = "0"
            tbAC.Text = "0"
            tbSN.Text = "N"

            Exit Sub
        Else
            tbStatus2.Text = "Boletim não encontrado!"
            tbStatus2.ForeColor = Color.Red
            btConsultar.Enabled = False
            tbMedia.Text = ""
            tbFaltas.Text = "0"
            tbAC.Text = "0"
            tbSN.Text = "N"

        End If
    End Sub

    Private Sub cbTurmas3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbTurmas3.KeyPress
        e.Handled = True

    End Sub

    Private Sub cbTurmas3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbTurmas3.TextChanged

        If (My.Computer.FileSystem.FileExists(tAno3.Value & "\professor.txt") = True) Then
            cbDisciplinas.Text = ""
            cbDisciplinas2.Text = ""
            cbDisciplinas3.Text = ""
            CarregaDisciplina(cbTurmas3.Text)
        End If

        AjustaJanela(False)

        dgBoletim.Visible = False
        exibeBoletim = 0

        If (My.Computer.FileSystem.FileExists(String.Format("{0}\{1}\{2}\{3}.txt", tAno3.Value, cbTurmas3.Text, cbDisciplinas3.Text, tBimestre3.Value))) Then
            tbStatus3.Text = "Boletim encontrado!"
            tbStatus3.ForeColor = Color.Blue
            btExcluir.Enabled = True
            Exit Sub
        Else
            tbStatus3.Text = "Boletim não encontrado!"
            tbStatus3.ForeColor = Color.Red
            btExcluir.Enabled = False
        End If

    End Sub

    Private Sub tbMedia2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbMedia2.LostFocus

        Try

            Dim nota As Integer
            nota = CInt(tbMedia2.Text)

            If nota < 5 Then
                tbMedia2.ForeColor = Color.Red
            Else
                tbMedia2.ForeColor = Color.Blue
            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub cbTurmas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbTurmas.SelectedIndexChanged

    End Sub

    Private Sub cbTurmas2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbTurmas2.SelectedIndexChanged

    End Sub

    Private Sub cbTurmas3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbTurmas3.SelectedIndexChanged

    End Sub

    Private Sub cbDisciplinas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbDisciplinas.SelectedIndexChanged

    End Sub
End Class
