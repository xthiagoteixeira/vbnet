Imports DevExpress.Spreadsheet
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraRichEdit
Imports Mais_Escola.Helpers
Imports Microsoft.VisualBasic.FileIO
Imports System.ComponentModel
Imports System.IO

Public Class frmImpExcel

    Dim IndiceNota = 0, IndiceFalta = 0
    Dim MeuBimestre As String
    Dim MeuBimestre2 As String

    Dim contadorArquivos = 0
    Dim nroAluno = 1

    ReadOnly arrayArquivos(99) As String
    ReadOnly arrayArquivosTipo(99) As String

    Private Sub AtualizarBarra_Invoke(value As Integer, contadorArquivos As Integer)
        BeginInvoke(New MethodInvoker(Function() AtualizarBarra(value, contadorArquivos)))
    End Sub

    Private Function AtualizarBarra(value As Integer, contadorArquivos As Integer) As Boolean

        barra.Properties.Maximum = contadorArquivos
        barra.Text = value

        barra.PerformStep()
        barra.Update()

        Return True
    End Function

    Private Sub DefinirProgresso(Status As String, Turma As String, Disciplina As String, texto As String)
        BeginInvoke(New MethodInvoker(Function() AtualizarDG(Status, Turma, Disciplina, texto)))
    End Sub

    Private Function AtualizarDG(Status As String, Turma As String, Disciplina As String, texto As String) As Boolean

        viewBoletim.AddNewRow()

        Dim rowHandle = viewBoletim.GetRowHandle(viewBoletim.DataRowCount)

        If (viewBoletim.IsNewItemRow(rowHandle)) Then

            If Status = "S" Then


                viewBoletim.SetRowCellValue(rowHandle, viewBoletim.Columns(0), ImageList.Images(1))
                viewBoletim.SetRowCellValue(rowHandle, viewBoletim.Columns(1), Turma)
                viewBoletim.SetRowCellValue(rowHandle, viewBoletim.Columns(2), Disciplina)
                viewBoletim.SetRowCellValue(rowHandle, viewBoletim.Columns(3), texto)

                'viewBoletim.AddNewRow({ImageList.Images(1), Turma, Disciplina, texto})
            Else

                viewBoletim.SetRowCellValue(rowHandle, viewBoletim.Columns(0), ImageList.Images(0))
                viewBoletim.SetRowCellValue(rowHandle, viewBoletim.Columns(1), Turma)
                viewBoletim.SetRowCellValue(rowHandle, viewBoletim.Columns(2), Disciplina)
                viewBoletim.SetRowCellValue(rowHandle, viewBoletim.Columns(3), texto)

                'DGBoletim.Rows.Add({ImageList.Images(0), Turma, Disciplina, texto})
            End If

        End If

        Return True
    End Function

    Sub Carregar_Arquivos()

        contadorArquivos = 0
        Array.Clear(arrayArquivos, 0, 99)

        Dim nomearquivo As String = Path.GetFileName(tbArquivo.Text)
        Dim importacao_alunos = tbArquivo.Text
        importacao_alunos = importacao_alunos.Replace(nomearquivo, String.Empty)
        importacao_alunos = Trim(importacao_alunos)

        If cbTodosArquivos.Checked = True Then

            If nomearquivo.Contains(".xlsx") = True Then

                ' /// TODOS ARQUIVOS DA PASTA...
                For Each nomearquivo In Directory.GetFiles(importacao_alunos, "*.xlsx")

                    Dim caminhoArquivoTxt As String = nomearquivo.Replace(".xlsx", ".txt")
                    Using workbook As New Workbook
                        workbook.LoadDocument(nomearquivo, DevExpress.Spreadsheet.DocumentFormat.OpenXml)
                        workbook.SaveDocument(caminhoArquivoTxt, DevExpress.Spreadsheet.DocumentFormat.Csv)
                    End Using
                    arrayArquivos(contadorArquivos) = caminhoArquivoTxt
                    arrayArquivosTipo(contadorArquivos) = "Mapão XLS"

                    ' Verifica se está em HTML
                    Dim fi As New FileInfo(caminhoArquivoTxt)
                    ' \\\ OBTER TAMANHO DO ARQUIVO
                    If fi.Length() = "0" Then
                        Dim server As RichEditDocumentServer = New RichEditDocumentServer
                        server.LoadDocument(nomearquivo, DevExpress.XtraRichEdit.DocumentFormat.Html)
                        server.SaveDocument(caminhoArquivoTxt, DevExpress.XtraRichEdit.DocumentFormat.PlainText)
                        arrayArquivosTipo(contadorArquivos) = "Mapão HTML"

                    End If
                    contadorArquivos += 1
                Next

            ElseIf nomearquivo.Contains(".xltx") = True Then

                ' /// TODOS ARQUIVOS DA PASTA...
                For Each nomearquivo In Directory.GetFiles(importacao_alunos, "*.xltx")

                    Dim caminhoArquivoTxt As String = nomearquivo.Replace(".xltx", ".txt")
                    Using workbook As New Workbook
                        workbook.LoadDocument(nomearquivo, DevExpress.Spreadsheet.DocumentFormat.OpenXml)
                        workbook.SaveDocument(caminhoArquivoTxt, DevExpress.Spreadsheet.DocumentFormat.Csv)
                    End Using
                    arrayArquivos(contadorArquivos) = caminhoArquivoTxt
                    arrayArquivosTipo(contadorArquivos) = "Mapão XLS"

                    ' Verifica se está em HTML
                    Dim fi As New FileInfo(caminhoArquivoTxt)
                    ' \\\ OBTER TAMANHO DO ARQUIVO
                    If fi.Length() = "0" Then
                        Dim server As RichEditDocumentServer = New RichEditDocumentServer
                        server.LoadDocument(nomearquivo, DevExpress.XtraRichEdit.DocumentFormat.Html)
                        server.SaveDocument(caminhoArquivoTxt, DevExpress.XtraRichEdit.DocumentFormat.PlainText)
                        arrayArquivosTipo(contadorArquivos) = "Mapão HTML"

                    End If
                    contadorArquivos += 1
                Next

            ElseIf nomearquivo.Contains(".xls") = True Then

                ' /// TODOS ARQUIVOS DA PASTA...
                For Each nomearquivo In Directory.GetFiles(importacao_alunos, "*.xls")
                    Dim caminhoArquivoTxt As String = nomearquivo.Replace(".xls", ".txt")
                    Using workbook As New Workbook
                        workbook.LoadDocument(nomearquivo, DevExpress.Spreadsheet.DocumentFormat.Xls)
                        workbook.SaveDocument(caminhoArquivoTxt, DevExpress.Spreadsheet.DocumentFormat.Csv)
                    End Using
                    arrayArquivos(contadorArquivos) = caminhoArquivoTxt
                    arrayArquivosTipo(contadorArquivos) = "Mapão XLS"

                    ' Verifica se está em HTML
                    Dim fi As New FileInfo(caminhoArquivoTxt)
                    ' \\\ OBTER TAMANHO DO ARQUIVO
                    If fi.Length() = "0" Then
                        Dim server As RichEditDocumentServer = New RichEditDocumentServer
                        server.LoadDocument(nomearquivo, DevExpress.XtraRichEdit.DocumentFormat.Html)
                        server.SaveDocument(caminhoArquivoTxt, DevExpress.XtraRichEdit.DocumentFormat.PlainText)
                        arrayArquivosTipo(contadorArquivos) = "Mapão HTML"

                    End If
                    contadorArquivos += 1
                Next

            End If

        Else

            ' ******************... SOMENTE UM ARQUIVO...*************************
            'If nomearquivo.Contains(".pdf") = True Then

            '    arrayArquivos(contadorArquivos) = tbArquivo.Text
            '    arrayArquivosTipo(contadorArquivos) = "Mapão PDF"

            If nomearquivo.Contains(".xltx") = True Then

                '... SOMENTE UM ARQUIVO...
                Dim caminhoArquivoTxt As String = tbArquivo.Text.Replace(".xltx", ".txt")
                Using workbook As New Workbook
                    workbook.LoadDocument(tbArquivo.Text, DevExpress.Spreadsheet.DocumentFormat.OpenXml)
                    workbook.SaveDocument(caminhoArquivoTxt, DevExpress.Spreadsheet.DocumentFormat.Csv)
                End Using
                arrayArquivos(contadorArquivos) = caminhoArquivoTxt
                arrayArquivosTipo(contadorArquivos) = "Mapão XLS"

                ' Verifica se está em HTML
                '
                Dim fi As New FileInfo(caminhoArquivoTxt)
                ' \\\ OBTER TAMANHO DO ARQUIVO
                If fi.Length() = "0" Then
                    Dim server As RichEditDocumentServer = New RichEditDocumentServer
                    server.LoadDocument(tbArquivo.Text, DevExpress.XtraRichEdit.DocumentFormat.Html)
                    server.SaveDocument(caminhoArquivoTxt, DevExpress.XtraRichEdit.DocumentFormat.PlainText)
                    arrayArquivosTipo(contadorArquivos) = "Mapão HTML"
                End If

            ElseIf nomearquivo.Contains(".xlsx") = True Then

                '... SOMENTE UM ARQUIVO...
                Dim caminhoArquivoTxt As String = tbArquivo.Text.Replace(".xlsx", ".txt")
                Using workbook As New Workbook
                    workbook.LoadDocument(tbArquivo.Text, DevExpress.Spreadsheet.DocumentFormat.OpenXml)
                    workbook.SaveDocument(caminhoArquivoTxt, DevExpress.Spreadsheet.DocumentFormat.Csv)
                End Using
                arrayArquivos(contadorArquivos) = caminhoArquivoTxt
                arrayArquivosTipo(contadorArquivos) = "Mapão XLS"

                ' Verifica se está em HTML
                '
                Dim fi As New FileInfo(caminhoArquivoTxt)
                ' \\\ OBTER TAMANHO DO ARQUIVO
                If fi.Length() = "0" Then
                    Dim server As RichEditDocumentServer = New RichEditDocumentServer
                    server.LoadDocument(tbArquivo.Text, DevExpress.XtraRichEdit.DocumentFormat.Html)
                    server.SaveDocument(caminhoArquivoTxt, DevExpress.XtraRichEdit.DocumentFormat.PlainText)
                    'arrayArquivosTipo(contadorArquivos) = "Mapão HTML"
                    arrayArquivosTipo(contadorArquivos) = "Mapão HTML"
                End If

            ElseIf nomearquivo.Contains(".xls") = True Then

                '... SOMENTE UM ARQUIVO...
                Dim caminhoArquivoTxt As String = tbArquivo.Text.Replace(".xls", ".txt")
                Using workbook As New Workbook
                    workbook.LoadDocument(tbArquivo.Text, DevExpress.Spreadsheet.DocumentFormat.Xls)
                    workbook.SaveDocument(caminhoArquivoTxt, DevExpress.Spreadsheet.DocumentFormat.Csv)
                End Using
                arrayArquivos(contadorArquivos) = caminhoArquivoTxt
                arrayArquivosTipo(contadorArquivos) = "Mapão XLS"

                ' Verifica se está em HTML
                Dim fi As New FileInfo(caminhoArquivoTxt)
                ' \\\ OBTER TAMANHO DO ARQUIVO
                If fi.Length() = "0" Then
                    Dim server As RichEditDocumentServer = New RichEditDocumentServer
                    server.LoadDocument(tbArquivo.Text, DevExpress.XtraRichEdit.DocumentFormat.Html)
                    server.SaveDocument(caminhoArquivoTxt, DevExpress.XtraRichEdit.DocumentFormat.PlainText)
                    'arrayArquivosTipo(contadorArquivos) = "Mapão HTML"
                    arrayArquivosTipo(contadorArquivos) = "Mapão HTML"

                End If

            End If

        End If

    End Sub

    Private Sub cbEscolhaBimestre_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbEscolhaBimestre.SelectedIndexChanged

        If cbEscolhaBimestre.Text = "1º Bimestre" Then
            MeuBimestre = "1"
            IndiceNota = 1
            IndiceFalta = 2

        ElseIf cbEscolhaBimestre.Text = "2º Bimestre" Then
            MeuBimestre = "2"
            IndiceNota = 3
            IndiceFalta = 4

        ElseIf cbEscolhaBimestre.Text = "3º Bimestre" Then
            MeuBimestre = "3"
            IndiceNota = 5
            IndiceFalta = 6

        ElseIf cbEscolhaBimestre.Text = "4º Bimestre" Then
            MeuBimestre = "4"
            IndiceNota = 7
            IndiceFalta = 8

        ElseIf cbEscolhaBimestre.Text = "2º Avaliação Final" Then
            MeuBimestre = "FINAL"
            IndiceNota = 9
            IndiceFalta = 10

        ElseIf cbEscolhaBimestre.Text = "4º Avaliação Final" Then
            MeuBimestre = "FINAL"
            IndiceNota = 9
            IndiceFalta = 10

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles tbArquivo.TextChanged
        btCarregar.Enabled = True
    End Sub

    Private Sub btCarregar_Click(sender As Object, e As EventArgs) Handles btCarregar.Click

        ' ... Checa bimestre escolhido ...
        If cbEscolhaBimestre.Text <> String.Empty And tbArquivo.Text <> String.Empty Then
            btCarregar.Enabled = False
            Carregar_Arquivos()

            bwCarregar.RunWorkerAsync()
        Else
            MsgBox("Escolher o bimestre!", MsgBoxStyle.Information, "Importação de Boletins")
        End If

    End Sub

    Private Sub frmImportacaoBoletins_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Erro_Form = "frmImpBoletins"
        tAno.Value = AnoVigente

        Dim dh As New DataHelper(DSparametr.simpleDS)
        gridBoletim.DataSource = dh.DataSet
        gridBoletim.DataMember = dh.DataMember

        viewBoletim.Columns(0).Width = 20
        viewBoletim.Columns(1).Width = 50
        viewBoletim.Columns(2).Width = 70
        viewBoletim.Columns(3).Width = 100


    End Sub

    Private Sub bwCarregar_DoWork(sender As Object, e As DoWorkEventArgs) Handles bwCarregar.DoWork

        ' Try
        ' ... Carrega os arquivos ...
        ' Ler os arquivos ... (TXT)
        Dim contador = 0
        Dim contaTracos = 0

        While (arrayArquivos(contador) IsNot Nothing)

            Dim arrayBoletim_M(500, 500) As String
            Dim arrayBoletim_F(500, 500) As String
            Dim arrayBoletim_AC(500, 500) As String
            Dim arrayBoletim_Aulas(500) As String

            Dim arrayNF1(500), arrayNF2(500), arrayNF3(500), arrayNF4(500)
            Dim arrayCheck_M(500), arrayCheck_F(500)

            Dim iDisciplina = 0
            Dim Nro_Turma = 0
            Dim NomeTurma = String.Empty

            '////////////////////////////////// MAPAO XLS ////////////////
            If arrayArquivosTipo(contador) = "Mapão XLS" Then

                '      Using leitor As New TextFieldParser(arrayArquivos(contador).ToString, System.Text.Encoding.UTF7)

                Using leitor As New TextFieldParser(arrayArquivos(contador).ToString, System.Text.Encoding.UTF7)

                    Dim linhaAtual As String()
                    Dim nroAluno_Turma = 1

                    Dim Nro_Disciplina = 0
                    'Informamos que será importado com Delimitação
                    leitor.TextFieldType = FieldType.Delimited
                    'Informamos o Delimitador
                    leitor.SetDelimiters(",")
                    leitor.TrimWhiteSpace = False

                    linhaAtual = leitor.ReadFields()

                    ' /////////// CONTEUDO DESTA TURMA ARQUIVO ...
                    While Not leitor.EndOfData

                        If NomeTurma = String.Empty Then ' /////////////////////////// Nome da Turma
                            While (NomeTurma = String.Empty)
                                Try
                                    linhaAtual = leitor.ReadFields()
                                    If linhaAtual(0).Contains("Turma:") = True Then
                                        NomeTurma = String.Format("{0}", linhaAtual(0).ToString)
                                        NomeTurma = NomeTurma.Replace("Turma:", String.Empty)
                                        NomeTurma = NomeTurma.Replace("MANHA", String.Empty)
                                        NomeTurma = NomeTurma.Replace("TARDE", String.Empty)
                                        NomeTurma = NomeTurma.Replace("NOITE", String.Empty)
                                        NomeTurma = NomeTurma.Replace("INTEGRAL", String.Empty)
                                        NomeTurma = NomeTurma.Replace("  ", String.Empty)

                                    End If
                                Catch ex As Exception
                                End Try

                            End While

                            NomeTurma = NomeTurma.Trim(NomeTurma)
                            NomeTurma = NomeTurma.TrimStart(NomeTurma)
                            NomeTurma = NomeTurma.TrimEnd(NomeTurma)

                            ' ///////////// BIMESTRE...
                            While (linhaAtual(0).Contains("Tipo de Fechamento") = False)
                                linhaAtual = leitor.ReadFields()
                            End While

                            Dim BimestreArquivo = linhaAtual(0).ToString
                            'If BimestreArquivo = linhaAtual(4).ToString = "" Then
                            '    BimestreArquivo = linhaAtual(3).ToString
                            'End If
                            If BimestreArquivo.Contains("PRIMEIRO") = True Or BimestreArquivo.Contains("Primeiro") = True Then
                                BimestreArquivo = "1"
                            ElseIf BimestreArquivo.Contains("SEGUNDO") = True Or BimestreArquivo.Contains("Segundo") = True Then
                                BimestreArquivo = "2"
                            ElseIf BimestreArquivo.Contains("TERCEIRO") = True Or BimestreArquivo.Contains("Terceiro") = True Then
                                BimestreArquivo = "3"
                            ElseIf BimestreArquivo.Contains("QUARTO") = True Or BimestreArquivo.Contains("Quarto") = True Then
                                BimestreArquivo = "4"
                            ElseIf BimestreArquivo.Contains("FINAL") = True Or BimestreArquivo.Contains("Final") = True Then
                                BimestreArquivo = "FINAL"
                            End If

                            If (MeuBimestre <> BimestreArquivo) Then
                                DefinirProgresso("F", NomeTurma, "-", "Bimestre não corresponde com o arquivo.")
                                Exit While
                            ElseIf BimestreArquivo = "FINAL" Then

                                If cbEscolhaBimestre.EditValue = "2º Avaliação Final" Then
                                    MeuBimestre2 = "2AF"
                                ElseIf cbEscolhaBimestre.EditValue = "4º Avaliação Final" Then
                                    MeuBimestre2 = "4AF"
                                End If
                            Else
                                MeuBimestre2 = BimestreArquivo
                            End If

                            linhaAtual = leitor.ReadFields()
                            If NomeTurma.Contains("MULTISSERIADA") = True Then
                                While (linhaAtual(1).Contains("Tipo") = False)
                                    linhaAtual = leitor.ReadFields()
                                    If linhaAtual(1).Contains("N1") = True Then
                                        NomeTurma = String.Format("{0}N1", NomeTurma)
                                        Exit While
                                    ElseIf linhaAtual(1).Contains("N2") = True Then
                                        NomeTurma = String.Format("{0}N2", NomeTurma)
                                        Exit While
                                    ElseIf linhaAtual(1).Contains("N3") = True Then
                                        NomeTurma = String.Format("{0}N3", NomeTurma)
                                    End If
                                End While
                            End If

                            '/////////////// TURMA do ALUNO ...
                            SQL = String.Format(
                                            "SELECT codigo_trma, classe FROM turma WHERE bloqueado='0' AND classe LIKE '%{0}%';",
                                            NomeTurma)
                            Nro_Turma = MySQL_consulta_campo(SQL, "codigo_trma")
                            If Nro_Turma = "0" Then
                                '... INFORMA o nome da TURMA...
                                DefinirProgresso("F", NomeTurma, "***", "Turma não cadastrada.")
                                '... SAI DESTE ARQUIVO...
                                Exit While
                            Else
                                DefinirProgresso("S", NomeTurma, "***", "Turma cadastrada.")
                            End If

                            ' ... LINHA das Disciplinas...
                            'While (linhaAtual(1).Contains("ARTE") = False)
                            '    linhaAtual = leitor.ReadFields()
                            'End While

                            ' ... LINHA das Disciplinas...
                            While (linhaAtual(0).Contains("GOE") = False)
                                linhaAtual = leitor.ReadFields()
                            End While
                            linhaAtual = leitor.ReadFields()
                            '       linhaAtual = leitor.ReadFields()

                        End If

                        '... Demais disciplinas...
                        While Not (IsNumeric(linhaAtual(0).ToString))

                            If linhaAtual(0).ToString <> String.Empty Then

                                Dim ProcuraDisciplina = linhaAtual(0).ToString

                                ' DISCIPLINAS DIFERENTES NOMES                                    
                                If ProcuraDisciplina = "LINGUAESTRANGEIRAINGLES" Or ProcuraDisciplina = "LINGUA ESTRANGEIRA INGLES" Then
                                    ProcuraDisciplina = "INGLES"
                                    'ElseIf ProcuraDisciplina = "LINGUAPORTUGUESAELITERATURA" Or ProcuraDisciplina = "LINGUA PORTUGUESA E LITERATURA" Then
                                    '    ProcuraDisciplina = "LITERATURA"
                                ElseIf ProcuraDisciplina = "EDUCACAOFISICA" Then
                                    ProcuraDisciplina = "EDUCACAO FISICA"
                                ElseIf ProcuraDisciplina = "LINGUAPORTUGUESA" Then
                                    ProcuraDisciplina = "LINGUA PORTUGUESA"
                                End If

                                contaTracos = 0
                                Dim Nro_NotasFreq = 0
                                ' ... Checa as disciplinas do ALUNO ...                                    
                                SQL = String.Format("SELECT codigo_disc, disciplina FROM disciplinas WHERE bloqueado='0' AND disciplina='{0}';", ProcuraDisciplina)
                                Nro_Disciplina = MySQL_consulta_campo(SQL, "codigo_disc").ToString
                                If Nro_Disciplina = "0" Then
                                    DefinirProgresso("F", NomeTurma, ProcuraDisciplina, "Disciplina não cadastrada.")
                                    contaTracos = 120
                                Else
                                    ' Consulta de Existe dentro do sistema o CONSOLIDADO...
                                    SQL = String.Format("SELECT cod_nf FROM notasfreq WHERE anovigente='{0}' AND cod_bimestre='{1}' AND turma='{2}' AND disciplina='{3}';",
                                                                tAno.Value, MeuBimestre2, Nro_Turma, Nro_Disciplina)
                                    Nro_NotasFreq = MySQL_consulta_campo(SQL, "cod_nf")
                                    contaTracos = 0
                                End If

                                ' ..... INDICES .....
                                ' Nome da Disciplina
                                arrayNF1(iDisciplina) = Trim(ProcuraDisciplina)
                                ' Número da Disciplina
                                arrayNF2(iDisciplina) = Nro_Disciplina
                                ' Número do Boletim
                                arrayNF3(iDisciplina) = Nro_NotasFreq

                                iDisciplina += 1

                            End If
                            ' Próxima linha...
                            linhaAtual = leitor.ReadFields()
                        End While

                        '... Procura o primeiro aluno ...                            
                        While (linhaAtual(0).ToString <> "01")
                            linhaAtual = leitor.ReadFields()
                        End While

                        Dim iDisciplina2 = 0

                        '... Preenche os boletins na variável ...
                        While (linhaAtual(0).ToString.Contains("Aulas") = False)

                            nroAluno = linhaAtual(0).ToString
                            Dim lerColuna = 1

                            Try
                                ' Percorre o Aluno
                                For x = 0 To iDisciplina * 5 Step 1

                                    If linhaAtual(x + lerColuna).ToString <> String.Empty Then

                                        ' *** NOTAS ***
                                        lerColuna += 1
                                        nroAluno_Turma = nroAluno

                                        If ((linhaAtual(x + lerColuna).ToString) = "-") Then
                                            'Conta Traços
                                            arrayCheck_M(iDisciplina2) += 1
                                            'arrayBoletim_M(iDisciplina2, nroAluno) = 19
                                            SQL = String.Format("SELECT status FROM aluno WHERE anovigente='{0}' AND turma='{1}' AND nro='{2}';",
                                                   tAno.Value, Nro_Turma, nroAluno)
                                            arrayBoletim_M(iDisciplina2, nroAluno) = TesteEvasaoEscolar_Nome((MySQL_consulta_campo(SQL, "status").ToString))
                                            'arrayBoletim_M(iDisciplina, nroAluno) = EvasaoEscolar(tAno.Text, Nro_Turma, nroAluno)
                                        Else
                                            arrayBoletim_M(iDisciplina2, nroAluno) = linhaAtual(x + lerColuna).ToString.Replace(",", ".")
                                        End If

                                        If cbEscolhaBimestre.Text.Contains("Final") = False Then

                                            lerColuna += 1
                                            ' VERIFICA SE EXISTE FALTAS ...
                                            If ((linhaAtual(x + lerColuna).ToString) = "-") Then
                                                arrayBoletim_F(iDisciplina2, nroAluno) = 0
                                            Else
                                                arrayBoletim_F(iDisciplina2, nroAluno) = linhaAtual(x + lerColuna).ToString.Replace(",", ".")
                                            End If

                                            lerColuna += 1
                                            ' VERIFICA SE EXISTE AC ...
                                            If ((linhaAtual(x + lerColuna).ToString) = "-") Then
                                                arrayBoletim_AC(iDisciplina2, nroAluno) = 0
                                            Else
                                                arrayBoletim_AC(iDisciplina2, nroAluno) = linhaAtual(x + lerColuna).ToString.Replace(",", ".")
                                            End If

                                        End If
                                        lerColuna += 1
                                        iDisciplina2 += 1

                                    End If
                                Next
                            Catch ex As Exception
                            End Try

                            ' ... Próximo Aluno
                            iDisciplina2 = 0
                            linhaAtual = leitor.ReadFields()

                        End While

                        Dim t = 0
                        ' *** AULAS DADAS ***
                        For x = 0 To iDisciplina * 5 Step 1
                            If (linhaAtual(x).ToString.Contains("Aulas") = True) Then
                                Dim temp As String = linhaAtual(x).ToString.Replace("Aulas Dadas:", String.Empty)
                                arrayBoletim_Aulas(t) = Trim(temp)
                                t += 1
                            End If
                        Next
                        Exit While

                        ' /////////// FIM DO CONTEUDO DESTA TURMA ARQUIVO ...
                    End While

                    ' ANALISA...
                    If Nro_Turma <> 0 Then
                        '... Procura boletins EXISTENTES ...
                        For i = 0 To iDisciplina - 1

                            'If (arrayCheck_M(i) = arrayCheck_F(i)) Or (arrayNF2(i) = "0") Then
                            If (arrayCheck_M(i) = nroAluno_Turma And arrayCheck_F(i) = nroAluno_Turma) Or (arrayNF2(i) = "0") Then
                                ' NAO EXISTE BOLETIM ...
                                DefinirProgresso("F", NomeTurma, arrayNF1(i).ToString, "Boletim/Disciplina não encontrada.")
                            Else
                                ' EXISTE BOLETIM ...
                                ' DefinirProgresso("S", NomeTurma, arrayNF1(i).ToString, "Boletim pronto para carregar!")
                                If (arrayNF3(i) <> "0") Then

                                    If cbAtualizarNotas.Checked = False Then
                                        DefinirProgresso("F", NomeTurma, arrayNF1(i).ToString, "Boletim já existe!")

                                    Else
                                        ' Delete boletim / notasfreq
                                        ' 
                                        SQL = "DELETE FROM boletim WHERE cod_boletim='" & arrayNF3(i).ToString & "';"
                                        MySQL_atualiza(SQL)
                                        SQL = "DELETE FROM notasfreq WHERE cod_nf='" & arrayNF3(i).ToString & "';"
                                        MySQL_atualiza(SQL)


                                        ' /// INSERIR BOLETIM ...
                                        SQL = String.Format("INSERT INTO notasfreq (turma, disciplina, cod_bimestre, qtdadeaulas, previsaoaulas, anovigente, dt_criacao,dt_atualizacao) values('{0}', '{1}', '{2}', '{5}', '{5}', '{3}', '{4}', '{4}'); SELECT LAST_INSERT_ID() AS cod_nf;",
                                                      Nro_Turma, arrayNF2(i).ToString, MeuBimestre2, tAno.Value, Format(Date.Now, "yyyy-MM-dd HH:mm:ss").ToString, arrayBoletim_Aulas(i).ToString)
                                        Dim nroBoletim = MySQL_atualiza(SQL)

                                        ' ... /// INSERIR NOTAS / TURMA / DISCIPLINA...
                                        nroAluno = 1
                                        While (arrayBoletim_M(i, nroAluno) <> Nothing)

                                            If cbEscolhaBimestre.Text.Contains("Final") = False Then
                                                SQL = String.Format("INSERT INTO boletim (cod_boletim, nro_aluno, M, F, AC, S, porcentagem) values('{0}', '{1}', '{2}', '{3}', '{5}', 'N', '{4}');",
                                                                                                        nroBoletim, nroAluno, arrayBoletim_M(i, nroAluno).ToString,
                                                                                                        arrayBoletim_F(i, nroAluno).ToString,
                                                                                                        Resultado_Porcentagem(arrayBoletim_F(i, nroAluno).ToString,
                                                                                                                              arrayBoletim_AC(i, nroAluno).ToString,
                                                                                                                              arrayBoletim_Aulas(i).ToString),
                                                                                                        arrayBoletim_AC(i, nroAluno).ToString)
                                            Else
                                                SQL = String.Format(
                                                            "INSERT INTO boletim (cod_boletim, nro_aluno, M, F, AC, S, porcentagem) values('{0}', '{1}', '{2}', '0', '0', '1', '0');",
                                                            nroBoletim, nroAluno, arrayBoletim_M(i, nroAluno).ToString)
                                            End If
                                            MySQL_atualiza(SQL)
                                            nroAluno += 1

                                        End While

                                        DefinirProgresso("S", NomeTurma, arrayNF1(i).ToString, "Atualizado com sucesso!")
                                        arquivoLog("Boletim",
                                                               String.Format("{0} - {1} Atualizado com sucesso!", NomeTurma, arrayNF1(i)))

                                    End If

                                Else
                                    ' DefinirProgresso("S", NomeTurma, arrayNF1(i).ToString, "Boletim pronto para carregar!")

                                    ' /// INSERIR BOLETIM ...
                                    SQL = String.Format("INSERT INTO notasfreq (turma, disciplina, cod_bimestre, qtdadeaulas, previsaoaulas, anovigente, dt_criacao,dt_atualizacao) values('{0}', '{1}', '{2}', '{5}', '{5}', '{3}', '{4}', '{4}'); SELECT LAST_INSERT_ID() AS cod_nf;",
                                                  Nro_Turma, arrayNF2(i).ToString, MeuBimestre2, tAno.Value, Format(Date.Now, "yyyy-MM-dd HH:mm:ss").ToString, arrayBoletim_Aulas(i).ToString)
                                    Dim nroBoletim = MySQL_atualiza(SQL)

                                    ' ... /// INSERIR NOTAS / TURMA / DISCIPLINA...
                                    nroAluno = 1
                                    While (arrayBoletim_M(i, nroAluno) <> Nothing)

                                        If cbEscolhaBimestre.Text.Contains("Final") = False Then
                                            SQL = String.Format("INSERT INTO boletim (cod_boletim, nro_aluno, M, F, AC, S, porcentagem) values('{0}', '{1}', '{2}', '{3}', '{5}', 'N', '{4}');",
                                                                                                    nroBoletim, nroAluno, arrayBoletim_M(i, nroAluno).ToString,
                                                                                                    arrayBoletim_F(i, nroAluno).ToString,
                                                                                                    Resultado_Porcentagem(arrayBoletim_F(i, nroAluno).ToString,
                                                                                                                          arrayBoletim_AC(i, nroAluno).ToString,
                                                                                                                          arrayBoletim_Aulas(i).ToString),
                                                                                                    arrayBoletim_AC(i, nroAluno).ToString)
                                        Else
                                            SQL = String.Format(
                                                        "INSERT INTO boletim (cod_boletim, nro_aluno, M, F, AC, S, porcentagem) values('{0}', '{1}', '{2}', '0', '0', '1', '0');",
                                                        nroBoletim, nroAluno, arrayBoletim_M(i, nroAluno).ToString)
                                        End If
                                        MySQL_atualiza(SQL)
                                        nroAluno += 1

                                    End While

                                    DefinirProgresso("S", NomeTurma, arrayNF1(i).ToString, "Inserido com sucesso!")
                                    arquivoLog("Boletim",
                                                   String.Format("{0} - {1} Importado com sucesso!", NomeTurma, arrayNF1(i)))

                                End If
                            End If
                        Next
                    End If
                End Using

                '////////////////////////////////// MAPAO HTML ///////////////
                '/////////////////////////////////////////////////////////////
            ElseIf arrayArquivosTipo(contador) = "Mapão HTML" Then

                Using leitor As New TextFieldParser(arrayArquivos(contador).ToString, System.Text.Encoding.UTF7)

                    Dim linhaAtual As String()
                    Dim nroAluno_Turma = 1

                    Dim Nro_Disciplina = 0
                    'Informamos que será importado com Delimitação
                    leitor.TextFieldType = FieldType.Delimited
                    'Informamos o Delimitador
                    leitor.SetDelimiters(" ")
                    linhaAtual = leitor.ReadFields()

                    ' /////////// CONTEUDO DESTA TURMA ARQUIVO ...
                    While Not leitor.EndOfData

                        If NomeTurma = String.Empty Then

                            ' /////////////////////////// Nome da Turma \\\\\\\\\\\\\\\\\\\\\\\
                            While (NomeTurma = String.Empty)
                                linhaAtual = leitor.ReadFields()
                                If linhaAtual(0).Contains("Turma:") = True Then
                                    NomeTurma = String.Format("{0} {1} {2}", linhaAtual(1).ToString, linhaAtual(2).ToString, linhaAtual(3).ToString)
                                End If
                            End While

                            ' ///////////// BIMESTRE...
                            While (linhaAtual(0).Contains("Tipo") = False)
                                linhaAtual = leitor.ReadFields()
                            End While

                            Dim BimestreArquivo = linhaAtual(4).ToString
                            If linhaAtual(4).ToString = "" Then
                                BimestreArquivo = linhaAtual(3).ToString
                            End If


                            If BimestreArquivo.Contains("PRIMEIRO") = True Or BimestreArquivo.Contains("Primeiro") = True Then
                                BimestreArquivo = "1"
                            ElseIf BimestreArquivo.Contains("SEGUNDO") = True Or BimestreArquivo.Contains("Segundo") = True Then
                                BimestreArquivo = "2"
                            ElseIf BimestreArquivo.Contains("TERCEIRO") = True Or BimestreArquivo.Contains("Terceiro") = True Then
                                BimestreArquivo = "3"
                            ElseIf BimestreArquivo.Contains("QUARTO") = True Or BimestreArquivo.Contains("Quarto") = True Then
                                BimestreArquivo = "4"
                            ElseIf BimestreArquivo.Contains("FINAL") = True Or BimestreArquivo.Contains("Final") = True Then
                                BimestreArquivo = "FINAL"
                            End If

                            If (MeuBimestre <> BimestreArquivo) Then
                                DefinirProgresso("F", NomeTurma, "-", "Bimestre não corresponde com o arquivo.")
                                Exit While
                            ElseIf BimestreArquivo = "FINAL" Then

                                If cbEscolhaBimestre.EditValue = "2º Avaliação Final" Then
                                    MeuBimestre2 = "2AF"
                                ElseIf cbEscolhaBimestre.EditValue = "4º Avaliação Final" Then
                                    MeuBimestre2 = "4AF"
                                End If
                            Else
                                MeuBimestre2 = BimestreArquivo
                            End If

                            linhaAtual = leitor.ReadFields()
                            If NomeTurma.Contains("MULTISSERIADA") = True Then
                                While (linhaAtual(0).Contains("Tipo") = False)
                                    linhaAtual = leitor.ReadFields()
                                End While

                                'Acrescenta Nível
                                NomeTurma = String.Format("{0} {1}", NomeTurma, linhaAtual(8))

                            End If

                            '/////////////// TURMA do ALUNO...
                            SQL = String.Format(
                                    "SELECT codigo_trma, classe FROM turma WHERE bloqueado='0' AND classe LIKE '%{0}%';",
                                    NomeTurma)
                            Nro_Turma = MySQL_consulta_campo(SQL, "codigo_trma")
                            If Nro_Turma = "0" Then
                                '... INFORMA o nome da TURMA...
                                DefinirProgresso("F", NomeTurma, "***", "Turma não cadastrada.")
                                '... SAI DESTE ARQUIVO...
                                Exit While
                            Else
                                DefinirProgresso("S", NomeTurma, "***", "Turma cadastrada.")
                            End If

                            ' ... LINHA das Disciplinas...
                            While (linhaAtual(0).ToString <> "GOE")
                                linhaAtual = leitor.ReadFields()
                            End While
                            linhaAtual = leitor.ReadFields()
                            'linhaAtual = leitor.ReadFields()

                        End If



                        Dim FormaDisciplina = linhaAtual.Length
                        Dim ProcuraDisciplina = " "

                        '... Demais disciplinas...
                        While Not (IsNumeric(linhaAtual(0).ToString))
                            FormaDisciplina = linhaAtual.Length
                            ProcuraDisciplina = " "

                            If linhaAtual(0).ToString <> String.Empty Then
                                For i = 0 To FormaDisciplina - 1
                                    ProcuraDisciplina = Trim(String.Format("{0} {1}", ProcuraDisciplina, linhaAtual(i)))
                                Next

                                ' DISCIPLINAS DIFERENTES NOMES                                    
                                If ProcuraDisciplina = "LINGUAESTRANGEIRAINGLES" Or
                                    ProcuraDisciplina = "LINGUA ESTRANGEIRA INGLES" Then
                                    ProcuraDisciplina = "INGLES"
                                    'ElseIf ProcuraDisciplina = "LINGUAPORTUGUESAELITERATURA" Or ProcuraDisciplina = "LINGUA PORTUGUESA E LITERATURA" Then
                                    '    ProcuraDisciplina = "LITERATURA"
                                ElseIf ProcuraDisciplina = "EDUCACAOFISICA" Then
                                    ProcuraDisciplina = "EDUCACAO FISICA"
                                    'ElseIf _
                                    '    ProcuraDisciplina = "LINGUAPORTUGUESA" Or
                                    '    ProcuraDisciplina = "LINGUA PORTUGUESA" Then
                                    '    ProcuraDisciplina = "PORTUGUES"
                                End If

                                contaTracos = 0
                                Dim Nro_NotasFreq = 0
                                ' ... Checa as disciplinas do ALUNO ...                                    
                                SQL = String.Format("SELECT codigo_disc, disciplina FROM disciplinas WHERE bloqueado='0' AND disciplina='{0}';", ProcuraDisciplina)
                                Nro_Disciplina = MySQL_consulta_campo(SQL, "codigo_disc").ToString
                                If Nro_Disciplina = "0" Then
                                    DefinirProgresso("F", NomeTurma, ProcuraDisciplina, "Disciplina não cadastrada.")
                                    contaTracos = 120
                                Else
                                    ' Consulta de Existe dentro do sistema o CONSOLIDADO...
                                    SQL = String.Format(
                                            "SELECT cod_nf FROM notasfreq WHERE anovigente='{0}' AND cod_bimestre='{1}' AND turma='{2}' AND disciplina='{3}';",
                                            tAno.Value, MeuBimestre2, Nro_Turma, Nro_Disciplina)
                                    Nro_NotasFreq = MySQL_consulta_campo(SQL, "cod_nf")
                                    contaTracos = 0
                                End If

                                ' ..... INDICES .....
                                ' Nome da Disciplina
                                arrayNF1(iDisciplina) = Trim(ProcuraDisciplina)
                                ' Número da Disciplina
                                arrayNF2(iDisciplina) = Nro_Disciplina
                                ' Número do Boletim
                                arrayNF3(iDisciplina) = Nro_NotasFreq

                                iDisciplina += 1

                            End If

                            linhaAtual = leitor.ReadFields()


                        End While

                        '... Procura o primeiro aluno ...                            
                        While (linhaAtual(0).ToString <> "01")
                            linhaAtual = leitor.ReadFields()
                        End While

                        Dim nroAluno_Temporario = 1
                        nroAluno_Turma = 1
                        iDisciplina = 0


                        '... Preenche os boletins na variável ...
                        While (linhaAtual(0).ToString <> "Aulas")

                            ' ... Indice das Disciplinas                                                              
                            nroAluno = CInt(linhaAtual(0).ToString)

                            If nroAluno_Temporario <> nroAluno Then
                                nroAluno_Temporario = linhaAtual(0).ToString
                                iDisciplina = 0
                                ' contaTracos = 0
                            End If
                            linhaAtual = leitor.ReadFields()

                            ' VERIFICA SE EXISTE NOTAS...
                            If ((linhaAtual(0).ToString) = "-") Then

                                'Conta Traços
                                arrayCheck_M(iDisciplina) += 1

                                'Pega o nro total de alunos da classe
                                nroAluno_Turma = nroAluno

                                'arrayBoletim_M(iDisciplina2, nroAluno) = 19
                                SQL = String.Format("SELECT status FROM aluno WHERE anovigente='{0}' AND turma='{1}' AND nro='{2}';",
                                               tAno.Value, Nro_Turma, nroAluno)
                                arrayBoletim_M(iDisciplina, nroAluno) = TesteEvasaoEscolar((MySQL_consulta_campo(SQL, "status").ToString))

                            Else
                                arrayBoletim_M(iDisciplina, nroAluno) = linhaAtual(0).ToString.Replace(",", ".")
                                'Pega o nro total de alunos da classe
                                nroAluno_Turma = nroAluno

                            End If
                            linhaAtual = leitor.ReadFields()


                            If cbEscolhaBimestre.Text.Contains("Final") = False Then

                                ' VERIFICA SE EXISTE FALTAS ...
                                If ((linhaAtual(0).ToString) = "-") Then
                                    'Conta Traços
                                    arrayCheck_F(iDisciplina) += 1

                                    arrayBoletim_F(iDisciplina, nroAluno) = 0
                                Else
                                    arrayBoletim_F(iDisciplina, nroAluno) = linhaAtual(0).ToString.Replace(",", ".")
                                End If
                                linhaAtual = leitor.ReadFields()

                                ' VERIFICA SE EXISTE AC ...
                                If ((linhaAtual(0).ToString) = "-") Then
                                    arrayBoletim_AC(iDisciplina, nroAluno) = "0"
                                Else
                                    If linhaAtual(0).ToString.Contains("-") Then
                                        arrayBoletim_AC(iDisciplina, nroAluno) = "0"
                                    Else
                                        arrayBoletim_AC(iDisciplina, nroAluno) = linhaAtual(0).ToString.Replace(",", ".")
                                    End If

                                End If
                                linhaAtual = leitor.ReadFields()

                            End If

                            iDisciplina += 1

                        End While

                        iDisciplina = 0
                        Try
                            '... Preenche as Aulas Dadas
                            While (linhaAtual(0).ToString <> " ")
                                If linhaAtual(0).ToString = "Aulas" Then
                                    arrayBoletim_Aulas(iDisciplina) = linhaAtual(2).ToString
                                    iDisciplina += 1
                                End If
                                linhaAtual = leitor.ReadFields()
                            End While
                        Catch ex As Exception
                            Exit While

                        End Try

                        ' /////////// FIM DO CONTEUDO DESTA TURMA ARQUIVO ...
                    End While

                    ' ANALISA...
                    If Nro_Turma <> 0 Then
                        '... Procura boletins EXISTENTES ...
                        For i = 0 To iDisciplina - 1

                            If (arrayCheck_M(i) = nroAluno_Turma And arrayCheck_F(i) = nroAluno_Turma) Or (arrayNF2(i) = "0") Then
                                ' NAO EXISTE BOLETIM ...
                                DefinirProgresso("F", NomeTurma, arrayNF1(i).ToString, "Boletim/Disciplina não encontrada.")
                            Else
                                '... EXISTE BOLETIM ...
                                '... DefinirProgresso("S", NomeTurma, arrayNF1(i).ToString, "Boletim pronto para carregar!")
                                'Or (arrayNF3(i) <> Nothing) *****************
                                If (arrayNF3(i) <> "0") Then

                                    If cbAtualizarNotas.Checked = False Then
                                        DefinirProgresso("F", NomeTurma, arrayNF1(i).ToString, "Boletim já existe!")

                                    Else

                                        ' Delete boletim / notasfreq
                                        ' 
                                        SQL = "DELETE FROM boletim WHERE cod_boletim='" & arrayNF3(i).ToString & "';"
                                        MySQL_atualiza(SQL)
                                        SQL = "DELETE FROM notasfreq WHERE cod_nf='" & arrayNF3(i).ToString & "';"
                                        MySQL_atualiza(SQL)


                                        SQL = String.Format("INSERT INTO notasfreq (turma, disciplina, cod_bimestre, qtdadeaulas, previsaoaulas, anovigente, dt_criacao,dt_atualizacao) values('{0}', '{1}', '{2}', '{5}', '{5}', '{3}', '{4}', '{4}'); SELECT LAST_INSERT_ID() AS cod_nf;",
                                                  Nro_Turma, arrayNF2(i).ToString, MeuBimestre2, tAno.Value, Format(Date.Now, "yyyy-MM-dd HH:mm:ss").ToString, arrayBoletim_Aulas(i).ToString)
                                        Dim nroBoletim = MySQL_atualiza(SQL)

                                        ' ... /// INSERIR NOTAS / TURMA / DISCIPLINA ...
                                        nroAluno = 1
                                        While (arrayBoletim_M(i, nroAluno) <> Nothing)

                                            If cbEscolhaBimestre.Text.Contains("Final") = False Then
                                                SQL = String.Format("INSERT INTO boletim (cod_boletim, nro_aluno, M, F, AC, S, porcentagem) values('{0}', '{1}', '{2}', '{3}', '{5}', 'N', '{4}');",
                                                                                                    nroBoletim, nroAluno, arrayBoletim_M(i, nroAluno).ToString,
                                                                                                    arrayBoletim_F(i, nroAluno).ToString,
                                                                                                    Resultado_Porcentagem(arrayBoletim_F(i, nroAluno).ToString,
                                                                                                                          arrayBoletim_AC(i, nroAluno).ToString,
                                                                                                                          arrayBoletim_Aulas(i).ToString),
                                                                                                    arrayBoletim_AC(i, nroAluno).ToString)
                                            Else
                                                SQL = String.Format(
                                                        "INSERT INTO boletim (cod_boletim, nro_aluno, M, F, AC, S, porcentagem) values('{0}', '{1}', '{2}', '0', '0', '1', '0');",
                                                        nroBoletim, nroAluno, arrayBoletim_M(i, nroAluno).ToString)
                                            End If
                                            MySQL_atualiza(SQL)
                                            nroAluno += 1

                                            If arrayBoletim_M(i, nroAluno) = Nothing Then
                                                nroAluno += 1
                                                If arrayBoletim_M(i, nroAluno) = Nothing Then
                                                    Exit While
                                                End If
                                            End If

                                        End While

                                        DefinirProgresso("S", NomeTurma, arrayNF1(i).ToString, "Atualizado com sucesso!")
                                        arquivoLog("Boletim", String.Format("{0} - {1} Atualizado com sucesso!", NomeTurma, arrayNF1(i)))

                                    End If

                                Else

                                    ' /// INSERIR BOLETIM ...
                                    SQL = String.Format("INSERT INTO notasfreq (turma, disciplina, cod_bimestre, qtdadeaulas, previsaoaulas, anovigente, dt_criacao,dt_atualizacao) values('{0}', '{1}', '{2}', '{5}', '{5}', '{3}', '{4}', '{4}'); SELECT LAST_INSERT_ID() AS cod_nf;",
                                              Nro_Turma, arrayNF2(i).ToString, MeuBimestre2, tAno.Value, Format(Date.Now, "yyyy-MM-dd HH:mm:ss").ToString, arrayBoletim_Aulas(i).ToString)
                                    Dim nroBoletim = MySQL_atualiza(SQL)

                                    ' ... /// INSERIR NOTAS / TURMA / DISCIPLINA...
                                    nroAluno = 1
                                    While (arrayBoletim_M(i, nroAluno) <> Nothing)

                                        If cbEscolhaBimestre.Text.Contains("Final") = False Then
                                            SQL = String.Format("INSERT INTO boletim (cod_boletim, nro_aluno, M, F, AC, S, porcentagem) values('{0}', '{1}', '{2}', '{3}', '{5}', 'N', '{4}');",
                                                                                                nroBoletim, nroAluno, arrayBoletim_M(i, nroAluno).ToString,
                                                                                                arrayBoletim_F(i, nroAluno).ToString,
                                                                                                Resultado_Porcentagem(arrayBoletim_F(i, nroAluno).ToString,
                                                                                                                      arrayBoletim_AC(i, nroAluno).ToString,
                                                                                                                      arrayBoletim_Aulas(i).ToString),
                                                                                                arrayBoletim_AC(i, nroAluno).ToString)
                                        Else
                                            SQL = String.Format(
                                                    "INSERT INTO boletim (cod_boletim, nro_aluno, M, F, AC, S, porcentagem) values('{0}', '{1}', '{2}', '0', '0', '1', '0');",
                                                    nroBoletim, nroAluno, arrayBoletim_M(i, nroAluno).ToString)
                                        End If
                                        MySQL_atualiza(SQL)
                                        nroAluno += 1

                                        If arrayBoletim_M(i, nroAluno) = Nothing Then
                                            nroAluno += 1
                                            If arrayBoletim_M(i, nroAluno) = Nothing Then
                                                Exit While
                                            End If
                                        End If

                                    End While

                                    DefinirProgresso("S", NomeTurma, arrayNF1(i).ToString, "Inserido com sucesso!")
                                    arquivoLog("Boletim",
                                               String.Format("{0} - {1} Importado com sucesso!", NomeTurma, arrayNF1(i)))

                                End If
                            End If
                        Next
                    End If
                End Using

            End If

            '// PROXIMO ARQUIVO...
            contador += 1
            bwCarregar.ReportProgress(contador)

        End While

    End Sub

    Private Sub bwCarregar_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bwCarregar.ProgressChanged

        AtualizarBarra_Invoke(e.ProgressPercentage, contadorArquivos)

    End Sub

    Private Sub bwCarregar_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bwCarregar.RunWorkerCompleted

        btCarregar.Enabled = True
        AtualizarBarra_Invoke(contadorArquivos, contadorArquivos)
        MsgBox("Sucesso!", MsgBoxStyle.Information, "Importação de Notas")
        btImprimir.Enabled = True

    End Sub

    Private Sub btImprimir_Click(sender As Object, e As EventArgs) Handles btImprimir.Click

        frmPreview_Titulo = "Relatório - Trazer Mapão!"
        Dim Link As New PrintableComponentLink(New PrintingSystem()) With {.Component = gridBoletim}
        AddHandler Link.CreateMarginalHeaderArea, AddressOf frmPreview_Padrao
        Link.CreateDocument()
        Link.ShowPreview()

    End Sub

    Private Sub PictureEdit1_Click(sender As Object, e As EventArgs) Handles PictureEdit1.Click

        ' Abre direto no desktop...
        openFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        openFD.Title = "Abrir arquivo"
        'openFD.Filter = "MAPÃO - Excel|*.xltx;*.xls;*.xlsx|BOLETINS - PDF (*.pdf)|*.pdf|Todos os arquivos (*.*)|*.*"
        openFD.Filter = "Arquivo em Excel|*.xltx;*.xls;*.xlsx"
        openFD.FilterIndex = 1
        openFD.ShowDialog()

        tbArquivo.Text = openFD.FileName

    End Sub

    Private Sub frmImpMapao_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Try
            bwCarregar.CancelAsync()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub viewBoletim_MasterRowExpanded(sender As Object, e As CustomMasterRowEventArgs) Handles viewBoletim.MasterRowExpanded
        Dim grid As GridView = sender
        Dim detail As GridView = grid.GetDetailView(e.RowHandle, e.RelationIndex)
        detail.OptionsView.ColumnAutoWidth = True
        detail.BestFitColumns()


    End Sub
End Class
