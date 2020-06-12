Imports System.ComponentModel
Imports System.IO

Public Delegate Sub IncrementProgressDelegate()
Public Delegate Sub Delegação(texto As String)
Public Delegate Sub Delegação2(Status As String, texto As String)

Public Class frmImpListas

    Dim ArquivoTipo As String
    Dim contadorTurmas = 0
    Dim contadorArquivos = 0
    Dim m, Final

    Public Sub DefinirBarra(texto As String)
        On Error Resume Next
        If barra.InvokeRequired Then
            Dim d As New Delegação(AddressOf DefinirBarra)
            barra.Text = barra.Invoke(d, texto)
            barra.Properties.Maximum = Final
            barra.EditValue = m
        Else
            barra.Text = texto
            barra.Properties.Maximum = Final
            barra.EditValue = m
        End If
    End Sub

    Sub ImportaTxt(substituirTxt As Boolean)

        Final = lstTurma.Items.Count - 1

        For inicio = 0 To Final

            m = inicio
            Dim linhaTxt = inicio + 1 & "/" & Final + 1 & " | Importando: " & lstTurma.Items(inicio) & "..."
            DefinirBarra(linhaTxt)

            ' Verifica o código da turma...
            SQL = String.Format("SELECT codigo_trma FROM turma WHERE bloqueado='0' AND classe='{0}';", lstTurma.Items(inicio))
            Dim codigoTurma = MySQL_consulta_campo(SQL, "codigo_trma")

            Dim caminhoArquivo As String = lstArquivos.Items(inicio)
            Dim PDF_conteudo As String = ExtractTextFromPDF(caminhoArquivo)
            Dim PDF_partes As String() = PDF_conteudo.Split(ControlChars.CrLf.ToCharArray)

            ' DELETE alunos acima da nova lista...
            SQL = String.Format("DELETE FROM aluno WHERE anovigente='{0}' AND turma='{1}' AND nro>'0';", tAno.Value, codigoTurma)
            MySQL_atualiza(SQL)
            Dim indice = 0
            Dim indice_Linha = 0
            Dim EntrouArquivo = False

            ' Entra para analisar o arquivo.
            For Each part As String In PDF_partes

                Dim linhaAtual As String() = part.Split(New Char() {" "c})
                ' ... Verifica se já entrou no arquivo ...
                ' If (EntrouArquivo = True And part.ToString.Contains("Página") = False) Then

                If (EntrouArquivo = True And part.ToString.Contains("Página") = False) Or (linhaAtual(0).ToString.Length = 1 Or linhaAtual(0).ToString.Length = 2) Then
                    If (part.ToString.Contains("/") = False And part <> String.Empty) And (part.ToString.Contains("ENSINO") = False And part.ToString.Contains("MEDIO") = False) Then
                        part = part & " " & PDF_partes(indice + 2).ToString
                        linhaAtual = part.Split(New Char() {" "c})

                        If PDF_partes(indice + 2).ToString.Contains("/") = False And part <> String.Empty And (part.ToString.Contains("ENSINO") = False And part.ToString.Contains("MEDIO") = False) Then
                            part = part & " " & PDF_partes(indice + 4)
                            linhaAtual = part.Split(New Char() {" "c})

                            If PDF_partes(indice + 4).Contains("/") = False Then
                                part = part & " " & PDF_partes(indice + 6)
                                linhaAtual = part.Split(New Char() {" "c})
                            End If

                        End If
                    End If
                End If

                Dim posicao = 0
                Try
                    If linhaAtual(0).ToString.Contains("ENSINO") = True Or (linhaAtual(1).ToString.Length = 1 Or linhaAtual(1).ToString.Length = 2) Then
                        posicao = 0
                    Else
                        posicao = 1
                    End If
                Catch ex As Exception
                End Try

                ' ... Encontrar o primeiro aluno ... 
                'While (part.Contains("/") = True And part.ToString <> String.Empty And (linhaAtual(0).ToString.Length = 1 Or linhaAtual(0).ToString.Length = 2))
                While (part.Contains("/") = True And part.ToString <> String.Empty And (linhaAtual(0).ToString.Length = 1 Or linhaAtual(0).ToString.Length = 2))

                    If IsNumeric(linhaAtual(0)) = False Then
                        Exit While
                    End If

                    EntrouArquivo = True

                    Dim aluno_Nro, aluno_Nome, aluno_RA, aluno_Digito, aluno_Nascimento, Status_Aluno
                    aluno_Nro = linhaAtual(1 - posicao).ToString
                    aluno_Nome = linhaAtual(2 - posicao).ToString

                    Dim i = 3
                    While (linhaAtual(i - posicao).ToString.Contains("0")) = False
                        aluno_Nome = aluno_Nome & " " & linhaAtual(i - posicao).ToString
                        i = i + 1
                    End While

                    aluno_RA = linhaAtual(i - posicao).ToString
                    aluno_Digito = linhaAtual(i + 1 - posicao).ToString

                    Try
                        aluno_Nascimento = linhaAtual(i + 3 - posicao).ToString
                    Catch ex As Exception
                        aluno_Nascimento = linhaAtual(i + 2 - posicao).ToString
                        aluno_Digito = String.Empty
                    End Try

                    Status_Aluno = "0"

                    Try
                        Status_Aluno = TesteEvasaoEscolar_Nome(linhaAtual(i + 4 - posicao).ToString)
                    Catch ex As Exception
                        Status_Aluno = "0"
                    End Try

                    SQL = String.Format("INSERT INTO aluno (nome, turma, nro, anovigente, ra, ra_digito, data, status) values('{0}', {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}');",
                                 aluno_Nome.ToString.Replace("'", String.Empty), codigoTurma, aluno_Nro.ToString, tAno.Value, aluno_RA.ToString,
                                 aluno_Digito.ToString, aluno_Nascimento.ToString, Status_Aluno)
                    MySQL_atualiza(SQL)
                    Exit While

                    indice = indice + 1

                End While

                'If EntrouNaLista = True Then
                ' Monta a Linha
                'End If

                ''Monta a linha para cadastrar
                '' Se a linha terminar, sem demais dados.
                'Dim part_Anterior = PDF_partes(indice - 2).ToString()

                ''//// Linha Anterior

                ''//// Fim da Linha Anterior


                indice = indice + 1

            Next

            'Using leitor As New TextFieldParser(conteudoPDF)

            '        Dim linhaAtual As String()
            '        'Informamos que será importado com Delimitação
            '        leitor.TextFieldType = FieldType.Delimited
            '        'Informamos o Delimitador
            '        leitor.SetDelimiters(";")

            '        ' /////////// CONTEUDO DESTA TURMA ARQUIVO...
            '        While Not leitor.EndOfData

            '            linhaAtual = leitor.ReadFields()

            '            ' PRIMEIRO ALUNO
            '            'While (IsNumeric(linhaAtual(2)) = True)

            '            '    Dim aluno_Nro = linhaAtual(2).ToString
            '            '    Dim aluno_Nome = linhaAtual(3).ToString
            '            '    Dim aluno_RA = linhaAtual(4).ToString
            '            '    Dim aluno_Digito = linhaAtual(5).ToString
            '            '    Dim aluno_Nascimento = linhaAtual(7).ToString
            '            '    Dim Status_Aluno

            '            '    'Identifica qual é o arquivo, se na primeira coluna está o Tipo de Ensino, ou NÃO.
            '            '    If linhaAtual(0).ToString.Contains("ENSINO") = True Then

            '            '        aluno_Nro = linhaAtual(2).ToString
            '            '        aluno_Nome = linhaAtual(3).ToString
            '            '        aluno_RA = linhaAtual(4).ToString
            '            '        aluno_Digito = linhaAtual(5).ToString
            '            '        aluno_Nascimento = linhaAtual(7).ToString
            '            '        Status_Aluno = "0"

            '            '        If linhaAtual(8).ToString = "" Then
            '            '            Status_Aluno = "0"
            '            '        Else
            '            '            Status_Aluno = TesteEvasaoEscolar_Nome(linhaAtual(8).ToString)
            '            '        End If

            '            '        '  Pega o sobrenome
            '            '        linhaAtual = leitor.ReadFields()
            '            '        If linhaAtual(3).ToString <> "" Then
            '            '            aluno_Nome = aluno_Nome & " " & linhaAtual(3).ToString
            '            '        End If

            '            '    Else

            '            '        aluno_Nro = linhaAtual(1).ToString
            '            '        aluno_Nome = linhaAtual(2).ToString
            '            '        aluno_RA = linhaAtual(3).ToString
            '            '        aluno_Digito = linhaAtual(5).ToString
            '            '        aluno_Nascimento = linhaAtual(7).ToString.Replace("P", "/")
            '            '        Status_Aluno = "0"

            '            '        If linhaAtual(8).ToString = "" Then
            '            '            Status_Aluno = "0"
            '            '        Else
            '            '            Status_Aluno = TesteEvasaoEscolar_Nome(linhaAtual(8).ToString)
            '            '        End If

            '            '        '  Pega o sobrenome
            '            '        linhaAtual = leitor.ReadFields()
            '            '        If linhaAtual(3).ToString <> "" Then
            '            '            aluno_Nome = aluno_Nome & " " & linhaAtual(3).ToString
            '            '        End If

            '            '    End If





            '            '    SQL = String.Format("INSERT INTO aluno (nome, turma, nro, anovigente, ra, ra_digito, data, status) values('{0}', {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}');",
            '            '         aluno_Nome.ToString.Replace("'", ""), codigoTurma, aluno_Nro.ToString, tAno.Value, aluno_RA.ToString,
            '            '         aluno_Digito.ToString, aluno_Nascimento.ToString, Status_Aluno)
            '            '    MySQL_atualiza(SQL)
            '            '    '... Lista até o final ...

            '            'End While

            '        End While

            '    End Using


        Next
    End Sub

    Sub Refresh_Arquivos()

        contadorArquivos = 0

        Dim nomearquivo As String = Path.GetFileName(lbArquivo.EditValue)
        importacao_alunos = lbArquivo.EditValue
        importacao_alunos = importacao_alunos.Replace(nomearquivo, String.Empty)
        importacao_alunos = Trim(importacao_alunos)

        '... Identifica o arquivo que o cliente quer.
        If _
            InStr(lbArquivo.EditValue, ".pdf") Or InStr(lbArquivo.EditValue, ".Pdf") Or
            InStr(lbArquivo.EditValue, ".PDF") Then
            ArquivoTipo = "PDF"
        End If

        Dim nomearq As String
        lstArquivos.Items.Clear()
        For Each nomearq In Directory.GetFiles(importacao_alunos, "*.*")

            If ArquivoTipo = "PDF" Then
                If _
                    InStr(lbArquivo.EditValue, ".pdf") Or InStr(lbArquivo.EditValue, ".Pdf") Or
                    InStr(lbArquivo.EditValue, ".PDF") Then
                    lstArquivos.Items.Add(New String(nomearq))
                    contadorArquivos = contadorArquivos + 1
                End If
            End If
        Next
        lbArquivos.Text = contadorArquivos
    End Sub

    Sub Refresh_Turmas()

        contadorTurmas = 0

        ' Pega as turmas...
        SQL = "SELECT classe FROM turma WHERE bloqueado='0' ORDER BY classe ASC;"
        Dim Turma = MySQL_combobox(SQL)
        lstTurma.Items.Clear()
        For Each r In Turma.Rows
            lstTurma.Items.Add(r("classe").ToString)
        Next

        contadorTurmas = Turma.Rows.Count
        lbTurmas.Text = Turma.Rows.Count
    End Sub

    Private Sub frmImpListas2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        tAno.Value = AnoVigente
        Refresh_Turmas()

    End Sub

    Private Sub bwMaisEscola_DoWork(sender As Object, e As DoWorkEventArgs) Handles bwMaisEscola.DoWork
        If ArquivoTipo = "TXT" Then
            ImportaTxt(False)
        ElseIf ArquivoTipo = "PDF" Then
            ImportaTxt(True)
        End If
    End Sub

    Private Sub bwMaisEscola_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) _
        Handles bwMaisEscola.RunWorkerCompleted
        LiberaPublicacao("aluno")

        MsgBox("Importação finalizada com sucesso!", MsgBoxStyle.Information, "Mais Escola!")

        ' Visualiza o relatório ...
        If cbVTurmas.Checked = True Then
            Try
                Dim fRpt As New frmRpt_Auxiliar
                frmRpt_anovigente = tAno.Value
                frmRpt_Tipo = "TurmaTodas"
                SQL_frmRPT = meuRPT2("Lista de Alunos", "Reunião de Pais e Mestres")
                fRpt.Show()
            Catch ex As Exception
            End Try
        End If

        Me.Close()

    End Sub

    Private Sub pbProcurar_Click(sender As Object, e As EventArgs) Handles pbProcurar.Click
        ' Abre direto no desktop...

        openFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

        openFD.Title = "Abrir arquivo"
        openFD.Filter = "Lista Piloto PDF (*.pdf)|*.pdf"
        openFD.FilterIndex = 1
        openFD.ShowDialog()

        lbArquivo.Text = openFD.FileName
    End Sub

    Private Sub arquivoSobe_Click(sender As Object, e As EventArgs) Handles arquivoOrigSobe.Click
        Dim Index As Integer = lstArquivos.SelectedIndex    'Index of selected item
        Dim Swap As Object = lstArquivos.SelectedItem       'Selected Item
        If Not (Swap Is Nothing) Then 'If something is selected...
            If Index <= 0 Then
                Exit Sub
            End If
            lstArquivos.Items.RemoveAt(Index)                   'Remove it
            lstArquivos.Items.Insert(Index - 1, Swap)           'Add it back in one spot up
            lstArquivos.SelectedItem = Swap                     'Keep this item selected
        End If
    End Sub

    Private Sub arquivoDesce_Click(sender As Object, e As EventArgs) Handles arquivoOrigDesce.Click
        Dim Index As Integer = lstArquivos.SelectedIndex    'Index of selected item
        Dim Swap As Object = lstArquivos.SelectedItem       'Selected Item
        If Not (Swap Is Nothing) Then 'If something is selected...
            If Index + 1 >= lstArquivos.Items.Count Then
                Exit Sub
            End If
            lstArquivos.Items.RemoveAt(Index)                   'Remove it
            lstArquivos.Items.Insert(Index + 1, Swap)           'Add it back in one spot up
            lstArquivos.SelectedItem = Swap                     'Keep this item selected
        End If
    End Sub

    Private Sub arquivoRefresh_Click(sender As Object, e As EventArgs) Handles arquivoRefresh.Click
        Refresh_Arquivos()
    End Sub

    Private Sub arquivoRetira_Click(sender As Object, e As EventArgs) Handles arquivoRetira.Click

        Dim Index As Integer = lstArquivos.SelectedIndex - 1

        If lstArquivos.SelectedIndex <> -1 Then
            ' MsgBox("Escolha uma turma!", MsgBoxStyle.Information, "Mais Escola!")
            contadorArquivos -= 1
            'Remove da lista
            lstArquivos.Items.Remove(lstArquivos.SelectedItems(0))
            Try
                lstArquivos.SelectedIndex = Index
            Catch ex As Exception
            End Try
            lbArquivos.Text = contadorArquivos
        End If

    End Sub

    Private Sub turmaSobe_Click(sender As Object, e As EventArgs) Handles turmaDestSobe.Click

        Dim Index As Integer = lstTurma.SelectedIndex    'Index of selected item
        Dim Swap As Object = lstTurma.SelectedItem       'Selected Item
        If Not (Swap Is Nothing) Then 'If something is selected...
            If Index <= 0 Then
                Exit Sub
            End If
            lstTurma.Items.RemoveAt(Index)                   'Remove it
            lstTurma.Items.Insert(Index - 1, Swap)           'Add it back in one spot up
            lstTurma.SelectedItem = Swap                     'Keep this item selected
        End If

    End Sub

    Private Sub turmaDesce_Click(sender As Object, e As EventArgs) Handles turmaDestDesce.Click

        Dim Index As Integer = lstTurma.SelectedIndex    'Index of selected item
        Dim Swap As Object = lstTurma.SelectedItem       'Selected Item
        If Not (Swap Is Nothing) Then 'If something is selected...
            If Index + 1 >= lstTurma.Items.Count Then
                Exit Sub
            End If
            lstTurma.Items.RemoveAt(Index)                   'Remove it
            lstTurma.Items.Insert(Index + 1, Swap)           'Add it back in one spot up
            lstTurma.SelectedItem = Swap                     'Keep this item selected
        End If

    End Sub

    Private Sub turmaRefresh_Click(sender As Object, e As EventArgs) Handles turmaRefresh.Click
        Refresh_Turmas()
    End Sub

    Private Sub turmaRetira_Click(sender As Object, e As EventArgs) Handles turmaRetira.Click

        Dim Index As Integer = lstTurma.SelectedIndex - 1

        If lstTurma.SelectedIndex <> -1 Then
            'MsgBox("Escolha uma turma!", MsgBoxStyle.Information, "Mais Escola!")
            contadorTurmas -= 1
            'Remove da lista
            lstTurma.Items.Remove(lstTurma.SelectedItems(0))
            Try
                lstTurma.SelectedIndex = Index
            Catch ex As Exception
            End Try
            lbTurmas.Text = contadorTurmas
        End If

    End Sub

    Private Sub btImportar_Click(sender As Object, e As EventArgs) Handles btImportar.Click

        If contadorArquivos <> contadorTurmas Then
            MsgBox("Antes de importar," & vbCrLf & "é necessário ter o mesmo número de arquivos relacionado às turmas!",
                   MsgBoxStyle.Information, "Mais Escola!")
        Else

            If MsgBox("Deseja importar os alunos?", MsgBoxStyle.YesNo, "Importar lista de alunos") = MsgBoxResult.Yes _
                Then
                ' Aguardar...
                DefinirBarra("Carregando...")
                ' Verifica se contém lista...
                bwMaisEscola.RunWorkerAsync()
            End If

        End If

    End Sub

    Private Sub openFD_FileOk(sender As Object, e As CancelEventArgs) Handles openFD.FileOk

    End Sub

    Private Sub lbArquivo_TextChanged(sender As Object, e As EventArgs) Handles lbArquivo.TextChanged

        If lbArquivo.Text = String.Empty Or lbArquivo.Text = "OpenFileDialog1" Then
            MsgBox("Por favor, escolha o arquivo antes de continuar!", MsgBoxStyle.Information, "Mais Escola!")
        Else
            Refresh_Arquivos()
        End If

    End Sub

End Class