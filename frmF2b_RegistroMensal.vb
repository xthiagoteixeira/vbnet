Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraEditors

Public Class frmF2b_RegistroMensal
    Dim idCodNF, idGrupo, CodigoFechamento, CodigoCliente, Novo_idCodNF, idCodNF_Referencia, DiasVencimento
    
    Dim MinhaDG As DataTable

    Sub Refresh_Sacados()

        If cbGrupo.text  <> "" Then
            
            Try
                SQL = "SELECT " _
                              & "clientes.Nome, " _
                              & "tabela.valor_atual AS Mensalidade, " _
                              & "grupos.grupo, " _
                              & "clientes.bloqueado AS Bloqueado, " _
                              & "clientes.idCliente AS Codigo, " _
                              & "tabela.Nome AS NomeTabela, " _
                              & "clientes.bloqueado AS CodigoBloqueado " _
                              & "FROM " _
                              & "clientes " _
                              & "INNER JOIN tabela ON tabela.idPagTabela = clientes.idPagTabela INNER JOIN grupos ON clientes.idGrupo = grupos.idGrupo WHERE grupos.grupo='" & cbGrupo.Text & "' ORDER BY clientes.bloqueado, clientes.Nome ASC;"
                gridSacados.DataSource = Nothing
                gridSacados.DataSource = MySQL_datatable(SQL)

                RepositoryItemCheckEdit2.ValueChecked = Convert.ToByte(1)
                        RepositoryItemCheckEdit2.ValueUnchecked = Convert.ToByte(0)

                        viewSacados.Columns(0).Width = "100"
                        viewSacados.Columns(1).Width = "50"
                        viewSacados.Columns(2).Width = "50"

                        viewSacados.Columns(0).OptionsColumn.AllowEdit = False
                        viewSacados.Columns(1).OptionsColumn.AllowEdit = False
                        viewSacados.Columns(2).OptionsColumn.AllowEdit = False

                        viewSacados.Columns(3).Width = "30"
                        viewSacados.Columns(3).OptionsColumn.AllowEdit = True
                        viewSacados.Columns(3).ColumnEdit = RepositoryItemCheckEdit2

                        viewSacados.Columns(4).Visible = False
                        viewSacados.Columns(5).Visible = False
                viewSacados.Columns(6).Visible = False


            Catch ex As Exception
                        MsgBox("Erro: " & ex.Message, MsgBoxStyle.Information, "Mais Projetos!")

                    End Try
        End If
        

    End Sub


    Sub Refresh_Lancamento()
        Try

            SQL = "SELECT " _
                       & "grupos.grupo, " _
                       & "notafiscal.periodo_inicial AS Inicio, " _
                       & "notafiscal.periodo_final AS Final, " _
                       & "notafiscal.vencimento AS Vencimento, " _
                       & "notafiscal.Fechamento AS Fechamento, notafiscal.fechamento AS codigoFechamento, notafiscal.idNF AS CodigoNF " _
                       & "FROM " _
                       & "grupos " _
                       & "INNER JOIN notafiscal ON notafiscal.idGrupo = grupos.idGrupo ORDER BY notafiscal.idNF DESC;"
            gridLancamento.DataSource = Nothing
            gridLancamento.DataSource = MySQL_datatable(SQL)

            RepositoryItemCheckEdit1.ValueChecked = Convert.ToByte(0)
            RepositoryItemCheckEdit1.ValueUnchecked = Convert.ToByte(1)
            viewLancamento.Columns(4).ColumnEdit = RepositoryItemCheckEdit1

            Alinhar_Lancamento()
        Catch ex As Exception

        End Try


    End Sub

    Sub Alinhar_Lancamento()

        Try
            viewLancamento.Columns(0).Width = "50"
            viewLancamento.Columns(1).Width = "50"
            viewLancamento.Columns(2).Width = "50"
            viewLancamento.Columns(3).Width = "100"
            viewLancamento.Columns(4).Width = "50"
            viewLancamento.Columns(5).Visible = False
            viewLancamento.Columns(6).Visible = False
            viewLancamento.Columns(7).Visible = False

        Catch ex As Exception
        End Try

    End Sub

    Sub Refresh_Boletos(idCodNF)

        SQL = String.Format("SELECT clientes.Nome, notafiscal_boleto.valor_bruto AS ValorBruto, notafiscal_boleto.status, notafiscal.idNF, clientes.idCliente, clientes.CNPJ, clientes.email, clientes.email2, clientes.vencimento AS Vencimento FROM notafiscal INNER JOIN notafiscal_boleto ON notafiscal.idNF = notafiscal_boleto.idNF INNER JOIN clientes ON notafiscal_boleto.idCliente = clientes.idCliente INNER JOIN tabela ON clientes.idPagTabela = tabela.idPagTabela WHERE notafiscal.idNF='{0}' ORDER BY notafiscal_boleto.status, notafiscal_boleto.valor_bruto ASC;", idCodNF)
        MinhaDG = MySQL_datatable(SQL)
        gridBoletos.DataSource = Nothing
        gridBoletos.DataSource = MinhaDG
        Alinhar_Boletos()

    End Sub

    Sub Alinhar_Boletos()

        Try
            viewBoletos.Columns(0).Width = "100"
            viewBoletos.Columns(1).Width = "25"
            viewBoletos.Columns(2).Width = "25"
            viewBoletos.Columns(3).Visible = False
            viewBoletos.Columns(4).Visible = False
            viewBoletos.Columns(5).Visible = False
            viewBoletos.Columns(6).Visible = False
            viewBoletos.Columns(7).Visible = False
            viewBoletos.Columns(8).Visible = False
            viewBoletos.Columns(9).Visible = False
            viewBoletos.Columns(10).Visible = False

        Catch ex As Exception
        End Try

    End Sub

    Private Sub frmAdmNF_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '... Adiciona os grupos disponíveis
        SQL = "SELECT Grupo FROM grupos WHERE bloqueado='0' ORDER BY Grupo ASC;"
        cbGrupo.Properties.DataSource = MySQL_combobox(SQL)
        cbGrupo.Properties.DisplayMember = "Grupo"
        
        Refresh_Lancamento()
        Refresh_Sacados()

    End Sub

    Private Sub bwEmitirBoletos_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bwEmitirBoletos.DoWork

        Try

            For Each r In MinhaDG.Rows
                                                                
                Dim NomeCliente = r("Nome").ToString
                Dim CodigoCliente = r("idCliente").ToString
                Dim ValorBruto = r("ValorBruto").ToString
                Dim Email01 = r("email").ToString
                Dim Email02 = r("email2").ToString
                Dim MeuCNPJ = r("CNPJ").ToString
                Dim Retorno As String = F2B_Cobranca(CodigoCliente, ValorBruto,
                "B", CodigoCliente, cbGrupo.Text & "! - Boleto", NomeCliente, Email01, Email02, MeuCNPJ, DiasVencimento)

                Dim matriz As String() = Retorno.Split(";")
                numeroDocumento = matriz(0).ToString
                cobrancaURL = matriz(1).ToString

                'If Enviado_Retorno = True Then
                SQL = String.Format("UPDATE notafiscal_boleto SET nroCobranca='{0}', url='{3}' WHERE idNF='{1}' AND idCliente='{2}';", numeroDocumento, Novo_idCodNF, CodigoCliente, cobrancaURL)
                MySQL_atualiza(SQL)

            Next

        Catch ex As Exception
            MsgBox("Erro: " & ex.Message, MsgBoxStyle.Information, "Erro na emissão!")
        End Try

    End Sub

    Private Sub btCriar_Click(sender As Object, e As EventArgs) Handles btCriar.Click

        If (f2b_ContaNro = "" Or f2b_ContaSenha = "" Or f2b_NomeCompleto = "") Or (f2b_ContaNro = "0" Or f2b_ContaSenha = "0" Or f2b_NomeCompleto = "0") Then
            MsgBox("Preencher as informações da conta f2b", MsgBoxStyle.Information, "Configuração")
        Else

            DiasVencimento = DateDiff("d", Now, cbVencimento.EditValue)
            
            If cbDataFinal.Text <> "" And cbDataInicial.Text <> "" And cbGrupo.Text <> "" And DiasVencimento >= 0 Then
                                
                ' Checa se os campos estão ok...

                If MsgBox(String.Format("Efetuar novos registros?{0}Período: {1} à {2}", vbCrLf, cbDataInicial.EditValue, cbDataFinal.EditValue), MsgBoxStyle.YesNo, "Criar Novos Registros") = MsgBoxResult.Yes Then

                    barra.Properties.Paused = False

                    ' Cria a nota fiscal...
                    SQL = String.Format("INSERT INTO notafiscal (periodo_inicial, periodo_final, emissao, vencimento, idGrupo, idUsuarioDSK) VALUES ('{0}', '{1}','{2}', '{3}','{4}', '{5}'); SELECT LAST_INSERT_ID() AS idNF;", Format(CDate(cbDataInicial.Text), "yyyy-MM-dd"), Format(CDate(cbDataFinal.Text), "yyyy-MM-dd"), Format(Date.Now, "yyyy-MM-dd"), Format(CDate(cbVencimento.Text), "yyyy-MM-dd"), idGrupo, Usuario_Codigo)
                    idCodNF = MySQL_atualiza(SQL)

                    If idCodNF <> -1 Or idCodNF = 0 Then

                        SQL = String.Format("SELECT * FROM clientes INNER JOIN grupos ON grupos.idGrupo = clientes.idgrupo WHERE grupos.grupo='{0}' AND clientes.bloqueado='0';", cbGrupo.Text)
                        Dim MeusClientes = MySQL_datatable(SQL)

                        For Each r In MeusClientes.Rows

                            Dim MeuCliente = r("idCliente").ToString

                            ' Valor Atual (tabela)
                            SQL = String.Format("SELECT valor_atual, valor_multa, idPagTabela FROM tabela WHERE idPagTabela='{0}';", r("idPagTabela").ToString)
                            Dim ValorCliente = MySQL_consulta_campo(SQL, "valor_atual")
                            Dim ValorMulta = MySQL_consulta_campo(SQL, "valor_multa")
                            Dim ValorBruto = 0.0

                            ' Consulta a notafiscal anterior...
                            SQL = String.Format("SELECT notafiscal_boleto.idCliente, notafiscal_boleto.status, notafiscal_boleto.idNF, notafiscal.periodo_inicial, notafiscal.periodo_final, notafiscal_boleto.valor_bruto AS ValorBruto, notafiscal_boleto.status FROM notafiscal INNER JOIN notafiscal_boleto ON notafiscal.idNF = notafiscal_boleto.idNF WHERE notafiscal_boleto.idNF='{0}' AND notafiscal_boleto.idCliente='{1}' AND (notafiscal_boleto.status='Paga' OR notafiscal_boleto.status='Registrada') ORDER BY notafiscal_boleto.status, notafiscal_boleto.valor_bruto DESC;", idCodNF_Referencia, MeuCliente)

                            Dim ConsultaStatus = MySQL_consulta_campo(SQL, "status")
                            Dim ConsultaValor = MySQL_consulta_campo(SQL, "ValorBruto")

                            If ConsultaStatus = "Registrada" Then
                                ValorBruto = ConsultaValor + ValorCliente + ValorMulta
                            Else
                                ValorBruto = ValorCliente
                            End If

                            Dim ValorBruto_MySQL As String = ValorBruto
                            ValorBruto_MySQL = ValorBruto_MySQL.Replace(",", ".")

                            ' INSERIR (notafiscal_boleto)
                            SQL = String.Format("INSERT INTO notafiscal_boleto (idNF, idCliente, status, tarifa, nroCobranca, valor_bruto) VALUES('{0}', '{1}', 'Registrada', '0', '0', '{2}');", idCodNF, MeuCliente, ValorBruto_MySQL)
                            MySQL_atualiza(SQL)

                        Next

                        Refresh_Lancamento()
                        Refresh_Boletos(idCodNF)
                        Novo_idCodNF = idCodNF

                        MsgBox(String.Format("Nota Fiscal [ {0} ] criada com sucesso!", idCodNF), MsgBoxStyle.Information, "Pronto para emitir!")
                         
                        ' Inicia emissão...
                        bwEmitirBoletos.RunWorkerAsync()

                    Else
                        MsgBox("Não foi possível criar a nota fiscal", MsgBoxStyle.Information, "Informação")
                    End If

                End If

            Else
                MsgBox("Preencher os campos!", MsgBoxStyle.Information, "Informação")

            End If

        End If

    End Sub


    Private Sub viewLancamento_RowClick(sender As Object, e As RowClickEventArgs) Handles viewLancamento.RowClick

        Try
            Dim View As GridView = sender
            idCodNF_Referencia = View.GetRowCellDisplayText(e.RowHandle, View.Columns("CodigoNF"))
            CodigoFechamento = View.GetRowCellDisplayText(e.RowHandle, View.Columns("codigoFechamento"))
            cbGrupo.Text = View.GetRowCellDisplayText(e.RowHandle, View.Columns("grupo")).ToString
            idGrupo = grupoID(cbGrupo.Text)

            'Pega a data, e acrescenta 1 mês.
            Dim Inicio As Date = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Inicio"))
            Dim Final As Date = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Final"))

            cbDataInicial.EditValue = Inicio.AddMonths(1)
            cbDataFinal.EditValue = Final.AddMonths(1)

            Refresh_Boletos(idCodNF_Referencia)

            If CodigoFechamento = 1 Then
                btCriar.Enabled = False             
            Else
                btCriar.Enabled = True      
            End If
        Catch ex As Exception

        End Try


    End Sub

    Private Sub viewLancamento_RowStyle(sender As Object, e As RowStyleEventArgs) Handles viewLancamento.RowStyle

        ' Quem Fechou e NAO Fechou
        '
        If (e.RowHandle >= 0) Then
            Dim View As GridView = sender
            Dim category = View.GetRowCellDisplayText(e.RowHandle, View.Columns("codigoFechamento"))

            If category = "0" Then
                e.Appearance.BackColor = Color.LightGreen
                e.Appearance.BackColor2 = Color.White
            Else
                e.Appearance.BackColor = Color.White
                e.Appearance.BackColor2 = Color.LightYellow
            End If
        End If

    End Sub

    Private Sub bwEmitirBoletos_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwEmitirBoletos.RunWorkerCompleted
        barra.Properties.Paused = True

        MsgBox("Emissão de boletos finalizado!", MsgBoxStyle.Information)

        '... Sinaliza o abertura.
        SQL = String.Format("UPDATE notafiscal SET fechamento='1' WHERE idNF='{0}';", Novo_idCodNF)
        MySQL_atualiza(SQL)

    End Sub

    Private Sub cbGrupo_EditValueChanged(sender As Object, e As EventArgs) Handles cbGrupo.EditValueChanged
          Refresh_Sacados()

    End Sub

    Private Sub viewBoletos_CustomColumnDisplayText(sender As Object, e As CustomColumnDisplayTextEventArgs) Handles viewBoletos.CustomColumnDisplayText

        Try
            Dim view As ColumnView = TryCast(sender, ColumnView)
            If e.Column.FieldName = "ValorBruto" Then
                Dim price As Decimal = Convert.ToDecimal(e.Value)
                e.DisplayText = String.Format(ciBR, "{0:c}", price)
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub viewBoletos_RowStyle(sender As Object, e As RowStyleEventArgs) Handles viewBoletos.RowStyle

        ' Status de cada boleto
        '
        If (e.RowHandle >= 0) Then
            Dim View As GridView = sender
            Dim category = View.GetRowCellDisplayText(e.RowHandle, View.Columns("status"))

            If category = "Paga" Then
                e.Appearance.BackColor = Color.LightGreen
                e.Appearance.BackColor2 = Color.White

            ElseIf category = "Registrada" Then
                e.Appearance.BackColor = Color.White
                e.Appearance.BackColor2 = Color.Yellow

            ElseIf category = "Cancelada" Then
                e.Appearance.BackColor = Color.White
                e.Appearance.BackColor2 = Color.Salmon
            End If

        End If

    End Sub

    Private Sub RepositoryItemCheckEdit2_CheckedChanged(sender As Object, e As EventArgs) Handles RepositoryItemCheckEdit2.CheckedChanged

        '...  Dim View As GridView = sender
        Dim info As GridHitInfo = viewSacados.CalcHitInfo(gridSacados.PointToClient(Cursor.Position))
        CodigoCliente = viewSacados.GetRowCellDisplayText(info.RowHandle, viewSacados.Columns("Codigo"))

        '... Disciplina...
        Dim obj As CheckEdit = sender
        If obj.Checked = True Then
            SQL = String.Format("UPDATE clientes SET bloqueado='1' WHERE idCliente='{0}';", CodigoCliente)
            MySQL_atualiza(SQL)
            Refresh_Sacados()
        Else
            SQL = String.Format("UPDATE clientes SET bloqueado='0' WHERE idCliente='{0}';", CodigoCliente)
            MySQL_atualiza(SQL)
            Refresh_Sacados()
        End If

    End Sub

    Private Sub viewSacados_RowClick(sender As Object, e As RowClickEventArgs) Handles viewSacados.RowClick

        Dim View As GridView = sender
        CodigoCliente = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Codigo"))

    End Sub

  Private Sub viewSacados_RowStyle(sender As Object, e As RowStyleEventArgs) Handles viewSacados.RowStyle

        If (e.RowHandle >= 0) Then
            Dim View As GridView = sender
            Dim category = View.GetRowCellDisplayText(e.RowHandle, View.Columns("CodigoBloqueado"))
            If category = "0" Then
                e.Appearance.BackColor = Color.LightGreen
                e.Appearance.BackColor2 = Color.White
            Else
                e.Appearance.BackColor = Color.LightSalmon
                e.Appearance.BackColor2 = Color.White
            End If
        End If
    End Sub

    Private Sub viewSacados_CustomColumnDisplayText(sender As Object, e As CustomColumnDisplayTextEventArgs) Handles viewSacados.CustomColumnDisplayText

        Try
            Dim view As ColumnView = TryCast(sender, ColumnView)
            If e.Column.FieldName = "Mensalidade" Then
                Dim price As Decimal = Convert.ToDecimal(e.Value)
                e.DisplayText = String.Format(ciBR, "{0:c}", price)
            End If
        Catch ex As Exception
        End Try

    End Sub
End Class