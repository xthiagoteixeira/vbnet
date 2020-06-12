Imports DevExpress.XtraEditors
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraPrinting

Public Class frmAdmTurma
    Dim nroturma_clicada = 0
    Dim nometurma_clicada = ""

    Public Sub RefreshTurmas()

        try
            SQL =
                "SELECT Classe, Periodo, Ensino, Bloqueado, codigo_trma AS Codigo, Bloqueado AS CodigoBloqueado FROM turma ORDER BY Bloqueado, Classe ASC;"
            gridTurmas.DataSource = Nothing
            gridTurmas.DataSource = MySQL_datatable(SQL)

            RepositoryItemCheckEdit1.ValueChecked = Convert.ToByte(1)
            RepositoryItemCheckEdit1.ValueUnchecked = Convert.ToByte(0)

            viewTurmas.Columns(0).OptionsColumn.AllowEdit = False
            viewTurmas.Columns(1).OptionsColumn.AllowEdit = True
            viewTurmas.Columns(2).OptionsColumn.AllowEdit = True
            viewTurmas.Columns(3).OptionsColumn.AllowEdit = True

            viewTurmas.Columns(1).ColumnEdit = RepositoryItemComboBox1

            ' Preencher o Periodo...
            RepositoryItemComboBox1.Items.Clear()
            RepositoryItemComboBox1.Items.Add("Manhã")
            RepositoryItemComboBox1.Items.Add("Tarde")
            RepositoryItemComboBox1.Items.Add("Noite")
            RepositoryItemComboBox1.Items.Add("Integral")
            RepositoryItemComboBox1.Items.Add("Intermediário")
            RepositoryItemComboBox1.Items.Add("Vespertino")

            viewTurmas.Columns(2).ColumnEdit = RepositoryItemComboBox2

            ' Preencher Ensino...
            RepositoryItemComboBox2.Items.Clear()
            RepositoryItemComboBox2.Items.Add("EJA")
            RepositoryItemComboBox2.Items.Add("Fundamental")
            RepositoryItemComboBox2.Items.Add("Integral")
            RepositoryItemComboBox2.Items.Add("Médio")
            RepositoryItemComboBox2.Items.Add("Técnico")
            RepositoryItemComboBox2.Items.Add("Superior")

            viewTurmas.Columns(3).ColumnEdit = RepositoryItemCheckEdit1
            viewTurmas.Columns(4).Visible = False
            viewTurmas.Columns(5).Visible = False
        Catch ex As Exception

        End Try
        
    End Sub

    Private Sub frmManTurma_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Erro_Form = "frmAdmTurma"
        RefreshTurmas()
    End Sub

    Private Sub btCadastrar_Click(sender As Object, e As EventArgs) Handles btCadastrar.Click

        If Not (tbClasse.Text = "" Or cbPeriodo.Text = "" Or cbEnsino.Text = "" Or cbPeriodo.Text = "") Then

            tbClasse.Text = Trim(tbClasse.Text)

            If _
                MsgBox(String.Format("Efetuar o cadastro de: {0} ?", tbClasse.Text), MsgBoxStyle.YesNo, "Alteração") =
                DialogResult.No Then
                MsgBox("Operação Cancelada", MsgBoxStyle.Information, "Cancelada")
                Exit Sub

            Else

                Dim SQL2 = String.Format("SELECT classe FROM turma WHERE classe LIKE '{0}';", tbClasse.Text)
                Dim Retorno = MySQL_consulta_campo(SQL2, "classe")

                If Retorno <> "0" Then
                    MsgBox("Já existe uma turma cadastrada!", MsgBoxStyle.Information, "Informação")
                    Exit Sub
                Else

                    Dim SQL =
                            String.Format(
                                "INSERT INTO turma (classe, periodo, bloqueado, ensino) VALUES ('{0}', '{1}', '0', '{2}');",
                                tbClasse.Text, cbPeriodo.Text, cbEnsino.Text)
                    Dim Retorno2 = MySQL_atualiza(SQL)

                    'arquivoLog("Administrativo", "Cadastrou a turma: " & tbClasse.Text)

                    ' LIBERA NO boletimweb, PARA SINCRONIZAR...
                    SQL = "UPDATE boletimweb SET sincronizado='1' WHERE tabela='turma';"
                    MySQL_atualiza(SQL)

                    If Retorno2 = "0" Then
                        MsgBox("Turma não foi cadastrada!", MsgBoxStyle.Information, "Cadastro!")
                    Else
                        MsgBox("Turma inserida com sucesso!", MsgBoxStyle.Information, "Cadastro!")
                    End If
                    RefreshTurmas()

                End If
            End If
        Else
            MsgBox("Preencher os campos", MsgBoxStyle.Information, "Informação")
        End If

        tbClasse.Text = ""
        cbPeriodo.Text = ""
    End Sub

    Private Sub viewTurmas_RowClick(sender As Object, e As RowClickEventArgs) Handles viewTurmas.RowClick

        Dim View As GridView = sender
        nroturma_clicada = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Codigo"))
        nometurma_clicada = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Classe"))
        tbClasse.Text = nometurma_clicada

        cbPeriodo.Text = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Periodo"))
        cbEnsino.Text = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Ensino"))
    End Sub

    Private Sub viewTurmas_RowStyle(sender As Object, e As RowStyleEventArgs) Handles viewTurmas.RowStyle

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

    Private Sub RepositoryItemCheckEdit1_CheckedChanged(sender As Object, e As EventArgs) _
        Handles RepositoryItemCheckEdit1.CheckedChanged

        viewTurmas.PostEditor()

        Dim info As GridHitInfo = viewTurmas.CalcHitInfo(gridTurmas.PointToClient(Cursor.Position))
        nometurma_clicada = viewTurmas.GetRowCellDisplayText(info.RowHandle, viewTurmas.Columns("Classe"))
        nroturma_clicada = viewTurmas.GetRowCellDisplayText(info.RowHandle, viewTurmas.Columns("Codigo"))

        Dim obj As CheckEdit = sender
        If obj.Checked = True Then

            SQL = String.Format("UPDATE turma SET bloqueado='1' WHERE codigo_trma='{0}';", nroturma_clicada)
            MySQL_atualiza(SQL)

            'arquivoLog("Administrativo", "Bloqueou a turma: " & nometurma_clicada)
        Else

            SQL = String.Format("UPDATE turma SET bloqueado='0' WHERE codigo_trma='{0}';", nroturma_clicada)
            MySQL_atualiza(SQL)

            'arquivoLog("Administrativo", "Desbloqueou a turma: " & nometurma_clicada)
        End If

        ' LIBERA NO boletimweb, PARA SINCRONIZAR...
        SQL = "UPDATE boletimweb SET sincronizado='1' WHERE tabela='turma';"
        MySQL_atualiza(SQL)
        RefreshTurmas()
    End Sub

    Private Sub RepositoryItemComboBox1_Click(sender As Object, e As EventArgs) Handles RepositoryItemComboBox1.Click
        viewTurmas.PostEditor()

        Dim info As GridHitInfo = viewTurmas.CalcHitInfo(gridTurmas.PointToClient(Cursor.Position))
        nroturma_clicada = viewTurmas.GetRowCellDisplayText(info.RowHandle, viewTurmas.Columns("Codigo"))
    End Sub

    Private Sub RepositoryItemComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) _
        Handles RepositoryItemComboBox1.KeyPress
        e.Handled = True
    End Sub

    Private Sub RepositoryItemComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles RepositoryItemComboBox1.SelectedIndexChanged
        ' Periodo...
        '
        Dim edit As ComboBoxEdit = sender
        SQL = String.Format("UPDATE turma SET periodo='{0}' WHERE codigo_trma='{1}';", edit.EditValue, nroturma_clicada)
        MySQL_atualiza(SQL)
    End Sub

    Private Sub RepositoryItemComboBox2_Click(sender As Object, e As EventArgs) Handles RepositoryItemComboBox2.Click
        viewTurmas.PostEditor()

        Dim info As GridHitInfo = viewTurmas.CalcHitInfo(gridTurmas.PointToClient(Cursor.Position))
        nroturma_clicada = viewTurmas.GetRowCellDisplayText(info.RowHandle, viewTurmas.Columns("Codigo"))
    End Sub

    Private Sub RepositoryItemComboBox2_KeyPress(sender As Object, e As KeyPressEventArgs) _
        Handles RepositoryItemComboBox2.KeyPress
        e.Handled = True
    End Sub

    Private Sub RepositoryItemComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles RepositoryItemComboBox2.SelectedIndexChanged
        ' Ensino...
        '
        Dim edit As ComboBoxEdit = sender
        SQL = String.Format("UPDATE turma SET ensino='{0}' WHERE codigo_trma='{1}';", edit.EditValue, nroturma_clicada)
        MySQL_atualiza(SQL)
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub cbPeriodo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cbPeriodo.KeyPress
        e.Handled = True
    End Sub

    Private Sub cbEnsino_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cbEnsino.KeyPress
        e.Handled = True
    End Sub

    Private Sub btExcluir_Click(sender As Object, e As EventArgs) Handles btExcluir.Click

        If _
            MsgBox(String.Format("Excluir a turma: {0} ?", nometurma_clicada), MsgBoxStyle.YesNo, "Excluir") =
            DialogResult.No Then

            MsgBox("Operação Cancelada", MsgBoxStyle.Information, "Cancelada")
            Exit Sub
        Else

            ' Verifica se EXISTE BOLETINS cadastrados...
            Dim SQL2 =
                    String.Format(
                        "SELECT COUNT(cod_nf) AS NroBoletins FROM notasfreq WHERE turma='" & nroturma_clicada & "';")
            Dim Retorno = MySQL_consulta_campo(SQL2, "NroBoletins")

            If Retorno <> "0" Then
                MsgBox("Existem boletins cadastrados com esta turma!", MsgBoxStyle.Information, "Operação Cancelada")
                Exit Sub
            Else

                SQL = String.Format("DELETE FROM turma WHERE codigo_trma='{0}';", nroturma_clicada)
                MySQL_atualiza(SQL)

                tbClasse.Text = ""
                cbPeriodo.Text = ""
                cbEnsino.Text = ""

                ' LIBERA NO boletimweb, PARA SINCRONIZAR...
                SQL = "UPDATE boletimweb SET sincronizado='1' WHERE tabela='turma';"
                MySQL_atualiza(SQL)

                MsgBox("Turma excluída com sucesso!", MsgBoxStyle.Information, "Excluir")
                RefreshTurmas()

                'arquivoLog("Administrativo", "Excluiu a turma: " & nometurma_clicada)

            End If

        End If
    End Sub

    Private Sub btImprimir_Click(sender As Object, e As EventArgs) Handles btImprimir.Click

        'frmPreview_Titulo = "Relatório: Gerenciar Turmas"
        'Dim Link As New PrintableComponentLink(New PrintingSystem()) With {.Component = gridTurmas}
        'AddHandler Link.CreateMarginalHeaderArea, AddressOf frmPreview_Padrao
        'Link.CreateDocument()
        'Link.ShowPreview()

    End Sub

    Private Sub btAlterar_Click(sender As Object, e As EventArgs) Handles btAlterar.Click
        If _
         MsgBox(String.Format("Alterar a turma: {0} para: {1}?", nometurma_clicada, tbClasse.Text), MsgBoxStyle.YesNo, "Excluir") =
         DialogResult.No Then

            MsgBox("Operação Cancelada", MsgBoxStyle.Information, "Cancelada")
            Exit Sub
        Else
                     
                SQL = String.Format("UPDATE turma SET classe='{1}' WHERE codigo_trma='{0}';", nroturma_clicada, tbClasse.Text)
                MySQL_atualiza(SQL)

                tbClasse.Text = ""
                cbPeriodo.Text = ""
                cbEnsino.Text = ""

                ' LIBERA NO boletimweb, PARA SINCRONIZAR...
                SQL = "UPDATE boletimweb SET sincronizado='1' WHERE tabela='turma';"
                MySQL_atualiza(SQL)

                MsgBox("Turma alterada com sucesso!", MsgBoxStyle.Information, "Alteração")
                RefreshTurmas()

                'arquivoLog("Administrativo", "Excluiu a turma: " & nometurma_clicada)
                       

        End If
    End Sub
End Class