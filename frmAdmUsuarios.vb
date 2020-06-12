Imports MySql.Data.MySqlClient
Imports System.DBNull
Imports System.Data.DataTable
Imports System.Data
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraPrinting

Public Class frmAdmUsuarios

    Dim codigoUsuario = 0
    Dim usuarioClicado

    Private Sub frmGerencial2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        RefreshUsuarios()

    End Sub

    Public Sub RefreshUsuarios()

        Try
            SQL = "SELECT Usuario, Bloqueado, idUsuarioDSK AS codigoUsuario, Bloqueado AS codigoBloqueado FROM usuariodsk ORDER BY bloqueado, usuario ASC;"
            gridUsuario.DataSource = Nothing
            gridUsuario.DataSource = MySQL_datatable(SQL)

            RepositoryItemCheckEdit1.ValueChecked = Convert.ToByte(1)
            RepositoryItemCheckEdit1.ValueUnchecked = Convert.ToByte(0)

            viewUsuario.Columns(0).OptionsColumn.AllowEdit = False
            viewUsuario.Columns(1).OptionsColumn.AllowEdit = True
            viewUsuario.Columns(1).ColumnEdit = RepositoryItemCheckEdit1
            viewUsuario.Columns(2).Visible = False
            viewUsuario.Columns(3).Visible = False

        Catch ex As Exception
        End Try

    End Sub

    Private Sub btEfetivar_Click(sender As Object, e As EventArgs) Handles btSalvar.Click

        If tbSenha.Text <> tbSenha2.Text Then
            MsgBox("Ao repetir a senha, ela não confere.", MsgBoxStyle.Information, "Senha")
            tbSenha.SelectAll()
            tbSenha.Focus()

        Else

            'Procura se já existe o usuário. 
            SQL = String.Format("SELECT idUsuarioDSK, usuario FROM usuariodsk WHERE usuario='{0}';", tbUsuario.Text)
            Dim Nome = MySQL_consulta_campo(SQL, "idUsuarioDSK")

            If Nome <> "0" Then
                'Existe...
                If MsgBox(String.Format("Efetuar a alteração: {0}De: [{1}] {0}Para: [{2}] ?", vbCrLf, usuarioClicado, tbUsuario.Text), MsgBoxStyle.YesNo, "Alteração") = DialogResult.Yes Then

                    SQL = String.Format("UPDATE usuariodsk SET senha='{1}' WHERE idUsuarioDSK='{0}';", codigoUsuario, tbSenha.Text)
                    MySQL_atualiza(SQL)
                    tbUsuario.Text = ""
                    tbSenha.Text = ""
                    tbSenha2.Text = ""

                    MsgBox("Sucesso!", MsgBoxStyle.Information, "Alteração")
                    RefreshUsuarios()

                End If

            Else
                'Não Existe...
                If MsgBox(String.Format("Cadastrar o usuário: {0} ?", tbUsuario.Text), MsgBoxStyle.YesNo, "Adicionar usuário") = DialogResult.Yes Then

                    SQL = String.Format("INSERT INTO usuariodsk (usuario, senha, bloqueado) VALUES ('{0}', '{1}', '0');", tbUsuario.Text, tbSenha.Text)
                    MySQL_atualiza(SQL)
                    tbUsuario.Text = ""
                    tbSenha.Text = ""
                    tbSenha2.Text = ""

                    MsgBox("Sucesso!", MsgBoxStyle.Information, "Alteração")
                    RefreshUsuarios()

                End If

            End If
        End If


    End Sub

    Private Sub viewUsuario_RowClick(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles viewUsuario.RowClick

        Dim View As GridView = sender

        codigoUsuario = View.GetRowCellDisplayText(e.RowHandle, View.Columns("codigoUsuario"))
        usuarioClicado = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Usuario"))
        tbUsuario.Text = usuarioClicado

    End Sub

    Private Sub viewUsuario_RowStyle(sender As Object, e As RowStyleEventArgs) Handles viewUsuario.RowStyle

        If (e.RowHandle >= 0) Then
            Dim View As GridView = sender
            Dim category = View.GetRowCellDisplayText(e.RowHandle, View.Columns("codigoBloqueado"))

            If category = "0" Then
                e.Appearance.BackColor = Color.LightGreen
                e.Appearance.BackColor2 = Color.White
            Else
                e.Appearance.BackColor = Color.LightSalmon
                e.Appearance.BackColor2 = Color.White
            End If
        End If

    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click

        frmPreview_Titulo = "Relatório: Gerenciar Usuários"
        Dim Link As New PrintableComponentLink(New PrintingSystem()) With {.Component = gridUsuario}
        AddHandler Link.CreateMarginalHeaderArea, AddressOf frmPreview_Padrao
        Link.CreateDocument()
        Link.ShowPreview()

    End Sub
End Class