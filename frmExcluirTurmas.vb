Public Class frmExcluirTurmas

    Dim codigoConsulta = 0
    Dim nroTurma

    Sub RefreshDataGrid()

        SQL = "SELECT notasfreq.anovigente AS Ano, notasfreq.cod_bimestre AS Bimestre, " _
           & "turma.classe AS Turma, " _
           & "disciplinas.disciplina AS Disciplina " _
           & "FROM " _
           & "notasfreq " _
           & "INNER JOIN disciplinas ON notasfreq.disciplina = disciplinas.codigo_disc " _
           & "INNER JOIN turma ON notasfreq.turma = turma.codigo_trma " _
           & "WHERE turma.codigo_trma='" & nroTurma & "' " _
           & "ORDER BY notasfreq.anovigente, notasfreq.cod_bimestre, turma.classe, disciplinas.disciplina ASC;"
        gridBoletim.DataSource = Nothing
        gridBoletim.DataSource = MySQL_datatable(SQL)

    End Sub

    Sub RefreshTurmas(Bloqueadas As Boolean)

        If Bloqueadas = True Then
            SQL = "SELECT classe FROM turma WHERE bloqueado='1' ORDER BY classe ASC;"
            cbTurma.Properties.ValueMember = "classe"
            cbTurma.Properties.DisplayMember = "classe"
            cbTurma.Properties.DataSource = MySQL_datatable(SQL)
        Else
            SQL = "SELECT classe FROM turma WHERE bloqueado='0' ORDER BY classe ASC;"
            cbTurma.Properties.ValueMember = "classe"
            cbTurma.Properties.DisplayMember = "classe"
            cbTurma.Properties.DataSource = MySQL_datatable(SQL)
        End If

    End Sub

    Private Sub btExcluir_Click(sender As Object, e As EventArgs) Handles btExcluir.Click

        SQL = "DELETE FROM turma WHERE codigo_trma='" & nroTurma & "';"
        MySQL_atualiza(SQL)

        MsgBox("Retirado: " & nroTurma & " - " & cbTurma.Text, MsgBoxStyle.Information, "Gerenciar Turmas")

        If grupoTurmas.SelectedIndex = 0 Then
            'Refresh Liberadas
            RefreshTurmas(False)
            codigoConsulta = 0
        ElseIf grupoTurmas.SelectedIndex = 1 Then
            'Refresh Bloqueadas
            RefreshTurmas(True)
            codigoConsulta = 1
        End If

    End Sub

    Private Sub RadioGroup1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles grupoTurmas.SelectedIndexChanged

        If grupoTurmas.SelectedIndex = 0 Then
            'Refresh Liberadas
            RefreshTurmas(False)
            codigoConsulta = 0
        ElseIf grupoTurmas.SelectedIndex = 1 Then
            'Refresh Bloqueadas
            RefreshTurmas(True)
            codigoConsulta = 1
        End If

    End Sub

    Private Sub cbTurma_TextChanged(sender As Object, e As EventArgs) Handles cbTurma.TextChanged
        nroTurma = Consulta_Turma(cbTurma.Text, codigoConsulta)
        RefreshDataGrid()

    End Sub
End Class