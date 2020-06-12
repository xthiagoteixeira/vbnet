Imports System.Web.Services
Imports System.ComponentModel
Imports System.Net

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Delib11
    Inherits System.Web.Services.WebService

    ' TRAZER E BAIXA...
    '    
    <WebMethod()> Function trazerFicha(UsuarioWeb As String, SenhaWeb As String, Escola As String, AnoVigente As String, CodigoTurma As String) As DataSet

        System.Net.ServicePointManager.Expect100Continue = False
        Dim FichaAluno
       
        SQL = "SELECT " _
            & "delib11_nf.anovigente, " _
            & "delib11_nf.turma, " _
            & "delib11_nf.nro_aluno " _
            & "FROM " _
            & "delib11_nf " _
            & "WHERE delib11_nf.anovigente='" & AnoVigente & "' AND delib11_nf.turma='" & CodigoTurma & "' AND delib11_nf.sincronizado='1' ORDER BY delib11_nf.anovigente, delib11_nf.turma, delib11_nf.nro_aluno;"
        FichaAluno = MySQL_consulta_tabela(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))
      
       Return FichaAluno

    End Function

    ' TRAZER E BAIXA...
    '    
    <WebMethod()> Function trazerFicha2(UsuarioWeb As String, SenhaWeb As String, Escola As String, AnoVigente As String, CodigoTurma As String, NroAluno As String) As DataSet

        System.Net.ServicePointManager.Expect100Continue = False
        Dim FichaAluno

        SQL = "SELECT " _
            & "delib11_nf.anovigente, " _
            & "delib11_nf.turma, " _
            & "delib11_nf.nro_aluno, " _
            & "delib11_boletim.idDelib_cat, " _
            & "delib11_boletim.idDelib_des, " _
            & "delib11_nf.sincronizado " _
            & "FROM " _
            & "delib11_nf " _
            & "INNER JOIN delib11_boletim ON delib11_nf.idDelibNF = delib11_boletim.idDelibNF WHERE delib11_nf.anovigente='" & AnoVigente & "' AND delib11_nf.turma='" & CodigoTurma & "' AND delib11_nf.nro_aluno='" & NroAluno & "' AND delib11_nf.sincronizado='1' ORDER BY delib11_nf.anovigente, delib11_nf.turma, delib11_nf.nro_aluno;"
        FichaAluno = MySQL_consulta_tabela(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))

        Try
            ' Da Baixa no SERVIDOR...
            SQL = String.Format("UPDATE delib11_nf SET sincronizado='0' WHERE anovigente='{0}' AND turma='{1}' AND nro_aluno='{2}';", AnoVigente, CodigoTurma, NroAluno)
            MySQL_atualiza(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))
        Catch ex As Exception
        End Try

        Return FichaAluno

    End Function

    ' CONTAR Deliberações (nro delib)...
    '
    <WebMethod()> Function trazerNro(UsuarioWeb As String, SenhaWeb As String, Escola As String, AnoVigente As String) As String

        System.Net.ServicePointManager.Expect100Continue = False
        Dim Nro

        Try
            SQL = String.Format("SELECT COUNT(*) AS nro FROM delib11_nf WHERE sincronizado='1' AND anovigente='{0}';", AnoVigente)
            Nro = MySQL_consulta_campo(SQL, "nro", String.Format("{0}database={1}", CONEXAOBD, Escola))
        Catch ex As Exception
            Nro = "0"
        End Try

        Return Nro

    End Function


    ' '' TRAZER NOTAS - CONSULTAR (boletim)...
    ' ''
    ' ''
    ''<WebMethod()> Function consultarNF(UsuarioWeb As String, SenhaWeb As String, Escola As String, AnoVigente As String, Bimestre As String) As DataSet

    ''    ServicePointManager.Expect100Continue = False
    ''    Dim Boletim

    ''    Try

    ''        SQL = "SELECT " _
    ''            & "notasfreq.cod_nf AS cod_nf, " _
    ''            & "notasfreq.cod_bimestre AS cod_bimestre, " _
    ''            & "notasfreq.turma AS CodigoTurma, " _
    ''            & "notasfreq.disciplina AS CodigoDisciplina, " _
    ''            & "disciplinas.disciplina As disciplina, " _
    ''            & "turma.classe AS classe, " _
    ''            & "notasfreq.anovigente AS AnoVigente, " _
    ''            & "notasfreq.sincronizado_delib11 " _
    ''            & "FROM " _
    ''            & "notasfreq " _
    ''            & "INNER JOIN turma ON turma.codigo_trma = notasfreq.turma " _
    ''            & "INNER JOIN disciplinas ON disciplinas.codigo_disc = notasfreq.disciplina " _
    ''            & "WHERE notasfreq.anovigente='" & AnoVigente & "' AND notasfreq.cod_bimestre='" & Bimestre & "' AND notasfreq.sincronizado_delib11='1';"
    ''        Boletim = MySQL_consulta_tabela(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))

    ''    Catch ex As Exception
    ''    End Try


    ''    Return Boletim

    ''End Function
    ' DELIBERAÇÃO 11 - CONSULTAR (quem tá com nota vermelha)...

    <WebMethod()> Function GradeTurma(UsuarioWeb As String, SenhaWeb As String, Escola As String, AnoVigente As String, Bimestre As String, Turma As String) As DataSet

        ServicePointManager.Expect100Continue = False
        SQL = "SELECT " _
                & "disciplinas.disciplina AS Disciplina, " _
                & "boletim.nro_aluno AS Nro, " _
                & "aluno.nome AS Nome, " _
                & "boletim.M, " _
                & "boletim.F, turma.codigo_trma AS NroTurma " _
                & "FROM " _
                & "usuariodsk " _
                & "INNER JOIN professor_grupos ON usuariodsk.codigo = professor_grupos.idUsuarioDSK " _
                & "INNER JOIN notasfreq ON notasfreq.disciplina = professor_grupos.codigo_disc AND professor_grupos.codigo_trma = notasfreq.turma " _
                & "INNER JOIN boletim ON boletim.cod_boletim = notasfreq.cod_nf " _
                & "INNER JOIN aluno ON aluno.turma = notasfreq.turma AND aluno.anovigente = notasfreq.anovigente AND aluno.nro = boletim.nro_aluno " _
                & "INNER JOIN turma ON professor_grupos.codigo_trma = turma.codigo_trma " _
                & "INNER JOIN disciplinas ON professor_grupos.codigo_disc = disciplinas.codigo_disc " _
                & "WHERE " _
                & "notasfreq.anovigente='" & AnoVigente & "' AND  " _
                & "notasfreq.cod_bimestre='" & Bimestre & "' AND  " _
                & "turma.classe='" & Turma & "' AND  " _
                & "usuariodsk.usuario = '" & UsuarioWeb & "' AND " _
                & "boletim.M<5 ORDER BY disciplinas.disciplina, boletim.nro_aluno ASC;"

        Dim boletim = MySQL_consulta_tabela(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))
        Return boletim

    End Function

    ' DELIBERAÇÃO 11 - CONSULTAR (quem tá com nota vermelha)...

    <WebMethod()> Function idDelibNF(UsuarioWeb As String, SenhaWeb As String, Escola As String, AnoVigente As String, Turma As String, NroAluno As String) As String
        
        ServicePointManager.Expect100Continue = False

        'Primeiro Consulta se Existe...
        SQL = "SELECT idDelibNF FROM delib11_nf WHERE anovigente='" & AnoVigente & "' AND turma='" & Turma & "' AND nro_aluno='" & NroAluno & "';"
        Dim Retorno = MySQL_consulta_campo(SQL, "idDelibNF", String.Format("{0}database={1}", CONEXAOBD, Escola))
        
        If Retorno = "0" Then
            'Se não existir... cria uma idDelibNF...
            SQL = "INSERT INTO delib11_nf (anovigente, turma, nro_aluno) VALUES('" & AnoVigente & "','" & Turma & "','" & NroAluno & "'); SELECT LAST_INSERT_ID() AS idDelibNF;"
            Retorno = MySQL_atualiza(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))
        End If
        
        Return Retorno

    End Function

    ' DELIBERAÇÃO 11 - OPCOES (consulta categoria)...
    '
    '
    <WebMethod()> Function consultaOpcoes(UsuarioWeb As String, SenhaWeb As String, Escola As String, Categoria As String) As DataSet

        ServicePointManager.Expect100Continue = False
        Dim Deb11
        Try
            SQL = String.Format("SELECT delib11_categoria.idDelib_cat, delib11_categoria.categoria, delib11_descricao.idDelib_des, delib11_descricao.descricao FROM delib11_categoria INNER JOIN delib11_descricao ON delib11_categoria.idDelib_cat = delib11_descricao.idDelib_cat WHERE delib11_categoria.categoria='{0}';", Categoria)
            Deb11 = MySQL_consulta_tabela(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))
        Catch ex As Exception
        End Try

        Return Deb11

    End Function

    ' DELIBERAÇÃO 11 - Ficha_Aluno (delib11_notasfreq - consulta descrição já cadastrada)...
    '
    <WebMethod()> Function consultaFicha(UsuarioWeb As String, SenhaWeb As String, Escola As String, idDelib_cat As Integer, idDelib_des As Integer, notasfreq As Integer, nro_aluno As Integer) As String

        ServicePointManager.Expect100Continue = False
        Dim retorno

        Try

            SQL = "SELECT iddelib_des FROM delib11_boletim WHERE iddelibnf='" & notasfreq & "' AND iddelib_cat='" & idDelib_cat & "' AND iddelib_des='" & idDelib_des & "';"
            retorno = MySQL_consulta_campo(SQL, "iddelib_des", String.Format("{0}database={1}", CONEXAOBD, Escola))

        Catch ex As Exception
        End Try

        Return retorno

    End Function

    ' DELIBERAÇÃO 11 - Ficha_Aluno (delib11_notasfreq)...
    '
    <WebMethod()> Function preencherFicha(UsuarioWeb As String, SenhaWeb As String, Escola As String, idDelib_cat As Integer, idDelib_des As Integer, notasfreq As Integer, nro_aluno As Integer, Opcao_Exclusiva As Integer) As String

        '... Opcao Exclusiva!
        ' -> 1 (incluir)
        ' -> 2 (excluir)
        '
        ServicePointManager.Expect100Continue = False
        Dim retorno = "Não foi possível executar esta operação!"

        Try

            If Opcao_Exclusiva = 1 Then

                'incluir...
                '
                SQL = String.Format("INSERT INTO delib11_boletim (iddelib_cat, iddelib_des, iddelibnf) VALUES ('{0}', '{1}', '{2}');", idDelib_cat, idDelib_des, notasfreq)
                Dim Ret = MySQL_atualiza(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))

            ElseIf Opcao_Exclusiva = 2 Then

                'excluir...
                '
                SQL = String.Format("DELETE FROM delib11_boletim WHERE iddelibnf='{0}' AND iddelib_cat='{1}' AND iddelib_des='{2}';", notasfreq, idDelib_cat, idDelib_des)
                Dim Ret = MySQL_atualiza(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))

            Else
                retorno = "Não foi possível executar esta operação!"
            End If

            ' NECESSARIO NOTIFICAR NOTASFREQ (notasfreq - delib11_notasfreq)
            '
            SQL = String.Format("UPDATE delib11_nf SET sincronizado='1' WHERE iddelibnf={0}", notasfreq)
            MySQL_atualiza(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))

        Catch ex As Exception
        End Try

        Return retorno

    End Function

    ' DELIBERAÇÃO 11 - CATEGORIAS GERAIS ()...
    '
    '
    <WebMethod()> Function consultaCategoria(UsuarioWeb As String, SenhaWeb As String, Escola As String) As DataSet

        ServicePointManager.Expect100Continue = False
        Dim Deb11

        Try

            SQL = "SELECT " _
                & "delib11_categoria.idDelib_cat, " _
                & "delib11_categoria.categoria " _
                & "FROM " _
                & "delib11_categoria;"
            Deb11 = MySQL_consulta_tabela(SQL, String.Format("{0}database={1}", CONEXAOBD, Escola))

        Catch ex As Exception
        End Try

        Return Deb11

    End Function

    '**************************** TRAZER deliberação 11 (no programa)***************************
    '
    

End Class