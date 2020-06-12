Imports DevExpress.XtraPrinting

Public Class frmRpt

    Private Sub frmRptRelatorio_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Efetiva a pesquisa do relatório
        Try
            meuRPT_Parametro = cbRevenda.EditValue
            meuRPT_DataFinal = Format(CDate(dataFinal.EditValue), "yyyy-MM-dd")
            meuRPT_DataInicio = Format(CDate(dataInicio.EditValue), "yyyy-MM-dd")

            Dim SQL = meuRPT(cbRelatorio.EditValue, cbEstilo.EditValue)

            reportDocument_dataset = MySQL_dataset(SQL)
            reportDocument.DataSource = reportDocument_dataset
            reportDocument.DataMember = reportDocument_dataset.Tables(0).TableName

            DocumentViewer1.DocumentSource = reportDocument
            reportDocument.CreateDocument()

        Catch ex As Exception
            MsgBox("Consulta não encontrada!" & vbCrLf & "Mensagem: " & ex.Message, MsgBoxStyle.Information, "Relatório!")
        End Try

    End Sub

End Class