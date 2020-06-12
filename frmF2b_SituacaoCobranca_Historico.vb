Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraGrid.Views.Grid

Public Class frmF2b_SituacaoCobranca_Historico

    Public ClienteAtual As String

    Public Sub New(Cliente As String)

        InitializeComponent()
        Consulta_Historico(Cliente, Nothing, Nothing)

        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Public Function Consulta_Historico(Cliente As String, Inicio As Date, Final As Date)

        If Cliente <> "" Then
            ClienteAtual = Cliente

            If Inicio = Nothing And Final = Nothing Then
                Final = Format(Date.Now, "dd/MM/yyyy")
                Inicio = Date.Now.AddMonths(-3).ToString
            End If

            ' Consulta Histórico do Cliente...
            '
            SQL = String.Format("SELECT clientes.Nome AS Nome, notafiscal.vencimento AS Vencimento, notafiscal_boleto.valor_bruto AS ValorBruto, notafiscal_boleto.`status` AS Status, notafiscal_boleto.nroCobranca AS nroCobranca FROM notafiscal INNER JOIN notafiscal_boleto ON notafiscal_boleto.idNF = notafiscal.idNF INNER JOIN clientes ON notafiscal_boleto.idCliente = clientes.idCliente WHERE clientes.Nome='{0}' AND notafiscal.vencimento BETWEEN ('{1}') AND ('{2}');", ClienteAtual, Format(Inicio, "yyyy-MM-dd"), Format(Final, "yyyy-MM-dd"))
            gridHistorico.DataSource = Nothing
            gridHistorico.DataSource = MySQL_consulta_datagrid(SQL)
            viewHistorico.Columns(0).Width = "250"
            viewHistorico.Columns(1).Width = "70"
            viewHistorico.Columns(2).Width = "80"
            viewHistorico.Columns(3).Width = "80"
            viewHistorico.Columns(4).Width = "80"
        End If

    End Function

    Private Sub frmF2b_SituacaoCobranca_Historico_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Período de 90 dias anteriores...
        '
        dataFim.EditValue = Format(Date.Now, "dd/MM/yyyy").ToString
        dataInicio.EditValue = Date.Now.AddMonths(-6)

        Consulta_Historico(ClienteAtual, Nothing, Nothing)

    End Sub

    Private Sub viewHistorico_RowStyle(sender As Object, e As RowStyleEventArgs) Handles viewHistorico.RowStyle

        ' Status de cada boleto
        '
        If (e.RowHandle >= 0) Then
            Dim View As GridView = sender
            Dim category = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Status"))

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

    Private Sub dataInicio_EditValueChanged(sender As Object, e As EventArgs) Handles dataInicio.EditValueChanged
        Consulta_Historico(ClienteAtual, dataInicio.EditValue, dataFim.EditValue)

    End Sub

    Private Sub dataFim_EditValueChanged(sender As Object, e As EventArgs) Handles dataFim.EditValueChanged
        Consulta_Historico(ClienteAtual, dataInicio.EditValue, dataFim.EditValue)

    End Sub

    Private Sub viewHistorico_CustomColumnDisplayText(sender As Object, e As CustomColumnDisplayTextEventArgs) Handles viewHistorico.CustomColumnDisplayText
        Try
            Dim view As ColumnView = TryCast(sender, ColumnView)
            If e.Column.FieldName = "ValorBruto" Then
                Dim price As Decimal = Convert.ToDecimal(e.Value)
                e.DisplayText = String.Format(ciBR, "{0:c}", price)
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class