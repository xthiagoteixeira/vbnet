Imports System.Runtime.InteropServices

Public Class frmSuporteRapido

    Public Const MOD_ALT As Integer = &H0 'Alt key
    Public Const WM_HOTKEY As Integer = &H312

    <DllImport("User32.dll")> _
    Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr, _
                        ByVal id As Integer, ByVal fsModifiers As Integer, _
                        ByVal vk As Integer) As Integer
    End Function

    <DllImport("User32.dll")> _
    Public Shared Function UnregisterHotKey(ByVal hwnd As IntPtr, _
                        ByVal id As Integer) As Integer
    End Function

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam.ToInt32
            Select Case (id.ToString)
                Case "100"
                    'Pressionou botão... abre o form...
                    Me.Visible = True

                    'Me.Size = New System.Drawing.Size(347, 202)
                    Me.WindowState = FormWindowState.Normal
            End Select
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub btTeamviewer_Click(sender As Object, e As EventArgs) Handles btTeamviewer.Click

        Try
            'SE EXISTIR, ELE CHAMA O ARQUIVO ...
            Process.Start(String.Format("{0}\" & NomeDoProjeto & "\TeamViewerQS_pt.exe", meucaminhorelatorio))
        Catch ex As Exception
            'SE NAO EXISTIR, ELE BAIXA...
            Me.Hide()
            frmTeamViewer.Show()
        End Try
             
    End Sub

    Private Sub btManutencao_Click(sender As Object, e As EventArgs) Handles btManutencao.Click
        Me.Hide()
        frmManutencao.Show()
        
    End Sub

    Private Sub frmSuporteRapido_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '#################################### VERIFICA O PROCESSO (MAIS ESCOLA) SE JÁ ESTÁ ATIVADO #################################
        Dim sb2() As Process     ' Gera um array de processos
        sb2 = Process.GetProcessesByName("Mais Escola")
        If sb2.Length = 0 Then  ' Se o tamanho do array for = 0 quer dizer que o processo NAO está ativo

            'ABRE O PROGRAMA
            Me.WindowState = FormWindowState.Normal

        End If
        '############################################## FIM DE VERIFICACAO ###########################################

        RegisterHotKey(Me.Handle, 100, MOD_ALT, Keys.F10)
        '   RegisterHotKey(Me.Handle, 200, MOD_ALT, Keys.F12)

    End Sub

    Private Sub frmSuporteRapido_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '#################################### VERIFICA O PROCESSO SE JÁ ESTÁ ATIVADO #################################
        Dim sb() As Process     ' Gera um array de processos
        sb = Process.GetProcessesByName("Mais Escola")
        If sb.Length = 0 Then  ' Se o tamanho do array for = 0 quer dizer que o processo NAO está ativo
            End
        Else
            Me.Visible = False
            e.Cancel = True
        End If
        '############################################## FIM DE VERIFICACAO ###########################################
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs)
        ' Abre navegador pra pedir ajuda
        Try
            Process.Start("BrowserDemo.exe")
        Catch ex As Exception
        End Try

    End Sub
End Class
