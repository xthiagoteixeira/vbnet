Imports System.Net
Public Class frmTeamviewer

    Dim whereToSave As String 'Where the program save the file
    Dim Endereco_Download As String = "http://www.teamviewer.com/TeamViewerQS_pt.exe"
    Delegate Sub ChangeTextsSafe(ByVal length As Long, ByVal position As Integer, ByVal percent As Integer, ByVal speed As Double)
    Delegate Sub DownloadCompleteSafe(ByVal cancelled As Boolean)

    Public Sub DownloadComplete(ByVal cancelled As Boolean)

        If cancelled Then
            Me.Label4.Text = "Cancelado!"
            MessageBox.Show("Download cancelado!", "Cancelado!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Me.Label4.Text = "Baixado com sucesso!"
           ' MessageBox.Show("Successfully downloaded!", "All OK", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Try
                Me.Hide()
                'SE EXISTIR, ELE CHAMA O ARQUIVO ...
                Process.Start(String.Format("{0}\{1}\TeamViewerQS_pt.exe", meucaminhorelatorio, NomeDoProjeto))
                                                
            Catch ex As Exception
              
            End Try

        End If

        Me.ProgressBar1.Value = 0
        Me.Label5.Text = "Baixando: "
       ' Me.Label6.Text = "Save to: "
        Me.Label3.Text = "Tamanho do arquivo: "
        Me.Label2.Text = "Velocidade: "
        Me.Label4.Text = ""

    End Sub

    Public Sub ChangeTexts(ByVal length As Long, ByVal position As Integer, ByVal percent As Integer, ByVal speed As Double)

        Me.Label3.Text = String.Format("Tamanho do arquivo: {0} KB", Math.Round((length / 1024), 2))

        Me.Label5.Text = "BAIXANDO: " & Endereco_Download

        Me.Label4.Text = String.Format("Baixando {0} KB de {1}KB ({2}%)", Math.Round((position / 1024), 2), Math.Round((length / 1024), 2), Me.ProgressBar1.Value)

        If speed = -1 Then
            Me.Label2.Text = "Velocidade: calculando..."
        Else
            Me.Label2.Text = String.Format("Velocidade: {0} KB/s", Math.Round((speed / 1024), 2))
        End If

        Me.ProgressBar1.Value = percent


    End Sub
    
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        'Creating the request and getting the response
        Dim theResponse As HttpWebResponse
        Dim theRequest As HttpWebRequest
        Try 'Checks if the file exist
            
            theRequest = WebRequest.Create(Endereco_Download)
            theResponse = theRequest.GetResponse
        Catch ex As Exception

            MessageBox.Show("Erro ao baixar o arquivo, uma das possíveis causas:" & ControlChars.CrLf & _
                            "1) Arquivo não existe" & ControlChars.CrLf & _
                            "2) Problemas de acesso à internet", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Dim cancelDelegate As New DownloadCompleteSafe(AddressOf DownloadComplete)

            Me.Invoke(cancelDelegate, True)

            Exit Sub
        End Try
        Dim length As Long = theResponse.ContentLength 'Size of the response (in bytes)
        Dim safedelegate As New ChangeTextsSafe(AddressOf ChangeTexts)

        Me.Invoke(safedelegate, length, 0, 0, 0) 'Invoke the TreadsafeDelegate
        Dim writeStream As New IO.FileStream(Me.whereToSave, IO.FileMode.Create)

        'Replacement for Stream.Position (webResponse stream doesn't support seek)
        Dim nRead As Integer

        'To calculate the download speed
        Dim speedtimer As New Stopwatch
        Dim currentspeed As Double = -1
        Dim readings As Integer = 0

        Do

            If BackgroundWorker1.CancellationPending Then 'If user abort download
                Exit Do
            End If

            speedtimer.Start()

            Dim readBytes(4095) As Byte
            Dim bytesread As Integer = theResponse.GetResponseStream.Read(readBytes, 0, 4096)

            nRead += bytesread
            Dim percent As Short = (nRead * 100) / length

            Me.Invoke(safedelegate, length, nRead, percent, currentspeed)

            If bytesread = 0 Then Exit Do

            writeStream.Write(readBytes, 0, bytesread)

            speedtimer.Stop()

            readings += 1
            If readings >= 5 Then 'For increase precision, the speed it's calculated only every five cicles
                currentspeed = 20480 / (speedtimer.ElapsedMilliseconds / 1000)
                speedtimer.Reset()
                readings = 0
            End If
        Loop

        'Close the streams
        theResponse.GetResponseStream.Close()
        writeStream.Close()

        If Me.BackgroundWorker1.CancellationPending Then

            IO.File.Delete(Me.whereToSave)

            Dim cancelDelegate As New DownloadCompleteSafe(AddressOf DownloadComplete)

            Me.Invoke(cancelDelegate, True)

            Exit Sub

        End If

        Dim completeDelegate As New DownloadCompleteSafe(AddressOf DownloadComplete)

        Me.Invoke(completeDelegate, False)

    End Sub

    Private Sub mainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.whereToSave = String.Format("{0}\{1}\TeamViewerQS_pt.exe", meucaminhorelatorio, NomeDoProjeto)
        Me.BackgroundWorker1.RunWorkerAsync() 'Start download

    End Sub

    'Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
    '    Me.BackgroundWorker1.CancelAsync() 'Send cancel request
    'End Sub
     
End Class
