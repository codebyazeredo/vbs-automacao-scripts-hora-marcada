Option Explicit

Const DIAS_RETENCAO_LOGS = 60
Const INTERVALO_POLL_MS = 2000

Dim FSO
Dim WshShell
Dim pastaBase
Dim pastaBat
Dim pastaLog
Dim pastaLock
Dim pastaTmp
Dim arquivoConfigPath
Dim arquivoEstado
Dim arquivoConfigHandle
Dim estado
Dim linha
Dim partes
Dim nomeBat
Dim listaHorarios
Dim horarios
Dim horarioAtual
Dim caminhoBat
Dim caminhoLogDiario
Dim caminhoLogErroDiario
Dim nomeLogDiario
Dim nomeLogErroDiario
Dim houveErro
Dim executouAlgumBat
Dim aspas
Dim logBufferNormal
Dim logBufferErro
Dim i
Dim chaveEstado
Dim hojeStr
Dim minutosAgora
Dim minutosAlvo
Dim jobsAtivos

Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set jobsAtivos = CreateObject("Scripting.Dictionary")

aspas = Chr(34)
logBufferNormal = ""
logBufferErro = ""

pastaBase = FSO.GetParentFolderName(WScript.ScriptFullName)
pastaBat = pastaBase & "\bat"
pastaLog = pastaBase & "\logs"
pastaLock = pastaBase & "\locks"
pastaTmp = pastaBase & "\tmp"

arquivoConfigPath = pastaBase & "\config_horario.ini"
arquivoEstado = pastaBase & "\estado_horario.ini"

If Not FSO.FolderExists(pastaBat) Then
    WScript.Echo "ERRO: Pasta bat nao encontrada: " & pastaBat
    WScript.Quit 1
End If

If Not FSO.FolderExists(pastaLog) Then FSO.CreateFolder pastaLog
If Not FSO.FolderExists(pastaLock) Then FSO.CreateFolder pastaLock
If Not FSO.FolderExists(pastaTmp) Then FSO.CreateFolder pastaTmp

If Not FSO.FileExists(arquivoConfigPath) Then
    WScript.Echo "ERRO: config_horario.ini nao encontrado: " & arquivoConfigPath
    WScript.Quit 1
End If

LimparLogsAntigos

Set estado = CarregarEstado()

houveErro = False
executouAlgumBat = False
hojeStr = ObterHojeStr()
minutosAgora = Hour(Now) * 60 + Minute(Now)

nomeLogDiario = "execucao_" & Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & ".log"
nomeLogErroDiario = "erros_" & Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & ".log"

caminhoLogDiario = pastaLog & "\" & nomeLogDiario
caminhoLogErroDiario = pastaLog & "\" & nomeLogErroDiario

Log "================================================================"
Log "INÍCIO DA EXECUÇÃO: " & FormatDateTime(Now, 0)
Log "================================================================"

LogError "================================================================"
LogError "INÍCIO DOS REGISTROS DE ERRO: " & FormatDateTime(Now, 0)
LogError "================================================================"

Set arquivoConfigHandle = FSO.OpenTextFile(arquivoConfigPath, 1, False)

Do Until arquivoConfigHandle.AtEndOfStream
    linha = Trim(arquivoConfigHandle.ReadLine)
    If linha <> "" Then
        If Left(linha, 1) <> "#" And Left(linha, 1) <> ";" Then
            partes = Split(linha, "=")
            If UBound(partes) >= 1 Then
                nomeBat = Trim(partes(0))
                listaHorarios = Trim(partes(1))
                If nomeBat <> "" And listaHorarios <> "" Then
                    caminhoBat = pastaBat & "\" & nomeBat
                    If FSO.FileExists(caminhoBat) Then
                        horarios = Split(listaHorarios, ",")
                        For i = 0 To UBound(horarios)
                            horarioAtual = Trim(horarios(i))
                            If ValidarHorario(horarioAtual) Then
                                minutosAlvo = MinutosDoHorario(horarioAtual)
                                chaveEstado = nomeBat & "|" & horarioAtual
                                If minutosAgora >= minutosAlvo Then
                                    If Not estado.Exists(chaveEstado) Or estado(chaveEstado) <> hojeStr Then
                                        DispararJob nomeBat, caminhoBat, horarioAtual, chaveEstado
                                    End If
                                End If
                            Else
                                houveErro = True
                                LogError "[" & FormatDateTime(Now, 3) & "] Horário inválido para " & nomeBat & ": " & horarioAtual
                                LogError ""
                            End If
                        Next
                    Else
                        houveErro = True
                        LogError "[" & FormatDateTime(Now, 3) & "] " & nomeBat
                        LogError "  Status: ERRO | Arquivo BAT não encontrado em: " & caminhoBat
                        LogError ""
                    End If
                Else
                    houveErro = True
                    LogError "[" & FormatDateTime(Now, 3) & "] Configuração inválida no arquivo INI: " & linha
                    LogError ""
                End If
            Else
                houveErro = True
                LogError "[" & FormatDateTime(Now, 3) & "] Linha mal formatada no arquivo INI: " & linha
                LogError ""
            End If
        End If
    End If
Loop

arquivoConfigHandle.Close
Set arquivoConfigHandle = Nothing

Do While jobsAtivos.Count > 0
    VerificarJobsConcluidos
    If jobsAtivos.Count > 0 Then
        WScript.Sleep INTERVALO_POLL_MS
    End If
Loop

SalvarEstado estado, arquivoEstado

If Not executouAlgumBat Then
    Log "[" & FormatDateTime(Now, 3) & "] NENHUM BAT PENDENTE PARA EXECUÇÃO"
    Log ""
End If

Log "================================================================"
Log "FIM DA EXECUÇÃO: " & FormatDateTime(Now, 3)
Log "================================================================"
Log ""

LogError "================================================================"
LogError "FIM DOS REGISTROS DE ERRO: " & FormatDateTime(Now, 0)
LogError "================================================================"
LogError ""

If Not houveErro Then
    GravarNoArquivo caminhoLogDiario, logBufferNormal
Else
    GravarNoArquivo caminhoLogErroDiario, logBufferErro
End If

Set estado = Nothing
Set jobsAtivos = Nothing
Set WshShell = Nothing
Set FSO = Nothing

WScript.Quit 0

Sub DispararJob(nomeBatJob, caminhoBatJob, horarioJob, chaveEstadoJob)
    Dim caminhoLockJob
    Dim sufixo
    Dim caminhoSaidaJob
    Dim caminhoExitJob
    Dim caminhoDoneJob
    Dim caminhoWrapperJob
    Dim arquivoWrapper
    Dim job

    sufixo = Replace(nomeBatJob, ".bat", "") & "_" & Replace(horarioJob, ":", "") & "_" & AgoraUnix()
    caminhoLockJob = pastaLock & "\" & Replace(nomeBatJob, ".bat", "") & "_" & Replace(horarioJob, ":", "") & ".lock"

    If Not CriarLock(caminhoLockJob) Then
        Log "[" & FormatDateTime(Now, 3) & "] " & nomeBatJob & " (agendado " & horarioJob & ") | IGNORADO (Já em execução)"
        Log ""
        Exit Sub
    End If

    caminhoSaidaJob = pastaTmp & "\" & sufixo & "_saida.tmp"
    caminhoExitJob = pastaTmp & "\" & sufixo & "_exit.tmp"
    caminhoDoneJob = pastaTmp & "\" & sufixo & "_done.tmp"
    caminhoWrapperJob = pastaTmp & "\" & sufixo & "_wrapper.cmd"

    Set arquivoWrapper = FSO.CreateTextFile(caminhoWrapperJob, True)
    arquivoWrapper.WriteLine "@echo off"
    arquivoWrapper.WriteLine "cd /d " & aspas & pastaBat & aspas
    arquivoWrapper.WriteLine "call " & aspas & caminhoBatJob & aspas & " > " & aspas & caminhoSaidaJob & aspas & " 2>&1"
    arquivoWrapper.WriteLine "echo %errorlevel% > " & aspas & caminhoExitJob & aspas
    arquivoWrapper.WriteLine "echo done > " & aspas & caminhoDoneJob & aspas
    arquivoWrapper.Close
    Set arquivoWrapper = Nothing

    WshShell.Run "cmd.exe /d /c " & aspas & caminhoWrapperJob & aspas, 0, False

    executouAlgumBat = True

    Set job = New JobExecucao
    job.NomeBat = nomeBatJob
    job.Horario = horarioJob
    job.ChaveEstado = chaveEstadoJob
    job.CaminhoLock = caminhoLockJob
    job.CaminhoSaida = caminhoSaidaJob
    job.CaminhoExit = caminhoExitJob
    job.CaminhoDone = caminhoDoneJob
    job.CaminhoWrapper = caminhoWrapperJob
    job.Inicio = Now

    jobsAtivos.Add chaveEstadoJob & "|" & sufixo, job
End Sub

Sub VerificarJobsConcluidos()
    Dim chaves
    Dim chave
    Dim job
    Dim codigoSaida
    Dim conteudoSaida
    Dim arquivoTxt
    Dim fimJob
    Dim k

    chaves = jobsAtivos.Keys

    For k = 0 To UBound(chaves)
        chave = chaves(k)
        Set job = jobsAtivos(chave)

        If FSO.FileExists(job.CaminhoDone) Then
            fimJob = Now
            codigoSaida = -1
            conteudoSaida = ""

            If FSO.FileExists(job.CaminhoExit) Then
                Set arquivoTxt = FSO.OpenTextFile(job.CaminhoExit, 1, False)
                If Not arquivoTxt.AtEndOfStream Then
                    codigoSaida = Trim(arquivoTxt.ReadLine)
                End If
                arquivoTxt.Close
                Set arquivoTxt = Nothing
            End If

            If FSO.FileExists(job.CaminhoSaida) Then
                Set arquivoTxt = FSO.OpenTextFile(job.CaminhoSaida, 1, False)
                Do Until arquivoTxt.AtEndOfStream
                    linha = arquivoTxt.ReadLine
                    If Trim(linha) <> "" Then
                        conteudoSaida = conteudoSaida & linha & vbCrLf
                    End If
                Loop
                arquivoTxt.Close
                Set arquivoTxt = Nothing
            End If

            If IsNumeric(codigoSaida) And CLng(codigoSaida) = 0 Then
                Log "[" & FormatDateTime(fimJob, 3) & "] " & job.NomeBat & " (agendado " & job.Horario & ")"
                Log "  Status: OK (Saída: 0) | Duração: " & DateDiff("s", job.Inicio, fimJob) & "s"
                Log ""
                estado(job.ChaveEstado) = hojeStr
            Else
                houveErro = True
                LogError "[" & FormatDateTime(fimJob, 3) & "] " & job.NomeBat & " (agendado " & job.Horario & ")"
                LogError "  Status: ERRO (Saída: " & codigoSaida & ") | Duração: " & DateDiff("s", job.Inicio, fimJob) & "s"
                If Trim(conteudoSaida) <> "" Then
                    LogError "  Detalhes: " & Replace(Trim(conteudoSaida), vbCrLf, " ")
                Else
                    LogError "  Detalhes: Nenhuma mensagem retornada pelo BAT."
                End If
                LogError ""
                estado(job.ChaveEstado) = hojeStr
            End If

            LimparArquivoSeExistir job.CaminhoSaida
            LimparArquivoSeExistir job.CaminhoExit
            LimparArquivoSeExistir job.CaminhoDone
            LimparArquivoSeExistir job.CaminhoWrapper
            RemoverLock job.CaminhoLock

            jobsAtivos.Remove chave
        End If
    Next
End Sub

Sub LimparArquivoSeExistir(caminho)
    On Error Resume Next
    If FSO.FileExists(caminho) Then FSO.DeleteFile caminho, True
    On Error GoTo 0
End Sub

Sub Log(texto)
    logBufferNormal = logBufferNormal & texto & vbCrLf
End Sub

Sub LogError(texto)
    logBufferErro = logBufferErro & texto & vbCrLf
End Sub

Sub GravarNoArquivo(caminho, conteudo)
    Dim arquivo
    Set arquivo = FSO.OpenTextFile(caminho, 8, True)
    arquivo.Write conteudo
    arquivo.Close
    Set arquivo = Nothing
End Sub

Sub LimparLogsAntigos()
    Dim arquivo
    Dim limite
    limite = DateAdd("d", - DIAS_RETENCAO_LOGS, Now)

    For Each arquivo In FSO.GetFolder(pastaLog).Files
        If LCase(FSO.GetExtensionName(arquivo.Name)) = "log" Then
            If arquivo.DateLastModified < limite Then
                FSO.DeleteFile arquivo.Path, True
            End If
        End If
    Next
End Sub

Function AgoraUnix()
    AgoraUnix = DateDiff("s", DateSerial(1970, 1, 1), Now)
End Function

Function ObterHojeStr()
    ObterHojeStr = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)
End Function

Function ValidarHorario(h)
    Dim p, hh, mm
    ValidarHorario = False
    If InStr(h, ":") = 0 Then Exit Function
    p = Split(h, ":")
    If UBound(p) <> 1 Then Exit Function
    If Not IsNumeric(p(0)) Or Not IsNumeric(p(1)) Then Exit Function
    hh = CInt(p(0))
    mm = CInt(p(1))
    If hh < 0 Or hh > 23 Then Exit Function
    If mm < 0 Or mm > 59 Then Exit Function
    ValidarHorario = True
End Function

Function MinutosDoHorario(h)
    Dim p
    p = Split(h, ":")
    MinutosDoHorario = CInt(p(0)) * 60 + CInt(p(1))
End Function

Function CarregarEstado()
    Dim dict, arquivo, linhaEst, partesEst
    Set dict = CreateObject("Scripting.Dictionary")
    If Not FSO.FileExists(arquivoEstado) Then
        Set CarregarEstado = dict
        Exit Function
    End If
    Set arquivo = FSO.OpenTextFile(arquivoEstado, 1, False)
    Do Until arquivo.AtEndOfStream
        linhaEst = Trim(arquivo.ReadLine)
        If linhaEst <> "" Then
            partesEst = Split(linhaEst, "=")
            If UBound(partesEst) >= 1 Then
                dict(Trim(partesEst(0))) = Trim(partesEst(1))
            End If
        End If
    Loop
    arquivo.Close
    Set arquivo = Nothing
    Set CarregarEstado = dict
End Function

Sub SalvarEstado(dict, caminho)
    Dim arquivo, chaveEst
    Set arquivo = FSO.CreateTextFile(caminho, True)
    For Each chaveEst In dict.Keys
        arquivo.WriteLine chaveEst & "=" & dict(chaveEst)
    Next
    arquivo.Close
    Set arquivo = Nothing
End Sub

Function CriarLock(caminho)
    Dim arquivo
    CriarLock = False
    On Error Resume Next
    Set arquivo = FSO.CreateTextFile(caminho, False)
    If Err.Number = 0 Then
        arquivo.WriteLine "INICIO=" & FormatDateTime(Now, 0)
        arquivo.Close
        Set arquivo = Nothing
        CriarLock = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Sub RemoverLock(caminho)
    On Error Resume Next
    If FSO.FileExists(caminho) Then
        FSO.DeleteFile caminho, True
    End If
    On Error GoTo 0
End Sub

Class JobExecucao
    Public NomeBat
    Public Horario
    Public ChaveEstado
    Public CaminhoLock
    Public CaminhoSaida
    Public CaminhoExit
    Public CaminhoDone
    Public CaminhoWrapper
    Public Inicio
End Class