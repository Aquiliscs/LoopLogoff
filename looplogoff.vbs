Option Explicit

' Arrays dos feriados nacionais, estaduais e municipais
Dim feriadosNacionais, feriadosEstaduaisBahia, feriadosEstaduaisPernambuco
Dim feriadosMunicipaisPetrolina, feriadosMunicipaisJuazeiro
Dim feriadosMunicipaisBonfim, feriadosMunicipaisJacobina
Dim feriadosVariaveis

feriadosNacionais = Array("07/09", "25/12", "02/11", "12/02", "30/05", "13/02", "15/11", "29/03", "01/05", "01/01", "12/10", "21/04")
feriadosEstaduaisBahia = Array("02/07", "20/11")
feriadosEstaduaisPernambuco = Array("06/03", "20/11")
feriadosMunicipaisPetrolina = Array("21/09", "24/06", "15/08")
feriadosMunicipaisJuazeiro = Array("29/01", "15/08", "08/09")
feriadosMunicipaisBonfim = Array("28/05", "24/06", "17/01")
feriadosMunicipaisJacobina = Array("08/12", "11/08", "13/06", "24/06")
feriadosVariaveis = Array("24/02/2024", "25/02/2024", "26/02/2024", "05/04/2024", _
                        "16/02/2025", "17/02/2025", "18/02/2025", "28/03/2025", _
                        "08/02/2026", "09/02/2026", "10/02/2026", "27/03/2026", _
                        "21/02/2027", "22/02/2027", "23/02/2027", "02/04/2027", _
                        "13/02/2028", "14/02/2028", "15/02/2028", "24/03/2028", _
                        "05/02/2029", "06/02/2029", "07/02/2029", "20/04/2029", _
                        "17/02/2030", "18/02/2030", "19/02/2030", "05/04/2030", _
                        "09/02/2031", "10/02/2031", "11/02/2031", "18/04/2031", _
                        "25/02/2032", "26/02/2032", "27/02/2032", "02/04/2032", _
                        "14/02/2033", "15/02/2033", "16/02/2033", "15/04/2033")

' Função para obter o login do usuário ativo
Function ObterUsuarioLogado()
    Dim objShell, objExec, strNomeUsuario
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("whoami")
    strNomeUsuario = objExec.StdOut.ReadLine()
    If Err.Number <> 0 Then
        strNomeUsuario = "Desconhecido"
    End If
    On Error GoTo 0
    Set objExec = Nothing
    Set objShell = Nothing
    ObterUsuarioLogado = strNomeUsuario
End Function

' Função para obter o hostname do computador
Function ObterHostname()
    Dim objShell, objExec, strHostname
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("hostname")
    strHostname = objExec.StdOut.ReadLine()
    If Err.Number <> 0 Then
        strHostname = "Desconhecido"
    End If
    On Error GoTo 0
    Set objExec = Nothing
    Set objShell = Nothing
    ObterHostname = strHostname
End Function

' Função para verificar a hora
Function VerificarHora()
    Dim horaAtual
    horaAtual = Hour(Now)
    If horaAtual >= 18 Or horaAtual < 8 Then
        VerificarHora = True
    Else
        VerificarHora = False
    End If
End Function

' Função para registrar mensagens
Sub RegistrarMensagem(mensagem)
    Dim fso, arquivoLog
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arquivoLog = fso.OpenTextFile("C:\Temp\loglooplogoff.txt", 8, True)
    arquivoLog.WriteLine(Now & " - " & mensagem)
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
End Sub

' Função para verificar se uma data é feriado
Function Eferiado(data, feriados)
    Dim i
    For i = 0 To UBound(feriados)
        If data = feriados(i) Then
            Eferiado = True
            Exit Function
        End If
    Next
    Eferiado = False
End Function

' Função para adicionar o ano atual às datas
Function AdicionarAnoAtual(feriados)
    Dim ano, i
    ano = Year(Now)
    For i = 0 To UBound(feriados)
        feriados(i) = feriados(i) & "/" & ano
    Next
    AdicionarAnoAtual = feriados
End Function

' Obtém o hostname do computador
Dim hostname, cidade, estado
hostname = ObterHostname()
RegistrarMensagem "Hostname: " & hostname

' Determina a cidade e o estado com base no hostname
Select Case True
    Case InStr(hostname, "210101") > 0 Or InStr(hostname, "210102") > 0 Or InStr(hostname, "210104") > 0 Or InStr(hostname, "210106") > 0
        cidade = "Petrolina"
        estado = "Pernambuco"
    Case InStr(hostname, "210103") > 0
        cidade = "Juazeiro"
        estado = "Bahia"
    Case InStr(hostname, "210105") > 0
        cidade = "Senhor do Bonfim"
        estado = "Bahia"
    Case InStr(hostname, "210107") > 0
        cidade = "Jacobina"
        estado = "Bahia"
    Case Else
        cidade = "Desconhecida"
        estado = "Desconhecido"
End Select
RegistrarMensagem "Cidade: " & cidade & ", Estado: " & estado

' Adiciona o ano atual às datas de feriados
feriadosNacionais = AdicionarAnoAtual(feriadosNacionais)
feriadosEstaduaisBahia = AdicionarAnoAtual(feriadosEstaduaisBahia)
feriadosEstaduaisPernambuco = AdicionarAnoAtual(feriadosEstaduaisPernambuco)
feriadosMunicipaisPetrolina = AdicionarAnoAtual(feriadosMunicipaisPetrolina)
feriadosMunicipaisJuazeiro = AdicionarAnoAtual(feriadosMunicipaisJuazeiro)
feriadosMunicipaisBonfim = AdicionarAnoAtual(feriadosMunicipaisBonfim)
feriadosMunicipaisJacobina = AdicionarAnoAtual(feriadosMunicipaisJacobina)

' Lista de usuários permitidos
Dim usuariosPermitidos
usuariosPermitidos = Array(hostname & "\suporte", "sicredi\aquilis_souza", "sicredi\charlles_andrade")

' Verifica se o usuário logado está na lista de usuários permitidos
Dim usuarioLogado, i, usuarioPermitido
usuarioLogado = ObterUsuarioLogado()
RegistrarMensagem "Usuário Logado: " & usuarioLogado
usuarioPermitido = False

For i = 0 To UBound(usuariosPermitidos)
    If LCase(usuarioLogado) = LCase(usuariosPermitidos(i)) Then
        usuarioPermitido = True
        Exit For
    End If
Next

' Se o usuário for permitido, encerra o script
If usuarioPermitido Then
    RegistrarMensagem "Usuário permitido, encerrando script."
    WScript.Quit
End If

RegistrarMensagem "Usuário não permitido, entrando em loop."

' Se o usuário não for permitido, entra em loop de verificação de hora
Do
    Dim dataAtual
    dataAtual = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    Dim eFeriado
    eFeriado = Eferiado(dataAtual, feriadosNacionais) Or Eferiado(dataAtual, feriadosVariaveis)
    
    If estado = "Bahia" Then
        eFeriado = eFeriado Or Eferiado(dataAtual, feriadosEstaduaisBahia)
    ElseIf estado = "Pernambuco" Then
        eFeriado = eFeriado Or Eferiado(dataAtual, feriadosEstaduaisPernambuco)
    End If
    
    Select Case cidade
        Case "Petrolina"
            eFeriado = eFeriado Or Eferiado(dataAtual, feriadosMunicipaisPetrolina)
        Case "Juazeiro"
            eFeriado = eFeriado Or Eferiado(dataAtual, feriadosMunicipaisJuazeiro)
        Case "Senhor do Bonfim"
            eFeriado = eFeriado Or Eferiado(dataAtual, feriadosMunicipaisBonfim)
        Case "Jacobina"
            eFeriado = eFeriado Or Eferiado(dataAtual, feriadosMunicipaisJacobina)
    End Select
    
    ' Se estiver fora do horário permitido ou for feriado, faz logoff do usuário
    If VerificarHora() Or eFeriado Then
        RegistrarMensagem "Fora do horário permitido ou feriado, fazendo logoff do usuário."
        Do
            WScript.Sleep 1000 ' Espera 1 segundo
            Dim WshShell
            Set WshShell = WScript.CreateObject("WScript.Shell")
            WshShell.Run "shutdown.exe -l", 0, False
        Loop
    End If
    
    WScript.Sleep 60000 ' Espera 1 minuto
Loop
