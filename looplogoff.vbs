Option Explicit

' Arrays dos feriados nacionais, estaduais e municipais
Dim arrFeriadosNacionais, arrFeriadosEstaduaisBahia, arrFeriadosEstaduaisPernambuco
Dim arrFeriadosMunicipaisPetrolina, arrFeriadosMunicipaisJuazeiro
Dim arrFeriadosMunicipaisBonfim, arrFeriadosMunicipaisJacobina
Dim arrFeriadosVariaveis

arrFeriadosNacionais = Array("07/09", "25/12", "02/11", "12/02", "30/05", "13/02", "15/11", "29/03", "01/05", "01/01", "12/10", "21/04")
arrFeriadosEstaduaisBahia = Array("02/07", "20/11")
arrFeriadosEstaduaisPernambuco = Array("06/03", "20/11")
arrFeriadosMunicipaisPetrolina = Array("21/09", "24/06", "15/08")
arrFeriadosMunicipaisJuazeiro = Array("29/01", "15/08", "08/09")
arrFeriadosMunicipaisBonfim = Array("28/05", "24/06", "17/01")
arrFeriadosMunicipaisJacobina = Array("08/12", "11/08", "13/06", "24/06")
arrFeriadosVariaveis = Array("24/02/2024", "25/02/2024", "26/02/2024", "05/04/2024", _
                        "16/02/2025", "17/02/2025", "18/02/2025", "28/03/2025", _
                        "08/02/2026", "09/02/2026", "10/02/2026", "27/03/2026", _
                        "21/02/2027", "22/02/2027", "23/02/2027", "02/04/2027", _
                        "13/02/2028", "14/02/2028", "15/02/2028", "24/03/2028", _
                        "05/02/2029", "06/02/2029", "07/02/2029", "20/04/2029", _
                        "17/02/2030", "18/02/2030", "19/02/2030", "05/04/2030", _
                        "09/02/2031", "10/02/2031", "11/02/2031", "18/04/2031", _
                        "25/02/2032", "26/02/2032", "27/02/2032", "02/04/2032", _
                        "14/02/2033", "15/02/2033", "16/02/2033", "15/04/2033")

' Array dos números de série dos pendrives permitidos
Dim arrPendrivesPermitidos
arrPendrivesPermitidos = Array("010145e3efc0c86ae539")

' Função para obter o login do usuário ativo
Function GetLoggedUser()
    Dim objShell, objExec, strUserName
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("whoami")
    strUserName = objExec.StdOut.ReadLine()
    If Err.Number <> 0 Then
        strUserName = "Desconhecido"
    End If
    On Error GoTo 0
    Set objExec = Nothing
    Set objShell = Nothing
    GetLoggedUser = strUserName
End Function

' Função para obter o hostname do computador
Function GetHostName()
    Dim objShell, objExec, strHostName
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("hostname")
    strHostName = objExec.StdOut.ReadLine()
    If Err.Number <> 0 Then
        strHostName = "Desconhecido"
    End If
    On Error GoTo 0
    Set objExec = Nothing
    Set objShell = Nothing
    GetHostName = strHostName
End Function

' Função para verificar a hora
Function CheckTime()
    Dim currentHour, currentMinute
    currentHour = Hour(Now)
    currentMinute = Minute(Now)
    If (currentHour >= 8 And currentHour < 18) Or (currentHour = 18 And currentMinute <= 30) Then
        CheckTime = False
    Else
        CheckTime = True
    End If
End Function

' Função para verificar se é fim de semana
Function IsWeekend()
    Dim dayOfWeek
    dayOfWeek = Weekday(Now, vbMonday)
    If dayOfWeek > 5 Then ' 6 = Sábado, 7 = Domingo
        IsWeekend = True
    Else
        IsWeekend = False
    End If
End Function

' Função para verificar se algum pendrive permitido está conectado
Function IsPendriveAllowedConnected(arrPendrivesPermitidos)
    Dim objWMIService, colDisks, objDisk, i, serial
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colDisks = objWMIService.ExecQuery("Select * from Win32_DiskDrive")

    For Each objDisk in colDisks
        serial = objDisk.SerialNumber
        For i = 0 To UBound(arrPendrivesPermitidos)
            If serial = arrPendrivesPermitidos(i) Then
                IsPendriveAllowedConnected = True
                Set objWMIService = Nothing
                Set colDisks = Nothing
                Exit Function
            End If
        Next
    Next

    Set objWMIService = Nothing
    Set colDisks = Nothing
    IsPendriveAllowedConnected = False
End Function

' Função para verificar se uma data é feriado
Function IsHoliday(dateToCheck, arrHolidays)
    Dim i
    For i = 0 To UBound(arrHolidays)
        If dateToCheck = arrHolidays(i) Then
            IsHoliday = True
            Exit Function
        End If
    Next
    IsHoliday = False
End Function

' Função para adicionar o ano atual às datas
Function AddCurrentYear(arrHolidays)
    Dim yearNow, j
    yearNow = Year(Now)
    For j = 0 To UBound(arrHolidays)
        arrHolidays(j) = arrHolidays(j) & "/" & yearNow
    Next
    AddCurrentYear = arrHolidays
End Function

' Obtém o hostname do computador
Dim hostName, cityName, stateName
hostName = GetHostName()

' Determina a cidade e o estado com base no hostname
Select Case True
    Case InStr(hostName, "210101") > 0 Or InStr(hostName, "210102") > 0 Or InStr(hostName, "210104") > 0 Or InStr(hostName, "210106") > 0
        cityName = "Petrolina"
        stateName = "Pernambuco"
    Case InStr(hostName, "210103") > 0
        cityName = "Juazeiro"
        stateName = "Bahia"
    Case InStr(hostName, "210105") > 0
        cityName = "Senhor do Bonfim"
        stateName = "Bahia"
    Case InStr(hostName, "210107") > 0
        cityName = "Jacobina"
        stateName = "Bahia"
    Case Else
        cityName = "Desconhecida"
        stateName = "Desconhecido"
End Select

' Adiciona o ano atual às datas de feriados
arrFeriadosNacionais = AddCurrentYear(arrFeriadosNacionais)
arrFeriadosEstaduaisBahia = AddCurrentYear(arrFeriadosEstaduaisBahia)
arrFeriadosEstaduaisPernambuco = AddCurrentYear(arrFeriadosEstaduaisPernambuco)
arrFeriadosMunicipaisPetrolina = AddCurrentYear(arrFeriadosMunicipaisPetrolina)
arrFeriadosMunicipaisJuazeiro = AddCurrentYear(arrFeriadosMunicipaisJuazeiro)
arrFeriadosMunicipaisBonfim = AddCurrentYear(arrFeriadosMunicipaisBonfim)
arrFeriadosMunicipaisJacobina = AddCurrentYear(arrFeriadosMunicipaisJacobina)

' Lista de usuários permitidos
Dim arrAllowedUsers
arrAllowedUsers = Array(hostName & "\suporte", "sicredi\aquis_souza", "sicredi\charlles_andrade")

' Verifica se o usuário logado está na lista de usuários permitidos
Dim loggedUser, i, isUserAllowed
loggedUser = GetLoggedUser()
isUserAllowed = False

For i = 0 To UBound(arrAllowedUsers)
    If LCase(loggedUser) = LCase(arrAllowedUsers(i)) Then
        isUserAllowed = True
        Exit For
    End If
Next

' Se o usuário for permitido, encerra o script
If isUserAllowed Then
    WScript.Quit
End If

' Verificação de Feriados e Finais de Semana
Dim currentDate, isHolidayToday, isWeekendToday, isPendriveConnected
currentDate = Day(Now) & "/" & Month(Now)
isHolidayToday = IsHoliday(currentDate, arrFeriadosNacionais) Or IsHoliday(currentDate, arrFeriadosVariaveis)
isWeekendToday = IsWeekend()
isPendriveConnected = IsPendriveAllowedConnected(arrPendrivesPermitidos)

If stateName = "Bahia" Then
    isHolidayToday = isHolidayToday Or IsHoliday(currentDate, arrFeriadosEstaduaisBahia)
ElseIf stateName = "Pernambuco" Then
    isHolidayToday = isHolidayToday Or IsHoliday(currentDate, arrFeriadosEstaduaisPernambuco)
End If

Select Case cityName
    Case "Petrolina"
        isHolidayToday = isHolidayToday Or IsHoliday(currentDate, arrFeriadosMunicipaisPetrolina)
    Case "Juazeiro"
        isHolidayToday = isHolidayToday Or IsHoliday(currentDate, arrFeriadosMunicipaisJuazeiro)
    Case "Senhor do Bonfim"
        isHolidayToday = isHolidayToday Or IsHoliday(currentDate, arrFeriadosMunicipaisBonfim)
    Case "Jacobina"
        isHolidayToday = isHolidayToday Or IsHoliday(currentDate, arrFeriadosMunicipaisJacobina)
End Select

' Se for fim de semana ou feriado
If isHolidayToday Or isWeekendToday Then
    If isPendriveConnected Then
        WScript.Quit
    Else
        Do
            WScript.Sleep 1000 ' Espera 1 segundo
            Dim WshShell
            Set WshShell = WScript.CreateObject("WScript.Shell")
            WshShell.Popup "Pendrive não conectado. O sistema irá fazer logoff agora.", 5, "Atenção", 48
            WshShell.Run "shutdown.exe -l", 0, False
            Set WshShell = Nothing
        Loop
    End If
End If

' Loop de Verificação de Hora

Do
    If CheckTime() Then
        isPendriveConnected = IsPendriveAllowedConnected(arrPendrivesPermitidos)
        If isPendriveConnected Then
            WScript.Quit
        Else
            Do
                WScript.Sleep 1000 ' Espera 1 segundo
                Set WshShell = WScript.CreateObject("WScript.Shell")
                WshShell.Run "shutdown.exe -l", 0, False
                Set WshShell = Nothing
            Loop
        End If
    End If
    
    WScript.Sleep 60000 ' Espera 1 minuto
Loop
