Imports System.Data.OleDb
Imports System.Net.Mail
Imports Microsoft.VisualBasic.Devices
Imports Microsoft.Win32
Imports System.Xml
Imports System.Net.NetworkInformation
Imports System.Drawing.Printing
Imports System.Threading
Imports System.Management


Public Class SistemaOperacional

    Public Shared Function Linguagem()
        Dim ComputerInfo As ComputerInfo = New ComputerInfo
        Dim Cult_Instalada As String = ComputerInfo.InstalledUICulture.ToString
        Return Cult_Instalada
    End Function

    Public Shared Function Nome()
        Dim ComputerInfo As ComputerInfo = New ComputerInfo
        Dim Nom_SO As String = ComputerInfo.OSFullName.ToString
        Return Nom_SO
    End Function

    Public Shared Function Plataforma()
        Dim PlataformaDetectada As String
        Dim pa As String = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")
        PlataformaDetectada = IIf((String.IsNullOrEmpty(pa) Or String.Compare(pa, 0, "x86", 0, 3, True) = 0), 32, 64).ToString
        Return PlataformaDetectada & " Bits"
    End Function

    Public Shared Function ServicePack()
        Dim Versao_SO As String = Environment.OSVersion.ServicePack.ToString
        Return Versao_SO
    End Function

End Class

Public Class Memoria
    Public Shared Function MemoriaFisicaTotal()
        Dim ComputerInfo As ComputerInfo = New ComputerInfo
        Dim MemFisTot As String = FormataNumero(ComputerInfo.TotalPhysicalMemory.ToString)
        Return MemFisTot
    End Function

    Public Shared Function MemoriaFisicaDisponivel()
        Dim ComputerInfo As ComputerInfo = New ComputerInfo
        Dim MemFisTot As String = FormataNumero(ComputerInfo.AvailablePhysicalMemory.ToString)
        Return MemFisTot
    End Function

End Class

Public Class DiscosLocais
    Public Shared Function Nome(ByVal Drive As String)
        Dim Infodrive As IO.DriveInfo = New IO.DriveInfo(Drive)
        Return Infodrive.Name
    End Function

    Public Shared Function Rotulo(ByVal Drive As String)
        Dim Infodrive As IO.DriveInfo = New IO.DriveInfo(Drive)
        Return Infodrive.VolumeLabel
    End Function

    Public Shared Function Tipo(ByVal Drive As String)
        Dim Infodrive As IO.DriveInfo = New IO.DriveInfo(Drive)
        Return Infodrive.DriveType.ToString
    End Function

    Public Shared Function Formato(ByVal Drive As String)
        Dim Infodrive As IO.DriveInfo = New IO.DriveInfo(Drive)
        Return Infodrive.DriveFormat
    End Function

    Public Shared Function TamanhoTotal(ByVal Drive As String)
        Dim Infodrive As IO.DriveInfo = New IO.DriveInfo(Drive)
        Return FormataNumero(Infodrive.TotalSize)
    End Function

    Public Shared Function TamanhoDisponivel(ByVal Drive As String)
        Dim Infodrive As IO.DriveInfo = New IO.DriveInfo(Drive)
        Return FormataNumero(Infodrive.TotalFreeSpace)
    End Function

End Class

Public Class ProdutosMicrosoft
    Public Shared Function ChaveWindows()
        Dim ChaveAtualWindows As String = GetProductKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\", "DigitalProductId")
        Return ChaveAtualWindows
    End Function

    Public Shared Function ChaveOffice()
        Dim ChaveAtualOffice As String = GetProductKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\14.0\Registration\{90140000-0011-0000-1000-0000000FF1CE}", "DigitalProductId")
        Return ChaveAtualOffice
    End Function
End Class

Public Class Monitor
    Public Shared Function ResolucaoTelaAtual() As String
        Dim intX As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim intY As Integer = Screen.PrimaryScreen.Bounds.Height
        Return intX & " X " & intY
    End Function
    Public Shared Function PlacaVideo()
        Dim NomePlacadeVideo = String.Empty

        '   Dim WmiSelect As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_VideoController")
        '    For Each WmiResults As ManagementObject In WmiSelect.Get()
        ' NomePlacadeVideo = WmiResults.GetPropertyValue("Name").ToString & " com " & FormataNumero(WmiResults.GetPropertyValue("AdapterRam").ToString) & " Dedicado de RAM"
        ' If (Not String.IsNullOrEmpty(NomePlacadeVideo)) Then
        ' Exit For
        '  End If
        ' Next
        Return NomePlacadeVideo
    End Function
End Class

Public Class GravalogNovo
    Public Shared Function GravaLogNovo(ByVal Mensagem As String)
        Dim NMARQLOG As String = My.Application.Info.DirectoryPath & Date.Now.Year & Date.Now.Month & Date.Now.Day & Date.Now.Hour & Date.Now.Minute & Date.Now.Second & ".log"
        Dim ESCRITOR As IO.FileStream
        Dim ESCRITORTEXTO As IO.StreamWriter
        If System.IO.File.Exists(NMARQLOG) Then
            ESCRITOR = New System.IO.FileStream(NMARQLOG, IO.FileMode.Append, IO.FileAccess.Write)
        Else
            ESCRITOR = New System.IO.FileStream(NMARQLOG, IO.FileMode.CreateNew, IO.FileAccess.Write)
        End If
        ESCRITORTEXTO = New System.IO.StreamWriter(ESCRITOR)
        ESCRITORTEXTO.WriteLine(DateTime.Now & " - " & Mensagem)
        ESCRITORTEXTO.Close()
        Return True
    End Function
End Class

Public Class Computador
    Public Shared Function Nome()
        Dim NomeComputador As String = My.Computer.Name()
        Return NomeComputador
    End Function
    Public Shared Function UsuarioAtual()
        Dim Usuario As String = Environment.UserDomainName & "\" & Environment.UserName
        Return Usuario
    End Function
    Public Shared Function ControladorDominio()
        Dim logonserver As String = Environment.GetEnvironmentVariable("LOGONSERVER")
        Return logonserver
    End Function
    Public Shared Function Uptime()
        Dim time As String = String.Empty
        time += Math.Round(Environment.TickCount / 86400000) & " days, "
        time += Math.Round(Environment.TickCount / 3600000 Mod 24) & " hours, "
        time += Math.Round(Environment.TickCount / 120000 Mod 60) & " minutes, "
        time += Math.Round(Environment.TickCount / 1000 Mod 60) & " seconds."
        Return time
    End Function

    Public Shared Function TimeZoneAtual()
        Dim curTimeZone As TimeZone = TimeZone.CurrentTimeZone
        Dim MTimezoneAtual = curTimeZone.StandardName
        Return "TimeZone: " & MTimezoneAtual
    End Function


End Class

Public Class Processador
    Public Shared Function Nome()
        Dim m_LM As RegistryKey
        Dim m_HW As RegistryKey
        Dim m_Des As RegistryKey
        Dim m_System As RegistryKey
        Dim m_CPU As RegistryKey
        Dim m_Info As RegistryKey
        Dim Core As String = System.Environment.ProcessorCount()
        Dim NomeProcessador As String
        m_LM = Registry.LocalMachine
        m_HW = m_LM.OpenSubKey("HARDWARE")
        m_Des = m_HW.OpenSubKey("DESCRIPTION")
        m_System = m_Des.OpenSubKey("SYSTEM")
        m_CPU = m_System.OpenSubKey("CentralProcessor")
        m_Info = m_CPU.OpenSubKey("0")

        NomeProcessador = Core & " x " & m_Info.GetValue("ProcessorNameString") '& " " & m_Info.GetValue("Identifier") & " " & m_Info.GetValue("~Mhz") & "MHz"
        Return NomeProcessador
    End Function
End Class


Module Module1

    'Variaveis Gerais
    Dim NMARQLOG As String = My.Application.Info.DirectoryPath & "\Logs\" & "AuditCO_Cliente" & Date.Now.Year & Date.Now.Month & Date.Now.Day & Date.Now.Hour & Date.Now.Minute & Date.Now.Second & ".log"
    Dim ESCRITOR As IO.FileStream
    Dim ESCRITORTEXTO As IO.StreamWriter
    Dim NRPAREME As Integer = 1
    Dim NMDIRBKP As String
    Dim INDELARQ As Integer
    Dim INMAPRED As Integer
    Dim LETMAPRE As String
    Dim DSCAMMAP As String
    Dim NMDESTIN As String
    Dim ENDREMEX As String
    Dim ENDDESTX As String
    Dim ENDSMTPX As String
    Dim DSPORTAX As Integer = 0
    Dim DSUSUARI As String
    Dim DSSENHAX As String
    Dim AUXORIGX As String
    Dim AUXDESTX As String
    Dim EXTNMPES As String
    Dim INLIMDES As Integer
    Dim ComputerInfo As ComputerInfo = New ComputerInfo

    'Declaracao das Funcoes GetFileAtributes e GetPrivaProfileString
    Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Integer

    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
      ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, _
      ByVal lpDefault As String, _
      ByVal lpReturnedString As String, _
      ByVal nSize As Integer, _
      ByVal lpFileName As String) As Integer

    'Declaracao das Funcoes de Mapeamento de Recurso de Rede
    Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" _
                           (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, _
                           ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer

    Public Declare Function WNetCancelConnection2 Lib "mpr" Alias "WNetCancelConnection2A" _
                            (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer

    Public Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure

    Public Const ForceDisconnect As Integer = 1
    Public Const RESOURCETYPE_DISK As Long = &H1

    'Declaracao das Funcoes de Busca de Informacoes de Disco Locais
    Partial Public Class Win32API

        Public Declare Function GetDiskFreeSpace Lib "kernel32" _
            Alias "GetDiskFreeSpaceA" (ByVal RootPathName As String,
                                                     ByRef SectorsPerCluster As Integer,
                                                     ByRef BytesPerSector As Integer,
                                                     ByRef NumberOfFreeClusters As Integer,
                                                     ByRef TotalNumberOfClusters As Integer) As Integer

        Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
            Alias "GetDiskFreeSpaceExA" (ByVal RootPathName As String,
                                                         ByRef FreeBytesAvailableToCaller As Integer,
                                                         ByRef TotalNumberOfBytes As Integer,
                                                         ByRef TotalNumberOfFreeBytes As UInt32) As Integer

        Public Declare Function GetDriveType Lib "kernel32" _
             Alias "GetDriveTypeA" (ByVal nDrive As String) As Integer

    End Class

    Public Function MapeiaDrive(ByVal LetraDrive As String, ByVal Caminho As String) As Boolean

        Dim recursorede As NETRESOURCE
        Dim strUsername As String
        Dim strPassword As String

        recursorede = New NETRESOURCE
        recursorede.lpRemoteName = Caminho
        recursorede.lpLocalName = LetraDrive & ":"
        strUsername = PegaParametro("GERAL", "DSUSUARI")
        strPassword = PegaParametro("GERAL", "DSSENHAX")
        recursorede.dwType = RESOURCETYPE_DISK

        Dim resultado As Integer
        resultado = WNetAddConnection2(recursorede, strPassword, strUsername, 0)

        If resultado = 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function RemoveMapeamentoDrive(ByVal LetraDrive As String) As Boolean
        Dim rc As Integer
        rc = WNetCancelConnection2(LetraDrive & ":", 0, ForceDisconnect)

        If rc = 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    'Sub GravaLog
    Public Sub GravaLog(ByVal Mensagem As String)
        If System.IO.File.Exists(NMARQLOG) Then
            ESCRITOR = New System.IO.FileStream(NMARQLOG, IO.FileMode.Append, IO.FileAccess.Write)
        Else
            ESCRITOR = New System.IO.FileStream(NMARQLOG, IO.FileMode.CreateNew, IO.FileAccess.Write)
        End If
        ESCRITORTEXTO = New System.IO.StreamWriter(ESCRITOR)
        ESCRITORTEXTO.WriteLine(DateTime.Now & " - " & Mensagem)
        ESCRITORTEXTO.Close()
    End Sub

    'Funcao Pega Parametro do Arquivo .ini, .txt
    Public Function PegaParametro(ByRef CDSESSAO As String, ByRef CDPARAM As String) As String
        Dim NRTAMPAR As Integer = 0
        Dim DSPARAME As String = New String(Chr(0), 255)
        Dim NMARQINI As String = My.Application.Info.DirectoryPath & "\LimpaCache.ini"

        NRTAMPAR = GetPrivateProfileString(CDSESSAO, CDPARAM, "", DSPARAME, Len(DSPARAME), NMARQINI)
        PegaParametro = Microsoft.VisualBasic.Left(DSPARAME, NRTAMPAR)
        'If NRTAMPAR = Nothing Then
        'End
        'End If
    End Function

    Public Sub VerificaExecucaoPrograma(ByVal NomeAplicativo)
        Try
            Dim processo() As Process = Process.GetProcessesByName(NomeAplicativo)
            If processo.Length > 1 Then
                GravaLog("Já Existe Uma Instância do Programa Aberto. Finalizando Execução")
                Application.Exit()
            End If
        Catch
            GravaLog("Erro VerificaExecucaoPrograma : " & Err.Number & Err.Description)
        End Try
    End Sub
    Public Sub ConectaDB()
        Dim NMBANCOX As String
        Dim DSUSUBDX As String
        Dim DSSENBDX As String

        NMBANCOX = PegaParametro("BD", "NMBANCOX")
        DSUSUBDX = PegaParametro("BD", "DSUSUBDX")
        DSSENBDX = PegaParametro("BD", "DSSENBDX")

        Dim CONEXBDX As OleDbConnection = New OleDbConnection()
        CONEXBDX.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Application.Info.DirectoryPath & "\" & NMBANCOX
        CONEXBDX.Open()
    End Sub

    Public Sub ExecutaLimpezaDiretorio(ByVal NMDIRETO)
        Dim TMARQPES As Integer
        INLIMDES = PegaParametro("GERAL", "INLIMDES")
        Try
            If INLIMDES = 1 Then
                GravaLog("...Iniciando Limpeza do Diretório de Destino...")
                For Each TBARQLIS As String In IO.Directory.GetFiles(NMDIRETO, "*.*", IO.SearchOption.AllDirectories)
                    TMARQPES = TBARQLIS.Length
                    IO.File.Delete(TBARQLIS)
                    GravaLog("Arquivo Deletado : " & TBARQLIS)
                Next
            End If
        Catch
            GravaLog("Erro ExecutaLimpezaDiretorio : " & Err.Number & Err.Description)
        End Try
    End Sub

    Public Sub LimpaDiretorio(ByVal NMDIRETO)
        GravaLog("...Limpeza do Diretório: " & NMDIRETO)
        Try
            For Each TBDIRPES As String In IO.Directory.GetDirectories(NMDIRETO, "*.*", IO.SearchOption.AllDirectories)
                For Each TBARQLIS As String In IO.Directory.GetFiles(TBDIRPES, "*.*", IO.SearchOption.AllDirectories)
                    ExcluiArquivo(TBARQLIS)
                Next
                ExcluiDiretorio(TBDIRPES)
            Next
        Catch
            GravaLog("Erro LimpaDiretorio : " & Err.Number & Err.Description)
        End Try
    End Sub


    Public Function CopiaArquivo(ByVal AUXORIGX, ByVal AUXDESTX)
        Dim Sucess As Boolean = False
        Try
            IO.File.Copy(AUXORIGX, AUXDESTX, True)
            GravaLog("...Cópia de Arquivo Após Pesquisa...")
            GravaLog("Arquivo Copiado de : " & AUXORIGX & " para: " & AUXDESTX)
            Return True
        Catch
            GravaLog("Erro CopiaArquivo : " & Err.Number & Err.Description)
            Return False
        End Try
    End Function

    Public Function ExcluiArquivo(ByVal NMARQUIVO)
        Dim Sucess As Boolean = False
        Dim ArquivoFormatado As String = FormataTamanho(NMARQUIVO)
        Try
            IO.File.Delete(NMARQUIVO)
            GravaLog("Arquivo Deletado : " & ArquivoFormatado)
            Return True
        Catch
            GravaLog("Erro ExcluiArquivo : " & Err.Number & Err.Description)
            Return False
        End Try
    End Function

    Public Sub ExcluiDiretorio(TBDIRPES As String)
        Throw New NotImplementedException
    End Sub

    Public Sub PesquisaPasta(ByVal NMDIRETO, ByVal EXTNMPES)
        For Each TBARQPES As String In IO.Directory.GetFiles(NMDIRETO, EXTNMPES, IO.SearchOption.AllDirectories)
            AUXORIGX = TBARQPES
            AUXDESTX = NMDESTIN & "\" & IO.Path.GetFileName(TBARQPES)
            If CopiaArquivo(AUXORIGX, AUXDESTX) = True Then
                ExcluiArquivo(TBARQPES)
            Else
                Exit Sub
            End If
        Next
    End Sub

    Public Sub EnviaEmail(ByVal ENDREMEX As String, _
                          ByVal ENDDESTX As String, _
                          ByVal DSUSUARI As String, _
                          ByVal DSSENHAX As String, _
                          ByVal ENDSMTPX As String, _
                          ByVal DSPORTAX As Integer)
        Dim Mensagem As Net.Mail.MailMessage = New MailMessage(ENDREMEX, ENDDESTX)
        Dim AutenticationCredential As Net.NetworkCredential = New Net.NetworkCredential(DSUSUARI, DSSENHAX)
        Dim ClienteSMTP As Net.Mail.SmtpClient = New Net.Mail.SmtpClient(ENDSMTPX, DSPORTAX)
        Mensagem.IsBodyHtml = True
        'Mensagem.Attachments.Add(New Attachment(My.Application.Info.DirectoryPath & "\LimpaCache" & Date.Now.Year & Date.Now.Month & Date.Now.Day & ".log"))
        Mensagem.Subject = "Programa LimpaCache - Status Report"
        Mensagem.Body = "<P>Bom dia</P>" & _
                        "<P><B><FONT COLOR=""#FF0000"">Esta é uma Mensagem Automática. Favor Não Responder.</FONT></B></P>" & _
                        "<P>Segue Anexo o Log de Execução do Programa LimpaCache, no Servidor: <B>" & My.Computer.Name & "</B> .Favor analisar o conteúdo do Arquivo" & _
                        " e, na constatação de qualquer anomalia comunicar o Suporte - Lince.</P>" & _
                        "<BR>" & _
                        "<P>Obrigado.</P>"
        ClienteSMTP.EnableSsl = True
        ClienteSMTP.UseDefaultCredentials = False
        ClienteSMTP.Credentials = AutenticationCredential
        Try
            ClienteSMTP.Send(Mensagem)
        Catch ex As Exception
            GravaLog("Erro EnviaEmail : " & Err.Number & " - " & Err.Description)
        End Try
    End Sub

    Public Function FormataTamanho(ByVal NMARQUIVO)
        Dim DSARQPES As IO.FileInfo
        Dim TamanhoFormatadoArquivo As String
        Dim NumeroCerto As Integer
        Try
            DSARQPES = New IO.FileInfo(NMARQUIVO)
            If DSARQPES.Length <= 1 Then
                TamanhoFormatadoArquivo = NMARQUIVO & " - " & DSARQPES.Length & " byte."
                Return TamanhoFormatadoArquivo
            End If
            If DSARQPES.Length > 1 And DSARQPES.Length < 1000 Then
                TamanhoFormatadoArquivo = NMARQUIVO & " - " & DSARQPES.Length & " bytes."
                Return TamanhoFormatadoArquivo
            End If
            If DSARQPES.Length >= 1000 Then
                NumeroCerto = DSARQPES.Length / 1000
                TamanhoFormatadoArquivo = NMARQUIVO & " - " & NumeroCerto & " Megabytes."
                Return TamanhoFormatadoArquivo
            End If
            Return True
        Catch
            GravaLog("Erro CopiaArquivo : " & Err.Number & Err.Description)
            Return False
        End Try
    End Function

    Public Function FormataNumero(ByVal NUMERO)
        Dim TamanhoFormatadoNumero As String
        Dim NumeroCerto As Integer
        Try

            If NUMERO <= 1 Then
                TamanhoFormatadoNumero = NUMERO & "byte"
                Return TamanhoFormatadoNumero
            End If
            If NUMERO > 1 And NUMERO < 1024 Then
                TamanhoFormatadoNumero = NUMERO & "bytes"
                Return TamanhoFormatadoNumero
            End If
            If NUMERO >= 1000 And NUMERO < 1048576 Then
                TamanhoFormatadoNumero = NUMERO & "MB"
                Return TamanhoFormatadoNumero
            End If
            If NUMERO >= 1073741824 Then
                NumeroCerto = NUMERO / 1024
                NumeroCerto = NumeroCerto / 1024
                NumeroCerto = NumeroCerto / 1024
                TamanhoFormatadoNumero = NumeroCerto & "GB"
                Return TamanhoFormatadoNumero
            End If
            Return True
        Catch
            GravaLog("Erro FormataNumero : " & Err.Number & Err.Description)
            Return False
        End Try
    End Function

    Public Function GetProductKey(ByVal KeyPath As String, ByVal ValueName As String) As String
        Try
            Dim HexBuf As Object = My.Computer.Registry.GetValue(KeyPath, ValueName, 0)

            If HexBuf Is Nothing Then Return "N/A"

            Dim tmp As String = ""

            For l As Integer = LBound(HexBuf) To UBound(HexBuf)
                tmp = tmp & " " & Hex(HexBuf(l))
            Next

            Dim StartOffset As Integer = 52
            Dim EndOffset As Integer = 67
            Dim Digits(24) As String

            Digits(0) = "B" : Digits(1) = "C" : Digits(2) = "D" : Digits(3) = "F"
            Digits(4) = "G" : Digits(5) = "H" : Digits(6) = "J" : Digits(7) = "K"
            Digits(8) = "M" : Digits(9) = "P" : Digits(10) = "Q" : Digits(11) = "R"
            Digits(12) = "T" : Digits(13) = "V" : Digits(14) = "W" : Digits(15) = "X"
            Digits(16) = "Y" : Digits(17) = "2" : Digits(18) = "3" : Digits(19) = "4"
            Digits(20) = "6" : Digits(21) = "7" : Digits(22) = "8" : Digits(23) = "9"

            Dim dLen As Integer = 29
            Dim sLen As Integer = 15
            Dim HexDigitalPID(15) As String
            Dim Des(30) As String

            Dim tmp2 As String = ""

            For i = StartOffset To EndOffset
                HexDigitalPID(i - StartOffset) = HexBuf(i)
                tmp2 = tmp2 & " " & Hex(HexDigitalPID(i - StartOffset))
            Next

            Dim KEYSTRING As String = ""

            For i As Integer = dLen - 1 To 0 Step -1
                If ((i + 1) Mod 6) = 0 Then
                    Des(i) = "-"
                    KEYSTRING = KEYSTRING & "-"
                Else
                    Dim HN As Integer = 0
                    For N As Integer = (sLen - 1) To 0 Step -1
                        Dim Value As Integer = ((HN * 2 ^ 8) Or HexDigitalPID(N))
                        HexDigitalPID(N) = Value \ 24
                        HN = (Value Mod 24)

                    Next

                    Des(i) = Digits(HN)
                    KEYSTRING = KEYSTRING & Digits(HN)
                End If
            Next

            Return StrReverse(KEYSTRING)
        Catch

        End Try

    End Function

    Public Function DiscosLocaisLIB(ByVal LetraDisco)


        Dim RootPath As String = LetraDisco
        Dim SectorsInCluster As Integer = 0
        Dim BytesInSector As Integer = 0
        Dim NumberFreeClusters = 0
        Dim TotalNumberClusters = 0
        Call Win32API.GetDiskFreeSpace(RootPath, SectorsInCluster, BytesInSector, NumberFreeClusters, TotalNumberClusters)

        GravaLog("GetDiskSpace: Cluster livres em" & LetraDisco & NumberFreeClusters)

        Dim FreeBytes As Integer = 0
        Dim TotalBytes As Integer = 0
        Dim TotalFreeBytes As UInt32 = 0
        Call Win32API.GetDiskFreeSpaceEx(RootPath, FreeBytes, TotalBytes, TotalFreeBytes)

        GravaLog("GetDiskSpaceEx: Total de bytes livres em " & LetraDisco & TotalFreeBytes)

        Dim DriveType As Integer = Win32API.GetDriveType(RootPath)
        Dim DriveTypeName As String = String.Empty
        Select Case DriveType
            Case 2 : DriveTypeName = "Removível"
            Case 3 : DriveTypeName = "Fixo"
            Case 4 : DriveTypeName = "Remoto"
            Case 5 : DriveTypeName = "CD-Rom"
            Case 6 : DriveTypeName = "RAM Disk"
            Case Else : DriveTypeName = "Desconhecido"
        End Select

        GravaLog("GetDriveType: Tipo de Drive : " & DriveTypeName)
        Return True
    End Function

    Public Function ProgramasInstalados()
        Dim SubKey As RegistryKey
        'Abre a chave que consta os programas que possuem desinstalador
        Dim Key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", False)
        'Levanta Todos os Programas
        Dim SubKeyNames() As String = Key.GetSubKeyNames()
        'Lista todos os Programas
        For Index = 0 To Key.SubKeyCount - 1
            SubKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" & "\" & SubKeyNames(Index), False)
            'Valida se o Programa tem um Nome Valido para Exibiçao
            If Not SubKey.GetValue("DisplayName", "") Is "" Then
                GravaLog(CType(SubKey.GetValue("DisplayName", ""), String))
            End If
        Next
        Return True
    End Function

    Public Function CriaXML()
        Dim xmlPath As String = "c:\temp\" & Date.Now.Year & Date.Now.Month & Date.Now.Day & Date.Now.Hour & Date.Now.Minute & Date.Now.Second & ".xml"
        ' Cria um novo ficheiro XML com a codificação UTF8
        Dim xmlw As New XmlTextWriter(xmlPath, System.Text.Encoding.UTF8)
        xmlw.Formatting = Formatting.Indented

        xmlw.WriteStartDocument()

        ' Adiciona um comentário geral
        xmlw.WriteComment("Declaracao do Node Pai")

        ' Criar um elemento geral
        xmlw.WriteStartElement("AuditCo_Cliente")

        ' Criar o Node Filho "Computador" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho Computador")
            .WriteStartElement("Computador")
            .WriteElementString("Nome", Computador.Nome)
            .WriteElementString("Usuario", Computador.UsuarioAtual)
            .WriteElementString("ControladorDominio", Computador.ControladorDominio)
            .WriteElementString("Uptime", Computador.Uptime)
            .WriteElementString("TimeZone", Computador.TimeZoneAtual)
            .WriteEndElement()
        End With

        ' Criar o Node Filho "Processador" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho Processador")
            .WriteStartElement("Processador")
            .WriteElementString("Nome", Processador.Nome)
            .WriteEndElement()
        End With

        ' Criar o Node Filho "Informacoes SO" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho Informações do Sistema Operacional")
            .WriteStartElement("Informacoes_SO")
            .WriteElementString("NomeCompleto", SistemaOperacional.Nome())
            .WriteElementString("Versao", SistemaOperacional.ServicePack())
            .WriteElementString("Plataforma", SistemaOperacional.Plataforma())
            .WriteElementString("Linguagem", SistemaOperacional.Linguagem())
            .WriteEndElement()
        End With

        ' Criar o Node Filho "Memoria" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho Memoria")
            .WriteStartElement("Memoria")
            .WriteElementString("MemoriaTotal", Memoria.MemoriaFisicaTotal)
            .WriteElementString("MemoriaDisponivel", Memoria.MemoriaFisicaDisponivel())
            ' Dim TamanhoPente
            ' Dim WmiSelect As New ManagementObjectSearcher("root\CIMV2", "Select * from Win32_MemoryDevice")
            ' For Each WmiResults As ManagementObject In WmiSelect.Get()
            ' TamanhoPente = FormataNumero(WmiResults.GetPropertyValue("EndingAddress"))
            '.WriteElementString("PenteInstalado", TamanhoPente)
            'Next
            .WriteEndElement()
        End With

        ' Criar o Node Filho "DiscosLocais" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho DiscosLocais")
            .WriteStartElement("DiscosLocais")
            For Each drive As IO.DriveInfo In My.Computer.FileSystem.Drives
                If drive.IsReady = True And Not drive.DriveType = IO.DriveType.CDRom Then
                    .WriteElementString("Nome", DiscosLocais.Nome(drive.ToString))
                    '   .WriteElementString("Rotulo", DiscosLocais.Rotulo(drive.ToString))
                    .WriteElementString("Tipo", DiscosLocais.Tipo(drive.ToString))
                    .WriteElementString("Formato", DiscosLocais.Formato(drive.ToString))
                    .WriteElementString("TamanhoTotal", DiscosLocais.TamanhoTotal(drive.ToString))
                    .WriteElementString("TamanhoLivre", DiscosLocais.TamanhoDisponivel(drive.ToString))
                End If
                'If drive.DriveType = IO.DriveType.CDRom Then
                '.WriteElementString("Nome", DiscosLocais.Nome(drive.ToString))
                '   .WriteElementString("Rotulo", DiscosLocais.Rotulo(drive.ToString))
                '.WriteElementString("Tipo", DiscosLocais.Tipo(drive.ToString))
                ' .WriteElementString("Formato", DiscosLocais.Formato(drive.ToString))
                'End If
            Next drive
            .WriteEndElement()
        End With

        ' Criar o Node Filho "Rede" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho Rede")
            .WriteStartElement("Rede")
            If NetworkInterface.GetIsNetworkAvailable Then
                ' Obtem e define todos os objetos NetworkInterface para a maquina local
                Dim interfaces As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
                ' Percorre as interfaces
                For Each ni As NetworkInterface In interfaces
                    .WriteElementString("MacAddress", ni.GetPhysicalAddress().ToString())
                    .WriteElementString("Nome", ni.Name)
                    .WriteElementString("Descricao", ni.Description)
                    For Each addr As UnicastIPAddressInformation In ni.GetIPProperties.UnicastAddresses
                        .WriteElementString("EnderecoIP", addr.Address.ToString)
                    Next
                Next
            End If
            .WriteEndElement()
        End With

        ' Criar o Node Filho "Monitor" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho Monitor")
            .WriteStartElement("Monitor")
            .WriteElementString("PlacaVideo", Monitor.PlacaVideo)
            .WriteElementString("ResolucaoAtual", Monitor.ResolucaoTelaAtual)
            Dim ResolucaoPossivel As String = String.Empty
            'Dim WmiSelect As New ManagementObjectSearcher("root\CIMV2", "Select * from CIM_VideoControllerResolution")
            'For Each WmiResults As ManagementObject In WmiSelect.Get()
            'ResolucaoPossivel = WmiResults.GetPropertyValue("SettingID").ToString
            '.WriteElementString("ResolucaoPossivel", ResolucaoPossivel)
            'Next
            .WriteEndElement()
        End With

        ' Criar o Node Filho "Printers" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho Impressoras")
            .WriteStartElement("Impressoras")
            ' Find all printers installed
            For Each pkInstalledPrinters In PrinterSettings.InstalledPrinters
                .WriteElementString("Impressora", pkInstalledPrinters.ToString)
            Next pkInstalledPrinters
            .WriteEndElement()
        End With

        ' Criar o Node Filho "Produtos Microsoft" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho Produtos Microsoft")
            .WriteStartElement("ProdutosMicrosoft")
            .WriteElementString("Windows", ProdutosMicrosoft.ChaveWindows)
            .WriteElementString("Office", ProdutosMicrosoft.ChaveOffice)
            .WriteEndElement()
        End With

        ' Criar o Node Filho "ProgramasInstalados" e alguns dados
        With xmlw
            ' Adiciona um comentário Identificador Node Filho
            xmlw.WriteComment("Declaracao do Node Filho Programas Instalados")
            .WriteStartElement("ProgramasInstalados")
            Dim SubKey As RegistryKey
            'Abre a chave que consta os programas que possuem desinstalador
            Dim Key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", False)
            'Levanta Todos os Programas
            Dim SubKeyNames() As String = Key.GetSubKeyNames()
            'Lista todos os Programas
            For Index = 0 To Key.SubKeyCount - 1
                SubKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" & "\" & SubKeyNames(Index), False)
                'Valida se o Programa tem um Nome Valido para Exibiçao
                If Not SubKey.GetValue("DisplayName", "") Is "" Then
                    .WriteElementString("ProgramaInstalado", (CType(SubKey.GetValue("DisplayName", ""), String)))
                End If
            Next
            .WriteEndElement()
        End With

        xmlw.WriteEndElement() ' Insere Final XML Pai
        xmlw.WriteEndDocument()

        ' Fecha o documento XML
        xmlw.Flush()
        xmlw.Close()
        Return True
    End Function

    'Sub Main()
    'Do While True
    '       Thread.Sleep(25000)
    '       CriaXML()
    '       
    'Loop
    'End Sub
End Module
