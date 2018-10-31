Imports System.IO
Imports System.Net
Imports System.Data
Imports IBM.Data.DB2
Imports System.Reflection


Module ftpClient

    Const ITEMMAXNUM As Integer = 23
    Const REMFNMIDFMT As String = "yyyyMMddHHmm"
    Const REMFNTAILFMT As String = ".csv"
    Dim LOGDATETIME As String = DateTime.Now.ToString("yyyyMMddHHmmss")
    Dim LOGPREFIX As String = "LOG"
    Dim LOGFNEXTEND As String = ".log"
    Dim LOCALDNLOADDIR As String = LOGPREFIX & LOGDATETIME & "/"
    Dim LOCALLOGFN As String = LOGDATETIME & LOGFNEXTEND

    Const EXCEPT_UNDEF As String = "! Undef Exception: "
    Const EXCEPT_AOOR As String = "! ArgumentOutOfRangeException: "
    Const EXCEPT_OF As String = "! OverflowException: "
    Const EXCEPT_WEB As String = "! Web Exception: "
    Const EXCEPT_ARRUMENT As String = "! ArgumentException: "
    Const EXCEPT_MALFORMEDLINE As String = "! MalformedLine: "
    Const EXCEPT_BADIMAGEFORMAT As String = "! BadImageFormat: "
    Const EXCEPT_USRDEF As String = "! UsrDef Exception: "
    Const WRONG_FTP As String = "! FTP Error: "

    Dim loginConfFn As String = "config.txt"

    Dim FTP_CONF_SITE As String = Nothing
    Dim FTP_CONF_UID As String = Nothing
    Dim FTP_CONF_PWD As String = Nothing
    Dim FTP_CONF_REMDIR As String = Nothing
    Dim FTP_CONF_REMFNPREFIX As String = Nothing

    Dim DB_CONF_IP As String = Nothing
    Dim DB_CONF_PORT As String = Nothing
    Dim DB_CONF_DBNAME As String = Nothing
    Dim DB_CONF_UID As String = Nothing
    Dim DB_CONF_PWD As String = Nothing
    Dim DB_CONF_TO As String = Nothing ' Second.
    Dim DB_QURY_TABLE As String = Nothing
    Dim DB_QURY_FIELD As String = Nothing
    Dim DB_QUERY_DB2SEL_AWARDNO As String = Nothing

    Const ID_BUFFER_NUMMAX As UInt32 = 900000
    Dim sHostIP As String = Nothing
    Dim sUser As String = Nothing
    Dim sPwd As String = Nothing
    Dim aBufCvsToday(ITEMMAXNUM) As String
    Dim iBufCvsTodayNum As Integer
    Dim aAwardnoArrFrmDB(ID_BUFFER_NUMMAX) As String
    Dim iAwardnoArrSzFrmDB As Integer = 0
    Dim aAwardnoArrFrmCVS(ID_BUFFER_NUMMAX) As String
    Dim iAwardnoArrSzFrmCVS As Integer = 0

    Dim gROI() As String = Nothing ' Target items what we are looking for.
    Dim gDBconnStr As String = Nothing

    Dim gUserPause As Boolean = True
    Dim gDBG_ConsoleDispON As Boolean = False
    Dim gDBG_level As Integer = LogLvl.Dbg

    Enum LogLvl
        Err
        Warn
        Inf
        Dbg
    End Enum


    Sub Main(args As String())
        Dim sRemoFn As String = Nothing
        Dim sLatestCSVFn As String = Nothing

        If Not Directory.Exists(LOCALDNLOADDIR) Then
            My.Computer.FileSystem.CreateDirectory(LOCALDNLOADDIR)
        End If

        Console.WriteLine("Start execution")

        WriteLOG(LogLvl.Err, "+ " & "Misson.s", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        Try
            WriteLOG(LogLvl.Inf, "> " & "STEP1 set FTP login START", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            If SetAllConfig() = False Then
                WriteLOG(LogLvl.Err, "! " & "SetAllConfig NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            If SetFtpConfig(sHostIP, sUser, sPwd) = False Then
                WriteLOG(LogLvl.Err, "! " & "bDoFtpConfigCheck NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If

            WriteLOG(LogLvl.Inf, "> " & "STEP1 set FTP login END", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & "STEP2 select FTP data START", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            sRemoFn = RetFtpListLatestCSVfile(sHostIP, sUser, sPwd)
            WriteLOG(LogLvl.Dbg, "* " & "sRemoFn :" & sRemoFn, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If sRemoFn Is Nothing Then
                WriteLOG(LogLvl.Err, "! " & "sRemoFn NULL, Abort", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            WriteLOG(LogLvl.Dbg, "* " & "outputDir :" & System.AppDomain.CurrentDomain.BaseDirectory(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Dim remfnprefix As String = GetFtpConfigREMFNPREFIX()
            sLatestCSVFn = RetLatestCsvFn(aBufCvsToday, remfnprefix, REMFNMIDFMT)
            WriteLOG(LogLvl.Dbg, "* " & "sLatestCSVFn :" & sLatestCSVFn, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If sLatestCSVFn Is Nothing Then
                WriteLOG(LogLvl.Err, "! " & "sLatestCSVFn NULL", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If

            WriteLOG(LogLvl.Inf, "> " & "STEP2 select FTP data END", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & "STEP3 download FTP data START", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Dim remdir As String = GetFtpConfigREMDIR()
            If DownloadFtpFile(sHostIP & remdir & sRemoFn, sUser, sPwd, LOCALDNLOADDIR & sLatestCSVFn) <= 0 Then
                WriteLOG(LogLvl.Err, "! " & "DownloadFtpFile get 0 size or NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If

            WriteLOG(LogLvl.Inf, "> " & "STEP3 download FTP data END", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & "STEP4 parsing downloaded file START", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            If DoParseCsvFile(LOCALDNLOADDIR & sLatestCSVFn) = False Then
                WriteLOG(LogLvl.Err, "! " & "bDoParseCSVFile NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If

            WriteLOG(LogLvl.Inf, "> " & "STEP4 parsing downloaded file END", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & "STEP5 login database START", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            gDBconnStr = SetDbconnStr()
            DB_QUERY_DB2SEL_AWARDNO = SetDbQryDb2SelAwrdno()
            If DoDbSrvQuery(gDBconnStr, DB_QUERY_DB2SEL_AWARDNO) = False Then
                WriteLOG(LogLvl.Err, "! " & "bQueryDBSrv NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If

            WriteLOG(LogLvl.Inf, "> " & "STEP5 login database END", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & "STEP6 diff table refer to .cvs START", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            If DoSetDGroupForAwardNO(aAwardnoArrFrmDB, aAwardnoArrFrmCVS) = False Then
                WriteLOG(LogLvl.Err, "! " & "DoExceptBtwArrays2 NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If

            WriteLOG(LogLvl.Inf, "> " & "STEP6 diff table refer to .cvs END", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & "STEP7 sudo update table START", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            If DoUpdDBSrvForAwardNO(gDBconnStr) = False Then
                WriteLOG(LogLvl.Err, "! " & "bUpdDBSrvForAwardNO NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If

            WriteLOG(LogLvl.Inf, "> " & "STEP7 sudo update table END", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & "STEP8 officially update table START", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            If DoUpdDBSrvForAwardNODelete(gDBconnStr) = False Then
                WriteLOG(LogLvl.Err, "! " & "bUpdDBSrvForAwardNO NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If

            WriteLOG(LogLvl.Inf, "> " & "STEP8 officially update table END", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Err, "+ " & "Misson.e", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Console.WriteLine("End execution")

            If gDBG_ConsoleDispON = True Then
                Pause()
            End If
        Catch ex As Exception
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            WriteLOG(LogLvl.Err, "- " & "Misson.e.unexp", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Abort()
        End Try
    End Sub


    Public Function RetFtpListLatestCSVfile(ByVal remote As String, ByVal uid As String, ByVal pwd As String) As String
        Const CONNTIMEOUT As Integer = 1000 * 10 ' Set 10 sec.
        Const RETRYRND As Integer = 10
        Dim RunRetry As Boolean = True
        Dim attempts As Integer = 0
        Dim wrsp As FtpWebResponse = Nothing
        Dim fwrspConn As FtpWebResponse = Nothing
        Dim responseStream As Stream = Nothing
        Dim rs As StreamReader = Nothing
        Dim strReqUri As String = Nothing
        Dim fwreqConn As FtpWebRequest = Nothing

        If IsNothing(remote) Or IsNothing(uid) Or IsNothing(pwd) Then
            WriteLOG(LogLvl.Err, EXCEPT_AOOR & "fatal. invalid param", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid param")
            Return Nothing
        End If

        Dim remdir As String = GetFtpConfigREMDIR()
        Do While RunRetry
            strReqUri = remote & remdir
            fwreqConn = CType(FtpWebRequest.Create(strReqUri), FtpWebRequest)
            fwreqConn.Credentials = New NetworkCredential(uid, pwd)
            fwreqConn.KeepAlive = True
            fwreqConn.UseBinary = True
            fwreqConn.UsePassive = False
            fwreqConn.Timeout = CONNTIMEOUT
            fwreqConn.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
            fwreqConn.Method = WebRequestMethods.Ftp.ListDirectory

            WriteLOG(LogLvl.Dbg, "* " & "Connecting : " & strReqUri & " , Timeout : " & fwreqConn.Timeout & "ms", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Try
                wrsp = fwreqConn.GetResponse()
                fwrspConn = CType(wrsp, FtpWebResponse)

                WriteLOG(LogLvl.Dbg, "< " & fwrspConn.StatusDescription, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Dbg, "< " & fwrspConn.WelcomeMessage, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                responseStream = fwrspConn.GetResponseStream()
                rs = New StreamReader(responseStream)

                Dim sCVSfileTobecheck As String = Nothing
                Dim iIdxNowdateCVSfile As Integer = 0

                WriteLOG(LogLvl.Dbg, "< " & fwrspConn.StatusDescription, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                sCVSfileTobecheck = rs.ReadLine()
                While sCVSfileTobecheck IsNot Nothing
                    ''If IsTodayCVSfile(sCVSfileTobecheck) = True Then
                    aBufCvsToday(iIdxNowdateCVSfile) = sCVSfileTobecheck
                    iIdxNowdateCVSfile += 1
                    ''End If
                    sCVSfileTobecheck = rs.ReadLine()
                End While

                iBufCvsTodayNum = iIdxNowdateCVSfile
                RunRetry = False
                Dim remfnprefix As String = GetFtpConfigREMFNPREFIX()
                Return RetLatestCsvFn(aBufCvsToday, remfnprefix, REMFNMIDFMT)
            Catch exWebException As WebException
                If exWebException.Status = WebExceptionStatus.Timeout Then
                    If attempts < RETRYRND Then
                        attempts += 1
                        WriteLOG(LogLvl.Warn, "! " & "conRetry, attempts#" & attempts, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Else
                        RunRetry = False
                        WriteLOG(LogLvl.Warn, "! " & "conRetry out", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    End If
                Else
                    WriteLOG(LogLvl.Err, EXCEPT_WEB & exWebException.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    RunRetry = False
                    Return Nothing
                End If
            Catch ex As Exception
                WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                RunRetry = False
                Return Nothing
            Finally
                If Not wrsp Is Nothing Then wrsp.Dispose()
                If Not fwrspConn Is Nothing Then fwrspConn.Dispose()
                If Not rs Is Nothing Then rs.Dispose()
                If Not responseStream Is Nothing Then responseStream.Dispose()
                wrsp = Nothing
                fwrspConn = Nothing
                rs = Nothing
                responseStream = Nothing
            End Try
        Loop

        WriteLOG(LogLvl.Err, "! " & "conn NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        Return Nothing
    End Function


    Public Function DownloadFtpFile(ByRef remote As String, ByRef uid As String, ByRef pwd As String, ByRef local As String) As Integer
        Dim wrsp As WebResponse = Nothing
        Dim fwrspConn As FtpWebResponse = Nothing
        Dim CONNTIMEOUT As Integer = 1000 * 10 ' 10 sec
        Dim rs As IO.Stream = Nothing
        Dim fs As IO.FileStream = Nothing
        Dim buffer(2047) As Byte
        Dim read As Integer = 0
        Dim fwreqConn As FtpWebRequest = Nothing

        If IsNothing(remote) Or IsNothing(uid) Or IsNothing(pwd) Or IsNothing(local) Then
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & "fatal. invalid param", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid param")
            Return Nothing
        End If

        Try
            fwreqConn = CType(FtpWebRequest.Create(remote), FtpWebRequest)
            fwreqConn.Credentials = New NetworkCredential(uid, pwd)
            fwreqConn.KeepAlive = False
            fwreqConn.UseBinary = True
            fwreqConn.UsePassive = False
            fwreqConn.Timeout = CONNTIMEOUT
            fwreqConn.Method = WebRequestMethods.Ftp.DownloadFile

            WriteLOG(LogLvl.Dbg, "! " & fwreqConn.RequestUri.ToString, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            wrsp = fwreqConn.GetResponse()
            fwrspConn = CType(wrsp, FtpWebResponse)
            rs = fwrspConn.GetResponseStream
            fs = New IO.FileStream(local, IO.FileMode.Create)

            Do
                read = rs.Read(buffer, 0, buffer.Length)
                ''WriteLOG(LogLvl.Dbg, "* " & "r " & read & " bytes", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                fs.Write(buffer, 0, read)
            Loop Until read = 0

            WriteLOG(LogLvl.Inf, "* " & "dn " & fs.Position & " bytes", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Dbg, "- " & fwrspConn.StatusDescription, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            fs.Flush()

            Dim dnSz As Integer = CInt(fs.Position)

            Return dnSz
        Catch exDirNotFound As DirectoryNotFoundException
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & exDirNotFound.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return 0
        Catch exWeb As WebException
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & exWeb.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return 0
        Catch exUnauthorizedAccess As UnauthorizedAccessException
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & exUnauthorizedAccess.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return 0
        Finally
            If Not fwrspConn Is Nothing Then fwrspConn.Dispose()
            If Not wrsp Is Nothing Then wrsp.Dispose()
            If Not rs Is Nothing Then rs.Dispose()
            If Not fs Is Nothing Then fs.Dispose()
            fwrspConn = Nothing
            wrsp = Nothing
            rs = Nothing
            fs = Nothing
        End Try
    End Function


    Private Function SetAllConfig() As Boolean
        Dim fs As FileStream = Nothing
        Dim rs As StreamReader = Nothing
        Dim dat As String = Nothing
        Dim str As Array = Nothing
        Dim lines As String()

        Try
            fs = New FileStream(loginConfFn, FileMode.Open, FileAccess.Read)
            rs = New StreamReader(fs)

            dat = rs.ReadToEnd()
            lines = System.Text.RegularExpressions.Regex.Split(dat, Environment.NewLine)

            Dim Loginitem() As String = {
                "FTP_CONF_SITE",
                "FTP_CONF_UID",
                "FTP_CONF_PWD",
                "FTP_CONF_REMDIR",
                "FTP_CONF_REMFNPREFIX",
                "DB_CONF_IP",
                "DB_CONF_PORT",
                "DB_CONF_DBNAME",
                "DB_CONF_UID",
                "DB_CONF_PWD",
                "DB_CONF_TO",
                "DB_QURY_TABLE",
                "DB_QURY_FIELD"
                }

            str = Split(lines(0), "=")
            If str(0) = "FTP_CONF_SITE" Then FTP_CONF_SITE = str(1)
            str = Split(lines(1), "=")
            If str(0) = "FTP_CONF_UID" Then FTP_CONF_UID = str(1)
            str = Split(lines(2), "=")
            If str(0) = "FTP_CONF_PWD" Then FTP_CONF_PWD = str(1)
            str = Split(lines(3), "=")
            If str(0) = "FTP_CONF_REMDIR" Then FTP_CONF_REMDIR = str(1)
            str = Split(lines(4), "=")
            If str(0) = "FTP_CONF_REMFNPREFIX" Then FTP_CONF_REMFNPREFIX = str(1)
            str = Split(lines(5), "=")
            If str(0) = "DB_CONF_IP" Then DB_CONF_IP = str(1)
            str = Split(lines(6), "=")
            If str(0) = "DB_CONF_PORT" Then DB_CONF_PORT = str(1)
            str = Split(lines(7), "=")
            If str(0) = "DB_CONF_DBNAME" Then DB_CONF_DBNAME = str(1)
            str = Split(lines(8), "=")
            If str(0) = "DB_CONF_UID" Then DB_CONF_UID = str(1)
            str = Split(lines(9), "=")
            If str(0) = "DB_CONF_PWD" Then DB_CONF_PWD = str(1)
            str = Split(lines(10), "=")
            If str(0) = "DB_CONF_TO" Then DB_CONF_TO = str(1)
            str = Split(lines(11), "=")
            If str(0) = "DB_QURY_TABLE" Then DB_QURY_TABLE = str(1)
            str = Split(lines(12), "=")
            If str(0) = "DB_QURY_FIELD" Then DB_QURY_FIELD = str(1)
        Catch ex As Exception
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Finally
            If rs IsNot Nothing Then rs.Close()
            rs = Nothing
            If fs IsNot Nothing Then fs.Close()
            fs = Nothing
        End Try
        Return True
    End Function


    Private Function SetFtpConfig(ByRef h As String, ByRef u As String, ByRef p As String) As Boolean
        If (h IsNot Nothing) Or (u IsNot Nothing) Or (p IsNot Nothing) Then
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & "fatal. invalid param", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid param")
            Return Nothing
        End If

        h = "ftp://" & FTP_CONF_SITE
        u = FTP_CONF_UID
        p = FTP_CONF_PWD

        Try
            If IsNotProperConfig(h, u, p) = False Then
                WriteLOG(LogLvl.Err, WRONG_FTP & "Not properly config", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & "fatal. Not properly config")
                Return False
            End If
        Catch exFileNotFound As IO.FileNotFoundException
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & exFileNotFound.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Abort()
            Return False
        Catch ex As Exception
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Abort()
            Return False
        End Try
        Return True
    End Function


    Public Function IsNotProperConfig(ByVal h As String, ByVal u As String, ByVal p As String) As Boolean
        If IsNothing(h) Or IsNothing(u) Or IsNothing(p) Then
            Return False
        End If
        Return True
    End Function


    Public Function IsTodayCVSfile(ByVal sIn As String) As Boolean
        Dim cmpFmt As String = Nothing
        Dim sNowDatePtn As String = Nothing
        Dim cvsFileExtendS As String = Nothing

        If IsNothing(sIn) Then
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & "fatal. invalid param", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid param")
            Return Nothing
        End If

        cmpFmt = "yyyyMMdd"
        sNowDatePtn = DateTime.Now.ToString(cmpFmt)

        Dim remfnprefix As String = GetFtpConfigREMFNPREFIX()
        cvsFileExtendS = remfnprefix & REMFNMIDFMT

        Try
            If sIn.Length < (Len(cvsFileExtendS) + Len(REMFNTAILFMT)) Then
                WriteLOG(LogLvl.Err, EXCEPT_AOOR & "fatal. invalid param", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & "fatal. invalid param")
            End If

            If (0 <> String.Compare(sIn.Substring(Len(remfnprefix), Len(cmpFmt)), sNowDatePtn)) Then
                Return False
            End If

            If (0 <> String.Compare(sIn.Substring(Len(cvsFileExtendS), Len(REMFNTAILFMT)), REMFNTAILFMT)) Then
                Return False
            End If
            Return True
        Catch ex As Exception
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End Try
    End Function


    Public Function RetLatestCsvFn(ByVal t() As String, ByVal skipFnHeadFmt As String, ByVal FnMidFmt As String) As String
        Dim sCand As String = Nothing

        If IsNothing(t(0)) Or IsNothing(skipFnHeadFmt) Or IsNothing(FnMidFmt) Then
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & "fatal. invalid param", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid param")
            Return Nothing
        End If

        sCand = t(0)

        Try
            For i As Integer = 0 To t.GetLength(0) - 1
                If t(i) IsNot Nothing Then
                    Dim d1 As Double = Val(t(i).Substring(skipFnHeadFmt.Length, FnMidFmt.Length))
                    Dim d2 As Double = Val(sCand.Substring(skipFnHeadFmt.Length, FnMidFmt.Length))
                    If (d1 > d2) Then
                        sCand = t(i)
                    End If
                End If
            Next
            Return sCand
        Catch ex As OverflowException
            WriteLOG(LogLvl.Err, EXCEPT_OF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return Nothing
        Catch ex As ArgumentOutOfRangeException
            WriteLOG(LogLvl.Err, EXCEPT_AOOR & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return Nothing
        Catch ex As Exception
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return Nothing
        End Try
    End Function


    Public Function DoParseCsvFile(ByRef Filename As String) As Boolean
        Const DLM As String = ","
        Const AWARDNOFIELDNO As Int32 = 0
        Dim aRowCurrent As String()
        Dim tfpMyReader As FileIO.TextFieldParser = Nothing

        WriteLOG(LogLvl.Dbg, "* " & "Filename: " & Filename, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        Try
            tfpMyReader = New FileIO.TextFieldParser(Filename)
            tfpMyReader.TextFieldType = FileIO.FieldType.Delimited
            tfpMyReader.SetDelimiters(DLM)

            While Not tfpMyReader.EndOfData
                aRowCurrent = tfpMyReader.ReadFields()
                aAwardnoArrFrmCVS(iAwardnoArrSzFrmCVS) = aRowCurrent(AWARDNOFIELDNO)
                iAwardnoArrSzFrmCVS += 1
            End While

            ReDim Preserve aAwardnoArrFrmCVS(iAwardnoArrSzFrmCVS - 1)

            WriteLOG(LogLvl.Inf, "* " & "aAwardnoArrFrmCVS total " & aAwardnoArrFrmCVS.Length, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Dbg, "- " & "----------aAwardnoArrFrmCVS.s ", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            For Each tt As String In aAwardnoArrFrmCVS
                WriteLOG(LogLvl.Dbg, "- " & tt, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Next
            WriteLOG(LogLvl.Dbg, "- " & "----------aAwardnoArrFrmCVS.e ", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        Catch exArgu As System.ArgumentException
            WriteLOG(LogLvl.Err, EXCEPT_ARRUMENT & exArgu.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Catch exMfl As FileIO.MalformedLineException
            WriteLOG(LogLvl.Err, EXCEPT_MALFORMEDLINE & exMfl.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Catch ex As System.Exception
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Finally
            If Not tfpMyReader Is Nothing Then tfpMyReader.Close()
            tfpMyReader = Nothing
        End Try
        Return True
    End Function


    Public Function DoDbSrvQuery(ByVal constr As String, ByVal cmd As String) As Boolean
        Dim conn As DB2Connection = Nothing
        Dim command As DB2Command = Nothing
        Dim dr As DB2DataReader = Nothing
        Const DBTIMEOUT As Integer = 30 ' sec

        If (cmd Is Nothing) Or (constr Is Nothing) Then
            WriteLOG(LogLvl.Err, EXCEPT_USRDEF & "fatal. invalid param", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid param")
            Return False
        End If

        WriteLOG(LogLvl.Dbg, "* " & "constr: " & constr, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        Try
            conn = New DB2Connection(constr)
            conn.Open()
            If conn.IsOpen() = False Then
                WriteLOG(LogLvl.Err, "! " & "conn.Open NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            End If
            ''WriteLOG(LogLvl.Inf, "! " & "Conn SrvVer: " & conn.ServerVersion & " Database: " & conn.Database & " state: " & conn.State, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            command = conn.CreateCommand()
            command.CommandText = cmd
            command.CommandTimeout = DBTIMEOUT
            dr = command.ExecuteReader()

            If (dr.HasRows = True) Then
                While dr.Read()
                    For i As Integer = 0 To (dr.FieldCount - 1)
                        aAwardnoArrFrmDB(iAwardnoArrSzFrmDB) = dr.GetValue(i).ToString()
                        iAwardnoArrSzFrmDB += 1
                    Next
                End While
            End If

            ReDim Preserve aAwardnoArrFrmDB(iAwardnoArrSzFrmDB - 1)

            WriteLOG(LogLvl.Inf, "* " & "aAwardnoArrFrmDB total " & aAwardnoArrFrmDB.Length, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Dbg, "- " & "----------aAwardnoArrFrmDB.s", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            For Each tt As String In aAwardnoArrFrmDB
                WriteLOG(LogLvl.Dbg, "- " & tt, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Next
            WriteLOG(LogLvl.Dbg, "- " & "----------aAwardnoArrFrmDB.e", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Return True

        Catch ex As BadImageFormatException
            Console.WriteLine(ex.Message)
            WriteLOG(LogLvl.Err, EXCEPT_BADIMAGEFORMAT & ex.Message.Substring(0, ex.Message.IndexOf(".") + 1), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine(ex.Message)
            WriteLOG(LogLvl.Err, EXCEPT_BADIMAGEFORMAT & ex.Message.Substring(0, ex.Message.IndexOf(".") + 1), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Finally
            If Not conn Is Nothing Then conn.Dispose()
            If Not dr Is Nothing Then dr.Close()
            conn = Nothing
            dr = Nothing
        End Try
    End Function


    Public Function DoUpdDBSrvForAwardNO(ByVal constr As String) As Boolean
        Const DBTIMEOUT As Integer = 30 ' sec
        Dim sQryOptr As String = Nothing
        Dim conn As DB2Connection = Nothing
        Dim dr As DB2DataReader = Nothing
        Dim cmd As DB2Command = Nothing

        sQryOptr = "UPDATE " & DB_QURY_TABLE & " Set " & "FLAG='Y' " & "Where " & DB_QURY_FIELD & "="

        WriteLOG(LogLvl.Dbg, "! " & "gROI.Length: " & gROI.Length, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        For Each tt As String In gROI
            Try
                conn = New DB2Connection(constr)
                conn.Open()
                If conn.IsOpen() = False Then
                    WriteLOG(LogLvl.Err, "! " & "conn.Open NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Return False
                End If
                ''WriteLOG(LogLvl.Inf, "! " & "Conn SrvVer: " & conn.ServerVersion & " Database: " & conn.Database & " state: " & conn.State, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                cmd = conn.CreateCommand()
                cmd.CommandText = sQryOptr & "'" & tt & "'" & ";"
                cmd.CommandTimeout = DBTIMEOUT
                WriteLOG(LogLvl.Dbg, "* " & cmd.CommandText, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                dr = cmd.ExecuteReader()
                If (dr.HasRows = True) Then
                    While dr.Read()
                        WriteLOG(LogLvl.Err, dr.GetValue(0).ToString, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    End While
                End If
            Catch exBadImageFormat As BadImageFormatException
                WriteLOG(LogLvl.Err, EXCEPT_BADIMAGEFORMAT & exBadImageFormat.Message.Substring(0, exBadImageFormat.Message.IndexOf(".") + 1), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            Catch ex As Exception
                WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            Finally
                If Not conn Is Nothing Then conn.Dispose()
                If Not dr Is Nothing Then dr.Close()
                conn = Nothing
                dr = Nothing
            End Try
        Next
        Return True
    End Function


    Public Function DoUpdDBSrvForAwardNODelete(ByVal constr As String) As Boolean
        Const DBTIMEOUT As Integer = 30 ' sec
        Dim sQryOptr As String = Nothing
        Dim conn As DB2Connection = Nothing
        Dim dr As DB2DataReader = Nothing
        Dim cmd As DB2Command = Nothing

        sQryOptr = "delte from " & DB_QURY_TABLE & " where " & "(" & "FLAG='Y'" & " and " & DB_QURY_FIELD & "="

        WriteLOG(LogLvl.Dbg, "! " & "gROI.Length: " & gROI.Length, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        For Each tt As String In gROI
            Try
                conn = New DB2Connection(constr)
                conn.Open()
                If conn.IsOpen() = False Then
                    WriteLOG(LogLvl.Err, "! " & "conn.Open NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Return False
                End If
                ''WriteLOG(LogLvl.Inf, "! " & "Conn SrvVer: " & conn.ServerVersion & " Database: " & conn.Database & " state: " & conn.State, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                cmd = conn.CreateCommand()
                cmd.CommandText = sQryOptr & "'" & tt & "'" & ")" & ";"
                cmd.CommandTimeout = DBTIMEOUT
                WriteLOG(LogLvl.Dbg, "* " & cmd.CommandText, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                dr = cmd.ExecuteReader()
                If (dr.HasRows = True) Then
                    While dr.Read()
                        WriteLOG(LogLvl.Err, dr.GetValue(0).ToString, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    End While
                End If
            Catch exBadImageFormat As BadImageFormatException
                WriteLOG(LogLvl.Err, EXCEPT_BADIMAGEFORMAT & exBadImageFormat.Message.Substring(0, exBadImageFormat.Message.IndexOf(".") + 1), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            Catch ex As Exception
                WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            Finally
                If Not conn Is Nothing Then conn.Dispose()
                If Not dr Is Nothing Then dr.Close()
                conn = Nothing
                dr = Nothing
            End Try
        Next
        Return True
    End Function


    Public Function DoSetDGroupForAwardNO(ByVal listA As String(), ByVal listB As String()) As Boolean
        Try
            Dim aDiff = listA.Except(listB).ToArray()
            ReDim Preserve gROI(aDiff.Length - 1)
            gROI = aDiff

            WriteLOG(LogLvl.Inf, "* " & "gROI total " & gROI.Length, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            WriteLOG(LogLvl.Dbg, "- " & "----------gROI.s ", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            For Each tt As String In gROI
                WriteLOG(LogLvl.Dbg, "- " & tt, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Next
            WriteLOG(LogLvl.Dbg, "- " & "----------gROI.e ", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return True
        Catch ex As Exception
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End Try
    End Function


    Public Sub WriteLOG(ByVal hint As LogLvl, ByVal Msg As String, ByVal exceptLine As String)
        Dim currentName As String = Nothing
        Dim callName As String = Nothing
        Dim sPrefix As String = Nothing
        Dim aFile As FileStream = Nothing
        Dim sw As StreamWriter = Nothing

        Select Case hint
            Case LogLvl.Err
                Console.ForegroundColor = ConsoleColor.Red
                sPrefix = "err"
            Case LogLvl.Warn
                Console.ForegroundColor = ConsoleColor.Yellow
                sPrefix = "warn"
            Case LogLvl.Inf
                Console.ForegroundColor = ConsoleColor.White
                sPrefix = "inf"
            Case LogLvl.Dbg
                Console.ForegroundColor = ConsoleColor.Green
                sPrefix = "dbg"
            Case Else
                Throw New ApplicationException("LogLvl is invalid")
        End Select

        If gDBG_level >= hint Then
            currentName = New StackTrace(True).GetFrame(0).GetMethod().Name
            callName = New StackTrace(True).GetFrame(1).GetMethod().Name
            Try
                aFile = New FileStream(LOCALDNLOADDIR & LOCALLOGFN, FileMode.OpenOrCreate)
                aFile.Seek(0, SeekOrigin.End)
                sw = New StreamWriter(aFile)
                sw.WriteLine("{0}{1} {2} {3} @{4}#{5}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString(), sPrefix, Msg.Replace(vbCrLf, ""), callName, exceptLine)

                If gDBG_ConsoleDispON = True Then
                    DBG_display(Msg)
                End If
            Catch ex As Exception
                Console.WriteLine("Fatal. Error in WriteLOG")
                Throw New Exception("Fatal. Error in WriteLOG")
                ''Pause()
            Finally
                If Not sw Is Nothing Then sw.Close()
                If Not aFile Is Nothing Then aFile.Close()
                sw = Nothing
                aFile = Nothing
            End Try
        End If
    End Sub


    Public Sub Pause()
        Console.WriteLine("Any Key to exit ...")
        Console.Read()
    End Sub


    Public Sub Abort()
        Console.Write("Bat " & Assembly.GetExecutingAssembly.GetName.Name & " NOT executed properly. ")
        Console.WriteLine("Any Key to exit ...")
        Console.Read()
    End Sub


    Sub DBG_display(ByVal s As String)
        If gDBG_ConsoleDispON = True Then
            Console.WriteLine(s)
        End If
    End Sub


    Private Function GetFtpConfigSITE() As String
        Return FTP_CONF_SITE
    End Function


    Private Function GetFtpConfigUID() As String
        Return FTP_CONF_UID
    End Function


    Private Function GetFtpConfigPWD() As String
        Return FTP_CONF_PWD
    End Function


    Private Function GetFtpConfigREMDIR()
        Return FTP_CONF_REMDIR
    End Function


    Private Function GetFtpConfigREMFNPREFIX() As String
        Return FTP_CONF_REMFNPREFIX
    End Function


    Private Function GetDbConfIP() As String
        Return DB_CONF_IP
    End Function


    Private Function GetDbConfPORT() As String
        Return DB_CONF_PORT
    End Function


    Private Function GetDbConfDBNAME() As String
        Return DB_CONF_DBNAME
    End Function


    Private Function GetDbConfUID() As String
        Return DB_CONF_UID
    End Function


    Private Function GetDbConfPWD() As String
        Return DB_CONF_PWD
    End Function


    Private Function GetDBConfTO() As String
        Return DB_CONF_TO
    End Function


    Private Function GetDbQryTABLE() As String
        Return DB_QURY_TABLE
    End Function


    Private Function GetDbQryFIELD() As String
        Return DB_QURY_FIELD
    End Function


    Private Function SetDbconnStr() As String
        Dim ip As String = GetDbConfIP()
        Dim port As String = GetDbConfPORT()
        Dim dbName As String = GetDbConfDBNAME()
        Dim uid As String = GetDbConfUID()
        Dim pwd As String = GetDbConfPWD()
        Dim timeout As String = GetDBConfTO()

        If (IsHostIpValid(ip) = False) Or (IsPortValid(port) = False) Or (IsDbNameValid(dbName) = False) Or IsNothing(uid) Or IsNothing(pwd) Or IsNothing(timeout) Then
            WriteLOG(LogLvl.Err, EXCEPT_USRDEF & "fatal. invalid config", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid config")
            Return Nothing
        End If
        Return "Server=" & ip & ":" & port & ";" & "Database=" & dbName & ";" & "UID=" & uid & ";" & "PWD=" & pwd & ";" & "Connect Timeout=" & timeout
    End Function


    Private Function SetDbQryDb2SelAwrdno() As String
        Dim fld As String = GetDbQryFIELD()
        Dim tbl As String = GetDbQryTABLE()

        If IsNothing(fld) Or IsNothing(tbl) Then
            WriteLOG(LogLvl.Err, EXCEPT_USRDEF & "fatal. invalid config", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid config")
            Return Nothing
        End If
        Return "select " & fld & " from " & tbl & ";"
    End Function


    Private Function IsHostIpValid(ByVal inIP As String) As Boolean
        Dim ipaddr As IPAddress = Nothing
        If IsNothing(inIP) Then
            WriteLOG(LogLvl.Err, EXCEPT_USRDEF & "fatal. invalid config", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid config")
            Return False
        End If
        Dim is_valid As Boolean = IPAddress.TryParse(inIP, ipaddr)
        Return is_valid
    End Function


    Private Function IsPortValid(ByVal port As String) As Boolean
        If IsNothing(port) Then
            WriteLOG(LogLvl.Err, EXCEPT_USRDEF & "fatal. invalid config", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid config")
            Return False
        End If
        Dim portNum As Integer = CInt(port)
        If (portNum < 0) Or (portNum > 65536) Then
            Return False
        End If
        Return True
    End Function


    Private Function IsDbNameValid(ByVal db As String) As Boolean
        If IsNothing(db) Then
            WriteLOG(LogLvl.Err, EXCEPT_USRDEF & "fatal. invalid config", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & "fatal. invalid config")
            Return False
        End If
        Dim bIsValid As Boolean = db.StartsWith("UGD", True, Nothing)
        Return bIsValid
    End Function

End Module