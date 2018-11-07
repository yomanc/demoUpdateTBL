Imports System.IO
Imports System.Net
Imports System.Data
Imports System.Reflection
Imports demoUpdateEFTBL03.Logger


Public Class myFtpDownloader

    Dim FTP_CONF_SITE As String
    Dim FTP_CONF_UID As String
    Dim FTP_CONF_PWD As String
    Dim FTP_CONF_REMDIR As String
    Dim FTP_CONF_REMFNPREFIX As String
    Dim strReqUri As String = Nothing
    Dim aAwardnoArrFrmCVS(900000) As String
    Dim iAwardnoArrSzFrmCVS As Integer = 0
    Dim aBufCvsToday(ITEMMAXNUM) As String
    Dim iBufCvsTodayNum As Integer
    Dim sRemoFn As String = Nothing
    Dim sLatestCSVFn As String = Nothing
    Const ITEMMAXNUM As Integer = 23
    Dim loggerFtp = New Logger(LOCALDNLOADDIR, LOCALLOGFN)


    Public Sub New(fn As String)
        Dim fs As FileStream = Nothing
        Dim rs As StreamReader = Nothing
        Dim dat As String = Nothing
        Dim str As Array = Nothing
        Dim lines As String()
        Dim linetot As Integer = GetFileLine(fn)

        Try
            fs = New FileStream(fn, FileMode.Open, FileAccess.Read)
            rs = New StreamReader(fs)
            dat = rs.ReadToEnd()
            lines = System.Text.RegularExpressions.Regex.Split(dat, Environment.NewLine)

            For i As Integer = 0 To (linetot - 1)
                Dim idx As String = ""
                Dim key As String = ""
                str = Split(lines(i), "=")
                If IsNothing(str(0)) Or IsNothing(str(1)) Then
                    loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
                Else
                    idx = str(0)
                    key = str(1)
                End If

                If idx = "FTP_CONF_SITE" Then FTP_CONF_SITE = key
                If idx = "FTP_CONF_UID" Then FTP_CONF_UID = key
                If idx = "FTP_CONF_PWD" Then FTP_CONF_PWD = key
                If idx = "FTP_CONF_REMDIR" Then FTP_CONF_REMDIR = key
                If idx = "FTP_CONF_REMFNPREFIX" Then FTP_CONF_REMFNPREFIX = key
            Next
        Catch ex As Exception
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_UNDEF & INVALID_CONF_DESC)
        Finally
            If rs IsNot Nothing Then rs.Close()
            If fs IsNot Nothing Then fs.Close()
            rs = Nothing
            fs = Nothing
        End Try
    End Sub


    Public Function GetsLatestCSVFn() As String
        Return Me.sLatestCSVFn
    End Function


    Public Function SetsLatestCSVFn(ByVal s As String) As Boolean
        If IsNothing(s) Then
            Return False
        End If
        Me.sLatestCSVFn = s
        Return True
    End Function


    Public Function GetsRemoFn() As String
        Return Me.sRemoFn
    End Function

	
    Public Function GetaBufCvsToday() As String()
        Return Me.aBufCvsToday
    End Function


    Public Function GetaAwardnoArrFrmCVS() As String()
        Return Me.aAwardnoArrFrmCVS
    End Function


    Public Function GetHost() As String
        Return Me.FTP_CONF_SITE
    End Function


    Public Function GetUid() As String
        Return Me.FTP_CONF_UID
    End Function


    Public Function GetPWD() As String
        Return Me.FTP_CONF_PWD
    End Function


    Public Function GetFtpConfREMDIR() As String
        Return Me.FTP_CONF_REMDIR
    End Function


    Public Function GetFtpConfigREMFNPREFIX() As String
        Return Me.FTP_CONF_REMFNPREFIX
    End Function


    Public Function RetLatestCsvFn(ByVal t() As String, ByVal skipFnHeadFmt As String, ByVal FnMidFmt As String) As String
        Dim sCand As String = Nothing

        If IsNothing(t(0)) Or IsNothing(skipFnHeadFmt) Or IsNothing(FnMidFmt) Then
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
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
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_OF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return Nothing
        Catch ex As ArgumentOutOfRangeException
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_AOOR & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return Nothing
        Catch ex As Exception
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return Nothing
        End Try
    End Function


    Public Function DoParseCsvFile(ByRef Filename As String) As Boolean
        Const DLM As String = ","
        Const AWARDNOFIELDNO As Int32 = 0
        Dim aRowCurrent As String()
        Dim tfpMyReader As FileIO.TextFieldParser = Nothing

        loggerFtp.WriteLOG(LogLvl.Warn, "* " & "Filename: " & Filename, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

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

            loggerFtp.WriteLOG(LogLvl.Inf, "* " & "aAwardnoArrFrmCVS total " & aAwardnoArrFrmCVS.Length, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            ''aAwardnoArrFrmCVS.Distinct.ToArray()
            loggerFtp.WriteLOG(LogLvl.Dbg, "- " & "----------aAwardnoArrFrmCVS.s ", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            For Each tt As String In aAwardnoArrFrmCVS
                ''loggerFtp.WriteLOG(LogLvl.Dbg, "- " & tt, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Next
            loggerFtp.WriteLOG(LogLvl.Dbg, "- " & "----------aAwardnoArrFrmCVS.e ", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        Catch exArgu As System.ArgumentException
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_ARRUMENT & exArgu.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Catch exMfl As FileIO.MalformedLineException
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_MALFORMEDLINE & exMfl.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Catch ex As System.Exception
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Finally
            If Not tfpMyReader Is Nothing Then tfpMyReader.Close()
            tfpMyReader = Nothing
        End Try
        Return True
    End Function


    Public Function SetsRemoFn() As Boolean
        Const CONNTIMEOUT As Integer = 1000 * 10 ' ms
        Const RETRYRND As Integer = 10
        Dim RunRetry As Boolean = True
        Dim attempts As Integer = 0
        Dim wrsp As FtpWebResponse = Nothing
        Dim fwrspConn As FtpWebResponse = Nothing
        Dim responseStream As Stream = Nothing
        Dim rs As StreamReader = Nothing
        Dim fwreqConn As FtpWebRequest = Nothing
        Dim remote = "ftp://" & Me.getHost
        Dim uid = Me.getUid
        Dim pwd = Me.getPWD

        If IsNothing(remote) Or IsNothing(uid) Or IsNothing(pwd) Then
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_AOOR & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
            Return False
        End If

        Me.strReqUri = remote & getFtpConfREMDIR()

        Do While RunRetry
            fwreqConn = CType(FtpWebRequest.Create(Me.strReqUri), FtpWebRequest)
            fwreqConn.Credentials = New NetworkCredential(uid, pwd)
            fwreqConn.KeepAlive = True
            fwreqConn.UseBinary = True
            fwreqConn.UsePassive = False
            fwreqConn.Timeout = CONNTIMEOUT
            fwreqConn.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
            fwreqConn.Method = WebRequestMethods.Ftp.ListDirectory

            loggerFtp.WriteLOG(LogLvl.Dbg, "* " & "Connecting Uri : " & strReqUri, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Try
                wrsp = fwreqConn.GetResponse()
                fwrspConn = CType(wrsp, FtpWebResponse)

                loggerFtp.WriteLOG(LogLvl.Dbg, "< " & fwrspConn.StatusDescription, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                loggerFtp.WriteLOG(LogLvl.Dbg, "< " & fwrspConn.WelcomeMessage, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                responseStream = fwrspConn.GetResponseStream()
                rs = New StreamReader(responseStream)

                Dim sCVSfileTobecheck As String = Nothing
                Dim iIdxNowdateCVSfile As Integer = 0

                loggerFtp.WriteLOG(LogLvl.Dbg, "< " & fwrspConn.StatusDescription, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

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
                Dim remfnprefix As String = getFtpConfigREMFNPREFIX()

                Me.sRemoFn = RetLatestCsvFn(aBufCvsToday, remfnprefix, "yyyyMMddHHmm")

                Return True
            Catch exWebException As WebException
                If exWebException.Status = WebExceptionStatus.Timeout Then
                    If attempts < RETRYRND Then
                        attempts += 1
                        loggerFtp.WriteLOG(LogLvl.Warn, "! " & "conRetry, attempts#" & attempts, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Else
                        RunRetry = False
                        loggerFtp.WriteLOG(LogLvl.Warn, "! " & "conRetry out", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    End If
                Else
                    loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_WEB & exWebException.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    RunRetry = False
                    loggerFtp.WriteLOG(LogLvl.Err, "! " & "conn NG", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Return False
                End If
            Catch ex As Exception
                loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                RunRetry = False
                loggerFtp.WriteLOG(LogLvl.Err, "! " & "conn NG", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
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


    End Function


    Public Function getFileSize(ByVal uid As String, ByVal pwd As String, ByVal fp As String, ByVal fn As String) As Long
        Dim reqFtp As FtpWebRequest = CType(FtpWebRequest.Create(fp & fn), FtpWebRequest)
        Dim respFtp As FtpWebResponse = Nothing
        Dim sz As Long = 0
        Const RETRYRND As Integer = 5
        Const CONNTIMEOUT As Integer = 1000 * 10 'sec
        Dim attempts As Integer = 0
        Dim RunRetry As Boolean = True

        With reqFtp
            .Credentials = New NetworkCredential(uid, pwd)
            .KeepAlive = False
            .UseBinary = True
            .UsePassive = False
            .Timeout = CONNTIMEOUT
            .Method = WebRequestMethods.Ftp.GetFileSize
        End With

        Do While RunRetry
            Try
                respFtp = DirectCast(reqFtp.GetResponse(), FtpWebResponse)
                sz = respFtp.ContentLength()
                If sz > 0 Then
                    RunRetry = False
                End If
            Catch exWebException As WebException
                If exWebException.Status = WebExceptionStatus.Timeout Then
                    If attempts < RETRYRND Then
                        attempts += 1
                        loggerFtp.WriteLOG(LogLvl.Warn, "! " & "conRetry, attempts#" & attempts, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Else
                        RunRetry = False
                        loggerFtp.WriteLOG(LogLvl.Warn, "! " & "conRetry out", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    End If
                Else
                    loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_WEB & exWebException.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    RunRetry = False
                    ''Return 
                End If
            Catch ex As Exception
                loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                RunRetry = False
                ''Return
            Finally
                If Not respFtp Is Nothing Then respFtp.Dispose()
                respFtp = Nothing
            End Try
        Loop

        Return sz
    End Function

    
    Public Function downloadFile(ByRef remote As String, ByRef uid As String, ByRef pwd As String, ByRef local As String) As Integer
        Dim wrsp As WebResponse = Nothing
        Dim fwrspConn As FtpWebResponse = Nothing
        Dim CONNTIMEOUT As Integer = 1000 * 10 '10sec
        Dim rs As IO.Stream = Nothing
        Dim fs As IO.FileStream = Nothing
        Dim buffer(4095) As Byte
        Dim read As Integer = 0
        Dim fwreqConn As FtpWebRequest = Nothing

        If IsNothing(remote) Or IsNothing(uid) Or IsNothing(pwd) Or IsNothing(local) Then
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
            Return Nothing
        End If

        fwreqConn = CType(FtpWebRequest.Create("ftp://" & remote), FtpWebRequest)
        fwreqConn.Credentials = New NetworkCredential(uid, pwd)
        fwreqConn.KeepAlive = False
        fwreqConn.UseBinary = True
        fwreqConn.UsePassive = False
        fwreqConn.Timeout = CONNTIMEOUT
        fwreqConn.Method = WebRequestMethods.Ftp.DownloadFile

        loggerFtp.WriteLOG(LogLvl.Dbg, "! " & fwreqConn.RequestUri.ToString, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        Try
            fwrspConn = CType(fwreqConn.GetResponse(), FtpWebResponse)
            rs = fwrspConn.GetResponseStream
            fs = New IO.FileStream(local, IO.FileMode.Create)
            Do
                read = rs.Read(buffer, 0, buffer.Length)
                loggerFtp.WriteLOG(LogLvl.Dbg, "* " & "r " & read & " bytes", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                fs.Write(buffer, 0, read)
            Loop Until read = 0
            loggerFtp.WriteLOG(LogLvl.Warn, "* " & "dn " & fs.Position & " bytes", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            loggerFtp.WriteLOG(LogLvl.Dbg, "- " & fwrspConn.StatusDescription, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            fs.Flush()
            Dim dnSz As Integer = CInt(fs.Position)
            Return dnSz
        Catch exDirNotFound As DirectoryNotFoundException
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & exDirNotFound.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return 0
        Catch exWeb As WebException
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & exWeb.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return 0
        Catch exUnauthorizedAccess As UnauthorizedAccessException
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & exUnauthorizedAccess.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
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


    Private Function GetFileLine(ByVal fp As String) As Integer
        Dim sr As StreamReader = New StreamReader(fp)
        Dim strSplit() As String
        Dim fileCnt As String = sr.ReadToEnd()

        sr.Close()
        sr = Nothing
        strSplit = Split(fileCnt, vbCrLf)
		
        Return strSplit.GetUpperBound(0) + 1
    End Function


    ' [Depricated]
    Public Sub dbgTestConfigAll()
        loggerFtp.WriteLOG(LogLvl.Dbg, "* " & GetHost(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        loggerFtp.WriteLOG(LogLvl.Dbg, "* " & GetUid(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        loggerFtp.WriteLOG(LogLvl.Dbg, "* " & GetPWD(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    End Sub


    '[Depricated]Check if target fileName is today or not by certain format
    Public Function IsTodayCVSfile(ByVal s As String) As Boolean
        Dim cmpFmt As String = Nothing
        Dim cvsFileExtendS As String = Nothing

        If IsNothing(s) Then
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
            Return Nothing
        End If


        Dim remfnprefix As String = GetFtpConfigREMFNPREFIX()
        cvsFileExtendS = remfnprefix & "yyyyMMddHHmm"

        Try
            If s.Length < (Len(cvsFileExtendS) + Len(".csv")) Then
                loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_AOOR & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
            End If

            If (0 <> String.Compare(s.Substring(Len(remfnprefix), Len(cmpFmt)), DateTime.Now.ToString("yyyyMMdd"))) Then
                Return False
            End If

            If (0 <> String.Compare(s.Substring(Len(cvsFileExtendS), Len(".csv")), ".csv")) Then
                Return False
            End If
            Return True
        Catch ex As Exception
            loggerFtp.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End Try
    End Function

End Class
