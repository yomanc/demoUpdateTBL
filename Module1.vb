Imports System.IO
Imports System.Net
Imports System.Data
Imports IBM.Data.DB2
Imports System.Reflection


Module Module1

    Const ITEMMAXNUM As Integer = 23
    Const REMFNMIDFMT As String = "yyyyMMddHHmm"
    Const REMFNTAILFMT As String = ".csv"

    Dim LOGPREFIX As String = "LOG"
    Dim LOGDATETIME As String = DateTime.Now.ToString("yyyyMMddHHmmss")
    Dim LOCALDNLOADDIR As String = LOGPREFIX & LOGDATETIME & "/"
    Dim LOGFNEXTEND As String = ".log"
    Dim LOCALLOGFN As String = LOGDATETIME & LOGFNEXTEND

    Dim FTPPREFIX As String = "ftp://"
    Dim loginConfFn As String = "config.txt"

    Const EXCEPT_UNDEF As String = "! Undef Exception: "
    Const EXCEPT_AOOR As String = "! ArgumentOutOfRangeException: "
    Const EXCEPT_OF As String = "! OverflowException: "
    Const EXCEPT_WEB As String = "! Web Exception: "
    Const EXCEPT_ARRUMENT As String = "! ArgumentException: "
    Const EXCEPT_MALFORMEDLINE As String = "! MalformedLine: "
    Const EXCEPT_BADIMAGEFORMAT As String = "! BadImageFormat: "
    Const EXCEPT_USRDEF As String = "! UsrDef Exception: "
    Const WRONG_FTP As String = "! FTP Error: "
    
    Const INVALID_PARM_DESC As String = "fatal. invalid param"
    Const INVALID_CONF_DESC As String = "fatal. invalid config"
    Const NULL_INPUT_DESC As String = "warn. null input"
    Const MISSION_DESC_S As String = "Misson.s"
    Const MISSION_DESC_E_UNEXP As String = "Misson.e.unexp"
    Const MISSION_DESC_E_DONE As String = "Misson.e.done"
    
    Const HINT_STEP1_S As String = "STEP1 set FTP login conf S"
    Const HINT_STEP1_E As String = "STEP1 set FTP login conf E"
    Const HINT_STEP2_S As String = "STEP2 select FTP data S"
    Const HINT_STEP2_E As String = "STEP2 select FTP data E"
    Const HINT_STEP3_S As String = "STEP3 download FTP data S"
    Const HINT_STEP3_E As String = "STEP3 download FTP data E"
    Const HINT_STEP4_S As String = "STEP4 parsing downloaded file S"
    Const HINT_STEP4_E As String = "STEP4 parsing downloaded file E"
    Const HINT_STEP5_S As String = "STEP5 login database S"
    Const HINT_STEP5_E As String = "STEP5 login database E"
    Const HINT_STEP6_S As String = "STEP6 diff table refer To .cvs S"
    Const HINT_STEP6_E As String = "STEP6 diff table refer To .cvs E"
    Const HINT_STEP7_S As String = "STEP7 sudo update table S"
    Const HINT_STEP7_E As String = "STEP7 sudo update table E"
    Const HINT_STEP8_S As String = "STEP8 officially update table S"
    Const HINT_STEP8_E As String = "STEP8 officially update table E"

    Dim gUserPause As Boolean = True
    Dim gDBG_ConsoleDispON As Boolean = True
    Dim gDBG_level As Integer = LogLvl.Dbg

    Enum LogLvl
        Err
        Warn
        Inf
        Dbg
    End Enum


    Sub Main(args As String())
        If Not Directory.Exists(LOCALDNLOADDIR) Then
            My.Computer.FileSystem.CreateDirectory(LOCALDNLOADDIR)
        End If

        Console.WriteLine("Start execution")
		
        WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_S, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        Dim ftpdn = New myFtpDownloader
        Dim dbupdter = New myDbUpdater

        Try
            WriteLOG(LogLvl.Inf, "> " & HINT_STEP1_S, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            ftpdn.SetConf()
            ftpdn.dbgTestFtpConf()
            WriteLOG(LogLvl.Inf, "> " & HINT_STEP1_E, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & HINT_STEP2_S, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            ftpdn.SetsRemoFn()
            WriteLOG(LogLvl.Dbg, "* " & "sRemoFn :" & ftpdn.GetsRemoFn, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.GetsRemoFn Is Nothing Then
                WriteLOG(LogLvl.Err, "! " & "sRemoFn NULL, Abort", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            WriteLOG(LogLvl.Dbg, "* " & "outputDir :" & System.AppDomain.CurrentDomain.BaseDirectory(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Dim remfnprefix As String = ftpdn.getFtpConfigREMFNPREFIX()
            ftpdn.SetsLatestCSVFn(ftpdn.RetLatestCsvFn(ftpdn.GetaBufCvsToday, remfnprefix, REMFNMIDFMT))
            WriteLOG(LogLvl.Dbg, "* " & "sLatestCSVFn :" & ftpdn.GetsLatestCSVFn, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.GetsLatestCSVFn Is Nothing Then
                WriteLOG(LogLvl.Err, "! " & "sLatestCSVFn NULL", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            WriteLOG(LogLvl.Inf, "> " & HINT_STEP2_E, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & HINT_STEP3_S, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If DownloadFtpFile(FTPPREFIX & ftpdn.getHost() & ftpdn.getFtpConfigREMDIR & ftpdn.GetsRemoFn, ftpdn.getUid(), ftpdn.getPWD(), LOCALDNLOADDIR & ftpdn.GetsLatestCSVFn) <= 0 Then
                WriteLOG(LogLvl.Err, "! " & "DownloadFtpFile get 0 size or NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            WriteLOG(LogLvl.Inf, "> " & HINT_STEP3_E, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & HINT_STEP4_S, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.DoParseCsvFile(LOCALDNLOADDIR & ftpdn.GetsLatestCSVFn) = False Then
                WriteLOG(LogLvl.Err, "! " & "bDoParseCSVFile NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            WriteLOG(LogLvl.Inf, "> " & HINT_STEP4_E, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & HINT_STEP5_S, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            dbupdter.SetDbConfig()
            dbupdter.SetDbconnStr()
            dbupdter.SetDbQryDb2SelAwrdno()
            If dbupdter.DoDbSrvQuery(dbupdter.GetDbconnStr(), dbupdter.GetDbQryDb2SelAwrdno()) = False Then
                WriteLOG(LogLvl.Err, "! " & "bQueryDBSrv NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            WriteLOG(LogLvl.Inf, "> " & HINT_STEP5_E, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & HINT_STEP6_S, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If dbupdter.DoSetDGroupForAwardNO(dbupdter.GetaAwardnoArrFrmDB(), ftpdn.GetaAwardnoArrFrmCVS()) = False Then
                WriteLOG(LogLvl.Err, "! " & "DoExceptBtwArrays2 NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            WriteLOG(LogLvl.Inf, "> " & HINT_STEP6_E, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & HINT_STEP7_S, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If dbupdter.DoUpdDBSrvForAwardNO(dbupdter.GetDbconnStr()) = False Then
                WriteLOG(LogLvl.Err, "! " & "bUpdDBSrvForAwardNO NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_UNEXP , New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            WriteLOG(LogLvl.Inf, "> " & HINT_STEP7_E , New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Inf, "> " & HINT_STEP8_S , New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If dbupdter.DoUpdDBSrvForAwardNODelete(dbupdter.GetDbconnStr()) = False Then
                WriteLOG(LogLvl.Err, "! " & "bUpdDBSrvForAwardNO NG", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_UNEXP , New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Abort()
                Return
            End If
            WriteLOG(LogLvl.Inf, "> " & HINT_STEP8_E , New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_DONE, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Console.WriteLine("End execution")
			
            If gDBG_ConsoleDispON = True Then
                Pause()
            End If
        Catch ex As Exception
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            WriteLOG(LogLvl.Err, "- " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Abort()
        End Try
    End Sub


    Public Function DownloadFtpFile(ByRef remote As String, ByRef uid As String, ByRef pwd As String, ByRef local As String) As Integer
        Dim wrsp As WebResponse = Nothing
        Dim fwrspConn As FtpWebResponse = Nothing
        Dim CONNTIMEOUT As Integer = 1000 * 10 '10sec
        Dim rs As IO.Stream = Nothing
        Dim fs As IO.FileStream = Nothing
        Dim buffer(2047) As Byte
        Dim read As Integer = 0
        Dim fwreqConn As FtpWebRequest = Nothing

        If IsNothing(remote) Or IsNothing(uid) Or IsNothing(pwd) Or IsNothing(local) Then
            WriteLOG(LogLvl.Err, EXCEPT_UNDEF & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
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


    Public Function IsNotProperConfig(ByVal h As String, ByVal u As String, ByVal p As String) As Boolean
        If IsNothing(h) Or IsNothing(u) Or IsNothing(p) Then
            Return False
        End If
        Return True
    End Function


    Class myFtpDownloader
	
        Dim FTP_CONF_SITE As String
        Dim FTP_CONF_UID As String
        Dim FTP_CONF_PWD As String
        Dim FTP_CONF_REMDIR As String
        Dim FTP_CONF_REMFNPREFIX As String

        Dim RunRetry As Boolean = True
        Dim attempts As Integer = 0

        Dim responseStream As Stream = Nothing
        Dim rs As StreamReader = Nothing
        Dim strReqUri As String = Nothing
        Dim fwrspConn As FtpWebResponse = Nothing
        Dim fwreqConn As FtpWebRequest = Nothing

        Dim aAwardnoArrFrmCVS(900000) As String
        Dim iAwardnoArrSzFrmCVS As Integer = 0

        Dim aBufCvsToday(ITEMMAXNUM) As String
        Dim iBufCvsTodayNum As Integer

        Dim sRemoFn As String = Nothing
        Dim sLatestCSVFn As String = Nothing

		
        Public Function GetsLatestCSVFn() As String
            Return Me.sLatestCSVFn
        End Function
		
		
        Public Sub SetsLatestCSVFn(ByVal s As String)
            Me.sLatestCSVFn = s
        End Sub

		
        Public Function GetsRemoFn() As String
            Return Me.sRemoFn
        End Function
		
		
        Public Function GetaBufCvsToday() As String()
            Return Me.aBufCvsToday
        End Function
		
		
        Public Function GetaAwardnoArrFrmCVS() As String()
            Return Me.aAwardnoArrFrmCVS
        End Function

		
        Public Sub setHost(ByVal s As String)
            Me.FTP_CONF_SITE = s
        End Sub

		
        Public Function getHost() As String
            Return Me.FTP_CONF_SITE
        End Function

		
        Public Sub setUid(ByVal s As String)
            Me.FTP_CONF_UID = s
        End Sub

		
        Public Function getUid() As String
            Return Me.FTP_CONF_UID
        End Function

		
        Public Sub setPwd(ByVal s As String)
            Me.FTP_CONF_PWD = s
        End Sub

		
        Public Function getPWD() As String
            Return Me.FTP_CONF_PWD
        End Function

		
        Public Sub setFtpConfigREMDIR(ByVal s As String)
            Me.FTP_CONF_REMDIR = s
        End Sub

		
        Public Function getFtpConfigREMDIR() As String
            Return Me.FTP_CONF_REMDIR
        End Function

		
        Public Function getFtpConfigREMFNPREFIX() As String
            Return Me.FTP_CONF_REMFNPREFIX
        End Function

		
        Public Sub setFtpConfigREMFNPREFIX(ByVal s As String)
            Me.FTP_CONF_REMFNPREFIX = s
        End Sub

		
        Public Sub SetConf()
            Dim fs As FileStream = Nothing
            Dim rs As StreamReader = Nothing
            Dim dat As String = Nothing
            Dim str As Array = Nothing
            Dim lines As String()
            Dim linetot As Integer = GetFileLine(loginConfFn)

            Try
                fs = New FileStream(loginConfFn, FileMode.Open, FileAccess.Read)
                rs = New StreamReader(fs)

                dat = rs.ReadToEnd()
                lines = System.Text.RegularExpressions.Regex.Split(dat, Environment.NewLine)

                For i As Integer = 0 To (linetot - 1)
                    Dim idx As String = ""
                    Dim key As String = ""
                    str = Split(lines(i), "=")
                    If IsNothing(str(0)) Or IsNothing(str(1)) Then
                        WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
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
                WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_UNDEF & INVALID_CONF_DESC)
            Finally
                If rs IsNot Nothing Then rs.Close()
                If fs IsNot Nothing Then fs.Close()
                rs = Nothing
                fs = Nothing
            End Try
        End Sub

        
        Public Function IsTodayCVSfile(ByVal s As String) As Boolean
            Dim cmpFmt As String = Nothing
            Dim sNowDatePtn As String = Nothing
            Dim cvsFileExtendS As String = Nothing

            If IsNothing(s) Then
                WriteLOG(LogLvl.Err, EXCEPT_UNDEF & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
                Return Nothing
            End If

            cmpFmt = "yyyyMMdd"
            sNowDatePtn = DateTime.Now.ToString(cmpFmt)

            Dim remfnprefix As String = getFtpConfigREMFNPREFIX()
            cvsFileExtendS = remfnprefix & REMFNMIDFMT

            Try
                If s.Length < (Len(cvsFileExtendS) + Len(REMFNTAILFMT)) Then
                    WriteLOG(LogLvl.Err, EXCEPT_AOOR & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
                End If

                If (0 <> String.Compare(s.Substring(Len(remfnprefix), Len(cmpFmt)), sNowDatePtn)) Then
                    Return False
                End If

                If (0 <> String.Compare(s.Substring(Len(cvsFileExtendS), Len(REMFNTAILFMT)), REMFNTAILFMT)) Then
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
                WriteLOG(LogLvl.Err, EXCEPT_UNDEF & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
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
                ''aAwardnoArrFrmCVS.Distinct.ToArray()
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


        Public Sub SetsRemoFn()
            Const CONNTIMEOUT As Integer = 1000 * 10 ' sec
            Const RETRYRND As Integer = 10
            Dim RunRetry As Boolean = True
            Dim attempts As Integer = 0
            Dim wrsp As FtpWebResponse = Nothing
            Dim fwrspConn As FtpWebResponse = Nothing
            Dim responseStream As Stream = Nothing
            Dim rs As StreamReader = Nothing
            Dim strReqUri As String = Nothing
            Dim fwreqConn As FtpWebRequest = Nothing

            Dim remote = FTPPREFIX & Me.getHost
            Dim uid = Me.getUid
            Dim pwd = Me.getPWD

            If IsNothing(remote) Or IsNothing(uid) Or IsNothing(pwd) Then
                WriteLOG(LogLvl.Err, EXCEPT_AOOR & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
                ''Return
            End If

            Do While RunRetry
                strReqUri = remote & getFtpConfigREMDIR()
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
                    Dim remfnprefix As String = getFtpConfigREMFNPREFIX()

                    Me.sRemoFn = RetLatestCsvFn(aBufCvsToday, remfnprefix, REMFNMIDFMT)
                    Return
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
                        ''Return
                    End If
                Catch ex As Exception
                    WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    RunRetry = False
                    ''Return
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
        End Sub


        Public Function doFtpConn() As Boolean
            Return True
        End Function


        Public Sub dbgTestFtpConf()
            WriteLOG(LogLvl.Dbg, "* " & Me.getHost(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            WriteLOG(LogLvl.Dbg, "* " & Me.getUid(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            WriteLOG(LogLvl.Dbg, "* " & Me.getPWD(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        End Sub

    End Class


    Class myDbUpdater

        Dim DB_CONF_IP As String = Nothing
        Dim DB_CONF_PORT As String = Nothing
        Dim DB_CONF_DBNAME As String = Nothing
        Dim DB_CONF_UID As String = Nothing
        Dim DB_CONF_PWD As String = Nothing
        Dim DB_CONF_TO As String = Nothing ' second
        Dim DB_QURY_TABLE As String = Nothing
        Dim DB_QURY_FIELD As String = Nothing
        Dim DB_QUERY_DB2SEL_AWARDNO As String = Nothing
        Dim DB_CONNSTR As String = Nothing

        Dim aAwardnoArrFrmDB(900000) As String
        Dim iAwardnoArrSzFrmDB As Integer = 0

        Dim gROI() As String = Nothing


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

		
        Public Function GetaAwardnoArrFrmDB() As String()
            Return Me.aAwardnoArrFrmDB
        End Function

		
        Private Function GetDbConfIP() As String
            Return Me.DB_CONF_IP
        End Function

		
        Private Sub SetDbConfIP(ByVal i As String)
            If IsNothing(i) Then
                WriteLOG(LogLvl.Warn, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            End If
            Me.DB_CONF_IP = i
        End Sub

		
        Private Function GetDbConfPORT() As String
            Return Me.DB_CONF_PORT
        End Function

		
        Private Sub SetDbConfPORT(ByVal i As String)
            If IsNothing(i) Then
                WriteLOG(LogLvl.Warn, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            End If
            Me.DB_CONF_PORT = i
        End Sub

		
        Private Function GetDbConfDBNAME() As String
            Return Me.DB_CONF_DBNAME
        End Function

		
        Private Sub SetDbConfDBNAME(ByVal i As String)
            If IsNothing(i) Then
                WriteLOG(LogLvl.Warn, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            End If
            Me.DB_CONF_DBNAME = i
        End Sub

		
        Private Function GetDbConfUID() As String
            Return Me.DB_CONF_UID
        End Function

		
        Private Sub SetDbConfUID(ByVal i As String)
            If IsNothing(i) Then
                WriteLOG(LogLvl.Warn, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            End If
            Me.DB_CONF_UID = i
        End Sub

		
        Private Function GetDbConfPWD() As String
            Return Me.DB_CONF_PWD
        End Function

		
        Private Sub SetDbConfPWD(ByVal i As String)
            If IsNothing(i) Then
                WriteLOG(LogLvl.Warn, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            End If
            Me.DB_CONF_PWD = i
        End Sub

		
        Private Function GetDBConfTO() As String
            Return Me.DB_CONF_TO
        End Function

		
        Private Sub SetDBConfTO(ByVal i As String)
            If IsNothing(i) Then
                WriteLOG(LogLvl.Warn, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            End If
            Me.DB_CONF_TO = i
        End Sub

		
        Private Function GetDbQryTABLE() As String
            Return Me.DB_QURY_TABLE
        End Function

		
        Private Sub SetDbQryTABLE(ByVal i As String)
            If IsNothing(i) Then
                WriteLOG(LogLvl.Warn, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            End If
            Me.DB_QURY_TABLE = i
        End Sub

		
        Private Function GetDbQryFIELD() As String
            Return Me.DB_QURY_FIELD
        End Function

		
        Private Sub SetDbQryFIELD(ByVal i As String)
            If IsNothing(i) Then
                WriteLOG(LogLvl.Warn, "! " & "NULL input", New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            End If
            Me.DB_QURY_FIELD = i
        End Sub

		
        Public Function GetDbconnStr() As String
            Return Me.DB_CONNSTR
        End Function

		
        Public Sub SetDbconnStr()
            Dim ip As String = GetDbConfIP()
            Dim port As String = GetDbConfPORT()
            Dim dbName As String = GetDbConfDBNAME()
            Dim uid As String = GetDbConfUID()
            Dim pwd As String = GetDbConfPWD()
            Dim timeout As String = GetDBConfTO()

            If (IsHostIpValid(ip) = False) Or (IsPortValid(port) = False) Or (IsDbNameValid(dbName) = False) Or IsNothing(uid) Or IsNothing(pwd) Or IsNothing(timeout) Then
                WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
            End If
            Me.DB_CONNSTR = "Server=" & ip & ":" & port & ";" & "Database=" & dbName & ";" & "UID=" & uid & ";" & "PWD=" & pwd & ";" & "Connect Timeout=" & timeout
        End Sub

		
        Public Sub SetDbQryDb2SelAwrdno()
            Dim fld As String = GetDbQryFIELD()
            Dim tbl As String = GetDbQryTABLE()

            If IsNothing(fld) Or IsNothing(tbl) Then
                WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
            End If
            Me.DB_QUERY_DB2SEL_AWARDNO = "select " & fld & " from " & tbl & ";"
        End Sub

		
        Public Function GetDbQryDb2SelAwrdno() As String
            Return Me.DB_QUERY_DB2SEL_AWARDNO
        End Function

		
        Private Function IsDbNameValid(ByVal db As String) As Boolean
            Dim bIsValid As Boolean

            If IsNothing(db) Then
                WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
                Return False
            End If
            bIsValid = db.StartsWith("UGD", True, Nothing)
            Return bIsValid
        End Function

		
        Public Function DoUpdDBSrvForAwardNODelete(ByVal constr As String) As Boolean
            Const DBTIMEOUT As Integer = 30 ' sec
            Dim sQryOptr As String = Nothing
            Dim conn As DB2Connection = Nothing
            Dim dr As DB2DataReader = Nothing
            Dim cmd As DB2Command = Nothing

            sQryOptr = "delte from " & GetDbQryTABLE() & " where " & "(" & "FLAG='Y'" & " and " & GetDbQryFIELD() & "="

            WriteLOG(LogLvl.Dbg, "! " & "gROI.Length: " & Me.gROI.Length, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            For Each tt As String In Me.gROI
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

		
        Public Function DoUpdDBSrvForAwardNO(ByVal constr As String) As Boolean
            Const DBTIMEOUT As Integer = 30 ' sec
            Dim sQryOptr As String = Nothing
            Dim conn As DB2Connection = Nothing
            Dim dr As DB2DataReader = Nothing
            Dim cmd As DB2Command = Nothing

            sQryOptr = "UPDATE " & GetDbQryTABLE() & " Set " & "FLAG='Y' " & "Where " & GetDbQryFIELD() & "="

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

        Public Sub SetDbConfig()
            Dim fs As FileStream = Nothing
            Dim rs As StreamReader = Nothing
            Dim dat As String = Nothing
            Dim str As Array = Nothing
            Dim lines As String()
            Dim linetot As Integer = GetFileLine(loginConfFn)

            Try
                fs = New FileStream(loginConfFn, FileMode.Open, FileAccess.Read)
                rs = New StreamReader(fs)

                dat = rs.ReadToEnd()
                lines = System.Text.RegularExpressions.Regex.Split(dat, Environment.NewLine)

                For i As Integer = 0 To (linetot - 1)
                    str = Split(lines(i), "=")
                    If str(0) = "DB_CONF_IP" Then SetDbConfIP(str(1))
                    If str(0) = "DB_CONF_PORT" Then SetDbConfPORT(str(1))
                    If str(0) = "DB_CONF_DBNAME" Then SetDbConfDBNAME(str(1))
                    If str(0) = "DB_CONF_UID" Then SetDbConfUID(str(1))
                    If str(0) = "DB_CONF_PWD" Then SetDbConfPWD(str(1))
                    If str(0) = "DB_CONF_TO" Then SetDBConfTO(str(1))
                    If str(0) = "DB_QURY_TABLE" Then SetDbQryTABLE(str(1))
                    If str(0) = "DB_QURY_FIELD" Then SetDbQryFIELD(str(1))
                Next
            Catch ex As Exception
                WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
            Finally
                If rs IsNot Nothing Then rs.Close()
                rs = Nothing
                If fs IsNot Nothing Then fs.Close()
                fs = Nothing
            End Try
        End Sub


        'Public Function doDbConn() As Boolean
        '    Return True
        'End Function

		
        Public Function DoDbSrvQuery(ByVal constr As String, ByVal cmd As String) As Boolean
            Dim conn As DB2Connection = Nothing
            Dim command As DB2Command = Nothing
            Dim dr As DB2DataReader = Nothing
            Const DBTIMEOUT As Integer = 30 'in sec

            If (cmd Is Nothing) Or (constr Is Nothing) Then
                WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
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
                ''aAwardnoArrFrmDB.Distinct.ToArray() 
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

    End Class


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

	
    Private Function GetFileLine(ByVal strFilePath As String) As Integer
        Dim strSplit() As String
        Dim sr As StreamReader = New StreamReader(strFilePath)
        Dim fileCnt As String = sr.ReadToEnd()
		
        sr.Close()
        sr = Nothing
        strSplit = Split(fileCnt, vbCrLf)
        Return strSplit.GetUpperBound(0) + 1
    End Function

	
    Private Function IsHostIpValid(ByVal inIP As String) As Boolean
        Dim ipaddr As IPAddress = Nothing

        If IsNothing(inIP) Then
            WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
            Return False
        End If
        Dim is_valid As Boolean = IPAddress.TryParse(inIP, ipaddr)
        Return is_valid
    End Function

	
    Private Function IsPortValid(ByVal port As String) As Boolean
        If IsNothing(port) Then
            WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
            Return False
        End If
        Dim portNum As Integer = CInt(port)
        If (portNum < 0) Or (portNum > 65536) Then
            Return False
        End If
        Return True
    End Function

End Module