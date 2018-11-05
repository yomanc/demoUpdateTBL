Imports System.IO
Imports System.Net
Imports System.Data
Imports IBM.Data.DB2
Imports System.Reflection
Imports demoUpdateEFTBL03.myDbUpdater
Imports demoUpdateEFTBL03.myFtpDownloader
Imports demoUpdateEFTBL03.Logger


Public Module Module1

    Dim LOG_TIME As String = DateTime.Now.ToString("yyyyMMddHHmmss")
    Public LOCALDNLOADDIR As String = "LOG" & LOG_TIME & "/"
    Public LOCALLOGFN As String = LOG_TIME & ".log"

    Dim MyFTPDBconfig As String = "config.txt"
    Dim FTPPREFIX As String = "ftp://"

    Public logger = New Logger

	
    Sub Main(args As String())
        If logger.CreateLoggerDir(LOCALDNLOADDIR) = False Then
            logger.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & DIR_ALREADY_EXIST_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            logger.WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            logger.Abort()
        End If

        logger.WriteLOG(LogLvl.Dbg, "* " & "outputDir :" & System.AppDomain.CurrentDomain.BaseDirectory(), New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        Console.WriteLine("Start execution")
        logger.WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        Try
            Dim ftpdn = New myFtpDownloader
            Dim dbupdter = New myDbUpdater

            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP1_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.SetConf(MyFTPDBconfig) = False Then Return
            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP1_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP2_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.SetsRemoFn() = False Then Return
            'logger.WriteLOG(LogLvl.Dbg, "* " & "sRemoFn :" & ftpdn.GetsRemoFn, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.GetsRemoFn() Is Nothing Then Return

            Dim REMFNPREFIX As String = ftpdn.GetFtpConfigREMFNPREFIX
            Dim aBufCvsToday() As String = ftpdn.GetaBufCvsToday
            Dim fmt As String = "yyyyMMddHHmm"
            If ftpdn.SetsLatestCSVFn(ftpdn.RetLatestCsvFn(aBufCvsToday, REMFNPREFIX, fmt)) = False Then Return
            If ftpdn.GetsLatestCSVFn Is Nothing Then Return
            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP2_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP3_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Dim szExp As Integer = CInt(ftpdn.getFileSize(ftpdn.GetUid(), ftpdn.GetPWD(), "ftp://" & ftpdn.GetHost(), ftpdn.GetFtpConfREMDIR & ftpdn.GetsRemoFn))
            Dim szDn As Integer = ftpdn.downloadFile(FTPPREFIX & ftpdn.GetHost() & ftpdn.GetFtpConfREMDIR & ftpdn.GetsRemoFn, ftpdn.GetUid(), ftpdn.GetPWD(), LOCALDNLOADDIR & ftpdn.GetsLatestCSVFn)
            If (szDn <> szExp) Or (szDn <= 0) Then Return
            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP3_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP4_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.DoParseCsvFile(LOCALDNLOADDIR & ftpdn.GetsLatestCSVFn) = False Then Return
            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP4_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP5_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            ''dbupdter.ResetDbObj()
            dbupdter.SetDbObj()
            dbupdter.SetDbQryDb2SelAwrdno()
            If dbupdter.DoDbSrvQuery() = False Then Return
            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP5_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP6_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Dim ROI_amount = dbupdter.DoSetDGroupForAwardNO(dbupdter.GetaAwardnoArrFrmDB(), ftpdn.GetaAwardnoArrFrmCVS())
            If ROI_amount < 0 Then
                Return
            ElseIf ROI_amount = 0 Then
                logger.WriteLOG(LogLvl.Err, "! " & "Donot need update", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                logger.WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_DONE, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Console.WriteLine("End execution")
                Return
            End If
            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP6_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP7_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If dbupdter.DoUpdDBSrvForAwardNO(dbupdter.GetDbconnStr()) = False Then Return
            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP7_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP8_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If dbupdter.DoUpdDBSrvForAwardNODelete(dbupdter.GetDbconnStr()) = False Then Return
            logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP8_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            ''verify?

            logger.WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_DONE, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Console.WriteLine("End execution")

            If logger.gDBG_ConsoleDispON = True Then
                logger.Pause()
            End If
        Catch ex As Exception
            logger.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            logger.WriteLOG(LogLvl.Err, "- " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            logger.Abort()
        End Try
    End Sub

End Module