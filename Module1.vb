Imports System.IO
Imports System.Net
Imports System.Data
Imports IBM.Data.DB2
Imports System.Reflection
Imports demoUpdateEFTBL03.myDbUpdater
Imports demoUpdateEFTBL03.myFtpDownloader
Imports demoUpdateEFTBL03.Logger


Public Module Module1

    ' Share this
    Dim LOG_TIME As String = DateTime.Now.ToString("yyyyMMddHHmmss")
    Public LOCALDNLOADDIR As String = "LOG" & LOG_TIME & "/"
    Public LOCALLOGFN As String = LOG_TIME & ".log"


    Sub Main(args As String())

        Dim logger = New Logger(LOCALDNLOADDIR, LOCALLOGFN)

        logger.WriteLOG(LogLvl.Dbg, "* " & "outputDir :" & System.AppDomain.CurrentDomain.BaseDirectory(), New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        Console.WriteLine("Start execution")
        Logger.WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        Try
            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP1_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Dim ftpdn = New myFtpDownloader("config.txt")
            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP1_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP2_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.SetsRemoFn() = False Then Return
            'logger.WriteLOG(LogLvl.Dbg, "* " & "sRemoFn :" & ftpdn.GetsRemoFn, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.GetsRemoFn() Is Nothing Then Return

            Dim REMFNPREFIX As String = ftpdn.GetFtpConfigREMFNPREFIX
            Dim aBufCvsToday() As String = ftpdn.GetaBufCvsToday
            Dim fmt As String = "yyyyMMddHHmm"
            If ftpdn.SetsLatestCSVFn(ftpdn.RetLatestCsvFn(aBufCvsToday, REMFNPREFIX, fmt)) = False Then Return
            If ftpdn.GetsLatestCSVFn Is Nothing Then Return
            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP2_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP3_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Dim szExp As Integer = CInt(ftpdn.getFileSize(ftpdn.GetUid(), ftpdn.GetPWD(), "ftp://" & ftpdn.GetHost(), ftpdn.GetFtpConfREMDIR & ftpdn.GetsRemoFn))
            Dim szDn As Integer = ftpdn.downloadFile(ftpdn.GetHost() & ftpdn.GetFtpConfREMDIR & ftpdn.GetsRemoFn, ftpdn.GetUid(), ftpdn.GetPWD(), LOCALDNLOADDIR & ftpdn.GetsLatestCSVFn)
            If (szDn <> szExp) Or (szDn <= 0) Then Return
            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP3_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP4_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If ftpdn.DoParseCsvFile(LOCALDNLOADDIR & ftpdn.GetsLatestCSVFn) = False Then Return
            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP4_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP5_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            ''dbupdter.ResetDbObj()

            Dim dbupdter = New myDbUpdater("config.txt")

            dbupdter.ConnectDb()

            dbupdter.SetDbQryDb2SelAwrdno()

            If dbupdter.DoDbSrvQuery() = False Then Return
            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP5_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP6_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Dim ROI_amount = dbupdter.DoSetDGroupForAwardNO(dbupdter.GetaAwardnoArrFrmDB(), ftpdn.GetaAwardnoArrFrmCVS())
            If ROI_amount < 0 Then
                Return
            ElseIf ROI_amount = 0 Then
                Logger.WriteLOG(LogLvl.Err, "! " & "Donot need update", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Logger.WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_DONE, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Console.WriteLine("End execution")
                Return
            End If
            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP6_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP7_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If dbupdter.DoUpdDBSrvForAwardNO(dbupdter.GetDbconnStr()) = False Then Return
            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP7_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP8_S, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            If dbupdter.DoUpdDBSrvForAwardNODelete(dbupdter.GetDbconnStr()) = False Then Return
            Logger.WriteLOG(LogLvl.Inf, "- " & HINT_STEP8_E, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            ''verify?

            Logger.WriteLOG(LogLvl.Err, "+ " & MISSION_DESC_E_DONE, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())


            If dbupdter.IsConnected = True Then dbupdter.DisconnectDb()

            Console.WriteLine("End execution")

            If Logger.gDBG_ConsoleDispON = True Then
                Logger.Pause()
            End If
        Catch ex As Exception
            Logger.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Logger.WriteLOG(LogLvl.Err, "- " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Logger.Abort()
        End Try
    End Sub

End Module