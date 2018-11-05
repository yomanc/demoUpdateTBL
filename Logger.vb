Imports System.IO
Imports System.Reflection

Public Class Logger

    Enum LogLvl
        Err
        Warn
        Inf
        Dbg
    End Enum

    Public gUserPause As Boolean = True
    Public gDBG_ConsoleDispON As Boolean = True
    Public gDBG_level As Integer = LogLvl.Dbg

    Public Const EXCEPT_UNDEF As String = "! Undef Exception: "
    Public Const EXCEPT_AOOR As String = "! ArgumentOutOfRangeException: "
    Public Const EXCEPT_OF As String = "! OverflowException: "
    Public Const EXCEPT_WEB As String = "! Web Exception: "
    Public Const EXCEPT_ARRUMENT As String = "! ArgumentException: "
    Public Const EXCEPT_MALFORMEDLINE As String = "! MalformedLine: "
    Public Const EXCEPT_BADIMAGEFORMAT As String = "! BadImageFormat: "
    Public Const EXCEPT_USRDEF As String = "! UsrDef Exception: "
    Public Const WRONG_FTP As String = "! FTP Error: "

    Public Const UNKNOWN_ISSUE_DESC As String = "fatal. unknown issue"
    Public Const INVALID_PARM_DESC As String = "fatal. invalid param"
    Public Const INVALID_CONF_DESC As String = "fatal. invalid config"
    Public Const NULL_INPUT_DESC As String = "warn. null input"
    Public Const NULL_ALLOC_DESC As String = "fatl. invalid alloc"
    Public Const DIR_ALREADY_EXIST_DESC As String = "Dir should not be existed"
    Public Const NULL_SREMOFN_DESC As String = "sRemoFn NULL, Abort"
    Public Const NULL_SLATESTCSVFN_DESC As String = "sLatestCSVFn NULL"
    Public Const INVALID_GETFILESZ_DESC As String = "getFileSize 0 or NG"
    Public Const INVALID_DOWNLOADFILE_DESC As String = "downloadFile NG"
    Public Const INVALID_PARSECSVFILE_DESC As String = "bDoParseCSVFile NG"
    Public Const MISSION_DESC_S As String = "Misson.s"
    Public Const MISSION_DESC_E_UNEXP As String = "Misson.e.unexp"
    Public Const MISSION_DESC_E_DONE As String = "Misson.e.done"

    Public Const HINT_BANNER As String = "====="
    Public Const HINT_STEP1_S As String = HINT_BANNER & " STEP1 set FTP login conf S " & HINT_BANNER
    Public Const HINT_STEP1_E As String = HINT_BANNER & " STEP1 set FTP login conf E " & HINT_BANNER
    Public Const HINT_STEP2_S As String = HINT_BANNER & " STEP2 select data S " & HINT_BANNER
    Public Const HINT_STEP2_E As String = HINT_BANNER & " STEP2 select data E " & HINT_BANNER
    Public Const HINT_STEP4_S As String = HINT_BANNER & " STEP3 dn data S " & HINT_BANNER
    Public Const HINT_STEP4_E As String = HINT_BANNER & " STEP3 dn data E " & HINT_BANNER
    Public Const HINT_STEP5_S As String = HINT_BANNER & " STEP4 parsing data S " & HINT_BANNER
    Public Const HINT_STEP5_E As String = HINT_BANNER & " STEP4 parsing data E " & HINT_BANNER
    Public Const HINT_STEP6_S As String = HINT_BANNER & " STEP5 login database S " & HINT_BANNER
    Public Const HINT_STEP6_E As String = HINT_BANNER & " STEP5 login database E " & HINT_BANNER
    Public Const HINT_STEP7_S As String = HINT_BANNER & " STEP6 do diff table S " & HINT_BANNER
    Public Const HINT_STEP7_E As String = HINT_BANNER & " STEP6 do diff table E " & HINT_BANNER
    Public Const HINT_STEP8_S As String = HINT_BANNER & " STEP7 sudo upd table S " & HINT_BANNER
    Public Const HINT_STEP8_E As String = HINT_BANNER & " STEP7 sudo upd table E " & HINT_BANNER
    Public Const HINT_STEP9_S As String = HINT_BANNER & " STEP8 do upd table S " & HINT_BANNER
    Public Const HINT_STEP9_E As String = HINT_BANNER & " STEP8 do upd table E " & HINT_BANNER


    Public Sub WriteLOG(ByVal hint As LogLvl, ByVal Msg As String, ByVal curName As String, ByVal exceptLine As String)
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
            'currentName = New StackTrace(True).GetFrame(0).GetMethod().Name
            'callName = New StackTrace(True).GetFrame(1).GetMethod().Name
            Try
                aFile = New FileStream(LOCALDNLOADDIR & LOCALLOGFN, FileMode.OpenOrCreate)
                aFile.Seek(0, SeekOrigin.End)
                sw = New StreamWriter(aFile)
                'sw.WriteLine("{0}{1} {2} {3} @{4}#{5}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString(), sPrefix, Msg.Replace(vbCrLf, ""), callName, exceptLine)
                sw.WriteLine("{0}{1} {2} {3} @{4}#{5}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString(), sPrefix, Msg.Replace(vbCrLf, ""), curName, exceptLine)
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


    Public Function CreateLoggerDir(ByVal p As String) As Boolean
        If Not Directory.Exists(p) Then
            My.Computer.FileSystem.CreateDirectory(p)
            Return True
        Else
            Return False
        End If
    End Function


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

End Class
