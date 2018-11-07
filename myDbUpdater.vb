Imports System.IO
Imports System.Net
Imports System.Data
Imports IBM.Data.DB2
Imports System.Reflection
Imports demoUpdateEFTBL03.Logger
Imports demoUpdateEFTBL03.Module1
Imports System.Text ' For StringBuilder

Public Class myDbUpdater

    Dim DB_CONF_IP As String = Nothing
    Dim DB_CONF_PORT As String = Nothing
    Dim DB_CONF_DBNAME As String = Nothing
    Dim DB_CONF_UID As String = Nothing
    Dim DB_CONF_PWD As String = Nothing
    Dim DB_CONF_TO As String = Nothing 'second
    Dim DB_QURY_TABLE As String = Nothing
    Dim DB_QURY_FIELD As String = Nothing
    Dim DB_QUERY_DB2SEL_AWARDNO As String = Nothing
    Dim DB_CONNSTR As String = Nothing
    Dim aAwardnoArrFrmDB(900000) As String
    Dim iAwardnoArrSzFrmDB As Integer = 0
    Dim DB2conn As DB2Connection = Nothing
    Dim gROI() As String = Nothing
    Dim loggerDb = New Logger(LOCALDNLOADDIR, LOCALLOGFN)

    ' Constructor
    Public Sub New(Path As String)
        Dim fs As FileStream = Nothing
        Dim sr As StreamReader = Nothing
        Dim dat As String = Nothing
        Dim str As Array = Nothing
        Dim lines As String()
        Dim linetot As Integer = GetFileLine(Path)

        If IsNothing(Path) Then Return

        Try
            fs = New FileStream(Path, FileMode.Open, FileAccess.Read)
            sr = New StreamReader(fs)

            dat = sr.ReadToEnd()
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
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
        Finally
            ReleaseStreamReader(sr)
            ReleaseFileStream(fs)
        End Try
    End Sub


    Public Function DoSetDGroupForAwardNO(ByVal listA As String(), ByVal listB As String()) As Integer
        Try
            Dim aDiff = listA.Except(listB).ToArray()
            ReDim Preserve gROI(aDiff.Length - 1)
            gROI = aDiff

            loggerDb.WriteLOG(LogLvl.Inf, "* " & "gROI total " & gROI.Length, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            loggerDb.WriteLOG(LogLvl.Dbg, "- " & "----------gROI.s ", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            For Each tt As String In gROI
                loggerDb.WriteLOG(LogLvl.Dbg, "- " & tt, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Next
            loggerDb.WriteLOG(LogLvl.Dbg, "- " & "----------gROI.e ", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            Return aDiff.Length
        Catch ex As Exception
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_UNDEF & UNKNOWN_ISSUE_DESC)
            Return -1
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
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return
        End If
        Me.DB_CONF_IP = i
    End Sub


    Private Function GetDbConfPORT() As String
        Return Me.DB_CONF_PORT
    End Function


    Private Sub SetDbConfPORT(ByVal i As String)
        If IsNothing(i) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return
        End If
        Me.DB_CONF_PORT = i
    End Sub


    Private Function GetDbConfDBNAME() As String
        Return Me.DB_CONF_DBNAME
    End Function


    Private Sub SetDbConfDBNAME(ByVal i As String)
        If IsNothing(i) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return
        End If
        Me.DB_CONF_DBNAME = i
    End Sub


    Private Function GetDbConfUID() As String
        Return Me.DB_CONF_UID
    End Function


    Private Sub SetDbConfUID(ByVal i As String)
        If IsNothing(i) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return
        End If
        Me.DB_CONF_UID = i
    End Sub


    Private Function GetDbConfPWD() As String
        Return Me.DB_CONF_PWD
    End Function


    Private Sub SetDbConfPWD(ByVal i As String)
        If IsNothing(i) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return
        End If
        Me.DB_CONF_PWD = i
    End Sub


    Private Function GetDBConfTO() As String
        Return Me.DB_CONF_TO
    End Function


    Private Sub SetDBConfTO(ByVal i As String)
        If IsNothing(i) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return
        End If
        Me.DB_CONF_TO = i
    End Sub


    Private Function GetDbQryTABLE() As String
        Return Me.DB_QURY_TABLE
    End Function


    Private Sub SetDbQryTABLE(ByVal i As String)
        If IsNothing(i) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return
        End If
        Me.DB_QURY_TABLE = i
    End Sub


    Private Function GetDbQryFIELD() As String
        Return Me.DB_QURY_FIELD
    End Function


    Private Sub SetDbQryFIELD(ByVal i As String)
        If IsNothing(i) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & "NULL input", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return
        End If
        Me.DB_QURY_FIELD = i
    End Sub

    Public Sub SetDbconnStr(ByVal i As String)
        If IsNothing(i) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & "NULL input", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return
        End If
        Me.DB_CONNSTR = i
    End Sub


    Public Function GetDbconnStr() As String
        Return Me.DB_CONNSTR
    End Function


    Public Sub SetDbQryDb2SelAwrdno()
        Dim fld As String = GetDbQryFIELD()
        Dim tbl As String = GetDbQryTABLE()

        If IsNothing(fld) Or IsNothing(tbl) Then
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
        End If
        Me.DB_QUERY_DB2SEL_AWARDNO = "select " & fld & " from " & tbl & ";"

        loggerDb.WriteLOG(LogLvl.Dbg, Me.DB_QUERY_DB2SEL_AWARDNO, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    End Sub


    Public Function GetDbQryDb2SelAwrdno() As String
        Return Me.DB_QUERY_DB2SEL_AWARDNO
    End Function


    Public Function DoUpdDBSrvForAwardNO(ByVal constr As String) As Boolean
        Dim DBTIMEOUT As Integer = 30 ' sec
        Dim sQryOpr As StringBuilder = Nothing
        Dim dr As DB2DataReader = Nothing
        Dim cmd As DB2Command = Nothing

        Dim SudoKey As String = "000000000"

        Dim tarTbl = GetDbQryTABLE()
        Dim tarFld = GetDbQryFIELD()
        If IsNothing(constr) Or IsNothing(tarTbl) Or IsNothing(tarFld) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End If

        sQryOpr = New StringBuilder("UPDATE " & tarTbl & " Set " & "FLAG='Y' " & "Where " & tarFld & "=" & "'" & SudoKey & "'" & ";")

        loggerDb.WriteLOG(LogLvl.Dbg, "! " & "gROI.Length: " & gROI.Length, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        If DB2conn.IsOpen = False Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & "DB2conn disconnect", New StackTrace(True).GetFrame(0).GetMethod().Name, , New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End If

        For Each sQryOpd As String In gROI
            Try
                ' Preload a default format with value, then replace what we want
                sQryOpr.Replace(SudoKey, sQryOpd)
                cmd = createDB2Command(DB2conn, sQryOpr.ToString, DBTIMEOUT)
                loggerDb.WriteLOG(LogLvl.Dbg, "* " & cmd.CommandText, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            Finally
                sQryOpr.Replace(sQryOpd, SudoKey)
                releaseDB2DataReader(dr)
                releaseDB2Command(cmd)
            End Try
        Next

        Return True
    End Function


    Public Function DoUpdDBSrvForAwardNODelete(ByVal constr As String) As Boolean
        Const DBTIMEOUT As Integer = 30 ' sec
        Dim sQryOptr As StringBuilder = Nothing ' Refactor. Replace String object with StringBuilder object forward to better performance
        Dim dr As DB2DataReader = Nothing
        Dim cmd As DB2Command = Nothing
        Dim tarTbl = GetDbQryTABLE()
        Dim tarFld = GetDbQryFIELD()
        Dim SudoKey = New String("000000000")

        If IsNothing(constr) Or IsNothing(tarTbl) Or IsNothing(tarFld) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End If

        sQryOptr = New StringBuilder("delete from " & tarTbl & " where " & "(" & "FLAG='Y'" & " and " & tarFld & "=" & "'" & SudoKey & "'" & ")" & ";")

        loggerDb.WriteLOG(LogLvl.Dbg, "! " & "gROI.Length: " & Me.gROI.Length, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        If DB2conn.IsOpen = False Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & "DB2conn disconnect", New StackTrace(True).GetFrame(0).GetMethod().Name, , New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End If
        ''loggerDb.WriteLOG(LogLvl.Inf, "! " & "Conn SrvVer: " & conn.ServerVersion & " Database: " & conn.Database & " state: " & conn.State, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        For Each tt As String In Me.gROI
            Try
                ' Preload a default format with value, then replace what we want
                sQryOptr.Replace(SudoKey, tt)
                cmd = createDB2Command(DB2conn, sQryOptr.ToString, DBTIMEOUT)
                loggerDb.WriteLOG(LogLvl.Dbg, "* " & cmd.CommandText, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                cmd.ExecuteNonQuery()
            Catch exBadImageFormat As BadImageFormatException
                loggerDb.WriteLOG(LogLvl.Err, EXCEPT_BADIMAGEFORMAT & exBadImageFormat.Message.Substring(0, exBadImageFormat.Message.IndexOf(".") + 1), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            Catch ex As Exception
                loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            Finally
                sQryOptr.Replace(tt, SudoKey)
                releaseDB2DataReader(dr)
                releaseDB2Command(cmd)
            End Try
        Next
        Return True
    End Function


    Public Sub releaseDB2DataReader(ByRef dr As DB2DataReader)
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Sub


    Public Function createDB2Command(ByRef conn As DB2Connection, ByVal contxt As String, ByVal timeout As Integer) As DB2Command
        Dim cmd As DB2Command = Nothing

        cmd = conn.CreateCommand()
        cmd.CommandText = contxt
        cmd.CommandTimeout = timeout
        Return cmd
    End Function


    Public Sub releaseDB2Command(ByRef cmd As DB2Command)
        If Not cmd Is Nothing Then cmd.Dispose()
        cmd = Nothing
    End Sub


    Public Sub ReleaseFileStream(ByRef fs As FileStream)
        If Not fs Is Nothing Then fs.Close()
        fs = Nothing
    End Sub


    Public Sub ReleaseStreamReader(ByRef sr As StreamReader)
        If Not sr Is Nothing Then sr.Close()
        sr = Nothing
    End Sub


    Public Function ConnectDb() As Boolean
        Dim server As String = GetDbConfIP()
        Dim userId As String = GetDbConfUID()
        Dim password As String = GetDbConfPWD()
        Dim portNumber As String = GetDbConfPORT()
        Dim db As String = GetDbConfDBNAME()

        SetDbconnStr("Server=" & server & ":" & portNumber & ";Database=" & db & ";UID=" & userId & ";PWD=" & password)

        Try
            Me.DB2conn = New DB2Connection(GetDbconnStr)
            Me.DB2conn.Open()
            loggerDb.WriteLOG(LogLvl.Warn, "* " & "  Connected to the " + db + " database", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return True
        Catch ex As Exception
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            loggerDb.WriteLOG(LogLvl.Err, "- " & MISSION_DESC_E_UNEXP, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End Try
    End Function


    Public Sub DisconnectDb()
        If IsNothing(Me.DB2conn) Then Return

        If Not Me.DB2conn Is Nothing Then Me.DB2conn.Dispose()
        Me.DB2conn = Nothing
    End Sub


    Public Function IsConnected() As Boolean
        If IsNothing(Me.DB2conn) Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function DoDbSrvQuery() As Boolean
        'Dim conn As DB2Connection = Nothing
        Dim command As DB2Command = Nothing
        Dim dr As DB2DataReader = Nothing
        Const DBTIMEOUT As Integer = 30 'in sec

        Dim constr = Me.GetDbconnStr
        Dim cmd = Me.GetDbQryDb2SelAwrdno

        If (cmd Is Nothing) Or (constr Is Nothing) Then
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_PARM_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_PARM_DESC)
            Return False
        End If
        loggerDb.WriteLOG(LogLvl.Dbg, "* " & "constr: " & constr, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        If DB2conn.IsOpen = False Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & "conn.Open NG", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End If
        ''loggerDb.WriteLOG(LogLvl.Inf, "! " & "Conn SrvVer: " & conn.ServerVersion & " Database: " & conn.Database & " state: " & conn.State, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        Try
            command = createDB2Command(DB2conn, cmd, DBTIMEOUT)

            iAwardnoArrSzFrmDB = 0
            Array.Clear(aAwardnoArrFrmDB, 0, aAwardnoArrFrmDB.Length)

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

            loggerDb.WriteLOG(LogLvl.Inf, "* " & "aAwardnoArrFrmDB total " & aAwardnoArrFrmDB.Length, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            ''aAwardnoArrFrmDB.Distinct.ToArray() 
            loggerDb.WriteLOG(LogLvl.Dbg, "- " & "----------aAwardnoArrFrmDB.s", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            For Each tt As String In aAwardnoArrFrmDB
                ''loggerDb.WriteLOG(LogLvl.Dbg, "- " & tt, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Next
            loggerDb.WriteLOG(LogLvl.Dbg, "- " & "----------aAwardnoArrFrmDB.e", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return True
        Catch ex As BadImageFormatException
            Console.WriteLine(ex.Message)
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_BADIMAGEFORMAT & ex.Message.Substring(0, ex.Message.IndexOf(".") + 1), New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine(ex.Message)
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_BADIMAGEFORMAT & ex.Message.Substring(0, ex.Message.IndexOf(".") + 1), New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Finally
            releaseDB2DataReader(dr)
            releaseDB2Command(command)
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


    '[depricated]
    Public Sub dbgTestConfigAll()
        loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfIP(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfPORT(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfDBNAME(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfUID(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfPWD(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDBConfTO(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbQryTABLE(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbQryFIELD(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    End Sub

End Class
