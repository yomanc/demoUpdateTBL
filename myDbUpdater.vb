Imports System.IO
Imports System.Net
Imports System.Data
Imports IBM.Data.DB2
Imports System.Reflection
Imports demoUpdateEFTBL03.Logger
Imports demoUpdateEFTBL03.Module1


Public Class myDbUpdater

    Dim loginConfFn As String = "config.txt"
    Dim loggerDb = New Logger

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

    Dim gROI() As String = Nothing


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

	
    Public Function GetDbconnStr() As String
        Return Me.DB_CONNSTR
    End Function


    Public Sub SetDbconnStrAsNull()
        loggerDb.WriteLOG(LogLvl.Warn, "- ", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        Me.DB_CONNSTR = Nothing
    End Sub


    Public Sub SetDbconnStr()
        Dim ip As String = GetDbConfIP()
        Dim port As String = GetDbConfPORT()
        Dim dbName As String = GetDbConfDBNAME()
        Dim uid As String = GetDbConfUID()
        Dim pwd As String = GetDbConfPWD()
        Dim timeout As String = GetDBConfTO()

        If (IsHostIpValid(ip) = False) Or (IsPortValid(port) = False) Or (IsDbNameValid(dbName) = False) Or IsNothing(uid) Or IsNothing(pwd) Or IsNothing(timeout) Then
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
        End If
        Me.DB_CONNSTR = "Server=" & ip & ":" & port & ";" & "Database=" & dbName & ";" & "UID=" & uid & ";" & "PWD=" & pwd & ";" & "Connect Timeout=" & timeout
    End Sub


    Public Sub SetDbQryDb2SelAwrdnoAsNull()
        loggerDb.WriteLOG(LogLvl.Warn, "- ", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        Me.DB_QUERY_DB2SEL_AWARDNO = Nothing
    End Sub


    Public Sub SetDbQryDb2SelAwrdno()
        Dim fld As String = GetDbQryFIELD()
        Dim tbl As String = GetDbQryTABLE()

        If IsNothing(fld) Or IsNothing(tbl) Then
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
        End If
        Me.DB_QUERY_DB2SEL_AWARDNO = "select " & fld & " from " & tbl & ";"
    End Sub

	
    Public Function GetDbQryDb2SelAwrdno() As String
        Return Me.DB_QUERY_DB2SEL_AWARDNO
    End Function


    Public Sub ResetDbObj()
        Me.SetDbConfigAsNull()
        Me.SetDbconnStrAsNull()
        Me.SetDbQryDb2SelAwrdnoAsNull()
    End Sub

    Public Sub SetDbObj()
        Me.SetDbConfig()
        Me.SetDbconnStr()
    End Sub


    Private Function IsDbNameValid(ByVal db As String) As Boolean
        Dim bIsValid As Boolean

        If IsNothing(db) Then
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
            Return False
        End If

        bIsValid = db.StartsWith("UGD", True, Nothing)
        Return bIsValid
    End Function


    Public Function DoUpdDBSrvForAwardNO(ByVal constr As String) As Boolean
        Dim DBTIMEOUT As Integer = 30 ' sec
        Dim sQryOpr As String = Nothing
        Dim dr As DB2DataReader = Nothing
        Dim cmd As DB2Command = Nothing

        Dim tarTbl = GetDbQryTABLE()
        Dim tarFld = GetDbQryFIELD()
        If IsNothing(constr) Or IsNothing(tarTbl) Or IsNothing(tarFld) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Else
            sQryOpr = "UPDATE " & tarTbl & " Set " & "FLAG='Y' " & "Where " & tarFld & "="
            'sQryOpr = "UPDATE " & GetDbQryTABLE() & " Set " & GetDbQryFIELD() & "='5566' " & "Where " & GetDbQryFIELD() & "=" 'For TEST
        End If

        Dim conn As DB2Connection = New DB2Connection(constr)
        If IsNothing(conn) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_ALLOC_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        End If
        conn.Open()

        loggerDb.WriteLOG(LogLvl.Dbg, "! " & "gROI.Length: " & gROI.Length, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        For Each sQryOpd As String In gROI
            Try
                If conn.IsOpen() = False Then
                    loggerDb.WriteLOG(LogLvl.Err, "! " & "conn.Open NG", New StackTrace(True).GetFrame(0).GetMethod().Name, , New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Return False
                End If

                Dim sCmdtxt = sQryOpr & "'" & sQryOpd & "'" & ";"
                cmd = createDB2Command(conn, sCmdtxt, DBTIMEOUT)
                loggerDb.WriteLOG(LogLvl.Dbg, "* " & cmd.CommandText, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                dr = cmd.ExecuteReader()
                If (dr.HasRows = True) Then
                    While dr.Read()
                        loggerDb.WriteLOG(LogLvl.Err, dr.GetValue(0).ToString, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    End While
                End If

            Catch ex As Exception
                loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            Finally
                releaseDB2Command(cmd)
                releaseDB2DataReader(dr)
            End Try
        Next

        releaseDB2Connection(conn)

        Return True
    End Function


    Public Function DoUpdDBSrvForAwardNODelete(ByVal constr As String) As Boolean
        Const DBTIMEOUT As Integer = 30 ' sec
        Dim sQryOptr As String = Nothing
        Dim conn As DB2Connection = Nothing
        Dim dr As DB2DataReader = Nothing
        Dim cmd As DB2Command = Nothing
        Dim tarTbl = GetDbQryTABLE()
        Dim tarFld = GetDbQryFIELD()

        If IsNothing(constr) Or IsNothing(tarTbl) Or IsNothing(tarFld) Then
            loggerDb.WriteLOG(LogLvl.Err, "! " & NULL_INPUT_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Return False
        Else
            sQryOptr = "delete from " & tarTbl & " where " & "(" & "FLAG='Y'" & " and " & tarFld & "="
        End If


        loggerDb.WriteLOG(LogLvl.Dbg, "! " & "gROI.Length: " & Me.gROI.Length, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

        For Each tt As String In Me.gROI
            Try
                conn = New DB2Connection(constr)
                conn.Open()
                If conn.IsOpen() = False Then
                    loggerDb.WriteLOG(LogLvl.Err, "! " & "conn.Open NG", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    Return False
                End If
                ''loggerDb.WriteLOG(LogLvl.Inf, "! " & "Conn SrvVer: " & conn.ServerVersion & " Database: " & conn.Database & " state: " & conn.State, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                cmd = conn.CreateCommand()
                cmd.CommandText = sQryOptr & "'" & tt & "'" & ")" & ";"
                cmd.CommandTimeout = DBTIMEOUT
                loggerDb.WriteLOG(LogLvl.Dbg, "* " & cmd.CommandText, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

                dr = cmd.ExecuteReader()
                If (dr.HasRows = True) Then
                    While dr.Read()
                        loggerDb.WriteLOG(LogLvl.Err, dr.GetValue(0).ToString, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                    End While
                End If
            Catch exBadImageFormat As BadImageFormatException
                loggerDb.WriteLOG(LogLvl.Err, EXCEPT_BADIMAGEFORMAT & exBadImageFormat.Message.Substring(0, exBadImageFormat.Message.IndexOf(".") + 1), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            Catch ex As Exception
                loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
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


    Public Sub releaseDB2Connection(ByRef conn As DB2Connection)
        If Not conn Is Nothing Then conn.Dispose()
        conn = Nothing
    End Sub


    Public Sub releaseDB2Command(ByRef cmd As DB2Command)
        If Not cmd Is Nothing Then cmd.Dispose()
        cmd = Nothing
    End Sub


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

            Console.WriteLine("[DB_CONF_IP]" & GetDbConfIP())
            Console.WriteLine("[DB_CONF_PORT]" & GetDbConfPORT())
            Console.WriteLine("[DB_CONF_DBNAME]" & GetDbConfDBNAME())
            Console.WriteLine("[DB_CONF_UID]" & GetDbConfUID())
            Console.WriteLine("[DB_CONF_PWD]" & GetDbConfPWD())
            Console.WriteLine("[DB_CONF_TO]" & GetDBConfTO())
            Console.WriteLine("[DB_QURY_TABLE]" & GetDbQryTABLE())
            Console.WriteLine("[DB_QURY_FIELD]" & GetDbQryFIELD())

        Catch ex As Exception
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
        Finally
            If rs IsNot Nothing Then rs.Close()
            rs = Nothing
            If fs IsNot Nothing Then fs.Close()
            fs = Nothing
        End Try
    End Sub


    Public Sub SetDbConfigAsNull()
        loggerDb.WriteLOG(LogLvl.Warn, "- ", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        Try
            SetDbConfIP(Nothing)
            SetDbConfPORT(Nothing)
            SetDbConfDBNAME(Nothing)
            SetDbConfUID(Nothing)
            SetDbConfPWD(Nothing)
            SetDBConfTO(Nothing)
            SetDbQryTABLE(Nothing)
            SetDbQryFIELD(Nothing)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_UNDEF & ex.Message, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
        End Try

    End Sub


    Public Function doDbConn() As Boolean
        Return True
    End Function


    Public Function DoDbSrvQuery() As Boolean
        Dim conn As DB2Connection = Nothing
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

        Try
            conn = New DB2Connection(constr)
            conn.Open()
            If conn.IsOpen() = False Then
                loggerDb.WriteLOG(LogLvl.Err, "! " & "conn.Open NG", New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
                Return False
            End If
            ''loggerDb.WriteLOG(LogLvl.Inf, "! " & "Conn SrvVer: " & conn.ServerVersion & " Database: " & conn.Database & " state: " & conn.State, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())

            command = conn.CreateCommand()
            command.CommandText = cmd
            command.CommandTimeout = DBTIMEOUT
            dr = command.ExecuteReader()

            iAwardnoArrSzFrmDB = 0
            Array.Clear(aAwardnoArrFrmDB, 0, aAwardnoArrFrmDB.Length)

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
            If Not conn Is Nothing Then conn.Dispose()
            If Not dr Is Nothing Then dr.Close()
            conn = Nothing
            dr = Nothing
        End Try
    End Function


    Private Function IsHostIpValid(ByVal inIP As String) As Boolean
        If IsNothing(inIP) Then
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
            Return False
        End If

        Dim ipaddr As IPAddress = Nothing
        Dim is_valid As Boolean = IPAddress.TryParse(inIP, ipaddr)
		
        Return is_valid
    End Function

	
    Private Function IsPortValid(ByVal port As String) As Boolean
        If IsNothing(port) Then
            loggerDb.WriteLOG(LogLvl.Err, EXCEPT_USRDEF & INVALID_CONF_DESC, New StackTrace(True).GetFrame(0).GetMethod().Name, New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
            Throw New Exception(EXCEPT_USRDEF & INVALID_CONF_DESC)
            Return False
        End If

        Dim portNum As Integer = CInt(port)

        If (portNum < 0) Or (portNum > 65536) Then
            Return False
        End If

        Return True
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


    'Public Sub dbgTestConfigAll()
    '    loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfIP(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    '    loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfPORT(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    '    loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfDBNAME(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    '    loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfUID(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    '    loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbConfPWD(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    '    loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDBConfTO(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    '    loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbQryTABLE(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    '    loggerDb.WriteLOG(LogLvl.Dbg, "* " & GetDbQryFIELD(), New StackTrace(True).GetFrame(0).GetFileLineNumber().ToString())
    'End Sub

End Class
