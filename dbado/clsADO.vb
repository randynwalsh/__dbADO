Option Strict Off
Option Explicit On

Imports System
Imports System.Text
Imports System.IO
Imports System.Data

Imports System.Diagnostics
Imports System.Collections
Imports System.Web

Imports System.Web.Configuration
Imports System.Web.Security
Imports System.Configuration

Imports VB6 = Microsoft.VisualBasic
Imports System.Net
Imports System.Collections.Generic

' Imports System.Data.SQLite
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Data.Common

<ComVisible(True)>
Public Class clsHTAConnection
    Public myDBConnection As clsDBConnection
    Dim cachedfHandles As Collection

    ' windowDotLocation = C:\Users\rwalsh\Dropbox\_src\__HTA\tcl.hta
    ' startAccountName = "hta"
    Public Sub adoInit(ByVal windowDotLocation As String, ByVal startAccountName As String)
        Dim Errors As String = ""

        allowAllDosWritesChecked = True
        allowAllDosWritesValue = True
        allowAllDosReadsChecked = True
        allowAllDosReadsValue = True

        FullUrl = Field(Field(windowDotLocation, "#", 1), "?", 1)
        cachedfHandles = New Collection

        If startAccountName = "" Then mAccountName = DropExtension(GetFName(FullUrl)) Else mAccountName = startAccountName

        mMapPath = VB6.Replace(FullUrl, "file:///", "")
        mMapPath = VB6.Replace(mMapPath, "/", "\")

        If Not System.IO.Directory.Exists(mMapPath) Then
            If System.IO.File.Exists(mMapPath) = False Then Throw New Exception("File does not exist: " & mMapPath)
            mMapPath = GetFolder(mMapPath) ' drop index.hta
        End If

        If System.IO.Directory.Exists(MapRootPath("App_Data\_database\")) = False Then Throw New Exception("Directory does not exist: " & MapRootAccount(startAccountName))
        If System.IO.Directory.Exists(MapRootAccount(startAccountName)) = False Then Throw New Exception("Account '" & startAccountName & "' does not exist in " & MapRootPath("App_Data\_database\"))

        myDBConnection = New clsDBConnection(UrlAccount())
    End Sub

    Public Function AccountName() As String
        Return mAccountName
    End Function

    Public Function adoAttach(ByVal AccountName As String, ByVal ConnectionItem As String, ByVal UserID As String, ByVal Password As String, ByVal CreateIt As Boolean, ByVal TimeOutSecs As Integer)
        Return myDBConnection.adoAttach(AccountName, ConnectionItem, UserID, Password, CreateIt, TimeOutSecs)
    End Function

    Public Function startExe(ByVal commandline As String)
        Try
            Return Shell(commandline, AppWinStyle.NormalFocus)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
        End Try
    End Function

    Public Sub startDoc(ByVal doc As String, ByVal Args As String)
        Process.Start(doc, Args)
    End Sub

    ' Returns fHandle or **Error
    Public Function adoOpen(ByVal DictData As String, ByVal TableName As String) As String
        Dim ToFile As clsTableHandle = Nothing, Errors As String = Nothing, FromAccount As clsjdbAccount
        If myDBConnection.mAttachedDB Is Nothing Then FromAccount = myDBConnection.mUrlAccount Else FromAccount = myDBConnection.mAttachedDB

        If LCase(DictData) <> "dict" Then DictData = ""
        Try
            If Not myDBConnection.dbOpen(DictData, TableName, ToFile, Errors, myDBConnection.mAttachedDB) Then
                If Errors = "" Then Return Nothing Else Return "**" & Errors
            End If
            Dim fHandle As String = VB6.LCase(ToFile.MyDBAccount.AcountName & "*" & DictData & TableName)
            If cachedfHandles.Contains(fHandle) Then cachedfHandles.Remove(fHandle)
            cachedfHandles.Add(ToFile, fHandle)
            Return fHandle
        Catch ex As Exception
            Return "**" & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Function adoDeleteFile(ByVal TableHandle As String) As String
        Try
            If cachedfHandles.Contains(TableHandle) Then
                Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
                myDBConnection.dbDeleteFile(fHandle)
                cachedfHandles.Remove(TableHandle)
            End If
            Return Nothing
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & ex.StackTrace)
        End Try
        Throw New Exception("File not found")
    End Function

    Public Function adoListFiles() As String
        Try
            Return Join(myDBConnection.dbListFiles().ToArray, Chr(254))
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoTimeDate() As String
        Return myDBConnection.dbTimeDate()
    End Function

    Public Sub adoClearCachedTables()
        Try
            ' Force Open to re-open
            For Each Account As clsjdbAccount In myDBConnection.mCachedDBAccounts
                Account.Tables.Clear()
            Next
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Public Function adoCreateFile(ByVal DictData As String, ByVal TableName As String) As String
        Dim ToFile As clsTableHandle = Nothing, Errors As String = Nothing
        If VB6.LCase(DictData) <> "dict" Then DictData = ""

        Try
            If Not myDBConnection.OpenBool(DictData, TableName, ToFile, Errors, True, Nothing) Then
                If Errors = "" Then Return Nothing Else Return "**" & Errors
            End If
            Dim fHandle As String = VB6.LCase(ToFile.MyDBAccount.AcountName & "*" & DictData & TableName)
            If cachedfHandles.Contains(fHandle) Then cachedfHandles.Remove(fHandle)
            cachedfHandles.Add(ToFile, fHandle)
            Return fHandle
        Catch ex As Exception
            Return "**" & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Function adoPrimaryKeyColumnName(ByVal TableHandle As String) As String
        Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
        Return fHandle.PrimaryKeyColumnName
    End Function


    Function adoReadXML(ByVal TableHandle As String, ByVal ItemID As String, ByVal IsReadU As Boolean) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            Dim XmlDataRecord As Object = Nothing
            If Not fHandle.dbReadXML(ItemID, XmlDataRecord, IsReadU) Then Return Nothing
            Return XML_Obj2Str(XmlDataRecord)
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & "table Handle " & TableHandle & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoReadJSON(ByVal TableHandle As String, ByVal ItemID As String, ByVal IsReadU As Boolean) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            Dim JSonDataRecord As Object = Nothing
            If Not fHandle.dbReadJSon(ItemID, JSonDataRecord, IsReadU) Then Return Nothing
            Return JSON_Obj2Str(JSonDataRecord)
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoReadV(ByVal TableHandle As String, ByVal ItemID As String, ByVal FieldNo As Short, ByVal IsReadU As Boolean) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            Dim Result As Object = ""
            ' ByRef DataRecord As Object, ByVal TableHandle As clsTableHandle, ByVal ItemID As String, ByVal FieldNo As Short
            If Not myDBConnection.dbReadV(Result, fHandle, ItemID, FieldNo, IsReadU) Then Return Nothing
            If Result Is Nothing Then Return ""
            Return CStr(Result)
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & "table Handle " & TableHandle & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoRead(ByVal TableHandle As String, ByVal ItemID As String, ByVal IsReadU As Boolean) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            Dim Result As String = ""
            If Not fHandle.dbReadBool(ItemID, Result, IsReadU) Then Return Nothing
            If Result Is Nothing Then Return ""
            Return STX(Result)
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & "table Handle " & TableHandle & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoClearFile(ByVal TableHandle As String) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            fHandle.ClearFile()
            Return Nothing
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & "table Handle " & TableHandle & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoWrite(ByVal Item As String, ByVal TableHandle As String, ByVal ItemID As String, ByVal IsWriteU As Boolean) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            ' ByRef DataRecord As Object, ByVal TableHandle As clsTableHandle, ByVal ItemID As String, ByVal FieldNo As Short
            fHandle.dbWrite(XTS(Item), ItemID, IsWriteU)
            Return Nothing
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & "table Handle " & TableHandle & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoWriteXML(ByVal XmlItem As String, ByVal TableHandle As String, ByVal ItemID As String, ByVal IsWriteU As Boolean) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            fHandle.dbWriteXML(XML_Str2Obj(XmlItem), ItemID, IsWriteU)
            Return Nothing
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & "table Handle " & TableHandle & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoWriteJSON(ByVal JSonItem As String, ByVal TableHandle As String, ByVal ItemID As String, ByVal IsWriteU As Boolean) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            fHandle.dbWriteJSon(JSON(JSonItem), ItemID, IsWriteU)
            Return Nothing
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & "table Handle " & TableHandle & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoWriteJSON(ByVal Line As String, ByVal TableHandle As String, ByVal ItemID As String, ByVal FieldNo As Short, ByVal IsWriteU As Boolean) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            myDBConnection.dbWriteV(Line, fHandle, ItemID, FieldNo, IsWriteU)
            Return Nothing
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    '   dbDelete, dbSelectVX, dbClearFile
    Function adoDelete(ByVal TableHandle As String, ByVal ItemID As String) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            fHandle.dbDelete(ItemID)
            Return Nothing
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & "table Handle " & TableHandle & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoSelectVX(ByVal ColumnList As String, ByVal TableHandle As String, ByVal WhereClause As String) As String
        Try
            Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
            Dim SL As rSelectList = fHandle.SelectFileX(ColumnList, WhereClause)
            If ColumnList = "" Then Return SL.GetListOfPKs Else Return SL.GetJSONString
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Dim fhandleCookies As String = Nothing

    Function adoSetCookie(ByVal name As String, ByVal value As String)
        If fhandleCookies = Nothing Then
            fhandleCookies = Me.adoCreateFile("", "jsbCookies")
        End If
        adoWrite(STX(value), fhandleCookies, name, False)
        Return True
    End Function

    Function adoGetCookie(ByVal name As String, ByVal optionalvalue As String)
        If fhandleCookies = Nothing Then
            fhandleCookies = Me.adoCreateFile("", "jsbCookies")
        End If
        Dim Result As String = XTS(adoRead(fhandleCookies, name, False))
        If Result Is Nothing Then Result = optionalvalue
        Return Result
    End Function

    Function adoDeleteCookie(ByVal name As String)
        If fhandleCookies = Nothing Then
            fhandleCookies = Me.adoCreateFile("", "jsbCookies")
        End If
        adoDelete(fhandleCookies, name)
        Return True
    End Function

    Function adoSqlSelect(ByVal SQLCommand As String) As String
        Try
            Dim SL As rSelectList = myDBConnection.dbSqlSelect(SQLCommand)
            Return SL.GetJSONString
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoSqlScalar(ByVal SQLCommand As String) As String
        Try
            Return C2Str(myDBConnection.dbSqlScalar(SQLCommand))
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function adoGetDrives() As String
        Dim S As New StringBuilder, ap As String = "["
        For Each D As DriveInfo In DriveInfo.GetDrives
            If D.IsReady Then
                S.Append(ap & "{")
                S.Append("VolumeLabel: '" & D.VolumeLabel & "',")
                S.Append("AvailableFreeSpace: " & D.AvailableFreeSpace & ",")
                S.Append("RootDirectory: '" & VB6.Replace(D.RootDirectory.ToString, "\", "\\") & "',")

                S.Append("DriveFormat: '" & D.DriveFormat & "',")
                S.Append("DriveType: '" & D.DriveType & "',")
                '
                S.Append("TotalFreeSpace: " & D.TotalFreeSpace & ",")
                S.Append("TotalSize: " & D.TotalSize & "")

                S.Append("}")
                ap = "," & vbCrLf
            End If
        Next

        S.Append("]")
        Return S.ToString
    End Function

    Function adoGetDDL(ByVal TableHandle As String)
        Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
        If Not TypeOf fHandle Is clsTableHandleAdo Then Return Nothing

        Dim adoHandle As clsTableHandleAdo = fHandle
        If adoHandle.SqlActiveConnection Is Nothing Then Throw New Exception("Ado Table handle not valid")
        If adoHandle.SqlActiveConnection.State = ConnectionState.Closed Then adoHandle.SqlActiveConnection.Open()

        Dim dt As DataTable = New DataTable(adoHandle.TableName)
        ' Read Item from database for data adapter update
        Dim DA As IDbDataAdapter = NewDataAdapter("SELECT * FROM [" & adoHandle.TableName & "] WHERE 1=0", adoHandle.SqlActiveConnection)

        SetupSqlCommands(DA, adoHandle.TableName, adoHandle.PrimaryKeyColumnName)
        CType(DA, System.Data.Common.DbDataAdapter).Fill(dt)

        Try
            Dim js As System.Collections.Generic.Dictionary(Of System.String, System.Object) = JSON("{}")

            For Each C As DataColumn In dt.Columns

                Dim js_defs As System.Collections.Generic.Dictionary(Of System.String, System.Object) = JSON("{}")

                js_defs.Add("DataType", Field(C.DataType.ToString, ".", 2))
                js_defs.Add("AllowDBNull", C.AllowDBNull)
                js_defs.Add("AutoIncrement", C.AutoIncrement)
                js_defs.Add("AutoIncrementSeed", C.AutoIncrementSeed)
                js_defs.Add("AutoIncrementStep", C.AutoIncrementStep)
                js_defs.Add("Caption", C.Caption)
                js_defs.Add("ColumnName", C.ColumnName)
                js_defs.Add("MaxLength", C.MaxLength)
                js_defs.Add("DefaultValue", C2Str(C.DefaultValue))
                js_defs.Add("Ordinal", C.Ordinal)
                js_defs.Add("ReadOnly", C.ReadOnly)
                js_defs.Add("Unique", C.Unique)
                js_defs.Add("Prefix", C.Prefix)

                For Each PropertyName As String In C.ExtendedProperties.Keys
                    If TypeOf C.ExtendedProperties.Item(PropertyName) Is String Then
                        js_defs.Add(PropertyName, C.ExtendedProperties.Item(PropertyName))

                    ElseIf TypeOf C.ExtendedProperties.Item(PropertyName) Is Boolean Then
                        js_defs.Add(PropertyName, C.ExtendedProperties.Item(PropertyName))

                    ElseIf IsNumeric(C.ExtendedProperties.Item(PropertyName)) Then
                        js_defs.Add(PropertyName, C.ExtendedProperties.Item(PropertyName))

                    End If

                Next

                js.Add(C.ColumnName, js_defs)
            Next

            Return C2Str(js, False)
        Catch ex As Exception
            Throw New Exception("**" & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Public Function adoCreateColumn(ByVal TableHandle As String, ByVal ColName As String, ByVal DefinedSize As Integer, ByVal dotNetType As String, ByVal NullsAllowed As Boolean, ByVal AutoKey As Boolean, ByVal AllowZeroLength As Boolean, ByVal DefaultValue As String, ByVal Description As String) As String
        Dim Errors As String = ""

        Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
        If Not TypeOf fHandle Is clsTableHandleAdo Then Return "TableHande invalid"

        Dim adoHandle As clsTableHandleAdo = fHandle
        If adoHandle.SqlActiveConnection Is Nothing Then Return "Ado Table handle not valid"
        If adoHandle.SqlActiveConnection.State = ConnectionState.Closed Then adoHandle.SqlActiveConnection.Open()

        If CreateColumn(adoHandle.SqlActiveConnection, Nothing, fHandle.TableName, ColName, DefinedSize, dotNetType, NullsAllowed, AutoKey, AllowZeroLength, DefaultValue, Description, Errors) Then Return ""
        Return Errors
    End Function

    Public Function adoRenameColumn(ByVal TableHandle As String, ByVal OldColName As String, ByVal NewColName As String) As String
        Dim Errors As String = ""

        Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
        If Not TypeOf fHandle Is clsTableHandleAdo Then Return "TableHande invalid"

        Dim adoHandle As clsTableHandleAdo = fHandle
        If adoHandle.SqlActiveConnection Is Nothing Then Return "Ado Table handle not valid"
        If adoHandle.SqlActiveConnection.State = ConnectionState.Closed Then adoHandle.SqlActiveConnection.Open()

        If RenameColumn(adoHandle.SqlActiveConnection, Nothing, fHandle.TableName, OldColName, NewColName, Errors) Then Return ""
        Return Errors
    End Function

    Public Function adoDeleteColumn(ByVal TableHandle As String, ByVal ColName As String) As String
        Dim Errors As String = ""

        Dim fHandle As clsTableHandle = cachedfHandles(TableHandle)
        If Not TypeOf fHandle Is clsTableHandleAdo Then Return "TableHande invalid"

        Dim adoHandle As clsTableHandleAdo = fHandle
        If adoHandle.SqlActiveConnection Is Nothing Then Return "Ado Table handle not valid"
        If adoHandle.SqlActiveConnection.State = ConnectionState.Closed Then adoHandle.SqlActiveConnection.Open()

        If DeleteColumn(adoHandle.SqlActiveConnection, Nothing, fHandle.TableName, ColName, Errors) Then Return ""
        Return Errors
    End Function


    Function adoBrowseForFile(Optional ByVal Title As String = "Open File Dialog", Optional ByVal InitialDirectory As String = "C:\", Optional ByVal Filter As String = "All files (*.*)|*.*|All files (*.*)|*.*") As String
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.Title = Title
        fd.InitialDirectory = InitialDirectory
        fd.Filter = Filter
        fd.RestoreDirectory = True
        If fd.ShowDialog() = DialogResult.OK Then Return fd.FileName Else Return ""
    End Function
End Class


Public Class clsDBConnection

    ' Each account is an open database connection.
    '
    ' To add new types of connections, do the following:
    '
    '    1. Add code to recognize the type in dbInitialize
    '    2. Create a new defs for all "If ConnectionType =" statements
    '        see: ListFiles, OpenBool
    '    3. Create a derived child class of clsTableHandle for the appropriate DB
    '
    Public Const JSB_SelectLists As String = "JSB_SelectLists"
    Public Const JSB_MD As String = "MD"


    ' SYSTEM only variables
    Public mUserID As String = ""
    Dim mPassword As String = ""

    Dim mPointerFileHandle As clsTableHandle = Nothing  ' or clsTableHandleDos
    Dim mMDHandle As clsTableHandle = Nothing  ' or clsTableHandleDos
    Dim mSystemTableHandle As clsTableHandle = Nothing

    Private mSelectLists() As rSelectList ' Active/Nothing SelectN's

    Public mCachedDBAccounts As Collection   ' clsjdbAccount

    Public mAttachedDB As clsjdbAccount = Nothing
    Public mUrlAccount As clsjdbAccount = Nothing


    Public Sub New()
        MakeNew(UrlAccount())
    End Sub

    Public Sub New(UrlAccount As String)
        MakeNew(UrlAccount)
    End Sub

    Public Sub MakeNew(UrlAccount As String)
        Dim Errors As String = ""
        mCachedDBAccounts = New Collection

        Dim mMyAccount As clsjdbAccount = dbOpenAccount(UrlAccount, Errors)

        mUrlAccount = mMyAccount ' By default all create-file's will happen here
        mCachedDBAccounts.Add(mMyAccount, ".")

        If Not OpenBool("", MapRootAccount("") & "\SYSTEM\", mSystemTableHandle, Errors, True) Then Throw New Exception(Errors)
        If Not OpenBool("", JSB_SelectLists, mPointerFileHandle, Errors, True, mMyAccount) Then Throw New Exception(Errors)
        If Not OpenBool("", JSB_MD, mMDHandle, Errors) Then Throw New Exception(Errors)
        '' Copy all MD items from NEWACT as a starter
        'If Not OpenBool("", JSB_MD, mMDHandle, Errors, True) Then Throw New Exception(Errors)
        'My.Computer.FileSystem.CopyDirectory(MapRootAccount("NEWACT"), MapRootAccount(""), True)
        'End If

        dbClearSelectAll()
    End Sub

    Public Function adoAttach(ByVal AccountName As String, ByVal ConnectionItem As String, ByVal UserID As String, ByVal Password As String, ByVal CreateIt As Boolean, ByVal TimeOutSecs As Integer)
        Dim Errors As String = ""

        If AccountName = "" Then
            mAttachedDB = Nothing
            Return Nothing
        End If

        ' Check already open Account collection
        If mCachedDBAccounts.Contains(VB6.UCase("@" & AccountName & UserID)) Then
            mAttachedDB = mCachedDBAccounts(VB6.UCase("@" & GetFName(AccountName) & UserID))
            Return mAttachedDB
        End If

        Dim dbAccount As New clsjdbAccount
        If dbAccount.Init(AccountName, ConnectionItem, Errors, CreateIt, TimeOutSecs, UserID, Password) = False Then Throw New Exception(Errors)

        mCachedDBAccounts.Add(dbAccount, VB6.UCase("@" & AccountName & UserID))
        mAttachedDB = dbAccount

        Return dbAccount
    End Function

    ' Looks in SYSTEM for a definition of the account, then tries to make a connection
    ' If successful, caches a handle for the account in mCachedAccounts
    ' Open a database ACCOUNT
    ' Resolution priority
    '    If a ConnectionString is given, it is used.
    '      If the AccountName is empty, is a derived by parsing the MapPath
    '        If the AccountName contains "\" its is used as a path
    '           Else the MapPath should contain the path to the filefolder base
    '        
    '
    '  The Connection string may be <1> Type, <2> Provider String
    '
    '  Connection Types are
    '    D  : Data Provider String (adox)
    '    F  : File Folder Path
    '    C  : Configuration Section of Web Config
    '    GS : Google SpreadSheet
    '
    Function dbOpenAccount(ByVal AccountName As String, ByRef Errors As String, Optional ByRef UserID As String = "", Optional ByRef Password As String = "", Optional ByRef TimeOutSecs As Short = 0, Optional ByVal CreateIt As Boolean = False) As clsjdbAccount
        Dim ConnectString As String = ""

        ' Check already open Account collection
        If mCachedDBAccounts.Contains(VB6.UCase(GetFName(AccountName) & UserID)) Then Return mCachedDBAccounts(VB6.UCase(GetFName(AccountName) & UserID))

        ' Read definition from the SYSTEM table to know where this account is located
        If Not SystemTableHandle() Is Nothing Then
            If Not SystemTableHandle.dbReadBool(GetFName(AccountName), ConnectString) Then ConnectString = ""
            If VB6.InStr(ConnectString, vbCrLf) Then ConnectString = VB6.Replace(ConnectString, vbCrLf, Chr(254))

            Dim ConnectionType As String = Extract(ConnectString, 1)
            ' Format Connect string
            If VB6.Right(ConnectionType, 1) <> "!" AndAlso Extract(ConnectString, 4) <> "" Then
                ConnectString = Replace(ConnectString, 1, 0, 0, ConnectionType & "!")
                ConnectString = Replace(ConnectString, 4, 0, 0, STX(AesEncrypt(Extract(ConnectString, 4))))
                SystemTableHandle.dbWrite(ConnectString, GetFName(AccountName))
            End If
        End If

        Dim dbAccount As New clsjdbAccount
        If dbAccount.Init(AccountName, ConnectString, Errors, CreateIt, TimeOutSecs, UserID, Password) = False Then Return Nothing


        mCachedDBAccounts.Add(dbAccount, VB6.UCase(GetFName(AccountName) & UserID))
        dbOpenAccount = dbAccount
    End Function

    Function dbAttachDB(ByVal AccountName As String, ByRef Errors As String, Optional ByRef UserID As String = "", Optional ByRef Password As String = "") As Boolean
        Dim dbAccount As clsjdbAccount = dbOpenAccount(AccountName, Errors, UserID, Password)
        If dbAccount Is Nothing Then Return False
        mAttachedDB = dbAccount
        dbClearCachedTables()
        Return True
    End Function

    Public ReadOnly Property dbAttachedDBName() As String
        Get
            Return mAttachedDB.AcountName
        End Get
    End Property

    Public ReadOnly Property dbAttachedDBType() As String
        Get
            If mAttachedDB Is Nothing Then Return ""
            Return mAttachedDB.ConnectionType
        End Get
    End Property

    Protected Overrides Sub Finalize()
        mPointerFileHandle = Nothing
        mMDHandle = Nothing
        mSystemTableHandle = Nothing
        mCachedDBAccounts = Nothing
        MyBase.Finalize()
    End Sub

    Public Function dbListFiles() As ArrayList
        Dim FNames As New ArrayList

        If dbAttachedDBType = "D" Then
            If mAttachedDB.AdoConnection.State = ConnectionState.Closed Then mAttachedDB.AdoConnection.Open()

            If mAttachedDB.AdoConnection Is Nothing Then Return FNames
            Dim schema As System.Data.DataTable = Nothing

            If TypeOf mAttachedDB.AdoConnection Is OleDb.OleDbConnection Then
                schema = CType(mAttachedDB.AdoConnection, OleDb.OleDbConnection).GetSchema("tables")

            ElseIf TypeOf mAttachedDB.AdoConnection Is SqlClient.SqlConnection Then
                schema = CType(mAttachedDB.AdoConnection, SqlClient.SqlConnection).GetSchema("tables")

            ElseIf TypeOf mAttachedDB.AdoConnection Is SQLite.SQLiteConnection Then
                schema = CType(mAttachedDB.AdoConnection, SQLite.SQLiteConnection).GetSchema("tables")

            ElseIf TypeOf mAttachedDB.AdoConnection Is Odbc.OdbcConnection Then
                schema = CType(mAttachedDB.AdoConnection, Odbc.OdbcConnection).GetSchema("tables")

#If IncludeCompactFrameWork Then
        ElseIf TypeOf AdoConnection Is SqlServerCe.SqlCeConnection Then
            schema =CType(AdoConnection, SqlServerCe.SqlCeConnection).GetSchema ("tables")
#End If
            End If

            For Each dr As DataRow In schema.Rows
                FNames.Add(dr("TABLE_NAME").ToString)
            Next
            schema.Dispose()
            mAttachedDB.AdoConnection.Close()
            Return FNames
        End If

        If dbAttachedDBType = "J" Then
            ' Login
            Dim JsbServer As String = mAttachedDB.ConnectionString & "/ServerListFiles"
            Dim Result As String = ""
            If Not UrlFetch("GET", JsbServer, "", "", Result, "", False, 33) Then Return FNames

            If VB6.Left(Result, 1) <> "{" Then Return FNames
            Dim JSonResult As Object = JSON(Result)
            If Val(JSonResult("_restFunctionResult")) = 0 Then Return FNames
            Return New ArrayList(VB6.Split(JsonNthValue(JSonResult, 1), Chr(254)))
        End If

        Dim dDir As DirectoryInfo
        If dbAttachedDBType = "F" Then
            dDir = New DirectoryInfo(mAttachedDB.ConnectionString)
        Else
            dDir = New DirectoryInfo(MapRootAccount(""))
        End If

        Dim fFileSystemInfo As FileSystemInfo
        Dim I As Integer = 1
        For Each fFileSystemInfo In dDir.GetFileSystemInfos()
            If (fFileSystemInfo.Attributes And FileAttributes.Directory) <> 0 Then
                FNames.Add(fFileSystemInfo.Name)
            End If
        Next

        Return FNames
    End Function

    Function dbDate() As Integer
        Return Fix(Now.ToOADate - DateSerial(1967, 12, 31).ToOADate)
    End Function

    Function dbTime() As Integer
        Return Hour(Now) * 60 ^ 2 + Minute(Now) * 60 + Second(Now) + (Now.Millisecond / 1000)
    End Function

    Function dbTimeDate() As String
        Return Now().ToString("yyy-MM-dd HH:mm:ss") ' Makes it sortable as yyyy-mm-dd hh:mm:ss
    End Function

    Function dbSelectList(ByRef ListNo As Object) As rSelectList
        On Error GoTo Nope

        If TypeOf ListNo Is rSelectList Then
            Return ListNo

        ElseIf IsNumeric(ListNo) Then
            Return mSelectLists(ListNo)

        Else
            Return mSelectLists(0)
        End If

        Exit Function
Nope:
        Err.Clear()
    End Function

    Public Function SystemTableHandle() As clsTableHandle ' or clsTableHandleDos
        Return mSystemTableHandle
    End Function

    Public Function PFHandle() As clsTableHandle ' or clsTableHandleDos
        Return mPointerFileHandle
    End Function

    Public Function MDHandle() As clsTableHandle ' or clsTableHandleDos
        Return mMDHandle
    End Function

    Public ReadOnly Property dbWho() As String
        Get
            Return UrlAccount() & " " & dbAttachedDBName
        End Get
    End Property

    Public ReadOnly Property ConnectString() As String
        Get
            If mAttachedDB Is Nothing Then Return ""
            Return mAttachedDB.ConnectionString
        End Get
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return mUserID
        End Get
    End Property

    Public ReadOnly Property Password() As String
        Get
            Return mPassword
        End Get
    End Property


    'Public Function MakeSqlTableName(ByVal DictData As String, ByVal TableName As String) As String
    '    Dim CH As String
    '    Dim TmpResult As String
    '    Dim I As Short

    '    ' Strip possible DICT/DATA from tablename
    '    If VB6.Left(VB6.UCase(TableName), 5) = "DICT " Then
    '        TableName = Field(VB6.Mid(TableName, 6), ",", 1)
    '        If DictData = "" Then DictData = "DICT"

    '    ElseIf VB6.Left(VB6.UCase(TableName), 5) = "DATA " Then
    '        TableName = VB6.Mid(TableName, 6)
    '    End If

    '    If VB6.UCase(DictData) = "DICT" Then TableName = "MetaDictionary" ' TableName & "]D"

    '    ' Return a TableName that is acceptable by SQL standards
    '    CH = VB6.Mid(TableName, 1, 1)
    '    If (CH >= "a" And CH <= "z") Or (CH >= "A" And CH <= "Z") Then
    '        TmpResult = CH
    '    Else
    '        TmpResult = "X_"
    '    End If

    '    For I = 2 To Len(TableName)
    '        CH = VB6.Mid(TableName, I, 1)
    '        If (CH >= "a" And CH <= "z") Or (CH >= "A" And CH <= "Z") Then
    '            TmpResult = TmpResult & CH
    '        ElseIf (CH >= "0" And CH <= "9") Then
    '            TmpResult = TmpResult & CH
    '        Else
    '            TmpResult = TmpResult & "_"
    '        End If
    '    Next I

    '    'If TmpResult = "MEMO" Then TmpResult = "." & TmpResult
    '    MakeSqlTableName = TmpResult
    'End Function

    Public Function dbOpen(ByVal DictData As String, ByVal TableName As String, ByRef ToFile As clsTableHandle, ByRef Errors As String, Optional ByVal FromAccount As clsjdbAccount = Nothing) As Boolean
        Dim QREC As Object = "", QType As String
        Dim NewActName As String, NewFName As String
        Dim IsDict As Boolean

        If FromAccount Is Nothing Then FromAccount = mUrlAccount
        If TableName = "" Then Err.Raise(1, , "Open(No TableName)")

        ToFile = Nothing
        Do While VB6.UCase(VB6.Left(TableName, 5)) = "DICT " Or VB6.UCase(VB6.Left(TableName, 5)) = "DATA "
            DictData = VB6.Left(TableName, 4)
            TableName = VB6.Mid(TableName, 6)
        Loop

        If DictData = "" Then DictData = "DATA" Else DictData = VB6.UCase(DictData)
        IsDict = DictData = "DICT"

        If FromAccount.Tables.Contains(DictData & "_" & VB6.UCase(TableName)) Then
            ToFile = FromAccount.Tables.Item(DictData & "_" & VB6.UCase(TableName))
            Return True
        End If

        ' Check local copy before Q filing
        If OpenBool(DictData, TableName, ToFile, Errors, False, FromAccount) Then Return True
        If FromAccount IsNot mAttachedDB Then
            If OpenBool(DictData, TableName, ToFile, Errors, False, mAttachedDB) Then Return True
        End If

        If MDHandle() IsNot Nothing Then
            If TableName = "" OrElse VB6.LCase(TableName) = VB6.LCase(MDHandle.TableName) Then
                ToFile = MDHandle()
                Return True
            End If

            ' Read the MD filedef and check for a Q ptr
            If dbRead(QREC, MDHandle, TableName) Then
                ' MD may contain:
                '  <D> entries: DOS file directions in <2>
                '  <Q> entries: option AccountName in <2>, current Account used if empty, optional TableName in <3> - current tablename used if missing)
                '   entries: Provider entry in <2>, TableName in <3>
                If VB6.InStr(QREC, vbCrLf) Then QREC = VB6.Replace(QREC, vbCrLf, Chr(254))
                QType = VB6.LCase(Field(QREC, Chr(254), 1))
                NewActName = Field(QREC, Chr(254), 2)
                NewFName = Field(QREC, Chr(254), 3)
                If VB6.InStr(QType, Chr(253)) Then
                    For I As Integer = 1 To DCount(QType, Chr(253))
                        Dim FT As String = VB6.LCase(Field(QType, Chr(253), I))
                        If FT = "q" Or FT = "qd" Or FT = "d" Or FT = "f" Then
                            NewActName = Field(NewActName, Chr(253), I)
                            NewFName = Field(NewFName, Chr(253), I)
                            QType = FT
                            Exit For
                        End If
                    Next
                End If

                If QType = "q" Or QType = "qd" Then

                    ' Change filename?
                    If NewFName <> "" And NewFName <> TableName Then TableName = NewFName
                    If QType = "qd" And IsDict Then NewActName = ""

                    If NewActName <> "" And NewActName <> FromAccount.AcountName Then
                        FromAccount = dbOpenAccount(NewActName, Errors, mUserID, mPassword, 0, False)
                        If FromAccount Is Nothing Then Return False
                    End If

                ElseIf QType = "d" Then
                    ' NewActName is a provider string
                    Dim TAct As clsjdbAccount = New clsjdbAccount
                    If Not TAct.Init("", NewActName, Errors) Then Return False
                    If NewFName = "" Then NewFName = TableName
                    Return OpenBool(DictData, NewFName, ToFile, Errors, False, TAct)

                ElseIf QType = "f" Then
                    If IsDict Then
                        If NewFName = "" Then Return False ' No dict give in line 3
                        TableName = NewFName
                    Else
                        TableName = NewActName
                    End If
                End If
            End If
        End If

        Dim CanOpen As Boolean = OpenBool(DictData, TableName, ToFile, Errors, False, FromAccount)
        If CanOpen Then Return True

        If IsDict Then
            ' If the DATA portion exists, create the DICT if necessary
            CanOpen = OpenBool("", TableName, ToFile, Errors, False, FromAccount)
            If Not CanOpen Then Return False
            Return OpenBool(DictData, TableName, ToFile, Errors, True, FromAccount)
        End If

        ' Is it in the attacheddb?
        If FromAccount Is mUrlAccount AndAlso mAttachedDB IsNot Nothing Then
            Return OpenBool(DictData, TableName, ToFile, Errors, False, mAttachedDB)
        End If

        Return False
    End Function


    ' Create a TABLE file
    Public Function dbCreateFile(ByVal DictData As String, ByVal TableName As String, ByRef TableHandle As clsTableHandle, ByRef Errors As String) As Boolean
        dbCreateFile = OpenBool(DictData, TableName, TableHandle, Errors, True)
    End Function

    Public Sub ClearReadUCache()
        If mAttachedDB Is Nothing Then Exit Sub

        For Each Table As clsTableHandle In mAttachedDB.Tables
            If Table.ReadUCache IsNot Nothing AndAlso Table.ReadUCache.Count Then Table.ReadUCache = New Collection
        Next
    End Sub

    Public Sub dbExitServer()
        If mAttachedDB Is Nothing Then Exit Sub

        For Each Table As clsTableHandle In mAttachedDB.Tables
            Table.dbExitServer()
        Next

        For Each Account As clsjdbAccount In mCachedDBAccounts
            Account.closeAdo()
        Next
    End Sub

    Public Sub dbTransactionBegin()
        If mAttachedDB Is Nothing Then Exit Sub

        For Each Table As clsTableHandle In mAttachedDB.Tables
            Table.dbTransactionBegin()
        Next
    End Sub

    Public Sub dbTransactionEnd()
        If mAttachedDB Is Nothing Then Exit Sub

        For Each Table As clsTableHandle In mAttachedDB.Tables
            Table.dbTransactionCommit()
        Next
    End Sub

    ' Open a TABLE in this account
    Public Function OpenBool(ByVal DictData As String, ByVal TableName As String, ByRef TableHandle As clsTableHandle, ByRef Errors As String, Optional ByVal CreateIt As Boolean = False, Optional ByVal FromAccount As clsjdbAccount = Nothing) As Boolean
        Dim mTableHandle As clsTableHandle = Nothing, IsDos As Boolean = False, FirstChance = FromAccount Is Nothing
        If FromAccount Is Nothing Then FromAccount = mAttachedDB
        If FromAccount Is Nothing Then FromAccount = mUrlAccount

        If TableName = "" Then Err.Raise(1, , "OpenBool(No TableName)")
        If DictData = "" Then DictData = "DATA" Else DictData = VB6.UCase(DictData)

        Dim IsDict As Boolean = DictData = "DICT"
        IsDos = IsDict OrElse VB6.Left(TableName, 1) = "." OrElse VB6.InStr(TableName, "\")

        Dim IsHttp As Boolean = VB6.Left(LCase(TableName), 7) = "http://" Or VB6.Left(LCase(TableName), 8) = "https://"
        If IsDict AndAlso IsHttp Then Return False

        ' Check if table is already open
        If FromAccount.Tables.Contains(DictData & "_" & VB6.UCase(TableName)) Then
            TableHandle = FromAccount.Tables.Item(DictData & "_" & VB6.UCase(TableName))
            Return True
        End If

        ' Give DOS the first chance at this table
        If FirstChance And Not IsHttp Then
            mTableHandle = New clsTableHandleDos
            CType(mTableHandle, clsTableHandleDos).IsDict = IsDict
            If mTableHandle.dbOpenBool(FromAccount, TableName, Errors, False) Then
                TableHandle = mTableHandle
                FromAccount.Tables.Add(TableHandle, DictData & "_" & VB6.UCase(TableName))
                Return True
            End If
        End If

        If IsDict Then
            mTableHandle = New clsTableHandleDos
            CType(mTableHandle, clsTableHandleDos).IsDict = True

        ElseIf IsHttp Then
            mTableHandle = New clsTableHandleHttp

        ElseIf IsDos Then
            ' assume DOS
            mTableHandle = New clsTableHandleDos

        ElseIf FromAccount.ConnectionType = "D" Then
            mTableHandle = New clsTableHandleAdo

            ' ElseIf FromAccount.ConnectionType = "EXX" Then
            '    mTableHandle = New clsTableHandleExchangeServer

        Else
            ' assume DOS
            mTableHandle = New clsTableHandleDos
        End If

        If mTableHandle.dbOpenBool(FromAccount, TableName, Errors, CreateIt) Then
            TableHandle = mTableHandle
            FromAccount.Tables.Add(TableHandle, DictData & "_" & VB6.UCase(TableName))
            Return True
        End If

        Dim Errors2 As String = ""
        If IsDict Then
            ' If the data portion exists, auto create a DICT
            Dim op As clsTableHandle = Nothing
            If Not OpenBool("", TableName, op, Errors2) Then Return False
            CreateIt = True
        Else
            ' Attempt a second open, to see if it's a DOS file
            If TypeOf mTableHandle Is clsTableHandleDos Then Return False
            mTableHandle = New clsTableHandleDos
        End If

        If mTableHandle.dbOpenBool(FromAccount, TableName, Errors2, CreateIt) Then
            TableHandle = mTableHandle
            FromAccount.Tables.Add(TableHandle, DictData & "_" & VB6.UCase(TableName))
            Return True
        End If

        Return False
    End Function

    Function dbRead(ByRef DataRecord As Object, ByVal TableHandle As clsTableHandle, ByRef ItemID As String, Optional ByVal IsReadU As Boolean = False) As Boolean
        Dim jDataRecord As String = ""
        Dim ItemFound As Boolean

        If Not IsReference(TableHandle) Then Throw New Exception("Invalid file (OBJECT) handle")
        If TableHandle Is Nothing Then Return False

        ItemFound = TableHandle.dbReadBool(ItemID, jDataRecord, IsReadU)
        If ItemFound Then DataRecord = NumTest(jDataRecord)
        Return ItemFound
    End Function

    Function dbReadJSon(ByRef jsonDataRecord As Object, ByVal TableHandle As clsTableHandle, ByRef ItemID As String, Optional ByVal IsReadU As Boolean = False) As Boolean
        jsonDataRecord = Nothing
        If Not IsReference(TableHandle) Then Throw New Exception("Invalid file (OBJECT) handle")
        If TableHandle Is Nothing Then Return False
        Return TableHandle.dbReadJSon(ItemID, jsonDataRecord, IsReadU)
    End Function

    Function dbReadXML(ByRef XmlDataRecord As Object, ByVal TableHandle As clsTableHandle, ByRef ItemID As String, Optional ByVal IsReadU As Boolean = False) As Boolean
        XmlDataRecord = Nothing
        If Not IsReference(TableHandle) Then Throw New Exception("Invalid file (OBJECT) handle")
        If TableHandle Is Nothing Then Return False
        Return TableHandle.dbReadXML(ItemID, XmlDataRecord, IsReadU)
    End Function

    Function dbReadNext(ByVal rSelectOrListNo As Object, ByRef Result As Object) As Boolean
        Dim sResult As String = ""
        Dim js As rSelectList

        js = dbSelectList(rSelectOrListNo)
        If js Is Nothing Then Return False

        dbReadNext = js.ReadNextBool(sResult)

        Result = sResult
    End Function

    Function dbReadNextValue(ByRef rSelectOrListNo As Object, ByRef Result As Object, ByVal Value As Integer) As Object
        Dim js As rSelectList

        js = dbSelectList(rSelectOrListNo)
        If js Is Nothing Then Return Nothing


        dbReadNextValue = js.ReadNextBool(Result, Value)
    End Function

    Function dbReadNextSubValue(ByRef rSelectOrListNo As Object, ByRef Result As Object, ByRef Value As Integer, ByRef SubValue As Integer) As Object
        Dim sResult As String = ""
        Dim js As rSelectList

        js = dbSelectList(rSelectOrListNo)
        If js Is Nothing Then Return Nothing

        dbReadNextSubValue = js.ReadNextBool(sResult, Value, SubValue)
        Result = sResult
    End Function

    Function dbReadV(ByRef DataRecord As Object, ByVal TableHandle As clsTableHandle, ByVal ItemID As String, ByVal FieldNo As Short, ByVal IsReadU As Boolean) As Boolean
        Dim Result As String = ""
        If Val(FieldNo) = 0 Then Return ItemID

        If Not IsReference(TableHandle) Then Throw New Exception("Invalid file (OBJECT) handle")
        If TableHandle Is Nothing Then Return False

        If TableHandle.dbReadBool(ItemID, Result, IsReadU) Then
            If VB6.InStr(Result, vbCrLf) Then Result = VB6.Replace(Result, vbCrLf, Chr(254))

            Result = Extract(Result, FieldNo)

            If IsReference(DataRecord) Then
                DataRecord = Result
            Else
                DataRecord = NumTest(Result)
            End If
            Return True
        End If

        Return False
    End Function

    Function dbSelectV(ByVal TableHandle As clsTableHandle) As rSelectList
        If IsReference(TableHandle) Then
            Return TableHandle.SelectFile
        Else
            dbSelectV = mSelectLists(0)
        End If
    End Function

    Function dbSelectVX(ByVal ColumnList As String, ByVal TableHandle As clsTableHandle, ByVal WhereClause As String, ByVal doMerging As Boolean) As rSelectList
        If IsReference(TableHandle) Then
            Dim result As rSelectList = TableHandle.SelectFileX(ColumnList, WhereClause)

            ' Is there an active Select?
            If doMerging AndAlso mSelectLists(0) IsNot Nothing AndAlso mSelectLists(0).EOF = False Then
                Dim LimitList As String = mSelectLists(0).GetListOfPKs
                If LimitList <> "" Then result.LimitToListOfPKs(LimitList)
            End If

            Return result
        Else
            Return mSelectLists(0)
        End If
    End Function

    Sub dbSelectN(ByVal TableHandle As clsTableHandle, ByVal ListNo As Short)
        'If Not TypeOf TableHandle Is Object  Then Err.Raise 9999, "clsDBConnection", "Invalid file (OBJECT) handle"

        If ListNo > VB6.UBound(mSelectLists) Then ReDim Preserve mSelectLists(ListNo)

        mSelectLists(ListNo) = TableHandle.SelectFile
    End Sub

    Function dbSqlSelect(ByVal SqlCommand As String) As rSelectList
        Dim SlResult As New rSelectList, ConnectionType As String = mAttachedDB.ConnectionType

        If ConnectionType = "D" Then
            ' A Data Provider String
            SlResult.PrimeSqlSelect(mAttachedDB.AdoConnection, SqlCommand, False)
            SlResult.ActiveSelect()
            Return SlResult

        ElseIf ConnectionType = "J" Then
            ' J : JSB Server http://some.domain/jsb/ServerLogin
            Dim JsbServer As String = mAttachedDB.ConnectionString
            Dim myTimeOut As Integer = 33 ' seconds

            Dim Result As String = "", Errors As String = "", Url As String = JsbServer & "/ServerSqlSelect?_p1=" & UrlEncode(SqlCommand)

            If Not UrlFetch("GET", Url, "", "", Result, "", False, myTimeOut) Then Throw New Exception("network error:" & Result)

            Dim JSonResult As Object = restFunctionResult(Result, "_p3")

            If TypeOf JSonResult("_p2") Is ArrayList Then
                Dim dt As DataTable = Json2DataTable(JSonResult("_p2"), "results", Errors)
                SlResult.SetDataTable(dt, False)
                Return SlResult
            End If

            SlResult.SetDynamicArray("", False)
            Return SlResult

        Else
            Throw New Exception("No SQL-SELECT on non-SQL databases")
        End If

    End Function

    Function restFunctionResult(ByRef Result As String, errTag As String) As Object
        If VB6.Left(Result, 1) <> "{" Then Throw New Exception(Result)
        Dim JSonResult As Dictionary(Of String, Object) = Nothing
        JSonResult = JSON(Result)
        If JSonResult.ContainsKey("_restFunctionResult") = False Then Throw New Exception("JSON call missing tag _restFunctionResult")

        If Val(JSonResult("_restFunctionResult")) <> 0 Then Return JSonResult
        If JSonResult.ContainsKey(errTag) Then Throw New Exception(JSonResult(errTag))
        If JSonResult.ContainsKey("ERRORS") Then Throw New Exception(JSonResult("ERRORS"))
        If JSonResult.ContainsKey("errors") Then Throw New Exception(JSonResult("errors"))
        Throw New Exception("json result missing tag " & errTag)
    End Function

    Function dbSqlScalar(ByVal SqlCommand As String) As Object
        If mAttachedDB.AdoConnection.State = ConnectionState.Closed Then mAttachedDB.AdoConnection.Open()
        Dim DA As IDbCommand = NewCommand(SqlCommand, mAttachedDB.AdoConnection)
        Dim Result = DA.ExecuteScalar()
        mAttachedDB.AdoConnection.Close()
        Return Result
    End Function

    Sub dbWrite(ByVal Item As Object, ByVal TableHandle As clsTableHandle, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        'If Not TypeOf TableHandle Is clsTableHandle Then Err.Raise 9999, "clsDBConnection", "Invalid file (OBJECT) handle"
        If TableHandle Is Nothing Then Return
        TableHandle.dbWrite(Item, ItemID, IsWriteU)
    End Sub

    Sub dbWriteJSon(ByVal Item As Object, ByVal TableHandle As clsTableHandle, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        'If Not TypeOf TableHandle Is clsTableHandle Then Err.Raise 9999, "clsDBConnection", "Invalid file (OBJECT) handle"
        If TableHandle Is Nothing Then Return
        TableHandle.dbWriteJSon(Item, ItemID, IsWriteU)
    End Sub

    Sub dbWriteXML(ByVal Item As Object, ByVal TableHandle As clsTableHandle, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        'If Not TypeOf TableHandle Is clsTableHandle Then Err.Raise 9999, "clsDBConnection", "Invalid file (OBJECT) handle"
        If TableHandle Is Nothing Then Return
        TableHandle.dbWriteXML(Item, ItemID, IsWriteU)
    End Sub

    Sub dbWriteV(ByVal Str_Renamed As String, ByVal TableHandle As clsTableHandle, ByVal ItemID As String, ByVal FieldNo As Integer, Optional ByVal IsWriteU As Boolean = False)
        'If Not TypeOf TableHandle Is Object  Then Err.Raise 9999, "clsDBConnection", "Invalid file (OBJECT) handle"
        Dim Item As String = ""

        Call TableHandle.dbReadBool(ItemID, Item, IsWriteU)

        Item = Replace(Item, FieldNo, 0, 0, CStr(FieldNo))

        TableHandle.dbWrite(Item, ItemID, IsWriteU)
    End Sub

    Public Sub Execute(ByVal SQL As String)
        Dim DA As New SqlClient.SqlCommand
        DA.CommandText = SQL
        DA.Connection = mAttachedDB.AdoConnection
        DA.ExecuteScalar()
    End Sub

    Public Sub dbClearCachedTables()
        ' Force Open to re-open
        For Each Account As clsjdbAccount In mCachedDBAccounts
            Account.Tables.Clear()
        Next
    End Sub

    Public Sub dbDeleteFile(ByVal TableHandle As clsTableHandle)
        TableHandle.dbDeleteFile()

        If TypeOf TableHandle Is clsTableHandleDos AndAlso CType(TableHandle, clsTableHandleDos).IsDict Then
            TableHandle.MyDBAccount.Tables.Remove("DICT_" & VB6.UCase(TableHandle.TableName))
        Else
            TableHandle.MyDBAccount.Tables.Remove("DATA_" & VB6.UCase(TableHandle.TableName))
        End If
    End Sub

    Sub dbClearFile(ByVal TableHandle As clsTableHandle)
        TableHandle.ClearFile()
    End Sub

    Sub dbClearSelect(ByVal ListNo As Short)
        mSelectLists(ListNo) = Nothing
    End Sub

    Sub dbClearSelectAll()
        Dim I As Short
        ReDim mSelectLists(10)
        mSelectLists(0) = New rSelectList
        For I = 1 To VB6.UBound(mSelectLists)
            mSelectLists(I) = Nothing
        Next
    End Sub

    Sub dbDelete(ByVal TableHandle As clsTableHandle, ByVal ItemID As String)
        TableHandle.dbDelete(ItemID)
        If mMDHandle Is TableHandle Then
            For Each act As clsjdbAccount In mCachedDBAccounts
                If act.ConnectionType = "F" Then
                    Dim RTables As New Collection
                    For Each Table As clsTableHandle In act.Tables
                        If VB6.UCase(Table.TableName) = VB6.UCase(ItemID) Then RTables.Add(Table)
                    Next
                    For Each table In RTables
                        act.Tables.Remove(table)
                    Next

                End If
            Next
        End If
    End Sub

    Sub dbDeleteList(ByVal ListName As String)
        If Not PFHandle() Is Nothing Then PFHandle.dbDelete(ListName)
    End Sub

    Sub dbWriteList(ByVal Item As Object, ByVal ListName As String)
        If PFHandle() Is Nothing Then
            If OpenBool("", JSB_SelectLists, mPointerFileHandle, True) = False Then
                Throw New Exception("Can't create JSB_SelectLists table")
            End If
        End If
        PFHandle.dbWrite(Item, ListName)
    End Sub

    ' FORM Takes Chr(AM) delimited list and makes it the ACTIVE select_struct 
    Sub dbFormList(ByRef DynamicArray As Object, ByVal ListNo As Short)
        If TypeOf DynamicArray Is rSelectList Then
            mSelectLists(ListNo) = DynamicArray

        ElseIf IsJsonObj(DynamicArray) Then
            Dim rlist As New rSelectList
            Dim SelectedItems As Object = JsonChild(DynamicArray, "SelectedItems")
            Dim SelectedIDs As Object = JsonChild(DynamicArray, "SelectedItemIDs")
            If IsArray(SelectedIDs) Then SelectedIDs = Join(SelectedIDs, Chr(254))
            rlist.SetJSonArray(SelectedIDs, SelectedItems, JsonChild(DynamicArray, "OnlyReturnItemIDs"))
            mSelectLists(ListNo) = rlist
        Else
            Dim rlist As New rSelectList
            rlist.SetDynamicArray(C2Str(DynamicArray), True)
            mSelectLists(ListNo) = rlist
        End If
    End Sub

    Function dbMakeSelect(ByRef DynamicArray As Object) As rSelectList
        If TypeOf DynamicArray Is rSelectList Then Return DynamicArray
        Dim rlist As New rSelectList

        If IsJsonObj(DynamicArray) Then
            Dim SelectedItems As Object = JsonChild(DynamicArray, "SelectedItems")
            Dim SelectedIDs As Object = JsonChild(DynamicArray, "SelectedItemIDs")
            If IsArray(SelectedIDs) Then SelectedIDs = Join(SelectedIDs, Chr(254))
            rlist.SetJSonArray(SelectedIDs, SelectedItems, JsonChild(DynamicArray, "OnlyReturnItemIDs"))
        Else
            rlist.SetDynamicArray(C2Str(DynamicArray), True)
        End If

        Return rlist
    End Function

    Sub dbGetListN(ByVal SavedListName As String, ByVal ListNo As Short)
        If ListNo > VB6.UBound(mSelectLists) Then ReDim Preserve mSelectLists(ListNo)
        On Error Resume Next
        mSelectLists(ListNo) = dbGetList(SavedListName)
        Err.Clear()
    End Sub

    Public Function dbGetList(ByVal ListName As String) As rSelectList
        Dim SL As New rSelectList
        Dim Item As String = ""
        If PFHandle() Is Nothing Then Return SL
        If PFHandle.dbReadBool(ListName, Item) Then SL.SetDynamicArray(Item, True)
        Return SL
    End Function

    Function dbReadList(ByRef rSelectOrListNo As Object, ByRef Result As Object) As Boolean
        Dim js As rSelectList
        js = dbSelectList(rSelectOrListNo)
        If js Is Nothing Then Return False
        If js.OnlyReturnItemIDs Then Result = js.GetListOfPKs Else Result = JSON(js.GetJSONString)
        dbReadList = True
    End Function

    Public Function dbCreateFile(ByVal FileSpec As String, ByRef Errors As String) As Boolean
        Dim TblHandle As Object = Nothing, DictData As String = "", DummyRec As String = ""

        If VB6.UCase(VB6.Left(FileSpec, 12)) = "CREATE-FILE " Then FileSpec = VB6.Mid(FileSpec, 13)

        Do While VB6.UCase(VB6.Left(FileSpec, 5)) = "DICT " Or VB6.UCase(VB6.Left(FileSpec, 5)) = "DATA "
            DictData = VB6.Left(FileSpec, 4)
            FileSpec = VB6.Mid(FileSpec, 6)
        Loop

        ' Create a SQL Lite database?
        If VB6.LCase(VB6.Right(FileSpec, 4)) = ".db3" Then
            FileSpec = VB6.Replace(FileSpec, "/", "\")
            If VB6.InStr(FileSpec, ":") = 0 And VB6.InStr(FileSpec, "\") = 0 Then FileSpec = "App_Data\_database\" & FileSpec
            Dim AName = DropExtension(GetFName(FileSpec))

            If Me.dbRead(DummyRec, mSystemTableHandle, AName) Then
                Errors = AName & " already exists in system"
                Return False
            Else
                Try
                    FileSpec = MapPath(FileSpec)
                    SQLite.SQLiteConnection.CreateFile(FileSpec)

                    ' Place connection string in system
                    Me.dbWrite("D" & Chr(254) & "Data Source=" & FileSpec & ";Version=3;New=False;Compress=False;Pooling=True;", mSystemTableHandle, AName)
                    Return True

                Catch ex As Exception
                    Errors = ex.Message
                    Return False
                End Try
            End If
        End If

        If DictData = "" Then
            ' Create the dict
            If Not dbCreateFile("DICT", Field(FileSpec, ",", 1), TblHandle, Errors) Then Return False
        End If

        ' create the DictData type
        dbCreateFile = dbCreateFile(DictData, FileSpec, TblHandle, Errors)
    End Function

    ' Save an active select list
    Public Sub dbSaveList(ByRef SelectOrItem As Object, ByVal ListName As String)
        If TypeOf SelectOrItem Is rSelectList Then
            dbWriteList(SelectOrItem, ListName) ' As item
        Else
            If mSelectLists(SelectOrItem) Is Nothing Then Exit Sub
            dbWriteList(mSelectLists(SelectOrItem).GetListOfPKs, ListName)
        End If
    End Sub

    Public Function dbSelectCount(ByRef ListNo As Object) As Integer
        If TypeOf ListNo Is rSelectList Then
            Return ListNo.RealCount
        ElseIf IsNumeric(ListNo) Then
            Return mSelectLists(ListNo).Count
        Else
            Return mSelectLists(0).Count
        End If
    End Function

    Public Function dbActiveSelect(Optional ByVal FileOrListNo As Object = Nothing) As Boolean
        Dim js As rSelectList
        js = dbSelectList(FileOrListNo)
        If js Is Nothing Then Return False
        Return js.ActiveSelect()
    End Function

End Class

Public Class clsjdbAccount
    Const iDefaultTimeoutSeconds As Short = 30
    Const iPrimaryKeyType As Short = 1
    Public ConnectionType As String = ""
    Public AcountName As String = ""
    Public Tables As Collection = New Collection ' of clsTableHandle
    '  The Connection string may be <1> Type, <2> Provider String
    '
    '  Connection Types are
    '    D  : Data Provider String (adox)
    '    F  : File Folder Path
    '    C  : Configuration Section of Web Config
    '    GS : Google SpreadSheet
    '    J: JSB Server

    ' Type P - Provider
    Public ConnectionString As String = "" ' May be valid directory for filesystem db
    ' Type D - ADO
    Private mAdoConnection As IDbConnection = Nothing  ' Connect to database where this table resides

    Public rpcsid As String
    Public rpcgid As String

    Public Sub closeAdo()
        If mAdoConnection IsNot Nothing AndAlso mAdoConnection.State = ConnectionState.Open Then
            mAdoConnection.Close()
        End If
    End Sub

    Public Function AdoConnection() As IDbConnection
        Return mAdoConnection
    End Function

    Function Init(ByVal AccountName As String, ByVal ConnectionItem As String, ByRef Errors As String, Optional ByVal CreateIt As Boolean = False, Optional ByVal TimeOutSecs As Integer = 0, Optional ByVal UserID As String = "", Optional ByVal Password As String = "") As Boolean
        Dim sPath As String = "", JetDB As Boolean, I As Short, J As Short
        Me.Tables = New Collection ' of clsTableHandle

        ConnectionItem = VB6.Replace(ConnectionItem, vbCrLf, Chr(254))
        ConnectionType = VB6.UCase(Extract(ConnectionItem, 1))

        If UserID = "" Then UserID = Extract(ConnectionItem, 3)
        If Password = "" Then
            Password = Extract(ConnectionItem, 4)
            If VB6.Right(ConnectionType, 1) = "!" Then
                Try
                    Password = AesDecrypt(XTS(Password))
                Catch ex As Exception
                End Try
            End If
        End If
        If VB6.Right(ConnectionType, 1) = "!" Then ConnectionType = VB6.Left(ConnectionType, Len(ConnectionType) - 1)
        ConnectionString = Extract(ConnectionItem, 2)

        Do
            I = VB6.InStr(ConnectionString, "<<")
            J = VB6.InStr(ConnectionString, ">>")
            If I Then I = I + 2

            ' If the connection string has parts encoded, decode them
            If I > 0 And J > 0 And I < J Then
                Dim substr = VB6.Mid(ConnectionString, I, J - I)
                Dim Tag As String = Field(substr, "|", 1)
                Dim DftValue As String = Field(substr, "|", 2)

                Select Case VB6.LCase(Tag)
                    Case "root"
                        DftValue = MapRootPath(".")

                    Case "rootaccount"
                        DftValue = MapRootAccount("")

                    Case "app_data"
                        DftValue = MapRootPath("App_Data")

                    Case "app_code"
                        DftValue = MapRootPath("App_Code")

                    Case "bin"
                        DftValue = MapRootPath("bin")

                    Case "_database", "database"
                        DftValue = MapRootPath("App_Data/_database")

                    Case "username", "userid"
                        If Len(UserID) Then DftValue = UserID

                    Case "password", "passwd"
                        If Len(Password) Then DftValue = Password

                End Select
                DftValue = VB6.Replace(DftValue, "<<", "")
                DftValue = VB6.Replace(DftValue, ">>", "")
                ConnectionString = VB6.Left(ConnectionString, I - 3) & DftValue & VB6.Mid(ConnectionString, J + 2)
            Else
                Exit Do
            End If
        Loop

        If TimeOutSecs = 0 Then TimeOutSecs = iDefaultTimeoutSeconds

        If ConnectionType = "D" Then
            ' A Data Provider String

        ElseIf ConnectionType = "C" Then
            ' C  : Configuration Section of Web Config

        ElseIf ConnectionType = "Jx" Then ' moved to upper level, not valid here anymore
            ' J : JSB Server http://some.domain/jsb/ServerLogin

        ElseIf ConnectionType = "Px" Then ' moved to upper level, not valid here anymore
            ' J : JSB Server http://some.domain/jsb/ServerLogin
            ConnectionString = ConnectionString & "serverlogin"
            ConnectionType = "J"

        ElseIf ConnectionType = "GS" Then
            ' GS : Google SpreadSheet

        ElseIf ConnectionType = "EX" Then
            ' EX  : Exchange Server

        ElseIf ConnectionType = "F" Then
            ' F  : File Folder Path
            ConnectionType = "F"

        Else
            ' Check for Disk Folder
            If AccountName = "" Or AccountName = "." Then AccountName = UrlAccount()
            ConnectionType = "F"

            If VB6.InStr(AccountName, "\") Or VB6.Mid(AccountName, 2, 1) = ":" Then
                If System.IO.Directory.Exists(AccountName) Then ConnectionString = Path.GetFullPath(AccountName)
                ' AccountName = GetFName(AccountName)

            Else
                Dim TestDir As String = MapRootAccount(AccountName)
                If System.IO.Directory.Exists(TestDir) Then ConnectionString = Path.GetFullPath(TestDir)
            End If
        End If

        If ConnectionType = "F" Then
            If ConnectionString = "" Then
                Errors = "Unknown account " & AccountName
                Return False
            End If
            Try
                If System.IO.Directory.Exists(ConnectionString) = False Then
                    If System.IO.Directory.Exists(MapPath(ConnectionString)) Then
                        ConnectionString = MapPath(ConnectionString)
                    Else
                        Return Nothing
                    End If
                End If
            Catch ex As Exception
                Errors = ex.Message
                Return False
            End Try

            'ElseIf ConnectionType = "EX" Then
            '   Try
            '      ExchangeServer = New ExchangeService
            '      ExchangeServer.TraceEnabled = True
            '      ExchangeServer.Url = New Uri(Extract(ConnectionString, 1)) ' "https://idfgpost.idfg.state.id.us/EWS/Exchange.asmx"

            '      If UserID Then
            '         ExchangeServer.UseDefaultCredentials = False
            '         ExchangeServer.Credentials = New WebCredentials(UserID, Password) ' "idfg\rwalsh" - Password
            '      End If

            '      Dim TestServerThere As Object = ExchangeServer.GetInboxRules
            '      If TestServerThere Is Nothing Then
            '         Errors = "Unable to locate Exchange Server: " & ExchangeServer.Url.ToString
            '         Return False
            '      End If

            '   Catch ex As Exception
            '      Errors = ex.Message
            '      Return False
            '   End Try

        ElseIf ConnectionType = "J" Then
            ' Login
            Dim JsbServer = ConnectionString
            Dim Result As String = "", rtnHeader As String = ""
            If VB6.InStr(JsbServer, "?") = 0 Then JsbServer &= "?username=" & UrlEncode(UserID) & "&password=" & UrlEncode(Password)

            If Not UrlFetch("GET", JsbServer, "", "", Result, rtnHeader, False, 33) Then
                Errors = "network error:" & Result
                Return False
            End If

            If VB6.Left(Result, 1) = "{" Then
                Dim JSonResult As Object = JSON(Result)
                If Val(JSonResult("_restFunctionResult")) = 0 Then
                    Errors = JSonResult("_p4")
                    Return False
                End If
                If JSonResult.ContainsKey("_rpcgid_") Then rpcgid = JSonResult("_rpcgid_") Else rpcgid = ""
                If JSonResult.ContainsKey("_rpcsid_") Then rpcsid = JSonResult("_rpcsid_") Else rpcsid = ""

            Else
                Errors = Result
                Return False
            End If
            ConnectionString = Field(LCase(ConnectionString), "/serverlogin", 1)

        ElseIf ConnectionType = "D" Then
            JetDB = VB6.InStr(LCase(ConnectionString), ".jet.")

            If VB6.InStr(AccountName, "\") Then
                sPath = GetFolder(AccountName)
                AccountName = GetFName(AccountName)

            ElseIf JetDB And VB6.InStr(LCase(ConnectionString), "data source=") Then
                I = VB6.InStr(LCase(ConnectionString), "data source=") + Len("Data Source=")
                sPath = VB6.Mid(ConnectionString, I)
                If VB6.InStr(sPath, ";") Then sPath = VB6.Left(sPath, VB6.InStr(sPath, ";") - 1)

                If VB6.Left(sPath, 1) = ".\" Then ' map relitive paths
                    Try
                        sPath = MapPath(VB6.Mid(sPath, 3))
                    Catch ex As Exception
                    End Try

                ElseIf VB6.Left(sPath, 1) <> "\" And VB6.Mid(sPath, 2, 1) <> ":" Then
                    Try
                        sPath = MapPath(sPath) '  or MapPath & 
                    Catch ex As Exception
                    End Try
                End If

                Dim RHI As String = VB6.InStr(I, ConnectionString, ";"), RHS As String = ""
                If RHI > 0 Then RHS = VB6.Mid(ConnectionString, VB6.InStr(I, ConnectionString, ";"))

                ConnectionString = VB6.Left(ConnectionString, I - 1) & sPath & RHS
                If AccountName = "" Then AccountName = GetFName(sPath)
            End If

            If VB6.UCase(VB6.Right(AccountName, 4)) = ".MDB" Then AccountName = VB6.Left(AccountName, Len(AccountName) - 4)

            Dim CN As String = ConnectionString
            CN = VB6.Replace(CN, "data source=", "Data Source=")
            CN = VB6.Replace(CN, "provider=", "Provider=")
            CN = VB6.Replace(CN, "persist security info=", "Persist Security Info=")
            CN = VB6.Replace(CN, "=true", "=True")
            CN = VB6.Replace(CN, "=false", "=False")
            CN = VB6.Replace(CN, "mode=", "Mode=")

            ' Check for upper case stuff-
            CN = VB6.Replace(CN, "DATA SOURCE=", "Data Source=")
            CN = VB6.Replace(CN, "PROVIDER=", "Provider=")
            CN = VB6.Replace(CN, "PERSIST SECURITY INO=", "Persist Security Info=")
            CN = VB6.Replace(CN, "=TRUE", "=True")
            CN = VB6.Replace(CN, "=FALSE", "=False")
            CN = VB6.Replace(CN, "MODE=", "Mode=")

            CN = VB6.Replace(CN, "Provider=SQLOLEDB.1;", "")
            CN = VB6.Replace(CN, "Provider=SQLNCLI.1;", "")
            CN = VB6.Replace(CN, "Provider=SQLOLEDB;", "")
            CN = VB6.Replace(CN, "Provider=SQLNCLI;", "")

            Dim IsODBC As Boolean = False, IsAccess As Boolean = False, IsSql As Boolean = False

            If VB6.InStr(CN, "MSDASQL.") Then IsODBC = 1 Else If VB6.InStr(CN, "Provider") Then IsAccess = True Else IsSql = True


            If UserID <> "" Then
                ' Do we need to add a UserName and Password? User ID=sa;Password=pwd; or UID=sa;PWD=passwd
                I = VB6.InStr(LCase(CN), "user id=")

                If I Then
                    I = I + Len("user id=")
                    J = VB6.InStr(I, CN, ";")
                    If J = 0 Then J = Len(CN) + 1
                    CN = VB6.Left(CN, I - 1) & UserID & VB6.Mid(CN, J)
                Else
                    CN = CN & ";User ID=" & UserID
                End If
            End If

            If Password <> "" Then
                ' Do we need to add a UserName and Password? User ID=sa;Password=pwd; or UID=sa;PWD=passwd
                I = VB6.InStr(LCase(CN), "password=")

                If I Then
                    I = I + Len("password=")
                    J = VB6.InStr(I, CN, ";")
                    If J = 0 Then J = Len(CN) + 1
                    CN = VB6.Left(CN, I - 1) & Password & VB6.Mid(CN, J)
                Else
                    CN = CN & ";Password=" & Password
                End If
            End If

            If IsAccess Then
                CN = VB6.Replace(CN, ";Password=", ";Jet OLEDB:Database Password=")
                CN = VB6.Replace(CN, "Mode=ReadWrite|Share Deny None", "mode=19")
                If VB6.InStr(LCase(CN), "mode=") = 0 Then CN &= ";mode=19"
            End If

            If IsODBC Then
                CN = VB6.Replace(CN, ";Password=", ";Pwd=")
                CN = VB6.Replace(CN, ";Persist Security Info=True", "")
                CN = VB6.Replace(CN, ";Persist Security Info=False", "")
                CN = VB6.Replace(CN, ";User ID=", ";UID=")
            End If

            If IsSql Then
                If VB6.InStr(LCase(CN), "connect timeout=") = 0 Then CN = CN & "; Connect Timeout=2"
            End If

            Try

                mAdoConnection = NewConnection(CN)
                mAdoConnection.Open()

            Catch ex1 As Exception
                Errors = ex1.Message
                mAdoConnection = Nothing
                Return False
            End Try

            ' Success!
            Me.ConnectionString = CN
            mAdoConnection.Close()
        End If

        Me.AcountName = AccountName
        Tables = New Collection
        Return True
    End Function

    Protected Overrides Sub Finalize()
        Tables = Nothing
        mAdoConnection = Nothing
    End Sub
End Class

Public MustInherit Class clsTableHandle
    Public ReadUCache As Collection
    Public MyDBAccount As clsjdbAccount
    Public OneChunkRecord As Boolean ' Has data stored in a field called "ItemContent"
    Public MemoColumnName As String = "ItemContent"
    Public LastItemOut As String = ""
    Public LastHeader As String = ""

    Public MustOverride ReadOnly Property TableName() As String
    Public MustOverride Function dbReadBool(ByVal ItemID As String, ByRef Result As String, Optional ByVal IsReadU As Boolean = False) As Boolean
    Public MustOverride Function dbReadJSon(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
    Public MustOverride Function dbReadXML(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
    Public MustOverride Sub dbWrite(ByVal Item As String, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
    Public MustOverride Sub dbWriteJSon(ByVal Item As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
    Public MustOverride Sub dbWriteXML(ByVal Item As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
    Public MustOverride Function SelectFile() As rSelectList
    Public MustOverride Function SelectFileX(ByVal ColumnList As String, ByVal WhereClause As String) As rSelectList
    Public MustOverride Sub ClearFile()
    Public MustOverride Sub dbTransactionBegin()
    Public MustOverride Sub dbTransactionCommit()
    Public MustOverride Sub dbExitServer()

    Public MustOverride Sub dbDeleteFile()
    Public MustOverride Sub dbDelete(ByVal ItemID As String)
    Public MustOverride ReadOnly Property ColumnNames() As Collection
    Public MustOverride ReadOnly Property PrimaryKeyColumnName() As String
    Public MustOverride Function dbOpenBool(ByVal dbAccount As clsjdbAccount, ByVal TableName As String, ByRef Errors As String, Optional ByVal CreateIt As Boolean = False) As Boolean
    Public MustOverride ReadOnly Property ColumnPosition(ByVal ColumnName As String) As Integer


End Class

Public Class clsTableHandleAdo
    Inherits clsTableHandle

    Dim mSqlActiveConnection As IDbConnection = Nothing
    Dim mPrimaryKeyColumnName As String
    Dim mColumnNames As Collection ' of Strings - No "[" "]"'s
    Dim mColumnNamesString As String ' Seperated by commas with [ ]'s
    Dim PKQuote As String

    Dim mTableName As String

    Public mTransaction As IDbTransaction = Nothing

    Public Overrides Sub dbTransactionCommit()
        Dim UsrTrans As IDbTransaction = mTransaction
        mTransaction = Nothing
        If UsrTrans Is Nothing Then Return
        UsrTrans.Commit()
    End Sub

    ' Start User transaction or error out
    Public Overrides Sub dbTransactionBegin()
        If SqlActiveConnection.State = ConnectionState.Closed Then SqlActiveConnection.Open()
        mTransaction = SqlActiveConnection.BeginTransaction()
    End Sub

    Public Overrides Sub dbExitServer()

    End Sub


    Public Sub dbWriteX(ByVal Item As Object, ByVal ItemID As String, ByVal isJSon As Boolean, ByVal XML As Boolean, ByVal IsWriteU As Boolean)
        Dim ReRead As String = "", WasInCache As Boolean = False

        ' Validate parameters
        If SqlActiveConnection.State = ConnectionState.Closed Then SqlActiveConnection.Open()
        If IsDBNull(ItemID) OrElse ItemID = "" Then Throw New Exception("clsTableHandleAdo-" & mTableName & "-dbWriteX: Bad Primary Key")

        ' Did we pass optimistic Locking?
        Dim CacheName As String = ItemID
        If isJSon Then CacheName &= Chr(254) & "_JSON"
        If XML Then CacheName &= Chr(254) & "XML"

        If ReadUCache.Contains(CacheName) Then
            Dim CachedItem As String = ReadUCache(CacheName)
            If dbReadX(ItemID, ReRead, False, isJSon, XML) Then
                If ReRead <> CachedItem Then
                    If mTransaction IsNot Nothing Then
                        mTransaction.Rollback()
                        mTransaction = Nothing
                        Throw New Exception("Optimistic Locking Failed - Item has been modified by another process. Transaction aborted")
                    End If
                    Throw New Exception("Optimistic Locking Failed - Item has been modified by another process.")
                End If
            Else
                If mTransaction IsNot Nothing Then
                    mTransaction.Rollback()
                    mTransaction = Nothing
                    Throw New Exception("Optimistic Locking Failed - Item has been deleted by another process. Transaction aborted")
                End If
                Throw New Exception("Optimistic Locking Failed - Item has been deleted by another process")
            End If
            WasInCache = True
        End If

        ' Build list of column names
        Dim ColumnNames As New ArrayList
        If isJSon Then
            ' Names are in record
            If TypeOf Item Is String Then Item = JSON(Item)
            For Each C As KeyValuePair(Of String, Object) In Item
                If mColumnNames.Contains(C.Key) Then ColumnNames.Add(C.Key)
            Next
            If ColumnNames.Contains(mPrimaryKeyColumnName) = False Then ColumnNames.Add(mPrimaryKeyColumnName)

        ElseIf XML Then
            ' Names are in record
            If TypeOf Item Is String Then Item = XML_Str2Obj(Item)
            For Each C As System.Xml.XmlElement In Item.DocumentElement
                If mColumnNames.Contains(C.Name) Then ColumnNames.Add(C.Name)
            Next
            If ColumnNames.Contains(mPrimaryKeyColumnName) = False Then ColumnNames.Add(mPrimaryKeyColumnName)

        ElseIf OneChunkRecord Then
            ColumnNames.Add(mPrimaryKeyColumnName)
            ColumnNames.Add(MemoColumnName)

        Else
            ColumnNames.Add(mColumnNamesString)
        End If

        Dim SqlSelect As String = "SELECT [" & VB6.Join(ColumnNames.ToArray, "],[") & "] FROM [" & mTableName & "] WHERE [" & mPrimaryKeyColumnName & "] = " & PKQuote & VB6.Replace(ItemID, "'", "''") & PKQuote

        ' Read Item from database for data adapter update
        Dim DA As IDbDataAdapter = NewDataAdapter(SqlSelect, SqlActiveConnection), ds As New DataSet
        SetupSqlCommands(DA, mTableName, mPrimaryKeyColumnName, , mTransaction)
        Try
            CType(DA, System.Data.Common.DbDataAdapter).Fill(ds)
        Catch ex As Exception
            If PKQuote = "" AndAlso IsNumeric(ItemID) = False Then
                Throw New Exception("SQL Column Type mis-match on column [" & mPrimaryKeyColumnName & "] with value " & ItemID)
            End If
            Throw ex
        End Try

        Dim dt As DataTable = ds.Tables(0)
        Dim RowHandle As DataRow
        Dim isNew As Boolean = dt.Rows.Count = 0
        If isNew Then
            RowHandle = dt.NewRow
            RowHandle(mPrimaryKeyColumnName) = ItemID
            dt.Rows.Add(RowHandle)
        Else
            RowHandle = dt.Rows(0)
        End If

        If isJSon Then
            For Each C As KeyValuePair(Of String, Object) In Item
                If mColumnNames.Contains(C.Key) Then
                    Try
                        If C.Key <> mPrimaryKeyColumnName Then
                            If dt.Columns(C.Key).DataType Is GetType(System.Boolean) Then

                                If IsDBNull(RowHandle(C.Key)) Then
                                    If CStr(C.Value) <> "" Then RowHandle(C.Key) = C2Bool(C.Value)

                                ElseIf CStr(C.Value) = "" Then
                                    If RowHandle(C.Key) <> False Then RowHandle(C.Key) = False

                                ElseIf RowHandle(C.Key) <> C2Bool(C.Value) Then
                                    RowHandle(C.Key) = C2Bool(C.Value)

                                End If

                            ElseIf dt.Columns(C.Key).DataType Is GetType(System.Guid) Then
                                Dim V As New Guid()

                                If CStr(C.Value) <> "" Then V = Guid.Parse(C.Value)

                                If IsDBNull(RowHandle(C.Key)) Then
                                    If CStr(C.Value) <> "" Then RowHandle(C.Key) = V

                                ElseIf CStr(C.Value) = "" Then
                                    If RowHandle(C.Key).ToString <> "" Then RowHandle(C.Key) = V

                                ElseIf RowHandle(C.Key).ToString <> C.Value Then
                                    RowHandle(C.Key) = V

                                End If

                            ElseIf DateFld(dt.Columns(C.Key).DataType) Then
                                If IsDBNull(RowHandle(C.Key)) Then
                                    If CStr(C.Value) <> "" Then RowHandle(C.Key) = CDate(C.Value)

                                ElseIf CStr(C.Value) = "" Then
                                    RowHandle(C.Key) = DBNull.Value

                                ElseIf RowHandle(C.Key) <> CDate(C.Value) Then
                                    RowHandle(C.Key) = CDate(C.Value)
                                End If

                            ElseIf NumericFld(dt.Columns(C.Key).DataType) Then
                                If IsDBNull(RowHandle(C.Key)) Then
                                    If CStr(C.Value) <> "" Then RowHandle(C.Key) = CNum(C.Value)
                                ElseIf RowHandle(C.Key) <> CNum(C.Value) Then
                                    RowHandle(C.Key) = CNum(C.Value)
                                End If

                            Else
                                RowHandle(C.Key) = C.Value
                            End If
                        End If
                    Catch ex As Exception
                        Throw New Exception("Table: " & mTableName & "; Column: " & C.Key & "; Error: " & ex.Message)
                    End Try
                End If
            Next
        ElseIf XML Then
            For Each C As System.Xml.XmlElement In Item.DocumentElement
                If mColumnNames.Contains(C.Name) Then
                    Try
                        If C.Name <> mPrimaryKeyColumnName Then
                            If dt.Columns(C.Name).DataType Is GetType(System.Boolean) Then
                                If IsDBNull(RowHandle(C.Name)) Then
                                    If C.InnerText <> "" Then RowHandle(C.Name) = C2Bool(C.InnerText)

                                ElseIf C.InnerText = "" Then
                                    If RowHandle(C.Name) <> False Then RowHandle(C.Name) = False

                                ElseIf RowHandle(C.Name) <> C2Bool(C.InnerText) Then
                                    RowHandle(C.Name) = C2Bool(C.InnerText)

                                End If

                            ElseIf DateFld(dt.Columns(C.Name).DataType) Then
                                If IsDBNull(RowHandle(C.Name)) Then
                                    If C.InnerText <> "" Then RowHandle(C.Name) = CDate(C.InnerText)

                                ElseIf C.InnerText = "" Then
                                    RowHandle(C.Name) = DBNull.Value

                                ElseIf RowHandle(C.Name) <> CDate(C.InnerText) Then
                                    RowHandle(C.Name) = CDate(C.InnerText)
                                End If

                            ElseIf NumericFld(dt.Columns(C.Name).DataType) Then
                                If IsDBNull(RowHandle(C.Name)) Then
                                    If C.InnerText <> "" Then RowHandle(C.Name) = CNum(C.InnerText)
                                ElseIf RowHandle(C.Name) <> CNum(C.InnerText) Then
                                    RowHandle(C.Name) = CNum(C.InnerText)
                                End If

                            Else
                                RowHandle(C.Name) = C.InnerText
                            End If
                        End If
                    Catch ex As Exception
                        Throw New Exception("Table: " & mTableName & "; Column: " & C.Name & "; Error: " & ex.Message)
                    End Try
                End If
            Next
        ElseIf OneChunkRecord Then
            ' Get data from recordoDB As clsDBConnection,set
            RowHandle(MemoColumnName) = Item
        Else
            ' Write Each <atrNo> to mapped fields, 1 by 1
            Dim FastWrite() As String = VB6.Split(Item, Chr(254))
            ReDim Preserve FastWrite(dt.Columns.Count)
            Dim J As Short = 0
            For Each Column In dt.Columns

                Try
                    If Column.ColumnName <> mPrimaryKeyColumnName Then

                        If dt.Columns(Column.ColumnName).DataType Is GetType(System.Boolean) Then
                            If IsDBNull(RowHandle(Column.ColumnName)) Then
                                If FastWrite(J) <> "" Then RowHandle(Column.ColumnName) = C2Bool(FastWrite(J))

                            ElseIf FastWrite(J) = "" Then
                                If RowHandle(Column.ColumnName) <> False Then RowHandle(Column.ColumnName) = False

                            ElseIf RowHandle(Column.ColumnName) <> C2Bool(FastWrite(J)) Then
                                RowHandle(Column.ColumnName) = C2Bool(FastWrite(J))

                            End If

                        ElseIf DateFld(dt.Columns(Column.ColumnName).DataType) Then
                            If IsDBNull(RowHandle(Column.ColumnName)) Then
                                If FastWrite(J) <> "" Then RowHandle(Column.ColumnName) = CDate(FastWrite(J))

                            ElseIf FastWrite(J) = "" Then
                                RowHandle(Column.ColumnName) = DBNull.Value

                            ElseIf RowHandle(Column.ColumnName) <> CDate(FastWrite(J)) Then
                                RowHandle(Column.ColumnName) = CDate(FastWrite(J))
                            End If

                        ElseIf NumericFld(dt.Columns(Column.ColumnName).DataType) Then
                            If IsDBNull(RowHandle(Column.ColumnName)) Then
                                If FastWrite(J) <> "" Then RowHandle(Column.ColumnName) = CNum(FastWrite(J))
                            ElseIf RowHandle(Column.ColumnName) <> CNum(FastWrite(J)) Then
                                RowHandle(Column.ColumnName) = CNum(FastWrite(J))
                            End If

                        Else
                            RowHandle(Column.ColumnName) = FastWrite(J)
                        End If
                    End If

                    J = J + 1
                Catch ex As Exception
                    Throw New Exception("Table: " & mTableName & "; Column: " & Column.ColumnName & "; Error: " & ex.Message)
                End Try
            Next
        End If

        ' Update
        Try
            DA.Update(ds)

        Catch ex As Exception
            '  mTransaction.Rollback()
            '  mTransaction = Nothing
            Throw New Exception("Table: " & mTableName & "; Error: " & ex.Message)
        End Try

        If mTransaction Is Nothing Then SqlActiveConnection.Close()

        ' Update cache as necessary
        If IsWriteU Then dbReadX(ItemID, ReRead, True, isJSon, XML)
    End Sub

    Public Overrides ReadOnly Property TableName() As String
        Get
            TableName = mTableName
        End Get
    End Property

    Public ReadOnly Property ColumnNamesString() As String
        Get
            ColumnNamesString = mColumnNamesString
        End Get
    End Property

    Public Overrides ReadOnly Property ColumnNames() As Collection
        Get
            ColumnNames = mColumnNames
        End Get
    End Property

    Public Overrides ReadOnly Property ColumnPosition(ByVal ColumnName As String) As Integer
        Get
            ColumnName = VB6.LCase(ColumnName)
            For i As Integer = 1 To mColumnNames.Count
                If VB6.LCase(mColumnNames(i)) = ColumnName Then Return i
            Next
            Return 0
        End Get
    End Property

    Public ReadOnly Property SqlActiveConnection() As IDbConnection
        Get
            SqlActiveConnection = mSqlActiveConnection
        End Get
    End Property

    Public Overrides ReadOnly Property PrimaryKeyColumnName() As String
        Get
            PrimaryKeyColumnName = mPrimaryKeyColumnName
        End Get
    End Property

    Public Overrides Function dbOpenBool(ByVal dbAccount As clsjdbAccount, ByVal TableName As String, ByRef Errors As String, Optional ByVal CreateIt As Boolean = False) As Boolean

        Dim RS As clsResultSet
        Dim Fld As DataColumn

        mTableName = TableName
        mPrimaryKeyColumnName = "ItemID"

        If TableName = "" Then
            Errors = "OpenBool(No TableName)"
            Return False
        End If

        If dbAccount.AdoConnection Is Nothing Then
            Errors = "OpenBool(No AdoConnection)"
            Return False
        End If

        mSqlActiveConnection = dbAccount.AdoConnection
        ReadUCache = New Collection

        ' Open file
        RS = New clsResultSet
        If SqlActiveConnection.State = ConnectionState.Closed Then SqlActiveConnection.Open()

        Try
            RS.Open("SELECT * FROM [" & mTableName & "] WHERE 1=0", SqlActiveConnection)

        Catch ex As Exception
            Errors = ex.Message
            If CreateIt = False Then Return False ' Failed to open

            Try
                Dim CMD As IDbCommand = NewCommand("Create Table [" & TableName & "]([" & mPrimaryKeyColumnName & "] nvarchar(255) NOT NULL, [" & MemoColumnName & "] [text])", SqlActiveConnection, mTransaction)
                CMD.ExecuteNonQuery()

                Try
                    CMD = NewCommand("CREATE INDEX Idx_" & mPrimaryKeyColumnName & " ON [" & mTableName & "] ([" & mPrimaryKeyColumnName & "])", SqlActiveConnection, mTransaction)
                    CMD.ExecuteNonQuery()
                Catch exIgnore As Exception
                End Try

                PKQuote = "'"
                RS.Open("SELECT * FROM [" & mTableName & "] WHERE 1=0", SqlActiveConnection)



            Catch ex2 As Exception
                Errors = ex2.Message
                Return False
            End Try

        End Try


        If RS.Columns.Count = 0 Then Return False

        ' You get one, or the other, but not both
        mColumnNames = New Collection
        If RS.Columns.Contains(mPrimaryKeyColumnName) AndAlso RS.Columns.Contains(MemoColumnName) Then
            OneChunkRecord = True
            PKQuote = "'"
        Else
            OneChunkRecord = False

            ' Build list of ColumnName in Ordinal Order (Excluding 1st col, which is the PK Name
            Dim FirstCol As Boolean = True
            For Each Fld In RS.Columns
                If FirstCol Then
                    mPrimaryKeyColumnName = Fld.ColumnName
                    If NumericFld(Fld.DataType) Then
                        PKQuote = ""
                    ElseIf DateFld(Fld.DataType) Then
                        PKQuote = SqlDateDelimiter(SqlActiveConnection)
                    Else
                        PKQuote = "'"
                    End If
                End If

                mColumnNames.Add(Fld.ColumnName, Fld.ColumnName)

                FirstCol = False
            Next Fld
        End If

        ' Build "Select" list of column names
        mColumnNamesString = ""
        For Each ColumnName As String In mColumnNames
            If mColumnNamesString.Length Then mColumnNamesString &= ","
            mColumnNamesString &= "[" & ColumnName & "]"
        Next

        If mTransaction Is Nothing Then SqlActiveConnection.Close()
        MyDBAccount = dbAccount
        RS = Nothing
        dbOpenBool = True
    End Function

    Private Function NumericFld(ByVal DataType As System.Type) As Boolean
        Return DataType Is GetType(System.Single) _
         OrElse DataType Is GetType(System.Double) _
         OrElse DataType Is GetType(System.Decimal) _
         OrElse DataType Is GetType(System.Byte) _
         OrElse DataType Is GetType(System.Int16) _
         OrElse DataType Is GetType(System.Int32) _
         OrElse DataType Is GetType(System.Int64) _
         OrElse DataType Is GetType(System.UInt16) _
         OrElse DataType Is GetType(System.UInt32) _
         OrElse DataType Is GetType(System.UInt64)
    End Function

    Private Function DateFld(ByVal DataType As System.Type) As Boolean
        Return DataType Is GetType(System.DateTime) OrElse DataType Is GetType(System.TimeSpan)
    End Function

    Public Function ColumnExists(ByVal pColumnName As String) As Boolean
        Return mColumnNames.Contains(pColumnName)
    End Function

    Public Function FirstColumn() As String
        On Error Resume Next
        FirstColumn = mColumnNames(1)
    End Function

    Public Function GetSchema() As Object
        ' Validate parameters
        If SqlActiveConnection Is Nothing Then Throw New Exception("Ado Table handle not valid")
        If SqlActiveConnection.State = ConnectionState.Closed Then SqlActiveConnection.Open()

        Dim Schema As DataTable
        Dim Restrictions(2) As String
        Restrictions(2) = mTableName

        If TypeOf SqlActiveConnection Is OleDb.OleDbConnection Then
            Schema = CType(SqlActiveConnection, OleDb.OleDbConnection).GetSchema("Columns", Restrictions)
        ElseIf TypeOf SqlActiveConnection Is SQLite.SQLiteConnection Then
            Schema = CType(SqlActiveConnection, SQLite.SQLiteConnection).GetSchema("Columns", Restrictions)
        ElseIf TypeOf SqlActiveConnection Is SqlClient.SqlConnection Then
            Schema = CType(SqlActiveConnection, SqlClient.SqlConnection).GetSchema("Columns", Restrictions)
        ElseIf TypeOf SqlActiveConnection Is Odbc.OdbcConnection Then
            Schema = CType(SqlActiveConnection, Odbc.OdbcConnection).GetSchema("Columns", Restrictions)
        Else
            Throw New Exception("Not logged in!")
        End If

        If mTransaction Is Nothing Then SqlActiveConnection.Close()

        Dim SB As New StringBuilder, FirstRow As Boolean = True
        SB.AppendLine("[")
        For Each row As DataRow In Schema.Rows
            If Not FirstRow Then SB.AppendLine(",")
            SB.Append("{")
            Dim FirstCol As Boolean = True
            For Each col As DataColumn In Schema.Columns
                If Not FirstCol Then SB.AppendLine(",")
                SB.Append("""" & col.ColumnName & """:")
                SB.Append("""" & row(col) & """")
                FirstCol = False
            Next
            SB.AppendLine("}")
            FirstRow = False
        Next
        SB.AppendLine("]")
        Dim S As String = SB.ToString
        Return JSON(S)
    End Function

    Public Overrides Function dbReadBool(ByVal ItemID As String, ByRef Result As String, Optional ByVal IsReadU As Boolean = False) As Boolean
        Return dbReadX(ItemID, Result, IsReadU, False, False)
    End Function

    Public Overrides Function dbReadJSon(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
        If Not dbReadX(ItemID, Result, IsReadU, True, False) Then Return False
        Result = JSON(Result)
        Return True
    End Function

    Public Overrides Function dbReadXML(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
        If Not dbReadX(ItemID, Result, IsReadU, False, True) Then Return False
        Result = XML_Str2Obj(Result)
        Return True
    End Function

    Private Function dbReadX(ByVal ItemID As String, ByRef Result As Object, ByVal IsReadU As Boolean, ByVal JSon As Boolean, ByVal XML As Boolean) As Boolean
        Dim SqlSelect As String = ""

        ' Validate parameters
        If SqlActiveConnection Is Nothing Then Throw New Exception("Ado Table handle not valid")
        If IsDBNull(ItemID) OrElse ItemID = "" Then Return False
        If SqlActiveConnection.State = ConnectionState.Closed Then SqlActiveConnection.Open()

        If XML Or JSon Then
            SqlSelect = "SELECT " & mColumnNamesString & " FROM [" & mTableName & "] WHERE [" & mPrimaryKeyColumnName & "] = " & PKQuote & VB6.Replace(ItemID, "'", "''") & PKQuote
        ElseIf OneChunkRecord Then
            SqlSelect = "SELECT [" & mPrimaryKeyColumnName & "], " & MemoColumnName & " FROM [" & mTableName & "] WHERE [" & mPrimaryKeyColumnName & "] = " & PKQuote & VB6.Replace(ItemID, "'", "''") & PKQuote
        Else
            SqlSelect = "SELECT " & mColumnNamesString & " FROM [" & mTableName & "] WHERE [" & mPrimaryKeyColumnName & "] = " & PKQuote & VB6.Replace(ItemID, "'", "''") & PKQuote
        End If

        Dim AccessReader As IDataReader = Nothing
        Try
            '   RowHandle.Open(SqlSelect, SqlActiveConnection)
            Dim AccessCommand As IDbCommand = NewCommand(SqlSelect, SqlActiveConnection, mTransaction)
            AccessReader = AccessCommand.ExecuteReader()

        Catch ex As Exception
            If AccessReader IsNot Nothing AndAlso Not AccessReader.IsClosed Then AccessReader.Close()
            Return False
        End Try

        If AccessReader.Read = False Then
            If Not AccessReader.IsClosed Then AccessReader.Close()
            If mTransaction Is Nothing Then SqlActiveConnection.Close()
            Return False
        End If

        If JSon Then
            ' Get data from recordoDB As clsDBConnection,set
            Dim Results As String = ""
            For FldI As Integer = 0 To AccessReader.FieldCount - 1
                Results &= VB6.Left(",", Len(Results)) & """" & AccessReader.GetName(FldI) & """:"

                If IsDBNull(AccessReader.Item(FldI)) Then
                    If AccessReader.GetFieldType(FldI) Is GetType(System.Boolean) Then Results &= "false" Else Results &= """"""

                Else
                    If AccessReader.GetFieldType(FldI) Is GetType(System.Boolean) Then
                        If AccessReader.Item(FldI) Then Results &= "true" Else Results &= "false"

                    ElseIf AccessReader.GetFieldType(FldI) Is GetType(System.Byte()) Then
                        Dim byteArray As Byte() = AccessReader.Item(FldI)
                        If VB6.UBound(byteArray) = 7 Then
                            If BitConverter.IsLittleEndian Then Array.Reverse(byteArray)
                            Dim longValue As Long = BitConverter.ToInt64(byteArray, 0)
                            Results &= CStr(longValue)
                        Else
                            Results &= JSonEncodeString(System.Text.Encoding.Default.GetString(AccessReader.Item(FldI)))
                        End If
                    Else
                        Results &= JSonEncodeString(AccessReader.Item(FldI).ToString)
                    End If
                End If
            Next FldI
            Result = "{" & Results & "}"


        ElseIf XML Then
            ' Get data from recordoDB As clsDBConnection,set
            Result = "<record>"
            For FldI As Integer = 0 To AccessReader.FieldCount - 1
                Result &= "<" & AccessReader.GetName(FldI) & ">"

                If IsDBNull(AccessReader.Item(FldI)) Then
                    If AccessReader.GetFieldType(FldI) Is GetType(System.Boolean) Then Result = Result & "false" Else Result = Result & """"""
                Else
                    If AccessReader.GetFieldType(FldI) Is GetType(System.Boolean) Then
                        If AccessReader.Item(FldI) Then Result &= "true" Else Result &= "false"

                    ElseIf AccessReader.GetFieldType(FldI) Is GetType(System.Byte()) Then
                        Dim byteArray As Byte() = AccessReader.Item(FldI)
                        If VB6.UBound(byteArray) = 7 Then
                            If BitConverter.IsLittleEndian Then Array.Reverse(byteArray)
                            Dim longValue As Long = BitConverter.ToInt64(byteArray, 0)
                            Result &= CStr(longValue)
                        Else
                            Result &= JSonEncodeString(System.Text.Encoding.Default.GetString(AccessReader.Item(FldI)))
                        End If
                    Else
                        Result &= XMLEncodeString(AccessReader.Item(FldI).ToString)
                    End If
                End If

                Result &= "</" & AccessReader.GetName(FldI) & ">"
            Next FldI
            Result &= "</record>"

        Else
            If OneChunkRecord Then
                ' Get data from recordoDB As clsDBConnection,set
                Result = AccessReader.Item(MemoColumnName)
            Else
                ' Get data from recordoDB As clsDBConnection,set
                Dim iResults As String = ""
                For FldI As Integer = 0 To AccessReader.FieldCount - 1
                    If IsDBNull(AccessReader.Item(FldI)) Then
                        If AccessReader.GetFieldType(FldI) Is GetType(System.Boolean) Then iResults &= Chr(254) & "false" Else iResults &= Chr(254) & ""
                    Else
                        If AccessReader.GetFieldType(FldI) Is GetType(System.Boolean) Then
                            If AccessReader.Item(FldI) Then iResults &= Chr(254) & "true" Else iResults &= Chr(254) & "false"

                        ElseIf AccessReader.GetFieldType(FldI) Is GetType(System.Byte()) Then
                            Dim byteArray As Byte() = AccessReader.Item(FldI)
                            If VB6.UBound(byteArray) = 7 Then
                                If BitConverter.IsLittleEndian Then Array.Reverse(byteArray)
                                Dim longValue As Long = BitConverter.ToInt64(byteArray, 0)
                                iResults &= CStr(longValue)
                            Else
                                iResults &= JSonEncodeString(System.Text.Encoding.Default.GetString(AccessReader.Item(FldI)))
                            End If
                        Else
                            iResults &= Chr(254) & AccessReader.Item(FldI).ToString
                        End If
                    End If
                Next FldI
                Result = VB6.Mid(iResults, 2)
            End If
        End If

        AccessReader.Close()
        If mTransaction Is Nothing Then SqlActiveConnection.Close()

        ' We do this for optimistic locking
        If JSon Then ItemID &= Chr(254) & "_JSON"
        If XML Then ItemID &= Chr(254) & "XML"
        If ReadUCache.Contains(ItemID) Then ReadUCache.Remove(ItemID)
        If IsReadU Then ReadUCache.Add(Result, ItemID)

        Return True
    End Function

    Public Overrides Sub dbWrite(ByVal Item As String, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        Call dbWriteX(Item, ItemID, False, False, IsWriteU)
    End Sub

    Public Overrides Sub dbWriteJson(ByVal Item As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        Call dbWriteX(Item, ItemID, True, False, IsWriteU)
    End Sub

    Public Overrides Sub dbWriteXML(ByVal Item As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        Call dbWriteX(Item, ItemID, False, True, IsWriteU)
    End Sub

    Protected Overrides Sub Finalize()
        mSqlActiveConnection = Nothing
        MyBase.Finalize()
    End Sub

    Public Overrides Function SelectFile() As rSelectList
        Dim SelectHandle As New rSelectList
        If SqlActiveConnection Is Nothing Then Throw New Exception("Ado Table handle not valid")
        SelectHandle.PrimeSqlSelect(Me, "SELECT [" & mPrimaryKeyColumnName & "] FROM [" & mTableName & "]", True)
        Return SelectHandle
    End Function

    Public Overrides Function SelectFileX(ByVal ColumnList As String, ByVal WhereClause As String) As rSelectList
        Dim SelectHandle As New rSelectList, OnlyReturnItemIDs As Boolean
        Dim Limit As String = ""
        If SqlActiveConnection Is Nothing Then Throw New Exception("Ado Table handle not valid")
        Dim IsSqlLite As Boolean = TypeOf SqlActiveConnection Is SQLite.SQLiteConnection

        Dim WantItemID As Boolean = False, WantItemContent As Boolean = False, hasStarColumn As Boolean = False, AttributeNumsReferenced As New Collection

        WhereClause = VB6.LTrim(WhereClause)
        If VB6.Left(VB6.LCase(WhereClause), 6) = "where " Then WhereClause = VB6.Mid(WhereClause, 7)
        If VB6.InStr(WhereClause, "'") = 0 Then WhereClause = VB6.Replace(WhereClause, """", "'")
        If VB6.InStr(VB6.LCase(WhereClause), " like ") Then
            WhereClause = VB6.Replace(WhereClause, "'[", "'%")
            WhereClause = VB6.Replace(WhereClause, "]'", "%'")
        End If

        activeColumns(ColumnList & " " & WhereClause, False, WantItemID, WantItemContent, hasStarColumn, AttributeNumsReferenced)


        ' Accept ItemID anywhere as the primarykey
        If mPrimaryKeyColumnName <> "ItemID" And WantItemID Then
            ColumnList = " " & ColumnList & " "
            ColumnList = VB6.Replace(ColumnList, " ItemID ", "[" & mPrimaryKeyColumnName & "]", 1, -1, CompareMethod.Text)
            ColumnList = VB6.Replace(ColumnList, " ItemID,", "[" & mPrimaryKeyColumnName & "],", 1, -1, CompareMethod.Text)
            ColumnList = VB6.Replace(ColumnList, ",ItemID ", ",[" & mPrimaryKeyColumnName & "]", 1, -1, CompareMethod.Text)
            ColumnList = VB6.Replace(ColumnList, ",ItemID,", ",[" & mPrimaryKeyColumnName & "],", 1, -1, CompareMethod.Text)
            ColumnList = VB6.Replace(ColumnList, "[ItemID]", "[" & mPrimaryKeyColumnName & "]", 1, -1, CompareMethod.Text)
            ColumnList = VB6.LTrim(RTrim(ColumnList))

            WhereClause = " " & WhereClause & " "
            WhereClause = VB6.Replace(WhereClause, " ItemID ", "[" & mPrimaryKeyColumnName & "]", 1, -1, CompareMethod.Text)
            WhereClause = VB6.Replace(WhereClause, " ItemID,", "[" & mPrimaryKeyColumnName & "],", 1, -1, CompareMethod.Text)
            WhereClause = VB6.Replace(WhereClause, ",ItemID ", ",[" & mPrimaryKeyColumnName & "]", 1, -1, CompareMethod.Text)
            WhereClause = VB6.Replace(WhereClause, ",ItemID,", ",[" & mPrimaryKeyColumnName & "],", 1, -1, CompareMethod.Text)
            WhereClause = VB6.Replace(WhereClause, "[ItemID]", "[" & mPrimaryKeyColumnName & "]", 1, -1, CompareMethod.Text)
            WhereClause = VB6.LTrim(RTrim(WhereClause))
        End If

        If WantItemContent Then
            If mColumnNames.Contains("ItemContent") = False Then
                ColumnList = " " & ColumnList & " "
                ColumnList = VB6.Replace(ColumnList, " ItemContent ", " ", 1, -1, CompareMethod.Text)
                ColumnList = VB6.Replace(ColumnList, " ItemContent,", " ", 1, -1, CompareMethod.Text)
                ColumnList = VB6.Replace(ColumnList, ",ItemContent ", ",", 1, -1, CompareMethod.Text)
                ColumnList = VB6.Replace(ColumnList, ",ItemContent,", ",", 1, -1, CompareMethod.Text)
                ColumnList = VB6.LTrim(RTrim(ColumnList))
                If ColumnList <> "" Then ColumnList = "*," & ColumnList Else ColumnList = "*"
            End If
        End If

        If AttributeNumsReferenced.Count Then
            ColumnList = "*"
        Else
            ' Drop ItemID, ItemContent and *An
            Dim CA As ArrayList = activeColumns(ColumnList, False, False, False, False, New Collection)
            If CA.Count > 1 Then
                ColumnList = "[" & VB6.Join(CA.ToArray, "],[") & "]"
            Else
                ColumnList = VB6.Join(CA.ToArray, ",")
            End If
        End If


        Dim TopBotton As String = VB6.LCase(Field(ColumnList, " ", 1))
        If TopBotton = "top" Or TopBotton = "bottom" Then
            ColumnList = VB6.LTrim(VB6.Mid(ColumnList, Len(TopBotton) + 1))

            Dim TNum As String = Field(ColumnList, " ", 1)
            ColumnList = VB6.LTrim(VB6.Mid(ColumnList, Len(TNum) + 1))

            If IsSqlLite Then
                Limit = " LIMIT " & TNum
                TopBotton = ""
            Else
                TopBotton &= " " & TNum & " "
            End If

        Else
            TopBotton = ""
        End If

        OnlyReturnItemIDs = ColumnList = ""
        If OnlyReturnItemIDs Then ColumnList = "[" & mPrimaryKeyColumnName & "]"

        If WhereClause <> "" Then
            If VB6.Left(LCase(WhereClause), 8) = "order by" Then
                WhereClause = " " & WhereClause
            ElseIf VB6.Left(LCase(WhereClause), 8) = "by" Then
                WhereClause = " order " & WhereClause
            Else
                WhereClause = " WHERE " & WhereClause & " COLLATE NOCASE "
            End If
        End If

        SelectHandle.PrimeSqlSelect(SqlActiveConnection, "SELECT " & TopBotton & ColumnList & " FROM [" & mTableName & "]" & WhereClause & Limit, OnlyReturnItemIDs)

        SelectHandle.ActiveSelect()
        Dim mTable As DataTable = SelectHandle.mSelectedRows.Table

        ' Dim S() As String = VB6.Split(ColumnList, ",")

        If WantItemID And mTable.Columns.Contains("ItemID") Then WantItemID = False
        If WantItemContent And mTable.Columns.Contains("ItemContent") Then WantItemContent = False
        Dim NeedToBuildItemContent As Boolean = WantItemContent Or (AttributeNumsReferenced.Count > 0 And Not mTable.Columns.Contains("ItemContent"))

        If WantItemID Or WantItemContent Or NeedToBuildItemContent Then
            If WantItemID Then mTable.Columns.Add("ItemID", Type.GetType("System.String"))
            If NeedToBuildItemContent Then mTable.Columns.Add("ItemContent", Type.GetType("System.String"))

            For Each row As DataRow In mTable.Rows
                If WantItemID Then row("ItemID") = row(mPrimaryKeyColumnName)
                If NeedToBuildItemContent Then
                    Dim Results As String = ""
                    For Each Col As DataColumn In mTable.Columns
                        If Col.ColumnName <> "ItemContent" Then
                            Results &= VB6.Left(Chr(254), Len(Results))

                            If IsDBNull(row(Col)) Then
                                If Col.DataType Is GetType(System.Boolean) Then Results &= "false"
                            Else
                                If Col.DataType Is GetType(System.Boolean) Then
                                    If row(Col) Then Results &= "true" Else Results &= "false"

                                ElseIf Col.DataType Is GetType(System.Byte()) Then
                                    Dim byteArray As Byte() = row(Col)
                                    If VB6.UBound(byteArray) = 7 Then
                                        If BitConverter.IsLittleEndian Then Array.Reverse(byteArray)
                                        Dim longValue As Long = BitConverter.ToInt64(byteArray, 0)
                                        Results &= CStr(longValue)
                                    Else
                                        Results &= System.Text.Encoding.Default.GetString(row(Col))
                                    End If
                                Else
                                    Results &= row(Col).ToString
                                End If
                            End If
                        End If
                    Next
                    row("ItemContent") = Results
                End If
            Next
        End If

        If AttributeNumsReferenced.Count Then
            For Each Atr As String In AttributeNumsReferenced
                If mTable.Columns.Contains(Atr) = False Then mTable.Columns.Add(Atr, Type.GetType("System.String"))
            Next
            For Each row As DataRow In mTable.Rows
                Dim IC As String = row("ItemContent")
                For Each Atr As String In AttributeNumsReferenced
                    Dim AtrNo As Integer = Val(VB6.Mid(Atr, 3))
                    If AtrNo = 0 Then row(Atr) = row(PrimaryKeyColumnName) Else row(Atr) = Extract(IC, AtrNo)
                Next
            Next
            If WantItemContent = False Then mTable.Columns.Remove("ItemContent")
        End If

        SelectHandle.LimitToListOfPKs(SelectHandle.GetListOfPKs)

        Return SelectHandle
    End Function


    Public Sub Execute(ByVal mConnection As IDbConnection, ByVal SQL As String)
        If mConnection.State = ConnectionState.Closed Then mConnection.Open()
        Dim DA As IDbCommand = NewCommand(SQL, mConnection, mTransaction)
        DA.ExecuteScalar()
        If mTransaction Is Nothing Then mConnection.Close()
    End Sub

    Public Overrides Sub ClearFile()
        ' Validate parameters
        Execute(SqlActiveConnection, "DELETE FROM [" & mTableName & "]")
    End Sub

    Public Overrides Sub dbDeleteFile()
        Execute(SqlActiveConnection, "DROP TABLE [" & mTableName & "]")
    End Sub

    Public Overrides Sub dbDelete(ByVal ItemID As String)
        Dim ReRead As String = ""

        ' Validate parameters
        If SqlActiveConnection Is Nothing Then Throw New Exception("Ado Table handle not valid")
        If IsDBNull(ItemID) OrElse ItemID = "" Then Throw New Exception("clsTableHandleAdo-" & mTableName & "-dbDelete-Bad Primary Key")

        ' Did we pass optimistic Locking?
        'If ReadUCache.Contains(ItemID) Then
        '   If dbReadBool(ItemID, ReRead, False) Then
        '      If ReRead <> ReadUCache(ItemID) Then Throw New Exception("Optimistic Locking Failed - Item has been modified by another process")
        '   End If
        'End If

        ' Write Item to database
        Dim RowHandle As clsResultSet = New clsResultSet
        RowHandle.Open("SELECT * FROM [" & mTableName & "] WHERE [" & mPrimaryKeyColumnName & "] = " & PKQuote & VB6.Replace(ItemID, "'", "''") & PKQuote, SqlActiveConnection)

        If RowHandle.BOF And RowHandle.EOF Then Exit Sub

        Try
            RowHandle.Delete(mTableName, mPrimaryKeyColumnName, Nothing)

        Catch ex As Exception
            Throw New Exception("Table: " & mTableName & "; Error: " & ex.Message)
        End Try

        If ReadUCache.Contains(ItemID) Then ReadUCache.Remove(ItemID)
        If ReadUCache.Contains(ItemID & Chr(254) & "_JSON") Then ReadUCache.Remove(ItemID & Chr(254) & "_JSON")
        If ReadUCache.Contains(ItemID & Chr(254) & "XML") Then ReadUCache.Remove(ItemID & Chr(254) & "XML")
    End Sub
End Class

Public Class clsTableHandleDos
    Inherits clsTableHandle

    Const DictionarySuffix As String = "_D"

    Dim mFilePath As String

    Dim mColumnNames As New Collection

    Dim mTableName As String
    Dim mPrimaryKeyColumnName As String

    Public IsDict As Boolean = False

    Public Overrides ReadOnly Property TableName() As String
        Get
            TableName = mTableName
        End Get
    End Property

    Public Overrides ReadOnly Property ColumnNames() As Collection
        Get
            Return mColumnNames
        End Get
    End Property

    Public Overrides ReadOnly Property ColumnPosition(ByVal ColumnName As String) As Integer
        Get
            For I As Integer = 1 To mColumnNames.Count
                If mColumnNames(I) = ColumnName Then Return I
            Next
            Return 0
        End Get
    End Property

    Public Overrides ReadOnly Property PrimaryKeyColumnName() As String
        Get
            If mPrimaryKeyColumnName = "" Then
                mPrimaryKeyColumnName = "ItemID"

                ' potential JSON item?
                Dim O As Object = Nothing
                If IsJsonRecords(O) Then
                    For Each keyPair As KeyValuePair(Of String, Object) In CType(O, System.Collections.Generic.Dictionary(Of String, Object))
                        mPrimaryKeyColumnName = keyPair.Key
                        Exit For
                    Next
                End If
            End If

            Return mPrimaryKeyColumnName
        End Get
    End Property

    Public Function FilePath() As String
        FilePath = mFilePath
    End Function

    Public Function allowAllDosWrites() As Boolean
        If Not allowAllDosWritesChecked Then
            allowAllDosWritesChecked = True
            Dim SC As String = ConfigurationManager.AppSettings("allowAllDosWrites")
            If Not SC Is Nothing Then allowAllDosWritesValue = CNum(SC)
        End If

        Return allowAllDosWritesValue
    End Function

    Public Function allowAllDosReads() As Boolean

        If Not allowAllDosReadsChecked Then
            allowAllDosReadsChecked = True
            Dim SC As String = ConfigurationManager.AppSettings("allowAllDosReads")
            If SC Is Nothing Then allowAllDosReadsValue = allowAllDosWrites() Else allowAllDosReadsValue = CNum(SC)
        End If

        Return allowAllDosReadsValue
    End Function

    Function validUserAccountWrite(ByVal pathname As String) As Boolean
        If allowAllDosWrites() Then
            Return True
        Else
            If VB6.Left(FullUrl, 17) = "http://localhost:" Then Return True
            Dim FilePath = VB6.LCase(Path.GetFullPath(pathname))
            Dim UrlPath As String = VB6.LCase(MapRootAccount("."))
            Return VB6.Left(FilePath, Len(UrlPath)) = UrlPath
        End If
    End Function

    Function validUserAccountRead(ByVal pathname As String) As Boolean
        Dim FilePath = VB6.LCase(Path.GetFullPath(pathname))
        If VB6.Left(FullUrl, 17) = "http://localhost:" Then Return True

        If allowAllDosWrites() Then
            Return True
        Else
            Dim UrlPath As String = VB6.LCase(MapRootAccount("."))
            If VB6.Left(FilePath, Len(UrlPath)) = UrlPath Then Return True

            UrlPath = VB6.LCase(MapRootAccount("SYSPROG"))
            If VB6.Left(FilePath, Len(UrlPath)) = UrlPath Then Return True

            UrlPath = VB6.LCase(MapRootAccount("MODELER"))
            If VB6.Left(FilePath, Len(UrlPath)) = UrlPath Then Return True

            UrlPath = VB6.LCase(MapRootAccount("SYSTEM"))
            If VB6.Left(FilePath, Len(UrlPath)) = UrlPath Then Return True
        End If

        Return False
    End Function

    Private Function CreatePath(ByVal PathName As String) As Boolean
        Dim SI As Short

        If allowAllDosWrites() = False AndAlso Not validUserAccountWrite(PathName) Then Throw New Exception("DOS file write outside of user account not permitted. Add allowAllDosWrites=1 to your web.config appSettings.  " & allowAllDosWrites())

        If VB6.Right(PathName, 1) <> "\" Then PathName = PathName & "\"
        If System.IO.Directory.Exists(PathName) = True Then Return True

        For SI = 2 To Len(PathName)
            If VB6.Mid(PathName, SI, 1) = "\" Then
                If System.IO.Directory.Exists(VB6.Left(PathName, SI - 1)) = False Then MkDir(VB6.Left(PathName, SI - 1))
            End If
        Next SI

        Return True
    End Function

    Function tablePath(ByVal dbAccount As clsjdbAccount, ByVal TableName As String) As String
        Dim FilePath As String = ""
        ' Setup FilePath
        If VB6.Left(TableName, 3) = "..\" Or TableName = ".." Then
            FilePath = MapRootPath("") & TableName

        ElseIf VB6.Left(TableName, 2) = ".\" Then
            FilePath = MapRootPath("") & VB6.Mid(TableName, 3)

        ElseIf TableName = "." Then
            FilePath = MapRootPath("")

        ElseIf VB6.Left(TableName, 1) = "\" Or VB6.Mid(TableName, 2, 1) = ":" Then
            FilePath = TableName

        Else
            Dim BasePath As String = ""
            If dbAccount.ConnectionType = "F" Then BasePath = dbAccount.ConnectionString Else BasePath = MapRootAccount("")
            If IsDict Then
                FilePath = BasePath & "\JSB_Dictionaries\" & GetFName(TableName)
                If VB6.UCase(VB6.Right(FilePath, 4)) <> ".DCT" Then FilePath &= ".DCT\"
            Else
                FilePath = BasePath & TableName
            End If
        End If

        FilePath = Path.GetFullPath(FilePath)

        If VB6.Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"

        Return FilePath

    End Function

    Public Overrides Function dbOpenBool(ByVal dbAccount As clsjdbAccount, ByVal TableName As String, ByRef Errors As String, Optional ByVal CreateIt As Boolean = False) As Boolean
        MyDBAccount = dbAccount
        mTableName = TableName
        mFilePath = tablePath(dbAccount, TableName)

        mColumnNames = New Collection
        mColumnNames.Add("ItemContent")

        If Not validUserAccountRead(mFilePath) Then Throw New Exception("DOS file read outside of user account not permitted")

        If System.IO.Directory.Exists(VB6.Left(mFilePath, Len(mFilePath) - 1)) = False Then
            Errors = mFilePath & " does not exist"
            If CreateIt = False Then Return False ' Failed to open
            Try
                If CreatePath(mFilePath) = False Then Return False
            Catch ex As Exception
                Errors = ex.Message
            End Try
        End If

        ReadUCache = New Collection
        MyDBAccount = dbAccount
        OneChunkRecord = True

        Return True
    End Function

    Public Overrides Function dbReadBool(ByVal ItemID As String, ByRef Result As String, Optional ByVal IsReadU As Boolean = False) As Boolean
        ' Validate parameters

        If mFilePath = "" Then Throw New Exception("Table handle not valid")
        If IsDBNull(ItemID) OrElse ItemID = "" Then Return False

        Dim mFullPath As String = Path.GetFullPath(LCase(mFilePath & DosEncodeID(ItemID)))
        If Not validUserAccountRead(mFilePath) Then Throw New Exception("DOS file read outside of user account not permitted")

        If System.IO.File.Exists(mFullPath) = False Then
            Dim oldFullPath = ""
            Try
                oldFullPath = Path.GetFullPath(LCase(mFilePath & ItemID))
                If System.IO.File.Exists(oldFullPath) = False Then Return False
            Catch ex As Exception
                Return False
            End Try

            Try
                Dim b() As Byte = System.IO.File.ReadAllBytes(oldFullPath)
                System.IO.File.WriteAllBytes(mFullPath, b)
                System.IO.File.Delete(oldFullPath)
            Catch ex As Exception
                mFullPath = oldFullPath
            End Try
        End If

        If PCodeFile(ItemID) Then
            Dim objReader As StreamReader = New StreamReader(mFullPath)
            Result = objReader.ReadToEnd()
            objReader.Close()

        Else
            Dim b() As Byte = System.IO.File.ReadAllBytes(mFullPath)
            Dim s As New StringBuilder
            For i As Int16 = 0 To b.Length - 1
                s.Append(Chr(b(i)))
            Next
            Result = System.Text.Encoding.Default.GetString(b) ' determine encoding automatically (true)
            Result = s.ToString

            If CrlfFile(ItemID) Then
                Result = VB6.Replace(Result, Chr(239) & Chr(191) & Chr(189), Chr(254))
                Result = VB6.Replace(Result, Chr(195) & Chr(190), Chr(254))
                Result = VB6.Replace(Result, Chr(13) & Chr(10), Chr(254))
                Result = VB6.Replace(Result, Chr(10), Chr(254))
                Result = VB6.Replace(Result, Chr(13), Chr(254))
            End If
        End If

        If IsReadU Then
            If ReadUCache.Contains(mFullPath) Then ReadUCache.Remove(mFullPath)
            ReadUCache.Add(Result, mFullPath)
        End If

        Return True
    End Function

    Public Overrides Sub dbWriteJSon(ByVal Item As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        If IsJsonObj(Item) Then Item = JSON_Obj2Str(Item)
        dbWrite(Item, ItemID, IsWriteU)
    End Sub

    Public Overrides Sub dbWriteXML(ByVal Item As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        If IsXmlObj(Item) Then Item = XML_Obj2Str(Item)
        dbWrite(Item, ItemID, IsWriteU)
    End Sub

    Public Overrides Sub dbWrite(ByVal Item As String, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        Dim ReRead As String = "", Errors As String = "", mFullPath As String = ""

        ' Validate parameters
        If mFilePath = "" Then Throw New Exception("Ado Table handle not valid")
        If IsDBNull(ItemID) OrElse ItemID = "" Then
            Throw New Exception("clsTableHandleDos-" & mTableName & "-dbWrite-Bad Primary Key")
        End If

        mFullPath = Path.GetFullPath(mFilePath & DosEncodeID(ItemID))
        If Not validUserAccountWrite(mFilePath) Then Throw New Exception("DOS file write outside of user account not permitted")

        ' Did we pass optimistic Locking?
        If ReadUCache.Contains(mFullPath) Then
            If dbReadBool(ItemID, ReRead, False) Then
                If ReRead <> ReadUCache(mFullPath) Then Throw New Exception("Optimistic Locking Failed - Item has been modified by another process")
            End If
            ReadUCache.Remove(mFullPath)
        End If

        If Item Is Nothing Then Item = ""
        If PCodeFile(ItemID) Then
            Dim objReader As StreamWriter
            objReader = New StreamWriter(mFullPath)
            objReader.Write(Item)
            objReader.Close()
        Else
            If CrlfFile(ItemID) And Item <> "" Then Item = VB6.Replace(Item, Chr(254), vbCrLf)
            System.IO.File.WriteAllBytes(mFullPath, System.Text.Encoding.Default.GetBytes(Item))
        End If

        ' Update cache as necessary
        If IsWriteU Then dbReadBool(ItemID, ReRead, True)
    End Sub

    Public Overrides Function dbReadJSon(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
        Dim Item As String = ""
        If Not dbReadBool(ItemID, Item, IsReadU) Then Return False
        If VB6.Left(Item, 1) = "{" Then
            If CrlfFile(ItemID) Then Item = VB6.Replace(Item, Chr(254), " ")
            Result = JSON(Item)
        Else
            Result = JSON("{ItemID:" & JSonEncodeString(ItemID) & ",ItemContent:" & JSonEncodeString(Item) & "}")
        End If

        Return True
    End Function

    Public Overrides Function dbReadXML(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
        Dim Item As String = ""
        If Not dbReadBool(ItemID, Item, IsReadU) Then Return False
        If VB6.Left(Item, 1) = "<" Then
            Result = JSON(Item)
        Else
            Result = JSON("<ItemID>" & JSonEncodeString(ItemID) & "</ItemID><ItemContent>" & JSonEncodeString(Item) & "</ItemContent>")
        End If

        Return True
    End Function

    Private Function CrlfFile(ByVal mFullPath As String) As Boolean
        Return Not IsBinaryFile(mFullPath)

        'Dim Ext As String = VB6.LCase(GetExtension(mFullPath))
        'If Ext = "TXT" Or Ext = "" Or Ext = "JS" Or Ext = "VB" Or Ext = "CSS" Or Ext = "C" Or Ext = "CS" Or Ext = "HTM" Or Ext = "HTML" Or Ext = "ASPX" Or Ext = "BAT" Or Ext = "ASP" Or Ext = "YAML" Or Ext = "CONFIG" Then Return True
        ''Dim info As New FileInfo(mFullPath)
        ''Dim TableName = VB6.UCase(GetFName(info.DirectoryName))
        ''Return TableName = "MD" Or VB6.Left(TableName, 4) = "JSB_"
        'Return False
    End Function

    Private Function IsBinaryFile(ByVal mFullPath As String) As Boolean
        Dim Ext As String = VB6.LCase(GetExtension(mFullPath))
        Static BinaryList As New ArrayList
        If Len(Ext) < 2 Then Return False
        If BinaryList.Count = 0 Then
            BinaryList.Add("264")
            BinaryList.Add("3g2")
            BinaryList.Add("3gp")
            BinaryList.Add("asf")
            BinaryList.Add("asx")
            BinaryList.Add("avi")
            BinaryList.Add("bik")
            BinaryList.Add("dash")
            BinaryList.Add("dat")
            BinaryList.Add("dvr")
            BinaryList.Add("flv")
            BinaryList.Add("h264")
            BinaryList.Add("m2t")
            BinaryList.Add("m2ts")
            BinaryList.Add("m4v")
            BinaryList.Add("mkv")
            BinaryList.Add("mod")
            BinaryList.Add("mov")
            BinaryList.Add("mp4")
            BinaryList.Add("mpeg")
            BinaryList.Add("mpg")
            BinaryList.Add("mswmm")
            BinaryList.Add("mts")
            BinaryList.Add("ogv")
            BinaryList.Add("prproj")
            BinaryList.Add("rec")
            BinaryList.Add("rmvb")
            BinaryList.Add("swf")
            BinaryList.Add("tod")
            BinaryList.Add("tp")
            BinaryList.Add("ts")
            BinaryList.Add("vob")
            BinaryList.Add("webm")
            BinaryList.Add("wmv")
            '   Audio files
            BinaryList.Add("3ga")
            BinaryList.Add("aac")
            BinaryList.Add("aiff")
            BinaryList.Add("amr")
            BinaryList.Add("ape")
            BinaryList.Add("asf")
            BinaryList.Add("asx")
            BinaryList.Add("cda")
            BinaryList.Add("dvf")
            BinaryList.Add("flac")
            BinaryList.Add("gp4")
            BinaryList.Add("gp5")
            BinaryList.Add("gpx")
            BinaryList.Add("logic")
            BinaryList.Add("m4a")
            BinaryList.Add("m4b")
            BinaryList.Add("m4p")
            BinaryList.Add("midi")
            BinaryList.Add("mp3")
            BinaryList.Add("ogg")
            BinaryList.Add("pcm")
            BinaryList.Add("rec")
            BinaryList.Add("snd")
            BinaryList.Add("sng")
            BinaryList.Add("uax")
            BinaryList.Add("wav")
            BinaryList.Add("wma")
            BinaryList.Add("wpl")
            '  Bitmap images
            BinaryList.Add("bmp")
            BinaryList.Add("dib")
            BinaryList.Add("dng")
            BinaryList.Add("dt2")
            BinaryList.Add("emf")
            BinaryList.Add("gif")
            BinaryList.Add("ico")
            BinaryList.Add("icon")
            BinaryList.Add("jpeg")
            BinaryList.Add("jpg")
            BinaryList.Add("pcx")
            BinaryList.Add("pic")
            BinaryList.Add("png")
            BinaryList.Add("psd")
            BinaryList.Add("raw")
            BinaryList.Add("tga")
            BinaryList.Add("thm")
            BinaryList.Add("tif")
            BinaryList.Add("tiff")
            BinaryList.Add("wbmp")
            BinaryList.Add("wdp")
            BinaryList.Add("webp")

            '   Digital camera RAW photos
            BinaryList.Add("arw")
            BinaryList.Add("cr2")
            BinaryList.Add("crw")
            BinaryList.Add("dcr")
            BinaryList.Add("dng")
            BinaryList.Add("fpx")
            BinaryList.Add("mrw")
            BinaryList.Add("nef")
            BinaryList.Add("orf")
            BinaryList.Add("pcd")
            BinaryList.Add("ptx")
            BinaryList.Add("raf")
            BinaryList.Add("raw")
            BinaryList.Add("rw2")
            '   Vector graphics
            BinaryList.Add("cdr")
            BinaryList.Add("csh")
            BinaryList.Add("drw")
            BinaryList.Add("emz")
            BinaryList.Add("odg")
            BinaryList.Add("pic")
            BinaryList.Add("sda")
            BinaryList.Add("svg")
            BinaryList.Add("swf")
            BinaryList.Add("wmf")
            '   Graphics file types
            BinaryList.Add("abr")
            BinaryList.Add("ai")
            BinaryList.Add("ani")
            BinaryList.Add("cdt")
            BinaryList.Add("cpt")
            BinaryList.Add("djvu")
            BinaryList.Add("eps")
            BinaryList.Add("fla")
            BinaryList.Add("icns")
            BinaryList.Add("ico")
            BinaryList.Add("icon")
            BinaryList.Add("mdi")
            BinaryList.Add("odg")
            BinaryList.Add("pic")
            BinaryList.Add("ps")
            BinaryList.Add("psb")
            BinaryList.Add("psd")
            BinaryList.Add("pzl")
            BinaryList.Add("vsdx")
            '   3D graphics
            BinaryList.Add("3d")
            BinaryList.Add("3ds")
            BinaryList.Add("c4d")
            BinaryList.Add("dgn")
            BinaryList.Add("dwfx")
            BinaryList.Add("dwg")
            BinaryList.Add("dxf")
            BinaryList.Add("max")
            BinaryList.Add("pro")
            BinaryList.Add("pts")
            BinaryList.Add("skp")
            BinaryList.Add("stl")
            BinaryList.Add("u3d")
            BinaryList.Add("x_t")
            '   Font files
            BinaryList.Add("eot")
            BinaryList.Add("otf")
            BinaryList.Add("ttc")
            BinaryList.Add("ttf")
            BinaryList.Add("woff")
            '   Documents
            BinaryList.Add("abw")
            BinaryList.Add("aww")
            BinaryList.Add("chm")
            BinaryList.Add("cnt")
            BinaryList.Add("dbx")
            BinaryList.Add("djvu")
            BinaryList.Add("doc")
            BinaryList.Add("docm")
            BinaryList.Add("docx")
            BinaryList.Add("dot")
            BinaryList.Add("dotm")
            BinaryList.Add("dotx")
            BinaryList.Add("epub")
            BinaryList.Add("gp4")
            BinaryList.Add("gp5")
            BinaryList.Add("ind")
            BinaryList.Add("indd")
            BinaryList.Add("key")
            BinaryList.Add("mht")
            BinaryList.Add("mpp")
            BinaryList.Add("mpt")
            BinaryList.Add("odf")
            BinaryList.Add("ods")
            BinaryList.Add("odt")
            BinaryList.Add("ott")
            BinaryList.Add("oxps")
            BinaryList.Add("pdf")
            BinaryList.Add("pmd")
            BinaryList.Add("pot")
            BinaryList.Add("potx")
            BinaryList.Add("pps")
            BinaryList.Add("ppsx")
            BinaryList.Add("ppt")
            BinaryList.Add("pptm")
            BinaryList.Add("pptx")
            BinaryList.Add("prn")
            BinaryList.Add("prproj")
            BinaryList.Add("pub")
            BinaryList.Add("pwi")
            BinaryList.Add("rep")
            BinaryList.Add("rtf")
            BinaryList.Add("sdd")
            BinaryList.Add("sdw")
            BinaryList.Add("shs")
            BinaryList.Add("snp")
            BinaryList.Add("sxw")
            BinaryList.Add("tpl")
            BinaryList.Add("vsd")
            BinaryList.Add("wlmp")
            BinaryList.Add("wpd")
            BinaryList.Add("wps")
            BinaryList.Add("wri")
            BinaryList.Add("xps")

            '   Simple text files
            BinaryList.Add("lrc")
            BinaryList.Add("nfo")
            BinaryList.Add("opml")
            BinaryList.Add("pts")
            BinaryList.Add("rep")
            BinaryList.Add("rtf")
            BinaryList.Add("srt")

            '   E-book files
            BinaryList.Add("azw")
            BinaryList.Add("azw3")
            BinaryList.Add("cbr")
            BinaryList.Add("cbz")
            BinaryList.Add("epub")
            BinaryList.Add("fb2")
            BinaryList.Add("iba")
            BinaryList.Add("ibooks")
            BinaryList.Add("lit")
            BinaryList.Add("mobi")
            BinaryList.Add("pdf")
            '   Spreadsheets
            BinaryList.Add("ods")
            BinaryList.Add("sdc")
            BinaryList.Add("sxc")
            BinaryList.Add("xls")
            BinaryList.Add("xlsm")
            BinaryList.Add("xlsx")
            '   Microsoft Office files
            BinaryList.Add("accdb")
            BinaryList.Add("accdt")
            BinaryList.Add("doc")
            BinaryList.Add("docm")
            BinaryList.Add("docx")
            BinaryList.Add("dot")
            BinaryList.Add("dotm")
            BinaryList.Add("dotx")
            BinaryList.Add("mdb")
            BinaryList.Add("mpd")
            BinaryList.Add("mpp")
            BinaryList.Add("mpt")
            BinaryList.Add("one")
            BinaryList.Add("onepkg")
            BinaryList.Add("pot")
            BinaryList.Add("potx")
            BinaryList.Add("pps")
            BinaryList.Add("ppsx")
            BinaryList.Add("ppt")
            BinaryList.Add("pptm")
            BinaryList.Add("pptx")
            BinaryList.Add("pst")
            BinaryList.Add("pub")
            BinaryList.Add("snp")
            BinaryList.Add("thmx")
            BinaryList.Add("vsd")
            BinaryList.Add("vsdx")
            BinaryList.Add("xls")
            BinaryList.Add("xlsm")
            BinaryList.Add("xlsx")
            '   misc
            BinaryList.Add("cab")
            BinaryList.Add("lng")
            BinaryList.Add("res")
            BinaryList.Add("swf")
            BinaryList.Add("vhd")
            BinaryList.Add("vmx")

            '   Virtualization software related files
            BinaryList.Add("ova")
            BinaryList.Add("ovf")
            BinaryList.Add("pvm")
            BinaryList.Add("vdi")
            BinaryList.Add("vmdk")
            BinaryList.Add("vmem")
            BinaryList.Add("vmwarevm")
            BinaryList.Add("gdoc")

            BinaryList.Add("gsheet")
            BinaryList.Add("gslides")
            BinaryList.Add("eml")
            BinaryList.Add("flv")

            '   Internet related files
            BinaryList.Add("ashx")
            BinaryList.Add("atom")
            BinaryList.Add("crdownload")
            BinaryList.Add("dlc")
            BinaryList.Add("download")

            '   Email files
            BinaryList.Add("dbx")
            BinaryList.Add("eml")
            BinaryList.Add("ldif")
            BinaryList.Add("mht")
            BinaryList.Add("msg")
            BinaryList.Add("pst")
            BinaryList.Add("vcf")

            '   File extensions blocked by mail clients
            BinaryList.Add("chm")
            BinaryList.Add("com")
            BinaryList.Add("cpl")
            BinaryList.Add("eml")
            BinaryList.Add("exe")
            BinaryList.Add("inf")
            BinaryList.Add("mdb")
            BinaryList.Add("msi")
            BinaryList.Add("prg")
            BinaryList.Add("reg")
            BinaryList.Add("scr")
            BinaryList.Add("shs")

            '   Possibly dangerous files
            BinaryList.Add("bin")
            BinaryList.Add("com")
            BinaryList.Add("cpl")
            BinaryList.Add("dll")
            BinaryList.Add("drv")
            BinaryList.Add("exe")
            BinaryList.Add("jar")
            BinaryList.Add("ocx")
            BinaryList.Add("pcx")
            BinaryList.Add("scr")
            BinaryList.Add("shs")
            BinaryList.Add("swf")
            BinaryList.Add("sys")
            BinaryList.Add("vxd")
            BinaryList.Add("wmf")

            ' Archives()
            BinaryList.Add("7z")
            BinaryList.Add("7zip")

            BinaryList.Add("ace")
            BinaryList.Add("air")
            BinaryList.Add("apk")
            BinaryList.Add("arc")
            BinaryList.Add("arj")
            BinaryList.Add("asec")
            BinaryList.Add("bar")
            BinaryList.Add("bin")
            BinaryList.Add("cab")
            BinaryList.Add("cbr")
            BinaryList.Add("cbz")
            BinaryList.Add("cso")
            BinaryList.Add("deb")
            BinaryList.Add("dlc")
            BinaryList.Add("gz")
            BinaryList.Add("gzip")
            BinaryList.Add("hqx")
            BinaryList.Add("inv")
            BinaryList.Add("ipa")
            BinaryList.Add("isz")
            BinaryList.Add("jar")
            BinaryList.Add("msu")
            BinaryList.Add("nbh")
            BinaryList.Add("pak")
            BinaryList.Add("rar")

            BinaryList.Add("tar")
            BinaryList.Add("tar.gz")
            BinaryList.Add("tgz")
            BinaryList.Add("uax")
            BinaryList.Add("webarchive")

        End If
        Return BinaryList.Contains(Ext)

    End Function

    Private Function PCodeFile(ByVal ItemID As String) As Boolean
        Dim Ext As String
        Ext = VB6.UCase(GetExtension(ItemID))
        Return Ext = "PCD" Or Ext = "PCS" Or Ext = "PCF"
    End Function

    Public Overrides Function SelectFile() As rSelectList ' (Optional ByVal FroTables As String = "", Optional ByVal FilterTxt s String = "", Optional ByVal OrderByTxt As String = ""A) As rSelectList
        Dim SelectHandle As New rSelectList
        Dim FormList As String = ""

        ' Validate parameters
        If mFilePath = "" Then Throw New Exception("Ado Table handle not valid")
        Try
            Dim dDir As New DirectoryInfo(mFilePath)
            Dim fFileSystemInfo As FileSystemInfo
            For Each fFileSystemInfo In dDir.GetFileSystemInfos()
                If (fFileSystemInfo.Attributes And FileAttributes.Directory) = 0 Then
                    FormList = FormList & Chr(254) & DosDecodeID(fFileSystemInfo.Name)
                End If
            Next
        Catch ex As Exception
        End Try

        SelectHandle.SetDynamicArray(VB6.Mid(FormList, 2), True)

        Return SelectHandle
    End Function

    Public Function IsJsonRecords(ByRef firstJSonRec As Object) As Boolean
        Dim dDir1 As New DirectoryInfo(mFilePath), ItemID As String = ""
        Dim fFileSystemInfo1 As FileSystemInfo, Tags As New Collection, JsonCnt As Integer = 0

        mPrimaryKeyColumnName = ""
        For Each fFileSystemInfo1 In dDir1.GetFileSystemInfos()
            If (fFileSystemInfo1.Attributes And FileAttributes.Directory) = 0 Then
                ItemID = DosDecodeID(fFileSystemInfo1.Name)
                If ItemID <> "" AndAlso VB6.Left(ItemID, 2) <> "__" Then
                    Dim Item As String = ""
                    If dbReadBool(ItemID, Item, False) Then
                        If VB6.Left(Item, 1) = "{" Then
                            Try
                                firstJSonRec = JSON(Item)
                                For Each keyPair As KeyValuePair(Of String, Object) In CType(firstJSonRec, System.Collections.Generic.Dictionary(Of String, Object))
                                    If mPrimaryKeyColumnName = "" Then
                                        mPrimaryKeyColumnName = keyPair.Key

                                    ElseIf mPrimaryKeyColumnName = keyPair.Key Then
                                        JsonCnt += 1
                                        If JsonCnt = 10 Then Return True ' assume if 10 items are good, then all is good
                                    Else
                                        mPrimaryKeyColumnName = "ItemID"
                                        Return False
                                    End If

                                    Exit For
                                Next
                            Catch ex2 As Exception
                                mPrimaryKeyColumnName = "ItemID"
                                Return False
                            End Try
                        Else
                            mPrimaryKeyColumnName = "ItemID"
                            Return False
                        End If
                    Else
                        mPrimaryKeyColumnName = "ItemID"
                        Return False
                    End If
                End If
            End If
        Next

        If mPrimaryKeyColumnName = "" Then mPrimaryKeyColumnName = "ItemID" Else Return True
        Return Nothing
    End Function

    Public Overrides Function SelectFileX(ByVal ColumnList As String, ByVal WhereClause As String) As rSelectList
        Dim RS As rSelectList
        Dim FormList As String = "", returnOnlyItemIDs As Boolean
        Dim Top As Boolean = False
        Dim Bottom As Boolean = False
        Dim TopBottomCnt As Long = 0
        Dim SpecialFields As New Collection
        Dim FirstID As String = ""
        Dim mTable As New DataTable(mTableName)
        Dim JSonItem As Boolean = False
        Dim AttributeNumsReferenced As New Collection
        ColumnList = VB6.RTrim(LTrim(ColumnList))
        Dim WantItemID As Boolean = ColumnList = ""
        Dim WantItemContent As Boolean = False
        Dim hasStarColumn As Boolean = False

        ' Validate parameters
        If mFilePath = "" Then Throw New Exception("Ado Table handle not valid")

        WhereClause = VB6.LTrim(WhereClause)
        If VB6.Left(VB6.LCase(WhereClause), 6) = "where " Then WhereClause = VB6.Mid(WhereClause, 7)
        If VB6.InStr(WhereClause, "'") = 0 Then WhereClause = VB6.Replace(WhereClause, """", "'")
        If VB6.InStr(VB6.LCase(WhereClause), " like ") Then
            WhereClause = VB6.Replace(WhereClause, "'[", "'%")
            WhereClause = VB6.Replace(WhereClause, "]'", "%'")
        End If


        ColumnList = VB6.LTrim(ColumnList)
        Dim token As String = VB6.LCase(Field(ColumnList, " ", 1))

        Top = token = "top"
        Bottom = token = "bottom"

        If Top Or Bottom Then
            ColumnList = VB6.Mid(ColumnList, Len(token) + 1)
            ColumnList = VB6.LTrim(ColumnList)
            Dim TNum As String = Field(ColumnList, " ", 1)
            ColumnList = VB6.LTrim(VB6.Mid(ColumnList, Len(TNum) + 1))
            TopBottomCnt = Val(TNum)
        End If

        returnOnlyItemIDs = ColumnList = ""

        ' Dim S() As String = VB6.Split(ColumnList, ",")
        Dim S As ArrayList = activeColumns(ColumnList & " " & WhereClause, False, WantItemID, WantItemContent, hasStarColumn, AttributeNumsReferenced)
        Dim NeedItemID As Boolean = WantItemID Or returnOnlyItemIDs
        Dim NeedItemContent As Boolean = WantItemContent

        For Each ColName In S
            Dim LColName As String = VB6.LCase(ColName)
            If LColName <> "itemid" AndAlso LColName <> "itemcontent" AndAlso LColName <> "*" AndAlso LColName <> "" Then
                If SpecialFields.Contains(ColName) = False Then SpecialFields.Add(ColName, ColName)
            End If
        Next

        ' potential JSON item?
        If Not returnOnlyItemIDs Or SpecialFields.Count Then
            Dim O As Object = Nothing
            If IsJsonRecords(O) Then
                JSonItem = True
                If hasStarColumn Then NeedItemID = True
                For Each keyPair As KeyValuePair(Of String, Object) In CType(O, System.Collections.Generic.Dictionary(Of String, Object))
                    Try
                        If keyPair.Value Is Nothing Then
                            mTable.Columns.Add(keyPair.Key, Type.GetType("System.String"))
                        Else
                            mTable.Columns.Add(keyPair.Key, keyPair.Value.GetType)
                        End If
                    Catch ex As Exception
                    End Try
                Next
            Else
                If hasStarColumn Then
                    NeedItemContent = True
                    NeedItemID = True
                End If
            End If
        End If

        If NeedItemID Then mTable.Columns.Add("ItemID", Type.GetType("System.String"))
        If NeedItemContent Then mTable.Columns.Add("ItemContent", Type.GetType("System.String"))

        For Each AtrNo As String In AttributeNumsReferenced
            If mTable.Columns.Contains(AtrNo) = False Then mTable.Columns.Add(AtrNo, Type.GetType("System.String"))
        Next

        Dim dDir As New DirectoryInfo(mFilePath)
        Dim fFileSystemInfo As FileSystemInfo
        For Each fFileSystemInfo In dDir.GetFileSystemInfos()
            If (fFileSystemInfo.Attributes And FileAttributes.Directory) Then
                If NeedItemID And Not NeedItemContent Then
                    Dim R As DataRow = mTable.NewRow
                    R("ItemID") = "[" & DosDecodeID(fFileSystemInfo.Name) & "]"
                    mTable.Rows.Add(R)
                End If
            Else
                Dim R As DataRow = mTable.NewRow
                Dim Item As String = ""
                Dim ItemID As String = DosDecodeID(fFileSystemInfo.Name)
                If NeedItemID Then R("ItemID") = ItemID

                If JSonItem Or NeedItemContent Or SpecialFields.Count > 0 Or AttributeNumsReferenced.Count > 0 Then
                    If dbReadBool(ItemID, Item, False) Then
                        If NeedItemContent Then R("ItemContent") = Item
                        For Each Atr As String In AttributeNumsReferenced
                            Dim AtrNo As Integer = Val(VB6.Mid(Atr, 3))
                            If AtrNo = 0 Then
                                R(Atr) = ItemID
                            Else
                                R(Atr) = Field(Item, Chr(254), AtrNo)
                            End If

                        Next

                        If JSonItem And VB6.Left(Item, 1) = "{" Then
                            Try
                                Dim O As Object = JSON(Item)
                                For Each keyPair As KeyValuePair(Of String, Object) In CType(O, System.Collections.Generic.Dictionary(Of String, Object))
                                    Try
                                        If mTable.Columns.Contains(keyPair.Key) = False Then
                                            If keyPair.Value IsNot Nothing Then mTable.Columns.Add(keyPair.Key, keyPair.Value.GetType)
                                        End If
                                        If keyPair.Value IsNot Nothing Then
                                            If mTable.Columns(keyPair.Key).DataType = GetType(System.Boolean) Then
                                                If keyPair.Value Then R(keyPair.Key) = "true" Else R(keyPair.Key) = "false"
                                            Else
                                                R(keyPair.Key) = keyPair.Value
                                            End If
                                        End If

                                    Catch ex As Exception

                                    End Try
                                Next
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                End If

                mTable.Rows.Add(R)
            End If
        Next
        mTable.AcceptChanges()

        RS = New rSelectList
        RS.SetDataTable(mTable, returnOnlyItemIDs)

        ' every *axxx needs to be [*axxx]
        For Each AtrNo As String In AttributeNumsReferenced
            WhereClause = VB6.Replace(WhereClause, AtrNo, "[" & AtrNo & "*]", , , CompareMethod.Text)
        Next
        WhereClause = VB6.Replace(WhereClause, "[[*a", "[*a", , , CompareMethod.Text)
        WhereClause = VB6.Replace(WhereClause, "*]]", "*]")
        WhereClause = VB6.Replace(WhereClause, "*]", "]")
        If WhereClause <> "" Then RS.FilterList(WhereClause)

        ' Get only Top of Bot?
        If Top Then
            RS.TopFilter(TopBottomCnt)
        ElseIf Bottom Then
            RS.BottomFilter(TopBottomCnt)
        End If

        ' Remove any column not in ColumnList
        AttributeNumsReferenced = New Collection
        WantItemID = ColumnList = ""
        WantItemContent = False
        hasStarColumn = False

        S = activeColumns(ColumnList, False, WantItemID, WantItemContent, hasStarColumn, AttributeNumsReferenced)
        If hasStarColumn Then
            If JSonItem Then
                If NeedItemID And Not WantItemID Then mTable.Columns.Remove("ItemID")
                If NeedItemContent And Not WantItemContent Then mTable.Columns.Remove("ItemContent")
            End If
        Else
            Dim RemoveCols As New ArrayList
            For Each Col As DataColumn In mTable.Columns
                If S.Contains(Col.ColumnName) = False Then
                    If AttributeNumsReferenced.Contains(Col.ColumnName) = False Then
                        If Col.ColumnName = "ItemID" And WantItemID Then
                        ElseIf Col.ColumnName = "ItemContent" And WantItemContent Then
                        Else
                            RemoveCols.Add(Col)
                        End If
                    End If
                End If
            Next
            For Each Col As DataColumn In RemoveCols
                If RS.mSelectedRows.Columns.Contains(Col.ColumnName) Then RS.mSelectedRows.Columns.Remove(Col.ColumnName)
            Next
        End If


        Return RS
    End Function

    Public Overrides Sub dbTransactionBegin()
    End Sub

    Public Overrides Sub dbTransactionCommit()
    End Sub

    Public Overrides Sub dbExitServer()
    End Sub

    Public Overrides Sub ClearFile()
        Dim List As New Collection

        ' Validate parameters
        If mFilePath = "" Then Throw New Exception("Ado Table handle not valid")
        If Not validUserAccountWrite(mFilePath) Then Throw New Exception("DOS file delete outside of user account not permitted")

        ' Remove all items

        Dim dDir As New DirectoryInfo(mFilePath)
        Dim fFileSystemInfo As FileSystemInfo
        For Each fFileSystemInfo In dDir.GetFileSystemInfos()
            If (fFileSystemInfo.Attributes And FileAttributes.Directory) = 0 And
             (fFileSystemInfo.Attributes And FileAttributes.Hidden) = 0 And
             (fFileSystemInfo.Attributes And FileAttributes.System) = 0 And
             (fFileSystemInfo.Attributes And FileAttributes.Device) = 0 Then
                System.IO.File.Delete(fFileSystemInfo.FullName)
            End If
        Next

    End Sub

    Public Overrides Sub dbDeleteFile()
        System.IO.Directory.Delete(mFilePath, True)
    End Sub

    Public Overrides Sub dbDelete(ByVal ItemID As String)
        ' Validate parameters
        If mFilePath = "" Then Throw New Exception("Ado Table handle not valid")
        If IsDBNull(ItemID) OrElse ItemID = "" Then Throw New Exception("clsTableHandleDos-" & mTableName & "-dbDeleteFile-Bad Primary Key")

        Dim mFullPath As String = Path.GetFullPath(mFilePath & DosEncodeID(ItemID)), ReRead As String = ""
        If Not validUserAccountWrite(mFilePath) Then Throw New Exception("DOS file delete outside of user account not permitted")

        ' Did we pass optimistic Locking?
        If ReadUCache.Contains(mFullPath) Then
            If dbReadBool(ItemID, ReRead, False) Then
                If ReRead <> ReadUCache(mFullPath) Then Throw New Exception("Optimistic Locking Failed - Item has been modified by another process")
            End If
        End If

        System.IO.File.Delete(mFilePath & DosEncodeID(ItemID))

        If ReadUCache.Contains(mFullPath) Then ReadUCache.Remove(mFullPath)
    End Sub

End Class

Public Class clsTableHandleHttp
    Inherits clsTableHandle

    Dim myTimeOut As Integer = 33 ' seconds

    Dim mColumnNames As New Collection
    Dim AM As String = Chr(254), VM As String = Chr(253), SVM As String = Chr(252)


    Dim mTableName As String

    Public Overrides ReadOnly Property TableName() As String
        Get
            TableName = mTableName
        End Get
    End Property

    Public Overrides ReadOnly Property ColumnNames() As Collection
        Get
            Return mColumnNames
        End Get
    End Property

    Public Overrides ReadOnly Property PrimaryKeyColumnName() As String
        Get
            If mColumnNames Is Nothing OrElse mColumnNames.Count = 0 Then Return ""
            PrimaryKeyColumnName = mColumnNames(1)
        End Get
    End Property

    Public Overrides ReadOnly Property ColumnPosition(ByVal ColumnName As String) As Integer
        Get
            For i As Integer = 1 To mColumnNames.Count
                If mColumnNames(i) = ColumnName Then Return i
            Next
            Return 0
        End Get
    End Property

    ' Tablename is http://blablabla/bla.com

    Public Overrides Function dbOpenBool(ByVal dbAccount As clsjdbAccount, ByVal TableName As String, ByRef Errors As String, Optional ByVal CreateIt As Boolean = False) As Boolean
        mTableName = TableName
        mColumnNames = New Collection
        mColumnNames.Add("ItemContent")
        MyDBAccount = dbAccount
        Return True
    End Function

    Public Overrides Function dbReadBool(ByVal ItemID As String, ByRef Result As String, Optional ByVal IsReadU As Boolean = False) As Boolean
        Return dbReadWriteBool(ItemID, "", Result, "GET")
    End Function

    Public Overrides Function dbReadJSon(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
        If Not dbReadWriteBool(ItemID, "", Result, "GET") Then Return False
        Result = JSON(Result)
        Return True
    End Function

    Public Overrides Function dbReadXML(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
        If Not dbReadWriteBool(ItemID, "", Result, "GET") Then Return False
        Result = XML_Str2Obj(Result)
        Return True
    End Function

    Public Overrides Sub dbWrite(ByVal Item As String, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        If dbReadWriteBool(ItemID, C2Str(Item), LastItemOut, "POST") = False Then Throw New Exception(LastHeader)
    End Sub

    Public Overrides Sub dbWriteJSon(ByVal Item As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        If IsJsonObj(Item) Then Item = JSON_Obj2Str(Item)
        If dbReadWriteBool(ItemID, C2Str(Item), LastItemOut, "POST") = False Then Throw New Exception(LastHeader)
    End Sub

    Public Overrides Sub dbWriteXML(ByVal Item As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
        If IsXmlObj(Item) Then Item = XML_Obj2Str(Item)
        If dbReadWriteBool(ItemID, C2Str(Item), LastItemOut, "POST") = False Then Throw New Exception(LastHeader)
    End Sub

    Public Overrides Sub dbDelete(ByVal ItemID As String)
        If dbReadWriteBool(ItemID, "", LastItemOut, "DELETE") = False Then Throw New Exception(LastHeader)
    End Sub

    Private Function dbReadWriteBool(ByVal ItemID As String, ByVal ItemIn As String, ByRef ItemOut As String, ByVal DefaultMethod As String) As Boolean
        ' setup basic request
        Dim Method As String = Field(ItemID, AM, 2)
        Dim Headers As String = Field(ItemID, AM, 3)
        Dim Body As String = Field(ItemID, AM, 4)
        Dim NoRedirecting As Boolean = Field(ItemID, AM, 5) <> ""
        Dim Url As String = mTableName

        Dim ShowDbg As Boolean = VB6.InStr(ItemID, ".axd") = 0 And VB6.InStr(ItemID, ".css") = 0 And VB6.InStr(ItemID, ".js") = 0 And VB6.InStr(ItemID, ".png") = 0 And VB6.InStr(ItemID, ".bmp") = 0 And VB6.InStr(ItemID, "/css?") = 0

        If Method = "" Then Method = DefaultMethod

        ItemID = Field(ItemID, AM, 1)
        If VB6.Left(ItemID, 1) = "/" Then ItemID = VB6.Mid(ItemID, 1)
        If Len(ItemID) > 0 Then
            If VB6.Right(Url, 1) <> "/" Then Url &= "/"
            Url &= ItemID
        End If

        If Body <> "" Then ItemIn = Body
        Return UrlFetch(Method, Url, Headers, ItemIn, ItemOut, LastHeader, NoRedirecting, myTimeOut)
    End Function

    Public Overrides Function SelectFile() As rSelectList ' (Optional ByVal FroTables As String = "", Optional ByVal FilterTxt s String = "", Optional ByVal OrderByTxt As String = ""A) As rSelectList
        Throw New Exception("SelectFileX not supported")
        Return Nothing
    End Function

    Public Overrides Function SelectFileX(ByVal ColumnList As String, ByVal WhereClause As String) As rSelectList
        Throw New Exception("SelectFileX not supported")
        Return Nothing
    End Function

    Public Overrides Sub dbTransactionBegin()
    End Sub

    Public Overrides Sub dbTransactionCommit()
    End Sub

    Public Overrides Sub dbExitServer()
    End Sub

    Public Overrides Sub ClearFile()
        Throw New Exception("ClearFile not supported")
    End Sub

    Public Overrides Sub dbDeleteFile()
        Throw New Exception("http DeleteFile not supported")
    End Sub


End Class

Public Class clsResultSet
    Dim mDataset As DataSet = Nothing ' used only for .update's
    Dim mTable As DataTable = Nothing
    Public abosolutePosition As Long
    Dim DA As IDbDataAdapter


    Public Sub New(ByVal List As String)
        mTable = New DataTable("FormList")
        mTable.Columns.Add("ItemID", Type.GetType("System.String"))

        abosolutePosition = 0
        DA = Nothing
        If List.Length > 0 Then
            Dim FormList() As String = VB6.Split(List, Chr(254))

            For FormListI As Integer = 0 To VB6.UBound(FormList)
                Dim R As DataRow = mTable.NewRow
                R("ItemID") = FormList(FormListI)
                mTable.Rows.Add(R)
            Next
            mTable.AcceptChanges()
        End If
    End Sub

    Public Sub FilterList(ByVal Filter As String)
        Dim OrderBY As String = ""

        If Filter = "" Then Return

        mTable.CaseSensitive = False

        Dim I As Integer = VB6.InStr(LCase(Filter), "order by")
        If I Then
            OrderBY = VB6.Mid(Filter, I + Len("order by") + 1)
            Filter = VB6.Left(Filter, I - 1)
        End If

        Dim DRs() As DataRow
        Try
            DRs = mTable.Select(Filter, OrderBY)
        Catch ex As Exception
            If OrderBY <> "" Then
                Throw New Exception("Unable to apply filter: " & Filter & " order by " & OrderBY & "; Error: " & ex.Message)
            Else
                Throw New Exception("Unable to apply filter: " & Filter & "; Error: " & ex.Message)
            End If
        End Try

        If VB6.UBound(DRs) >= 0 Then
            mTable = DRs.CopyToDataTable
        Else
            mTable = New DataTable("FormList")
            mTable.Columns.Add("ItemID", Type.GetType("System.String"))
        End If
        abosolutePosition = 0

    End Sub

    Public Sub TopLimit(n As Integer)
        If mTable.Rows.Count > n Then
            mTable = mTable.AsEnumerable().Skip(0).Take(n).CopyToDataTable()
        End If
    End Sub

    Public Sub BottomLimit(n As Integer)
        If mTable.Rows.Count > n Then
            mTable = mTable.AsEnumerable().Skip(mTable.Rows.Count - n).Take(n).CopyToDataTable()
        End If
    End Sub

    Public ReadOnly Property Table() As DataTable
        Get
            Return mTable
        End Get
    End Property

    Public ReadOnly Property RecordCount() As Long
        Get
            Return mTable.Rows.Count
        End Get
    End Property

    Public ReadOnly Property IsOpen() As Boolean
        Get
            Return Not mTable Is Nothing
        End Get
    End Property

    Public ReadOnly Property IsClosed() As Boolean
        Get
            Return mTable Is Nothing
        End Get
    End Property

    Public Function Open(ByVal Sql As String, ByVal Connection As IDbConnection) As Boolean
        mTable = GetDataTable(Sql, Connection)
        abosolutePosition = 0
        Return True
    End Function

    Private Function GetDataTable(ByVal Sql As String, ByVal Connection As IDbConnection) As DataTable
        If Connection.State = ConnectionState.Closed Then Connection.Open()

        Using myCommand As IDbCommand = NewCommand(Sql, Connection)
            Using myReader As IDataReader = myCommand.ExecuteReader()
                Dim myTable As New DataTable()
                myTable.Load(myReader)
                Connection.Close()
                Return myTable
            End Using
        End Using
    End Function

    Public Function setTable(ByVal dataTable As DataTable) As Boolean
        mTable = dataTable
        abosolutePosition = 0
        Return True
    End Function

    Public ReadOnly Property EOF() As Boolean
        Get
            Return abosolutePosition >= RecordCount()
        End Get
    End Property

    Public ReadOnly Property BOF() As Boolean
        Get
            Return abosolutePosition <= 0
        End Get
    End Property

    Public Sub MoveFirst()
        abosolutePosition = 0
    End Sub

    Public Sub MoveLast()
        abosolutePosition = RecordCount() - 1
    End Sub

    Public Sub MoveNext()
        If Not EOF() Then abosolutePosition = abosolutePosition + 1
    End Sub

    Public Sub MovePrevious()
        If Not BOF() Then abosolutePosition = abosolutePosition - 1
    End Sub

    Public Function Fields() As DataRow
        If abosolutePosition >= 0 AndAlso abosolutePosition < RecordCount() Then Return mTable.Rows(abosolutePosition) Else Return Nothing
    End Function

    Public ReadOnly Property Columns() As DataColumnCollection
        Get
            Return mTable.Columns
        End Get
    End Property

    Public Sub Update(ByVal TableName As String, ByVal PK1 As String, ByVal PK2 As String)
        SetupSqlCommands(DA, TableName, PK1, PK2)
        DA.Update(mDataset)
    End Sub

    Public Sub Delete(ByVal TableName As String, ByVal PK1 As String, ByVal PK2 As String)
        Fields.Delete()
        Update(TableName, PK1, PK2)
    End Sub

    Public Sub AddNew()
        mTable.Rows.Add(mTable.NewRow)
        abosolutePosition = mTable.Rows.Count - 1
    End Sub

    Public Sub New()
        ' Not to do
    End Sub

End Class

Public Class rSelectList
    Public TblName As String ' Used to make ID into PF table
    Public PK1 As String
    Public PK2 As String
    Public mSelectedRows As clsResultSet ' Record data
    Public PrimedSqlSelect As String
    Public OnlyReturnItemIDs As Boolean = True ' for GetList

    Dim myConnection As IDbConnection = Nothing

    Public Sub SetJSonArray(ByRef PrimaryKeys As String, ByRef JsonRows As Object, ByVal OnlyReturnItemIDs As Boolean)
        ' Build a datatable based off the tags in JsonObject
        If OnlyReturnItemIDs Then
            SetDynamicArray(PrimaryKeys, OnlyReturnItemIDs)
            Return
        End If

        TblName = ""
        PrimedSqlSelect = ""
        myConnection = Nothing
        Me.OnlyReturnItemIDs = OnlyReturnItemIDs
        mSelectedRows = New clsResultSet()

        If TypeOf JsonRows Is String Then JsonRows = JSON(JsonRows)
        If JsonRows.Count = 0 Then Return

        ' Loop on each Tag in JsonObject(0)
        '    Create Column
        ' repeat
        Dim myTable As New DataTable
        Dim JsonPair = JsonRows(0)
        For Each Col As Object In JsonPair
            Dim Key As String = Col.Key
            Dim Value As Object = Col.Value
            If Value Is Nothing Then
                Value = ""
                For Each jRow In JsonRows
                    If jRow(Key) IsNot Nothing Then
                        Value = jRow(Key)
                        Exit For
                    End If
                Next
            End If
            myTable.Columns.Add(Key, Value.GetType)
        Next

        ' Loop on Each Row in JsonObject
        '  C = table.NewRow
        '   loop on each column
        '     c(col) = Row(Col)
        '   repeat
        For Each JRow In JsonRows
            Dim DRow As DataRow = myTable.NewRow
            For Each Col As DataColumn In myTable.Columns
                Dim V As Object = JsonChild(JRow, Col.ColumnName)
                If IsNothing(V) Then
                    DRow(Col) = DBNull.Value
                Else
                    DRow(Col) = V
                End If
            Next
            myTable.Rows.Add(DRow)
        Next

        mSelectedRows.setTable(myTable)
    End Sub

    Public Sub SetDynamicArray(ByRef List As String, ByVal OnlyReturnItemIDs As Boolean)
        ' No lock, Read Item from database
        mSelectedRows = New clsResultSet(List)
        TblName = ""
        PrimedSqlSelect = ""
        myConnection = Nothing
        Me.OnlyReturnItemIDs = OnlyReturnItemIDs
    End Sub

    Public Sub SetDataTable(ByRef DataTable As DataTable, ByVal OnlyReturnItemIDs As Boolean)
        ' No lock, Read Item from database
        mSelectedRows = New clsResultSet
        mSelectedRows.setTable(DataTable)
        TblName = DataTable.TableName
        PrimedSqlSelect = ""
        myConnection = Nothing
        Me.OnlyReturnItemIDs = OnlyReturnItemIDs
    End Sub

    Public Sub PrimeSqlSelect(ByRef Connection As IDbConnection, ByVal SQL As String, ByVal OnlyReturnItemIDs As Boolean)
        TblName = ""
        mSelectedRows = New clsResultSet
        PrimedSqlSelect = SQL
        myConnection = Connection
        Me.OnlyReturnItemIDs = OnlyReturnItemIDs
    End Sub


    Public ReadOnly Property RowHandle() As clsResultSet
        Get
            Return mSelectedRows
        End Get
    End Property

    Public Function Primed() As Boolean
        If mSelectedRows Is Nothing Then Return False
        If mSelectedRows.IsClosed AndAlso PrimedSqlSelect <> "" Then Return True
        Return False
    End Function

    Public Function ActiveSelect() As Boolean
        If mSelectedRows Is Nothing Then Return False

        If Primed() Then
            mSelectedRows.Open(PrimedSqlSelect, myConnection)
        End If

        Return Not RowHandle.EOF
    End Function

    Public Sub FilterList(ByVal Filter As String)
        If Not ActiveSelect() Then Return
        mSelectedRows.FilterList(Filter)
    End Sub

    Public Sub TopFilter(N As Integer)
        Dim Cnt As Integer = Count()
        If Cnt <= N Then Return
        mSelectedRows.TopLimit(N)
    End Sub

    Public Sub BottomFilter(N As Integer)
        Dim Cnt As Integer = Count()
        If Cnt <= N Then Return
        mSelectedRows.BottomLimit(N)
    End Sub

    Public Function Count() As Integer
        If Not ActiveSelect() Then Return Nothing

        If mSelectedRows.RecordCount <> -1 Then
            Count = mSelectedRows.RecordCount
            Exit Function
        End If

        mSelectedRows.MoveLast()
        If mSelectedRows.RecordCount <> -1 Then
            Count = mSelectedRows.RecordCount
            mSelectedRows.MoveFirst()
            Exit Function
        End If

        Dim CNT As Integer
        CNT = 0
        mSelectedRows.MoveFirst()
        Do
            If mSelectedRows.EOF Then Exit Do
            RowHandle.MoveNext()
            CNT = CNT + 1
        Loop
        Count = CNT
        mSelectedRows.MoveFirst()
    End Function

    Public Function ReadPrev() As String
        If Not ActiveSelect() Then Return Nothing

        If mSelectedRows.BOF Then Return Nothing
        mSelectedRows.MovePrevious()

        If mSelectedRows.Columns.Contains("ItemID") Then
            If IsDBNull(mSelectedRows.Fields("ItemID")) Then ReadPrev = "" Else ReadPrev = mSelectedRows.Fields("ItemID").ToString
        Else
            If IsDBNull(mSelectedRows.Fields(0)) Then ReadPrev = "" Else ReadPrev = mSelectedRows.Fields(0).ToString
        End If
    End Function

    Public Function EOF() As Boolean
        If Not ActiveSelect() Then Return True
        Return mSelectedRows.EOF
    End Function

    Public Function BOF() As Boolean
        If Not ActiveSelect() Then Return False
        Return mSelectedRows.BOF
    End Function

    Public Function ReadNextBool(ByRef ID As String, Optional ByVal Value As Integer = 0, Optional ByVal SubValue As Integer = 0) As Boolean
        If Not ActiveSelect() Then Return False

        If mSelectedRows.EOF Then Return False
        If mSelectedRows.Columns.Contains("ItemID") Then
            If IsDBNull(mSelectedRows.Fields("ItemID")) Then ID = "" Else ID = mSelectedRows.Fields("ItemID").ToString
        Else
            If mSelectedRows.Fields.ItemArray.Length = 0 Or IsDBNull(mSelectedRows.Fields(0)) Then ID = "" Else ID = mSelectedRows.Fields(0).ToString
        End If
        mSelectedRows.MoveNext()

        ReadNextBool = True
    End Function

    Public Sub LimitToListOfPKs(ByVal PrimaryKeys As String)
        If Not ActiveSelect() Then Return
        PrimaryKeys = Chr(254) & LCase(PrimaryKeys) & Chr(254)

        Dim Rows As DataRowCollection = mSelectedRows.Table.Rows, RowI As Integer = 0, hasItemId As Boolean = mSelectedRows.Columns.Contains("ItemID")

        Do While RowI < Rows.Count
            Dim ItemId As String = ""
            Dim Row As DataRow = Rows(RowI)

            If hasItemId Then
                If Not IsDBNull(Row("ItemID")) Then ItemId = Row("ItemID").ToString
            Else
                If Not IsDBNull(Row(0)) Then ItemId = Row(0).ToString
            End If

            If VB6.InStr(PrimaryKeys, Chr(254) & LCase(ItemId) & Chr(254)) = 0 Then
                Rows.RemoveAt(RowI)
            Else
                RowI = RowI + 1
            End If
        Loop
    End Sub

    ' Return only PrimaryKeys
    Public Function GetListOfPKs() As String
        Dim builder As New StringBuilder, firstTime = True, cma = "", absolutePosition As Integer
        If Not ActiveSelect() Then Return ""
        absolutePosition = mSelectedRows.abosolutePosition
        mSelectedRows.MoveFirst()
        Do While Not mSelectedRows.EOF

            If mSelectedRows.Columns.Contains("ItemID") Then
                If IsDBNull(mSelectedRows.Fields("ItemID")) Then builder.Append(cma) Else builder.Append(cma & mSelectedRows.Fields("ItemID").ToString)
            Else
                If IsDBNull(mSelectedRows.Fields(0)) Then builder.Append(cma) Else builder.Append(cma & mSelectedRows.Fields(0).ToString)
            End If

            mSelectedRows.MoveNext()
            If firstTime Then
                cma = Chr(254)
                firstTime = False
            End If
        Loop
        mSelectedRows.abosolutePosition = absolutePosition
        Return builder.ToString
    End Function

    Public Function GetJSONString() As String
        If OnlyReturnItemIDs Then Return GetListOfPKs()
        If Not ActiveSelect() Then Return "[]"

        Dim dt As DataTable = mSelectedRows.Table

        ' Return an array of JSON structure from a flat file?
        If dt.Columns.Count = 2 AndAlso dt.Rows.Count > 0 AndAlso dt.Columns(0).Caption = "ItemID" AndAlso dt.Columns(1).Caption = "ItemContent" Then
            If Not IsDBNull(dt.Rows(0)(1)) AndAlso VB6.Left(dt.Rows(0)(1), 1) = "{" Then
                Try
                    Dim jObj As Object = JSON(dt.Rows(0)(1))
                    Dim Sb1 As New StringBuilder()
                    Sb1.Append("[")
                    For i As Integer = 0 To dt.Rows.Count - 1
                        If i > 0 Then Sb1.Append(",")
                        Sb1.Append(dt.Rows(i)(1))
                    Next
                    Sb1.Append("]")
                    Return Sb1.ToString
                Catch ex As Exception
                End Try
            End If
        End If

        ' Return a JSON structure from a SQL Table

        Dim Sb As New StringBuilder()
        Dim StrDc As String() = New String(dt.Columns.Count - 1) {}

        For i As Integer = 0 To dt.Columns.Count - 1
            StrDc(i) = JSonEncodeString(dt.Columns(i).Caption) & ":"
        Next
        Sb.Append("[")
        For i As Integer = 0 To dt.Rows.Count - 1
            If i > 0 Then Sb.Append(",")
            Sb.Append("{")
            For j As Integer = 0 To dt.Columns.Count - 1
                If j > 0 Then Sb.Append(",")
                Dim DataType As System.Type = dt.Columns(j).DataType

                If IsNothing(dt.Rows(i)(j)) OrElse IsDBNull(dt.Rows(i)(j)) Then
                    Sb.Append(StrDc(j) & "null")

                ElseIf DataType Is GetType(System.Boolean) Then
                    If dt.Rows(i)(j) Then Sb.Append(StrDc(j) & "true") Else Sb.Append(StrDc(j) & "false")

                ElseIf NumericFld(DataType) Then
                    Sb.Append(StrDc(j) & dt.Rows(i)(j).ToString())

                Else
                    Sb.Append(StrDc(j) & JSonEncodeString(dt.Rows(i)(j).ToString()))
                End If
            Next
            Sb.Append("}")
        Next

        Sb.Append("]")
        Return Sb.ToString
    End Function
End Class


'Public Class clsTableHandleExchangeServer
'   Inherits clsTableHandle

'   Dim ExchangeServer As ExchangeService
'   Dim mColumnnames As Collection = New Collection()
'   Dim mTableName As String = ""

'   Public Overrides Function dbOpenBool(ByVal dbAccount As clsjdbAccount, ByVal TableName As String, ByRef Errors As String, Optional ByVal CreateIt As Boolean = False) As Boolean
'      mTableName = TableName

'      MyDBAccount = dbAccount
'      mTableName = TableName

'      mColumnnames = New Collection
'      mColumnnames.Add("ItemContent")
'      ExchangeServer = MyDBAccount.ExchangeServer

'      If VB6.LCase(TableName) = "inbox" Then Return True
'      If VB6.LCase(TableName) = "outbox" Then Return True

'      Errors = TableName & " not found.  Use inbox or outbox"
'      Return False
'   End Function

'   Public Overrides Function dbReadBool(ByVal ItemID As String, ByRef Result As String, Optional ByVal IsReadU As Boolean = False) As Boolean
'      Dim JRec As New jsonObject
'      If Not dbReadJSon(ItemID, JRec, IsReadU) Then Return False
'      Result = JSON_Obj2Str(JRec)
'      Return True
'   End Function

'   Public Overrides Function dbReadJSon(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
'      Dim MsgId = New ItemId(ItemID)
'      Dim MailMsg As Microsoft.Exchange.WebServices.Data.EmailMessage

'      If VB6.LCase(mTableName) = "outbox" Then Return False

'      Try
'         MailMsg = Item.Bind(ExchangeServer, MsgId)
'      Catch ex As Exception
'         If DirectCast(ex, Microsoft.Exchange.WebServices.Data.ServiceResponseException).ErrorCode <> ServiceError.ErrorInvalidId Then Throw ex
'         Return False
'      End Try

'      Dim JRec As New jsonObject
'      JRec.Add("Sender", MailMsg.Sender)
'      JRec.Add("ReceivedBy", MailMsg.ReceivedBy)
'      JRec.Add("BccRecipients", GetJSonValue(MailMsg.BccRecipients, 0))
'      JRec.Add("CcRecipients", GetJSonValue(MailMsg.CcRecipients, 0))
'      JRec.Add("Subject", MailMsg.Subject)
'      JRec.Add("DateTimeReceived", MailMsg.DateTimeReceived)

'      JRec.Add("DisplayTo", MailMsg.DisplayTo)
'      JRec.Add("Importance", MailMsg.Importance)
'      JRec.Add("InReplyTo", MailMsg.InReplyTo)
'      JRec.Add("IsRead", MailMsg.IsRead)
'      JRec.Add("HasAttachments", MailMsg.HasAttachments)
'      JRec.Add("DisplayCc", MailMsg.DisplayCc)
'      JRec.Add("IsNew", MailMsg.IsNew)

'      JRec.Add("BodyType", MailMsg.Body.BodyType.ToString)
'      JRec.Add("Body", MailMsg.Body.ToString())

'      Result = JRec
'      Return True
'   End Function

'   Public Overrides Function dbReadXML(ByVal ItemID As String, ByRef Result As Object, Optional ByVal IsReadU As Boolean = False) As Boolean
'      Throw New Exception("ExchangeServer does not support XML")
'      Return False
'   End Function

'   ' ItemID is TO
'   ' Item<1> is subject
'   ' Item<2> is CcRecipients ;
'   ' Item<3> is 
'   Public Overrides Sub dbWrite(ByVal Item As String, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
'      Dim JItem As New System.Collections.Generic.Dictionary(Of String, Object)

'      JItem.Add("To", ItemID)
'      JItem.Add("Subject", Extract(Item, 1))
'      JItem.Add("Cc", Extract(Item, 2))
'      JItem.Add("Bcc", Extract(Item, 3))
'      JItem.Add("BodyType", Extract(Item, 4))
'      JItem.Add("Body", Extract(Item, 5))


'      dbWriteJSon(JItem, ItemID, IsWriteU)
'   End Sub

'   Sub ParseEmailRecipients(ByVal R As EmailAddressCollection, ByVal S As String)
'      Dim D As String = ";"
'      If VB6.InStr(S, Chr(253)) Then D = Chr(253) Else If VB6.InStr(S, ";") Then D = ";" Else If VB6.InStr(S, ",") Then D = ","
'      For Each Adr As String In VB6.Split(S, D)
'         If VB6.InStr(Adr, "<") And VB6.InStr(Adr, ">") Then
'            Dim Name As String = Field(Adr, ">", 1)
'            Dim EMail As String = Field(Adr, ">", 2)
'            R.Add(Name, EMail)
'         Else
'            R.Add(Adr)
'         End If
'      Next
'   End Sub

'   Sub ParseEmailAttachments(ByVal A As AttachmentCollection, ByVal S As String)
'      Dim D As String = ";"
'      If VB6.InStr(S, Chr(253)) Then D = Chr(253) Else If VB6.InStr(S, ";") Then D = ";" Else If VB6.InStr(S, ",") Then D = ","
'      For Each f As String In VB6.Split(S, D)
'         If VB6.InStr(f, "<") And VB6.InStr(f, ">") Then
'            Dim Name As String = Field(f, ">", 1)
'            Dim fileName As String = Field(f, ">", 2)
'            A.AddFileAttachment(Name, fileName).IsInline = True
'         Else
'            A.AddFileAttachment(f).IsInline = True
'         End If
'      Next
'   End Sub

'   Public Overrides Sub dbWriteJSon(ByVal JItem As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
'      Dim Body As String = ""

'      If VB6.LCase(mTableName) = "inbox" Then Throw New Exception("Write to outbox not inbox")
'      If Not IsJsonObj(JItem) Then Throw New Exception("WriteJSON Item is not a JSON type")

'      ' Create an email message and provide it with connection 
'      ' configuration information by using an ExchangeService object named service.
'      Dim NewMessage As New EmailMessage(ExchangeServer)

'      ' Set properties on the email message.
'      If JItem.ContainsKey("To") AndAlso JItem("To") <> "" Then ParseEmailRecipients(NewMessage.ToRecipients, JItem("To"))
'      If JItem.ContainsKey("Subject") AndAlso JItem("Subject") <> "" Then NewMessage.Subject = JItem("Subject")
'      If JItem.ContainsKey("Body") AndAlso JItem("Body") <> "" Then Body = JItem("Body")
'      NewMessage.Body = Body

'      If JItem.ContainsKey("Cc") AndAlso JItem("Cc") <> "" Then ParseEmailRecipients(NewMessage.CcRecipients, JItem("Cc"))
'      If JItem.ContainsKey("Bcc") AndAlso JItem("Bcc") <> "" Then ParseEmailRecipients(NewMessage.BccRecipients, JItem("Bcc"))

'      If JItem.ContainsKey("BodyType") AndAlso JItem("BodyType") <> "" Then
'         If VB6.LCase(JItem("BodyType")) = "html" Then NewMessage.Body.BodyType = BodyType.HTML Else NewMessage.Body.BodyType = BodyType.Text
'      Else
'         If VB6.InStr(Body, "<") AndAlso VB6.InStr(Body, ">") Then NewMessage.Body.BodyType = BodyType.HTML Else NewMessage.Body.BodyType = BodyType.Text
'      End If

'      If JItem.ContainsKey("Attachments") AndAlso JItem("Attachments") <> "" Then ParseEmailAttachments(NewMessage.Attachments, JItem("Attachments"))

'      ' Send the email message and save a copy.
'      NewMessage.SendAndSaveCopy()
'   End Sub

'   Public Overrides Sub dbWriteXML(ByVal Item As Object, ByVal ItemID As String, Optional ByVal IsWriteU As Boolean = False)
'      If IsXmlObj(Item) Then Item = XML_Obj2Str(Item)
'   End Sub

'   Public Overrides Sub dbTransactionBegin()
'   End Sub

'   Public Overrides Sub dbTransactionCommit()
'   End Sub

'   Public Overrides Sub dbExitServer()
'   End Sub

'   Public Overrides Sub ClearFile()
'      Throw New Exception("Exchange Server clearfile")
'   End Sub

'   Public Overrides Sub dbDeleteFile()
'      Throw New Exception("Exchange Server can't deletefile")
'   End Sub

'   Public Overrides Sub dbDelete(ByVal ItemID As String)
'      Dim MsgId = New ItemId(ItemID)
'      Dim MailMsg As Microsoft.Exchange.WebServices.Data.EmailMessage

'      Try
'         MailMsg = Item.Bind(ExchangeServer, MsgId)
'      Catch ex As Exception
'         Return
'      End Try
'      MailMsg.Delete(DeleteMode.MoveToDeletedItems)
'   End Sub

'   ' Return 
'   Public Overrides Function SelectFile() As rSelectList
'      Return SelectFileX("", "")
'   End Function

'   Public Overrides Function SelectFileX(ByVal ColumnList As String, ByVal WhereClause As String) As rSelectList
'      Dim FormList As String = ""
'      Dim findResults As FindItemsResults(Of Item)
'      Dim returnOnlyItemIDs As Boolean = ColumnList = "" AndAlso WhereClause = ""
'      Dim Top As Boolean = False
'      Dim Bottom As Boolean = False
'      Dim TopBottomCnt As Long = 0

'      If WhereClause <> "" And False Then

'         Dim SearchFilterCollection As New List(Of SearchFilter)
'         SearchFilterCollection.Add(New SearchFilter.IsEqualTo(ItemSchema.Id, "xxx")) ' also substring, less, greater, etc;

'         ' Create the search filter.
'         Dim searchFilter As SearchFilter = New SearchFilter.SearchFilterCollection(LogicalOperator.Or, SearchFilterCollection.ToArray())

'         Dim view As New ItemView(1)
'         ' Order the search results by the DateTimeReceived in descending order.
'         view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending)

'         ' Set the traversal to shallow. (Shallow is the default option; other options are Associated and SoftDeleted.)
'         view.Traversal = ItemTraversal.Shallow


'         findResults = ExchangeServer.FindItems(WellKnownFolderName.Inbox, searchFilter, view)
'      Else
'         findResults = ExchangeServer.FindItems(WellKnownFolderName.Inbox, New ItemView(30)) '  1000 is the number of mails to fetch
'      End If

'      ColumnList = VB6.LTrim(ColumnList)
'      Dim token As String = VB6.LCase(Field(ColumnList, " ", 1))

'      Top = token = "top"
'      Bottom = token = "bottom"

'      If Top Or Bottom Then
'         ColumnList = VB6.Mid(ColumnList, Len(token) + 1)
'         ColumnList = VB6.LTrim(ColumnList)
'         Dim TNum As String = Field(ColumnList, " ", 1)
'         ColumnList = VB6.Mid(ColumnList, Len(TNum) + 1)
'         TopBottomCnt = Val(TNum)
'      End If

'      Dim A As New ArrayList
'      For Each MailMsg As Microsoft.Exchange.WebServices.Data.EmailMessage In findResults.Items
'         If returnOnlyItemIDs Then
'            FormList = FormList & Chr(254) & MailMsg.Id.ToString
'         Else
'            'this needs to be here to receive the message body
'            'Dim messageBody As MessageBody = New Microsoft.Exchange.WebServices.Data.MessageBody()
'            'Dim items As New List(Of Item)()
'            'If findResults.Items.Count > 0 Then
'            '   ' Prevent the exception
'            '   For Each item2 As Item In findResults
'            '      items.Add(item2)
'            '   Next
'            'End If
'            'ExchangeServer.LoadPropertiesForItems(items, PropertySet.FirstClassProperties)

'            Try
'               ' Dim JRec As jsonObject = JSON_Obj2Json(MailMsg)
'               Dim JRec As New jsonObject
'               JRec.Add("ItemID", MailMsg.Id)
'               JRec.Add("IsNew", MailMsg.IsNew)
'               JRec.Add("DisplayTo", MailMsg.DisplayTo)
'               JRec.Add("Subject", MailMsg.Subject)
'               ' JRec.Add("Body", MailMsg.Body.ToString())
'               JRec.Add("DateTimeReceived", MailMsg.DateTimeReceived)
'               JRec.Add("BccRecipients", GetJSonValue(MailMsg.BccRecipients, 0))
'               JRec.Add("CcRecipients", GetJSonValue(MailMsg.CcRecipients, 0))
'               JRec.Add("Importance", MailMsg.Importance)
'               JRec.Add("InReplyTo", MailMsg.InReplyTo)
'               JRec.Add("IsRead", MailMsg.IsRead)
'               JRec.Add("HasAttachments", MailMsg.HasAttachments)
'               JRec.Add("DisplayCc", MailMsg.DisplayCc)
'               JRec.Add("ReceivedBy", MailMsg.ReceivedBy)
'               JRec.Add("Sender", MailMsg.Sender)
'               A.Add(JRec)

'            Catch ex As Exception

'            End Try

'         End If
'      Next

'      Dim RS As New rSelectList
'      If returnOnlyItemIDs Then
'         FormList = VB6.Mid(FormList, 2)
'         RS.SetDynamicArray(FormList, True)

'      Else
'         ' By making this a data table, we can Filter and Select the results
'         Dim Errors As String = ""
'         Dim DT As DataTable = Json2DataTable(A, "Exchange", Errors)
'         If Errors <> "" Then Throw New Exception(Errors)

'         RS.SetDataTable(DT, returnOnlyItemIDs)
'         If WhereClause <> "" Then RS.FilterList(WhereClause)
'      End If

'      ' Get only Top of Bot?
'      If Top Then

'      End If
'      Return RS
'   End Function

'   Public Overrides ReadOnly Property TableName() As String
'      Get
'         TableName = mTableName
'      End Get
'   End Property

'   Public Overrides ReadOnly Property ColumnNames() As Collection
'      Get
'         Return mColumnnames
'      End Get
'   End Property

'   Public Overrides ReadOnly Property ColumnPosition(ByVal ColumnName As String) As Integer
'      Get
'         For I As Integer = 1 To mColumnnames.Count
'            If mColumnnames(I) = ColumnName Then Return I
'         Next
'         Return 0
'      End Get
'   End Property

'   Public Overrides ReadOnly Property PrimaryKeyColumnName() As String
'      Get
'         PrimaryKeyColumnName = "ItemID"
'      End Get
'   End Property

'   Public Function FilePath() As String
'      FilePath = mTableName
'   End Function


'End Class