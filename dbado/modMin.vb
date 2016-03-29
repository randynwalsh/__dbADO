Imports System.Security.Cryptography
Imports System.Text
Imports System.Collections.Specialized
Imports System.Web.Script.Serialization
Imports System.Xml
Imports System.IO
Imports System.Net
Imports VB6 = Microsoft.VisualBasic

Public Module modMain
   Public SessionCookies As New CookieContainer
   Public FullUrl As String = ""
   Public mMapPath As String = "" ' C:\Users\rwalsh\Dropbox\_src\__HTA\
   Public mAccountName As String = ""
   Public allowAllDosWritesChecked As Boolean = False
   Public allowAllDosWritesValue As Boolean = False
   Public allowAllDosReadsChecked As Boolean = False
   Public allowAllDosReadsValue As Boolean = False


   Private Const adEmpty As Integer = 0
   Private Const adTinyInt As Integer = 16
   Private Const adSmallInt As Integer = 2
   Private Const adInteger As Integer = 3
   Private Const adBigInt As Integer = 20
   Private Const adUnsignedTinyInt As Integer = 17
   Private Const adUnsignedSmallInt As Integer = 18
   Private Const adUnsignedInt As Integer = 19
   Private Const adUnsignedBigInt As Integer = 24
   Private Const adSingle As Integer = 4
   Private Const adDouble As Integer = 5
   Private Const adCurrency As Integer = 6
   Private Const adDecimal As Integer = 14
   Private Const adNumeric As Integer = 131
   Private Const adBoolean As Integer = 11
   Private Const adError As Integer = 10
   Private Const adUserDefined As Integer = 131
   Private Const adVariant As Integer = 12
   Private Const adIDispatch As Integer = 9
   Private Const adIUnknown As Integer = 13
   Private Const adGUID As Integer = 72
   Private Const adDate As Integer = 7
   Private Const adDBDate As Integer = 133
   Private Const adDBTime As Integer = 134
   Private Const adDBTimeStamp As Integer = 135
   Private Const adBSTR As Integer = 8
   Private Const adChar As Integer = 129
   Private Const adVarChar As Integer = 200
   Private Const adLongVarChar As Integer = 201
   Private Const adWChar As Integer = 130
   Private Const adVarWChar As Integer = 202
   Private Const adLongVarWChar As Integer = 203
   Private Const adBinary As Integer = 128
   Private Const adVarBinary As Integer = 204
   Private Const adLongVarBinary As Integer = 205

   Public Function FullPath(ByVal FPath As String) As Boolean
      Return System.IO.Path.GetFullPath(FPath)
   End Function

   Public Function JsonChild(ByVal ObjectRef As Object, cname As String) As Object
      If IsJsonObj(ObjectRef) Then
         Dim O As System.Collections.Generic.Dictionary(Of System.String, System.Object) = ObjectRef
         If ObjectRef.ContainsKey(cname) Then Return ObjectRef(cname)
      End If
      Return Nothing
   End Function

   ' Parses a SqlStatement and gets the list of columns
   Function activeColumns(ByVal sqlWhere As String, ByVal toLowerCase As Boolean, ByRef hasItemID As Boolean, ByRef hasItemContent As Boolean, ByRef hasStarColumn As Boolean, ByRef AttributeNumsReferenced As Collection) As ArrayList
      '   sqlWhere = VB6.Replace(sqlWhere, "*a", "~", , , CompareMethod.Text)
      Dim tokens As ArrayList = SplitX(sqlWhere, " ", 5)
      Dim hasFunctions As Boolean, hasOperators As Boolean

      Dim results As New ArrayList
      For i As Integer = 0 To tokens.Count - 1
         Dim otoken As String = tokens(i)
         Dim token As String = LCase(otoken)
         Dim ntoken As String = ""
         Dim starIsAllColumns As Boolean = True

         If i < tokens.Count - 1 Then ntoken = LCase(tokens(i + 1))

         Select Case True
            Case token = "order" And ntoken = "by"
               i += 1

            Case token = "or" Or token = "and" Or token = "like" Or _
               token = "gt" Or token = "lt" Or token = "ge" Or token = "le" _
               Or token = "ne" Or token = "eq" Or token = "from" Or token = "where" _
               Or token = "select"

            Case Left(token, 2) = "*a"
               Dim AtrNo As Integer = Val(Mid(token, 3))

               If toLowerCase Then
                  If AttributeNumsReferenced.Contains(token) = False Then AttributeNumsReferenced.Add(token, token)
               Else
                  If AttributeNumsReferenced.Contains(otoken) = False Then AttributeNumsReferenced.Add(otoken, otoken)
               End If

            Case token = "*" And starIsAllColumns ' represents all columns
               hasStarColumn = True
               results.Add(token)

            Case token = "itemid"
               hasItemID = True
               If toLowerCase Then results.Add(token) Else results.Add("ItemID")

            Case token = "itemcontent"
               hasItemContent = True
               If toLowerCase Then results.Add(token) Else results.Add("ItemContent")

               ' special case of [columname]
            Case Left(token, 1) = "[" And Right(token, 1) = "]"
               otoken = Mid(otoken, 2, Len(otoken) - 2)
               token = LCase(otoken)
               If toLowerCase Then results.Add(token) Else results.Add(otoken)

            Case Left(token, 1) = """" And Right(token, 1) = """" And (InStr("=<>!", Left(ntoken, 1)) Or ntoken = "like")
               otoken = Mid(otoken, 2, Len(otoken) - 2)
               token = LCase(otoken)
               If toLowerCase Then results.Add(token) Else results.Add(otoken)

               ' not a function
            Case IsAlpha(Left(token, 1)) And InStr(token, "(")
               hasFunctions = True

            Case IsAlpha(Left(token, 1))
               If toLowerCase Then results.Add(token) Else results.Add(otoken)

            Case InStr("+-*/%", Left(token, 1))
               hasOperators = True

         End Select

         starIsAllColumns = token = ","
      Next
      Return results
   End Function

   Public Function AdoNumericFld(ByVal AdoDBFldType As Integer) As Boolean
      Select Case AdoDBFldType
         Case adUnsignedBigInt ' An 8-byte unsigned integer (DBTYPE_UI8).
            Return True
         Case adUnsignedInt ' A 4-byte unsigned integer (DBTYPE_UI4).
            Return True
         Case adUnsignedSmallInt ' A 2-byte unsigned integer (DBTYPE_UI2).
            Return True
         Case adUnsignedTinyInt ' A 1-byte unsigned integer (DBTYPE_UI1).
            Return True
         Case adBigInt ' An 8-byte signed integer (DBTYPE_I8).
            Return True
         Case adSmallInt ' A 2-byte signed integer (DBTYPE_I2).
            Return True
         Case adTinyInt ' A 1-byte signed integer (DBTYPE_I1).
            Return True
         Case adInteger ' A 4-byte signed integer (DBTYPE_I4).
            Return True
         Case adBoolean ' A Boolean value (DBTYPE_BOOL).
            Return True
         Case adNumeric ' An exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
            Return True
         Case adSingle ' A single-precision floating point value (DBTYPE_R4).
            Return True
         Case adDecimal ' An exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
            Return True
         Case adDouble ' A double-precision floating point value (DBTYPE_R8).
            Return True
         Case adCurrency ' A currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an 8-byte signed integer scaled by 10,000.
            Return True
      End Select
      Return False
   End Function

   Public Function AdoVarFld(ByVal AdoDBFldType As Integer) As Boolean
      Select Case AdoDBFldType
         Case adLongVarBinary
            Return True
         Case adLongVarChar
            Return True
         Case adLongVarWChar
            Return True
         Case adVarBinary
            Return True
         Case adVarChar
            Return True
         Case adVarWChar
            Return True
      End Select
      Return False
   End Function

   Public Function CreateColumn(ByVal iDBConnection As Object, ByRef oDBCatalog As Object, ByVal TableName As String, ByVal ColName As String, ByVal DefinedSize As Integer, ByVal dotNetType As String, ByVal AllowDBNull As Boolean, ByVal AutoIncrement As Boolean, ByVal AllowZeroLength As Boolean, ByVal DefaultValue As String, ByVal Description As String, ByRef Errors As String) As Boolean
      Dim Tbl As Object ' ADOX.Table
      Dim AdoXColumn As Object ' ADOX.Column

      If InStr(dotNetType, ".") Then dotNetType = Field(dotNetType, ".", 2)
      dotNetType = UCase(Left(dotNetType, 1)) & LCase(Mid(dotNetType, 2))

      ' Check paramenters for validity
      If TableName = "" Then Return False
      If ColName = "" Then Return False

      If oDBCatalog Is Nothing Then
         Try
            oDBCatalog = VB6.CreateObject("ADOX.Catalog")
            Dim CS As String = iDBConnection.ConnectionString ' "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\rwalsh\Dropbox\_src\__HTA\SFSR_2014_JMedits_December2014 (testsent).accdb;mode=19"

            Dim cnAdoConnection As Object = VB6.CreateObject("ADODB.Connection")
            cnAdoConnection.Open(CS)

            oDBCatalog.ActiveConnection = cnAdoConnection
         Catch ex As Exception
            Errors = ex.Message
            Return False
         End Try
      End If

      Try
         Tbl = oDBCatalog.Tables(TableName)
      Catch ex As Exception
         Errors = "Table not found. " & ex.Message
         Return False
      End Try

      ' Does column already exist?
      Try
         AdoXColumn = Tbl.Columns(ColName)
         Errors = "Column already exists"
         Return True
      Catch ex As Exception
      End Try

      ' New Column
      AdoXColumn = VB6.CreateObject("ADOX.Column")
      AdoXColumn.Name = ColName
      Try
         AdoXColumn.ParentCatalog = oDBCatalog
      Catch ex As Exception
         Errors = ex.Message
         Return False
      End Try

      If dotNetType <> "Boolean" Then
         If AllowDBNull Then
            AdoXColumn.Properties("Nullable").Value = True
         Else
            AdoXColumn.Properties("Nullable").Value = False
         End If
      End If

      If AutoIncrement Then
         AdoXColumn.Properties("Jet OLEDB:Allow Zero Length") = True

         If dotNetType = "Guid" Then
            AdoXColumn.Properties("Jet OLEDB:AutoGenerate") = True
         Else
            AdoXColumn.Properties("Autoincrement").Value = True
            AdoXColumn.Properties("Seed").Value = 1
            AdoXColumn.Properties("Increment").Value = 1
         End If

      End If

      Select Case dotNetType
         Case "Boolean"
            AdoXColumn.Type = adBoolean

         Case "Byte"
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adGUID

         Case "Guid"
            AdoXColumn.Properties("Jet OLEDB:Allow Zero Length") = True
            AdoXColumn.Type = adGUID

         Case "Char"
            AdoXColumn.Properties("Jet OLEDB:Compressed UniCode Strings").Value = True
            If DefinedSize > 255 Then
               ' use memo field if over 255
               AdoXColumn.Type = adLongVarWChar
            Else
               AdoXColumn.DefinedSize = DefinedSize
               AdoXColumn.Type = adVarWChar
            End If

         Case "Datetime"
            AdoXColumn.Type = adDate

         Case "Decimal"
            AdoXColumn.Precision = 18
            AdoXColumn.NumericScale = 4
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adNumeric

         Case "Double"
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adDouble

         Case "Int16"
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adSmallInt

         Case "Int32"
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adInteger

         Case "Int64"
            AdoXColumn.Properties("Default").Value = 0
            ' obvious problems with possible overflow
            AdoXColumn.Type = adInteger

         Case "Sbyte"
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adUnsignedTinyInt

         Case "Single"
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adSingle

         Case "String"
            AdoXColumn.Properties("Jet OLEDB:Compressed UniCode Strings").Value = True
            If DefinedSize > 255 Then
               ' use memo field if over 255
               AdoXColumn.Type = adLongVarWChar
            Else
               AdoXColumn.DefinedSize = DefinedSize
               AdoXColumn.Type = adVarWChar
            End If

         Case "Timespan"
            AdoXColumn.Type = adDate

         Case "Uint16"
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adSmallInt

         Case "Uint32"
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adInteger

         Case "Uint64"
            AdoXColumn.Properties("Default").Value = 0
            AdoXColumn.Type = adInteger

         Case Else
            AdoXColumn.Type = adBinary
      End Select

      If AllowZeroLength Then AdoXColumn.Properties("Jet OLEDB:Allow Zero Length").Value = AllowZeroLength

      If DefaultValue <> "" Then
         If AdoNumericFld(AdoXColumn.Type) Then
            AdoXColumn.Properties("Default").Value = Val(DefaultValue)
         ElseIf dotNetType = "Boolean" Then
            If DefaultValue = "True" Or DefaultValue = "1" Or DefaultValue = "-1" Then
               AdoXColumn.Properties("Default").Value = 1
            Else
               AdoXColumn.Properties("Default").Value = 0
            End If
         Else
            AdoXColumn.Properties("Default").Value = "'" & DefaultValue & "'"
         End If
      End If

      Try
         If Description <> "" Then AdoXColumn.Properties("Description").Value = "'" & DefaultValue & "'"
      Catch ex As Exception
      End Try

      Try
         Tbl.Columns.Append(AdoXColumn)
      Catch ex As Exception
         Try
            AdoXColumn.Name = "[" & AdoXColumn.Name & "]"  ' MS SQL bug
            Tbl.Columns.Append(AdoXColumn)
         Catch ex2 As Exception
            Errors = ex.Message
            Return False
         End Try
      End Try

      Return True
   End Function

   Public Function RenameColumn(ByVal IDbConnection As Object, ByRef oDBCatalog As Object, ByVal TableName As String, ByVal OldColName As String, ByVal NewColName As String, ByRef Errors As String) As Boolean
      Dim Tbl As Object ' ADOX.Table
      Dim AdoC As Object ' ADOX.Column

      ' Check paramenters for validity
      If TableName = "" Then Return False
      If OldColName = "" Then Return False


      If oDBCatalog Is Nothing Then
         Try
            oDBCatalog = VB6.CreateObject("ADOX.Catalog")
            Dim CS As String = IDbConnection.ConnectionString ' "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\rwalsh\Dropbox\_src\__HTA\SFSR_2014_JMedits_December2014 (testsent).accdb;mode=19"

            Dim cnAdoConnection As Object = VB6.CreateObject("ADODB.Connection")
            cnAdoConnection.Open(CS)

            oDBCatalog.ActiveConnection = cnAdoConnection
         Catch ex As Exception
            Errors = ex.Message
            Return False
         End Try
      End If

      ' Does table exist?
      Try
         Tbl = oDBCatalog.Tables(TableName)
      Catch ex As Exception
         Errors = ex.Message
         Return False
      End Try

      ' Does column exist?
      Try
         AdoC = Tbl.Columns(OldColName)
      Catch ex As Exception
         Errors = ex.Message
         Return False
      End Try


      ' New Column Name
      Try
         If AdoC.Name <> NewColName And NewColName <> "" Then AdoC.Name = NewColName
      Catch ex As Exception
         Errors = ex.Message
         Return False
      End Try

      Return True
   End Function

   Public Function DeleteColumn(ByVal IDbConnection As Object, ByRef oDBCatalog As Object, ByVal TableName As String, ByVal ColName As String, ByRef Errors As String) As Boolean
      Dim Tbl As Object ' ADOX.Table

      ' Check paramenters for validity
      If TableName = "" Then Return False
      If ColName = "" Then Return False

      If oDBCatalog Is Nothing Then
         Try
            oDBCatalog = VB6.CreateObject("ADOX.Catalog")
            Dim CS As String = IDbConnection.ConnectionString ' "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\rwalsh\Dropbox\_src\__HTA\SFSR_2014_JMedits_December2014 (testsent).accdb;mode=19"

            Dim cnAdoConnection As Object = VB6.CreateObject("ADODB.Connection")
            cnAdoConnection.Open(CS)

            oDBCatalog.ActiveConnection = cnAdoConnection
         Catch ex As Exception
            Errors = ex.Message
            Return False
         End Try
      End If

      ' Does table exist?
      Try
         Tbl = oDBCatalog.Tables(TableName)
      Catch ex As Exception
         Errors = ex.Message
         Return False
      End Try

      ' New Column Name
      Try
         Tbl.Columns.Delete(ColName)
      Catch ex As Exception
         Errors = ex.Message
         Return False
      End Try

      Return True
   End Function

   Public Function IsAlphaNum(ByVal strInputText As String) As Boolean
      Dim IsAlpha As Boolean = False
      If System.Text.RegularExpressions.Regex.IsMatch(strInputText, "^[a-zA-Z0-9]+$") Then
         IsAlpha = True
      Else
         IsAlpha = False
      End If
      Return IsAlpha
   End Function

   Public Function IsAlpha(ByRef Value As Object) As Boolean
      ' Use the ALPHA public function to determine whether expression is an
      ' alphabetic or nonalphabetic string.
      '
      ' Expr: If Expr contains only the Characters a through z or A through Z, Expr Value of 1 returns.
      '       If Expr contains any other Character or is an Empty Value, Expr Value of zero returns.
      '       If expression is the Null Value, Null returns.
      '
      Dim I As Object
      Dim L As Integer
      Dim C As String
      Dim EXPR As String

      EXPR = Value

      L = Len(EXPR)
      For I = 1 To L
         C = UCase(Mid(EXPR, I, 1))
         If C < "A" Or C > "Z" Then Return False
      Next I

      ' Empty string returns 0 (False)
      Return L > 0
   End Function

   Private Function isFileUnicode(ByVal path As String) As Boolean
      Dim enc As System.Text.Encoding = Nothing
      Dim file As System.IO.FileStream = New System.IO.FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)
      If file.CanSeek Then
         Dim bom As Byte() = New Byte(3) {} ' Get the byte-order mark, if there is one
         file.Read(bom, 0, 4)
         If (bom(0) = &HEF AndAlso bom(1) = &HBB AndAlso bom(2) = &HBF) OrElse (bom(0) = &HFF AndAlso bom(1) = &HFE) OrElse (bom(0) = &HFE AndAlso bom(1) = &HFF) OrElse (bom(0) = 0 AndAlso bom(1) = 0 AndAlso bom(2) = &HFE AndAlso bom(3) = &HFF) Then ' ucs-4
            Return True
         Else
            Return False
         End If

         ' Now reposition the file cursor back to the start of the file
         file.Seek(0, System.IO.SeekOrigin.Begin)
      Else
         Return False
      End If
   End Function

   Public Function MapPath(ByVal relPath As String) As String
      If Left(relPath, 1) = "\" Or Mid(relPath, 2, 1) = ":" Then Return VB6.Replace(relPath, "/", "\")
      If Left(relPath, 1) = "/" Then relPath = Mid(relPath, 2)
      If Left(relPath, 2) = "~/" Then relPath = Mid(relPath, 3)
      Return mMapPath & VB6.Replace(relPath, "/", "\")
   End Function

   Public Function MapRootPath(ByVal relPath As String) As String
      If Left(relPath, 1) = "\" Or Mid(relPath, 2, 1) = ":" Then Return VB6.Replace(relPath, "/", "\")
      Return mMapPath & VB6.Replace(relPath, "/", "\")
   End Function

   Public Function MapRootAccount(ByVal AccountName As String) As String
      If AccountName = "" Then AccountName = UrlAccount()
      If Left(AccountName, 1) = "." Then Return mMapPath
      Return MapRootPath("App_Data\_database\" & AccountName & "\")
   End Function


   ' Name of HTA file
   Public Function UrlAccount() As String
      Return mAccountName
   End Function

   Function IsOdd(ByVal I As Integer) As Boolean
      Return (I Mod 2) = 1
   End Function

   Function IsEven(ByVal I As Integer) As Boolean
      Return (I Mod 2) = 0
   End Function


   Function AesEncrypt(Str As String, Optional ByVal Key As String = "") As String
      If Str = "" Then Return ""

      Dim myAes As New RijndaelManaged '  AesManaged
      myAes.Mode = CipherMode.CFB
      myAes.IV = New Byte(15) {&H77, &H81, &H15, &H7E, &H26, &H29, &HB0, &H94, &HF0, &HE3, &HDD, &H48, &HC4, &HD7, &H86, &H11}
      myAes.Padding = PaddingMode.Zeros

      Dim K() As Byte = strToBytes(Key & "32o4908go293hohg98fh40gh")
      ReDim Preserve K(15)
      myAes.Key = K

      Dim inBlock As Byte() = strToBytes(Str)
      Dim xfrm As ICryptoTransform = myAes.CreateEncryptor()
      Dim outBlock() As Byte = xfrm.TransformFinalBlock(inBlock, 0, inBlock.Length)
      ReDim Preserve outBlock(inBlock.Length - 1)
      Return bytesToString(outBlock)
   End Function

   Function AesDecrypt(Str As String, Optional ByVal Key As String = "") As String
      If Str = "" Then Return ""

      Dim myAes As New RijndaelManaged '  AesManaged
      myAes.Mode = CipherMode.CFB
      myAes.IV = New Byte(15) {&H77, &H81, &H15, &H7E, &H26, &H29, &HB0, &H94, &HF0, &HE3, &HDD, &H48, &HC4, &HD7, &H86, &H11}
      myAes.Padding = PaddingMode.Zeros

      Dim K() As Byte = strToBytes(Key & "32o4908go293hohg98fh40gh")
      ReDim Preserve K(15)
      myAes.Key = K

      Dim inBlock As Byte() = strToBytes(Str)
      Dim OrginalLen As Integer = inBlock.Length
      Dim AddBlanks As Integer = inBlock.Length + 16 - (inBlock.Length Mod 16)
      ReDim Preserve inBlock(AddBlanks - 1)
      Dim xfrm As ICryptoTransform = myAes.CreateDecryptor
      Dim outBlock() As Byte = xfrm.TransformFinalBlock(inBlock, 0, inBlock.Length)
      ReDim Preserve outBlock(OrginalLen - 1)
      Return bytesToString(outBlock)
   End Function

   Function strToBytes(s As String) As Byte()
      Dim B(Len(s) - 1) As Byte
      For I As Integer = 0 To Len(s) - 1
         B(I) = Asc(Mid(s, I + 1, 1))
      Next
      Return B
   End Function

   Function bytesToString(b As Byte()) As String
      Dim S As New StringBuilder
      For i As Integer = 0 To UBound(b)
         S.Append(Chr(b(i)))
      Next
      Return S.ToString
   End Function


   Public Function CNum(ByRef vValue As Object) As Object
      If IsDBNull(vValue) OrElse IsNothing(vValue) Then Return 0

      Select Case vValue.GetType
         Case GetType(Boolean)
            If vValue Then Return 1 Else Return 0

         Case GetType(Byte), GetType(UShort), GetType(UInteger), GetType(ULong), GetType(SByte), GetType(Short), _
                     GetType(Integer), GetType(Long), GetType(Decimal), GetType(Single), GetType(Double)
            Return vValue

         Case GetType(String)
            Return Val(vValue)

         Case GetType(Date)
            Return False

         Case GetType(Char)
            If vValue >= "0" AndAlso vValue <= "9" Then Return CByte(vValue)

         Case GetType(Object)
            Return Val(C2Str(vValue))
      End Select

      Return 0
   End Function

   Public Function C2str2(ByVal O As Object, Del As Integer) As String
      If TypeOf O Is Collection OrElse TypeOf O Is NameObjectCollectionBase.KeysCollection Or IsArray(O) Or TypeOf O Is ArrayList Then
         Dim S As New StringBuilder, Firstcall As Boolean = True, subDel As Integer = Del - 1
         If Del = Asc(",") Then
            S.Append("[")
            subDel = Del
         End If

         For Each E As Object In O
            S.Append(IIf(Firstcall, "", Chr(Del)) & C2str2(E, subDel))
            Firstcall = False
         Next
         If Del = Asc(",") Then S.Append("]")
         Return S.ToString
      Else
         Return C2Str(O)
      End If
   End Function

   Public Function IsJsonObj(ByVal j As Object) As Boolean
      Return TypeOf j Is System.Collections.Generic.Dictionary(Of System.String, System.Object)
   End Function

   Public Function JSON_Obj2Str(ByVal Obj As Object) As String
      Return (New System.Web.Script.Serialization.JavaScriptSerializer).Serialize(Obj)
   End Function


   ' Use the INDEX public function to return the starting Character position for the specified occurrence of substring in string.
   ' SearchStr:    is any valid string, and is examined for the substring expression.
   ' Cnt:          specifies whiCh occurrence of substring is to be located.
   '
   ' When substring is found and ff it meets the occurrence criterion, the
   ' starting Character position of the substring returns.  If substring is an
   ' empty string, 1 returns.  If the specified occurrence of the substring is
   ' not found, or If string or substring evaluate to the null Value, zero returns.
   '

   Public Function Index(ByVal SearchStr As String, ByVal Delimiter As String, Optional ByVal Cnt As Integer = 1) As Integer
      Dim S As Integer

      If Cnt <= 0 Then Return 0
      If Delimiter = "" Then Return 1

      Do
         S = InStr(S + 1, SearchStr, Delimiter)
         If S = 0 Then Return 0
         Cnt = Cnt - 1
      Loop While Cnt > 0

      Return S
   End Function

   Public Function Field(ByRef EXPR As String, ByRef Delimiter As String, ByRef Occurrence As Integer, Optional ByRef Col1 As Object = Nothing, Optional ByRef Col2 As Object = Nothing) As String
      ' Use the Field public function to return a substring located between specified Delimiters in string.
      '
      ' Delimiter is any Character, including field mark, Value mark, and
      ' subValue mark. It Delimits the start and end of the substring.  If
      ' Delimiter equals more than one Character, only the first Character is
      ' used.  Delimiters are not returned with the substring.
      '
      ' Occurrence specifies whiCh occurrence of the Delimiter is to be used as
      ' a terminator.  If occurrence is less than one, one is assumed.
      Dim SI As Object
      Dim L As Integer

      If IsDBNull(EXPR) Or EXPR = "" Or Occurrence = 0 Then
         Col1 = 0
         Col2 = 1
         Return Nothing
      End If

      If Occurrence > 1 Then
         SI = Index(EXPR, Delimiter, Occurrence - 1)
         If SI = 0 Then
            Col1 = 0
            Col2 = 0
            Return Nothing
         End If
         Col1 = SI
         SI = SI + Len(Delimiter)
      Else
         SI = 1
         Col1 = 0
      End If

      Col2 = InStr(SI, EXPR, Delimiter)
      If Col2 = 0 Then Col2 = Len(EXPR) + 1

      L = Col2 - SI
      Field = Mid(EXPR, SI, L)
   End Function

   ' CountDel (Count) returns the number of bUnique times a Del is repeated in a string value
   ' If Str is an empty string, the lenght of P is returned
   Public Function CountDel(ByVal str As String, ByVal Del As String) As Integer
      Dim S As Integer = 0, Cnt As Integer = 0

      If Del = "" Then Return Len(str)
      Do
         S = InStr(S + 1, str, Del)
         If S = 0 Then Return Cnt
         Cnt = Cnt + 1
      Loop
   End Function

   Public Function DCount(ByVal str As String, ByVal Del As String) As Integer
      If str = "" Then Return 0 Else Return CountDel(str, Del) + 1
   End Function

   ' Returns the path portion of a pathname
   Public Function GetFolder(ByVal SFile As String) As String
      For lCount As Integer = Len(SFile) To 1 Step -1
         If Mid$(SFile, lCount, 1) = "\" Or Mid$(SFile, lCount, 1) = "/" Then Return Left$(SFile, lCount)
      Next
      Return vbNullString
   End Function

   Public Function GetExtension(ByVal SFile As String) As String
      Dim lCount As Integer
      For lCount = Len(SFile) To 2 Step -1
         If InStr(":\/", Mid$(SFile, lCount, 1)) > 0 Then Exit For
         If Mid$(SFile, lCount, 1) = "." Then Return Mid$(SFile, lCount + 1)
      Next
      Return ""
   End Function

   Public Function DropExtension(ByVal SFile As String) As String
      Dim lCount As Integer

      For lCount = Len(SFile) To 2 Step -1
         If InStr(":\/", Mid$(SFile, lCount, 1)) > 0 Then Exit For
         If Mid$(SFile, lCount, 1) = "." Then Return Left$(SFile, lCount - 1)
      Next

      Return SFile
   End Function

   Public Function JSON_PrettyFormat(inputText As String) As String
      Dim escaped As Boolean = False
      Dim inquotes As Boolean = False
      Dim column As Integer = 0
      Dim indentation As Integer = 0
      Dim indentations As New Stack(Of Integer)()
      Dim sb As New StringBuilder()
      For Each x As Char In inputText
         sb.Append(x)
         column += 1
         If escaped Then
            escaped = False
         Else
            If x = "\" Then
               escaped = True
            ElseIf x = """" Then
               inquotes = Not inquotes
            ElseIf Not inquotes Then
               If x = "," Then
                  ' if we see a comma, go to next line, and indent to the same depth
                  sb.Append(vbCrLf)
                  column = 0
                  For i As Integer = 0 To indentation - 1
                     sb.Append(" ")
                     column += 1
                  Next
               ElseIf x = "[" OrElse x = "{" Then
                  ' if we open a bracket or brace, indent further (push on stack)
                  indentations.Push(indentation)
                  indentation = column
               ElseIf x = "]" OrElse x = "}" Then
                  ' if we close a bracket or brace, undo one level of indent (pop)
                  indentation = indentations.Pop()
               ElseIf x = ":" Then
                  sb.Append(" ")
                  column += 1
               End If
            End If
         End If
      Next
      Return sb.ToString()
   End Function

   Function JSonEncodeString(s As String) As String
      Dim r As String = """"
      Dim start As Integer = 1
      For i As Integer = 1 To Len(s)
         Dim b As Integer = Asc(Mid(s, i, 1))

         If b > 32 And b <> Asc("\") And b <> Asc("""") And b <> Asc("<") And b <> Asc(">") Then Continue For
         If start < i Then r &= Mid(s, start, i - start)

         Select Case b
            Case Asc("\"), Asc("""")
               r &= "\"
               r &= Chr(b)

            Case 10
               r &= "\n"

            Case 13
               r &= "\r"

            Case Else
               r &= "\u00" & Right("00" & Hex(b), 2)
         End Select

         start = i + 1
      Next

      If start <= Len(s) Then r &= Mid(s, start)
      r &= """"
      Return r
   End Function

   Public Function IsXmlObj(ByVal o As Object) As Boolean
      If TypeOf (o) Is XmlDocument Then Return True
      If TypeOf o Is XmlNode Then Return True
      Return False
   End Function

   Public Function XML_Obj2Str(ByVal Obj As Object) As String
      Dim objStreamWriter As New StringWriter
      Dim x As New System.Xml.Serialization.XmlSerializer(Obj.GetType)
      x.Serialize(objStreamWriter, Obj)
      Dim S As String = objStreamWriter.ToString
      objStreamWriter.Close()
      Return Mid(S, InStr(S, Chr(10)) + 1)
   End Function

   Public Function XML_Str2Obj(ByRef xmlString As String) As System.Xml.XmlDocument
      Dim xmlDoc As XmlDocument = New System.Xml.XmlDocument
      If xmlString = "" Then Return xmlDoc

      If Left(xmlString, 2) = "<?" Then
         Dim I As Integer = InStr(xmlString, ">")
         xmlString = Left(xmlString, I) & "<xml>" & Mid(xmlString, I + 1) & "</xml>"
      End If
      xmlDoc.LoadXml(xmlString)
      Return xmlDoc
   End Function

   Public Function C2Str(ByVal vValue As Object, Optional ByVal PrettyPrint As Boolean = False) As String
      Try
         If IsDBNull(vValue) OrElse vValue Is Nothing Then Return Nothing

         If TypeOf vValue Is String Then Return vValue
         If TypeOf vValue Is Guid Then Return vValue.ToString
         If TypeOf vValue Is rSelectList Then Return CType(vValue, rSelectList).GetListOfPKs
         If TypeOf vValue Is Collection OrElse TypeOf vValue Is NameObjectCollectionBase.KeysCollection Or IsArray(vValue) Or TypeOf vValue Is ArrayList Then Return C2str2(vValue, 254)
         If TypeOf vValue Is Boolean Then If vValue Then Return 1 Else Return 0

         If IsJsonObj(vValue) OrElse TypeOf vValue Is System.Collections.Generic.Dictionary(Of System.String, System.String) Then
            Dim S As String = JSON_Obj2Str(vValue)
            If PrettyPrint Then Return JSON_PrettyFormat(S) Else Return S
         End If

         If TypeOf vValue Is XmlDocument Then
            Return CType(vValue, XmlDocument).DocumentElement.InnerXml
         End If

         If TypeOf vValue Is XmlNode Then Return CType(vValue, XmlNode).InnerXml
         If TypeOf vValue Is KeyValuePair(Of String, Object) Then
            Return C2Str(CType(vValue, KeyValuePair(Of String, Object)).Value)
         End If

         If TypeOf vValue Is XmlAttributeCollection Then
            Dim XC As XmlAttributeCollection = CType(vValue, XmlAttributeCollection)
            Dim R As String = ""
            For i As Integer = 0 To XC.Count - 1
               If Len(R) Then R = R & " "
               R &= XC.ItemOf(i).OuterXml
            Next
            Return R
         End If

         If TypeOf vValue Is clsTableHandleAdo Then Return "ado:" & CType(vValue, clsTableHandle).TableName
         If TypeOf vValue Is clsTableHandleDos Then
            If CType(vValue, clsTableHandleDos).IsDict Then Return "dos:DICT " & CType(vValue, clsTableHandle).TableName Else Return "dos:" & CType(vValue, clsTableHandle).TableName
         End If
         If TypeOf vValue Is clsTableHandleHttp Then Return "http:" & CType(vValue, clsTableHandle).TableName
         If TypeOf vValue Is Byte() Then Return bytesToString(vValue)

         If TypeOf vValue Is XmlEntry Then
            If CType(vValue, XmlEntry).Count > 0 Then
               Return CType(vValue, XmlEntry).OuterXML(PrettyPrint)
            Else
               Return CType(vValue, XmlEntry).InnerXML(PrettyPrint)
            End If
         End If

         If vValue.ToString = "System.Web.SessionState.HttpSessionState" Then
            Dim S As New StringBuilder
            For I As Integer = 0 To vValue.Count - 1
               If I > 0 Then S.Append(Chr(254))
               S.Append(vValue.Keys(I).ToString() & " " & vValue(I).ToString())
            Next
            Return S.ToString
         End If

         Return vValue.ToString
      Catch ex As Exception
      End Try

      Return Nothing
   End Function

   Private Function AVS_Index(ByRef P As Object, ByRef StopOnNext As String, ByRef AmCnt As Integer, ByRef VMCnt As Integer, ByRef SvmCnt As Integer, ByRef L As Integer) As Integer
      ' Return starting position & Length (L)
      Dim EndStr, NewEndStr, PI As Integer
      Dim J As Short

      StopOnNext = Chr(255)
      EndStr = Len(P) + 1

      ' Position to A
      PI = 1
      If AmCnt > 0 Then
         StopOnNext = Chr(254)
         For J = 2 To AmCnt
            PI = InStr(PI, P, StopOnNext)
            If PI = 0 Then
               AVS_Index = 0
               L = 0
               Exit Function
            End If
            PI += 1
         Next J

         NewEndStr = InStr(PI, P, StopOnNext)
         If NewEndStr <> 0 And NewEndStr < EndStr Then EndStr = NewEndStr

      ElseIf AmCnt < 0 Then
         AVS_Index = 0
         StopOnNext = Chr(254)
         L = 0
         Exit Function ' No such position
      End If

      If VMCnt > 0 Then
         StopOnNext = Chr(253)
         For J = 2 To VMCnt
            PI = InStr(PI, P, StopOnNext)
            If PI >= EndStr Or PI = 0 Then
               AVS_Index = 0
               L = 0
               Exit Function
            End If
            PI += 1
         Next J

         NewEndStr = InStr(PI, P, StopOnNext)
         If NewEndStr <> 0 And NewEndStr < EndStr Then EndStr = NewEndStr

      ElseIf VMCnt < 0 Then
         AVS_Index = 0
         StopOnNext = Chr(253)
         L = 0
         Exit Function ' No such position
      End If

      If SvmCnt > 0 Then
         StopOnNext = Chr(252)
         For J = 2 To SvmCnt
            PI = InStr(PI, P, StopOnNext)
            If PI >= EndStr Or PI = 0 Then
               AVS_Index = 0
               L = 0
               Exit Function
            End If
            PI += 1
         Next J

         NewEndStr = InStr(PI, P, StopOnNext)
         If NewEndStr <> 0 And NewEndStr < EndStr Then EndStr = NewEndStr

      ElseIf SvmCnt < 0 Then
         AVS_Index = 0
         StopOnNext = Chr(252)
         L = 0
         Exit Function ' No such position
      End If

      L = EndStr - PI
      AVS_Index = PI
   End Function

   Enum splitTypes
      jsbBasicDelimeters = 1
      javaScriptDelimeters = 2
      visualBasicDelimeters = 3
      CShapeDelimeters = 4
      SqlDelimeters = 5
      NoSpecialDelimeters = 6
   End Enum

   Function SplitX(ByVal SplitS As String, ByVal Del As String, ByVal SplitType As Integer) As ArrayList
      Dim AL As New ArrayList
      If Len(SplitS) = 0 Then Return AL

      Dim StrDels As String = ""
      Dim lineComments As String = ""
      Dim EscapeChar As String = ""
      Dim LDel1 As String = Left(Del, 1)

      If SplitType = 0 Then
         AL.AddRange(Split(SplitS, Del))
         Return AL
      End If

      Select Case SplitType
         Case 1 ' JSB
            StrDels = """'`"
            lineComments = "//*'"
            EscapeChar = "\"
            Del = " "

         Case 2 ' JS 
            StrDels = """'"
            lineComments = "//"
            EscapeChar = "\"
            Del = " "

         Case 3 ' VB 
            StrDels = """'"
            lineComments = "'"
            EscapeChar = ""
            Del = " "

         Case 5 ' SQL
            StrDels = """'["
            lineComments = "//--"
            EscapeChar = ""
            Del = " "

         Case 6 ' C#
            StrDels = """'"
            lineComments = "//"
            EscapeChar = "\"
            Del = " "

         Case Else ' type 4 strings
            StrDels = """'"
            lineComments = ""
            EscapeChar = "\\"
      End Select

      Dim SplitResult As New StringBuilder
      Dim Split_ProcessingString As Boolean = False
      Dim SplitIsAtBOL As Boolean = True
      Dim SkipToNextBOL As Boolean = False
      Dim EndOnDel As String = ""

      For I = 1 To Len(SplitS)
         ' Get the next character
         Dim C As Char = Mid(SplitS, I, 1)
         SplitResult.Append(C)

         If SkipToNextBOL Then
            If C = Chr(13) OrElse C = Chr(10) OrElse C = Chr(254) Then SkipToNextBOL = False

         ElseIf Split_ProcessingString Then
            ' Done with string?
            If C = EndOnDel Then
               Split_ProcessingString = False

               If (SplitType <> 4 AndAlso SplitType <= 6) Then
                  Dim nc = Mid(SplitS, I + 1, Del.Length)
                  If (nc <> "" AndAlso nc <> Del) Then SplitResult.Append(Del & Chr(0))
               End If

               ' Escaping a char?
            ElseIf InStr(EscapeChar, C) > 0 AndAlso Len(EscapeChar) > 0 Then
               I += 1
               SplitResult.Append(Mid(SplitS, I, 1))

            ElseIf C = Chr(13) OrElse C = Chr(10) OrElse C = Chr(254) Then
               If EndOnDel <> "`" AndAlso SplitType <> 4 AndAlso SplitType <= 6 Then Split_ProcessingString = False

            ElseIf (EndOnDel = "]" AndAlso C = "[") Then
               SplitResult.Length -= 1
               SplitResult.Append(Chr(0))
               Split_ProcessingString = False
               I -= 1
            End If

            ' Starting a string?
         ElseIf InStr(StrDels, C) > 0 Then
            Split_ProcessingString = True
            EndOnDel = C

            If C = "[" Then EndOnDel = "]"
            If C = "(" Then EndOnDel = ")"
            If C = "{" Then EndOnDel = "}"

            SplitIsAtBOL = False

            If SplitType = 1 Then
               If C = "`" Then EscapeChar = "" Else EscapeChar = "\"
            End If

            ' begining of line?
         ElseIf SplitIsAtBOL AndAlso Len(lineComments) > 0 AndAlso InStr(lineComments, C) > 0 Then
            If C = "/" Or C = "-" Then
               If Mid(SplitS, I + 1, 1) = C Then SkipToNextBOL = True
            Else
               SkipToNextBOL = True
            End If
            SplitIsAtBOL = False

         ElseIf C = LDel1 AndAlso Mid(SplitS, I, Len(Del)) = Del Then
            SplitResult.Append(Chr(0))
            SplitIsAtBOL = False

         ElseIf C = Chr(13) OrElse C = Chr(10) OrElse C = Chr(254) Then
            SplitIsAtBOL = True

         Else
            SplitIsAtBOL = False

            ' Types 1,2,3,5,6 = JSB, JS, SQL, C#
            If SplitType <> 4 AndAlso SplitType <= 6 Then
               ' Starting a number?
               If (C >= "0" And C <= "9") Then
                  I += 1
                  C = Mid(SplitS, I, 1)
                  While (C >= "0" And C <= "9")
                     SplitResult.Append(C)
                     I += 1
                     C = Mid(SplitS, I, 1)
                  End While

                  If C = "." Then
                     SplitResult.Append(C)
                     I += 1
                     C = Mid(SplitS, I, 1)
                     While (C >= "0" And C <= "9")
                        SplitResult.Append(C)
                        I += 1
                        C = Mid(SplitS, I, 1)
                     End While
                  End If
                  I -= 1

                  ' Starting an identifier?
               ElseIf IsAlpha(C) Or C = "_" Or (C = "*" And InStr("0123456789_abcdefghijklmnopqrstuvwxyz", LCase(Mid(SplitS, I + 1, 1)))) Then
                  I += 1
                  C = Mid(SplitS, I, 1)
                  While IsAlphaNum(C) Or C = "_" Or C = "."
                     SplitResult.Append(C)
                     I += 1
                     C = Mid(SplitS, I, 1)
                  End While

                  ' function call?
                  If C = "(" Then
                     SplitResult.Append(C)
                     I += 1
                  End If
                  I -= 1

                  ' White space
               ElseIf C = " " Or C = Chr(9) Then
                  I += 1
                  C = Mid(SplitS, I, 1)
                  While C = " " Or C = Chr(9)
                     SplitResult.Append(C)
                     I += 1
                     C = Mid(SplitS, I, 1)
                  End While
                  I -= 1

                  ' Starting a conditional operator
               ElseIf InStr("<>=!", C) > 0 Then
                  I += 1

                  C = Mid(SplitS, I, 1)
                  While InStr("<>=!", C) > 0
                     SplitResult.Append(C)
                     I += 1
                     C = Mid(SplitS, I, 1)
                  End While
                  I -= 1

               ElseIf (InStr("&|", C) > 0) Then
                  I += 1
                  If (C = Mid(SplitS, I, 1)) Then
                     SplitResult.Append(C)
                     I += 1
                     C = Mid(SplitS, I, 1)
                  End If
                  I -= 1
               End If

               If I < SplitS.Length Then SplitResult.Append(Chr(0))
            End If
         End If
      Next

      SplitS = SplitResult.ToString


      If SplitType <> 4 AndAlso SplitType <= 6 Then
         SplitS = VB6.Replace(SplitS, Del & Chr(0), Chr(0))
         While InStr(SplitS, Chr(0) & Chr(0)) > 0
            SplitS = VB6.Replace(SplitS, Chr(0) & Chr(0), Chr(0))
         End While

         While Left(SplitS, 1) = Chr(0)
            SplitS = Mid(SplitS, 2)
         End While

         While Right(SplitS, 1) = Chr(0)
            SplitS = Left(SplitS, Len(SplitS) - 1)
         End While

         AL.AddRange(Split(SplitS, Chr(0)))
      Else
         AL.AddRange(Split(SplitS, Del & Chr(0)))
      End If

      Return AL
   End Function

   ' Tablename is jsbserver: www.blablabla.com|tablename
   Public Function UrlEncode(ByVal X As String) As String
      ' Return HttpContext.Current.Server.UrlDecode(X)
      Dim A As New StringBuilder

      For i As Integer = 1 To Len(X)
         Dim C As String = Mid(X, i, 1)
         If C = " " Then
            C = "+"

         ElseIf C >= "A" AndAlso C <= "Z" Then
         ElseIf C >= "a" AndAlso C <= "z" Then
         ElseIf C >= "0" AndAlso C <= "9" Then
         ElseIf C = "-" Or C = "_" Or C = "." Or C = "~" Then

         Else
            C = "%" & Right("0" & DTX(Asc(C)), 2)
         End If
         A.Append(C)
      Next
      Return A.ToString
   End Function

   Public Function DTX(ByRef D As Object) As String
      If IsDBNull(D) Then Return "" Else Return Hex(CNum(D))
   End Function

   ' Use the XTD public function to convert a hexadecimal number to its decimal equivalent.
   Public Function Xtd(ByRef X As String) As Object
      If IsDBNull(X) Then
         Return System.DBNull.Value
      Else
         Return CDbl(Convert.ToInt64(X, 16))
      End If
   End Function

   ' Change a As String of hex "65696200D0A" into a As String of ascii chars "ABC"
   Function XTS(ByVal x As String) As String
      If IsOdd(Len(x)) Then
         Dim nx As String = "0" & x
         Return XTS(nx)
      End If

      Dim s2 As New StringBuilder

      For i As Integer = 0 To Len(x) - 2 Step 2
         s2.Append(Chr(Convert.ToByte(x.Substring(i, 2), 16)))
      Next

      Return s2.ToString
   End Function

   ' Change a As String of Ascii chars "ABC" into a As String of  hex "65696200D0A"
   Function STX(ByVal s As String) As String
      Dim A(Len(s) - 1) As String
      For I As Integer = 0 To Len(s) - 1
         Dim C As String = Hex(Asc(s.Substring(I, 1)))
         If Len(C) = 1 Then A(I) &= "0" & C Else A(I) = C
      Next
      Return String.Join("", A)
   End Function

   Public Function UrlDecode(ByVal X As String) As String
      ' Return HttpContext.Current.Server.UrlDecode(X)
      Dim A As New StringBuilder
      For i As Integer = 1 To Len(X)
         Dim C As String = Mid(X, i, 1)
         If C = "%" Then
            C = Mid(X, i + 1, 2)
            C = Chr(Xtd(C))
            i = i + 2
         ElseIf C = "+" Then
            C = " "
         End If
         A.Append(C)
      Next
      Return A.ToString
   End Function

   Function Hex2(C As Char) As String
      Dim b As Byte = Asc(C)
      If Asc(b) < 16 Then Return "0" & Hex(b) Else Return Hex(b)
   End Function

   ' Escape funny characters
   Public Function DosEncodeID(ByVal IID As String) As String
      Dim C As Char, NewIID As String = ""
      Dim I As Short
      For I = 1 To Len(IID)
         C = Mid(IID, I, 1)
         If InStr("$%\+|/<*>:?*""", C) Then
            NewIID &= "$" & Hex2(C)

         ElseIf I = 1 AndAlso C = "." Then
            NewIID &= "$" & Hex2(C)

            'ElseIf C = Chr(160) Then
            '   NewIID = NewIID & Chr(32)

         Else
            NewIID = NewIID & C
         End If
      Next I
      Return NewIID
   End Function

   Public Function DosDecodeID(ByVal IID As String) As String
      If InStr(IID, "$") = 0 Then Return IID

      Dim R As String = "", HX1 As String, HX2 As String
      For Idx As Integer = 1 To Len(IID)
         Dim C As Char = Mid(IID, Idx, 1)
         If C = "$" Then

            HX1 = Mid(IID, Idx + 1, 1)
            HX2 = Mid(IID, Idx + 2, 1)
            If InStr("0123456789ABCDEF", HX1) AndAlso InStr("0123456789ABCDEF", HX2) Then
               R &= Chr(Xtd(HX1 & HX2))
               Idx += 2
            Else
               R &= C
            End If
         Else
            R &= C
         End If
      Next
      Return R
   End Function

   Public Function C2Bool(ByVal O As Object) As Boolean
      Try
         If TypeOf O Is String Then Return Len(O) <> 0 AndAlso LCase(CStr(O)) <> "false" AndAlso LCase(CStr(O)) <> "no"
         Return CBool(O)
      Catch ex As Exception
         Return False
      End Try

      Return Nothing
   End Function

   ' Returns the file.ext portion of a pathname
   Public Function GetFName(ByVal SFile As String) As String
      Dim C As Char
      For lCount As Integer = Len(SFile) To 1 Step -1
         C = Mid$(SFile, lCount, 1)
         If C = "\" Or C = ":" Or C = "/" Then Return Mid$(SFile, lCount + 1)
      Next
      Return SFile
   End Function

   ' Use the Extract public function to access the data contents of a specified field, Value, or subValue
   Public Function Extract(ByRef EXPR As String, ByVal AtrNo As Integer, Optional ByVal ValNo As Integer = 0, Optional ByVal SubNo As Integer = 0) As Object
      Dim StartI, L As Integer
      Dim StopOnNext As String = ""
      If EXPR = "" Or IsDBNull(EXPR) Then Return Nothing
      ' Get the AVS index position
      StartI = AVS_Index(EXPR, StopOnNext, AtrNo, ValNo, SubNo, L)
      If StartI = 0 Or L <= 0 Then Return Nothing
      Return Mid(EXPR, StartI, L)
   End Function

   Function Replace(ByVal EXPR As String, ByVal AtrNo As Integer, ByVal ValNo As Integer, ByVal SubNo As Integer, ByVal NewStr As String) As String
      Dim PI, L As Integer
      Dim StopOnNext As String = ""
      PI = AVS_Index(EXPR, StopOnNext, AtrNo, ValNo, SubNo, L)
      If PI = 0 Then Return Insert(EXPR, AtrNo, ValNo, SubNo, NewStr)
      Return Left(EXPR, PI - 1) + NewStr + Mid(EXPR, PI + L)
   End Function

   Public Function Insert(ByRef EXPR As String, ByVal AmCnt As Integer, ByVal VMCnt As Integer, ByVal SvmCnt As Integer, ByRef NewStr As Object) As Object
      ' Use the Insert public function to return a dynamic array that has a new field,
      ' Value, or subValue inserted into the specified dynamic array.
      ' Result:        is the dynamic array to be modified.
      ' NewStr:   specifies the Value of the new element to be inserted.
      ' AmCnt,
      ' VmCnt,
      ' SvmCnt:   specify the Type and position of the new element to be inserted.
      '
      Dim PI As Integer
      Dim EndStr As Integer
      Dim NewEndStr As Integer
      Dim J As Integer
      Dim StopOnNext As String
      Dim pChar As String
      Dim Result As String

      ' Position to A

      Result = EXPR
      PI = 1
      StopOnNext = Chr(255)
      EndStr = InStr(Result, Chr(255))
      If EndStr = 0 Then EndStr = Len(Result) + 1

      If AmCnt > 0 Then
         For J = 2 To AmCnt
            PI = InStr(PI, Result, Chr(254))
            If PI = 0 Then
               Dim AddMarks As Integer = AmCnt - J + 1
               Result = Result & New String(Chr(254), AddMarks)
               PI = EndStr + AddMarks
               EndStr += AddMarks
               Exit For
            Else
               PI += 1
            End If
         Next J
         NewEndStr = InStr(PI, Result, Chr(254))
         If NewEndStr <> 0 And NewEndStr < EndStr Then EndStr = NewEndStr
         StopOnNext = Chr(254)
      Else
         If AmCnt < 0 Then
            If EndStr > PI Then
               Result = Mid(Result, 1, EndStr - 1) & Chr(254) & Mid(Result, EndStr)
               EndStr += 1
            End If
            StopOnNext = Chr(254)
            PI = EndStr
         End If
      End If

      If VMCnt > 0 Then
         For J = 2 To VMCnt
            PI = InStr(PI, Result, Chr(253))
            If PI >= EndStr Or PI = 0 Then
               Dim AddMarks As Integer = VMCnt - J + 1
               Result = Mid(Result, 1, EndStr - 1) & New String(Chr(253), AddMarks) & Mid(Result, EndStr)
               PI = EndStr + AddMarks
               EndStr += AddMarks
               Exit For
            Else
               PI += 1
            End If
         Next J
         NewEndStr = InStr(PI, Result, Chr(253))
         If NewEndStr <> 0 And NewEndStr < EndStr Then EndStr = NewEndStr
         StopOnNext = Chr(253)
      Else
         If VMCnt < 0 Then
            NewEndStr = InStr(PI, Result, Chr(254))
            If NewEndStr <> 0 And NewEndStr < EndStr Then EndStr = NewEndStr
            If EndStr > PI Then
               Result = Mid(Result, 1, EndStr - 1) & Chr(253) & Mid(Result, EndStr)
               EndStr = EndStr + 1
            End If
            StopOnNext = Chr(253)
            PI = EndStr
         End If
      End If

      If SvmCnt > 0 Then
         For J = 2 To SvmCnt
            PI = InStr(PI, Result, Chr(252))
            If PI >= EndStr Or PI = 0 Then
               Dim AddMarks As Integer = SvmCnt - J + 1
               Result = Mid(Result, 1, EndStr - 1) & New String(Chr(252), AddMarks) & Mid(Result, EndStr)
               PI = EndStr + AddMarks
               EndStr += AddMarks
               Exit For
            Else
               PI += 1
            End If
         Next J
         StopOnNext = Chr(252)
      Else
         If SvmCnt < 0 Then
            NewEndStr = InStr(PI, Result, Chr(253))
            If NewEndStr <> 0 And NewEndStr < EndStr Then EndStr = NewEndStr
            If EndStr > PI Then
               Result = Mid(Result, 1, EndStr - 1) & Chr(252) & Mid(Result, EndStr)
               EndStr = EndStr + 1
            End If
            StopOnNext = Chr(252)
            PI = EndStr
         End If
      End If

      pChar = Mid(Result, PI, 1)
      If pChar <> "" And pChar <= StopOnNext Then
         Insert = Mid(Result, 1, PI - 1) & NewStr & StopOnNext & Mid(Result, PI)
      Else
         Insert = Mid(Result, 1, PI - 1) & NewStr & Mid(Result, PI)
      End If
   End Function

   Public Function JSON(ByRef str As String) As Object ' System.Collections.Generic.Dictionary(Of System.String, System.Object)
      If Left(str, 1) = "[" Then
         Dim T As New System.Web.Script.Serialization.JavaScriptSerializer
         T.MaxJsonLength = Integer.MaxValue
         Dim O As Object = T.DeserializeObject("{""schema"":" & str & "}")
         O = ConvertJSonArrays2ArrayList(O)
         Return O("schema")
      Else
         Return ConvertJSonArrays2ArrayList((New System.Web.Script.Serialization.JavaScriptSerializer).DeserializeObject(str))
      End If
   End Function

   Public Function JsonNthValue(ByVal j As Object, I As Integer) As Object
      If I >= CType(j, System.Collections.Generic.Dictionary(Of String, Object)).Keys.Count Then Return Nothing
      Return j(CType(j, System.Collections.Generic.Dictionary(Of String, Object)).Keys(I))
   End Function

   Public Function ConvertJSonArrays2ArrayList(ByVal O As Object) As Object
      If Not TypeOf O Is System.Collections.Generic.Dictionary(Of System.String, System.Object) Then Return O

      Dim Result As System.Collections.Generic.Dictionary(Of System.String, System.Object) = New System.Collections.Generic.Dictionary(Of System.String, System.Object)
      For Each Item As KeyValuePair(Of String, Object) In O
         If TypeOf Item.Value Is Array Then
            Dim AL As New ArrayList
            For Each ArrayElement As Object In Item.Value
               AL.Add(ConvertJSonArrays2ArrayList(ArrayElement))
            Next
            Result.Add(Item.Key, AL)
         Else
            Result.Add(Item.Key, ConvertJSonArrays2ArrayList(Item.Value))
         End If
      Next
      Return Result
   End Function

   Public Function NumTest(ByRef V As String) As Object
      On Error GoTo VChkExit

      If V = "" Then Return Nothing

      If CStr(Val(V)) = V Then
         NumTest = Val(V)
         Exit Function
      End If

      NumTest = V
      Exit Function

VChkExit:
      Err.Clear()
      NumTest = V
   End Function

   Function NumericFld(ByVal DataType As Type) As Boolean
      If DataType Is GetType(System.Boolean) Then Return True
      If DataType Is GetType(System.Decimal) Then Return True
      If DataType Is GetType(System.Int16) Then Return True
      If DataType Is GetType(System.Int32) Then Return True
      If DataType Is GetType(System.Int64) Then Return True
      If DataType Is GetType(System.Double) Then Return True
      If DataType Is GetType(System.Single) Then Return True
      If DataType Is GetType(System.SByte) Then Return True
      If DataType Is GetType(System.UInt16) Then Return True
      If DataType Is GetType(System.UInt32) Then Return True
      If DataType Is GetType(System.Byte) Then Return True
      Return False
   End Function

   ' Method is "GET", "POST", "PUT", "DELETE"
   ' Headers is VM seperated strings, "xxx: yyy" & VM & ...
   '
   Function UrlFetch(ByVal Method As String, ByVal Url As String, ByVal Headers As String, ByVal Body As String, ByRef Result As String, ByRef RtnHeader As String, Optional ByVal NoRedirecting As Boolean = False, Optional ByVal secTimeOut As Long = 30) As Boolean
      Dim AM As String = Chr(254), VM As String = Chr(253), SVM As String = Chr(252)
      Dim ShowDbg As Boolean = False

      If Headers = "" Then
         Headers = "Cache-Control: no-cache"
         Headers &= VM & "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
         Headers &= VM & "Accept-Encoding: gzip, deflate, sdch"
         Headers &= VM & "Accept-Language: en-US,en;q=0.8"
         Headers &= VM & "Cache-Control: max-age=0"
         Headers &= VM & "Connection: keep-alive"
         If Method = "POST" Or Method = "PUT" Then Headers &= VM & "Content-Type: application/x-www-form-urlencoded"
      End If

      For redirects As Integer = 1 To 2
         Dim Req As HttpWebRequest = WebRequest.Create(New Uri(Url))
         If Method = "" Then Method = "GET"
         Req.Method = Method ' GET / POST / DELETE
         Req.CookieContainer = SessionCookies

         Req.Timeout = secTimeOut * 1000 ' convert to milliseconds

         ' add headers
         For Each sHeader As String In Split(Headers, VM)
            Dim i As Integer = InStr(sHeader, "=")
            If i = 0 Then i = InStr(sHeader, ":")
            If InStr(sHeader & ":", ":") < i Then i = InStr(sHeader, ":")
            If i > 0 Then
               Dim Key As String = Left(sHeader, i - 1)
               Dim S2 As String = LTrim(Mid(sHeader, i + 1))

               If ShowDbg Then Debug.Print("    request.header." & Key & ":" & S2)
               Try
                  Select Case LCase(Key)
                     Case "keep-alive"
                        If LCase(S2) = "true" Then Req.KeepAlive = True

                     Case "accept"
                        Req.Accept = S2

                     Case "user-agent"
                        Req.UserAgent = S2

                     Case "content-type"
                        Req.ContentType = S2

                     Case "content-length"
                        ' Ignore Req.ContentLength = Body.Length

                     Case "connection"
                        If LCase(S2) = "keep-alive" Then Req.KeepAlive = True

                     Case "timeout"
                        Req.Timeout = CInt(S2)

                     Case "date"

                     Case "cache-control" ' =no-cache
                        Req.Headers.Set(Key, S2)

                     Case "accept-encoding" ' gzip
                        '    Req.Headers.Set(Key, S2)

                     Case "accept-language"
                        Req.Headers.Set(Key, S2)

                     Case "cookie"
                        Req.Headers.Set(Key, S2)

                     Case "host"
                        Req.Host = Field(Field(Url, "/", 3), "?", 1) ' "webservice.lotek.com"


                     Case "origin"
                        ' We want an origin
                        ' Req.Host = S2

                     Case "referer"
                        ' can't be modified directly 
                        '  Req.Headers.Set("Origin", S2)
                        Req.Referer = Field(Url, "?", 1)

                     Case "dnt"
                        Req.Headers.Set(Key, S2)

                     Case "transfer-encoding"
                        Req.TransferEncoding = S2

                     Case Else
                        Req.Headers.Set(Key, S2)

                  End Select
               Catch ex As Exception
                  If ShowDbg Then Debug.Print("       !!!!!! unable to set header " & Key & ". error: " & ex.Message)
               End Try
            End If

         Next

         ' Put body into request
         If Body = "" Or Method = "GET" Then
            Req.ContentLength = 0
         Else
            Dim encoding As New System.Text.ASCIIEncoding
            Dim BodyBytes As Byte() = encoding.GetBytes(Body)
            Req.ContentLength = BodyBytes.Length
            Dim newStream As Stream = Req.GetRequestStream()
            newStream.Write(BodyBytes, 0, BodyBytes.Length)
            newStream.Close()
         End If

         '  FetchRequest 
         Dim Response As WebResponse, StatusCode As Integer, Status As String
         Try
            Response = Req.GetResponse()
            StatusCode = DirectCast(Response, System.Net.HttpWebResponse).StatusCode
            Status = DirectCast(Response, System.Net.HttpWebResponse).StatusDescription

            ' Get Content
            ' Dim Response = FetchRequest(Req, myTimeOut, NoRedirecting)
            Dim CT As String = LCase(Response.ContentType)
            If InStr(CT, "json") Then
               Dim TempStream As New MemoryStream
               Dim bytesRead As Integer = 0
               Dim sResult As New StringBuilder
               Do
                  Dim inBuf(4096) As Byte
                  bytesRead = Response.GetResponseStream.Read(inBuf, 0, 4096)
                  For i As Integer = 0 To bytesRead - 1
                     sResult.Append(Chr(CInt(inBuf(i))))
                  Next
               Loop While bytesRead > 0
               Result = sResult.ToString

            ElseIf InStr(CT, "text") Or InStr(CT, "application") Or InStr(CT, "utf") Or Response.ContentLength = -1 Then
               Dim responseReader As StreamReader
               Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")
               responseReader = New StreamReader(Response.GetResponseStream())
               Result = responseReader.ReadToEnd()
               responseReader.Close()

            Else

               'Read the answer from the website and store it into a stream
               Dim objStream As Stream
               objStream = Response.GetResponseStream
               Dim inBuf(DirectCast(Response, System.Net.HttpWebResponse).ContentLength) As Byte ' 10 MByte Max --- use Response.ContentLength??

               Dim bytesToRead As Integer = CInt(inBuf.Length)
               Dim bytesRead As Integer = 0
               While bytesToRead > 0
                  Dim n As Integer = objStream.Read(inBuf, bytesRead, bytesToRead)
                  If n = 0 Then Exit While
                  bytesRead += n
                  bytesToRead -= n
               End While

               objStream.Close()
               Response.Close()

               ReDim Preserve inBuf(bytesRead - 1)

               ' Convert to string
               Result = System.Text.Encoding.Default.GetString(inBuf)
            End If

            ' Build header
            RtnHeader = StatusCode & AM & CStr(StatusCode) & ":" & Status
            For Each Key As String In Response.Headers.AllKeys
               RtnHeader &= AM & Key & "=" & Response.Headers(Key)
            Next
            RtnHeader &= AM & "ResponseUri=" & Response.ResponseUri.AbsoluteUri

         Catch ex As WebException
            Dim resp As HttpWebResponse = CType(ex.Response, HttpWebResponse)
            Result = ex.Message
            If resp IsNot Nothing Then
               StatusCode = resp.StatusCode
               Status = resp.StatusDescription
            Else
               Status = ex.Message
               StatusCode = 400
            End If
            RtnHeader = StatusCode & AM & CStr(StatusCode) & ":" & Status
            Return False
         End Try

         ' Success??
         If Method = "POST" Then
            If StatusCode = 200 Or StatusCode = 201 Or StatusCode = 202 Then
               Return True
            End If
            Method = "GET"
         Else
            If StatusCode = 200 Or StatusCode = 201 Then Return True
         End If

         ' if not redirecting we are done!
         If Not shouldRedirect(StatusCode) Or NoRedirecting Then Return False

         Dim OldURL As String = Url

         ' = = = = = Do Redirect = = = =
         Url = Response.Headers.Get("Location")
         If Url = "" Then Return True

         If ShowDbg Then Debug.Print("")
         If ShowDbg Then Debug.Print("--- REDIRECTING from " & OldURL & " to " & Url & " ---")
         If ShowDbg Then Debug.Print("")

         Headers = ""
         For Each Key As String In Response.Headers.AllKeys
            If Len(Headers) Then Headers &= VM
            If Key = "Set-Cookie" Then Headers &= Key & "=" & Response.Headers(Key)
         Next

         If ShowDbg Then Debug.Print("Redirecting to " + Url)
      Next

      ' too many redirects
      Return False

   End Function

   ' True if the specified HTTP status code is one for which the Get utility should automatically redirect.
   Function shouldRedirect(statusCode As Integer) As Boolean
      Select Case statusCode
         Case 301, 302, 303, 307
            Return True
         Case Else
            Return False
      End Select
   End Function

   Function MimeType(ByVal Ext As String, ByVal defaultType As String) As String
      Ext = LCase(Right(Ext, 4))
      If Ext = ".csv" Then
         Return "text/csv"
      ElseIf Ext = ".tsv" Then
         Return "text/tab-seperated-values"
      ElseIf Ext = ".htm" Or Ext = "html" Then
         Return "text/html"
      ElseIf Ext = ".doc" Then
         Return "application/msword"
      ElseIf Ext = ".ods" Then
         Return "application/x-vnd.oasis.opendocument.spreadsheet"
      ElseIf Ext = ".odt" Then
         Return "application/vnd.oasis.opendocument.text"
      ElseIf Ext = ".rtf" Then
         Return "text/rtf"
      ElseIf Ext = ".txt" Then
         Return "text/plain"
      ElseIf Ext = ".xls" Then
         Return "text/vnd.ms-excel"
      ElseIf Ext = "xlsx" Then
         Return "text/vnd.openformats-officedocument.spreadsheetml.sheet"
      ElseIf Ext = ".pdf" Then
         Return "text/pdf"
      ElseIf Ext = ".ppt" Or Ext = "pps" Then
         Return "vnd.ms-powerpoint"
      ElseIf Ext = ".wmf" Then
         Return "image/x-wmf"
      ElseIf Ext = ".jpg" Then
         Return "image/jpeg"
      ElseIf Ext = ".gif" Then
         Return "image/gif"
      ElseIf Ext = ".png" Then
         Return "image/png"
      ElseIf Ext = ".bmp" Then
         Return "image/bmp"
      End If
      Return defaultType
   End Function


   Public Function NewCommandBuilder(ByVal DataAdapter As IDbDataAdapter) As System.ComponentModel.Component

      If TypeOf DataAdapter Is SQLite.SQLiteDataAdapter Then
         Return New SQLite.SQLiteCommandBuilder(CType(DataAdapter, SQLite.SQLiteDataAdapter))
      ElseIf TypeOf DataAdapter Is OleDb.OleDbDataAdapter Then
         Return New OleDb.OleDbCommandBuilder(CType(DataAdapter, OleDb.OleDbDataAdapter))
      ElseIf TypeOf DataAdapter Is SqlClient.SqlDataAdapter Then
         Return New SqlClient.SqlCommandBuilder(CType(DataAdapter, SqlClient.SqlDataAdapter))
      ElseIf TypeOf DataAdapter Is Odbc.OdbcDataAdapter Then
         Return New Odbc.OdbcCommandBuilder(CType(DataAdapter, Odbc.OdbcDataAdapter))
#If IncludeCompactFrameWork Then
        ElseIf TypeOf DataAdapter Is SqlServerCe.SqlCeDataAdapter Then
            Return New SqlServerCe.SqlCeCommandBuilder(CType(DataAdapter, SqlServerCe.SqlCeDataAdapter))
#End If
      Else
         Return Nothing
      End If
   End Function

   Public Function NewCommand(ByVal SelectStatement As String, ByVal cnSql As IDbConnection, Optional ByVal TransactionHandle As IDbTransaction = Nothing) As IDbCommand
      If TypeOf cnSql Is SQLite.SQLiteConnection Then
         Return New SQLite.SQLiteCommand(SelectStatement, CType(cnSql, SQLite.SQLiteConnection))
      ElseIf TypeOf cnSql Is OleDb.OleDbConnection Then
         Return New OleDb.OleDbCommand(SelectStatement, CType(cnSql, OleDb.OleDbConnection), CType(TransactionHandle, OleDb.OleDbTransaction))
      ElseIf TypeOf cnSql Is SqlClient.SqlConnection Then
         Return New SqlClient.SqlCommand(SelectStatement, CType(cnSql, SqlClient.SqlConnection), CType(TransactionHandle, SqlClient.SqlTransaction))
      ElseIf TypeOf cnSql Is Odbc.OdbcConnection Then
         Return New Odbc.OdbcCommand(SelectStatement, CType(cnSql, Odbc.OdbcConnection), CType(TransactionHandle, Odbc.OdbcTransaction))
#If IncludeCompactFrameWork Then
        ElseIf TypeOf cnSql Is SqlServerCe.SqlCeConnection Then
            Return New SqlServerCe.SqlCeCommand(SelectStatement, CType(cnSql, SqlServerCe.SqlCeConnection), CType(TransactionHandle, SqlServerCe.SqlCeTransaction))
#End If
      Else
         Return Nothing
      End If
   End Function

   Public Function NewParameter(ByVal cmd As IDbCommand) As IDbDataParameter
      If TypeOf cmd Is SQLite.SQLiteCommand Then
         Return New SQLite.SQLiteParameter
      ElseIf TypeOf cmd Is OleDb.OleDbCommand Then
         Return New OleDb.OleDbParameter
      ElseIf TypeOf cmd Is SqlClient.SqlCommand Then
         Return New SqlClient.SqlParameter
      ElseIf TypeOf cmd Is Odbc.OdbcCommand Then
         Return New Odbc.OdbcParameter
#If IncludeCompactFrameWork Then
        ElseIf TypeOf cmd Is SqlServerCe.SqlCeCommand Then
            Return New SqlServerCe.SqlCeParameter
#End If
      Else
         Return Nothing
      End If
   End Function

   Public Function NewParameter(ByVal cmd As IDbCommand, ByVal ParameterName As String) As IDbDataParameter
      Dim zIDbDataParameter As IDbDataParameter = NewParameter(cmd)
      If zIDbDataParameter Is Nothing Then Return Nothing

      If TypeOf cmd Is OleDb.OleDbCommand Then
         If Left(ParameterName, 1) = "@" Then
            ParameterName = Mid(ParameterName, 2)
            cmd.CommandText = VB6.Replace(cmd.CommandText, "@" & ParameterName, "?")
         End If
      End If

      zIDbDataParameter.ParameterName = ParameterName
      Return zIDbDataParameter
   End Function


   Public Function NewParameter(ByVal cmd As IDbCommand, ByVal ParameterName As String, ByVal SourceColumn As String) As IDbDataParameter
      Dim zIDbDataParameter As IDbDataParameter = NewParameter(cmd, ParameterName)
      If zIDbDataParameter Is Nothing Then Return Nothing
      zIDbDataParameter.SourceColumn = SourceColumn
      Return zIDbDataParameter
   End Function

   Public Function NewParameter(ByVal cmd As IDbCommand, ByVal ParameterName As String, ByVal ParameterValue As Object) As IDbDataParameter
      Dim zIDbDataParameter As IDbDataParameter = NewParameter(cmd, ParameterName)
      If zIDbDataParameter Is Nothing Then Return Nothing

      zIDbDataParameter.Value = ParameterValue
      Return zIDbDataParameter
   End Function

   Public Function NewParameter(ByVal cmd As IDbCommand, ByVal ParameterName As String, ByVal SDBType As DbType, ByVal ParameterValue As Object) As IDbDataParameter
      Dim zIDbDataParameter As IDbDataParameter = NewParameter(cmd, ParameterName, ParameterValue)
      If zIDbDataParameter Is Nothing Then Return Nothing

      zIDbDataParameter.DbType = SDBType
      Return zIDbDataParameter
   End Function

   Public Function NewConnection(ByVal ConnectionString As String) As IDbConnection
      Dim cnSql As IDbConnection = Nothing
      Const DropoOLE As String = "provider=sqloledb.1;"

      If Left(LCase(ConnectionString), Len(DropoOLE)) = DropoOLE Then ConnectionString = Mid(ConnectionString, Len(DropoOLE) + 1)
      ConnectionString = VB6.Replace(ConnectionString, " =", "=")

      If InStr(LCase(ConnectionString), "msdasql.1") > 0 Then
         Dim P As Integer = InStr(LCase(ConnectionString), "data source")
         If P > 0 Then ConnectionString = Left(ConnectionString, P - 1) & "DSN" & Mid(ConnectionString, P + 11)
         cnSql = New Odbc.OdbcConnection(VB6.Replace(ConnectionString, "Data Source", "DNS"))


      ElseIf InStr(LCase(ConnectionString), "provider=") > 0 Then
         cnSql = New OleDb.OleDbConnection(ConnectionString)

#If IncludeCompactFrameWork Then
        ElseIf InStr(LCase(ConnectionString), "data source=") > 0 Then

            cnSql = New System.Data.SqlServerCe.SqlCeConnection(ConnectionString)
#End If

      ElseIf Left(LCase(ConnectionString), Len("data source=")) = "data source=" Then
         cnSql = New SQLite.SQLiteConnection(ConnectionString)

      Else

         cnSql = New SqlClient.SqlConnection(ConnectionString)
      End If

      '  cnSql.Open()
      Return cnSql
   End Function

   'Public Function NewParameter(ByVal cmd As IDbCommand, ByVal ParameterName As String, ByVal ParameterValue As Object) As IDbDataParameter
   '    Dim zIDbDataParameter As IDbDataParameter = NewParameter(cmd)
   '    If zIDbDataParameter Is Nothing Then Return Nothing

   '    zIDbDataParameter.ParameterName = ParameterName
   '    zIDbDataParameter.Value = ParameterValue
   '    Return zIDbDataParameter
   'End Function

   Public Function NewDataAdapter(ByVal cmd As IDbCommand) As IDbDataAdapter
      If TypeOf cmd Is SQLite.SQLiteCommand Then
         Return New SQLite.SQLiteDataAdapter
      ElseIf TypeOf cmd Is OleDb.OleDbCommand Then
         Return CType(New OleDb.OleDbDataAdapter(CType(cmd, OleDb.OleDbCommand)), IDbDataAdapter)
      ElseIf TypeOf cmd Is SqlClient.SqlCommand Then
         Return CType(New SqlClient.SqlDataAdapter(CType(cmd, SqlClient.SqlCommand)), IDbDataAdapter)
      ElseIf TypeOf cmd Is Odbc.OdbcCommand Then
         Return CType(New Odbc.OdbcDataAdapter(CType(cmd, Odbc.OdbcCommand)), IDbDataAdapter)
#If IncludeCompactFrameWork Then
        ElseIf TypeOf cmd Is SqlServerCe.SqlCeCommand Then
            Return CType(New SqlServerCe.SqlCeDataAdapter(CType(cmd, SqlServerCe.SqlCeCommand)), IDbDataAdapter)
#End If
      Else
         Return Nothing
      End If
   End Function

   Public Function NewDataAdapter(ByVal cnSql As IDbConnection) As IDbDataAdapter
      If TypeOf cnSql Is SQLite.SQLiteConnection Then
         Return New SQLite.SQLiteDataAdapter
      ElseIf TypeOf cnSql Is OleDb.OleDbConnection Then
         Return New OleDb.OleDbDataAdapter
      ElseIf TypeOf cnSql Is SqlClient.SqlConnection Then
         Return New SqlClient.SqlDataAdapter
      ElseIf TypeOf cnSql Is Odbc.OdbcConnection Then
         Return New Odbc.OdbcDataAdapter
#If IncludeCompactFrameWork Then
        ElseIf TypeOf cnSql Is SqlServerCe.SqlCeConnection Then
            Return New SqlServerCe.SqlCeDataAdapter
#End If
      Else
         Return Nothing
      End If
   End Function

   Public Function NewDataAdapter(ByVal SelectStatement As String, ByVal cnSql As IDbConnection) As IDbDataAdapter
      SelectStatement = SqlSelectFormat(SelectStatement, cnSql)

      If TypeOf cnSql Is SQLite.SQLiteConnection Then
         Return New SQLite.SQLiteDataAdapter(SelectStatement, CType(cnSql, SQLite.SQLiteConnection))
      ElseIf TypeOf cnSql Is OleDb.OleDbConnection Then
         Return New OleDb.OleDbDataAdapter(SelectStatement, CType(cnSql, OleDb.OleDbConnection))
      ElseIf TypeOf cnSql Is SqlClient.SqlConnection Then
         Return New SqlClient.SqlDataAdapter(SelectStatement, CType(cnSql, SqlClient.SqlConnection))
      ElseIf TypeOf cnSql Is Odbc.OdbcConnection Then
         Return New Odbc.OdbcDataAdapter(SelectStatement, CType(cnSql, Odbc.OdbcConnection))
#If IncludeCompactFrameWork Then
        ElseIf TypeOf cnSql Is SqlServerCe.SqlCeConnection Then
            Return New SqlServerCe.SqlCeDataAdapter(SelectStatement, CType(cnSql, SqlServerCe.SqlCeConnection))
#End If
      Else
         Return Nothing
      End If
   End Function

   Public Function SqlDateDelimiter(ByVal cnSql As IDbConnection) As Char
      If IsSQLServer(cnSql) Then Return "'" Else Return "#"
   End Function

   Public Function IsSQLServer(ByVal cnSql As IDbConnection) As Boolean
      If TypeOf cnSql Is OleDb.OleDbConnection Then
         Dim Provider As String = UCase(CType(cnSql, OleDb.OleDbConnection).Provider)
         Return InStr(Provider, "SQLOLEDB") Or InStr(Provider, "SQLNCLI")
      ElseIf TypeOf cnSql Is SqlClient.SqlConnection Or TypeOf cnSql Is SQLite.SQLiteConnection Then
         Return True
      ElseIf TypeOf cnSql Is Odbc.OdbcConnection Then
         Return True
#If IncludeCompactFrameWork Then
        ElseIf TypeOf cnSql Is SqlServerCe.SqlCeConnection Then
            Return True
#End If
      Else
         Return False
      End If
   End Function

   Public Function SqlSelectFormat(ByVal SelectStatement As String, ByVal cnSql As IDbConnection) As String
      Dim WorkingStr As String = SelectStatement, I As Integer, J As Integer
      Dim TopBottom As String = "", FromStr As String, Fields() As String, FieldName As String, AliasStr As String

      If TypeOf cnSql Is OleDb.OleDbConnection Then
         ' Drop Alias: XX.YY As [XX.YY]
         If InStr(LCase(SelectStatement), " as [") > 0 Then
            If Left(LCase(LTrim(SelectStatement)), 7) <> "select " Then Return SelectStatement
            WorkingStr = Trim(Mid(WorkingStr, 8))
            If Left(LTrim(LCase(WorkingStr)), 4) = "top " Then
               WorkingStr = Trim(Mid(WorkingStr, 5))
               I = InStr(WorkingStr, " ")
               If I = 0 Then Return SelectStatement
               TopBottom = " Top " & Left(WorkingStr, I)

            ElseIf Left(LTrim(LCase(WorkingStr)), 7) = "bottom " Then
               WorkingStr = Trim(Mid(WorkingStr, 8))
               I = InStr(WorkingStr, " ")
               If I = 0 Then Return SelectStatement
               TopBottom = " Top " & Left(WorkingStr, I)
            End If

            I = InStr(LCase(WorkingStr), " from ")
            If I = 0 Then Return SelectStatement
            FromStr = Mid(WorkingStr, I)
            WorkingStr = Left(WorkingStr, I - 1)

            Fields = Split(WorkingStr, ",")
            For J = 0 To UBound(Fields)
               Dim Field As String = Fields(J)
               I = InStr(LCase(Field), " as ")
               If I > 0 Then
                  FieldName = Trim(Left(Field, I - 1))
                  AliasStr = Trim(Mid(Field, I + 4))
                  If Left(AliasStr, 1) = "[" And Right(AliasStr, 1) = "]" Then AliasStr = Mid(AliasStr, 2, Len(AliasStr) - 2)
                  If LCase(FieldName) = LCase(AliasStr) Then Fields(J) = FieldName
               End If
            Next

            SelectStatement = "Select " & TopBottom & " " & String.Join(",", Fields.ToArray) & FromStr
         End If

         ' Dates need to be in #1/1/2005#
         Return SelectStatement

      Else
         ' Dates need to be in '1/1/2005'
         Return SelectStatement
      End If
   End Function

   Public Function BeginTransaction(ByVal cnSql As IDbConnection) As IDbTransaction
      If TypeOf cnSql Is SQLite.SQLiteConnection Then
         Return cnSql.BeginTransaction

      ElseIf TypeOf cnSql Is OleDb.OleDbConnection Then
         Return cnSql.BeginTransaction

      ElseIf TypeOf cnSql Is SqlClient.SqlConnection Then
         Return cnSql.BeginTransaction

      ElseIf TypeOf cnSql Is Odbc.OdbcConnection Then
         Return cnSql.BeginTransaction

#If IncludeCompactFrameWork Then
        ElseIf TypeOf cnSql Is SqlServerCe.SqlCeConnection Then
            Return cnSql.BeginTransaction
#End If

      Else
         Return Nothing
      End If
   End Function

   Public Sub SetupSqlCommands(ByVal DA As IDbDataAdapter, ByVal TblName As String, ByVal PK As String, Optional ByVal PK2 As String = "", Optional ByVal TransactionHandle As IDbTransaction = Nothing)
      Dim cb As System.ComponentModel.Component '  SqlClient.SqlCommandBuilder
      cb = NewCommandBuilder(DA)

      If TypeOf DA Is OleDb.OleDbDataAdapter Then
         CType(cb, OleDb.OleDbCommandBuilder).QuotePrefix = "["
         CType(cb, OleDb.OleDbCommandBuilder).QuoteSuffix = "]"
         DA.SelectCommand.Transaction = CType(TransactionHandle, OleDb.OleDbTransaction)
         DA.InsertCommand = CType(cb, OleDb.OleDbCommandBuilder).GetInsertCommand
         DA.UpdateCommand = BuildUpdateCommand(DA, TblName, PK, PK2)
         DA.UpdateCommand.Connection = DA.SelectCommand.Connection
         DA.DeleteCommand = BuildDeleteCommand(DA, TblName, PK, PK2)
         DA.DeleteCommand.Connection = DA.SelectCommand.Connection

      ElseIf TypeOf DA Is SQLite.SQLiteDataAdapter Then
         CType(cb, SQLite.SQLiteCommandBuilder).QuotePrefix = "["
         CType(cb, SQLite.SQLiteCommandBuilder).QuoteSuffix = "]"
         DA.SelectCommand.Transaction = CType(TransactionHandle, SQLite.SQLiteTransaction)
         DA.InsertCommand = CType(cb, SQLite.SQLiteCommandBuilder).GetInsertCommand
         DA.UpdateCommand = BuildUpdateCommand(DA, TblName, PK, PK2)
         DA.UpdateCommand.Connection = DA.SelectCommand.Connection
         DA.DeleteCommand = BuildDeleteCommand(DA, TblName, PK, PK2)
         DA.DeleteCommand.Connection = DA.SelectCommand.Connection

      ElseIf TypeOf DA Is SqlClient.SqlDataAdapter Then
         CType(cb, SqlClient.SqlCommandBuilder).QuotePrefix = "["
         CType(cb, SqlClient.SqlCommandBuilder).QuoteSuffix = "]"
         DA.SelectCommand.Transaction = CType(TransactionHandle, SqlClient.SqlTransaction)
         DA.InsertCommand = CType(cb, SqlClient.SqlCommandBuilder).GetInsertCommand
         DA.UpdateCommand = BuildUpdateCommand(DA, TblName, PK, PK2)
         DA.UpdateCommand.Connection = DA.SelectCommand.Connection
         DA.DeleteCommand = BuildDeleteCommand(DA, TblName, PK, PK2)
         DA.DeleteCommand.Connection = DA.SelectCommand.Connection

      ElseIf TypeOf DA Is Odbc.OdbcDataAdapter Then
         CType(cb, Odbc.OdbcCommandBuilder).QuotePrefix = "["
         CType(cb, Odbc.OdbcCommandBuilder).QuoteSuffix = "]"
         DA.SelectCommand.Transaction = CType(TransactionHandle, Odbc.OdbcTransaction)
         DA.InsertCommand = CType(cb, Odbc.OdbcCommandBuilder).GetInsertCommand
         DA.UpdateCommand = BuildUpdateCommand(DA, TblName, PK, PK2)
         DA.UpdateCommand.Connection = DA.SelectCommand.Connection
         DA.DeleteCommand = BuildDeleteCommand(DA, TblName, PK, PK2)
         DA.DeleteCommand.Connection = DA.SelectCommand.Connection

#If IncludeCompactFrameWork Then
            ElseIf TypeOf DA Is SqlServerCe.SqlCeDataAdapter Then
               CType(cb, SqlServerCe.SqlCeCommandBuilder).QuotePrefix = "["
               CType(cb, SqlServerCe.SqlCeCommandBuilder).QuoteSuffix = "]"
                DA.SelectCommand.Transaction = CType(TransactionHandle, SqlServerCe.SqlCeTransaction)
                DA.InsertCommand = CType(cb, SqlServerCe.SqlCeCommandBuilder).GetInsertCommand
                DA.UpdateCommand = BuildUpdateCommand(DA, TblName, PK, PK2)
                DA.UpdateCommand.Connection = DA.SelectCommand.Connection
                Try
                    DA.DeleteCommand = CType(cb, SqlServerCe.SqlCeCommandBuilder).GetDeleteCommand
                Catch ex As Exception
                    DA.DeleteCommand = BuildDeleteCommand(DA, TblName, PK, PK2)
                    DA.DeleteCommand.Connection = DA.SelectCommand.Connection
                End Try
#End If
      End If

      If Not TransactionHandle Is Nothing Then
         DA.InsertCommand.Transaction = TransactionHandle
         DA.UpdateCommand.Transaction = TransactionHandle
         DA.DeleteCommand.Transaction = TransactionHandle
         DA.SelectCommand.Transaction = TransactionHandle
      End If
   End Sub

   Public Function BuildUpdateCommand(ByVal DA As IDbDataAdapter, ByVal TableName As String, ByVal PKColumnName As String, Optional ByVal PKColumnName2 As String = "") As IDbCommand
      Dim P As IDbDataParameter, Upd As String = "", PK As IDbDataParameter, PK2 As IDbDataParameter

      Dim uCmd As IDbCommand = NewCommand(Upd, DA.SelectCommand.Connection)
      PK = NewParameter(uCmd)
      PK2 = NewParameter(uCmd)
      Upd = ""

      For Each P In DA.InsertCommand.Parameters
         If P.DbType <> DbType.Guid And P.SourceColumn <> "" Then
            If Upd <> "" Then Upd = Upd & ","
            Upd = Upd & "[" & P.SourceColumn & "] =" & P.ParameterName

            Dim NP As IDbDataParameter = NewParameter(uCmd)
            If TypeOf P Is SqlClient.SqlParameter Then CType(NP, SqlClient.SqlParameter).SqlDbType = CType(P, SqlClient.SqlParameter).SqlDbType
            If TypeOf P Is SQLite.SQLiteParameter Then CType(NP, SQLite.SQLiteParameter).DbType = CType(P, SQLite.SQLiteParameter).DbType
            If TypeOf P Is OleDb.OleDbParameter Then CType(NP, OleDb.OleDbParameter).OleDbType = CType(P, OleDb.OleDbParameter).OleDbType
            If TypeOf P Is Odbc.OdbcParameter Then CType(NP, Odbc.OdbcParameter).DbType = CType(P, Odbc.OdbcParameter).DbType
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(NP, SqlServerCe.SqlCeParameter).DbType = CType(P, SqlServerCe.SqlCeParameter).DbType
#End If
            NP.SourceColumn = P.SourceColumn()
            NP.DbType = P.DbType
            NP.Size = P.Size

            If TypeOf P Is SqlClient.SqlParameter Then CType(NP, SqlClient.SqlParameter).IsNullable = CType(P, SqlClient.SqlParameter).IsNullable
            If TypeOf P Is SQLite.SQLiteParameter Then CType(NP, SQLite.SQLiteParameter).IsNullable = CType(P, SQLite.SQLiteParameter).IsNullable
            If TypeOf P Is OleDb.OleDbParameter Then CType(NP, OleDb.OleDbParameter).IsNullable = CType(P, OleDb.OleDbParameter).IsNullable
            If TypeOf P Is Odbc.OdbcParameter Then CType(NP, Odbc.OdbcParameter).IsNullable = CType(P, Odbc.OdbcParameter).IsNullable
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(NP, SqlServerCe.SqlCeParameter).IsNullable = CType(P, SqlServerCe.SqlCeParameter).IsNullable
#End If
            NP.ParameterName = P.ParameterName
            NP.Direction = P.Direction
            NP.Precision = P.Precision

            uCmd.Parameters.Add(NP)
         End If

         If P.SourceColumn = PKColumnName Then
            If TypeOf P Is SqlClient.SqlParameter Then CType(PK, SqlClient.SqlParameter).SqlDbType = CType(P, SqlClient.SqlParameter).SqlDbType
            If TypeOf P Is SQLite.SQLiteParameter Then CType(PK, SQLite.SQLiteParameter).DbType = CType(P, SQLite.SQLiteParameter).DbType
            If TypeOf P Is OleDb.OleDbParameter Then CType(PK, OleDb.OleDbParameter).OleDbType = CType(P, OleDb.OleDbParameter).OleDbType
            If TypeOf P Is Odbc.OdbcParameter Then CType(PK, Odbc.OdbcParameter).DbType = CType(P, Odbc.OdbcParameter).DbType
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(PK, SqlServerCe.SqlCeParameter).DbType = CType(P, SqlServerCe.SqlCeParameter).DbType
#End If
            PK.SourceColumn = P.SourceColumn
            PK.DbType = P.DbType
            PK.SourceVersion = DataRowVersion.Original
            PK.Size = P.Size
            If TypeOf P Is SqlClient.SqlParameter Then CType(PK, SqlClient.SqlParameter).IsNullable = CType(P, SqlClient.SqlParameter).IsNullable
            If TypeOf P Is SQLite.SQLiteParameter Then CType(PK, SQLite.SQLiteParameter).IsNullable = CType(P, SQLite.SQLiteParameter).IsNullable
            If TypeOf P Is OleDb.OleDbParameter Then CType(PK, OleDb.OleDbParameter).IsNullable = CType(P, OleDb.OleDbParameter).IsNullable
            If TypeOf P Is Odbc.OdbcParameter Then CType(PK, Odbc.OdbcParameter).IsNullable = CType(P, Odbc.OdbcParameter).IsNullable
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(PK, SqlServerCe.SqlCeParameter).IsNullable = CType(P, SqlServerCe.SqlCeParameter).IsNullable
#End If

            PK.ParameterName = "@PK1"
            PK.Direction = P.Direction
            PK.Precision = P.Precision
         End If

         If P.SourceColumn = PKColumnName2 Then
            If TypeOf P Is SqlClient.SqlParameter Then CType(PK2, SqlClient.SqlParameter).SqlDbType = CType(P, SqlClient.SqlParameter).SqlDbType
            If TypeOf P Is SQLite.SQLiteParameter Then CType(PK2, SQLite.SQLiteParameter).DbType = CType(P, SQLite.SQLiteParameter).DbType
            If TypeOf P Is OleDb.OleDbParameter Then CType(PK2, OleDb.OleDbParameter).OleDbType = CType(P, OleDb.OleDbParameter).OleDbType
            If TypeOf P Is Odbc.OdbcParameter Then CType(PK2, Odbc.OdbcParameter).DbType = CType(P, Odbc.OdbcParameter).DbType
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(PK2, SqlServerCe.SqlCeParameter).DbType = CType(P, SqlServerCe.SqlCeParameter).DbType
#End If
            PK2.SourceColumn = P.SourceColumn
            PK2.DbType = P.DbType
            PK2.SourceVersion = DataRowVersion.Original
            PK2.Size = P.Size

            If TypeOf P Is SqlClient.SqlParameter Then CType(PK2, SqlClient.SqlParameter).IsNullable = CType(P, SqlClient.SqlParameter).IsNullable
            If TypeOf P Is SQLite.SQLiteParameter Then CType(PK2, SQLite.SQLiteParameter).IsNullable = CType(P, SQLite.SQLiteParameter).IsNullable
            If TypeOf P Is OleDb.OleDbParameter Then CType(PK2, OleDb.OleDbParameter).IsNullable = CType(P, OleDb.OleDbParameter).IsNullable
            If TypeOf P Is Odbc.OdbcParameter Then CType(PK2, Odbc.OdbcParameter).IsNullable = CType(P, Odbc.OdbcParameter).IsNullable
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(PK2, SqlServerCe.SqlCeParameter).IsNullable = CType(P, SqlServerCe.SqlCeParameter).IsNullable
#End If
            PK2.ParameterName = "@PK2"
            PK2.Direction = P.Direction
            PK2.Precision = P.Precision
         End If
      Next

      If PK.ParameterName = "" Then
         If TypeOf PK Is SqlClient.SqlParameter Then CType(PK, SqlClient.SqlParameter).SqlDbType = SqlDbType.Int
         If TypeOf PK Is SQLite.SQLiteParameter Then CType(PK, SQLite.SQLiteParameter).DbType = DbType.Int32
         If TypeOf PK Is OleDb.OleDbParameter Then CType(PK, OleDb.OleDbParameter).OleDbType = OleDb.OleDbType.Integer
         If TypeOf PK Is Odbc.OdbcParameter Then CType(PK, Odbc.OdbcParameter).DbType = DbType.Int32
#If IncludeCompactFrameWork Then
            If TypeOf PK Is SqlServerCe.SqlCeParameter Then CType(PK, SqlServerCe.SqlCeParameter).DbType = SqlDbType.Int
#End If

         PK.SourceColumn = PKColumnName
         PK.SourceVersion = DataRowVersion.Original

         If TypeOf PK Is SqlClient.SqlParameter Then CType(PK, SqlClient.SqlParameter).IsNullable = False
         If TypeOf PK Is SQLite.SQLiteParameter Then CType(PK, SQLite.SQLiteParameter).IsNullable = False
         If TypeOf PK Is OleDb.OleDbParameter Then CType(PK, OleDb.OleDbParameter).IsNullable = False
         If TypeOf PK Is Odbc.OdbcParameter Then CType(PK, Odbc.OdbcParameter).IsNullable = False
#If IncludeCompactFrameWork Then
            If TypeOf PK Is SqlServerCe.SqlCeParameter Then CType(PK, SqlServerCe.SqlCeParameter).IsNullable = False
#End If
         PK.ParameterName = "@PK1"
         PK.Direction = ParameterDirection.Input
         PK.Precision = 10
      End If

      Upd = "UPDATE [" & TableName & "] Set " & Upd & " Where [" & PKColumnName & "] = " & PK.ParameterName
      If PKColumnName2 <> "" Then Upd = Upd & " And [" & PKColumnName2 & "] = " & PK2.ParameterName
      uCmd.CommandText = Upd

      ' Add Where PK
      uCmd.Parameters.Add(PK)
      If PKColumnName2 <> "" Then uCmd.Parameters.Add(PK2)

      Return uCmd
   End Function

   Public Function BuildDeleteCommand(ByVal DA As IDbDataAdapter, ByVal TableName As String, ByVal PKColumnName As String, Optional ByVal PKColumnName2 As String = "") As IDbCommand
      Dim P As IDbDataParameter, Del As String = "", PK As IDbDataParameter, PK2 As IDbDataParameter

      Dim dCmd As IDbCommand = NewCommand(Del, DA.SelectCommand.Connection)
      PK = NewParameter(dCmd)
      PK2 = NewParameter(dCmd)

      ' Default for autokeys
      PK.ParameterName = "@" & PKColumnName
      PK.SourceVersion = DataRowVersion.Original
      If TypeOf PK Is SqlClient.SqlParameter Then CType(PK, SqlClient.SqlParameter).SqlDbType = SqlDbType.Int
      If TypeOf PK Is SQLite.SQLiteParameter Then CType(PK, SQLite.SQLiteParameter).DbType = DbType.Int32
      If TypeOf PK Is OleDb.OleDbParameter Then CType(PK, OleDb.OleDbParameter).OleDbType = OleDb.OleDbType.Integer
      If TypeOf PK Is Odbc.OdbcParameter Then CType(PK, Odbc.OdbcParameter).DbType = DbType.Int32
#If IncludeCompactFrameWork Then
        If TypeOf PK Is SqlServerCe.SqlCeParameter Then CType(PK, SqlServerCe.SqlCeParameter).DbType = SqlDbType.Int
#End If
      PK.Size = 4
      If TypeOf PK Is SqlClient.SqlParameter Then CType(PK, SqlClient.SqlParameter).IsNullable = False
      If TypeOf PK Is SQLite.SQLiteParameter Then CType(PK, SQLite.SQLiteParameter).IsNullable = False
      If TypeOf PK Is OleDb.OleDbParameter Then CType(PK, OleDb.OleDbParameter).IsNullable = False
      If TypeOf PK Is Odbc.OdbcParameter Then CType(PK, Odbc.OdbcParameter).IsNullable = False
#If IncludeCompactFrameWork Then
        If TypeOf PK Is SqlServerCe.SqlCeParameter Then CType(PK, SqlServerCe.SqlCeParameter).IsNullable = False
#End If
      PK.Direction = ParameterDirection.Input
      PK.SourceColumn = PKColumnName

      For Each P In DA.InsertCommand.Parameters
         If P.SourceColumn = PKColumnName Then
            If TypeOf P Is SqlClient.SqlParameter Then CType(PK, SqlClient.SqlParameter).SqlDbType = CType(P, SqlClient.SqlParameter).SqlDbType
            If TypeOf P Is SQLite.SQLiteParameter Then CType(PK, SQLite.SQLiteParameter).DbType = CType(P, SQLite.SQLiteParameter).DbType
            If TypeOf P Is OleDb.OleDbParameter Then CType(PK, OleDb.OleDbParameter).OleDbType = CType(P, OleDb.OleDbParameter).OleDbType
            If TypeOf P Is Odbc.OdbcParameter Then CType(PK, Odbc.OdbcParameter).DbType = CType(P, Odbc.OdbcParameter).DbType
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(PK, SqlServerCe.SqlCeParameter).DbType = CType(P, SqlServerCe.SqlCeParameter).DbType
#End If
            PK.SourceColumn = P.SourceColumn
            PK.DbType = P.DbType
            PK.SourceVersion = DataRowVersion.Original
            PK.Size = P.Size

            If TypeOf P Is SqlClient.SqlParameter Then CType(PK, SqlClient.SqlParameter).IsNullable = CType(P, SqlClient.SqlParameter).IsNullable
            If TypeOf P Is SQLite.SQLiteParameter Then CType(PK, SQLite.SQLiteParameter).IsNullable = CType(P, SQLite.SQLiteParameter).IsNullable
            If TypeOf P Is OleDb.OleDbParameter Then CType(PK, OleDb.OleDbParameter).IsNullable = CType(P, OleDb.OleDbParameter).IsNullable
            If TypeOf P Is Odbc.OdbcParameter Then CType(PK, Odbc.OdbcParameter).IsNullable = CType(P, Odbc.OdbcParameter).IsNullable
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(PK, SqlServerCe.SqlCeParameter).IsNullable = CType(P, SqlServerCe.SqlCeParameter).IsNullable
#End If
            PK.ParameterName = P.ParameterName
            PK.Direction = P.Direction
            PK.Precision = P.Precision
         End If

         If P.SourceColumn = PKColumnName2 Then
            If TypeOf P Is SqlClient.SqlParameter Then CType(PK2, SqlClient.SqlParameter).SqlDbType = CType(P, SqlClient.SqlParameter).SqlDbType
            If TypeOf P Is SQLite.SQLiteParameter Then CType(PK2, SQLite.SQLiteParameter).DbType = CType(P, SQLite.SQLiteParameter).DbType
            If TypeOf P Is OleDb.OleDbParameter Then CType(PK2, OleDb.OleDbParameter).OleDbType = CType(P, OleDb.OleDbParameter).OleDbType
            If TypeOf P Is Odbc.OdbcParameter Then CType(PK2, Odbc.OdbcParameter).DbType = CType(P, Odbc.OdbcParameter).DbType
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(PK2, SqlServerCe.SqlCeParameter).DbType = CType(P, SqlServerCe.SqlCeParameter).DbType
#End If
            PK2.SourceColumn = P.SourceColumn
            PK2.DbType = P.DbType
            PK2.SourceVersion = DataRowVersion.Original
            PK2.Size = P.Size

            If TypeOf P Is SqlClient.SqlParameter Then CType(PK2, SqlClient.SqlParameter).IsNullable = CType(P, SqlClient.SqlParameter).IsNullable
            If TypeOf P Is SQLite.SQLiteParameter Then CType(PK2, SQLite.SQLiteParameter).IsNullable = CType(P, SQLite.SQLiteParameter).IsNullable
            If TypeOf P Is OleDb.OleDbParameter Then CType(PK2, OleDb.OleDbParameter).IsNullable = CType(P, OleDb.OleDbParameter).IsNullable
            If TypeOf P Is Odbc.OdbcParameter Then CType(PK2, Odbc.OdbcParameter).IsNullable = CType(P, Odbc.OdbcParameter).IsNullable
#If IncludeCompactFrameWork Then
                If TypeOf P Is SqlServerCe.SqlCeParameter Then CType(PK2, SqlServerCe.SqlCeParameter).IsNullable = CType(P, SqlServerCe.SqlCeParameter).IsNullable
#End If
            PK2.ParameterName = P.ParameterName
            PK2.Direction = P.Direction
            PK2.Precision = P.Precision
         End If
      Next

      Del = "Delete From [" & TableName & "] Where [" & PKColumnName & "] = " & PK.ParameterName
      If PKColumnName2 <> "" Then Del = Del & " And [" & PKColumnName2 & "] =" & PK2.ParameterName
      dCmd.CommandText = Del

      dCmd.Parameters.Add(PK)
      If PKColumnName2 <> "" Then dCmd.Parameters.Add(PK2)

      Return dCmd
   End Function

   Function XMLEncodeString(s As String) As String
      Dim r As String = ""
      Dim start As Integer = 1
      For i As Integer = 1 To Len(s)
         Dim b As Integer = Asc(Mid(s, i, 1))

         If b > 32 And b <> Asc("\") And b <> Asc("""") And b <> Asc("<") And b <> Asc(">") Then Continue For
         If start < i Then r &= Mid(s, start, i - start)

         Select Case b
            Case Asc("\"), Asc("""")
               r &= "\"
               r &= Chr(b)

            Case 10
               r &= "\n"

            Case 13
               r &= "\r"

            Case Else
               r &= "\u00" & Right("00" & Hex(b), 2)
         End Select

         start = i + 1
      Next

      If start <= Len(s) Then r &= Mid(s, start)

      Return r
   End Function

   Public Function Json2DataTable(ByVal JsonArray As ArrayList, ByVal TableName As String, ByRef Errors As String) As DataTable
      Dim mTable As New DataTable(TableName)
      Errors = ""
      If JsonArray.Count = 0 Then Return mTable

      Dim JSon As System.Collections.Generic.Dictionary(Of String, Object)
      JSon = JsonArray(0)

      ' Create Columns
      For Each keyPair As KeyValuePair(Of String, Object) In JSon
         If TypeOf keyPair.Value Is Integer Then
            mTable.Columns.Add(keyPair.Key, Type.GetType("System.Int32"))

         ElseIf TypeOf keyPair.Value Is Boolean Then
            mTable.Columns.Add(keyPair.Key, Type.GetType("System.Boolean"))

         Else
            mTable.Columns.Add(keyPair.Key, Type.GetType("System.String"))
         End If
      Next

      ' Add Data
      For Each JSon In JsonArray
         Dim R As DataRow = mTable.NewRow
         For Each keyPair As KeyValuePair(Of String, Object) In JSon
            Try
               R(keyPair.Key) = keyPair.Value
            Catch ex As Exception
               Errors = ex.Message
               Return mTable
            End Try
         Next
         mTable.Rows.Add(R)
      Next
      Return mTable
   End Function


End Module


Public Class XmlEntry
   Public Name As String
   Public InnerAttributeText As String
   Public innerText As String
   Public ChildNodes As Collection
   Public Attributes As New Dictionary(Of String, String)
   Public postText As String

   Function ParseXML(ByVal Xml As String, Optional ByVal StartP As Integer = 1) As Integer
      StartP = SkipWS(Xml, StartP)

      Dim C As String = Mid(Xml, StartP, 1)
      If C <> "<" Then Return 0

      Dim EndP As Integer = InStr(StartP, Xml, ">")
      If EndP = 0 Then Return 0

      InnerAttributeText = Mid(Xml, StartP + 1, EndP - StartP - 1)
      Me.Name = Field(InnerAttributeText, " ", 1)

      If Left(Me.Name, 3) = "!--" Then
         Me.Name = Left(Me.Name, 3)
         EndP = InStr(StartP, Xml, "-->")
         If EndP = 0 Then EndP = Len(Xml) + 1 Else EndP = EndP + 3
         InnerAttributeText = Mid(Xml, StartP + 4, EndP - StartP - 7)
         Return EndP
      End If

      InnerAttributeText = Mid(InnerAttributeText, Len(Me.Name) + 2)
      If Right(InnerAttributeText, 1) = "/" Then
         InnerAttributeText = Left(InnerAttributeText, Len(InnerAttributeText) - 1)
         Me.ParseAttributes(InnerAttributeText)
         Return EndP + 1
      End If

      Me.Attributes = Me.ParseAttributes(InnerAttributeText)

      ' Look for next XML entry
      Dim IStartP As Integer = EndP + 1 ' Beginning of <xxx>*INNER TEXT</xxx>

      EndP = InStr(IStartP, Xml, "</" & Me.Name)
      If EndP = 0 Then
         If Me.Name = "area" Or Me.Name = "base" Or Me.Name = "br" Or Me.Name = "col" Or Me.Name = "command" Or Me.Name = "embed" _
            Or Me.Name = "hr" Or Me.Name = "img" Or Me.Name = "input" Or Me.Name = "keygen" Or Me.Name = "link" Or Me.Name = "meta" _
            Or Me.Name = "param" Or Me.Name = "source" Or Me.Name = "track" Or Me.Name = "wbr" Or Me.Name = "!doctype" Then Return IStartP

         Throw New Exception("Missing closing tag for " & Me.Name & " at postion " & IStartP)
      End If

      If Me.Name = "script" Then

      Else

         EndP = InStr(IStartP, Xml, "<")
         If EndP = 0 Then EndP = Len(Xml) + 1
      End If

      Me.innerText = VB6.Replace(Mid(Xml, IStartP, EndP - IStartP), Chr(254), vbCrLf)

      C = Mid(Xml, EndP, 2)
      Do Until C = "</" Or C = ""
         ' We have children at NextStartP
         Me.ChildNodes = New Collection
         Do
            Dim ChildEntry As XmlEntry = New XmlEntry
            EndP = ChildEntry.ParseXML(Xml, EndP)
            If Len(ChildEntry.Name) = 0 Then Exit Do
            Me.ChildNodes.Add(ChildEntry)

            IStartP = EndP
            EndP = InStr(EndP, Xml, "<")
            If EndP = 0 Then EndP = Len(Xml) + 1

            C = Mid(Xml, EndP, 2)
            If C = "</" Or Left(C, 1) <> "<" Or C = "" Then Exit Do
            ChildEntry.postText = VB6.Replace(Mid(Xml, IStartP, EndP - IStartP), Chr(254), vbCrLf)
         Loop


      Loop

      If C = "</" Then
         C = Mid(Xml, EndP + 2, Len(Me.Name))
         If C <> Me.Name Then Throw New Exception("Poorly formed XML, expected closing </" & Me.Name & ", got " & Field(Mid(Xml, EndP + 2, 30), ">", 1) & " at position " & EndP)
      End If

      EndP = InStr(EndP, Xml, ">")
      If EndP = 0 Then
         EndP = Len(Xml) + 1
      Else
         EndP += 1
      End If

      Return EndP
   End Function

   Function ParseAttributes(ByVal InnerAttributeText As String) As Dictionary(Of String, String)
      Dim StartI As Integer = 1
      Dim EqualI As Integer
      Me.Attributes = New Dictionary(Of String, String)

      Do
         EqualI = InStr(StartI, InnerAttributeText, "=")
         If EqualI = 0 Then Exit Do

         Dim Name As String = Mid(InnerAttributeText, StartI, EqualI - StartI)

         Dim C As String = LTrim(Left(Mid(InnerAttributeText, EqualI + 1, 20), 1))
         If C <> "'" And C <> """" Then Exit Do

         Dim StartP As Integer = InStr(EqualI, InnerAttributeText, C)
         Dim EndP As Integer = InStr(StartP + 1, InnerAttributeText, C)
         If EndP = 0 Then EndP = Len(InnerAttributeText) + 1

         Dim Value As String = Mid(InnerAttributeText, StartP + 1, EndP - StartP - 1)

         If Me.Attributes.ContainsKey(LCase(Name)) = False Then Me.Attributes.Add(LCase(Name), Value)

         StartI = EndP + 1
         C = LTrim(Left(Mid(InnerAttributeText, StartI + 1, 20), 1))
         If C = "" Then Exit Do

         StartI = InStr(StartI, InnerAttributeText, C)
         If StartI = 0 Then Exit Do
      Loop

      Return Attributes
   End Function

   Function SkipWS(ByVal str As String, ByRef StartP As Integer) As Integer
      Do
         Dim C As String = Mid(str, StartP, 1)
         If C = "" Then Return StartP
         If InStr("*" & vbCrLf & " " & Chr(254) & Chr(9), C) = 0 Then Return StartP
         StartP += 1
      Loop
      Return StartP
   End Function

   Function AttributeValues() As String
      Dim Result As String = ""
      For Each Attribute In Attributes
         If InStr(Attribute.Value, "'") <> 0 Then
            Result &= " " & Attribute.Key & "=""" & VB6.Replace(Attribute.Value, """", "\""") & """"
         Else
            Result &= " " & Attribute.Key & "='" & VB6.Replace(Attribute.Value, "'", "\'") & "'"
         End If
      Next
      Return Result
   End Function

   Function AttributeNames() As String
      Dim Result As String = ""
      For Each Attribute In Attributes
         Result &= Chr(254) & Attribute.Key
      Next
      Return Mid(Result, 2)
   End Function

   Function AttributesValue(ByVal i As Integer) As String
      If i > Attributes.Count Then Return "" Else Return Attributes(i)
   End Function

   Function AttributesName(ByVal i As Integer) As String
      If i > Attributes.Count Then Return "" Else Return Attributes.Keys(i)
   End Function

   ' This is a shot-cut for .ChildNodes["xxx"][0]
   Function Child(ByVal name As String, Optional ByVal idx As Integer = 1) As XmlEntry
      If Me.ChildNodes Is Nothing Then Return Nothing

      If IsNumeric(name) Then
         idx = Val(name)
         If idx <= 0 Then idx = 1
         If idx > Count() Then Return Nothing
         Return Me.ChildNodes(idx)
      End If

      If idx <= 0 Then idx = 1
      Dim CArray As ArrayList = getXmlEntries(name)
      If idx > CArray.Count Then Return Nothing

      Return CArray(idx - 1)
   End Function

   Function InnerXML(Optional ByVal WithFormatting As Boolean = False) As String
      Dim Result As String = Me.innerText
      If WithFormatting Then
         Do
            Result = LTrim(VB6.Replace(Result, vbCrLf & " ", vbCrLf))
         Loop While InStr(Result, vbCrLf & " ") > 0
      End If

      If Me.ChildNodes IsNot Nothing Then
         Dim Spaces As String = ""
         Dim CRLF As String = ""

         If WithFormatting Then
            Spaces = "   "
            CRLF = vbCrLf
         End If

         For I As Integer = 1 To Me.ChildNodes.Count
            Dim CXML As XmlEntry = Me.ChildNodes(I)
            Dim OutXML As String = CXML.OuterXML(WithFormatting)
            If Left(OutXML, 2) <> CRLF Then OutXML = CRLF & OutXML
            Result &= OutXML & LTrim(CXML.postText)
         Next

         Result = VB6.Replace(Result, CRLF, CRLF & Spaces)
         If Left(Result, 2) = CRLF Then Result = Mid(Result, 2)

         If WithFormatting AndAlso Right(Result, 2) <> CRLF Then Result &= CRLF
      End If

      Return Result
   End Function

   Function Count() As Integer
      If Me.ChildNodes Is Nothing Then Return 0
      Return Me.ChildNodes.Count
   End Function

   Function OuterXML(Optional ByVal WithFormatting As Boolean = False) As String
      Dim Result As String = "<" & Me.Name & Left(" ", Len(Me.InnerAttributeText)) & Me.InnerAttributeText

      If Len(Me.innerText) = 0 AndAlso Me.ChildNodes Is Nothing Then
         If Me.Name = "!--" Then Result &= "-->" Else Result &= "/>"
      Else
         Result &= ">"
         Result &= Me.InnerXML(WithFormatting)
         Result &= "</" & Me.Name & ">"
      End If

      Return Result
   End Function

   Function getXmlEntries(ByVal Name As String) As ArrayList
      Dim lname As String = LCase(Name)
      Dim Results As New ArrayList
      If ChildNodes Is Nothing Then Return Results

      For Each v As XmlEntry In ChildNodes
         If LCase(v.Name) = lname Then Results.Add(v)
      Next

      Return Results
   End Function

End Class

