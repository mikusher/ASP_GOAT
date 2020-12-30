<%
' dbschema

Const DBF_PKEY = 1
Const DBF_NULL = 2
Const DBF_IDENTITY = 4

' data type definitions
Const DBT_BIGINT = 1
Const DBT_BINARY = 2
Const DBT_BIT = 3
Const DBT_BLOB = 4
Const DBT_CHAR = 5
Const DBT_DATE = 6
Const DBT_DATETIME = 7
Const DBT_DECIMAL = 8
Const DBT_DOUBLE = 9
Const DBT_ENUM = 10
Const DBT_FLOAT = 11
Const DBT_IMAGE = 12
Const DBT_INT = 13
Const DBT_MEDIUMINT = 14
Const DBT_MONEY = 15
Const DBT_NCHAR = 16
Const DBT_NTEXT = 17
Const DBT_NUMERIC = 18
Const DBT_NVARCHAR = 19
Const DBT_REAL = 20
Const DBT_SET = 21
Const DBT_SMALLINT = 22
Const DBT_SMALLDATETIME = 23
Const DBT_SMALLMONEY = 24
Const DBT_TEXT = 25
Const DBT_TIME = 26
Const DBT_TIMESTAMP = 27
Const DBT_TINYINT = 28
Const DBT_UNIQUEIDENTIFIER = 29
Const DBT_VARBINARY = 30
Const DBT_VARCHAR = 31
Const DBT_YEAR =32

'---------------------------------------------------------------------
' CLASS - dbField
'---------------------------------------------------------------------

Class dbField
	Public Name
	Public Datatype
	Public Attr
	Public Size
	Public Precision
	
	Public Sub Create(sName, nSize, nPrecision, nAttr)
		mName = sName
		Size = nSize
		Precision = nPrecision
		mAttr = nAttr
	End Sub

	' PROPERTY - IsKey
	Public Property IsKey
		IsKey = (mAttr And DBF_PKEY <> 0)
	End Property

	Public Property IsNullable
		IsNullable = (mAttr And DBF_NULL <> 0)
	End Property

	Public Property IsIdentity
		IsIdentity = (mAttr And DBF_IDENTITY <> 0)
	End Property
End Class

'---------------------------------------------------------------------
' CLASS - dbTable
'---------------------------------------------------------------------

Class dbTable
	Private maField()
	Private nFields
	Private sTableName

	' constructor
	Private Sub Class_Initialize
		nFields = 0
	End Sub

	' add a new field to the table definition
	Public Sub AddField(oField)
		ReDim Preserve maField(nFields)
		Set maField(nFields) = oField
		nFields = nFields + 1
	End Sub

	' remove a field definition for this table
	Public Sub RemoveField(nIndex)
		Dim I

		If nIndex < 0 Or nIndex > (nFields - 1) Then Exit Sub
		' move the fields down (if nec)
		For I = nIndex + 1 To nFields - 1
			Set maField(I-1) = maField(I)
		Next
		nFields = nFields - 1
	End Sub

	' retrieve a field for the table definition by name
	Public Function GetField(sFieldName)
		Dim I

		For I = 0 To nFields - 1
			If maField(I).Name = sFieldName Then
				Set FindField = maField(I)
				Exit Function
			End If
		Next
		Set FindField = Nothing
	End Function

	' PROPERTY - Tablename
	Public Property Let Tablename(sValue)
		sTablename = sValue
	End Property

	Public Property Get Tablename(sValue)
		Tablename = sTablename
	End Property

	' PROPERTY - Field (retrieve field definition at nIndex)
	Public Property Let Field(oField)
		AddField oField
	End Property

	Public Property Set Field(nIndex)
		If nIndex < 0 Or nIndex > (nFields - 1) Then
			Set Field = Nothing
			Exit Sub
		End If
		Set Field = maField(nIndex - 1)
	End Property
End Class

'---------------------------------------------------------------------
' CLASS - dbSchema
'---------------------------------------------------------------------

Class dbSchema
	Public maTable()
	Private nTables

	' constructor
	Private Sub Class_Initialize
		nTables = 0
	End Sub

	' add a new field to the table definition
	Public Sub AddTable(oTable)
		ReDim Preserve maTable(nTables)
		Set maTable(nTables) = oTable
		nTables = nTables + 1
	End Sub

	' remove a field definition for this table
	Public Sub RemoveTable(nIndex)
		Dim I

		If nIndex < 0 Or nIndex > (nTables - 1) Then Exit Sub
		' move the fields down (if nec)
		For I = nIndex + 1 To nTables - 1
			Set maTable(I-1) = maTable(I)
		Next
		nTables = nTables - 1
	End Sub

	' PROPERTY - Field (retrieve field definition at nIndex)
	Public Property Let Table(oTable)
		AddTable oTable
	End Property

	Public Property Set Table(nIndex)
		If nIndex < 0 Or nIndex > (nTables - 1) Then
			Set Table = Nothing
			Exit Sub
		End If
		Set Field = maTable(nIndex - 1)
	End Property
End Class

'---------------------------------------------------------------------
' CLASS - SQLServerParser
'---------------------------------------------------------------------

Class SQLServerParser
	Private msErrorMsg
	Private msPathname
	Private msContents
	Private mSchema

	Private Class_Initialize
		mSchema = New dbSchema
	End Class

	' read all contents from the local filesystem
	Private Function ReadAll
		Dim oFSO, oFile
		Const ForReading = 1
		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		On Error Resume Next
		Set oFile = oFSO.OpenTextFile(msFilename, ForReading)
		If Err.Number <> 0 Then
			msErrorMsg = "Cannot open file: """ & msFilename & """ - " & Err.Description
			ReadAll = False
			Exit Function
		End If
		msContents = oFile.ReadAll
		If Err.Number <> 0 Then
			msErrorMsg = "Cannot open file: """ & msFilename & """ - " & Err.Description
			ReadAll = False
			Exit Function
		End If
		oFile.Close
		Set oFile = Nothing
		Set oFSO = Nothing
	End Function

	' PROPERTY - Filename
	Public Property Let Filename(sValue)
		msPathname = sValue
	End Property

	' parse the data type and return the constant
	Public Function ParseDatatype(sDatatype)
		Select Case Trim(LCase(sDatatype))
			Case "bigint" : ParseDatatype = DBT_BIGINT
			Case "binary" : ParseDatatype = DBT_BINARY
			Case "bit" : ParseDatatype = DBT_BIT
			Case "blob" : ParseDatatype = DBT_BLOB
			Case "char" : ParseDatatype = DBT_CHAR
			Case "date" : ParseDatatype = DBT_DATE
			Case "datetime" : ParseDatatype = DBT_DATETIME
			Case "decimal" : ParseDatatype = DBT_DECIMAL
			Case "double" : ParseDatatype = DBT_DOUBLE
			Case "enum" : ParseDatatype = DBT_ENUM 
			Case "float" : ParseDatatype = DBT_FLOAT
			Case "image" : ParseDatatype = DBT_IMAGE 
			Case "int" : ParseDatatype = DBT_INT
			Case "mediumint" : ParseDatatype = DBT_MEDIUMINT
			Case "money" : ParseDatatype = DBT_MONEY 
			Case "nchar" : ParseDatatype = DBT_NCHAR
			Case "ntext" : ParseDatatype = DBT_NTEXT
			Case "numeric" : ParseDatatype = DBT_NUMERIC
			Case "nvarchar" : ParseDatatype = DBT_NVARCHAR
			Case "real" : ParseDatatype = DBT_REAL
			Case "set" : ParseDatatype = DBT_SET
			Case "smallint" : ParseDatatype = DBT_SMALLINT
			Case "smalldatetime" : ParseDatatype = DBT_SMALLDATETIME
			Case "smallmoney" : ParseDatatype = DBT_SMALLMONEY
			Case "text" : ParseDatatype = DBT_TEXT
			Case "time" : ParseDatatype = DBT_TIME
			Case "timestamp" : ParseDatatype = DBT_TIMESTAMP
			Case "tinyint" : ParseDatatype = DBT_TINYINT
			Case "uniqueidentifier" : ParseDatatype = DBT_UNIQUEIDENTIFIER
			Case "varbinary" : ParseDatatype = DBT_VARBINARY
			Case "varchar" : ParseDatatype = DBT_VARCHAR
			Case "year" : ParseDatatype = DBT_YEAR
		End Select
	End Function

	' parse a table definition
	Private Sub ParseTableDef(sTableName, sFieldDef)
		Dim re, oMatches, oMatch, oTable, oField, nAttr, aPair

		Set oTable = New dbTable
		oTable.Tablename = sTableName

		Set re = New RegExp
		re.Pattern = "\[(\w+)\]\s+\[(\w+)\]\s+(IDENTITY\s+\(\d+,\s*\d+\)\s+NOT FOR REPLICATION)?\((\d+)(,\s*\d+)?\)?\s*(COLLATE\s+\w+)?\s+(NOT NULL|NULL)"
		re.IgnoreCase
		re.Multiline = True
		re.Global = True
		Set oMatches = re.Execute(msContents)
		For Each oMatch In oMatches
			Set oField = New dbField
			oField.Name = oMatch.SubMatches(0)
			oField.Datatype = ParseDatatype(oMatch.SubMatches(1))
			If InStr(oMatch.SubMatches(2), "IDENTITY") > 0 Then nAttr = nAttr Or DBF_IDENTITY
			If oMatch.SubMatches(3) <> "" Then
				If InStr(oMatch.SubMatches(3), ",") Then
					aPair = Split(oMatch.SubMatches(3), ",")
					oField.Size = Trim(aPair(0))
					oField.Precision = Trim(aPair(1))
				Else
					oField.Size = Trim(oMatch.SubMatches(3))
				End If
			End If
			If Trim(oMatch.SubMatches(5)) <> "NOT NULL" Then nAttr = nAttr Or DBF_NULL
			oField.Attr = nAttr
			Call oTable.AddField oField
		Next
		' finally - add the new table to the schema
		mSchema.AddTable(oTable)
	End Sub

	' parse all of the table sections
	Private Sub ParseTables
		Dim re, oMatches, oMatch

		Set re = New RegExp
		re.Pattern = "CREATE\s+TABLE\s+\[dbo\]\.\[(\w+)\]\s+\((\s\S)+\)\s+ON\s+\[PRIMARY\]"
		re.IgnoreCase
		re.Multiline = True
		re.Global = True
		Set oMatches = re.Execute(msContents)
		For Each oMatch In oMatches
			Call ParseTableDef(oMatch.SubMatches(0), oMatch.SubMatches(1))
		Next
	End Sub

	' parse all of the table sections
	Private Sub ParseConstraints
		Dim re, oMatches, oMatch

		Set re = New RegExp
		re.Pattern = "ALTER\s+TABLE\s+\[dbo\]\.\[(\w+)\]\s+\((\s\S)+\)\nGO"
		re.IgnoreCase
		re.Multiline = True
		re.Global = True
		Set oMatches = re.Execute(msContents)
		For Each oMatch In oMatches
			Call ParseConstraintDef(oMatch.SubMatches(0), oMatch.SubMatches(1))
		Next
	End Sub

	' parse all of the contents for the database schema script
	Public Function ParseAll
		' make sure we have contents to parse
		If msContents = "" And msFilename <> "" Then
			If Not ReadAll Then
				ParseAll = False
				Exit Function
			End If
		End If
		If msContents = "" Then
			msErrorMsg = "No contents were found to parse, maybe you need to set the ""Filename"" property"
			ParseAll = False
			Exit Function
		End If
	End Function
End Class
%>