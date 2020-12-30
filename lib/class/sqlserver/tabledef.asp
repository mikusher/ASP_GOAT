Class clsTypeDef
	Private mobjName
	Private mobjLength
	Private mobjPrec
	Private mobjScale
	Private mobjXType

	Private Sub Class_Initialize
		mobjName = Server.CreateObject("Scripting.Dictionary")		
		mobjLength = Server.CreateObject("Scripting.Dictionary")
		mobjPrec = Server.CreateObject("Scripting.Dictionary")
		mobjScale = Server.CreateObject("Scripting.Dictionary")
		mobjXType = Server.CreateObject("Scripting.Dictionary")
	End Sub

	'----------------------------------------------------------------
	' Load the system type configuration from the database

	Public Sub LoadFromDatabase
		Dim sStat, rsType, nXType

		sStat = "SELECT name, xtype, length, prec, scale " &_
				"FROM	systypes"
		Set rsType = adoOpenRecordset(sStat)
		Do Until rsType.EOF
			nXType = CInt(rsType.Fields("xtype").Value)
			mobjXType(rsType.Fields("name").Value) = nXType
			mobjName(nXType) = rsType.Fields("name").Value
			mobjLength(nXType) = rsType.Fields("length").Value
			mobjPrec(nXType) = rsType.Fields("prec").Value
			mobjScale(nXType) = rsType.Fields("scale").Value
			rsType.MoveNext
		Loop
	End Sub

	'----------------------------------------------------------------
	' XType

	Public Function XType(sTypeName)
		XType = mobjXType.Item(sTypeName)
	End Function

	'----------------------------------------------------------------
	' Name - name of the data type

	Public Function Name(nXType)
		Name = mobjName.Item(nXType)
	End Function

	'----------------------------------------------------------------
	' Length - maximum length for this data type

	Public Function Length(nXType)
		Length = mobjLength.Item(nXType)
	End Function

	'----------------------------------------------------------------
	' Precision - maximum precision for this data type

	Public Function Precision(nXType)
		Precision = mobjPrec.Item(nXType)
	End Function

	'----------------------------------------------------------------
	' Scale - maximum scale for this data type

	Public Function Scale(nXType)
		Scale = mobjScale.Item(nXType)
	End Function
End Class

Class clsColumnDef
	Private mstrName		' column name
	Private mintID			' object ID
	Private mintXType		' xtype
	Private mintPrec		' precision for the column
	Private mintScale		' scale for the column
	Private mboolNull		' allows null values?

	Private Sub Class_Initialize
		mstrName = ""
		mintID = 0
		mintXType = 0
		mintPrec = 0
		mintScale = 0
		mboolNull = False
	End Sub

	'----------------------------------------------------------------
	' Name PROPERTY

	Public Property Get Name
		Name = mstrName
	End Property

	Public Property Let Name(sValue)
		mstrName = sValue
	End Property

	'----------------------------------------------------------------
	' ID PROPERTY - object ID to which column belongs

	Public Property Get ID
		ID = mintID
	End Property

	Public Property Let ID(nValue)
		mintID = nValue
	End Property

	'----------------------------------------------------------------
	' XType PROPERTY - object ID to which column belongs

	Public Property Get XType
		XType = mintXType
	End Property

	Public Property Let XType(nValue)
		mintXType = nValue
	End Property

	'----------------------------------------------------------------
	' Prec PROPERTY - Precision for the column

	Public Property Get Precision
		Precision = mintPrec
	End Property

	Public Property Let Precision(nValue)
		mintPrec = nValue
	End Property

	'----------------------------------------------------------------
	' Scale PROPERTY - Scale for the column

	Public Property Get Scale
		Scale = mintScale
	End Property

	Public Property Let Scale(nValue)
		mintScale = nValue
	End Property

	'----------------------------------------------------------------
	' XType PROPERTY - object ID to which column belongs

	Public Property Get XType
		XType = mintXType
	End Property

	Public Property Let XType(nValue)
		mintXType = nValue
	End Property

	'----------------------------------------------------------------
	' IsNullable PROPERTY - allows null values?

	Public Property Get IsNullable
		IsNullable = mboolNull
	End Property

	Public Property Let IsNullable(boolValue)
		mboolNull = boolValue
	End Property
End Class

Class clsTableDef
	Private mstrName
	Private mstrError
	Private marrCol(0)
	Private mintCols
	Private mstrColList
	Private mobjType
	Private FSO_FORREADING

	Private Sub Class_Initialize
		ReDim marrCol(10)
		mintCols = 0
		mstrColList = ""
		FSO_FORREADING = 1
		Set mobjType = Nothing
	End Sub

	'----------------------------------------------------------------
	' Retrieve the column definitions for the SQL Server table

	Private Function RetrieveColumns
		Dim sStat, rsCol, oCol

		sStat = "select colorder, name, id, xtype, prec, scale, isnullable " &_
				"from	syscolumns " &_
				"where id = object_id('" & Replace(mstrName, "'", "''") & "')"
		Set rsCol = adoOpenRecordset(sStat)
		Do Until rsCol.EOF
			If mintCols > UBound(marrCol) Then
				ReDim Preserve marrCol(UBound(marrCol) + 10)
			End If
			Set oCol = New clsColumnDef
			oCol.Name = rsCol.Fields("name").Value
			oCol.ID = rsCol.Fields("ID").Value
			oCol.XType = rsCol.Fields("xtype").Value
			oCol.Precision = rsCol.Fields("prec").Value
			oCol.Scale = rsCol.Fields("scale").Value
			oCol.IsNullable = rsCol.Fields("isnullable").Value
			Set marrCol(mintCols) = oCol
			mintCols = mintCols + 1
			' build the list of column names (to keep the order)
			If mstrColList <> "" Then mstrColList = mstrColList & ","
			mstrColList = mstrColList & "," & rsCol.Fields("name").Value
			rsCol.MoveNext
		Loop
	End Function

	'----------------------------------------------------------------
	' Retrieve the table definition from the SQL Server

	Public Function Retrieve
		' make sure the name property is defined
		If mstrName = "" Then
			mstrError = "Unable to retrieve table definition - Name property undefined"
			Retrieve = False
			Exit Function
		End If
	End Sub

	'----------------------------------------------------------------
	' Retrieve the column definition for the specified column
	' RETURNS: True if column definition was found, false otherwise

	Public Function ColumnDef(sName, nXType, nPrecision, nScale, bIsNullable)
		Dim I, oCol

		For I = 0 To mintCols - 1
			Set oCol = marrCol(I)
			If oCol.Name = sName Then
				sName = oCol.Name
				nXType = oCol.XType
				nPrecision = oCol.Precision
				nScale = oCol.Scale
				bIsNullable = oCol.IsNullable
				ColumnDef = True
				Exit Function
			End If
		Next
		ColumnDef = False
	End Function

	' -------------------------------------------------------------------------
	' store an entire text file and return it
	
	Private Function StoreFile(sPathName, sContents)
		Const bOverwrite = True
		Dim oFSO, oFile
	
		' make sure the path exists first
		'If Not BuildPath(sPathName) Then
		'	StoreFile = False
		'	Exit Function
		'End If
		' now try storing the file
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		On Error Resume Next
		Set oFile = oFSO.CreateTextFile(Server.MapPath(sPathName), bOverwrite)
		If Err.Number <> 0 Then
			mstrError = "StoreFile 1 - " & Err.Number & " - " & Err.Description
			StoreFile = False
			Exit Function
		End If
		oFile.Write(sContents)
		If Err.Number <> 0 Then
			mstrError = "StoreFile 2 - " & Err.Number & " - " & Err.Description
			StoreFile = False
			Exit Function
		End If
		oFile.Close
		If Err.Number <> 0 Then
			mstrError = "StoreFile 3 - " & Err.Number & " - " & Err.Description
			StoreFile = False
			Exit Function
		End If
		On Error Goto 0
		StoreFile = True
	End Function

	' -------------------------------------------------------------------------
	' read an entire text file and puts the contents in arg 2 (sContents)
	' RETURNS: true on success, false otherwise
	
	Function RetrieveFile(sPathName, sContents, dtModified)
		Dim oFSO, oFile
	
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		If (oFSO.FileExists(Server.MapPath(sPathName))) Then
			On Error Resume Next
			Set oFile = oFSO.GetFile(Server.MapPath(sPathName))
			dtModified = oFile.DateLastModified
			If Err.Number <> 0 Then
				mstrError = "RetrieveFile - " & Err.Number & " - " & Err.Description
				RetrieveFile = False
				Exit Function
			End If		
			Set oFile = oFSO.OpenTextFile(Server.MapPath(sPathName), FSO_FORREADING)
			If Err.Number <> 0 Then
				mstrError = "RetrieveFile - " & Err.Number & " - " & Err.Description
				RetrieveFile = False
				Exit Function
			End If
			sContents = oFile.ReadAll
			If Err.Number <> 0 Then
				mstrError = "RetrieveFile - " & Err.Number & " - " & Err.Description
				RetrieveFile = False
				Exit Function
			End If
			On Error Goto 0
		Else
			dtModified = Now()
			sContents = ""
		End If
		Set oFSO = Nothing
		RetrieveFile = True
	End Function

	'----------------------------------------------------------------
	' Convert the definition to a string - to serialize to a file

	Public Function Serialize
		Dim oTxt, aCol

		' load the SQL server system type definitions
		If mobjType Is Nothing Then
			Set mobjType = New clsTypeDef
			mobjType.LoadFromDatabase
		End If
		' serialize the table information
		Set oTxt = New clsString
		oTxt.Add "CREATE TABLE "
		oTxt.Add mstrName
		oTxt.Add " ("
		oTxt.Add vbCrLf
		If mstrColList <> "" Then
			aCol = Split(mstrColList, ",")
			For I = 0 To UBound(aCol)
				For J = 0 To mintCols - 1
					Set oCol = marrCol(J)
					If oCol.Name = aCol(I) Then
						oTxt.Add "  "
						oTxt.Add oCol.Name
						nXType = mobjType.Name(oCol.XType)

						If (oCol.Precision > 0) Then
							oTxt.Add " ("
							oTxt.Add oCol.Precision
							If oCol.Scale > 0 Then
								oTxt.Add ","
								oTxt.Add oCol.Scale
							End If
							oTxt.Add ")"
							If oCol.IsNullable Then oTxt.Add " NULL"
						End If
						oTxt.Add vbCrLf
					End If
				Next
			Next
		End If
		oTxt.Add ")"
		oTxt.Add vbCrLf
		Serialize = oTxt.Value
	End Function

	'----------------------------------------------------------------
	' Parse the definition from a string - to deserialize from file

	Public Function Unserialize(sText)
		Dim oTxt, aCol, oCol, oReg, oMatch, oMatches, aPrec

		mstrName = ""
		mintCols = 0
		Set oReg = New RegExp
		oReg.Pattern = "\s+CREATE\s+TABLE\s+\(\w+)\s+\(([^\)])\)"
		oReg.Global = True
		oReg.IgnoreCase = True
		Set oMatches = oReg.Execute(sText)
		For Each oMatch In oMatches
			mstrName = oMatch.SubMatches(0)
			sColumns = oMatch.SubMatches(1)
			Exit For
		Next
		' abort if the table name was not found
		If mstrName = "" Then
			mstrError = "Unable to unserialize table definition - table name not found"
			Unserialize = False
			Exit Function
		End If
		' parse the column definitions
		oReg.Pattern = "\s+(\w+)\s+(\w+)(\s+\(\d+,?\d+?\))?(\s+NULL)?"
		Set oMatches = oReg.Execute(sColumns)
		For Each oMatch In oMatches
			If mintCols > UBound(marrCol) Then
				ReDim Preserve marrCol(UBound(marrCol) + 10)
			End If
			Set oCol = New clsColumnDef
			oCol.Name = oMatch.SubMatches(0)
			' oCol.ID = rsCol.Fields("ID").Value
			oCol.XType = mobjType.XType(oMatch.SubMatches(1))
			If oMatch.SubMatches(2) <> "" Then
				If InStr(oMatch.SubMatches(2), ",") > 0 Then
					aPrec = Split(Replace(Replace(oMatch.SubMatches(2), "(", ""), ")", ""), ",")
					oCol.Precision = aPrec(0)
					oCol.Scale = aPrec(1)
				Else
					oCol.Precision = oMatch.SubMatches(2)
					oCol.Scale = 0
				End If
			End If
			If oMatch.SubMatches(3) <> "" Then
				oCol.IsNullable = True
			Else
				oCol.IsNullable = False
			End If
			Set marrCol(mintCols) = oCol
			mintCols = mintCols + 1
		Next
	End Function

	'----------------------------------------------------------------
	' Name PROPERTY - name of the table

	Public Property Get Name
		Name = mstrName
	End Property

	Public Property Let Name(sValue)
		mstrName = sValue
	End Property

	'----------------------------------------------------------------
	' ColumnList PROPERTY - ordered list of column names

	Public Property Get ColumnList
		ColumnList = mstrColList
	End Property

End Class
