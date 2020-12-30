<%
' rsspub_lib.asp
'	Library for creating RSS (syndicated news feeds)

Class clsTemplate
	Private mobjFSO			' File System Object
	Private mdtModified		' Date file was last modified
	Private mstrContents	' contents of the template

	Const FSO_FORREADING = 1
	Const FSO_FORWRITING = 2

	'------------------------------------------------------------------------
	' Constructor for the template class

	Public Sub Class_Initialize
		Set mobjFSO = Nothing
	End Sub

	'------------------------------------------------------------------------
	' Read the contents of the template located at sPathname

	Public Function ReadTemplate(sPathname)
		Dim oFile

		If mobjFSO Is Nothing Then
			Set mobjFSO = Server.CreateObject("Scripting.FileSystemObject")
		End If
		On Error Resume Next
		Set oFile = oFSO.GetFile(Server.MapPath(sPathName))
		mdtModified = oFile.DateLastModified
		If Err.Number <> 0 Then
			mstrError = "ReadTemplate - " & Err.Number & " - " & Err.Description
			ReadTemplate = False
			Exit Function
		End If		
		Set oFile = oFSO.OpenTextFile(Server.MapPath(sPathName), FSO_FORREADING)
		If Err.Number <> 0 Then
			mstrError = "ReadTemplate - " & Err.Number & " - " & Err.Description
			ReadTemplate = False
			Exit Function
		End If
		mstrContents = oFile.ReadAll
		If Err.Number <> 0 Then
			mstrError = "ReadTemplate - " & Err.Number & " - " & Err.Description
			ReadTemplate = False
			Exit Function
		End If
		On Error Goto 0
		ReadTemplate = True
	End Function

	'------------------------------------------------------------------------
	' Write the contents of a file to the local filesystem

	Public Function WriteTemplate(sPathname)
		Dim oFile

		If mobjFSO Is Nothing Then
			Set mobjFSO = Server.CreateObject("Scripting.FileSystemObject")
		End If

		On Error Resume Next
		Set oFile = oFSO.CreateTextFile(Server.MapPath(sPathName), True)
		If Err.Number <> 0 Then
			mstrError = "WriteTemplate 1 - " & Err.Number & " - " & Err.Description
			WriteTemplate = False
			Exit Function
		End If
		oFile.Write(mstrContents)
		If Err.Number <> 0 Then
			mstrError = "WriteTemplate 2 - " & Err.Number & " - " & Err.Description
			WriteTemplate = False
			Exit Function
		End If
		oFile.Close
		If Err.Number <> 0 Then
			mstrError = "WriteTemplate 3 - " & Err.Number & " - " & Err.Description
			WriteTemplate = False
			Exit Function
		End If
		On Error Goto 0
		WriteTemplate = True
	End Function

	'------------------------------------------------------------------------
	' MacroSub

	Public Sub MacroSub(sMacroName, sValue)
		mstrContents = Replace(mstrContents, "##" & sMacroName & "##", sValue, 1, -1, vbTextCompare)
	End Sub

	'------------------------------------------------------------------------
	' PROPERTY - Contents

	Public Property Let Contents(strValue)
		mstrContents = strValue
	End Property

	Public Property Get Contents
		Contents = mstrContents
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - ErrorMsg

	Public Property Get ErrorMsg
		ErrorMsg = mstrError
	End Property
End Class

Class clsRDFArticle
	Public About ' As String
	Public Title ' As String
	Public Link ' As String
	Public Description ' As String
	Public DC_Creator ' As String
	Public DC_Subject ' As String
	Public DC_Date ' As String
	Public Slash_Section ' As String
	Public Slash_Comments ' As String
	Public Slash_HitParade ' As String
End Class

Class clsRDF
	Private marrArticle(0)		' array of articles
	Private mintArticles		' total articles
	Private mstrRSSTemplate		' RSS Template to create RSS feed
	Private mstrRSSPath			' Path where the RSS feed is written
	Private mstrError			' Error to report to the user

	Private mstrRDF_About		' URL w/info about this news feed
	Private mstrRDF_Title		' Title of the web site (publisher)
	Private mstrRDF_Link		' Link to the article overview web page
	Private mstrRDF_Description	' Description of the web page
	Private mstrRDF_LI			' List items for ordering articles

	Private mstrRDF_ResourceImage	' small logo image associated w/feed
	Private mstrRDF_ResourceTextInput	' web page w/search box for articles

	Private mstrDC_Rights		' copyright notice
	Private mstrDC_Creator		' creator (author e-mail address)
	Private mstrDC_Publisher	' publisher (name of company supplying feed)
	Private mstrDC_Subject		' subject material (genre or category)

	Private mstrSYN_UpdatePeriod ' how often feed is refreshed
	Private mstrSYN_UpdateFrequency ' how often feed is refreshed each period???

	Private mstrImage_RDFAbout	' about for the web site image
	Private mstrImage_Title		' title for the web page image
	Private mstrImage_URL		' URL for the web page image
	Private mstrImage_Link		' Link to the web site
	
	'------------------------------------------------------------------------
	' Constructor for the template class

	Public Sub Class_Initialize
		Set mobjFSO = Nothing
		mstrSyn_UpdatePeriod = "hourly"
		mstrSyn_UpdateFrequency = "1"
		mstrRSSPath = "/rss"
		mintArticles = 0
	End Sub

	'------------------------------------------------------------------------
	' XML encode a text string

	Private Function XMLEncode(sText)
		XMLEncode = Replace(Replace(Replace(Replace(Replace(sText, "&", "&amp;"), "'", "&apos;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;")
	End Function

	'------------------------------------------------------------------------
	' BuildArticleXML

	Private Function BuildArticleXML
		Dim sXML, oArt, I

		For I = 0 To mintArticles - 1
			Set oArt = marrArticle(I)
			sXML = sXML & "<item rdf:about=""" & XMLEncode(oArt.About) & """>" &_
				"<title>" & XMLEncode(oArt.Title) & "</title>" &_
				"<link>" & XMLEncode(oArt.Link) & "</link>" &_
				"<description>" & XMLEncode(oArt.Descripton) & "</description>" &_
				"<dc:creator>" & XMLEncode(oArt.DC_Creator) & "</dc:creator>" &_
				"<dc:subject>" & XMLEncode(oArt.DC_Subject) & "</dc:subject>" &_
				"<dc:date>" & XMLEncode(oArt.DC_Date) & "</dc:date>" &_
				"<slash:section>" & XMLEncode(oArt.Slash_Section) & "</slash:section>" &_
				"<slash:comments>" & XMLEncode(oArt.Slash_Comments) & "</slash:comments>" &_
				"<slash:hitparade>" & XMLEncode(oArt.Slash_HitParade) & "</slash:hitparade>" &_
				"</item>" & vbCrLf
		Next
		BuildArticleXML = sXML
	End Function

	'------------------------------------------------------------------------
	' Generate the current date and time in UTF (Universal Time Format)
	' like: "2003-11-26T21:13:06+08:00" for Pacific Standard Time

	Private Function CurrentUTFDateTime
		CurrentUTFDateTime = Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) &_
			"T" & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & "+08:00"
	End Function
	'------------------------------------------------------------------------
	' Write the RSS output file to the local filesystem
	' (requires write permission for the web server to the RSS directory)

	Public Function Publish
		Dim oTemplate

		' retrieve the template file
		If mstrRSSTemplate = "" Then
			mstrError = "clsRDF - You must define the RSSTemplate property"
			Publish = False
			Exit Function
		End If
		Set oTemplate = New clsTemplate
		If Not oTemplate.ReadTemplate(mstrRSSTemplate) Then
			mstrError = oTemplate.ErrorMsg
			Publish = False
			Exit Function
		End If

		' replace the macros inside the template file
		oTemplate.MacroSub "RDF:About", XMLEncode(mstrRDF_About)		' URL w/info about this news feed
		oTemplate.MacroSub "RDF:Title", XMLEncode(mstrRDF_Title)		' Title of company publishing the feed
		oTemplate.MacroSub "RDF:Link", XMLEncode(mstrRDF_Link)			' Link to the web site (article overview)
		oTemplate MacroSub "RDF:Description", XMLEncode(mstrRDF_Description)	' description of the web site
		oTemplate.MacroSub "RDF:ResourceImage", XMLEncode(mstrRDF_ResourceImage)	' small logo image associated w/feed
		oTemplate.MacroSub "RDF:ResourceTextInput", XMLEncode(mstrRDF_ResourceTextInput)	' web page w/search box for articles
		oTemplate.MacroSub "RDF:LI", mstrRDF_LI				' Article listing

		oTemplate.MacroSub "DC:Date", CurrentUTFDateTime				' date RSS feed was written (in UTF)
		oTemplate.MacroSub "DC:Rights", XMLEncode(mstrDC_Rights)		' copyright notice
		oTemplate.MacroSub "DC:Creator", XMLEncode(mstrDC_Creator)		' creator (author e-mail address)
		oTemplate.MacroSub "DC:Publisher", XMLEncode(mstrDC_Publisher)	' publisher (name of company supplying feed)
		oTemplate.MacroSub "DC:Subject", XMLEncode(mstrDC_Subject)		' subject material (genre or category)
	
		oTemplate.MacroSub "SYN:UpdatePeriod", XMLEncode(mstrSYN_UpdatePeriod) ' how often feed is refreshed
		oTemplate.MacroSub "SYN:UpdateFrequency", XMLEncode(mstrSYN_UpdateFrequency)

		oTemplate.MacroSub "IMAGE:RDFAbout", XMLEncode(mstrImage_RDFAbout)
		oTemplate.MacroSub "IMAGE:Title", XMLEncode(mstrImage_Title)
		oTemplate.MacroSub "IMAGE:URL", XMLEncode(mstrImage_URL)
		oTemplate.MacroSub "IMAGE:Link", XMLEncode(mstrImage_Link)
		oTemplate.MacroSub "ArticleList", BuildArticleXML

		' write the file to the local filesystem
		If Not oTemplate.WriteTemplate(mstrRSSTemplate) Then
			mstrError = oTemplate.ErrorMsg
			Publish = False
			Exit Function
		End If
		Publish = True
	End Function

	'------------------------------------------------------------------------
	' AddArticle

	Public Sub AddArticle(sAbout, sTitle, sLink, sDescription, sDCCreator, sDCSubject, _
		sDCDate, sSlashSection, sSlashComments, sSlashHitParade)
		Dim oArt

		Set oArt = New clsRDFArticle
		oArt.About = sAbout		' URL to the info about the article (or article itself)
		oArt.Title = sTitle 	' Title for the article
		oArt.Link = sLink		' URL to the article
		oArt.Description		' Synopsis of the article contents
		oArt.DC_Creator			' Author who created the article
		oArt.DC_Subject			' Category that the article was placed in
		oArt.DC_Date			' Date when article was first published
		oArt.Slash_Section		' SlashCode section where article should appear
		oArt.Slash_Comments		' Number of reader comments attached
		oArt.Slash_HitParade	' ???
		If mintArticles > UBound(marrArticle) Then
			ReDim Preserve marrArticle(UBound(marrArticle) + 10)
		End If
		Set marrArticle(mintArticles) = oArt
		mintArticles = mintArticles + 1

		' add to the article list "RDF:Seq ==> RDF:LI"
		mstrRDF_LI = mstrRDF_LI & "<rdf:li rdf:resource=""" & XMLEncode(oArt.Link) & """ />" & vbCrLf
	End Sub

	'------------------------------------------------------------------------
	' PROPERTY - RDF:About

	Public Property Let RDF_About(strValue)
		mstrRDF_About = strValue
	End Property

	Public Property Get RDF_About
		RDF_About = mstrRDF_About
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - RDF:Title

	Public Property Let RDF_Title(strValue)
		mstrRDF_Title = strValue
	End Property

	Public Property Get RDF_Title
		RDF_Title = mstrRDF_Title
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - RDF:Link

	Public Property Let RDF_Link(strValue)
		mstrRDF_Link = strValue
	End Property

	Public Property Get RDF_Link
		RDF_Link = mstrRDF_Link
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - image RDF:Resource

	Public Property Let RDF_ResourceImage(strValue)
		mstrRDF_ResourceImage = strValue
	End Property

	Public Property Get RDF_ResourceImage
		RDF_ResourceImage = mstrRDF_ResourceImage
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - textinput RDF:Resource

	Public Property Let RDF_ResourceTextInput(strValue)
		mstrRDF_ResourceTextInput = strValue
	End Property

	Public Property Get RDF_ResourceTextInput
		RDF_ResourceTextInput = mstrRDF_ResourceTextInput
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - DC:Rights

	Public Property Let DC_Rights(strValue)
		mstrDC_Rights = strValue
	End Property

	Public Property Get DC_Rights
		DC_Rights = mstrDC_Rights
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - DC:Creator

	Public Property Let DC_Creator(strValue)
		mstrDC_Creator = strValue
	End Property

	Public Property Get DC_Creator
		DC_Creator = mstrDC_Creator
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - DC:Publisher

	Public Property Let DC_Publisher(strValue)
		mstrDC_Publisher = strValue
	End Property

	Public Property Get DC_Publisher
		DC_Publisher = mstrDC_Publisher
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - DC:Subject

	Public Property Let DC_Subject(strValue)
		mstrDC_Subject = strValue
	End Property

	Public Property Get DC_Subject
		DC_Subject = mstrDC_Subject
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - SYN:UpdatePeriod

	Public Property Let SYN_UpdatePeriod(strValue)
		mstrSYN_UpdatePeriod = strValue
	End Property

	Public Property Get SYN_UpdatePeriod
		SYN_UpdatePeriod = mstrSYN_UpdatePeriod
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - SYN:UpdateFrequency

	Public Property Let SYN_UpdateFrequency(strValue)
		mstrSYN_UpdateFrequency = strValue
	End Property

	Public Property Get SYN_UpdateFrequency
		SYN_UpdateFrequency = mstrSYN_UpdateFrequency
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - Image RDF:About

	Public Property Let Image_RDFAbout(strValue)
		mstrImage_RDFAbout = strValue
	End Property

	Public Property Get Image_RDFAbout
		Image_RDFAbout = mstrImage_RDFAbout
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - Image Title

	Public Property Let Image_Title(strValue)
		mstrImage_Title = strValue
	End Property

	Public Property Get Image_Title
		Image_Title = mstrImage_Title
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - Image URL

	Public Property Let Image_URL(strValue)
		mstrImage_URL = strValue
	End Property

	Public Property Get Image_URL
		Image_URL = mstrImage_URL
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - Image Link

	Public Property Let Image_Link(strValue)
		mstrImage_Link = strValue
	End Property

	Public Property Get Image_Link
		Image_Link = mstrImage_Link
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - RSSTemplate

	Public Property Let RSSTemplate(strValue)
		mstrRSSTemplate = strValue
	End Property

	Public Property Get RSSTemplate
		RSSTemplate = mstrRSSTemplate
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - RSSPath

	Public Property Let RSSPath(strValue)
		mstrRSSPath = strValue
	End Property

	Public Property Get RSSPath
		RSSPath = mstrRSSPath
	End Property

	'------------------------------------------------------------------------
	' PROPERTY - ErrorMsg

	Public Property Get ErrorMsg
		ErrorMsg = mstrError
	End Property
End Class
%>