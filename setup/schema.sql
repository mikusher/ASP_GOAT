CREATE TABLE [dbo].[tblApplicationVar] (
	[VarID] [int] IDENTITY (1, 1) NOT NULL ,
	[VarName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[VarValue] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Label] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[HelpText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TabID] [int] NULL ,
	[IsRequired] [bit] NOT NULL ,
	[TypeID] [int] NULL ,
	[HasOptions] [bit] NOT NULL ,
	[MinValue] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MaxValue] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OrderNo] [smallint] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblApplicationVarOption] (
	[OptionID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[TypeID] [int] NULL ,
	[VarID] [int] NULL ,
	[OptionValue] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OptionLabel] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParentOptionID] [int] NOT NULL ,
	[IsValid] [bit] NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblApplicationVarTab] (
	[TabID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[TabName] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Introduction] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Summary] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Archive] [bit] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[OrderNo] [smallint] NOT NULL ,
	[Title] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblApplicationVarType] (
	[TypeID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[TypeCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TypeName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ASPConvertFunction] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HTMLInputType] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RegExValidate] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LabelPos] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OrderNo] [smallint] NOT NULL ,
	[HasOptions] [bit] NOT NULL ,
	[MinValue] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MaxValue] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsNumeric] [bit] NOT NULL ,
	[QuoteChar] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Archive] [bit] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblArticle] (
	[ArticleID] [int] IDENTITY (1, 1) NOT NULL ,
	[MagazineID] [int] NULL ,
	[AuthorID] [int] NOT NULL ,
	[Title] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LeadIn] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ArticleBody] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ShortComments] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[WordCount] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[CommentCount] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblArticleAuthor] (
	[AuthorID] [int] IDENTITY (1, 1) NOT NULL ,
	[Title] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Firstname] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Middlename] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Lastname] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Surname] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblArticleCategory] (
	[CategoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[ParentCategoryID] [int] NOT NULL ,
	[CategoryName] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[IconImage] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblArticleComment] (
	[CommentID] [int] IDENTITY (1, 1) NOT NULL ,
	[ArticleID] [int] NOT NULL ,
	[MemberID] [int] NULL ,
	[ParentCommentID] [int] NOT NULL ,
	[Subject] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Body] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ModPoints] [smallint] NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblArticleToCategory] (
	[ArticleID] [int] NOT NULL ,
	[CategoryID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblContactUs] (
	[ContactUsID] [int] IDENTITY (1, 1) NOT NULL ,
	[FirstName] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LastName] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address1] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address2] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[City] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StateCode] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ZipCode] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CountryID] [int] NULL ,
	[Email] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MailingList] [bit] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblCountry] (
	[CountryID] [int] IDENTITY (1, 1) NOT NULL ,
	[CountryName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SortOrder] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblDoc] (
	[DocID] [int] IDENTITY (1, 1) NOT NULL ,
	[AuthorID] [int] NOT NULL ,
	[FolderID] [int] NOT NULL ,
	[Title] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubTitle] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ShortDescription] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Body] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[BookID] [int] NULL ,
	[ParentDocID] [int] NOT NULL ,
	[SectionName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsInlineContent] [bit] NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[SectionNo] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AuthorNotes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TypeID] [int] NULL ,
	[ScriptName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblDocAuthor] (
	[AuthorID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserID] [int] NULL ,
	[Title] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FirstName] [varchar] (24) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MiddleName] [varchar] (24) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LastName] [varchar] (24) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EmailAddress] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Biography] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[Surname] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblDocBook] (
	[BookID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[FolderID] [int] NOT NULL ,
	[Title] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubTitle] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AuthorID] [int] NOT NULL ,
	[Version] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PublishDate] [datetime] NULL ,
	[ShowSectionNo] [bit] NOT NULL ,
	[AuthorNotes] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OrderNo] [int] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblDocFolder] (
	[FolderID] [int] IDENTITY (1, 1) NOT NULL ,
	[ParentFolderID] [int] NOT NULL ,
	[CreatedByUserID] [int] NULL ,
	[FolderName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ShortDescription] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DocumentCount] [int] NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblDocType] (
	[TypeID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[TypeName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblFaqAuthor] (
	[AuthorID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[UserID] [int] NULL ,
	[Title] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FirstName] [varchar] (24) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MiddleName] [varchar] (24) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LastName] [varchar] (24) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Email] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Biography] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblFaqDocument] (
	[DocumentID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[Title] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Synopsis] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Introduction] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Epilogue] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OrderNo] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[AuthorName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AuthorID] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblFaqQuestion] (
	[QuestionID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[DocumentID] [int] NOT NULL ,
	[Question] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Answer] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OrderNo] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblLang] (
	[LangCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CountryName] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[NativeLanguage] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FlagIcon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Published] [bit] NOT NULL ,
	[UserID] [int] NULL ,
	[PctComplete] [decimal](9, 2) NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblLangText] (
	[TextID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[EnglishText] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblLangTranslation] (
	[TranslationID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[LangCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TextID] [int] NOT NULL ,
	[Translation] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[MemberID] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblLink] (
	[LinkID] [int] IDENTITY (1, 1) NOT NULL ,
	[CategoryID] [int] NOT NULL ,
	[URL] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Label] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblLinkCategory] (
	[CategoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[CategoryName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblMember] (
	[MemberID] [int] IDENTITY (1, 1) NOT NULL ,
	[Username] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Password] [varchar] (64) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Firstname] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Middlename] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Lastname] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address1] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address2] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[City] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StateCode] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ZipCode] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CountryID] [int] NULL ,
	[EmailAddress] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmailAddressAlt] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DayPhone] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EvePhone] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BestCallTime] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ForumIcon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HomePage] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RatingNo] [int] NOT NULL ,
	[AuthCode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblMenuItem] (
	[ItemID] [int] IDENTITY (1, 1) NOT NULL ,
	[MenuID] [int] NOT NULL ,
	[ParentItemID] [int] NOT NULL ,
	[MenuName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[URL] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Content] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OrderNo] [smallint] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblMessage] (
	[MessageID] [int] IDENTITY (1, 1) NOT NULL ,
	[ParentMessageID] [int] NOT NULL ,
	[ThreadID] [int] NOT NULL ,
	[TopicID] [int] NOT NULL ,
	[MemberID] [int] NOT NULL ,
	[Subject] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MessageBody] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ModPoints] [tinyint] NOT NULL ,
	[ModClassID] [int] NULL ,
	[Messages] [int] NOT NULL ,
	[LastPost] [datetime] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblMessageConfig] (
	[ConfigID] [int] IDENTITY (1, 1) NOT NULL ,
	[ThemeName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MessageBoxOutlineColor] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MessageHeadBGColor] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MessageBodyBGColor] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UserInfoBGColor] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HomePageIcon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmailIcon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PrivateMessageIcon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EditIcon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReplyIcon] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ThreadHeadBGColor] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OrderNo] [int] NULL ,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblMessageEmail] (
	[EmailID] [int] IDENTITY (1, 1) NOT NULL ,
	[MessageID] [int] NOT NULL ,
	[FromMemberID] [int] NOT NULL ,
	[ToMemberID] [int] NOT NULL ,
	[Subject] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Body] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblMessagePrivate] (
	[PrivateID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[ThreadID] [int] NOT NULL ,
	[MessageID] [int] NOT NULL ,
	[FromMemberID] [int] NOT NULL ,
	[ToMemberID] [int] NOT NULL ,
	[Body] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReadDate] [datetime] NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblMessageProfile] (
	[ProfileID] [int] IDENTITY (1, 1) NOT NULL ,
	[MemberID] [int] NOT NULL ,
	[RankID] [int] NOT NULL ,
	[Username] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Password] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Location] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Email] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ForumIcon] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[NoPosts] [int] NOT NULL ,
	[NoReplies] [int] NOT NULL ,
	[TotalPosts] [int] NOT NULL ,
	[ShowEmail] [bit] NOT NULL ,
	[ModPoints] [int] NOT NULL ,
	[Biography] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ThemeID] [int] NULL ,
	[LastVisit] [datetime] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[HomePage] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblMessageTopic] (
	[TopicID] [int] IDENTITY (1, 1) NOT NULL ,
	[Title] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ShortComments] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Threads] [int] NOT NULL ,
	[Messages] [int] NOT NULL ,
	[LastPost] [datetime] NULL ,
	[OrderNo] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblModule] (
	[ModuleID] [int] IDENTITY (1, 1) NOT NULL ,
	[CategoryID] [int] NOT NULL ,
	[FolderName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Title] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Synopsis] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AuthorName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[VersionNo] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Size140Module] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SizeFullModule] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UpdateURL] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DoUpdateCheck] [bit] NOT NULL ,
	[CheckDays] [smallint] NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblModuleCategory] (
	[CategoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[CategoryName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FolderName] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ModuleCount] [smallint] NOT NULL ,
	[ActiveModuleCount] [smallint] NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblModuleGroup] (
	[GroupID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[GroupName] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[GroupCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[HasSize140Module] [bit] NOT NULL ,
	[HasSizeFullModule] [bit] NOT NULL ,
	[OrderNo] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblModuleGroupPos] (
	[PosID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[GroupID] [int] NOT NULL ,
	[ModuleID] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[OrderNo] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblModuleParam] (
	[ParamID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[ModuleID] [int] NOT NULL ,
	[ParamName] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ParamValue] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Label] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TypeID] [int] NOT NULL ,
	[MinValue] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MaxValue] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HelpText] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsRequired] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[OrderNo] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblModuleParamOption] (
	[OptionID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[TypeID] [int] NULL ,
	[ParamID] [int] NULL ,
	[OptionValue] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OptionLabel] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[ParentOptionID] [int] NOT NULL ,
	[IsValid] [bit] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblModuleParamType] (
	[TypeID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[TypeCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TypeName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ASPConvertFunction] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HTMLInputType] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[RegExValidate] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LabelPos] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OrderNo] [smallint] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[HasOptions] [bit] NOT NULL ,
	[MinValue] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MaxValue] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsNumeric] [bit] NOT NULL ,
	[QuoteChar] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblModuleResource] (
	[ResourceID] [int] IDENTITY (1, 1) NOT NULL ,
	[CurrentVersion] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[TypeCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PathName] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Content] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblModuleResourceType] (
	[TypeCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TypeName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblPoll] (
	[PollID] [int] IDENTITY (1, 1) NOT NULL ,
	[Question] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblPollAnswer] (
	[AnswerID] [int] IDENTITY (1, 1) NOT NULL ,
	[PollID] [int] NOT NULL ,
	[Answer] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Votes] [int] NOT NULL ,
	[OrderNo] [smallint] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblPollComment] (
	[CommentID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[PollID] [int] NOT NULL ,
	[MemberID] [int] NOT NULL ,
	[Subject] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Body] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ModPoints] [int] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[ParentCommentID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblPollIPAddress] (
	[PollID] [int] NOT NULL ,
	[IPAddress] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Created] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblQuote] (
	[QuoteID] [int] IDENTITY (1, 1) NOT NULL ,
	[Quote] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Author] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblRSSFeed] (
	[FeedID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[FeedName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Title] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FeedURL] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MaxItems] [int] NOT NULL ,
	[ShowDescription] [bit] NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[CacheHours] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblSiteStat] (
	[StatID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[HitCount] [int] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblState] (
	[StateCode] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StateName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblSuggestion] (
	[SuggestionID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[MemberID] [int] NULL ,
	[FromName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FromEmail] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Subject] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Body] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTask] (
	[TaskID] [int] IDENTITY (1, 1) NOT NULL ,
	[SiteID] [int] NULL ,
	[UserID] [int] NOT NULL ,
	[PriorityID] [int] NOT NULL ,
	[StatusID] [int] NOT NULL ,
	[Title] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CommentCount] [int] NOT NULL ,
	[PctComplete] [decimal](9, 2) NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTaskComment] (
	[CommentID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[ParentCommentID] [int] NOT NULL ,
	[TaskID] [int] NOT NULL ,
	[MemberID] [int] NOT NULL ,
	[Subject] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Body] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ModPoints] [smallint] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTaskMessage] (
	[MessageID] [int] IDENTITY (1, 1) NOT NULL ,
	[ParentMessageID] [int] NOT NULL ,
	[TaskID] [int] NOT NULL ,
	[UserID] [int] NOT NULL ,
	[Subject] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Body] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTaskPriority] (
	[PriorityID] [int] IDENTITY (1, 1) NOT NULL ,
	[PriorityName] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ColorCode] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OrderNo] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTaskStatus] (
	[StatusID] [int] IDENTITY (1, 1) NOT NULL ,
	[StatusName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comments] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OrderNo] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTheme] (
	[ThemeID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[ThemeName] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Synopsis] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[WebPath] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AuthorName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CreationDate] [datetime] NULL ,
	[TotalPosRating] [int] NOT NULL ,
	[TotalNegRating] [int] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblUser] (
	[UserID] [int] IDENTITY (1, 1) NOT NULL ,
	[Username] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Password] [varchar] (64) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Firstname] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Middlename] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Lastname] [varchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmailAddress] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmailAddressAlt] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DayPhone] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EvePhone] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BestCallTime] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblUserRight] (
	[RightID] [int] IDENTITY (1, 1) NOT NULL ,
	[ParentRightID] [int] NOT NULL ,
	[RightName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AdminMenuName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Hyperlink] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[HasAdd] [bit] NOT NULL ,
	[HasEdit] [bit] NOT NULL ,
	[HasDelete] [bit] NOT NULL ,
	[HasView] [bit] NOT NULL ,
	[Active] [bit] NOT NULL ,
	[Archive] [bit] NOT NULL ,
	[OrderNo] [int] NOT NULL ,
	[Created] [datetime] NOT NULL ,
	[Modified] [datetime] NOT NULL ,
	[AccessKey] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblUserToRight] (
	[UserID] [int] NOT NULL ,
	[RightID] [int] NOT NULL ,
	[CanAdd] [bit] NOT NULL ,
	[CanEdit] [bit] NOT NULL ,
	[CanDelete] [bit] NOT NULL ,
	[CanView] [bit] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblApplicationVar] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblApplicationVar_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblApplicationVar_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblApplicationVar_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF_tblApplicationVar_IsRequired] DEFAULT (0) FOR [IsRequired],
	CONSTRAINT [DF_tblApplicationVar_HasOptions] DEFAULT (0) FOR [HasOptions],
	CONSTRAINT [DF_tblApplicationVar_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [PK_tblApplicationVar] PRIMARY KEY  CLUSTERED 
	(
		[VarID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblApplicationVarOption] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblApplic__Paren__595B4002] DEFAULT (0) FOR [ParentOptionID],
	CONSTRAINT [DF__tblApplic__IsVal__5A4F643B] DEFAULT (1) FOR [IsValid],
	CONSTRAINT [DF_tblApplicationVarOption_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblApplic__Archi__5B438874] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblApplic__Modif__5C37ACAD] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblApplicationVarOption_OptionID] PRIMARY KEY  NONCLUSTERED 
	(
		[OptionID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblApplicationVarTab] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblApplic__Archi__5F141958] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblApplic__Modif__60083D91] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblApplic__Order__60FC61CA] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [PK_tblApplicationVarTab_TabID] PRIMARY KEY  NONCLUSTERED 
	(
		[TabID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblApplicationVarType] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblApplic__Label__51BA1E3A] DEFAULT ('LEFT') FOR [LabelPos],
	CONSTRAINT [DF__tblApplic__Order__52AE4273] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblApplic__HasOp__53A266AC] DEFAULT (0) FOR [HasOptions],
	CONSTRAINT [DF__tblApplic__IsNum__54968AE5] DEFAULT (0) FOR [IsNumeric],
	CONSTRAINT [DF__tblApplic__Archi__558AAF1E] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblApplic__Modif__567ED357] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblApplicationVarType_TypeID] PRIMARY KEY  NONCLUSTERED 
	(
		[TypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblArticle] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblArticle_AuthorID] DEFAULT (0) FOR [AuthorID],
	CONSTRAINT [DF_tblArticle_WordCount] DEFAULT (0) FOR [WordCount],
	CONSTRAINT [DF_tblArticle_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblArticle_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblArticle_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblArticle_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblArticl__Comme__681373AD] DEFAULT (0) FOR [CommentCount],
	CONSTRAINT [PK_tblArticle] PRIMARY KEY  NONCLUSTERED 
	(
		[ArticleID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblArticleAuthor] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblArticleAuthor_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblArticleAuthor_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblArticleAuthor_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblArticleAuthor_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblArticleAuthor] PRIMARY KEY  NONCLUSTERED 
	(
		[AuthorID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblArticleCategory] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblArticleCategory_ParentCategoryID] DEFAULT (0) FOR [ParentCategoryID],
	CONSTRAINT [DF_tblArticleCategory_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblArticleCategory_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblArticleCategory_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblArticleCategory_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblArticleCategory] PRIMARY KEY  NONCLUSTERED 
	(
		[CategoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblArticleComment] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblArticleComment_ParentCommentID] DEFAULT (0) FOR [ParentCommentID],
	CONSTRAINT [DF_tblArticleComment_Archive] DEFAULT (1) FOR [Archive],
	CONSTRAINT [DF_tblArticleComment_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblArticleComment_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblArticleComment] PRIMARY KEY  CLUSTERED 
	(
		[CommentID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblArticleToCategory] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblArticleToCategory] PRIMARY KEY  NONCLUSTERED 
	(
		[ArticleID],
		[CategoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblContactUs] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblContactUs_MailingList] DEFAULT (0) FOR [MailingList],
	CONSTRAINT [DF_tblContactUs_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblContactUs_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblContactUs_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblContactUs_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblContactUs] PRIMARY KEY  NONCLUSTERED 
	(
		[ContactUsID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblCountry] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblCountry_SortOrder] DEFAULT (0) FOR [SortOrder],
	CONSTRAINT [DF_tblCountry_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblCountry_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblCountry_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblCountry_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblCountry] PRIMARY KEY  NONCLUSTERED 
	(
		[CountryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblDoc] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblDoc_FolderID] DEFAULT (0) FOR [FolderID],
	CONSTRAINT [DF_tblDoc_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblDoc_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblDoc_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblDoc_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblDoc__ParentDo__7ABC33CD] DEFAULT (0) FOR [ParentDocID],
	CONSTRAINT [DF__tblDoc__IsInline__7BB05806] DEFAULT (0) FOR [IsInlineContent],
	CONSTRAINT [DF__tblDoc__OrderNo__7CA47C3F] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblDoc__SectionN__7D98A078] DEFAULT (0) FOR [SectionNo],
	CONSTRAINT [PK_tblDoc] PRIMARY KEY  CLUSTERED 
	(
		[DocID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblDocAuthor] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblDocAuthor_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblDocAuthor_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblDocAuthor_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblDocAuthor_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblDocAuthor] PRIMARY KEY  CLUSTERED 
	(
		[AuthorID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblDocBook] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblDocBoo__ShowS__00750D23] DEFAULT (1) FOR [ShowSectionNo],
	CONSTRAINT [DF__tblDocBoo__Order__0169315C] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblDocBoo__Archi__025D5595] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblDocBoo__Creat__035179CE] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblDocBook_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblDocBook_BookID] PRIMARY KEY  NONCLUSTERED 
	(
		[BookID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblDocFolder] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblDocFolder_ParentFolderID] DEFAULT (0) FOR [ParentFolderID],
	CONSTRAINT [DF_tblDocFolder_DocumentCount] DEFAULT (0) FOR [DocumentCount],
	CONSTRAINT [DF_tblDocFolder_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblDocFolder_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblDocFolder_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblDocFolder_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblDocFolder_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblDocFolder] PRIMARY KEY  CLUSTERED 
	(
		[FolderID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblDocType] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblDocTyp__Archi__0AF29B96] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblDocTyp__Creat__0BE6BFCF] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblDocType_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblDocType_TypeID] PRIMARY KEY  NONCLUSTERED 
	(
		[TypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblFaqAuthor] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblFaqAut__Archi__062DE679] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblFaqAut__Creat__07220AB2] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblFaqAut__Modif__08162EEB] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblFaqAuthor_AuthorID] PRIMARY KEY  NONCLUSTERED 
	(
		[AuthorID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblFaqDocument] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblFaqDoc__Order__56E8E7AB] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblFaqDoc__Activ__57DD0BE4] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF__tblFaqDoc__Archi__58D1301D] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblFaqDoc__Creat__59C55456] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblFaqDoc__Modif__5AB9788F] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblFaqDocument_DocumentID] PRIMARY KEY  NONCLUSTERED 
	(
		[DocumentID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblFaqQuestion] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblFaqQue__Order__5D95E53A] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblFaqQue__Activ__5E8A0973] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF__tblFaqQue__Archi__5F7E2DAC] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblFaqQue__Creat__607251E5] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblFaqQue__Modif__6166761E] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblFaqQuestion_QuestionID] PRIMARY KEY  NONCLUSTERED 
	(
		[QuestionID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblLang] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblLang__Publish__0EC32C7A] DEFAULT (0) FOR [Published],
	CONSTRAINT [DF__tblLang__OrderNo__0FB750B3] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblLang__Archive__10AB74EC] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblLang__Created__119F9925] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblLang__Modifie__1293BD5E] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblLang_LangCode] PRIMARY KEY  NONCLUSTERED 
	(
		[LangCode]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblLangText] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblLangTe__Archi__15702A09] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblLangTe__Creat__16644E42] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblLangTe__Modif__1758727B] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblLangText_TextID] PRIMARY KEY  NONCLUSTERED 
	(
		[TextID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblLangTranslation] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblLangTr__Archi__1A34DF26] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblLangTr__Creat__1B29035F] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblLangTr__Modif__1C1D2798] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblLangTr__Membe__23BE4960] DEFAULT (0) FOR [MemberID],
	CONSTRAINT [PK_tblLangTranslation_TranslationID] PRIMARY KEY  NONCLUSTERED 
	(
		[TranslationID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblLink] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblLink_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblLink_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblLink_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblLink_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblLink_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblLink] PRIMARY KEY  CLUSTERED 
	(
		[LinkID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblLinkCategory] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblLinkCategory_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblLinkCategory_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblLinkCategory_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblLinkCategory_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblLinkCategory_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblLinkCategory] PRIMARY KEY  CLUSTERED 
	(
		[CategoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMember] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblMember_RatingNo] DEFAULT (0) FOR [RatingNo],
	CONSTRAINT [DF_tblMember_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblMember_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblMember_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblMember_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblMember] PRIMARY KEY  NONCLUSTERED 
	(
		[MemberID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMenuItem] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblMenuItem_MenuID] DEFAULT (0) FOR [MenuID],
	CONSTRAINT [DF_tblMenuItem_ParentMenuID] DEFAULT (0) FOR [ParentItemID],
	CONSTRAINT [DF_tblMenuItem_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblMenuItem_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblMenuItem_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblMenuItem_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblMenuItem_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblMenuItem] PRIMARY KEY  CLUSTERED 
	(
		[ItemID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMessage] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblMessage_ParentMessageID] DEFAULT (0) FOR [ParentMessageID],
	CONSTRAINT [DF_tblMessage_ModPoints] DEFAULT (0) FOR [ModPoints],
	CONSTRAINT [DF_tblMessage_Messages] DEFAULT (0) FOR [Messages],
	CONSTRAINT [DF_tblMessage_LastPost] DEFAULT (getdate()) FOR [LastPost],
	CONSTRAINT [DF_tblMessage_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblMessage_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblMessage_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblMessage_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblMessage] PRIMARY KEY  CLUSTERED 
	(
		[MessageID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMessageConfig] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblMessageConfig_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblMessageConfig_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblMessageConfig_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblMessageConfig_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblMessageConfig] PRIMARY KEY  CLUSTERED 
	(
		[ConfigID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMessageEmail] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblMessageEmail_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblMessageEmail_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblMessageEmail_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblMessageEmail_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblMessageEmail] PRIMARY KEY  CLUSTERED 
	(
		[EmailID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMessagePrivate] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblMessag__Archi__54CB950F] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblMessag__Creat__55BFB948] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblMessag__Modif__56B3DD81] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblMessagePrivate_PrivateID] PRIMARY KEY  NONCLUSTERED 
	(
		[PrivateID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMessageProfile] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblMessageProfile_NoPosts] DEFAULT (0) FOR [NoPosts],
	CONSTRAINT [DF_tblMessageProfile_NoReplies] DEFAULT (0) FOR [NoReplies],
	CONSTRAINT [DF_tblMessageProfile_TotalPosts] DEFAULT (0) FOR [TotalPosts],
	CONSTRAINT [DF_tblMessageProfile_ShowEmail] DEFAULT (0) FOR [ShowEmail],
	CONSTRAINT [DF_tblMessageProfile_ModPoints] DEFAULT (0) FOR [ModPoints],
	CONSTRAINT [DF_tblMessageProfile_LastVisit] DEFAULT (getdate()) FOR [LastVisit],
	CONSTRAINT [DF_tblMessageProfile_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblMessageProfile_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblMessageProfile_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblMessageProfile_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblMessageProfile] PRIMARY KEY  CLUSTERED 
	(
		[ProfileID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblMessageTopic] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblMessageTopic_Threads] DEFAULT (0) FOR [Threads],
	CONSTRAINT [DF_tblMessageTopic_Messages] DEFAULT (0) FOR [Messages],
	CONSTRAINT [DF_tblMessageTopic_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblMessageTopic_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblMessageTopic_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblMessageTopic_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblMessageTopic_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblMessageTopic] PRIMARY KEY  CLUSTERED 
	(
		[TopicID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblModule] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblModule_DoUpdateCheck] DEFAULT (0) FOR [DoUpdateCheck],
	CONSTRAINT [DF_tblModule_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblModule_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblModule_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblModule] PRIMARY KEY  CLUSTERED 
	(
		[ModuleID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblModuleCategory] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblModuleCategory_ModuleCount] DEFAULT (0) FOR [ModuleCount],
	CONSTRAINT [DF_tblModuleCategory_ActiveModuleCount] DEFAULT (0) FOR [ActiveModuleCount],
	CONSTRAINT [DF_tblModuleCategory_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblModuleCategory_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblModuleCategory_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblModuleCategory_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblModuleCategory] PRIMARY KEY  CLUSTERED 
	(
		[CategoryID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblModuleGroup] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblModule__Activ__1975C517] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF__tblModule__Archi__1A69E950] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblModule__Creat__1B5E0D89] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblModule__Modif__1C5231C2] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblModule__HasSi__23F3538A] DEFAULT (0) FOR [HasSize140Module],
	CONSTRAINT [DF__tblModule__HasSi__24E777C3] DEFAULT (0) FOR [HasSizeFullModule],
	CONSTRAINT [DF__tblModule__Order__25DB9BFC] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [PK_tblModuleGroup_GroupID] PRIMARY KEY  NONCLUSTERED 
	(
		[GroupID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblModuleGroupPos] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblModule__Activ__1F2E9E6D] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF__tblModule__Archi__2022C2A6] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblModule__Creat__2116E6DF] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblModule__Modif__220B0B18] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblModule__Order__22FF2F51] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [PK_tblModuleGroupPos_PosID] PRIMARY KEY  NONCLUSTERED 
	(
		[PosID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblModuleParam] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblModule__IsReq__28B808A7] DEFAULT (0) FOR [IsRequired],
	CONSTRAINT [DF__tblModule__Archi__29AC2CE0] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblModule__Modif__2AA05119] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblModule__Order__2B947552] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [PK_tblModuleParam_ParamID] PRIMARY KEY  NONCLUSTERED 
	(
		[ParamID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblModuleParamOption] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblModule__Order__351DDF8C] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblModule__Archi__361203C5] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblModule__Modif__370627FE] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblModule__Paren__37FA4C37] DEFAULT (0) FOR [ParentOptionID],
	CONSTRAINT [DF__tblModule__IsVal__38EE7070] DEFAULT (1) FOR [IsValid],
	CONSTRAINT [PK_tblModuleParamOption_OptionID] PRIMARY KEY  NONCLUSTERED 
	(
		[OptionID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblModuleParamType] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblModule__Label__2E70E1FD] DEFAULT ('LEFT') FOR [LabelPos],
	CONSTRAINT [DF__tblModule__Order__2F650636] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblModule__Archi__30592A6F] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblModule__Modif__314D4EA8] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblModule__HasOp__324172E1] DEFAULT (0) FOR [HasOptions],
	CONSTRAINT [DF__tblModule__IsNum__46486B8E] DEFAULT (0) FOR [IsNumeric],
	CONSTRAINT [PK_tblModuleParamType_TypeID] PRIMARY KEY  NONCLUSTERED 
	(
		[TypeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblModuleResource] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblModuleResource_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblModuleResource_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblModuleResource_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblModuleResource] PRIMARY KEY  CLUSTERED 
	(
		[ResourceID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblModuleResourceType] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblModuleResourceType_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblModuleResourceType_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblModuleResourceType_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblModuleResourceType] PRIMARY KEY  CLUSTERED 
	(
		[TypeCode]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblPoll] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblPoll_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblPoll_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblPoll_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblPoll_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblPoll] PRIMARY KEY  CLUSTERED 
	(
		[PollID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblPollAnswer] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblPollAnswer_Votes] DEFAULT (0) FOR [Votes],
	CONSTRAINT [DF_tblPollAnswer_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblPollAnswer_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblPollAnswer_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblPollAnswer_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblPollAnswer_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblPollAnswer] PRIMARY KEY  CLUSTERED 
	(
		[AnswerID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblPollComment] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblPollCo__ModPo__4589517F] DEFAULT (0) FOR [ModPoints],
	CONSTRAINT [DF__tblPollCo__Archi__467D75B8] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblPollCo__Creat__477199F1] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblPollCo__Modif__4865BE2A] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblPollCo__Paren__4959E263] DEFAULT (0) FOR [ParentCommentID],
	CONSTRAINT [PK_tblPollComment_CommentID] PRIMARY KEY  NONCLUSTERED 
	(
		[CommentID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblPollIPAddress] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblPollIPAddress_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [PK_tblPollIPAddress] PRIMARY KEY  CLUSTERED 
	(
		[PollID],
		[IPAddress]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblQuote] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblQuote_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblQuote_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblQuote_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblQuote_LastModified] DEFAULT (getdate()) FOR [LastModified],
	CONSTRAINT [PK_tblQuote] PRIMARY KEY  CLUSTERED 
	(
		[QuoteID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblRSSFeed] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblRSSFee__MaxIt__75035A77] DEFAULT (0) FOR [MaxItems],
	CONSTRAINT [DF__tblRSSFee__ShowD__75F77EB0] DEFAULT (1) FOR [ShowDescription],
	CONSTRAINT [DF__tblRSSFee__Order__76EBA2E9] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF__tblRSSFee__Archi__77DFC722] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblRSSFee__Creat__78D3EB5B] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [DF__tblRSSFee__Cache__79C80F94] DEFAULT (2) FOR [CacheHours],
	CONSTRAINT [PK_tblRSSFeed_FeedID] PRIMARY KEY  NONCLUSTERED 
	(
		[FeedID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblSiteStat] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblSiteSt__HitCo__6AEFE058] DEFAULT (0) FOR [HitCount],
	CONSTRAINT [DF__tblSiteSt__Creat__6BE40491] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblSiteSt__Modif__6CD828CA] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblSiteStat_StatID] PRIMARY KEY  NONCLUSTERED 
	(
		[StatID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblState] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblState_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblState_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblState_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblState_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblState] PRIMARY KEY  CLUSTERED 
	(
		[StateCode]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblSuggestion] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblSugges__Activ__51300E55] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF__tblSugges__Archi__5224328E] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblSugges__Creat__531856C7] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblSugges__Modif__540C7B00] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblSuggestion_SuggestionID] PRIMARY KEY  NONCLUSTERED 
	(
		[SuggestionID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTask] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblTask_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblTask_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblTask_CommentCount] DEFAULT (0) FOR [CommentCount],
	CONSTRAINT [DF_tblTask_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblTask_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblTask] PRIMARY KEY  NONCLUSTERED 
	(
		[TaskID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTaskComment] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblTaskCo__Paren__1EF99443] DEFAULT (0) FOR [ParentCommentID],
	CONSTRAINT [DF__tblTaskCo__ModPo__1FEDB87C] DEFAULT (0) FOR [ModPoints],
	CONSTRAINT [DF__tblTaskCo__Archi__20E1DCB5] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblTaskCo__Creat__21D600EE] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblTaskCo__Modif__22CA2527] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblTaskComment_CommentID] PRIMARY KEY  NONCLUSTERED 
	(
		[CommentID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTaskMessage] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblTaskMessage_ParentMessageID] DEFAULT (0) FOR [ParentMessageID],
	CONSTRAINT [DF_tblTaskMessage_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblTaskMessage_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblTaskMessage_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblTaskMessage_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblTaskMessage] PRIMARY KEY  NONCLUSTERED 
	(
		[MessageID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTaskPriority] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblTaskPriority_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblTaskPriority_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblTaskPriority_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblTaskPriority_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblTaskPriority_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblTaskPriority] PRIMARY KEY  NONCLUSTERED 
	(
		[PriorityID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTaskStatus] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblTaskStatus_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblTaskStatus_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblTaskStatus_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblTaskStatus_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblTaskStatus_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblTaskStatus] PRIMARY KEY  NONCLUSTERED 
	(
		[StatusID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTheme] WITH NOCHECK ADD 
	CONSTRAINT [DF__tblTheme__TotalP__76619304] DEFAULT (0) FOR [TotalPosRating],
	CONSTRAINT [DF__tblTheme__TotalN__7755B73D] DEFAULT (0) FOR [TotalNegRating],
	CONSTRAINT [DF__tblTheme__Active__7849DB76] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF__tblTheme__Archiv__793DFFAF] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF__tblTheme__Create__7A3223E8] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF__tblTheme__Modifi__7B264821] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblTheme_ThemeID] PRIMARY KEY  NONCLUSTERED 
	(
		[ThemeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblUser] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblUser_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblUser_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblUser_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblUser_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblUser] PRIMARY KEY  NONCLUSTERED 
	(
		[UserID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblUserRight] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblUserRight_ParentRightID] DEFAULT (0) FOR [ParentRightID],
	CONSTRAINT [DF_tblUserRight_HasAdd] DEFAULT (0) FOR [HasAdd],
	CONSTRAINT [DF_tblUserRight_HasEdit] DEFAULT (0) FOR [HasEdit],
	CONSTRAINT [DF_tblUserRight_HasDelete] DEFAULT (0) FOR [HasDelete],
	CONSTRAINT [DF_tblUserRight_HasView] DEFAULT (0) FOR [HasView],
	CONSTRAINT [DF_tblUserRight_Active] DEFAULT (1) FOR [Active],
	CONSTRAINT [DF_tblUserRight_Archive] DEFAULT (0) FOR [Archive],
	CONSTRAINT [DF_tblUserRight_OrderNo] DEFAULT (0) FOR [OrderNo],
	CONSTRAINT [DF_tblUserRight_Created] DEFAULT (getdate()) FOR [Created],
	CONSTRAINT [DF_tblUserRight_Modified] DEFAULT (getdate()) FOR [Modified],
	CONSTRAINT [PK_tblUserRight] PRIMARY KEY  NONCLUSTERED 
	(
		[RightID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblUserToRight] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblUserToRight_CanAdd] DEFAULT (1) FOR [CanAdd],
	CONSTRAINT [DF_tblUserToRight_CanEdit] DEFAULT (1) FOR [CanEdit],
	CONSTRAINT [DF_tblUserToRight_CanDelete] DEFAULT (1) FOR [CanDelete],
	CONSTRAINT [DF_tblUserToRight_CanView] DEFAULT (1) FOR [CanView],
	CONSTRAINT [PK_tblUserToRight] PRIMARY KEY  NONCLUSTERED 
	(
		[UserID],
		[RightID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

