DROP TABLE IF EXISTS tblApplicationVar;

DROP TABLE IF EXISTS tblApplicationVarOption;

DROP TABLE IF EXISTS tblApplicationVarTab;

DROP TABLE IF EXISTS tblApplicationVarType;

DROP TABLE IF EXISTS tblArticle;

DROP TABLE IF EXISTS tblArticleAuthor;

DROP TABLE IF EXISTS tblArticleCategory;

DROP TABLE IF EXISTS tblArticleComment;

DROP TABLE IF EXISTS tblArticleToCategory;

DROP TABLE IF EXISTS tblContactUs;

DROP TABLE IF EXISTS tblCountry;

DROP TABLE IF EXISTS tblDoc;

DROP TABLE IF EXISTS tblDocAuthor;

DROP TABLE IF EXISTS tblDocBook;

DROP TABLE IF EXISTS tblDocFolder;

DROP TABLE IF EXISTS tblDocType;

DROP TABLE IF EXISTS tblFaqAuthor;

DROP TABLE IF EXISTS tblFaqDocument;

DROP TABLE IF EXISTS tblFaqQuestion;

DROP TABLE IF EXISTS tblLang;

DROP TABLE IF EXISTS tblLangText;

DROP TABLE IF EXISTS tblLangTranslation;

DROP TABLE IF EXISTS tblLink;

DROP TABLE IF EXISTS tblLinkCategory;

DROP TABLE IF EXISTS tblMember;

DROP TABLE IF EXISTS tblMenuItem;

DROP TABLE IF EXISTS tblMessage;

DROP TABLE IF EXISTS tblMessageConfig;

DROP TABLE IF EXISTS tblMessageEmail;

DROP TABLE IF EXISTS tblMessagePrivate;

DROP TABLE IF EXISTS tblMessageProfile;

DROP TABLE IF EXISTS tblMessageTopic;

DROP TABLE IF EXISTS tblModule;

DROP TABLE IF EXISTS tblModuleCategory;

DROP TABLE IF EXISTS tblModuleGroup;

DROP TABLE IF EXISTS tblModuleGroupPos;

DROP TABLE IF EXISTS tblModuleParam;

DROP TABLE IF EXISTS tblModuleParamOption;

DROP TABLE IF EXISTS tblModuleParamType;

DROP TABLE IF EXISTS tblModuleResource;

DROP TABLE IF EXISTS tblModuleResourceType;

DROP TABLE IF EXISTS tblPoll;

DROP TABLE IF EXISTS tblPollAnswer;

DROP TABLE IF EXISTS tblPollComment;

DROP TABLE IF EXISTS tblPollIPAddress;

DROP TABLE IF EXISTS tblQuote;

DROP TABLE IF EXISTS tblSiteStat;

DROP TABLE IF EXISTS tblState;

DROP TABLE IF EXISTS tblSuggestion;

DROP TABLE IF EXISTS tblTask;

DROP TABLE IF EXISTS tblTaskComment;

DROP TABLE IF EXISTS tblTaskMessage;

DROP TABLE IF EXISTS tblTaskPriority;

DROP TABLE IF EXISTS tblTaskStatus;

DROP TABLE IF EXISTS tblTheme;

DROP TABLE IF EXISTS tblUser;

DROP TABLE IF EXISTS tblUserRight;

DROP TABLE IF EXISTS tblUserToRight;

CREATE TABLE tblApplicationVar (
	VarID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	VarName varchar (100) NOT NULL ,
	VarValue text NOT NULL ,
	Label varchar (100) NULL ,
	HelpText text NULL ,
	TabID int NULL ,
	IsRequired tinyint NOT NULL default 0,
	TypeID int NULL ,
	HasOptions tinyint NOT NULL default 0,
	MinValue varchar (32) NULL ,
	MaxValue varchar (32) NULL ,
	OrderNo smallint NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL
);

CREATE TABLE tblApplicationVarOption (
	OptionID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	TypeID int NULL ,
	VarID int NULL ,
	OptionValue varchar (255) NOT NULL ,
	OptionLabel varchar (255) NOT NULL ,
	ParentOptionID int NOT NULL default 0,
	IsValid tinyint NOT NULL default 1,
	OrderNo int NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL 
);


CREATE TABLE tblApplicationVarTab (
	TabID int AUTO_INCREMENT   NOT NULL PRIMARY KEY,
	TabName varchar (32) NOT NULL ,
	Title varchar (100) NULL ,
	Introduction text NULL ,
	Summary text NULL ,
	OrderNo smallint NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL
);


CREATE TABLE tblApplicationVarType (
	TypeID int AUTO_INCREMENT   NOT NULL PRIMARY KEY,
	TypeCode varchar (4) NOT NULL ,
	TypeName varchar (100) NOT NULL ,
	ASPConvertFunction varchar (100) NULL ,
	HTMLInputType varchar (100) NULL ,
	RegExValidate varchar (255) NULL ,
	LabelPos varchar (32) NOT NULL default 'LEFT',
	OrderNo smallint NOT NULL default 0,
	HasOptions tinyint NOT NULL default 0,
	MinValue varchar (32) NULL ,
	MaxValue varchar (32) NULL ,
	IsNumeric tinyint NOT NULL default 0,
	QuoteChar varchar (1) NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL 
);


CREATE TABLE tblArticle (
	ArticleID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	MagazineID int NULL ,
	AuthorID int NOT NULL default 0,
	Title varchar (255) NOT NULL ,
	LeadIn text NULL ,
	ArticleBody text NULL ,
	ShortComments varchar (255) NULL ,
	Comments text NULL ,
	WordCount int NOT NULL default 0,
	CommentCount int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblArticleAuthor (
	AuthorID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	Title varchar (32) NULL ,
	Firstname varchar (32) NOT NULL ,
	Middlename varchar (32) NULL ,
	Lastname varchar (32) NOT NULL ,
	Surname varchar (32) NULL ,
	Comments text NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL
);


CREATE TABLE tblArticleCategory (
	CategoryID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ParentCategoryID int NOT NULL default 0,
	CategoryName varchar (40) NOT NULL ,
	Comments text NULL ,
	IconImage varchar (100) NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblArticleComment (
	CommentID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ArticleID int NOT NULL ,
	MemberID int NULL ,
	ParentCommentID int NOT NULL default 0,
	Subject varchar (100) NOT NULL ,
	Body text NULL ,
	ModPoints smallint NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblArticleToCategory (
	ArticleID int NOT NULL ,
	CategoryID int NOT NULL ,
	PRIMARY KEY (ArticleID, CategoryID)
);


CREATE TABLE tblContactUs (
	ContactUsID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	FirstName varchar (32) NULL ,
	LastName varchar (32) NULL ,
	Address1 varchar (40) NULL ,
	Address2 varchar (40) NULL ,
	City varchar (32) NULL ,
	StateCode varchar (2) NULL ,
	ZipCode varchar (10) NULL ,
	CountryID int NULL ,
	Email varchar (80) NULL ,
	Comments text NULL ,
	MailingList tinyint NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblCountry (
	CountryID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	CountryName varchar (50) NOT NULL ,
	SortOrder int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblDoc (
	DocID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ParentDocID int NOT NULL default 0,
	BookID int NULL ,
	TypeID int NULL ,
	AuthorID int NOT NULL ,
	FolderID int NOT NULL default 0,
	Title varchar (100) NOT NULL ,
	SubTitle varchar (255) NULL ,
	ShortDescription text NULL ,
	SectionName varchar (100) NULL ,
	IsInlineContent tinyint NOT NULL default 0,
	Body text NULL ,
	OrderNo int NOT NULL default 0,
	SectionNo varchar (32) NOT NULL default 0,
	AuthorNotes text NULL ,
	ScriptName varchar (100) NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblDocAuthor (
	AuthorID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	UserID int NULL ,
	Title varchar (10) NULL ,
	FirstName varchar (24) NOT NULL ,
	MiddleName varchar (24) NULL ,
	LastName varchar (24) NOT NULL ,
	Surname varchar (20) NULL ,
	EmailAddress varchar (100) NULL ,
	Biography text NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblDocBook (
	BookID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	FolderID int NOT NULL default 0,
	Title varchar (100) NOT NULL ,
	SubTitle varchar (255) NULL ,
	AuthorID int NOT NULL ,
	Version varchar (100) NULL ,
	PublishDate datetime NULL ,
	ShowSectionNo tinyint NOT NULL default 1,
	AuthorNotes text NULL ,
	OrderNo int NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblDocFolder (
	FolderID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ParentFolderID int NOT NULL default 0,
	CreatedByUserID int NULL ,
	FolderName varchar (50) NOT NULL ,
	ShortDescription text NULL ,
	DocumentCount int NOT NULL default 0,
	OrderNo int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblDocType (
	TypeID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	TypeName varchar (100) NOT NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblFaqAuthor (
	AuthorID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	UserID int NULL ,
	Title varchar (10) NULL ,
	FirstName varchar (24) NOT NULL ,
	MiddleName varchar (24) NULL ,
	LastName varchar (24) NOT NULL ,
	Email varchar (100) NULL ,
	Biography text NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblFaqDocument (
	DocumentID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	AuthorID int NULL ,
	AuthorName varchar (50) NULL ,
	Title varchar (100) NOT NULL ,
	Synopsis text NULL ,
	Introduction text NULL ,
	Epilogue text NULL ,
	OrderNo int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblFaqQuestion (
	QuestionID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	DocumentID int NOT NULL ,
	Question varchar (255) NOT NULL ,
	Answer text NULL ,
	OrderNo int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblLang (
	LangCode varchar (4) NOT NULL PRIMARY KEY,
	CountryName varchar (32) NOT NULL ,
	NativeLanguage varchar (50) NOT NULL ,
	FlagIcon varchar (50) NULL ,
	Published tinyint NOT NULL default 0,
	UserID int NULL ,
	PctComplete decimal(9, 2) NULL ,
	OrderNo int NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblLangText (
	TextID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	EnglishText varchar (255) NOT NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblLangTranslation (
	TranslationID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	MemberID int NOT NULL default 0,
	LangCode varchar (4) NOT NULL ,
	TextID int NOT NULL ,
	Translation varchar (255) NOT NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblLink (
	LinkID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	CategoryID int NOT NULL ,
	URL varchar (100) NOT NULL ,
	Label varchar (100) NOT NULL ,
	OrderNo int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblLinkCategory (
	CategoryID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	CategoryName varchar (50) NOT NULL ,
	OrderNo int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblMember (
	MemberID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	Username varchar (32) NULL ,
	Password varchar (64) NULL ,
	Firstname varchar (32) NULL ,
	Middlename varchar (32) NULL ,
	Lastname varchar (32) NULL ,
	Address1 varchar (40) NULL ,
	Address2 varchar (40) NULL ,
	City varchar (32) NULL ,
	StateCode varchar (2) NULL ,
	ZipCode varchar (10) NULL ,
	CountryID int NULL ,
	EmailAddress varchar (100) NULL ,
	EmailAddressAlt varchar (100) NULL ,
	DayPhone varchar (16) NULL ,
	EvePhone varchar (16) NULL ,
	BestCallTime char (1) NULL ,
	ForumIcon varchar (50) NULL ,
	HomePage varchar (100) NULL ,
	RatingNo int NOT NULL default 0,
	AuthCode varchar (20) NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblMenuItem (
	ItemID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	MenuID int NOT NULL default 0,
	ParentItemID int NOT NULL default 0,
	MenuName varchar (50) NOT NULL ,
	URL varchar (255) NULL ,
	Content text NULL ,
	OrderNo smallint NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblMessage (
	MessageID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ParentMessageID int NOT NULL default 0,
	ThreadID int NOT NULL ,
	TopicID int NOT NULL ,
	MemberID int NOT NULL ,
	Subject varchar (80) NOT NULL ,
	MessageBody text NULL ,
	ModPoints tinyint NOT NULL default 0,
	ModClassID int NULL ,
	Messages int NOT NULL default 0,
	LastPost datetime NOT NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblMessageConfig (
	ConfigID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ThemeName varchar (50) NULL ,
	MessageBoxOutlineColor varchar (16) NULL ,
	MessageHeadBGColor varchar (16) NULL ,
	MessageBodyBGColor varchar (16) NULL ,
	UserInfoBGColor varchar (16) NULL ,
	HomePageIcon varchar (50) NULL ,
	EmailIcon varchar (50) NULL ,
	PrivateMessageIcon varchar (50) NULL ,
	EditIcon varchar (50) NULL ,
	ReplyIcon varchar (50) NULL ,
	ThreadHeadBGColor varchar (16) NULL ,
	OrderNo int NULL ,
	Comments text NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblMessageEmail (
	EmailID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	MessageID int NOT NULL ,
	FromMemberID int NOT NULL ,
	ToMemberID int NOT NULL ,
	Subject varchar (100) NOT NULL ,
	Body varchar (50) NOT NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblMessagePrivate (
	PrivateID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ThreadID int NOT NULL ,
	MessageID int NOT NULL ,
	FromMemberID int NOT NULL ,
	ToMemberID int NOT NULL ,
	Body text NULL ,
	ReadDate datetime NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblMessageProfile (
	ProfileID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	MemberID int NOT NULL ,
	RankID int NOT NULL ,
	ThemeID int NULL ,
	Username varchar (20) NULL ,
	Password varchar (32) NULL ,
	Location varchar (50) NULL ,
	Email varchar (100) NULL ,
	ForumIcon varchar (255) NULL ,
	NoPosts int NOT NULL default 0,
	NoReplies int NOT NULL default 0,
	TotalPosts int NOT NULL default 0,
	ShowEmail tinyint NOT NULL default 0,
	ModPoints int NOT NULL default 0,
	Biography text NULL ,
	HomePage varchar (200) NULL ,
	LastVisit datetime NOT NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblMessageTopic (
	TopicID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	Title varchar (100) NOT NULL ,
	ShortComments text NULL ,
	Threads int NOT NULL default 0,
	Messages int NOT NULL default 0,
	LastPost datetime NULL ,
	OrderNo int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblModule (
	ModuleID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	CategoryID int NOT NULL ,
	FolderName varchar (20) NOT NULL ,
	Title varchar (100) NOT NULL ,
	Synopsis text NULL ,
	Description text NULL ,
	AuthorName varchar (50) NULL ,
	VersionNo varchar (16) NULL ,
	Size140Module varchar (200) NULL ,
	SizeFullModule varchar (200) NULL ,
	UpdateURL varchar (255) NULL ,
	DoUpdateCheck tinyint NOT NULL default 0,
	CheckDays smallint NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblModuleCategory (
	CategoryID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	CategoryName varchar (100) NOT NULL ,
	FolderName varchar (20) NOT NULL ,
	Description text NULL ,
	ModuleCount smallint NOT NULL default 0,
	ActiveModuleCount smallint NOT NULL default 0,
	OrderNo int NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblModuleGroup (
	GroupID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	GroupName varchar (32) NOT NULL ,
	GroupCode varchar (4) NOT NULL ,
	HasSize140Module tinyint NOT NULL default 0,
	HasSizeFullModule tinyint NOT NULL default 0,
	OrderNo smallint NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblModuleGroupPos (
	PosID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	GroupID int NOT NULL ,
	ModuleID int NOT NULL ,
	OrderNo smallint NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblModuleParam (
	ParamID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ModuleID int NOT NULL ,
	ParamName varchar (32) NOT NULL ,
	ParamValue varchar (255) NOT NULL ,
	Label varchar (100) NOT NULL ,
	TypeID int NOT NULL ,
	MinValue varchar (255) NULL ,
	MaxValue varchar (255) NULL ,
	HelpText text NULL ,
	IsRequired tinyint NOT NULL default 0,
	OrderNo int NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL 
);


CREATE TABLE tblModuleParamOption (
	OptionID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ParentOptionID int NOT NULL default 0,
	TypeID int NULL ,
	ParamID int NULL ,
	OptionValue varchar (255) NOT NULL ,
	OptionLabel varchar (255) NOT NULL ,
	IsValid tinyint NOT NULL default 1,
	OrderNo int NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL 
);


CREATE TABLE tblModuleParamType (
	TypeID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	TypeCode varchar (4) NOT NULL ,
	TypeName varchar (100) NOT NULL ,
	ASPConvertFunction varchar (100) NULL ,
	HTMLInputType varchar (100) NULL ,
	RegExValidate varchar (255) NULL ,
	LabelPos varchar (32) NOT NULL default 'LEFT',
	OrderNo smallint NOT NULL default 0,
	HasOptions tinyint NOT NULL default 0,
	MinValue varchar (32) NULL ,
	MaxValue varchar (32) NULL ,
	IsNumeric tinyint NOT NULL default 0,
	QuoteChar varchar (1) NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL 
);


CREATE TABLE tblModuleResource (
	ResourceID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	CurrentVersion varchar (16) NULL ,
	TypeCode varchar (4) NOT NULL ,
	PathName varchar (200) NULL ,
	Content text NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblModuleResourceType (
	TypeCode varchar (4) NOT NULL PRIMARY KEY,
	TypeName varchar (50) NOT NULL ,
	Description text NULL ,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblPoll (
	PollID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	Question varchar (255) NOT NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblPollAnswer (
	AnswerID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	PollID int NOT NULL ,
	Answer varchar (50) NOT NULL ,
	Votes int NOT NULL default 0,
	OrderNo smallint NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblPollComment (
	CommentID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ParentCommentID int NOT NULL default 0,
	PollID int NOT NULL ,
	MemberID int NOT NULL ,
	Subject varchar (100) NOT NULL ,
	Body text NULL ,
	ModPoints int NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Created timestamp NOT NULL 
);


CREATE TABLE tblPollIPAddress (
	PollID int NOT NULL ,
	IPAddress varchar (100) NOT NULL ,
	Created timestamp NOT NULL ,
	PRIMARY KEY (PollID, IPAddress)
);


CREATE TABLE tblQuote (
	QuoteID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	Quote text NOT NULL ,
	Author varchar (50) NOT NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblSiteStat (
	StatID int AUTO_INCREMENT   NOT NULL PRIMARY KEY,
	HitCount int NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblState (
	StateCode varchar (2) NOT NULL PRIMARY KEY,
	StateName varchar (50) NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL, 
	Created timestamp NOT NULL 
);


CREATE TABLE tblSuggestion (
	SuggestionID int AUTO_INCREMENT   NOT NULL PRIMARY KEY,
	MemberID int NULL ,
	FromName varchar (50) NULL ,
	FromEmail varchar (100) NULL ,
	Subject varchar (100) NOT NULL ,
	Body text NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblTask (
	TaskID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	SiteID int NULL ,
	UserID int NOT NULL ,
	PriorityID int NOT NULL ,
	StatusID int NOT NULL ,
	Title varchar (50) NOT NULL ,
	Comments text NULL ,
	CommentCount int NOT NULL default 0,
	PctComplete decimal(9, 2) NOT NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblTaskComment (
	CommentID int AUTO_INCREMENT   NOT NULL PRIMARY KEY,
	ParentCommentID int NOT NULL default 0,
	TaskID int NOT NULL ,
	MemberID int NOT NULL ,
	Subject varchar (100) NOT NULL ,
	Body text NULL ,
	ModPoints smallint NOT NULL default 0,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblTaskMessage (
	MessageID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ParentMessageID int NOT NULL default 0,
	TaskID int NOT NULL ,
	UserID int NOT NULL ,
	Subject varchar (80) NOT NULL ,
	Body text NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblTaskPriority (
	PriorityID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	PriorityName varchar (32) NULL ,
	Comments text NULL ,
	ColorCode varchar (16) NULL ,
	OrderNo int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblTaskStatus (
	StatusID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	StatusName varchar (50) NOT NULL ,
	Comments text NULL ,
	OrderNo int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblTheme (
	ThemeID int AUTO_INCREMENT   NOT NULL PRIMARY KEY,
	ThemeName varchar (100) NOT NULL ,
	Synopsis text NULL ,
	Description text NULL ,
	WebPath varchar (255) NULL ,
	AuthorName varchar (50) NULL ,
	CreationDate datetime NULL ,
	TotalPosRating int NOT NULL default 0,
	TotalNegRating int NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblUser (
	UserID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	Username varchar (32) NULL ,
	Password varchar (64) NULL ,
	Firstname varchar (32) NULL ,
	Middlename varchar (32) NULL ,
	Lastname varchar (32) NULL ,
	EmailAddress varchar (100) NULL ,
	EmailAddressAlt varchar (100) NULL ,
	DayPhone varchar (16) NULL ,
	EvePhone varchar (16) NULL ,
	BestCallTime char (1) NULL ,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblUserRight (
	RightID int AUTO_INCREMENT NOT NULL PRIMARY KEY,
	ParentRightID int NOT NULL default 0,
	RightName varchar (50) NOT NULL ,
	Hyperlink varchar (50) NOT NULL ,
	AdminMenuName varchar(50) NULL,
	AccessKey varchar (20) NULL ,
	HasAdd tinyint NOT NULL default 0,
	HasEdit tinyint NOT NULL default 0,
	HasDelete tinyint NOT NULL default 0,
	HasView tinyint NOT NULL default 0,
	Active tinyint NOT NULL default 1,
	Archive tinyint NOT NULL default 0,
	OrderNo int NOT NULL default 0,
	Modified timestamp NOT NULL ,
	Created timestamp NOT NULL 
);


CREATE TABLE tblUserToRight (
	UserID int NOT NULL ,
	RightID int NOT NULL ,
	CanAdd tinyint NOT NULL default 0,
	CanEdit tinyint NOT NULL default 0,
	CanDelete tinyint NOT NULL default 0,
	CanView tinyint NOT NULL default 0,
	PRIMARY KEY (UserID, RightID)
);

