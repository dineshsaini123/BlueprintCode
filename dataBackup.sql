USE [chitraData]
GO
/****** Object:  Table [dbo].[SLEDGER]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SLEDGER](
	[gledger] [nvarchar](40) NOT NULL,
	[SUBLEDGER] [nvarchar](50) NOT NULL,
	[DESCFORINVOICE] [nvarchar](40) NULL,
	[YEAROPENING] [float] NULL,
	[DISCATEGORY] [nvarchar](50) NULL,
	[DISTCODE] [nvarchar](50) NULL,
	[ADDRESS1] [nvarchar](50) NULL,
	[ADDRESS2] [nvarchar](50) NULL,
	[ADDRESS3] [nvarchar](50) NULL,
	[Phone] [nvarchar](50) NULL,
	[Owner] [float] NULL,
	[Offdays] [nvarchar](40) NULL,
	[OP] [float] NULL,
	[drcr] [nvarchar](50) NULL,
	[party] [nvarchar](100) NULL,
	[code] [nvarchar](50) NULL,
	[contactp] [nvarchar](50) NULL,
	[remarks] [nvarchar](250) NULL,
	[Category2] [nvarchar](15) NULL,
	[Category3] [nvarchar](15) NULL,
	[PartyRemarks] [nvarchar](100) NULL,
	[states] [nvarchar](50) NULL,
	[Email] [nvarchar](50) NULL,
	[mobile] [nvarchar](40) NULL,
	[Dup_party] [bit] NULL,
	[UPGuide] [bit] NULL,
	[UKGuide] [bit] NULL,
	[BBA_BCA] [bit] NULL,
	[Btech] [bit] NULL,
	[Ent_Guide] [bit] NULL,
	[Gar_Adhyan] [bit] NULL,
	[Mrt_Adhyan] [bit] NULL,
	[Bteck] [bit] NULL,
	[En_Guide] [bit] NULL,
	[balance] [float] NULL,
	[setupid] [tinyint] NULL,
	[fyear] [nvarchar](10) NULL,
	[printer_binder] [nvarchar](1) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SizeMaster]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SizeMaster](
	[size1] [nvarchar](10) NOT NULL,
	[size_info] [nvarchar](80) NULL,
	[fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL,
 CONSTRAINT [PK_SizeMaster] PRIMARY KEY CLUSTERED 
(
	[size1] ASC,
	[fyear] ASC,
	[setupid] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[setup2]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[setup2](
	[cname] [nvarchar](50) NULL,
	[add1] [nvarchar](50) NULL,
	[add2] [nvarchar](50) NULL,
	[city] [nvarchar](30) NULL,
	[yarfrom] [datetime] NULL,
	[yarto] [datetime] NULL,
	[phone1] [nvarchar](50) NULL,
	[phone2] [nvarchar](50) NULL,
	[phone3] [nvarchar](50) NULL,
	[phone4] [nvarchar](50) NULL,
	[fax] [nvarchar](50) NULL,
	[email] [nvarchar](100) NULL,
	[rem1] [nvarchar](255) NULL,
	[rem2] [nvarchar](255) NULL,
	[court] [nvarchar](100) NULL,
	[bankadviceno] [int] NULL,
	[uptt] [nvarchar](50) NULL,
	[cst] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[setup1]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[setup1](
	[cname] [nvarchar](50) NULL,
	[add1] [nvarchar](100) NULL,
	[add2] [nvarchar](100) NULL,
	[city] [nvarchar](30) NULL,
	[yarfrom] [smalldatetime] NULL,
	[yarto] [smalldatetime] NULL,
	[phone1] [nvarchar](50) NULL,
	[phone2] [nvarchar](50) NULL,
	[phone3] [nvarchar](50) NULL,
	[phone4] [nvarchar](50) NULL,
	[fax] [nvarchar](50) NULL,
	[email] [nvarchar](100) NULL,
	[rem1] [nvarchar](255) NULL,
	[rem2] [nvarchar](255) NULL,
	[court] [nvarchar](255) NULL,
	[bankadviceno] [int] NULL,
	[uptt] [nvarchar](50) NULL,
	[cst] [nvarchar](50) NULL,
	[fyear] [nvarchar](10) NULL,
	[updatedby] [smalldatetime] NULL,
	[updatedon] [smalldatetime] NULL,
	[othercompmasterdata] [bit] NULL,
	[setupid] [tinyint] NOT NULL,
	[viewallcomp] [bit] NULL,
	[manufacture] [bit] NULL,
	[invrpt] [nvarchar](100) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[setup]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[setup](
	[cname] [nvarchar](50) NULL,
	[add1] [nvarchar](50) NULL,
	[add2] [nvarchar](50) NULL,
	[city] [nvarchar](30) NULL,
	[yarfrom] [datetime] NULL,
	[yarto] [datetime] NULL,
	[phone1] [nvarchar](50) NULL,
	[phone2] [nvarchar](50) NULL,
	[phone3] [nvarchar](50) NULL,
	[phone4] [nvarchar](50) NULL,
	[fax] [nvarchar](50) NULL,
	[email] [nvarchar](100) NULL,
	[rem1] [nvarchar](255) NULL,
	[rem2] [nvarchar](255) NULL,
	[court] [nvarchar](100) NULL,
	[bankadviceno] [int] NULL,
	[uptt] [nvarchar](50) NULL,
	[cst] [nvarchar](50) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[RPTTEMPINDIS1]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[RPTTEMPINDIS1](
	[VDATE] [datetime] NULL,
	[VNO] [int] NULL,
	[SUBLEGER] [nvarchar](50) NULL,
	[DISTNAME] [nvarchar](30) NULL,
	[BNETAMT] [float] NULL,
	[BOOKCODE] [nvarchar](6) NULL,
	[GROUPCODE] [nvarchar](7) NULL,
	[VTYPE] [nvarchar](1) NULL,
	[userid] [int] NULL,
	[groupcheck] [nvarchar](10) NULL,
	[AGENTNAME ] [nvarchar](30) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[remarks]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[remarks](
	[Head] [nvarchar](100) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ReceiveIssueParty]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ReceiveIssueParty](
	[RecNo] [int] NULL,
	[Dates] [datetime] NULL,
	[PartyName] [nvarchar](100) NULL,
	[Particullar] [nvarchar](150) NULL,
	[Dr] [float] NULL,
	[Cr] [float] NULL,
	[Op] [float] NULL,
	[Remarks] [nvarchar](255) NULL,
	[Aouto] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[treport]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[treport](
	[text] [nvarchar](50) NULL,
	[genledger] [nvarchar](50) NULL,
	[subledger] [nvarchar](100) NULL,
	[OpeningBalance] [float] NULL,
	[vdate] [smalldatetime] NULL,
	[vtype] [nvarchar](3) NULL,
	[vno] [int] NULL,
	[narration] [nvarchar](255) NULL,
	[ad] [money] NULL,
	[ac] [money] NULL,
	[dorc] [nvarchar](1) NULL,
	[cbno] [nvarchar](50) NULL,
	[balance] [money] NULL,
	[seprator] [bit] NULL,
	[tab] [int] NULL,
	[sepwidth] [int] NULL,
	[period] [nvarchar](50) NULL,
	[header] [nvarchar](50) NULL,
	[sno] [int] IDENTITY(1,1) NOT NULL,
	[userid] [int] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[transportmaster]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[transportmaster](
	[Transportname] [nvarchar](30) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Winrpt]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Winrpt](
	[Date1] [smalldatetime] NULL,
	[OpDes] [nvarchar](150) NULL,
	[op] [float] NULL,
	[Receipt] [float] NULL,
	[Payment] [float] NULL,
	[Balance] [float] NULL,
	[FromDate] [smalldatetime] NULL,
	[ToDate] [smalldatetime] NULL,
	[Description] [nvarchar](100) NULL,
	[Party] [nvarchar](200) NULL,
	[Narration] [nvarchar](255) NULL,
	[Qty] [float] NULL,
	[Closing] [float] NULL,
	[Closing1] [float] NULL,
	[dr] [nvarchar](25) NULL,
	[cr] [nvarchar](25) NULL,
	[aouto] [int] IDENTITY(1,1) NOT NULL,
	[UID] [int] NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[VOUCHERS]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[VOUCHERS](
	[VoucherType] [nvarchar](1) NOT NULL,
	[VoucherDate] [datetime] NOT NULL,
	[VoucherNumber] [int] NOT NULL,
	[GenLedger] [nvarchar](40) NULL,
	[SubLedger] [nvarchar](40) NULL,
	[Amount] [float] NULL,
	[DebitorCredit] [nvarchar](1) NULL,
	[CBND] [nvarchar](20) NULL,
	[EntryNumber] [float] NULL,
	[DESCRIPTION] [nvarchar](50) NULL,
	[vsno] [int] IDENTITY(1,1) NOT NULL,
	[CashCheck] [bit] NULL,
	[setupid] [tinyint] NULL,
	[fyear] [nvarchar](10) NULL,
 CONSTRAINT [PK_VOUCHERS] PRIMARY KEY CLUSTERED 
(
	[VoucherDate] ASC,
	[VoucherNumber] ASC,
	[vsno] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[UsrePermission]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[UsrePermission](
	[Type] [nvarchar](25) NULL,
	[TaskName] [nvarchar](50) NULL,
	[Permission] [nvarchar](50) NULL,
	[Save] [nvarchar](50) NULL,
	[Delete] [nvarchar](50) NULL,
	[Edit] [nvarchar](50) NULL,
	[TaskType] [nvarchar](50) NULL,
	[UserName] [nvarchar](50) NULL,
	[password] [nvarchar](20) NULL,
	[fyear] [nvarchar](10) NULL,
	[userId] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[AGENTMASTER]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AGENTMASTER](
	[AGENTNAME] [nvarchar](30) NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tmps_LEDGER1]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tmps_LEDGER1](
	[gledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](50) NULL,
	[DESCFORINVOICE] [nvarchar](40) NULL,
	[YEAROPENING] [float] NULL,
	[DISCATEGORY] [nvarchar](50) NULL,
	[DISTCODE] [nvarchar](50) NULL,
	[ADDRESS1] [nvarchar](50) NULL,
	[ADDRESS2] [nvarchar](50) NULL,
	[ADDRESS3] [nvarchar](50) NULL,
	[Phone] [nvarchar](50) NULL,
	[Owner] [nvarchar](30) NULL,
	[Offdays] [nvarchar](40) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tmps_LEDGER]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tmps_LEDGER](
	[gledger] [nvarchar](40) NOT NULL,
	[SUBLEDGER] [nvarchar](50) NOT NULL,
	[DESCFORINVOICE] [nvarchar](40) NULL,
	[YEAROPENING] [float] NULL,
	[DISCATEGORY] [nvarchar](50) NULL,
	[DISTCODE] [nvarchar](50) NULL,
	[ADDRESS1] [nvarchar](50) NULL,
	[ADDRESS2] [nvarchar](50) NULL,
	[ADDRESS3] [nvarchar](50) NULL,
	[Phone] [nvarchar](50) NULL,
	[Owner] [float] NULL,
	[Offdays] [nvarchar](40) NULL,
	[OP] [float] NULL,
	[drcr] [nvarchar](50) NULL,
	[party] [nvarchar](100) NULL,
	[code] [nvarchar](50) NULL,
	[contactp] [nvarchar](50) NULL,
	[remarks] [nvarchar](250) NULL,
	[Category2] [nvarchar](15) NULL,
	[Category3] [nvarchar](15) NULL,
	[PartyRemarks] [nvarchar](100) NULL,
	[states] [nvarchar](50) NULL,
	[Email] [nvarchar](50) NULL,
	[mobile] [nvarchar](40) NULL,
	[Dup_party] [bit] NULL,
	[UPGuide] [bit] NULL,
	[UKGuide] [bit] NULL,
	[BBA_BCA] [bit] NULL,
	[Btech] [bit] NULL,
	[Ent_Guide] [bit] NULL,
	[Gar_Adhyan] [bit] NULL,
	[Mrt_Adhyan] [bit] NULL,
	[Bteck] [bit] NULL,
	[En_Guide] [bit] NULL,
	[balance] [float] NULL,
	[setupid] [tinyint] NULL,
	[fyear] [nvarchar](10) NULL,
	[printer_binder] [nvarchar](1) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TmpBook]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TmpBook](
	[BCode] [nvarchar](20) NULL,
	[BName] [nvarchar](50) NULL,
	[Qty] [int] NULL,
	[Login] [nvarchar](25) NULL,
	[Head] [nvarchar](25) NULL,
	[Ason] [nvarchar](25) NULL,
	[Auto] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TemprptTrialBalance]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TemprptTrialBalance](
	[Gledger] [nvarchar](40) NULL,
	[Subledger] [nvarchar](100) NULL,
	[OpeningBalance] [float] NULL,
	[DAmount] [float] NULL,
	[CAmount] [float] NULL,
	[ClosingBalance] [float] NULL,
	[userid] [int] NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tempLedgerRpt]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tempLedgerRpt](
	[Billtype] [nvarchar](50) NULL,
	[Bill] [int] NULL,
	[dates] [datetime] NULL,
	[Des] [nvarchar](250) NULL,
	[Dr] [float] NULL,
	[Cr] [float] NULL,
	[Balance] [float] NULL,
	[drcr] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tempLedger2]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tempLedger2](
	[dates] [datetime] NULL,
	[Billtype] [nvarchar](50) NULL,
	[Bill] [int] NULL,
	[Des] [nvarchar](250) NULL,
	[Dr] [float] NULL,
	[Cr] [float] NULL,
	[Balance] [float] NULL,
	[drcr] [nvarchar](50) NULL,
	[Party] [nvarchar](100) NULL,
	[OpBalance] [float] NULL,
	[Aouto] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tempLedger1]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tempLedger1](
	[dates] [datetime] NULL,
	[Billtype] [nvarchar](50) NULL,
	[Bill] [int] NULL,
	[Des] [nvarchar](250) NULL,
	[Dr] [float] NULL,
	[Cr] [float] NULL,
	[Balance] [float] NULL,
	[drcr] [nvarchar](50) NULL,
	[Party] [nvarchar](100) NULL,
	[OpBalance] [float] NULL,
	[Aouto] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[tempLedger]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tempLedger](
	[dates] [datetime] NULL,
	[Billtype] [nvarchar](50) NULL,
	[Bill] [int] NULL,
	[Des] [nvarchar](250) NULL,
	[Dr] [float] NULL,
	[Cr] [float] NULL,
	[Balance] [float] NULL,
	[drcr] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[TEMPCRITNOTE]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TEMPCRITNOTE](
	[IN] [int] NULL,
	[ID] [nvarchar](10) NULL,
	[GLD] [nvarchar](40) NULL,
	[SLD] [nvarchar](40) NULL,
	[BC] [nvarchar](6) NULL,
	[Q] [float] NULL,
	[R] [float] NULL,
	[D] [float] NULL,
	[D_] [float] NULL,
	[A] [float] NULL,
	[NA] [float] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[teacher]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[teacher](
	[School] [nvarchar](250) NULL,
	[Name] [nvarchar](50) NULL,
	[Agent] [nvarchar](50) NULL,
	[Subject] [nvarchar](50) NULL,
	[Address] [nvarchar](200) NULL,
	[Address2] [nvarchar](200) NULL,
	[District] [nvarchar](50) NULL,
	[City] [nvarchar](50) NULL,
	[Phone] [nvarchar](50) NULL,
	[DOB] [datetime] NULL,
	[Andate] [datetime] NULL,
	[auto] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[subledgertrail]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[subledgertrail](
	[subledger] [nvarchar](50) NULL,
	[yearopening] [float] NULL,
	[OPAMOUNTdebit] [float] NULL,
	[OPAMOUNTCREDIT] [float] NULL,
	[userid] [int] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[subjectName]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[subjectName](
	[subject] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[StockRegister]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[StockRegister](
	[Dates] [datetime] NULL,
	[BOOKCODE] [nvarchar](10) NULL,
	[Qty] [int] NULL,
	[Category] [nvarchar](20) NULL,
	[Issue_Receive] [nvarchar](15) NULL,
	[Issue_ReceveFrom] [nvarchar](50) NULL,
	[BinderName] [nvarchar](100) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[StockHead]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[StockHead](
	[Head] [nvarchar](50) NULL,
	[id] [int] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[states]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[states](
	[states] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[State]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[State](
	[CountryID] [nvarchar](6) NULL,
	[StateID] [nvarchar](6) NULL,
	[State] [nvarchar](50) NULL,
	[StateKey] [nvarchar](200) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[printsetup]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[printsetup](
	[reportname] [nvarchar](50) NULL,
	[comp] [bit] NOT NULL,
	[totalcolumn] [int] NULL,
	[paperwidth] [nvarchar](50) NULL,
	[totalline] [int] NULL,
	[paperheight] [nvarchar](50) NULL,
	[topmargin] [int] NULL,
	[bottommargin] [int] NULL,
	[leftmargin] [int] NULL,
	[rightmargin] [int] NULL,
	[header] [bit] NOT NULL,
	[footer] [bit] NOT NULL,
	[headertext] [nvarchar](50) NULL,
	[footertext] [nvarchar](50) NULL,
	[headerpos] [nvarchar](1) NULL,
	[footerpos] [nvarchar](1) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pass]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[pass](
	[pass] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[PaperMakeMaster]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[PaperMakeMaster](
	[papermaker_id] [nvarchar](30) NULL,
	[papermaker_name] [nvarchar](100) NULL,
	[add1] [nvarchar](100) NULL,
	[add2] [nvarchar](100) NULL,
	[city] [nvarchar](100) NULL,
	[pno1] [nvarchar](100) NULL,
	[pno2] [nvarchar](100) NULL,
	[mobile] [nvarchar](100) NULL,
	[faxno] [nvarchar](100) NULL,
	[eid] [nvarchar](100) NULL,
	[manufacturer] [bit] NULL,
	[Contact1] [nvarchar](100) NULL,
	[Contact2] [nvarchar](100) NULL,
	[Contact3] [nvarchar](100) NULL,
	[Contact4] [nvarchar](100) NULL,
	[Contact5] [nvarchar](100) NULL,
	[Mobile1] [nvarchar](100) NULL,
	[Mobile2] [nvarchar](100) NULL,
	[Mobile3] [nvarchar](100) NULL,
	[Mobile4] [nvarchar](100) NULL,
	[Mobile5] [nvarchar](100) NULL,
	[Eco] [nvarchar](50) NULL,
	[GSM] [nvarchar](50) NULL,
	[Size] [nvarchar](50) NULL,
	[Bright] [nvarchar](50) NULL,
	[PType] [nvarchar](50) NULL,
	[SizeValue1] [nvarchar](10) NULL,
	[SizeValue2] [nvarchar](10) NULL,
	[SizeValue3] [nvarchar](10) NULL,
	[Size1] [nvarchar](10) NULL,
	[Size2] [nvarchar](10) NULL,
	[Size3] [nvarchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Ordermnm]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Ordermnm](
	[id] [int] NULL,
	[Name] [nvarchar](100) NULL,
	[dates] [datetime] NULL,
	[Godwn] [nvarchar](10) NULL,
	[Status] [bit] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Menu]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Menu](
	[Type] [nvarchar](25) NULL,
	[Menu] [nvarchar](50) NULL,
	[MenuName] [nvarchar](50) NULL,
	[Hide] [nvarchar](1) NULL,
	[user] [nvarchar](25) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[MasterTbl]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[MasterTbl](
	[Name] [nvarchar](50) NULL,
	[Category] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Login]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Login](
	[Uname] [nvarchar](50) NULL,
	[Password] [nvarchar](50) NULL,
	[mnuGenralLedgermaster] [bit] NOT NULL,
	[mnuSubLedgermaster] [bit] NOT NULL,
	[mnuinvendmaster] [bit] NOT NULL,
	[mnucreendmaster] [bit] NOT NULL,
	[mnucashcountersalesmaster] [bit] NOT NULL,
	[mnudiscountcategorymaster] [bit] NOT NULL,
	[mnubookmaster] [bit] NOT NULL,
	[mnubookgroupmaster] [bit] NOT NULL,
	[mnuagentmaster] [bit] NOT NULL,
	[mnudistrictmaster] [bit] NOT NULL,
	[mnureportbookgroupmaster] [bit] NOT NULL,
	[rptcashbankbook] [bit] NOT NULL,
	[rptgenralledgerac] [bit] NOT NULL,
	[rptsubledgerac] [bit] NOT NULL,
	[rptalphawisesubledgerac] [bit] NOT NULL,
	[rptgenledgertrialbalance] [bit] NOT NULL,
	[rptgenledgeropentrial] [bit] NOT NULL,
	[rptsubledgertrialbalance] [bit] NOT NULL,
	[rptdistictwisesales] [bit] NOT NULL,
	[rptdistictwisesalesreturn] [bit] NOT NULL,
	[rptbookgroupwisesales] [bit] NOT NULL,
	[rptbankadvicereconcilation] [bit] NOT NULL,
	[rptbankadvice] [bit] NOT NULL,
	[mnuvoucherentry] [bit] NOT NULL,
	[mnusalesinvoice] [bit] NOT NULL,
	[mnucreditnoteitem] [bit] NOT NULL,
	[mnucashcountersales] [bit] NOT NULL,
	[mnucreditnote] [bit] NOT NULL,
	[mnudebitnote] [bit] NOT NULL,
	[createuser] [bit] NOT NULL,
	[mnutoolsetup] [bit] NOT NULL,
	[UserId] [smallint] NULL,
	[BAdd] [bit] NOT NULL,
	[Bedit] [bit] NOT NULL,
	[Bsave] [bit] NOT NULL,
	[BDelete] [bit] NOT NULL,
	[fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ItemCreation]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ItemCreation](
	[ItemCode] [int] NULL,
	[CourseName] [nvarchar](50) NULL,
	[ItemName] [nvarchar](100) NULL,
	[Price] [float] NULL,
	[PerHeadAmt] [int] NULL,
	[Head] [nvarchar](50) NULL,
	[Unit] [nvarchar](50) NULL,
	[Opening] [float] NULL,
	[OrderL] [int] NULL,
	[fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[issueregister]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[issueregister](
	[INVOICENO] [int] NOT NULL,
	[INVOICEDATE] [smalldatetime] NOT NULL,
	[BOOKCODE] [nvarchar](10) NULL,
	[BOOKNAME] [nvarchar](40) NULL,
	[QUANTITY] [float] NULL,
	[GROUPCODE] [nvarchar](7) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IssueDeppt]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IssueDeppt](
	[BillNo] [int] NULL,
	[Supplier] [nvarchar](100) NULL,
	[Dates] [datetime] NULL,
	[gp] [nvarchar](50) NULL,
	[ItemName] [nvarchar](100) NULL,
	[Qty] [float] NULL,
	[Unit] [nvarchar](50) NULL,
	[Price] [float] NULL,
	[Amt] [float] NULL,
	[TotalAmt] [float] NULL,
	[Remarks] [nvarchar](200) NULL,
	[Deppt] [nvarchar](100) NULL,
	[DemandNo] [int] NULL,
	[fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IssueC]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IssueC](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [smalldatetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](100) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DQurey_fields]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DQurey_fields](
	[DESCRIPTION] [nvarchar](80) NULL,
	[ADMFIELDNAME] [nvarchar](50) NULL,
	[ADMFIELDTYPE] [float] NULL,
	[admsortorder] [smallint] NULL,
	[AdmFieldSort] [smallint] NULL,
	[ascdec] [nvarchar](1) NULL,
	[SEARCHSTATUS] [nvarchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DQurey_critaria]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DQurey_critaria](
	[DESCRIPTION] [nvarchar](80) NULL,
	[SEARCHFIELDNAME1] [nvarchar](50) NULL,
	[SIMPLEOPERATOR1] [nvarchar](50) NULL,
	[SEARCHVALUE1] [nvarchar](50) NULL,
	[LOGICALOPERATOR1] [nvarchar](50) NULL,
	[SEARCHSTATUS] [nvarchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DQurey]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DQurey](
	[DESCRIPTION] [nvarchar](80) NULL,
	[CRITARIA] [nvarchar](255) NULL,
	[REPORTNAME] [nvarchar](50) NULL,
	[SEARCHSTATUS] [nvarchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DonationB]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DonationB](
	[EntNo] [int] NULL,
	[BookCode] [nvarchar](10) NULL,
	[Teacher] [nvarchar](50) NULL,
	[Amount] [float] NULL,
	[Others] [nvarchar](50) NULL,
	[sno] [int] NULL,
	[auto] [int] IDENTITY(1,1) NOT NULL,
	[setupid] [tinyint] NULL,
	[fyear] [nvarchar](10) NULL,
	[Status] [nvarchar](6) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DonationA]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DonationA](
	[EntNo] [int] NULL,
	[Dates] [datetime] NULL,
	[AgentName] [nvarchar](40) NULL,
	[SchoolName] [nvarchar](250) NULL,
	[setupid] [tinyint] NULL,
	[fyear] [nvarchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DNFB]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DNFB](
	[DNN] [float] NULL,
	[DND] [datetime] NULL,
	[GLD] [nvarchar](40) NULL,
	[SLD] [nvarchar](40) NULL,
	[A] [float] NULL,
	[DC] [nvarchar](1) NULL,
	[groupcode] [nvarchar](10) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DNFA]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DNFA](
	[DNN] [float] NULL,
	[DND] [datetime] NULL,
	[PGLD] [nvarchar](40) NULL,
	[PSLD] [nvarchar](40) NULL,
	[NA] [float] NULL,
	[N] [nvarchar](40) NULL,
	[DC] [nvarchar](1) NULL,
	[Agentname] [nvarchar](50) NULL,
	[BAuthorized] [bit] NULL,
	[print_yes] [nvarchar](50) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DISTRICTS]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DISTRICTS](
	[DISTRICTNAME] [nvarchar](30) NULL,
	[AGENTNAME] [nvarchar](30) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DispatchRegister]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DispatchRegister](
	[SNO] [int] NULL,
	[Date] [datetime] NULL,
	[Particulars] [nvarchar](240) NULL,
	[RR] [nvarchar](10) NULL,
	[GR] [int] NULL,
	[GR_DT] [datetime] NULL,
	[BDL] [nvarchar](100) NULL,
	[WT] [nvarchar](50) NULL,
	[Freight] [nvarchar](50) NULL,
	[Recd_dt] [datetime] NULL,
	[Freight_Paid] [nvarchar](50) NULL,
	[CNO] [int] NULL,
	[RNO] [int] NULL,
	[Remarks] [nvarchar](254) NULL,
	[gp] [nvarchar](50) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DISCCATS]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DISCCATS](
	[categorycode] [nvarchar](7) NULL,
	[groupcode] [nvarchar](7) NULL,
	[discountrate] [float] NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[AutoId] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Issueb]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Issueb](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [smalldatetime] NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](100) NULL,
	[BOOKCODE] [nvarchar](10) NULL,
	[QUANTITY] [float] NULL,
	[RATE] [float] NULL,
	[DISCOUNT] [float] NULL,
	[PRINTORDER] [float] NULL,
	[AMOUNT] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[SNO] [int] NOT NULL,
	[btno] [nvarchar](50) NULL,
	[mgdate] [nvarchar](20) NULL,
	[expdate] [nvarchar](20) NULL,
	[per] [nvarchar](50) NULL,
	[fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[IssueA]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IssueA](
	[INVOICENO] [int] NOT NULL,
	[INVOICEDATE] [smalldatetime] NOT NULL,
	[GENLEDGER] [nvarchar](40) NOT NULL,
	[SUBLEDGER] [nvarchar](100) NOT NULL,
	[ORDERNO] [nvarchar](10) NULL,
	[ORDERBY] [nvarchar](15) NULL,
	[ORDERDATE] [smalldatetime] NULL,
	[MARKA] [nvarchar](10) NULL,
	[BUNDLES] [nvarchar](15) NULL,
	[THROUGH] [nvarchar](50) NULL,
	[STATION] [nvarchar](50) NULL,
	[BILTYNO] [nvarchar](10) NULL,
	[BILTYDATE] [smalldatetime] NULL,
	[FREIGHT] [nvarchar](15) NULL,
	[WEIGHT] [nvarchar](10) NULL,
	[TXT1] [nvarchar](20) NULL,
	[TXT1A] [float] NULL,
	[TXT2] [nvarchar](20) NULL,
	[TXT2A] [float] NULL,
	[BAA] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[GAMOUNT] [float] NULL,
	[T2] [nvarchar](40) NULL,
	[THROUGH1] [nvarchar](50) NULL,
	[district] [nvarchar](30) NULL,
	[AdviceStatus] [nvarchar](20) NULL,
	[AdviceRemark] [nvarchar](100) NULL,
	[AgentName] [nvarchar](30) NULL,
	[aexp1] [nvarchar](50) NULL,
	[aexp2] [nvarchar](50) NULL,
	[aexp3] [nvarchar](50) NULL,
	[lexp1] [nvarchar](50) NULL,
	[lexp2] [nvarchar](50) NULL,
	[lexp3] [nvarchar](50) NULL,
	[aexp1am] [int] NULL,
	[aexp1rate] [int] NULL,
	[aexp2am] [int] NULL,
	[aexp2rate] [int] NULL,
	[aexp3am] [int] NULL,
	[aexp3rate] [int] NULL,
	[lexp1am] [int] NULL,
	[lexp1rate] [int] NULL,
	[lexp2am] [int] NULL,
	[lexp2rate] [int] NULL,
	[lexp3am] [int] NULL,
	[lexp3rate] [int] NULL,
	[RecAmt] [float] NULL,
	[notshow] [nvarchar](50) NULL,
	[fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Issue_ReceiveMaster]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Issue_ReceiveMaster](
	[Name] [nvarchar](50) NULL,
	[Category] [nvarchar](20) NULL,
	[id] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[InvoiceSubLedger]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[InvoiceSubLedger](
	[TaxHead] [nvarchar](50) NULL,
	[SubLedger] [nvarchar](50) NULL,
	[Up_Exup] [nvarchar](50) NULL,
	[auto] [numeric](18, 0) IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEEND]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEEND](
	[CGENLEDGER] [nvarchar](40) NULL,
	[CSUBLEDGER] [nvarchar](40) NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RATE] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[PrintOrder] [smallint] NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[type] [nvarchar](20) NULL,
	[auto] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[invoicedueamt]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[invoicedueamt](
	[INVOICENO] [int] NOT NULL,
	[INVOICEDATE] [smalldatetime] NOT NULL,
	[GENLEDGER] [nvarchar](40) NOT NULL,
	[SUBLEDGER] [nvarchar](50) NOT NULL,
	[NETAMOUNT] [float] NULL,
	[recamt] [float] NOT NULL,
	[balance] [float] NULL,
	[setupid] [tinyint] NULL,
	[fyear] [nvarchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICECtmp_spRet]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICECtmp_spRet](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [nvarchar](10) NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICECtmp_sp]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICECtmp_sp](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [nvarchar](10) NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[invoicectmp]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[invoicectmp](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL,
	[saleType] [nvarchar](100) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[invoicecr]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[invoicecr](
	[CreditAmt] [float] NULL,
	[CBND] [nvarchar](20) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[GenLedger] [nvarchar](40) NOT NULL,
	[SubLedger] [nvarchar](40) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEC_spRet]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEC_spRet](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEC_sp]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEC_sp](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEA_spRet]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEA_spRet](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[ORDERBY] [nvarchar](15) NULL,
	[ORDERDATE] [datetime] NULL,
	[MARKA] [nvarchar](10) NULL,
	[BUNDLES] [nvarchar](15) NULL,
	[THROUGH] [nvarchar](50) NULL,
	[STATION] [nvarchar](50) NULL,
	[BILTYNO] [nvarchar](10) NULL,
	[BILTYDATE] [datetime] NULL,
	[FREIGHT] [nvarchar](15) NULL,
	[WEIGHT] [nvarchar](10) NULL,
	[TXT1] [nvarchar](20) NULL,
	[TXT1A] [float] NULL,
	[TXT2] [nvarchar](20) NULL,
	[TXT2A] [float] NULL,
	[BAA] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[GAMOUNT] [float] NULL,
	[T2] [nvarchar](40) NULL,
	[THROUGH1] [nvarchar](50) NULL,
	[district] [nvarchar](30) NULL,
	[AdviceStatus] [nvarchar](20) NULL,
	[AdviceRemark] [nvarchar](100) NULL,
	[AgentName] [nvarchar](30) NULL,
	[transportname] [nvarchar](30) NULL,
	[godown] [nvarchar](5) NULL,
	[BAuthorized] [bit] NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEA_sp]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEA_sp](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[ORDERBY] [nvarchar](15) NULL,
	[ORDERDATE] [datetime] NULL,
	[MARKA] [nvarchar](10) NULL,
	[BUNDLES] [nvarchar](15) NULL,
	[THROUGH] [nvarchar](50) NULL,
	[STATION] [nvarchar](50) NULL,
	[BILTYNO] [nvarchar](10) NULL,
	[BILTYDATE] [datetime] NULL,
	[FREIGHT] [nvarchar](15) NULL,
	[WEIGHT] [nvarchar](10) NULL,
	[TXT1] [nvarchar](20) NULL,
	[TXT1A] [float] NULL,
	[TXT2] [nvarchar](20) NULL,
	[TXT2A] [float] NULL,
	[BAA] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[GAMOUNT] [float] NULL,
	[T2] [nvarchar](40) NULL,
	[THROUGH1] [nvarchar](50) NULL,
	[district] [nvarchar](30) NULL,
	[AdviceStatus] [nvarchar](20) NULL,
	[AdviceRemark] [nvarchar](100) NULL,
	[AgentName] [nvarchar](30) NULL,
	[transportname] [nvarchar](30) NULL,
	[godown] [nvarchar](5) NULL,
	[BAuthorized] [bit] NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEA_IssuedBind]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEA_IssuedBind](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](50) NULL,
	[ORDERBY] [nvarchar](15) NULL,
	[ORDERDATE] [datetime] NULL,
	[add1] [nvarchar](200) NULL,
	[add2] [nvarchar](200) NULL,
	[remarks] [nvarchar](150) NULL,
	[NetBook] [int] NULL,
	[BILTYNO] [nvarchar](10) NULL,
	[BILTYDATE] [datetime] NULL,
	[FREIGHT] [nvarchar](15) NULL,
	[WEIGHT] [nvarchar](10) NULL,
	[TXT1] [nvarchar](20) NULL,
	[TXT1A] [float] NULL,
	[TXT2] [nvarchar](20) NULL,
	[TXT2A] [float] NULL,
	[BAA] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[GAMOUNT] [float] NULL,
	[T2] [nvarchar](40) NULL,
	[THROUGH1] [nvarchar](50) NULL,
	[district] [nvarchar](30) NULL,
	[AdviceStatus] [nvarchar](20) NULL,
	[AdviceRemark] [nvarchar](100) NULL,
	[AgentName] [nvarchar](30) NULL,
	[advno] [int] NULL,
	[transportname] [nvarchar](30) NULL,
	[BAuthorized] [bit] NULL,
	[print_yes] [nvarchar](50) NULL,
	[godown] [nvarchar](5) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEA]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEA](
	[INVOICENO] [int] NOT NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[ORDERBY] [nvarchar](15) NULL,
	[ORDERDATE] [datetime] NULL,
	[MARKA] [nvarchar](10) NULL,
	[BUNDLES] [nvarchar](15) NULL,
	[THROUGH] [nvarchar](50) NULL,
	[STATION] [nvarchar](50) NULL,
	[BILTYNO] [nvarchar](10) NULL,
	[BILTYDATE] [datetime] NULL,
	[FREIGHT] [nvarchar](15) NULL,
	[WEIGHT] [nvarchar](10) NULL,
	[TXT1] [nvarchar](20) NULL,
	[TXT1A] [float] NULL,
	[TXT2] [nvarchar](20) NULL,
	[TXT2A] [float] NULL,
	[BAA] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[GAMOUNT] [float] NULL,
	[T2] [nvarchar](40) NULL,
	[THROUGH1] [nvarchar](50) NULL,
	[district] [nvarchar](30) NULL,
	[AdviceStatus] [nvarchar](20) NULL,
	[AdviceRemark] [nvarchar](100) NULL,
	[AgentName] [nvarchar](30) NULL,
	[advno] [int] NULL,
	[transportname] [nvarchar](30) NULL,
	[BAuthorized] [bit] NULL,
	[print_yes] [nvarchar](50) NULL,
	[Godown] [nvarchar](10) NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL,
	[permitno] [nvarchar](50) NULL,
	[formNo] [nvarchar](50) NULL,
 CONSTRAINT [PK_INVOICEA_1] PRIMARY KEY CLUSTERED 
(
	[INVOICENO] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Info]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Info](
	[vno] [nvarchar](50) NULL,
	[date] [datetime] NULL,
	[aname] [nvarchar](30) NULL,
	[city] [nvarchar](35) NULL,
	[scname] [nvarchar](50) NULL,
	[qty] [int] NULL,
	[district] [nvarchar](50) NULL,
	[setupid] [tinyint] NULL,
	[fyear] [nvarchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[GSMMaster]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GSMMaster](
	[GSM] [nvarchar](10) NOT NULL,
	[gsm_info] [nvarchar](80) NULL,
	[fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL,
 CONSTRAINT [PK_gsmMaster] PRIMARY KEY CLUSTERED 
(
	[GSM] ASC,
	[fyear] ASC,
	[setupid] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[GROUPS]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GROUPS](
	[groupcode] [nvarchar](7) NULL,
	[groupname] [nvarchar](50) NULL,
	[group1] [bit] NULL,
	[group2] [bit] NULL,
	[Group3] [bit] NULL,
	[Group4] [bit] NULL,
	[Category] [nvarchar](25) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[AutoId] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[groupHeading]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[groupHeading](
	[group1] [nvarchar](25) NULL,
	[group2] [nvarchar](25) NULL,
	[group3] [nvarchar](25) NULL,
	[group4] [nvarchar](25) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[grid_ini]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[grid_ini](
	[GridName] [nvarchar](20) NULL,
	[User] [nvarchar](20) NULL,
	[GridCols] [int] NULL,
	[FixRaw] [int] NULL,
	[WordWrap] [bit] NULL,
	[MergCellFix] [bit] NULL,
	[MergCell] [bit] NULL,
	[gridWidth] [int] NULL,
	[gridHieght] [int] NULL,
	[gridTop] [int] NULL,
	[gridLeft] [int] NULL,
	[gridminraw] [int] NULL,
	[Colwidth_0] [int] NULL,
	[Colwidth_1] [int] NULL,
	[Colwidth_2] [int] NULL,
	[Colwidth_3] [int] NULL,
	[Colwidth_4] [int] NULL,
	[Colwidth_5] [int] NULL,
	[Colwidth_6] [int] NULL,
	[Colwidth_7] [int] NULL,
	[Colwidth_8] [int] NULL,
	[Colwidth_9] [int] NULL,
	[Colwidth_10] [int] NULL,
	[Colwidth_11] [int] NULL,
	[Colwidth_12] [int] NULL,
	[Colwidth_13] [int] NULL,
	[Colwidth_14] [int] NULL,
	[Colwidth_15] [int] NULL,
	[Colwidth_16] [int] NULL,
	[Colwidth_17] [int] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Godownmaster]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Godownmaster](
	[Godwn] [nvarchar](20) NULL,
	[id] [int] NOT NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[GLEDGER_old]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GLEDGER_old](
	[Category] [nvarchar](11) NOT NULL,
	[gledger] [nvarchar](40) NOT NULL,
	[PLC] [bit] NOT NULL,
	[BSC] [bit] NOT NULL,
	[SLF] [bit] NOT NULL,
	[YEAROPENING] [float] NULL,
	[cashbankbook] [bit] NOT NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NOT NULL,
	[Auto] [numeric](18, 0) IDENTITY(1,1) NOT NULL,
 CONSTRAINT [PK_GLEDGER] PRIMARY KEY CLUSTERED 
(
	[gledger] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[GLEDGER]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[GLEDGER](
	[Category] [nvarchar](11) NULL,
	[gledger] [nvarchar](40) NOT NULL,
	[PLC] [bit] NULL,
	[BSC] [bit] NULL,
	[SLF] [bit] NULL,
	[YEAROPENING] [float] NULL,
	[cashbankbook] [bit] NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[Auto] [int] IDENTITY(1,1) NOT NULL,
 CONSTRAINT [PK_GLEDGER_1] PRIMARY KEY CLUSTERED 
(
	[gledger] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FirmName]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FirmName](
	[FirmName] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[FinishPurchase]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[FinishPurchase](
	[BillNo] [nvarchar](50) NULL,
	[Supplier] [nvarchar](100) NULL,
	[Dates] [datetime] NULL,
	[gp] [nvarchar](50) NULL,
	[ItemName] [nvarchar](100) NULL,
	[Qty] [float] NULL,
	[Unit] [nvarchar](50) NULL,
	[Price] [float] NULL,
	[Amt] [float] NULL,
	[TotalAmt] [float] NULL,
	[Remarks] [nvarchar](200) NULL,
	[Credit] [nvarchar](50) NULL,
	[challan] [nvarchar](50) NULL,
	[challan_date] [datetime] NULL,
	[clear] [nvarchar](5) NULL,
	[pno] [int] NULL,
	[chno] [nvarchar](25) NULL,
	[dated] [nvarchar](20) NULL,
	[chamt] [float] NULL,
	[bank] [nvarchar](50) NULL,
	[fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[financialyear]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[financialyear](
	[fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[ExsieDetail]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ExsieDetail](
	[Head] [nvarchar](50) NULL,
	[Rate] [float] NULL,
	[OrderId] [int] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CREDITA]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CREDITA](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[ORDERBY] [nvarchar](15) NULL,
	[ORDERDATE] [datetime] NULL,
	[MARKA] [nvarchar](15) NULL,
	[BUNDLES] [nvarchar](15) NULL,
	[THROUGH] [nvarchar](50) NULL,
	[STATION] [nvarchar](50) NULL,
	[BILTYNO] [nvarchar](10) NULL,
	[BILTYDATE] [datetime] NULL,
	[FREIGHT] [nvarchar](15) NULL,
	[WEIGHT] [nvarchar](10) NULL,
	[TXT1] [nvarchar](20) NULL,
	[TXT1A] [float] NULL,
	[TXT2] [nvarchar](20) NULL,
	[TXT2A] [float] NULL,
	[BAA] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[GAMOUNT] [float] NULL,
	[T2] [nvarchar](40) NULL,
	[THROUGH1] [nvarchar](50) NULL,
	[district] [nvarchar](30) NULL,
	[AgentName] [nvarchar](30) NULL,
	[RSNO] [nvarchar](10) NULL,
	[BAuthorized] [bit] NULL,
	[print_yes] [nvarchar](50) NULL,
	[Godown] [nvarchar](10) NULL,
	[mark1] [nvarchar](10) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Country]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Country](
	[CountryID] [nvarchar](6) NULL,
	[Country] [nvarchar](50) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CounterSale_Head]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CounterSale_Head](
	[1] [nvarchar](50) NULL,
	[2] [nvarchar](50) NULL,
	[3] [nvarchar](50) NULL,
	[4] [nvarchar](50) NULL,
	[5] [nvarchar](50) NULL,
	[6] [nvarchar](50) NULL,
	[7] [nvarchar](50) NULL,
	[8] [nvarchar](50) NULL,
	[9] [nvarchar](50) NULL,
	[10] [nvarchar](50) NULL,
	[11] [nvarchar](50) NULL,
	[12] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CounterSale]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CounterSale](
	[1] [float] NULL,
	[2] [float] NULL,
	[3] [float] NULL,
	[4] [float] NULL,
	[5] [float] NULL,
	[6] [float] NULL,
	[7] [float] NULL,
	[8] [float] NULL,
	[9] [float] NULL,
	[10] [float] NULL,
	[11] [float] NULL,
	[12] [float] NULL,
	[Agent] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[copyMaster]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[copyMaster](
	[BookNo] [nvarchar](100) NULL,
	[Book] [nvarchar](100) NULL,
	[ProductQuality] [nvarchar](100) NULL,
	[Trimsize] [nvarchar](100) NULL,
	[wastage] [int] NULL,
	[NoOfPages] [float] NULL,
	[PaperQuality] [nvarchar](100) NULL,
	[PaperMaker] [nvarchar](100) NULL,
	[NoOfCopy_pack] [float] NULL,
	[NoOfPackage] [float] NULL,
	[GSM] [nvarchar](100) NULL,
	[PaperSize] [nvarchar](100) NULL,
	[F_Fold] [int] NULL,
	[ups] [int] NULL,
	[Noofcopy] [int] NULL,
	[cover_ups] [float] NULL,
	[Cover_paper] [nvarchar](100) NULL,
	[Strip] [nvarchar](100) NULL,
	[FYear] [nvarchar](100) NULL,
	[rate] [nvarchar](100) NULL,
	[discount] [float] NULL,
	[Papercode] [nvarchar](100) NULL,
	[rulling] [nvarchar](100) NULL,
	[hardboard] [smallint] NULL,
	[TitleFrontgsm] [nvarchar](100) NULL,
	[TitleBackGsm] [nvarchar](100) NULL,
	[ExtraDiscount] [float] NULL,
	[CutSize] [float] NULL,
	[Weight] [float] NULL,
	[discontinue] [smallint] NULL,
	[CashDiscount] [float] NULL,
	[TaxDiscount] [float] NULL,
	[centercode] [nvarchar](100) NULL,
	[TypeofProduct] [nvarchar](100) NULL,
	[Createdby] [nvarchar](100) NULL,
	[createdon] [datetime] NULL,
	[updatedby] [nvarchar](100) NULL,
	[updatedon] [datetime] NULL,
	[setupid] [tinyint] NULL,
	[Opening] [nvarchar](100) NULL,
	[Auto] [int] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[College]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[College](
	[CollegeID] [int] IDENTITY(1,1) NOT NULL,
	[UniversityID] [int] NULL,
	[CityID] [int] NULL,
	[College] [nvarchar](50) NULL,
	[Add1] [nvarchar](50) NULL,
	[city] [nvarchar](100) NULL,
	[district] [nvarchar](35) NULL,
	[Pin] [nvarchar](10) NULL,
	[Phone1] [nvarchar](20) NULL,
	[Fax] [nvarchar](20) NULL,
	[Email] [nvarchar](25) NULL,
	[WebSite] [nvarchar](35) NULL,
	[Principalname] [nvarchar](50) NULL,
	[pphone] [nvarchar](20) NULL,
	[states] [nvarchar](20) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CNF1B_old]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CNF1B_old](
	[CNN] [float] NULL,
	[CND] [smalldatetime] NULL,
	[GLD] [nvarchar](40) NULL,
	[SLD] [nvarchar](100) NULL,
	[A] [float] NULL,
	[DC] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CNF1B]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CNF1B](
	[CNN] [float] NULL,
	[CND] [datetime] NULL,
	[GLD] [nvarchar](40) NULL,
	[SLD] [nvarchar](40) NULL,
	[A] [float] NULL,
	[DC] [nvarchar](1) NULL,
	[groupcode] [nvarchar](10) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CNF1A_old]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CNF1A_old](
	[CNN] [float] NOT NULL,
	[CND] [smalldatetime] NULL,
	[PGLD] [nvarchar](40) NULL,
	[PSLD] [nvarchar](100) NULL,
	[NA] [float] NULL,
	[N] [nvarchar](150) NULL,
	[DC] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL,
	[bAuthorized] [bit] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CNF1A]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CNF1A](
	[CNN] [float] NULL,
	[CND] [datetime] NULL,
	[PGLD] [nvarchar](40) NULL,
	[PSLD] [nvarchar](40) NULL,
	[NA] [float] NULL,
	[N] [nvarchar](40) NULL,
	[DC] [nvarchar](1) NULL,
	[Agentname] [nvarchar](30) NULL,
	[BAuthorized] [bit] NULL,
	[print_yes] [nvarchar](50) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[City_tobedel]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[City_tobedel](
	[DistrictID] [nvarchar](6) NULL,
	[CityID] [nvarchar](6) NULL,
	[City] [nvarchar](50) NULL,
	[CityKey] [nvarchar](200) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[City]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[City](
	[CityID] [int] IDENTITY(1,1) NOT NULL,
	[City] [nvarchar](35) NULL,
	[District] [nvarchar](35) NULL,
	[State] [nvarchar](35) NULL,
	[Country] [nvarchar](35) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CBMF]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CBMF](
	[GLD] [nvarchar](40) NULL,
	[SLD] [nvarchar](100) NULL,
	[Fyear] [nvarchar](10) NULL,
	[Createdby] [nvarchar](50) NULL,
	[createdon] [smalldatetime] NULL,
	[updatedby] [nvarchar](50) NULL,
	[updatedon] [smalldatetime] NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BinderBkReceive]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BinderBkReceive](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[remarks] [nvarchar](50) NULL,
	[Add1] [nvarchar](50) NULL,
	[add2] [nvarchar](50) NULL,
	[ToPay] [float] NULL,
	[Godown] [nvarchar](20) NULL,
	[NetBook] [int] NULL,
	[FirmName] [nvarchar](50) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BILTYRETURNREGISTER]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BILTYRETURNREGISTER](
	[SNO] [int] NULL,
	[Date] [datetime] NULL,
	[Particulars] [nvarchar](240) NULL,
	[RR] [nvarchar](10) NULL,
	[GR] [int] NULL,
	[GR_DT] [datetime] NULL,
	[BDL] [nvarchar](100) NULL,
	[WT] [nvarchar](50) NULL,
	[Freight] [nvarchar](50) NULL,
	[Recd_dt] [datetime] NULL,
	[Freight_Paid] [nvarchar](50) NULL,
	[CNO] [int] NULL,
	[RNO] [int] NULL,
	[Remarks] [nvarchar](254) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[billtrans]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[billtrans](
	[bill_id] [nvarchar](10) NULL,
	[firm_id] [nvarchar](10) NULL,
	[particulars] [nvarchar](100) NULL,
	[PAPERsize1] [nvarchar](10) NULL,
	[unit] [float] NULL,
	[type] [nvarchar](10) NULL,
	[quantity] [int] NULL,
	[amt] [int] NULL,
	[reams] [int] NULL,
	[sheets] [int] NULL,
	[wastage] [int] NULL,
	[wreams] [int] NULL,
	[wsheets] [int] NULL,
	[partyid] [nvarchar](10) NULL,
	[bookinfo] [nvarchar](80) NULL,
	[rate] [money] NULL,
	[atrate] [money] NULL,
	[plate] [int] NULL,
	[patrate] [money] NULL,
	[pamt] [int] NULL,
	[fontbk] [nvarchar](1) NULL,
	[billdate] [datetime] NULL,
	[ft] [nvarchar](50) NULL,
	[recdel] [nvarchar](1) NULL,
	[totaldelreams] [float] NULL,
	[totaldelsheets] [nvarchar](50) NULL,
	[totalrecreams] [float] NULL,
	[totalrecsheets] [float] NULL,
	[balancereams] [float] NULL,
	[balancesheets] [float] NULL,
	[totalrim] [float] NULL,
	[totalsht] [float] NULL,
	[OPENINGreams] [float] NULL,
	[OPENINGsheets] [float] NULL,
	[pscustomerid] [nvarchar](10) NULL,
	[pstatementno] [int] NULL,
	[pstdate] [datetime] NULL,
	[previousebillid] [nvarchar](10) NULL,
	[previousebilldate] [datetime] NULL,
	[Head1] [nvarchar](50) NULL,
	[Head2] [nvarchar](50) NULL,
	[Head3] [nvarchar](50) NULL,
	[Head4] [nvarchar](50) NULL,
	[Head5] [nvarchar](50) NULL,
	[Note] [nvarchar](255) NULL,
	[QtyRec] [nvarchar](20) NULL,
	[supp] [nvarchar](40) NULL,
	[PaparMake] [nvarchar](100) NULL,
	[BookNo] [nvarchar](100) NULL,
	[Inner] [nvarchar](10) NULL,
	[Inn_Printer] [nvarchar](50) NULL,
	[text_Printer] [nvarchar](50) NULL,
	[Exam_Printer] [nvarchar](50) NULL,
	[Supp_Printer] [nvarchar](50) NULL,
	[Title_Printer] [nvarchar](50) NULL,
	[PHead1] [nvarchar](50) NULL,
	[PHead2] [nvarchar](50) NULL,
	[PHead3] [nvarchar](50) NULL,
	[PHead4] [nvarchar](50) NULL,
	[PHead5] [nvarchar](50) NULL,
	[Binder] [nvarchar](50) NULL,
	[BookName] [nvarchar](50) NULL,
	[BookCode] [nvarchar](20) NULL,
	[PaperName] [nvarchar](50) NULL,
	[Set_Name] [nvarchar](50) NULL,
	[Qty] [int] NULL,
	[OrderPrinting] [int] NULL,
	[GPOrderPrinting] [int] NULL,
	[RemDetails] [nvarchar](100) NULL,
	[PaperDetails] [nvarchar](100) NULL,
	[Inners] [nvarchar](20) NULL,
	[text] [nvarchar](20) NULL,
	[Paper] [nvarchar](20) NULL,
	[NoOfForm] [nvarchar](20) NULL,
	[Types] [nvarchar](20) NULL,
	[Neg_Remarks] [nvarchar](100) NULL,
	[categories] [nvarchar](20) NULL,
	[binderName1] [nvarchar](100) NULL,
	[bookfont] [nvarchar](1) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[billmaster]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[billmaster](
	[bill_id] [nvarchar](50) NULL,
	[firm_id] [nvarchar](10) NULL,
	[party_id] [nvarchar](10) NULL,
	[dat] [datetime] NULL,
	[order_no] [smallint] NULL,
	[binder_id] [nvarchar](50) NULL,
	[total] [int] NULL,
	[totalrim] [float] NULL,
	[totalsht] [float] NULL,
	[logdetail] [datetime] NULL,
	[papersize1] [nvarchar](10) NULL,
	[totalprint] [float] NULL,
	[totalplate] [float] NULL,
	[totaldelreams] [float] NULL,
	[totaldelsheets] [nvarchar](50) NULL,
	[totalrecreams] [float] NULL,
	[totalrecsheets] [float] NULL,
	[balancereams] [float] NULL,
	[balancesheets] [float] NULL,
	[OPENINGreams] [float] NULL,
	[OPENINGsheets] [float] NULL,
	[pscustomerid] [nvarchar](10) NULL,
	[pstatementno] [int] NULL,
	[previousebillid] [nvarchar](10) NULL,
	[previousebilldate] [datetime] NULL,
	[calcmanual] [bit] NULL,
	[totalwrim] [float] NULL,
	[totalwsht] [float] NULL,
	[totalnoplate] [int] NULL,
	[totalnoream] [float] NULL,
	[totalnosht] [float] NULL,
	[OrderCancel] [nvarchar](50) NULL,
	[categories] [nvarchar](20) NULL,
	[PrinterName] [nvarchar](100) NULL,
	[Remarks] [nvarchar](100) NULL,
	[binderName] [nvarchar](100) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[bankstm]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[bankstm](
	[sno] [int] NULL,
	[ddNo] [nvarchar](100) NULL,
	[nameOfTheBank] [nvarchar](50) NULL,
	[PartyName] [nvarchar](50) NULL,
	[amount] [float] NULL,
	[dated] [datetime] NULL,
	[autoNo] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CREDITEND]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CREDITEND](
	[CGENLEDGER] [nvarchar](max) NOT NULL,
	[CSUBLEDGER] [nvarchar](max) NOT NULL,
	[GENLEDGER] [nvarchar](max) NOT NULL,
	[SUBLEDGER] [nvarchar](max) NOT NULL,
	[TEXT] [nvarchar](max) NOT NULL,
	[RATE] [float] NULL,
	[DEBITORCREDIT] [nvarchar](max) NULL,
	[PrintOrder] [smallint] NULL,
	[RYN] [nvarchar](max) NULL,
	[id] [int] NULL,
	[fyear] [nvarchar](max) NOT NULL,
	[Createdby] [nvarchar](max) NULL,
	[createdon] [datetime] NULL,
	[updatedby] [nvarchar](max) NULL,
	[updatedon] [datetime] NULL,
	[setupid] [tinyint] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CREDITCtmp]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CREDITCtmp](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CREDITC]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CREDITC](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CashRegister_conven]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CashRegister_conven](
	[SNO] [int] NULL,
	[Date] [datetime] NULL,
	[Particulars] [nvarchar](240) NULL,
	[RR] [nvarchar](10) NULL,
	[GR] [int] NULL,
	[GR_DT] [datetime] NULL,
	[BDL] [nvarchar](100) NULL,
	[CMNo] [int] NULL,
	[WT] [nvarchar](50) NULL,
	[Freight] [nvarchar](50) NULL,
	[Recd_dt] [datetime] NULL,
	[Freight_Paid] [nvarchar](50) NULL,
	[CNO] [int] NULL,
	[RNO] [int] NULL,
	[Remarks] [nvarchar](254) NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [nvarchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CashRegister]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CashRegister](
	[SNO] [int] NULL,
	[Date] [datetime] NULL,
	[Particulars] [nvarchar](240) NULL,
	[RR] [nvarchar](10) NULL,
	[GR] [int] NULL,
	[GR_DT] [datetime] NULL,
	[BDL] [nvarchar](100) NULL,
	[CMNo] [int] NULL,
	[WT] [nvarchar](50) NULL,
	[Freight] [nvarchar](50) NULL,
	[Recd_dt] [datetime] NULL,
	[Freight_Paid] [nvarchar](50) NULL,
	[CNO] [int] NULL,
	[RNO] [int] NULL,
	[Remarks] [nvarchar](254) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHCTMP_basilRet]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHCTMP_basilRet](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHCTMP_basil]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHCTMP_basil](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHCtmp]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHCtmp](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHC_basilRet]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHC_basilRet](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHC_basil]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHC_basil](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHC]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHC](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CashAndSalesrpt]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CashAndSalesrpt](
	[Bill] [int] NULL,
	[BillDate] [datetime] NULL,
	[Net] [float] NULL,
	[BookName] [nvarchar](150) NULL,
	[Qty] [int] NULL,
	[Mark] [nvarchar](10) NULL,
	[Cash_Sales] [nvarchar](10) NULL,
	[GP] [nvarchar](10) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHA_basilRet]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHA_basilRet](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[ORDERBY] [nvarchar](15) NULL,
	[ORDERDATE] [datetime] NULL,
	[MARKA] [nvarchar](10) NULL,
	[BUNDLES] [nvarchar](15) NULL,
	[THROUGH] [nvarchar](50) NULL,
	[STATION] [nvarchar](50) NULL,
	[BILTYNO] [nvarchar](10) NULL,
	[BILTYDATE] [datetime] NULL,
	[FREIGHT] [nvarchar](15) NULL,
	[WEIGHT] [nvarchar](10) NULL,
	[TXT1] [nvarchar](20) NULL,
	[TXT1A] [float] NULL,
	[TXT2] [nvarchar](20) NULL,
	[TXT2A] [float] NULL,
	[BAA] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[GAMOUNT] [float] NULL,
	[T2] [nvarchar](40) NULL,
	[THROUGH1] [nvarchar](50) NULL,
	[district] [nvarchar](30) NULL,
	[cashpartyname] [nvarchar](40) NULL,
	[AgentName] [nvarchar](30) NULL,
	[DISCAT] [nvarchar](1) NULL,
	[transportname] [nvarchar](30) NULL,
	[BAuthorized] [bit] NULL,
	[print_yes] [nvarchar](50) NULL,
	[Sale_Return] [nvarchar](15) NULL,
	[Godown] [nvarchar](10) NULL,
	[mark1] [nvarchar](10) NULL,
	[DISCAT2] [nvarchar](1) NULL,
	[DISCAT3] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHA_basil]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHA_basil](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[ORDERBY] [nvarchar](15) NULL,
	[ORDERDATE] [datetime] NULL,
	[MARKA] [nvarchar](10) NULL,
	[BUNDLES] [nvarchar](15) NULL,
	[THROUGH] [nvarchar](50) NULL,
	[STATION] [nvarchar](50) NULL,
	[BILTYNO] [nvarchar](10) NULL,
	[BILTYDATE] [datetime] NULL,
	[FREIGHT] [nvarchar](15) NULL,
	[WEIGHT] [nvarchar](10) NULL,
	[TXT1] [nvarchar](20) NULL,
	[TXT1A] [float] NULL,
	[TXT2] [nvarchar](20) NULL,
	[TXT2A] [float] NULL,
	[BAA] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[GAMOUNT] [float] NULL,
	[T2] [nvarchar](40) NULL,
	[THROUGH1] [nvarchar](50) NULL,
	[district] [nvarchar](30) NULL,
	[cashpartyname] [nvarchar](40) NULL,
	[AgentName] [nvarchar](30) NULL,
	[DISCAT] [nvarchar](1) NULL,
	[transportname] [nvarchar](30) NULL,
	[BAuthorized] [bit] NULL,
	[print_yes] [nvarchar](50) NULL,
	[Sale_Return] [nvarchar](15) NULL,
	[Godown] [nvarchar](10) NULL,
	[mark1] [nvarchar](10) NULL,
	[DISCAT2] [nvarchar](1) NULL,
	[DISCAT3] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHA]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHA](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[ORDERBY] [nvarchar](15) NULL,
	[ORDERDATE] [datetime] NULL,
	[MARKA] [nvarchar](10) NULL,
	[BUNDLES] [nvarchar](15) NULL,
	[THROUGH] [nvarchar](50) NULL,
	[STATION] [nvarchar](50) NULL,
	[BILTYNO] [nvarchar](10) NULL,
	[BILTYDATE] [datetime] NULL,
	[FREIGHT] [nvarchar](15) NULL,
	[WEIGHT] [nvarchar](10) NULL,
	[TXT1] [nvarchar](20) NULL,
	[TXT1A] [float] NULL,
	[TXT2] [nvarchar](20) NULL,
	[TXT2A] [float] NULL,
	[BAA] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[GAMOUNT] [float] NULL,
	[T2] [nvarchar](40) NULL,
	[THROUGH1] [nvarchar](50) NULL,
	[district] [nvarchar](30) NULL,
	[cashpartyname] [nvarchar](40) NULL,
	[AgentName] [nvarchar](30) NULL,
	[DISCAT] [nvarchar](1) NULL,
	[transportname] [nvarchar](30) NULL,
	[BAuthorized] [bit] NULL,
	[print_yes] [nvarchar](50) NULL,
	[Godown] [nvarchar](10) NULL,
	[DISCATII] [nvarchar](1) NULL,
	[DISCATIII] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BookStock]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BookStock](
	[EntryNo] [int] NULL,
	[Dates] [datetime] NULL,
	[BOOKCODE] [nvarchar](10) NULL,
	[Qty] [int] NULL,
	[Category] [nvarchar](20) NULL,
	[Issue_Receive] [nvarchar](15) NULL,
	[Binder_Code] [nvarchar](50) NULL,
	[Godown_Out] [nvarchar](30) NULL,
	[Godown_In] [nvarchar](30) NULL,
	[GodownHead] [nvarchar](15) NULL,
	[Remarks] [nvarchar](200) NULL,
	[auto] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BOOKS]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BOOKS](
	[BOOKCODE] [nvarchar](10) NOT NULL,
	[BOOKNAME] [nvarchar](40) NULL,
	[GROUPCODE] [nvarchar](7) NULL,
	[RATE] [float] NULL,
	[DISCOUNT] [float] NULL,
	[RetailPrice] [float] NULL,
	[RetailDis] [float] NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[BookDes] [nvarchar](100) NULL,
 CONSTRAINT [PK_BOOKS_1] PRIMARY KEY CLUSTERED 
(
	[BOOKCODE] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BookReceiveDet]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BookReceiveDet](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[BOOKCODE] [nvarchar](50) NULL,
	[TBook] [int] NULL,
	[LoosBook] [int] NULL,
	[TotalBook] [int] NULL,
	[NetBook] [int] NULL,
	[Book_Code] [nvarchar](20) NULL,
	[Remarks] [nvarchar](50) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BookMaster]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BookMaster](
	[book] [nvarchar](100) NULL,
	[class] [nvarchar](15) NULL,
	[book_info] [nvarchar](80) NULL,
	[bookfont] [nvarchar](1) NULL,
	[book_size] [nvarchar](10) NULL,
	[book_unit] [float] NULL,
	[atrate] [money] NULL,
	[wastage] [int] NULL,
	[patrate] [money] NULL,
	[websheet] [nvarchar](10) NULL,
	[party_id] [nvarchar](10) NULL,
	[Head1] [nvarchar](50) NULL,
	[Head2] [nvarchar](50) NULL,
	[Head3] [nvarchar](50) NULL,
	[Head4] [nvarchar](50) NULL,
	[Head5] [nvarchar](50) NULL,
	[HeadData1] [nvarchar](50) NULL,
	[HeadData2] [nvarchar](50) NULL,
	[HeadData3] [nvarchar](50) NULL,
	[HeadData4] [nvarchar](50) NULL,
	[HeadData5] [nvarchar](50) NULL,
	[DivideValue] [int] NULL,
	[BookNo] [nvarchar](10) NOT NULL,
	[Writer] [nvarchar](100) NULL,
	[TypeSetter] [nvarchar](100) NULL,
	[NegativeBy] [nvarchar](100) NULL,
	[Brand] [nvarchar](100) NULL,
	[Quality] [nvarchar](50) NULL,
	[Color] [nvarchar](50) NULL,
	[Lamination] [nvarchar](50) NULL,
	[Binder] [nvarchar](50) NULL,
	[Inn_Printer] [nvarchar](50) NULL,
	[text_Printer] [nvarchar](50) NULL,
	[Exam_Printer] [nvarchar](50) NULL,
	[Supp_Printer] [nvarchar](50) NULL,
	[Title_Printer] [nvarchar](50) NULL,
	[Inn_color] [nvarchar](50) NULL,
	[text_color] [nvarchar](50) NULL,
	[Exam_color] [nvarchar](50) NULL,
	[Supp_color] [nvarchar](50) NULL,
	[Title_color] [nvarchar](50) NULL,
	[Inn_pcode] [nvarchar](15) NULL,
	[text_pcode] [nvarchar](15) NULL,
	[Exam_pcode] [nvarchar](15) NULL,
	[Supp_pcode] [nvarchar](15) NULL,
	[Title_pcode] [nvarchar](15) NULL,
	[Edition] [nvarchar](100) NULL,
	[Inn_DBy] [int] NULL,
	[text_DBy] [int] NULL,
	[BrightNess] [nvarchar](50) NULL,
	[Exam_DBy] [int] NULL,
	[Supp_DBy] [int] NULL,
	[Title_DBy] [int] NULL,
	[Inn_Forms] [float] NULL,
	[text_Forms] [float] NULL,
	[Exam_Forms] [float] NULL,
	[Supp_Forms] [float] NULL,
	[Title_Forms] [float] NULL,
	[Inn_Bright] [nvarchar](15) NULL,
	[text_Bright] [nvarchar](15) NULL,
	[Exam_Bright] [nvarchar](15) NULL,
	[Supp_Bright] [nvarchar](15) NULL,
	[Title_Bright] [nvarchar](15) NULL,
	[price] [float] NULL,
	[GpNo] [float] NULL,
	[InternalPrint] [float] NULL,
	[txtHead6] [nvarchar](50) NULL,
	[txtHeadData6] [float] NULL,
	[cbosupp6] [nvarchar](50) NULL,
	[txtTextSupp6] [float] NULL,
	[cboPrinter6] [nvarchar](100) NULL,
	[cboColour6] [nvarchar](30) NULL,
	[txtPCode6] [nvarchar](50) NULL,
	[txtHead7] [nvarchar](50) NULL,
	[txtHeadData7] [float] NULL,
	[cbosupp7] [nvarchar](50) NULL,
	[txtTextSupp7] [float] NULL,
	[cboPrinter7] [nvarchar](100) NULL,
	[cboColour7] [nvarchar](30) NULL,
	[txtPCode7] [nvarchar](50) NULL,
	[txtHead8] [nvarchar](50) NULL,
	[txtHeadData8] [float] NULL,
	[cbosupp8] [nvarchar](50) NULL,
	[txtTextSupp8] [float] NULL,
	[cboPrinter8] [nvarchar](100) NULL,
	[cboColour8] [nvarchar](30) NULL,
	[txtPCode8] [nvarchar](50) NULL,
	[Return_CY] [float] NULL,
	[Sepimen_CY] [float] NULL,
	[Return_LY] [float] NULL,
	[Sepimen_LY] [float] NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BookGp]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BookGp](
	[Id] [int] NULL,
	[GroupCode] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BookDetails]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BookDetails](
	[BookNo] [nvarchar](50) NULL,
	[fyear] [nvarchar](15) NULL,
	[Price] [float] NULL,
	[TotalPrinted] [int] NULL,
	[Specimen] [int] NULL,
	[id] [int] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[bookcategory]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[bookcategory](
	[bookcategory] [nvarchar](50) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Book]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Book](
	[Bkname] [nvarchar](80) NULL,
	[AccessionNo] [nvarchar](30) NULL,
	[bkcode] [nvarchar](15) NULL,
	[bkauthorid] [int] NULL,
	[bkpublisherid] [int] NULL,
	[bksubject] [nvarchar](30) NULL,
	[bkedition] [nvarchar](10) NULL,
	[bkprice] [money] NULL,
	[bkrecdfrom] [nvarchar](30) NULL,
	[bkremark] [nvarchar](100) NULL,
	[bkisbn] [nvarchar](20) NULL,
	[bkrecddate] [datetime] NULL,
	[clcode] [smallint] NULL,
	[PageNo] [float] NULL,
	[ShelfNo] [nvarchar](10) NULL,
	[BKYEAR] [nvarchar](10) NULL,
	[BKVOLUME] [nvarchar](10) NULL,
	[BKBILLNO] [nvarchar](10) NULL,
	[BKBILLDATE] [datetime] NULL,
	[BKCALLCLASSNO] [nvarchar](10) NULL,
	[BKCALLBOOKNO] [nvarchar](10) NULL,
	[BKWITHNO] [nvarchar](10) NULL,
	[BKWITHDATE] [datetime] NULL,
	[Catid] [int] NULL,
	[Status] [bit] NULL,
	[bookbank] [bit] NULL,
	[classnumber] [nvarchar](255) NULL,
	[booknumber] [nvarchar](255) NULL
) ON [PRIMARY]
GO
/****** Object:  View [dbo].[BinderReceiveRegister]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[BinderReceiveRegister]
AS
SELECT     dbo.BinderBkReceive.INVOICEDATE, dbo.BookReceiveDet.BOOKCODE, SUM(dbo.BookReceiveDet.TotalBook) AS Qty, 'Issue' AS Issue, 
                      dbo.BinderBkReceive.Godown, dbo.BinderBkReceive.Fyear, dbo.BinderBkReceive.setupid, dbo.BinderBkReceive.SUBLEDGER
FROM         dbo.BinderBkReceive INNER JOIN
                      dbo.BookReceiveDet ON dbo.BinderBkReceive.INVOICENO = dbo.BookReceiveDet.INVOICENO AND 
                      dbo.BinderBkReceive.Fyear = dbo.BookReceiveDet.Fyear AND dbo.BinderBkReceive.setupid = dbo.BookReceiveDet.setupid
GROUP BY dbo.BinderBkReceive.INVOICEDATE, dbo.BookReceiveDet.BOOKCODE, dbo.BinderBkReceive.Godown, dbo.BinderBkReceive.Fyear, 
                      dbo.BinderBkReceive.setupid, dbo.BinderBkReceive.SUBLEDGER
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[41] 4[20] 2[12] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "BinderBkReceive"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 206
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "BookReceiveDet"
            Begin Extent = 
               Top = 6
               Left = 244
               Bottom = 206
               Right = 412
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'BinderReceiveRegister'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'BinderReceiveRegister'
GO
/****** Object:  Table [dbo].[CASHB_basilRet]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHB_basilRet](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[BOOKCODE] [nvarchar](10) NOT NULL,
	[QUANTITY] [float] NULL,
	[RATE] [float] NULL,
	[DISCOUNT] [float] NULL,
	[PRINTORDER] [float] NULL,
	[AMOUNT] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[sno] [int] IDENTITY(1,1) NOT NULL,
	[agentname] [nvarchar](30) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHB_basil]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHB_basil](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[BOOKCODE] [nvarchar](10) NOT NULL,
	[QUANTITY] [float] NULL,
	[RATE] [float] NULL,
	[DISCOUNT] [float] NULL,
	[PRINTORDER] [float] NULL,
	[AMOUNT] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[sno] [int] IDENTITY(1,1) NOT NULL,
	[agentname] [nvarchar](30) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CASHB]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CASHB](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[BOOKCODE] [nvarchar](10) NOT NULL,
	[QUANTITY] [float] NULL,
	[RATE] [float] NULL,
	[DISCOUNT] [float] NULL,
	[PRINTORDER] [float] NULL,
	[AMOUNT] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[sno] [int] IDENTITY(1,1) NOT NULL,
	[agentname] [nvarchar](30) NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL,
	[NoOfBox] [nvarchar](20) NULL,
	[Per_Box_GW] [nvarchar](20) NULL,
	[Per_Box_NW] [nvarchar](20) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[CREDITB]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CREDITB](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[BOOKCODE] [nvarchar](10) NOT NULL,
	[QUANTITY] [float] NULL,
	[RATE] [float] NULL,
	[DISCOUNT] [float] NULL,
	[PRINTORDER] [float] NULL,
	[AMOUNT] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[sno] [int] IDENTITY(1,1) NOT NULL,
	[fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  StoredProcedure [dbo].[Export_Invoicea]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[Export_Invoicea]

@INVOICENO int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@SUBLEDGER nvarchar (50),
@TXT1A nvarchar (50),
@TXT2 nvarchar (50),
@TXT2A nvarchar (50),
@T2 nvarchar (40),
@ORDERNO nvarchar(50),
@BUYERBATCHNO nvarchar (30),
@BRAND nvarchar (50),
@TERMSDELIVERY nvarchar (50),
@TERMSPAYMENT nvarchar (50),
@COUNTRYDEST nvarchar (25),
@PRECARRIAGE nvarchar (25),
@PLACEOFRECEIPT nvarchar (25),
@PORTLOADING nvarchar (25),
@PORTDISCHARGE nvarchar (25),
@TOTALPSC nvarchar (25),
@TOTALCARTONS nvarchar (25),
@TOTALCBM nvarchar (25),
@NETWEIGHT nvarchar (25),
@GWEIGHT nvarchar (25),
@NetAmount  float,
@NetAmount2 float,
@typeinvoce nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int


as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        


insert into Invoicea(INVOICENO,INVOICEDATE,IssueDate,DisPatchDate,SUBLEDGER,TXT1A,TXT2,TXT2A,T2,ORDERNO,ADVICEREMARK,BRAND,
TERMSDELEVERY,TERMSPAYMENT,COUNTRYDEST,PRECARRIAGE,PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,
TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT,NetAmount,GAMOUNT,typeofinvoice,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@INVOICENO,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@SUBLEDGER,
@TXT1A,@TXT2,@TXT2A,@T2,@ORDERNO,@BUYERBATCHNO,@BRAND,@TERMSDELIVERY,@TERMSPAYMENT,@COUNTRYDEST,@PRECARRIAGE,
@PLACEOFRECEIPT,@PORTLOADING,@PORTDISCHARGE,@TOTALPSC,@TOTALCARTONS,@TOTALCBM,@NETWEIGHT,@GWEIGHT,
@NetAmount,@NetAmount2,@typeinvoce,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)





  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  StoredProcedure [dbo].[Export_Casha1]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[Export_Casha1]

@INVOICENO int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@SUBLEDGER nvarchar (50),
@TXT1A nvarchar (50),
@TXT2 nvarchar (50),
@TXT2A nvarchar (50),
@T2 nvarchar (40),
@ORDERNO nvarchar(50),
@BUYERBATCHNO nvarchar (30),
@BRAND nvarchar (50),
@TERMSDELIVERY nvarchar (50),
@TERMSPAYMENT nvarchar (50),
@COUNTRYDEST nvarchar (25),
@PRECARRIAGE nvarchar (25),
@PLACEOFRECEIPT nvarchar (25),
@PORTLOADING nvarchar (25),
@PORTDISCHARGE nvarchar (25),
@TOTALPSC nvarchar (25),
@TOTALCARTONS nvarchar (25),
@TOTALCBM nvarchar (25),
@NETWEIGHT nvarchar (25),
@GWEIGHT nvarchar (25),
@NetAmount  float,
@NetAmount2 float,
@CurrencyValue nvarchar(10),
@typeinvoce nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int


as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        


insert into Casha(INVOICENO,INVOICEDATE,IssueDate,DisPatchDate,SUBLEDGER,TXT1A,TXT2,TXT2A,T2,ORDERNO,ADVICEREMARK,BRAND,
TERMSDELEVERY,TERMSPAYMENT,COUNTRYDEST,PRECARRIAGE,PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,
TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT,NetAmount,GAMOUNT,typeofinvoice,CurrencyValue,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@INVOICENO,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@SUBLEDGER,
@TXT1A,@TXT2,@TXT2A,@T2,@ORDERNO,@BUYERBATCHNO,@BRAND,@TERMSDELIVERY,@TERMSPAYMENT,@COUNTRYDEST,@PRECARRIAGE,
@PLACEOFRECEIPT,@PORTLOADING,@PORTDISCHARGE,@TOTALPSC,@TOTALCARTONS,@TOTALCBM,@NETWEIGHT,@GWEIGHT,
@NetAmount,@NetAmount2,@typeinvoce,@CurrencyValue,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)





  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  StoredProcedure [dbo].[Export_Casha]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[Export_Casha]

@INVOICENO int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@GENLEDGER nvarchar (50),
@SUBLEDGER nvarchar (50),
@TXT1A nvarchar (50),
@TXT2 nvarchar (50),
@TXT2A nvarchar (50),
@T2 nvarchar (40),
@ORDERNO nvarchar(50),
@BUYERBATCHNO nvarchar (30),
@BRAND nvarchar (50),
@TERMSDELIVERY nvarchar (50),
@TERMSPAYMENT nvarchar (50),
@COUNTRYDEST nvarchar (25),
@PRECARRIAGE nvarchar (25),
@PLACEOFRECEIPT nvarchar (25),
@PORTLOADING nvarchar (25),
@PORTDISCHARGE nvarchar (25),
@TOTALPSC nvarchar (25),
@TOTALCARTONS nvarchar (25),
@TOTALCBM nvarchar (25),
@NETWEIGHT nvarchar (25),
@GWEIGHT nvarchar (25),
@NetAmount  float,
@NetAmount2 float,
@CurrencyValue nvarchar(10),
@NetRate  float,
@typeinvoce nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int


as
BEGIN        
  
 SET NOCOUNT ON        
 SET ANSI_WARNINGS OFF        


insert into Casha(INVOICENO,INVOICEDATE,IssueDate,DisPatchDate,GENLEDGER,SUBLEDGER,TXT1A,TXT2,TXT2A,T2,ORDERNO,ADVICEREMARK,BRAND,
TERMSDELEVERY,TERMSPAYMENT,COUNTRYDEST,PRECARRIAGE,PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,
TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT,NetAmount,GAMOUNT,typeofinvoice,CurrencyValue,NetRate,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@INVOICENO,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@GENLEDGER,@SUBLEDGER,
@TXT1A,@TXT2,@TXT2A,@T2,@ORDERNO,@BUYERBATCHNO,@BRAND,@TERMSDELIVERY,@TERMSPAYMENT,@COUNTRYDEST,@PRECARRIAGE,
@PLACEOFRECEIPT,@PORTLOADING,@PORTDISCHARGE,@TOTALPSC,@TOTALCARTONS,@TOTALCBM,@NETWEIGHT,@GWEIGHT,
@NetAmount,@NetAmount2,@typeinvoce,@CurrencyValue,@NetRate,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)




SET ANSI_WARNINGS ON        
SET NOCOUNT OFF        

END
GO
/****** Object:  StoredProcedure [dbo].[insertData_subledger]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[insertData_subledger]

@gldger nvarchar (50),
@subledger nvarchar (50),
@address1 nvarchar (50),
@address2 nvarchar (50),
@phone nvarchar (25),
@FAX nvarchar (25),
@YEAROPENING float,
@bCityId nvarchar (10),
@radd1 nvarchar (50),
@radd2 nvarchar (50),
@rcityid nvarchar (10),
@rphone nvarchar (25),
@rfax nvarchar (25),

@contmail nvarchar (35),
@tpt nvarchar (100),
@range nvarchar (200),
@bank nvarchar (50),
@sadd1 nvarchar (50),
@slremark nvarchar (200),
@contperson nvarchar (40),
@contphone nvarchar (25),
@tinno nvarchar (50),
@TypeOfCust nvarchar(25),
@FYEAR nvarchar(10),
@setupid int,
@DESCFORINVOICE nvarchar (50)


as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        

  insert into SLEDGER(gledger,subledger,address1,address2,Phone,fax,YEAROPENING,bcityid,radd1,radd2,rcityid,rphone,rfax,
	email,tpt,nRange,bank,sadd1,slremark,contperson,contphone,tinno,TypeOfCust,FYEAR,setupid,DESCFORINVOICE )
  values(@gldger,@subledger,@address1,@address2,@Phone,@fax,@YEAROPENING,@bCityId,@radd1,@radd2,@rcityid,@rphone,@rfax, 
	@contmail,@tpt,@range,@bank,@sadd1,@slremark,@contperson,@contphone,@tinno,@TypeOfCust,@FYEAR,@setupid,@DESCFORINVOICE )

  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  Table [dbo].[INVOICEC]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEC](
	[INVOICENO] [int] NOT NULL,
	[INVOICEDATE] [datetime] NOT NULL,
	[GENLEDGER] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[GAMOUNT] [float] NULL,
	[Rate] [float] NULL,
	[AMOUNT] [float] NULL,
	[DEBITORCREDIT] [nvarchar](7) NULL,
	[TEXT] [nvarchar](20) NULL,
	[RYN] [nvarchar](1) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL,
	[autoid] [int] IDENTITY(1,1) NOT NULL,
	[saleType] [nvarchar](100) NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEB_spRet]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEB_spRet](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[BOOKCODE] [nvarchar](10) NOT NULL,
	[QUANTITY] [float] NULL,
	[RATE] [float] NULL,
	[DISCOUNT] [float] NULL,
	[PRINTORDER] [float] NULL,
	[AMOUNT] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[SNO] [int] IDENTITY(1,1) NOT NULL,
	[agentname] [nvarchar](30) NULL,
	[groupname] [nvarchar](20) NULL,
	[group1] [float] NULL,
	[group2] [float] NULL,
	[group3] [float] NULL,
	[group4] [float] NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEB_sp]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEB_sp](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[BOOKCODE] [nvarchar](10) NOT NULL,
	[QUANTITY] [float] NULL,
	[RATE] [float] NULL,
	[DISCOUNT] [float] NULL,
	[PRINTORDER] [float] NULL,
	[AMOUNT] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[SNO] [int] IDENTITY(1,1) NOT NULL,
	[agentname] [nvarchar](30) NULL,
	[groupname] [nvarchar](20) NULL,
	[group1] [float] NULL,
	[group2] [float] NULL,
	[group3] [float] NULL,
	[group4] [float] NULL,
	[Fyear] [nvarchar](10) NOT NULL,
	[setupid] [tinyint] NOT NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEB_IssuedBind]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEB_IssuedBind](
	[INVOICENO] [int] NULL,
	[INVOICEDATE] [datetime] NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[BOOKCODE] [nvarchar](100) NULL,
	[noofpack] [int] NULL,
	[TBook] [int] NULL,
	[LoosBook] [int] NULL,
	[TotalBook] [int] NULL,
	[NetBook] [int] NULL,
	[SNO] [int] IDENTITY(1,1) NOT NULL,
	[add1] [nvarchar](200) NULL,
	[add2] [nvarchar](200) NULL,
	[Crate] [float] NULL,
	[DisRate] [float] NULL,
	[Remarks] [nvarchar](50) NULL,
	[Book_Code] [nvarchar](10) NOT NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[INVOICEB]    Script Date: 09/05/2013 01:12:13 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[INVOICEB](
	[INVOICENO] [int] NOT NULL,
	[INVOICEDATE] [datetime] NOT NULL,
	[Genledger] [nvarchar](40) NULL,
	[SUBLEDGER] [nvarchar](40) NULL,
	[BOOKCODE] [nvarchar](10) NULL,
	[QUANTITY] [float] NULL,
	[RATE] [float] NULL,
	[DISCOUNT] [float] NULL,
	[PRINTORDER] [float] NULL,
	[AMOUNT] [float] NULL,
	[NETAMOUNT] [float] NULL,
	[SNO] [int] IDENTITY(1,1) NOT NULL,
	[agentname] [nvarchar](30) NULL,
	[Fyear] [nvarchar](10) NULL,
	[setupid] [tinyint] NULL
) ON [PRIMARY]
GO
/****** Object:  View [dbo].[DewaliAmt]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[DewaliAmt]
AS
SELECT     dbo.CASHC_basil.INVOICENO, dbo.CASHC_basil.AMOUNT, dbo.CASHC_basil.TEXT, dbo.CASHA_basil.AgentName
FROM         dbo.CASHC_basil INNER JOIN
                      dbo.CASHA_basil ON dbo.CASHC_basil.INVOICENO = dbo.CASHA_basil.INVOICENO
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "CASHC_basil"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 200
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "CASHA_basil"
            Begin Extent = 
               Top = 6
               Left = 238
               Bottom = 121
               Right = 393
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'DewaliAmt'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'DewaliAmt'
GO
/****** Object:  StoredProcedure [dbo].[Packing_Invoicea1]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[Packing_Invoicea1]

	@Packing int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@SUBLEDGER nvarchar (50),
@TXT1A nvarchar (50),
@TXT2 nvarchar (50),
@TXT2A nvarchar (50),
@T2 nvarchar (40),
@ORDERNO nvarchar(50),
@BUYERBATCHNO nvarchar (30),
@BRAND nvarchar (50),
@TERMSDELIVERY nvarchar (50),
@TERMSPAYMENT nvarchar (50),
@COUNTRYDEST nvarchar (25),
@PRECARRIAGE nvarchar (25),
@PLACEOFRECEIPT nvarchar (25),
@PORTLOADING nvarchar (25),
@PORTDISCHARGE nvarchar (25),
@TOTALPSC nvarchar (25),
@TOTALCARTONS nvarchar (25),
@TOTALCBM nvarchar (25),
@NETWEIGHT nvarchar (25),
@GWEIGHT nvarchar (25),
@typeinvoce nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int

AS
BEGIN
	
	SET NOCOUNT ON;

   insert into Invoicea(PACKINGNO,INVOICEDATE,IssueDate,DisPatchDate,SUBLEDGER,TXT1A,TXT2,TXT2A,T2,ORDERNO,ADVICEREMARK,BRAND,
TERMSDELEVERY,TERMSPAYMENT,COUNTRYDEST,PRECARRIAGE,PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,
TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT,typeofinvoice,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@Packing,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@SUBLEDGER,
@TXT1A,@TXT2,@TXT2A,@T2,@ORDERNO,@BUYERBATCHNO,@BRAND,@TERMSDELIVERY,@TERMSPAYMENT,@COUNTRYDEST,@PRECARRIAGE,
@PLACEOFRECEIPT,@PORTLOADING,@PORTDISCHARGE,@TOTALPSC,@TOTALCARTONS,@TOTALCBM,@NETWEIGHT,@GWEIGHT,
@typeinvoce,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)

SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF    
END
GO
/****** Object:  StoredProcedure [dbo].[Packing_Invoicea]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[Packing_Invoicea]

	@Packing int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@SUBLEDGER nvarchar (50),
@TXT1A nvarchar (50),
@TXT2 nvarchar (50),
@TXT2A nvarchar (50),
@T2 nvarchar (40),
@ORDERNO nvarchar(50),
@BUYERBATCHNO nvarchar (30),
@BRAND nvarchar (50),
@TERMSDELIVERY nvarchar (50),
@TERMSPAYMENT nvarchar (50),
@COUNTRYDEST nvarchar (25),
@PRECARRIAGE nvarchar (25),
@PLACEOFRECEIPT nvarchar (25),
@PORTLOADING nvarchar (25),
@PORTDISCHARGE nvarchar (25),
@TOTALPSC nvarchar (25),
@TOTALCARTONS nvarchar (25),
@TOTALCBM nvarchar (25),
@NETWEIGHT nvarchar (25),
@GWEIGHT nvarchar (25),
@typeinvoce nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int

AS
BEGIN
	
	SET NOCOUNT ON;

   insert into Invoicea(PACKINGNO,INVOICEDATE,IssueDate,DisPatchDate,SUBLEDGER,TXT1A,TXT2,TXT2A,T2,ORDERNO,ADVICEREMARK,BRAND,
TERMSDELEVERY,TERMSPAYMENT,COUNTRYDEST,PRECARRIAGE,PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,
TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT,typeofinvoice,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@Packing,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@SUBLEDGER,
@TXT1A,@TXT2,@TXT2A,@T2,@ORDERNO,@BUYERBATCHNO,@BRAND,@TERMSDELIVERY,@TERMSPAYMENT,@COUNTRYDEST,@PRECARRIAGE,
@PLACEOFRECEIPT,@PORTLOADING,@PORTDISCHARGE,@TOTALPSC,@TOTALCARTONS,@TOTALCBM,@NETWEIGHT,@GWEIGHT,
@typeinvoce,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)

SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF    
END
GO
/****** Object:  StoredProcedure [dbo].[Packing_Casha1]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[Packing_Casha1]

	@Packing int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@SUBLEDGER nvarchar (50),
@TXT1A nvarchar (50),
@TXT2 nvarchar (50),
@TXT2A nvarchar (50),
@T2 nvarchar (40),
@ORDERNO nvarchar(50),
@BUYERBATCHNO nvarchar (30),
@BRAND nvarchar (50),
@TERMSDELIVERY nvarchar (50),
@TERMSPAYMENT nvarchar (50),
@COUNTRYDEST nvarchar (25),
@PRECARRIAGE nvarchar (25),
@PLACEOFRECEIPT nvarchar (25),
@PORTLOADING nvarchar (25),
@PORTDISCHARGE nvarchar (25),
@TOTALPSC nvarchar (25),
@TOTALCARTONS nvarchar (25),
@TOTALCBM nvarchar (25),
@NETWEIGHT nvarchar (25),
@GWEIGHT nvarchar (25),
@typeinvoce nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int

AS
BEGIN
	
	SET NOCOUNT ON;

   insert into Casha(PACKINGNO,INVOICEDATE,IssueDate,DisPatchDate,SUBLEDGER,TXT1A,TXT2,TXT2A,T2,ORDERNO,ADVICEREMARK,BRAND,
TERMSDELEVERY,TERMSPAYMENT,COUNTRYDEST,PRECARRIAGE,PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,
TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT,typeofinvoice,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@Packing,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@SUBLEDGER,
@TXT1A,@TXT2,@TXT2A,@T2,@ORDERNO,@BUYERBATCHNO,@BRAND,@TERMSDELIVERY,@TERMSPAYMENT,@COUNTRYDEST,@PRECARRIAGE,
@PLACEOFRECEIPT,@PORTLOADING,@PORTDISCHARGE,@TOTALPSC,@TOTALCARTONS,@TOTALCBM,@NETWEIGHT,@GWEIGHT,
@typeinvoce,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)

SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF    
END
GO
/****** Object:  StoredProcedure [dbo].[Packing_Casha]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[Packing_Casha]

	@Packing int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@SUBLEDGER nvarchar (50),
@TXT1A nvarchar (50),
@TXT2 nvarchar (50),
@TXT2A nvarchar (50),
@T2 nvarchar (40),
@ORDERNO nvarchar(50),
@BUYERBATCHNO nvarchar (30),
@BRAND nvarchar (50),
@TERMSDELIVERY nvarchar (50),
@TERMSPAYMENT nvarchar (50),
@COUNTRYDEST nvarchar (25),
@PRECARRIAGE nvarchar (25),
@PLACEOFRECEIPT nvarchar (25),
@PORTLOADING nvarchar (25),
@PORTDISCHARGE nvarchar (25),
@TOTALPSC nvarchar (25),
@TOTALCARTONS nvarchar (25),
@TOTALCBM nvarchar (25),
@NETWEIGHT nvarchar (25),
@GWEIGHT nvarchar (25),
@typeinvoce nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int

AS
BEGIN
	
	SET NOCOUNT ON;

   insert into Casha(PACKINGNO,INVOICEDATE,IssueDate,DisPatchDate,SUBLEDGER,TXT1A,TXT2,TXT2A,T2,ORDERNO,ADVICEREMARK,BRAND,
TERMSDELEVERY,TERMSPAYMENT,COUNTRYDEST,PRECARRIAGE,PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,
TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT,typeofinvoice,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@Packing,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@SUBLEDGER,
@TXT1A,@TXT2,@TXT2A,@T2,@ORDERNO,@BUYERBATCHNO,@BRAND,@TERMSDELIVERY,@TERMSPAYMENT,@COUNTRYDEST,@PRECARRIAGE,
@PLACEOFRECEIPT,@PORTLOADING,@PORTDISCHARGE,@TOTALPSC,@TOTALCARTONS,@TOTALCBM,@NETWEIGHT,@GWEIGHT,
@typeinvoce,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)

SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF    
END
GO
/****** Object:  StoredProcedure [dbo].[insertData_Invoicea1]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[insertData_Invoicea1]

@INVOICENO int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@CentralExise nvarchar (50),
@GenLedger nvarchar (50),
@SUBLEDGER nvarchar (50),
@STATION nvarchar (30),
@ModeOfPayment nvarchar (30),
@THROUGH nvarchar (30),
@BUNDLES nvarchar (20),
@WEIGHT nvarchar (10),
@BILTYNO nvarchar (20),
@BILTYDate nvarchar (10),
@FREIGHT nvarchar (20),
@TXT1  nvarchar (20),
@GAmount  float,
@NetAmount  float,
@with_withoutFormc nvarchar(1),
@netrate nvarchar(1),
@typeinvoice nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int


as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        


insert into Invoicea(INVOICENO,INVOICEDATE,IssueDate,DisPatchDate,CentralExise,GenLedger,SUBLEDGER,STATION,
ModeOfPayment,THROUGH,BUNDLES,WEIGHT,BILTYNO,BILTYDate,FREIGHT,TXT1,GAmount,NetAmount,with_withoutFormc,netrate,typeofinvoice,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@INVOICENO,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@CentralExise,@GenLedger,@SUBLEDGER,@STATION,
@ModeOfPayment,@THROUGH,@BUNDLES,@WEIGHT,@BILTYNO,convert(smalldatetime,@BILTYDate,103),@FREIGHT,@TXT1,@GAmount,@NetAmount,@with_withoutFormc,@netrate,@typeinvoice,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)





  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  StoredProcedure [dbo].[insertData_Invoicea]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[insertData_Invoicea]

@INVOICENO int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@CentralExise nvarchar (50),
@GenLedger nvarchar (50),
@SUBLEDGER nvarchar (50),
@STATION nvarchar (30),
@ModeOfPayment nvarchar (30),
@THROUGH nvarchar (30),
@BUNDLES nvarchar (20),
@WEIGHT nvarchar (10),
@BILTYNO nvarchar (20),
@BILTYDate nvarchar (10),
@FREIGHT nvarchar (20),
@TXT1  nvarchar (20),
@GAmount  float,
@NetAmount  float,
@with_withoutFormc nvarchar(1),
@netrate nvarchar(1),
@typeinvoice nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int


as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        


insert into Invoicea(INVOICENO,INVOICEDATE,IssueDate,DisPatchDate,CentralExise,GenLedger,SUBLEDGER,STATION,
ModeOfPayment,THROUGH,BUNDLES,WEIGHT,BILTYNO,BILTYDate,FREIGHT,TXT1,GAmount,NetAmount,with_withoutFormc,netrate,typeofinvoice,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@INVOICENO,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@CentralExise,@GenLedger,@SUBLEDGER,@STATION,
@ModeOfPayment,@THROUGH,@BUNDLES,@WEIGHT,@BILTYNO,convert(smalldatetime,@BILTYDate,103),@FREIGHT,@TXT1,@GAmount,@NetAmount,@with_withoutFormc,@netrate,@typeinvoice,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)





  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  StoredProcedure [dbo].[insertData_Groups]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[insertData_Groups]

@groupcode nvarchar (10),
@groupname nvarchar (50),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int

as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        

  insert into Groups(groupcode,groupname,createdby,createdon,updatedby,updatedon,fyear,setupid)
  values(@groupcode,@groupname,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)

  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  View [dbo].[partyProfile1]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[partyProfile1]
AS
SELECT     DESCFORINVOICE AS PARTY, ADDRESS1, ADDRESS2, ADDRESS3, Phone, contactp AS [CONTACT PERSON], remarks, DISTCODE, UPGuide, UKGuide, 
                      BBA_BCA, Btech, Ent_Guide, Gar_Adhyan, Mrt_Adhyan, SUBLEDGER
FROM         dbo.SLEDGER
WHERE     (gledger = 'SUNDRY DEBTORS')
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "SLEDGER"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 208
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'partyProfile1'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'partyProfile1'
GO
/****** Object:  View [dbo].[partyProfile]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[partyProfile]
AS
SELECT     DESCFORINVOICE AS PARTY, ADDRESS1, ADDRESS2, ADDRESS3, Phone, contactp AS [CONTACT PERSON], remarks, DISTCODE, UPGuide, UKGuide, 
                      BBA_BCA, Btech, Ent_Guide, Gar_Adhyan, Mrt_Adhyan, SUBLEDGER, mobile
FROM         dbo.SLEDGER
WHERE     (gledger = 'SUNDRY DEBTORS') AND (LEN(DESCFORINVOICE) > 0)
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "SLEDGER"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 208
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'partyProfile'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'partyProfile'
GO
/****** Object:  View [dbo].[RawStockSummary]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[RawStockSummary]
AS
SELECT     'p' AS Vtype, Dates, ItemName, Qty, Unit, setupid, fyear
FROM         dbo.FinishPurchase
UNION ALL
SELECT     'I' AS Vtype, Dates, ItemName, Qty, Unit, setupid, fyear
FROM         dbo.IssueDeppt
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4[30] 2[40] 3) )"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 3
   End
   Begin DiagramPane = 
      PaneHidden = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 5
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'RawStockSummary'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'RawStockSummary'
GO
/****** Object:  StoredProcedure [dbo].[insertData_credita]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[insertData_credita]

@INVOICENO int,
@INVOICEDATE varchar (11),
@IssueDate varchar (11),
@DisPatchDate varchar (11),
@CentralExise nvarchar (50),
@GenLedger nvarchar (50),
@SUBLEDGER nvarchar (50),
@STATION nvarchar (30),
@ModeOfPayment nvarchar (30),
@THROUGH nvarchar (30),
@BUNDLES nvarchar (20),
@WEIGHT nvarchar (10),
@BILTYNO nvarchar (20),
@BILTYDate nvarchar (10),
@FREIGHT nvarchar (20),
@TXT1  nvarchar (20),
@GAmount  float,
@NetAmount  float,
@with_withoutFormc nvarchar(1),
@netrate nvarchar(1),
@typeinvoice nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int


as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        


insert into credita(INVOICENO,INVOICEDATE,IssueDate,DisPatchDate,CentralExise,GenLedger,SUBLEDGER,STATION,
ModeOfPayment,THROUGH,BUNDLES,WEIGHT,BILTYNO,BILTYDate,FREIGHT,TXT1,GAmount,NetAmount,with_withoutFormc,netrate,typeofinvoice,createdby,createdon,updatedby,updatedon,fyear,setupid)
values(@INVOICENO,convert(smalldatetime,@INVOICEDATE,103),convert(smalldatetime,@IssueDate,103),convert(smalldatetime,@DisPatchDate,103),@CentralExise,@GenLedger,@SUBLEDGER,@STATION,
@ModeOfPayment,@THROUGH,@BUNDLES,@WEIGHT,@BILTYNO,convert(smalldatetime,@BILTYDate,103),@FREIGHT,@TXT1,@GAmount,@NetAmount,@with_withoutFormc,@netrate,@typeinvoice,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)





  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  StoredProcedure [dbo].[insertData_copyMaster]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[insertData_copyMaster]

@TypeofProduct nvarchar (50),
@BookNo nvarchar (20),
@ProductQuality nvarchar (30),
@rulling nvarchar (30),
@noofpages nvarchar (5),
@rate float,
@Book nvarchar(60),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int


as

BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        
	
  insert into copyMaster(TypeofProduct,BookNo,ProductQuality,rulling,noofpages,rate,Book,createdby,createdon,updatedby,updatedon,fyear,setupid)
  values(@TypeofProduct,@BookNo,@ProductQuality,@rulling,@noofpages,@rate,@Book,@createdby,getdate(),@updatedby,getdate(),@fyear,@setupid)

  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  StoredProcedure [dbo].[UpdateData_subledger]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[UpdateData_subledger] (

@gldger nvarchar (50),
@subledger nvarchar (50),
@address1 nvarchar (50),
@address2 nvarchar (50),
@phone nvarchar (25),
@FAX nvarchar (25),
@YEAROPENING float,
@bCityId nvarchar (10),
@radd1 nvarchar (50),
@radd2 nvarchar (50),
@rcityid nvarchar (10),
@rphone nvarchar (25),
@rfax nvarchar (25),
@contmail nvarchar (35),
@tpt nvarchar (100),
@range nvarchar (200),
@bank nvarchar (50),
@sadd1 nvarchar (50),
@slremark nvarchar (200),
@contperson nvarchar (40),
@contphone nvarchar (25),
@tinno nvarchar (50),
@TypeOfCust nvarchar(25),
@FYEAR nvarchar(10),
@setupid int,
@DESCFORINVOICE nvarchar (50),
@lblsubledger nvarchar (50)

)

as

BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        



  update SLEDGER set gledger=@gldger,address1=@address1,address2=@address2,Phone=@Phone,fax=@fax,YEAROPENING=@YEAROPENING,
                     bcityid=@bcityid,radd1=@radd1,radd2=@radd2,rcityid=@rcityid,rphone=@rphone,rfax=@rfax,email=@contmail,
                     tpt=@tpt,nRange=@range,bank=@bank,sadd1=@sadd1,slremark=@slremark,contperson=@contperson,
                     contphone=@contphone,tinno=@tinno,TypeOfCust=@TypeOfCust,FYEAR=@FYEAR,DESCFORINVOICE=@DESCFORINVOICE
  where subledger=@lblsubledger and fyear=@fyear and setupid=@setupid

  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  StoredProcedure [dbo].[UpdateData_Groups]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[UpdateData_Groups]

@groupcode nvarchar (10),
@groupname nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int


as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        

	
  update Groups set groupname=@groupname,updatedby=@updatedby,updatedon=getdate() 
  where groupcode=@groupcode and setupid=@setupid AND FYEAR=@FYEAR

 

  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  StoredProcedure [dbo].[UpdateData_copyMaster]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[UpdateData_copyMaster]

@TypeofProduct nvarchar (50),
@BookNo nvarchar (20),
@ProductQuality nvarchar (30),
@rulling nvarchar (30),
@noofpages nvarchar (5),
@rate float,
@Book nvarchar(60),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@setupid int

as

BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        
	
  update copyMaster set TypeofProduct=@TypeofProduct,ProductQuality=@ProductQuality,rulling=@rulling,
  noofpages=@noofpages,rate=@rate,updatedby=@updatedby,updatedon =getdate()
  where BookNo=@BookNo and setupid=@setupid and fyear=@fyear 

  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  View [dbo].[TotalBookRec]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[TotalBookRec]
AS
SELECT     dbo.INVOICEB_spRet.BOOKCODE, SUM(dbo.INVOICEB_spRet.QUANTITY) AS Qty, SUM(dbo.INVOICEB_spRet.NETAMOUNT) AS Amount
FROM         dbo.INVOICEA_spRet LEFT OUTER JOIN
                      dbo.INVOICEB_spRet ON dbo.INVOICEA_spRet.setupid = dbo.INVOICEB_spRet.setupid AND 
                      dbo.INVOICEA_spRet.Fyear = dbo.INVOICEB_spRet.Fyear AND dbo.INVOICEA_spRet.INVOICENO = dbo.INVOICEB_spRet.INVOICENO
GROUP BY dbo.INVOICEB_spRet.BOOKCODE
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "INVOICEA_spRet"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 188
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 24
         End
         Begin Table = "INVOICEB_spRet"
            Begin Extent = 
               Top = 6
               Left = 244
               Bottom = 200
               Right = 412
            End
            DisplayFlags = 280
            TopColumn = 11
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'TotalBookRec'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'TotalBookRec'
GO
/****** Object:  View [dbo].[TotalBookIssue]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[TotalBookIssue]
AS
SELECT     dbo.BOOKS.BOOKCODE, dbo.BOOKS.BOOKNAME, SUM(dbo.INVOICEB_sp.QUANTITY) AS Issue, SUM(dbo.INVOICEB_sp.NETAMOUNT) AS NetAmount, 
                      dbo.BOOKS.GROUPCODE
FROM         dbo.BOOKS INNER JOIN
                      dbo.INVOICEB_sp ON dbo.BOOKS.BOOKCODE = dbo.INVOICEB_sp.BOOKCODE AND dbo.BOOKS.Fyear = dbo.INVOICEB_sp.Fyear AND 
                      dbo.BOOKS.setupid = dbo.INVOICEB_sp.setupid
GROUP BY dbo.BOOKS.BOOKCODE, dbo.BOOKS.BOOKNAME, dbo.BOOKS.GROUPCODE
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "BOOKS"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 200
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "INVOICEB_sp"
            Begin Extent = 
               Top = 6
               Left = 244
               Bottom = 200
               Right = 412
            End
            DisplayFlags = 280
            TopColumn = 11
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'TotalBookIssue'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'TotalBookIssue'
GO
/****** Object:  View [dbo].[ReceiveBook]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[ReceiveBook]
AS
SELECT     dbo.INVOICEB_spRet.BOOKCODE, dbo.INVOICEA_spRet.AgentName, SUM(dbo.INVOICEB_spRet.QUANTITY) AS Qty, 
                      SUM(dbo.INVOICEB_spRet.NETAMOUNT) AS Amount, dbo.INVOICEA_spRet.setupid, dbo.INVOICEA_spRet.Fyear
FROM         dbo.INVOICEA_spRet LEFT OUTER JOIN
                      dbo.INVOICEB_spRet ON dbo.INVOICEA_spRet.INVOICENO = dbo.INVOICEB_spRet.INVOICENO
GROUP BY dbo.INVOICEB_spRet.BOOKCODE, dbo.INVOICEA_spRet.AgentName, dbo.INVOICEA_spRet.setupid, dbo.INVOICEA_spRet.Fyear
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "INVOICEA_spRet"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 200
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 23
         End
         Begin Table = "INVOICEB_spRet"
            Begin Extent = 
               Top = 6
               Left = 244
               Bottom = 200
               Right = 412
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'ReceiveBook'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'ReceiveBook'
GO
/****** Object:  View [dbo].[SaleReturnRegister]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[SaleReturnRegister]
AS
SELECT     dbo.CREDITA.INVOICEDATE, dbo.CREDITB.BOOKCODE, SUM(dbo.CREDITB.QUANTITY) AS Qty, 'Sales' AS Sales, 'Issue' AS Issue, 
                      dbo.CREDITA.Godown, dbo.CREDITA.fyear, dbo.CREDITA.setupid
FROM         dbo.CREDITA INNER JOIN
                      dbo.CREDITB ON dbo.CREDITA.INVOICENO = dbo.CREDITB.INVOICENO AND dbo.CREDITA.fyear = dbo.CREDITB.fyear AND 
                      dbo.CREDITA.setupid = dbo.CREDITB.setupid
GROUP BY dbo.CREDITA.INVOICEDATE, dbo.CREDITB.BOOKCODE, dbo.CREDITB.QUANTITY, dbo.CREDITA.Godown, dbo.CREDITA.fyear, dbo.CREDITA.setupid
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "CREDITA"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "CREDITB"
            Begin Extent = 
               Top = 6
               Left = 244
               Bottom = 121
               Right = 412
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SaleReturnRegister'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SaleReturnRegister'
GO
/****** Object:  View [dbo].[SaleRegister]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[SaleRegister]
AS
SELECT     dbo.INVOICEA.INVOICEDATE, dbo.INVOICEB.BOOKCODE, SUM(dbo.INVOICEB.QUANTITY) AS Qty, 'Sales' AS Sales, 'Issue' AS Issue, 
                      dbo.INVOICEA.Godown, dbo.INVOICEA.Fyear, dbo.INVOICEA.setupid
FROM         dbo.INVOICEA INNER JOIN
                      dbo.INVOICEB ON dbo.INVOICEA.INVOICENO = dbo.INVOICEB.INVOICENO AND dbo.INVOICEA.Fyear = dbo.INVOICEB.Fyear AND 
                      dbo.INVOICEA.setupid = dbo.INVOICEB.setupid
GROUP BY dbo.INVOICEA.INVOICEDATE, dbo.INVOICEB.BOOKCODE, dbo.INVOICEB.QUANTITY, dbo.INVOICEA.Godown, dbo.INVOICEA.Fyear, 
                      dbo.INVOICEA.setupid
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[41] 4[20] 2[14] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "INVOICEA"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 160
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 32
         End
         Begin Table = "INVOICEB"
            Begin Extent = 
               Top = 8
               Left = 354
               Bottom = 159
               Right = 522
            End
            DisplayFlags = 280
            TopColumn = 9
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SaleRegister'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SaleRegister'
GO
/****** Object:  View [dbo].[SpecimenReturnRegister]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[SpecimenReturnRegister]
AS
SELECT     dbo.INVOICEA_spRet.INVOICEDATE, dbo.INVOICEB_spRet.BOOKCODE, SUM(dbo.INVOICEB_spRet.QUANTITY) AS Qty, 'Sales' AS Sales, 
                      'Issue' AS Issue, dbo.INVOICEA_spRet.godown, dbo.INVOICEA_spRet.Fyear, dbo.INVOICEA_spRet.setupid
FROM         dbo.INVOICEA_spRet INNER JOIN
                      dbo.INVOICEB_spRet ON dbo.INVOICEA_spRet.INVOICENO = dbo.INVOICEB_spRet.INVOICENO AND 
                      dbo.INVOICEA_spRet.Fyear = dbo.INVOICEB_spRet.Fyear AND dbo.INVOICEA_spRet.setupid = dbo.INVOICEB_spRet.setupid
GROUP BY dbo.INVOICEA_spRet.INVOICEDATE, dbo.INVOICEB_spRet.BOOKCODE, dbo.INVOICEB_spRet.QUANTITY, dbo.INVOICEA_spRet.godown, 
                      dbo.INVOICEA_spRet.Fyear, dbo.INVOICEA_spRet.setupid
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "INVOICEA_spRet"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "INVOICEB_spRet"
            Begin Extent = 
               Top = 6
               Left = 244
               Bottom = 121
               Right = 412
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SpecimenReturnRegister'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SpecimenReturnRegister'
GO
/****** Object:  View [dbo].[SpecimenRegister]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[SpecimenRegister]
AS
SELECT     dbo.INVOICEA_sp.INVOICEDATE, dbo.INVOICEB_sp.BOOKCODE, SUM(dbo.INVOICEB_sp.QUANTITY) AS Qty, 'Sales' AS Sales, 'Issue' AS Issue, 
                      dbo.INVOICEA_sp.godown, dbo.INVOICEA_sp.Fyear, dbo.INVOICEA_sp.setupid, dbo.INVOICEA_sp.AgentName
FROM         dbo.INVOICEA_sp INNER JOIN
                      dbo.INVOICEB_sp ON dbo.INVOICEA_sp.INVOICENO = dbo.INVOICEB_sp.INVOICENO AND dbo.INVOICEA_sp.Fyear = dbo.INVOICEB_sp.Fyear AND 
                      dbo.INVOICEA_sp.setupid = dbo.INVOICEB_sp.setupid
GROUP BY dbo.INVOICEA_sp.INVOICEDATE, dbo.INVOICEB_sp.BOOKCODE, dbo.INVOICEB_sp.QUANTITY, dbo.INVOICEA_sp.godown, dbo.INVOICEA_sp.Fyear, 
                      dbo.INVOICEA_sp.setupid, dbo.INVOICEA_sp.SUBLEDGER, dbo.INVOICEA_sp.AgentName
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[28] 4[10] 2[40] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "INVOICEA_sp"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 26
         End
         Begin Table = "INVOICEB_sp"
            Begin Extent = 
               Top = 6
               Left = 244
               Bottom = 121
               Right = 412
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 11
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SpecimenRegister'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'SpecimenRegister'
GO
/****** Object:  View [dbo].[QryTotalQuantity]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[QryTotalQuantity]
AS
SELECT     INVOICENO, SUM(QUANTITY) AS TQty, fyear, setupid
FROM         dbo.INVOICEB
GROUP BY INVOICENO, fyear, setupid
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "INVOICEB"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'QryTotalQuantity'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'QryTotalQuantity'
GO
/****** Object:  View [dbo].[QryInvoiceC]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[QryInvoiceC]
AS
SELECT DISTINCT INVOICENO, SUBLEDGER, fyear, setupid
FROM         dbo.INVOICEC
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "INVOICEC"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 200
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'QryInvoiceC'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'QryInvoiceC'
GO
/****** Object:  View [dbo].[ProductWiseSale]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[ProductWiseSale]
AS
SELECT     dbo.INVOICEB.BOOKCODE, dbo.copyMaster.TypeofProduct, dbo.copyMaster.rulling, dbo.copyMaster.NoOfPages, dbo.copyMaster.ProductQuality, 
                      dbo.INVOICEB.QUANTITY, dbo.INVOICEB.INVOICEDATE, MONTH(CONVERT(smalldatetime, dbo.INVOICEB.INVOICEDATE, 103)) AS MName, 
                      dbo.INVOICEB.setupid, dbo.INVOICEB.fyear
FROM         dbo.copyMaster INNER JOIN
                      dbo.INVOICEB ON dbo.copyMaster.BookNo = dbo.INVOICEB.BOOKCODE AND dbo.copyMaster.FYear = dbo.INVOICEB.fyear AND 
                      dbo.copyMaster.setupid = dbo.INVOICEB.setupid
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "copyMaster"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 197
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "INVOICEB"
            Begin Extent = 
               Top = 6
               Left = 235
               Bottom = 121
               Right = 387
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'ProductWiseSale'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'ProductWiseSale'
GO
/****** Object:  StoredProcedure [dbo].[insertData_creditb]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[insertData_creditb]

@INVOICENO int,
@INVOICEDATE varchar (11),
@SUBLEDGER nvarchar (50),
@PRINTORDER int,
@BOOKCODE varchar(10),
@QUANTITY int,
@RATE float,	
@NetRate float,
@Amount float,
@typeinvoice nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@itemname nvarchar (120),
@setupid int


as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        


insert into creditb(INVOICENO,INVOICEDATE,SUBLEDGER,PRINTORDER,BOOKCODE,QUANTITY,Rate,netrate,amount,typeofinvoice,createdby,
createdon,updatedby,updatedon,fyear,itemname,setupid)
values(@INVOICENO,convert(smalldatetime,@INVOICEDATE,103),@SUBLEDGER,@PRINTORDER,@BOOKCODE,@QUANTITY,@Rate,@netrate,@amount,@typeinvoice,
@createdby,getdate(),@updatedby,getdate(),@fyear,@itemname,@setupid)

  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  View [dbo].[IssueBook]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[IssueBook]
AS
SELECT     dbo.BOOKS.BOOKCODE, dbo.BOOKS.BOOKNAME, dbo.INVOICEB_sp.agentname, SUM(dbo.INVOICEB_sp.QUANTITY) AS Issue, 
                      SUM(dbo.INVOICEB_sp.NETAMOUNT) AS NetAmount, dbo.BOOKS.GROUPCODE, dbo.BOOKS.Fyear, dbo.BOOKS.setupid
FROM         dbo.BOOKS INNER JOIN
                      dbo.INVOICEB_sp ON dbo.BOOKS.BOOKCODE = dbo.INVOICEB_sp.BOOKCODE AND dbo.BOOKS.Fyear = dbo.INVOICEB_sp.Fyear AND 
                      dbo.BOOKS.setupid = dbo.INVOICEB_sp.setupid
GROUP BY dbo.BOOKS.BOOKCODE, dbo.BOOKS.BOOKNAME, dbo.INVOICEB_sp.agentname, dbo.BOOKS.GROUPCODE, dbo.BOOKS.Fyear, dbo.BOOKS.setupid
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "BOOKS"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 200
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 5
         End
         Begin Table = "INVOICEB_sp"
            Begin Extent = 
               Top = 6
               Left = 244
               Bottom = 200
               Right = 412
            End
            DisplayFlags = 280
            TopColumn = 11
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'IssueBook'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'IssueBook'
GO
/****** Object:  StoredProcedure [dbo].[insertData_Invoiceb]    Script Date: 09/05/2013 01:12:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[insertData_Invoiceb]

@INVOICENO int,
@INVOICEDATE varchar (11),
@SUBLEDGER nvarchar (50),
@PRINTORDER int,
@BOOKCODE varchar(10),
@QUANTITY int,
@RATE float,	
@NetRate float,
@Amount float,
@typeinvoice nvarchar(10),
@createdby nvarchar (50),
@updatedby nvarchar (25),
@fyear nvarchar (10),
@itemname nvarchar (120),
@setupid int


as
BEGIN        
  
  SET NOCOUNT ON        
  SET ANSI_WARNINGS OFF        


insert into Invoiceb(INVOICENO,INVOICEDATE,SUBLEDGER,PRINTORDER,BOOKCODE,QUANTITY,Rate,netrate,amount,typeofinvoice,createdby,
createdon,updatedby,updatedon,fyear,itemname,setupid)
values(@INVOICENO,convert(smalldatetime,@INVOICEDATE,103),@SUBLEDGER,@PRINTORDER,@BOOKCODE,@QUANTITY,@Rate,@netrate,@amount,@typeinvoice,
@createdby,getdate(),@updatedby,getdate(),@fyear,@itemname,@setupid)

  SET ANSI_WARNINGS ON        
  SET NOCOUNT OFF        
END
GO
/****** Object:  View [dbo].[AgentWiseSale]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[AgentWiseSale]
AS
SELECT     dbo.CASHB_basil.agentname, dbo.CASHB_basil.BOOKCODE, dbo.BOOKS.BOOKNAME, dbo.BOOKS.GROUPCODE, dbo.CASHB_basil.NETAMOUNT, 
                      dbo.CASHA_basil.Sale_Return
FROM         dbo.CASHB_basil INNER JOIN
                      dbo.BOOKS ON dbo.CASHB_basil.BOOKCODE = dbo.BOOKS.BOOKCODE INNER JOIN
                      dbo.CASHA_basil ON dbo.CASHB_basil.INVOICENO = dbo.CASHA_basil.INVOICENO
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "CASHB_basil"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 190
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "BOOKS"
            Begin Extent = 
               Top = 6
               Left = 228
               Bottom = 121
               Right = 380
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "CASHA_basil"
            Begin Extent = 
               Top = 6
               Left = 418
               Bottom = 121
               Right = 573
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'AgentWiseSale'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'AgentWiseSale'
GO
/****** Object:  View [dbo].[CashSaleRegister]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[CashSaleRegister]
AS
SELECT     dbo.CASHA.INVOICEDATE, dbo.CASHB.BOOKCODE, SUM(dbo.CASHB.QUANTITY) AS Qty, 'Sales' AS Sales, 'Issue' AS Issue, dbo.CASHA.Godown, 
                      dbo.CASHA.Fyear, dbo.CASHA.setupid
FROM         dbo.CASHA INNER JOIN
                      dbo.CASHB ON dbo.CASHA.INVOICENO = dbo.CASHB.INVOICENO AND dbo.CASHA.Fyear = dbo.CASHB.Fyear AND 
                      dbo.CASHA.setupid = dbo.CASHB.setupid
GROUP BY dbo.CASHA.INVOICEDATE, dbo.CASHB.BOOKCODE, dbo.CASHB.QUANTITY, dbo.CASHA.Godown, dbo.CASHA.Fyear, dbo.CASHA.setupid
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "CASHA"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 121
               Right = 209
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "CASHB"
            Begin Extent = 
               Top = 6
               Left = 247
               Bottom = 121
               Right = 415
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'CashSaleRegister'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'CashSaleRegister'
GO
/****** Object:  View [dbo].[BinderIssueRegister]    Script Date: 09/05/2013 01:12:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[BinderIssueRegister]
AS
SELECT     dbo.INVOICEA_IssuedBind.INVOICEDATE, dbo.INVOICEB_IssuedBind.BOOKCODE, SUM(dbo.INVOICEB_IssuedBind.TotalBook) AS Qty, 'Issue' AS Issue,
                       dbo.INVOICEA_IssuedBind.godown, dbo.INVOICEA_IssuedBind.Fyear, dbo.INVOICEA_IssuedBind.setupid, 
                      dbo.INVOICEA_IssuedBind.SUBLEDGER
FROM         dbo.INVOICEA_IssuedBind INNER JOIN
                      dbo.INVOICEB_IssuedBind ON dbo.INVOICEA_IssuedBind.INVOICENO = dbo.INVOICEB_IssuedBind.INVOICENO AND 
                      dbo.INVOICEA_IssuedBind.Fyear = dbo.INVOICEB_IssuedBind.Fyear AND 
                      dbo.INVOICEA_IssuedBind.setupid = dbo.INVOICEB_IssuedBind.setupid
GROUP BY dbo.INVOICEA_IssuedBind.INVOICEDATE, dbo.INVOICEB_IssuedBind.BOOKCODE, dbo.INVOICEA_IssuedBind.godown, 
                      dbo.INVOICEA_IssuedBind.Fyear, dbo.INVOICEA_IssuedBind.setupid, dbo.INVOICEA_IssuedBind.SUBLEDGER
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "INVOICEA_IssuedBind"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 200
               Right = 206
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "INVOICEB_IssuedBind"
            Begin Extent = 
               Top = 21
               Left = 377
               Bottom = 136
               Right = 545
            End
            DisplayFlags = 280
            TopColumn = 15
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 12
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'BinderIssueRegister'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'BinderIssueRegister'
GO
/****** Object:  Default [DF_TemprptTrialBalance_OpeningBalance]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[TemprptTrialBalance] ADD  CONSTRAINT [DF_TemprptTrialBalance_OpeningBalance]  DEFAULT ((0)) FOR [OpeningBalance]
GO
/****** Object:  Default [DF_TemprptTrialBalance_DAmount]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[TemprptTrialBalance] ADD  CONSTRAINT [DF_TemprptTrialBalance_DAmount]  DEFAULT ((0)) FOR [DAmount]
GO
/****** Object:  Default [DF_TemprptTrialBalance_CAmount]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[TemprptTrialBalance] ADD  CONSTRAINT [DF_TemprptTrialBalance_CAmount]  DEFAULT ((0)) FOR [CAmount]
GO
/****** Object:  Default [DF_TemprptTrialBalance_ClosingBalance]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[TemprptTrialBalance] ADD  CONSTRAINT [DF_TemprptTrialBalance_ClosingBalance]  DEFAULT ((0)) FOR [ClosingBalance]
GO
/****** Object:  Default [DF_treport_OpeningBalance]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[treport] ADD  CONSTRAINT [DF_treport_OpeningBalance]  DEFAULT ((0)) FOR [OpeningBalance]
GO
/****** Object:  Default [DF_treport_ad]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[treport] ADD  CONSTRAINT [DF_treport_ad]  DEFAULT ((0)) FOR [ad]
GO
/****** Object:  Default [DF_treport_ac]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[treport] ADD  CONSTRAINT [DF_treport_ac]  DEFAULT ((0)) FOR [ac]
GO
/****** Object:  Default [DF_treport_balance]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[treport] ADD  CONSTRAINT [DF_treport_balance]  DEFAULT ((0)) FOR [balance]
GO
/****** Object:  ForeignKey [FK_CASHB_BOOKS]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[CASHB]  WITH CHECK ADD  CONSTRAINT [FK_CASHB_BOOKS] FOREIGN KEY([BOOKCODE])
REFERENCES [dbo].[BOOKS] ([BOOKCODE])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[CASHB] CHECK CONSTRAINT [FK_CASHB_BOOKS]
GO
/****** Object:  ForeignKey [FK_CASHB_basil_BOOKS1]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[CASHB_basil]  WITH CHECK ADD  CONSTRAINT [FK_CASHB_basil_BOOKS1] FOREIGN KEY([BOOKCODE])
REFERENCES [dbo].[BOOKS] ([BOOKCODE])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[CASHB_basil] CHECK CONSTRAINT [FK_CASHB_basil_BOOKS1]
GO
/****** Object:  ForeignKey [FK_CASHB_basilRet_BOOKS]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[CASHB_basilRet]  WITH CHECK ADD  CONSTRAINT [FK_CASHB_basilRet_BOOKS] FOREIGN KEY([BOOKCODE])
REFERENCES [dbo].[BOOKS] ([BOOKCODE])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[CASHB_basilRet] CHECK CONSTRAINT [FK_CASHB_basilRet_BOOKS]
GO
/****** Object:  ForeignKey [FK_CREDITB_BOOKS]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[CREDITB]  WITH CHECK ADD  CONSTRAINT [FK_CREDITB_BOOKS] FOREIGN KEY([BOOKCODE])
REFERENCES [dbo].[BOOKS] ([BOOKCODE])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[CREDITB] CHECK CONSTRAINT [FK_CREDITB_BOOKS]
GO
/****** Object:  ForeignKey [FK_INVOICEB_BOOKS]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[INVOICEB]  WITH CHECK ADD  CONSTRAINT [FK_INVOICEB_BOOKS] FOREIGN KEY([BOOKCODE])
REFERENCES [dbo].[BOOKS] ([BOOKCODE])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[INVOICEB] CHECK CONSTRAINT [FK_INVOICEB_BOOKS]
GO
/****** Object:  ForeignKey [FK_INVOICEB_INVOICEA]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[INVOICEB]  WITH CHECK ADD  CONSTRAINT [FK_INVOICEB_INVOICEA] FOREIGN KEY([INVOICENO])
REFERENCES [dbo].[INVOICEA] ([INVOICENO])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[INVOICEB] CHECK CONSTRAINT [FK_INVOICEB_INVOICEA]
GO
/****** Object:  ForeignKey [FK_INVOICEB_IssuedBind_BOOKS]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[INVOICEB_IssuedBind]  WITH CHECK ADD  CONSTRAINT [FK_INVOICEB_IssuedBind_BOOKS] FOREIGN KEY([Book_Code])
REFERENCES [dbo].[BOOKS] ([BOOKCODE])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[INVOICEB_IssuedBind] CHECK CONSTRAINT [FK_INVOICEB_IssuedBind_BOOKS]
GO
/****** Object:  ForeignKey [FK_INVOICEB_sp_BOOKS]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[INVOICEB_sp]  WITH CHECK ADD  CONSTRAINT [FK_INVOICEB_sp_BOOKS] FOREIGN KEY([BOOKCODE])
REFERENCES [dbo].[BOOKS] ([BOOKCODE])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[INVOICEB_sp] CHECK CONSTRAINT [FK_INVOICEB_sp_BOOKS]
GO
/****** Object:  ForeignKey [FK_INVOICEB_spRet_BOOKS]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[INVOICEB_spRet]  WITH CHECK ADD  CONSTRAINT [FK_INVOICEB_spRet_BOOKS] FOREIGN KEY([BOOKCODE])
REFERENCES [dbo].[BOOKS] ([BOOKCODE])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[INVOICEB_spRet] CHECK CONSTRAINT [FK_INVOICEB_spRet_BOOKS]
GO
/****** Object:  ForeignKey [FK_INVOICEC_INVOICEA]    Script Date: 09/05/2013 01:12:13 ******/
ALTER TABLE [dbo].[INVOICEC]  WITH CHECK ADD  CONSTRAINT [FK_INVOICEC_INVOICEA] FOREIGN KEY([INVOICENO])
REFERENCES [dbo].[INVOICEA] ([INVOICENO])
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[INVOICEC] CHECK CONSTRAINT [FK_INVOICEC_INVOICEA]
GO
