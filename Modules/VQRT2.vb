Option Strict Off
Option Explicit On

Imports C1.C1Preview
Imports MySql.Data.MySqlClient
Imports VB = Microsoft.VisualBasic

Module VQRT2
    '#Const MARK = 0 ' Mark Ltg Version
    'MARK = Swap Status & Salesman
    Public SecurityBrancheCodes As String = "" '08-30-13
    Public SecurityLevel As String = "" '08-30-13

    Public DIST As Boolean '01-14-09
    '#Const DIST = 0 ' VQRD
    Public MFG As Boolean
    Public MaxNameLength As Int16 = 45 'Val(Me.rbnMaxNameTxt.Text) '12-23-12 45   Public
    Public MaxJobLength As Int16 = 40 '12-23-12
    Public DecFormat As String = "########0.00" '01-01-12 NoCents "########0") DecFormat for Nocents Whole Dollars
    Public LemonChiffon As Color = Color.LemonChiffon '01-18-13 ' 23 Times   DarkGrey 
    Public AntiqueWhite As Color = Color.AntiqueWhite '01-18-13   0  Times   DimGrey  
    Public LightGray As Color = Color.LightGray '01-18-13         5 Times   
    Public LightSkyBlue As Color = Color.LightSkyBlue '01-18-13   12 Times   
    Public PrtGrayScale As Boolean = False '01-18-13
    Public ServerPath As String = "" '01-03-18

    '#Const MFG = 0  'MFG version Sell,Cost,Margin and Status Column to REP# Column, Swap Status & Salesman
    'CHLO = MFG = 1  VQRTCHLO
    Public DAYB As Boolean
    '#Const DAYB = 0 'Set MFG =1 & DAYB =1  DayBrite Do Not Swap Status & Salesman 08-28-01
    'VQRTDAYB
    Public ELA As Boolean
    '#Const ELA = 0  '10-17-07 WNA Special for ELA - Instr RetrCode
    'REP  Columns   Sell   Comm$ Price2 Price3 Price4 Price5 Etc
    'MFG  Columns = Sell   Cost  DN15   DN10   DN5    Rep%
    'CHLO Columns = Sell   Cost  DN15   DN10   DN5    Rep%
    'DAYB Columns = Cost   Sell  Rep%   Price3 Price4 Price5 Etc
    'MOBR Columns = Cost   Sell  Rep%   Price3 Price4 Price5 Etc
    'DIVE Columns = Price1 Sell  Comm$  Price2 Price3 Price4 Price5 ' Price3 = Buy Price Total on Price3
    'TCAN Columns = Cost   Sell  Rep%   PI-Cd  West   Dist   Stk
    'Table
    Public doc As New C1PrintDocument
    Public RT As New C1.C1Preview.RenderTable
    Public RTotals As New C1.C1Preview.RenderText
    Public ra1 As New RenderArea
    Public ra2 As New RenderArea
    Public fs As Integer = 10 '05-03-10 - font size combobox
    Public strSql As String '04-21-08
    Public BranchSql As String '10-15-13 
    Public myConnectionString As String = "Database=saw8;Data Source=localhost;User Id=root;Password=saw987"
    Public myConnection As MySqlConnection
    Public myDataAdapter As MySqlDataAdapter
    Public myds As dsSaw8       'DataSet Detail Tab (This is the one to Save To)
    Public ds As dsSaw8         'DataSet LOOKUP TAB (Combined Table)
    Public dsRpt As dsSaw8      'DataSet REPORTS
    Public dsQuoteRealLU As dsSaw8 '02-11-09
    Public dsSESCOSpecifiers As dsSaw8 '04-17-12
    Public dsQutLU As dsSaw8
    Public dsQuote As dsSaw8
    Public dsadmin As dsSaw8 '08-13-13
    Public drarray() As DataRow 'Array of DataRows - Used in Printing Labels
    Public Adding As Boolean
    Public dsGrid As DataSet = New DataSet("dsGrid") '01-28-09
    Public table As DataTable = dsGrid.Tables.Add("Items")
    'Data Adapters for the Detail Tab
    Public daQuoteRealLU As MySql.Data.MySqlClient.MySqlDataAdapter
    Public daQuote As MySql.Data.MySqlClient.MySqlDataAdapter
    Public daProject As MySql.Data.MySqlClient.MySqlDataAdapter
    Public daQuoteLine As MySql.Data.MySqlClient.MySqlDataAdapter
    Public daQuoteLinePrice As MySql.Data.MySqlClient.MySqlDataAdapter
    Public daQuoteNotes As MySql.Data.MySqlClient.MySqlDataAdapter
    Public daQuoteSLS As MySql.Data.MySqlClient.MySqlDataAdapter
    Public daProjCust As MySql.Data.MySqlClient.MySqlDataAdapter
    Public daQutLU As MySql.Data.MySqlClient.MySqlDataAdapter
    Public daQuoteTo As MySql.Data.MySqlClient.MySqlDataAdapter
    Public drQRow As dsSaw8.QUTLU1Row '01-25-09
    Public drQToRow As dsSaw8.QuoteRealLURow '02-11-09
    Public drQline As dsSaw8.quotelinesRow '02-09-10
    Public dsSCross As New DataSet '02-13-13 WNA

    Public AGnam As String
    Public B As String
    Public BeginCode As String '01-15-09
    Public Bid As String '12-08-98 WNA
    Public BranchReporting As Boolean '05-31-07 JH
    Public BranchCodeRpt As String '06-15-10
    Public BrandReport As Boolean '05-16-13
    Public BrandReportMfg As String = "" '10-20-13  Philips and Cooper PHIL COOP
    Public BrandList As String = "" '10-20-13  Philips PHIL list = "DAYB,MCPH, 
    Public ForecastAllMfg As Boolean '05-14-15 JTC Public ForecastAllMfg Forecasting for MFGs Except Philips and SESCO
    Public DefChangeEstDelDate As String = ""     'Y/N    02/17/14  
    Public DefChgEstDelDateCodes As String = ""   'Status Codes Seperated by commas

    Public SLTCommAmt As Decimal '01-18-09
    Public SLTCommPct As Decimal '01-18-09
    Public JOBSER As String
    Public BrkMajorL1 As String
    Public BrkSecondL2 As String
    Public BrkThirdL3 As String '01-18-09 Public
    Public TgCol(130) As Int16 '02-03-09 
    Public TgHeading(130) As String 'Dim TgHeading(130) As String '02-18-09
    Public TgName(130) As String '02-15-09 
    Public TgWidth(130) As Single '02-09-09 
    Public DE As Single
    Public DT As String
    Public F As Int16 '01-14-09
    Public I As Int16
    Public MyControlNamesList As String '01-24-09
    'Public Mbackcolor As Integer
    'Public MBackCol As String
    Public MFGCustNumber As String '05-31-07 JH ' branch reporting
    Public multsrtrvs(8) As String '01-25-09
    Public NoMstrMatch As String
    Public Fds(2500) As String
    Public FirstP As Short
    Public fileExists As Boolean '11-24-08
    Public FirmName As String '05-08-08 JTC
    Public GT As Decimal '12-08-98 WNA
    Public GTG As Decimal
    Public HeaderPrtFlg As Short '12-08-98 WNA
    Public SESCO As Boolean '05-30-11
    Public ExcelQuoteFU As Boolean '04-28-15 JTC
    '01-15-99 WNA Public Look As String, TimeStamp As String
    '01-15-99 WNA Public Savefd(80) As String, FD(26)  As String
    '01-15-99 WNA Public LEV As Integer, FileNum As Integer, TotFiles As Integer
    'Public Lines As Integer
    'SEW 04-05-99 Public Size As Integer  '01-15-99 WNA , TotSize As Integer
    Public NoRept As Short 'tag 08-11-98
    '01-15-99 WNA Public ODate As String, ODate1 As String
    Public MAXJ As Short 'wlc03-24-97
    Public MaxCol As Int16 '02-03-09 = frmQuoteRpt.tg.Splits(0).DisplayColumns.Count - 1
    '01-15-99 WNA Public MonthI As Integer
    Public DetOrTot As String
    Public OrderBy As String '01-18-09
    'Public US As String '01-18-09 US = "Quote Code Sequence"
    'Public UH As String '01-18-09 :UH = "QUOTE CODE SEQUENCE"
    Public ExportExcelProductLines As Boolean = False '02-15-10
    Public Page As Integer '01-14-09
    Public PageBreak As Boolean = False '02-15-09
    Public PrevStat As String '01-13-09
    Public PQTCUST As String
    Public PREVCUST As String '12-08-98 WNA
    Public PrevCustName As String '01-30-06
    Public PREVSLS As String '06-11-99 WNA
    Public PrtMaxLines As Short
    Public ReadCt As Single
    Public RecordCt As Single
    Public RecordBreakCt As Single
    'Public RepType As Integer
    Public RepCustNumber As String '05-31-07 JH 'BRANCH REPORTING
    Public SavePrtFonts As Single
    Public SLT As Decimal '01-13-09
    Public SubCount As Decimal '08-10-99 WNA
    Public SubTotFlag As Short '06-14-99 WNA
    Public THDG As String '01-13-09
    Public TotalsOnly As Short
    Public WNS As String
    Public Wspcs As String
    ' Specific to VQRT
    'Public QRR As Decimal '12-08-98 WNA
    'moved down Public SubSeq As Integer   'Sub Menu was Mus% '02-21-02
    Public DI As String
    Public SI As String
    Public SortNeeded As String '02-21-02
    Public SelectionText As String ' tag 06-19-97
    Public RptL As Short
    Public RptSpecifiers As Short '01-28-09
    Public RealCustomerOnly As Boolean = 0 '03-12-14
    Public RealCustomer As Boolean = 0 '01-18-12 -C
    Public RealManufacturer As Boolean = 0 '01-18-12  - M
    Public RealQuoteTOOther As Boolean = 0 '01-31-12
    Public RealQuoteToAmtON As Boolean = 0 '01-06-14
    Public RealSLSCustomer As Boolean = 0 '01-18-12 
    Public RealArchitect As Boolean = 0 '01-18-12  - A
    Public RealEngineer As Boolean = 0 '01-18-12  - E
    Public RealLtgDesigner As Boolean = 0 '01-18-12 - L
    Public RealSpecifier As Boolean = 0 '01-18-12 - S
    Public RealContractor As Boolean = 0 '01-18-12 - T
    Public RealOther As Boolean = 0 'Specifier Other01-18-12 - O
    Public RealWithOneMfgCust As Boolean = 0 '03-22-13
    Public RealWithOneMfgCustCode As String = "" '03-22-13
    Public RealWithOneMfgCustSortJobName As Boolean = 0 '10-13-14 JTC
    Public RealALL As Boolean = 0 '01-18-12 
    Public RealExtByInfluencePercent As Boolean = 0 '02-04-12 '11-27-13 as boolean
    Public RealTgLookupExcel As Boolean = 0 '11-27-13
    Public RptNotes As Short '01-28-09
    Public RptMFGCust As Short 'tag 10-15-97
    Public RptMFG As Short 'tag 10-15-97
    Public RptCust As Short 'tag 10-15-97
    Public TerrSpecRegAllPaidUnPaid As String = "A" '09-10-12 A=All,P=Paid, U=Unpaid
    Public TOTDG As Decimal
    Public TOTD As Decimal
    Public TotSell As Decimal '06-14-99 WNA
    Public TotCost As Decimal '06-14-99 WNA
    Public TotProfit As Decimal '06-14-99 WNA
    Public TRCT As String '12-08-98 WNA
    Public Resp As Int16 '07-12-09
    Public UserID As String ' Denotes the User ID JTC
    Public UserDir As String ' = "C:\SAW8\USER\JTC\" & UserID '12-22-08 C:\SAW8\USER\JTC\ Set in FormLoad
    Public UserSysDir As String '= "C:\SAW8\USER\SYS\"
    Public UserDocDir As String '06-02-09
    Public UserPath As String '= "C:\SAW" '06-03-09
    Public UserPathImages As String ' UserPath & "\IMAGES\" '02-15-10
    Public UserPathHelp As String '08-27-10 = UserPath & "HELP\" '08-27-10
    Public SecurityAdministrator As Boolean '07-08-09
    Public ShowHideFileName As String
    Public ShowHideColMoved As Int16 = 0 '01-08-09 1 = needs to run ShowHideColMove()
    Public DebugOn As Boolean '07-25-09
    Public US As String
    Public UH As String
    Public MU As Short
    'Public LMargin As Short 'SEW 09-18-97
    Public Orientation As Short 'SEW 09-18-97
    Public PaperSize As Short '09-09-05 JTC Public PaperSize% = 9 = A4 International 8.27 X 11.69 (210x297mm) VREPORTS.INI PaperSize=9 At end of AUTO section
    'RptQutCode=2,  RptProj=3, RptEntryDate=4, RptSalesman=5, RptStatus=6 '02-19-02
    'RptBidDate=7, RptDescend=8, RptSpecif=12
    'SubSProj=1, SubSSls=2, SubSStatus=3, SubSBidDate=4, SubSDescend=5
    Enum RptMajorType ' 02-19-02 Public/Public By Default Long By Default
        RptQutCode = 2 ' MU% = 2
        RptProj = 3 ' MU% = 3
        RptEntryDate = 4 ' MU% = 4
        RptSalesman = 5 ' MU% = 5
        RptStatus = 6 ' MU% = 6
        RptBidDate = 7 ' MU% = 7
        RptDescend = 8 ' MU% = 8
        RptSpecif = 12 ' MU% = 12  Specifier Rpt
        RptLocation = 13 ' MU% = 13  Location Rpt
        RptRetrieval = 14 ' Retrieval Code
        RptMarketSegment = 15 'Market Segment 07-30-04
        RptFollowBy = 16 'FollowedBy 03-01-12
        RptEnteredBy = 17 'EnteredBy '05-14-13
    End Enum 'RepType
    Public RepType As RptMajorType
    Enum SubSortType ' 02-19-02 Public/Public By Default Long By Default
        SubSProj = 1 'MUS% = 1
        SubSSls = 2 'MUS% = 2
        SubSStatus = 3 'MUS% = 3
        SubSBidDate = 4 'MUS% = 4
        SubSDescend = 5 'MUS% = 5
        SubSProjCode = 6 '11-20-10
        SubSEnterDate = 7 '11-20-10 
        SubSSpecif = 12 'Salesman/Specifier
        SubSSelectBidDate = 14 '03-03-12 Job Name/BidDate
    End Enum 'subseq     subseq = SubSSpecif = Specifier with Major sort by Salesman-
    Public SubSeq As SubSortType
    Enum TotalLevels
        '01-25-09 TotalLevels.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
        TotGt = 0
        TotLv1 = 1
        TotLv2 = 2
        TotLv3 = 3
        TotLv4 = 4
        TotPrt = 5
    End Enum
    Public Lev As TotalLevels
    'FixSell, FixCost, FixProfit, LampSell, LampCost, ProfitLamp, CommAmt, Commpct
    '01-25-09 TotalLevels.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt
    Public FixMargin As Decimal '09-08-09
    Public LpMargin As Decimal
    Public FixSell As Decimal
    Public FixCost As Decimal
    Public LampSell As Decimal
    Public LampCost As Decimal
    Public LampProfit As Decimal
    'Public FixMargin As Decimal '09-08-09
    'Public LPMargin As Decimal '09-08-09
    Public LnQuantityA As Decimal '09-08-09 
    Public QuantityA(5) As Decimal '09-08-09 
    Public FixSellExt As Decimal '09-11-09 
    Public FixCostExt As Decimal
    Public RealizSellExt As Decimal '12-10-09
    Public RealizSellAExt(5) As Decimal
    Public SellFixtureAExt(5) As Decimal '    GTSellFixture = GTSellFixture + FixSell
    Public CostFixtureAExt(5) As Decimal '        GTCostFixture = GTCostFixture + FixCost
    Public SellFixtureA(5) As Decimal '    GTSellFixture = GTSellFixture + FixSell
    'Public TOTDSellFixture(5) As Decimal '        TOTDSellFixture = TOTDSellFixture + FixSell
    Public CostFixtureA(5) As Decimal '        GTCostFixture = GTCostFixture + FixCost
    'Public TOTDCostFixture(5) As Decimal '        TOTDCostFixture = TOTDCostFixture + FixCost
    Public ProfitFixtureA(5) As Decimal '        GTProfitFixture = GTProfitFixture + FixProfit
    'Public TOTDProfitFixture(5) As Decimal '        TOTDProfitFixture = TOTDProfitFixture + FixProfit
    Public LampSellA(5) As Decimal '        GTLampSell = GTLampSell + LampSell
    'Public TOTDLampSell(5) As Decimal '        TOTDLampSell = TOTDLampSell + LampSell
    Public LampCostA(5) As Decimal '        GTLampCost = GTLampCost + LampCost
    Public ProfitLampA(5) As Decimal '        TOTDLampCost = TOTDLampCost + LampCost
    Public CommAmtA(5) As Decimal '        GTCommAmt = GTCommAmt + CommAmt '02-26-01 WNA
    Public CommPctA(5) As Decimal '        GTCommPct = (GTCommAmt / (GT + 0.0001)) * 100 '02-26-01 WNA
    'Public (5) as Decimal'        TOTD = TOTD + Amt
    'Public (5) as Decimal'        TOTDCommAmt = TOTDCommAmt + CommAmt '02-26-01 WNA
    'Public (5) as Decimal'       TOTDCommPct = (TOTDCommAmt / (TOTD + 0.0001)) * 100 '02-26-01 WNA
    'Public LnType As Short = 0 '09-09-09 From Vqut
    'Public LnMfg As Short = 1
    'Public LnDesc As Short = 2
    'Public LnProdID As Short = 3
    'Public LnLPMFG As Short = 4
    'Public LnLPDesc As Short = 5
    'Public LnLPProdID As Short = 6
    'Public LnSpecCrs As Short = 7
    'Public LnSPA As Short = 8
    'Public LnEntryDate As Short = 9
    'Public LnLastChgDate As Short = 10
    'Public LnLastChgBy As Short = 11
    'Public LnActive As Short = 12
    'Public LnComments As Short = 13
    'Public LnNoteLine As Short = 14
    'Public LnGot As Short = 15
    'Public LnLpCode As Short = 16
    'Public LnSeq As Short = 17
    'Public LnCode As Short = 18
    'Public LnBranchCode As Short = 19
    'Public LnCost As Short = 20
    'Public LnSell As Short = 21
    'Public LnComPer As Short = 22
    'Public LnUnitOverage As Short = 23
    'Public LnBkComm As Short = 24
    'Public LnPrc1 As Short = 25
    'Public LnPrc2 As Short = 26
    'Public LnPrc3 As Short = 27
    'Public LnPrc4 As Short = 28
    'Public LnPrc5 As Short = 29
    'Public LnPrc6 As Short = 30
    'Public LnPrc7 As Short = 31
    'Public LnPrc8 As Short = 32
    'Public LnPrc9 As Short = 33
    'Public LnPrc10 As Short = 34
    'Public LnLPCost As Short = 35
    'Public LnLPSell As Short = 36
    'Public LnLPQty As Short = 37
    'Public LnQty As Short = 38
    'Public LnOvgSplit As Short = 39
    'Public LnUM As Short = 40
    'Public LnID As Short = 41
    'Public LnQutID As Short = 42
    'Public LnComDol As Short = 43
    'Public LnESell As Short = 44
    'Public LnExtCost As Short = 45
    'Public LnExtComm As Short = 46

    'Variables to save Prt font settings 02-02-99
    'Public PrtFontNameSave As String
    'Public PrtFontSizeSave As Single
    'Public PrtFontBoldSave As Boolean
    'Public PrtFontItalicSave As Boolean
    'Public PrtFontUnderlineSave As Boolean
    'Public FF() As String '01-18-09
    Public Amt As Decimal '01-18-09
    Public CommAmt As Decimal
    Public Commpct As Decimal '01-18-09
    Public PrevLev1 As String
    Public CurrLev1 As String
    Public PrevLev2 As String
    Public CurrLev2 As String '02-09-09
    Public FixProfit As Decimal
    Public FixProfitPer As Decimal
    Public LampProfitPer As Decimal '02-09-09
    Public SortSeq As String '02-07-09
    Public SortDir As String '02-07-09
    Public SortCode As String '02-07-09
    Public PrimarySortSeq As String
    Public SeconarySortSeq As String '02-08-09
    Public TOTDSellFixture As Decimal
    Public GTSellFixture As Decimal
    Public TOTDCostFixture As Decimal
    Public GTCostFixture As Decimal
    Public TOTDProfitFixture As Decimal
    Public GTProfitFixture As Decimal
    Public TOTDLampSell As Decimal
    Public GTLampSell As Decimal
    Public TOTDLampCost As Decimal
    Public GTLampCost As Decimal
    Public TOTDCommAmt As Decimal
    Public TOTDCommPct As Decimal
    Public GTCommAmt As Decimal
    Public GTCommPct As Decimal
    Public RC As Integer '01-30-09 

    'Public FLN() As String ' Data Base Calls
    'Public QP() As String ' Data Base Calls
    'Public QKa() As String '"
    'Public QK As String '"
    'Public QL() As Short '"
    'Public DBL() As Short '"
    'Public FQ As Short '"
    Public QN As Short '"
    'Public QS As Short '"
    'Public OP As Short


    ' Keep in Alpha Seq
    Public AbortPrtFlag As Boolean ' Std Abort Print Job
    Public Cmd As String ' Std Command to main
    Public Zarg As String
    Public ZE As String
    Public FormResize As Short 'SEW 02-25-00
    Public TabStartFlag As Short 'SEW 02-25-00
    'Public XFactor As Single 'SEW 02-25-00
    'Public YFactor As Single 'SEW 02-25-00
    'Public DesignX As Short 'SEW 02-25-00
    'Public DesignY As Short 'SEW 02-25-00
    'Public InResize As Short 'SEW 02-25-00
    ' Public Printer As PowerPacks.Printing.Compatibility.VB6.Printer = New PowerPacks.Printing.Compatibility.VB6.Printer
    ' Public PrinterX As PowerPacks.Printing.Compatibility.VB6.Printer = New PowerPacks.Printing.Compatibility.VB6.Printer
    '

    Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Integer 'VB5 07-01-98 TSL
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As IntPtr) As IntPtr
    Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer '04-05-99 WNA
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Short, ByVal lpFileName As String) As Integer '01-18-07 jh
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer '01-18-07 jh
    Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Integer
    '
    Public Declare Function AbortDoc Lib "gdi32" (ByVal hdc As Short) As Short
    Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Short, ByVal fEnable As Short) As Short
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Short, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Short) As Short

    Public nRet As Integer
    Public nNewWidth As Integer
    Public widths() As Double '02-03-09

    'MySQL Globals '03-09-07 JH 
    'Public myConnectionString As String = "Database=saw8;Data Source=localhost;User Id=root;Password=saw987"
    'Public myConnection As MySqlConnection
    'Public myDataAdapter As MySqlDataAdapter
    'Public myds As dsSAW8       'DataSet Detail Tab (This is the one to Save To)
    'Public ds As dsSAW8         'DataSet LOOKUP TAB (Combined Table)
    'GLOBAL VARIABLES
    Public _rowHeight As Integer ' original row height
    Public _recSelWidth As Integer ' oringal record selector width
    Public _fontSize As Single ' original font size
    Public _colWidthstgQh As New ArrayList()
    Public _colWidthstgr As New ArrayList()
    Public _colWidthstg As New ArrayList()
    'tgQh
    'tgr
    'tg

    Public mDataView As DataView
    Public pRecordCountInt3 As Integer
    Public NameMySqlParameter As New MySqlParameter
    Public mMySqlCommand As New MySqlCommand
    Public ActiveFrm As String = ""
    Public fontSetting As Font
#Region "SELECT CODE VARS"  '02-13-13 WNA
    Public SelectCodeArray(50, 2) As String '07-29-09 jh
    Structure SelectCodeList '07-29-09 jh
        Public Header As String
        Public TypeName As String
    End Structure
    Public SCRecord As SelectCodeList '07-29-09 jh
#End Region


    Public Sub EnableOrDisable2(ByRef Enable As Short)
        If Enable = 0 Then
            'frmQuoteRpt.optOne.Enabled = False'frmQuoteRpt.chkSpecifiers.Enabled = False'frmQuoteRpt.chkNotes.Enabled = False'frmQuoteRpt.chkCustomerBreakdown.Enabled = False 'tag 10-15-97
            'frmQuoteRpt.chkMfgBreakdown.Enabled = False 'tag 10-15-97'frmquoterpt.optMfgCust.Enabled = False'tag 10-15-97'frmQuoteRpt.cmdFmtOK.Enabled = False 'frmQuoteRpt.cmdCancelRpt.Enabled = False
            'frmQuoteRpt.chkDisplay.Enabled = False 'frmQuoteRpt.chkCopytoExcel.Enabled = False  '12-03-98 WNA
        Else
            frmQuoteRpt.ChkSpecifiers.Enabled = True
            frmQuoteRpt.chkNotes.Enabled = True
            frmQuoteRpt.chkCustomerBreakdown.Enabled = True 'tag 10-15-97
            frmQuoteRpt.chkMfgBreakdown.Enabled = True 'tag 10-15-97
        End If
    End Sub

    Sub todolist() 'PunchList TODO      Instructions for Quote Realization Report
        'To Get All Specifiers to one MFG, Chk QuotTo:Manuf, Chk All Specifiers, Chk Only Quotes to One Mfg/Cust, Enter Mfg Code
        'To Get All Specifiers to one Customer, Chk QuotTo:Customer, Chk All Specifiers, Chk Only Quotes to One Mfg/Cust, Enter Customer Code
        'To Get One Mfg Only, Chk QuotTo:Manuf, On the next Screen After Select MFG Code - Enter MFg Code 

        ''Done **************************************
        '06-28-18 JH ADD GOT TO REALIZATION FOLLOW UP EXCEL REPORT
        '06-26-18 06-27-18  JH  ADD SPEC/CUST PROMPT ON QUOTE SUMMARY REPORT.  ADDS SEPERATE COLUMNS FOR EACH TYPE
        '             And Me.ChkSpecifiersCustInCols.Checked = True Then '06-25-18 06-27-18
        '01-19-18 JH  SERVERPATH NOT SETUP
        '12-05-17     spelling fix on specifier
        '05-01-17 JH  REALIZATION - ADD SOURCE QUOTE (LIGHT ABILITY)
        '02-21-17 JH  REALIZATION EXPORT TO EXCEL 02-20-17
        '01-31-17 JH  REALIZATION - ADD PROBABILITY TO GRID (FORM) ALREADY IN SQL B/C OF "INCLUDE ALL HEADER FIELDS" (light ability)
        '01-04-17 JH  NEW C1
        '12-07-16 JH  REALIZATION ON MFG, ERROR TMPSELLQ DOESN'T BELONG TO TABLE.  IT'S NOT IN THE SQL - ITS IN THE ONLY QUOTES WITH ONE MFG/CUST.  
        '             ALSO DOESN'T ASK INFLUENCER % QUESTION ON SPECIFIER REPORTS
        '10-17-16 JH  REALIZATION ONE MFG, INCLUDE SPECIFIER AND SPEC SLS PROMPT (TERRI MOORE SHARP)
        '10-17-16 WNA changed Max Length on Retrieval code field on product sales history report.
        '01-26-16 JH  REALIZATION NOT WORKING ON SEARCH STRING - ADDED IT TO THE SQL BUT KEPT IT IN SELECTHIT
        '11-30-15 JH  REMOVE AUTO SAVE TO \DATA FOLDER.  GOT SUPPORT EMAIL ABOUT "TOO MANY REPORT FILES...." - LET THEM SAVE IT WHEREEVER THEY WANT TO
        '11-11-15 WNA print SpecCross from header on realiztion reports.
        '10-08-15 JTC Fix Quote Summary with Include Blank Bid Dates Checked
        '10-08-15 10-06-15 JTC New Controls Version .67 All
        '09-28-15 JTC New C1 Controls .75
        '09-23-15 JTC Realization Totals Only set font to Consolas = Fixed length Set Back to prior setting
        '09-18-15 JTC Special Totals Only Format on these Reports If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True) Then
        '09-17-15 JTC fix smaller Grid Status field
        '06-30-15 JTC Add  and Q.Status <> '" & "NOREPT" & "'  to Realization Reports
        '06-29-15 JTC better 4 year message
        '06-26-15 JTC Fix Printing all years Spread Sheet by Year" Eliminate Balnk NCode & UCASE(NCODE) 
        '06-25-15 JTC Fix Both Spreadsheets for one Specifier Ucase NCODE
        '06-24-15 JTC Add option for "Spread Sheet by Year" 4 years with total of 4 tears
        '06-17-15 JTC add , OpenShare.Shared, OpenShare.Shared) to VSPEC.DAT'06-17-15 JTC add , OpenShare.Shared
        '06-01-15 JTC QuoteCode/JobName with SLS-1-14  MsgBox("Do you want Report in SLS-1-4 Sequence with One Salesman per page?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "SLS-1-4 Sequence with One Salesman per page option") 
        '05-28-15 Put status if First SQL If ForecastAllMfg = False And BCstatus <> "" Then SaveStrSQL += BCstatus 
        '05-28-15 Put Q.EstDelivDate or Q.EntryDate if First Temporary SQL in PrintReportQuote lines
        '05-28-15 JTC Fix Multiple MFgs in Select Line Items Not Forecasting Replace above BuildSQLDetail
        '05-28-15 JTC Customer Option Must Check Specifier or Customer In Line Items Reports Multiple Customer Codes("GES/AT,GES/BR")  in Select Line Items Not Forecasting BuildSQLDetail If rcode.Contains(",") = True Then 
        '05-28-15 JTC Multiple MFgs in Select MFG Line Items and Not Forecasting Replace Fix rcode = "KEEN,COOP,LOL,DAYB"
        '05-20-15 JTC Fix where Q.Status = 'ALL" Forecasting for MFGs Except Philips and SESCO If frmQuoteRpt.txtStatus.Text.ToUpper = "ALL" Then BC = "" 
        '05-20-15 JTC Alkways Save a Report even if No RepCustNumber If RepCustNumber Is Nothing Then GoTo NoFileName '05-14-15 JTC
        '05-15-15 JTC Loads BrandList & lets them Change on ForecastAllMFG
        '05-14-15 JTC Forecasting for MFGs Except Philips and SESCO If ForecastAllMfg = True Then Public 
        '05-08-15 JTC Help System Add Group Help - F1 
        '05-05-15 JTC Excel Fix Zero QuoteTo Amount on Contractors BOM print If ExcelQuoteFU = True Ask to Select on Quote Header Amout and Print it if Quote to Amt is Zero
        '05-01-15 JTC Added Help GoToMeeting
        '04-30-15 JTC ExcelQuoteFollowUP Added SLSQT from QT, Filter on Quote amount and Status
        '04-29-15 JTC Realization No other options allowed if ExcelQuoteFollowUp or SESCOJobListRealReport
        '04-28-15 JTC Public ExcelQuoteFU = True as Boolean switch 
        '04-27-15 JTC/JH add NameDetail.BUsinessType  ExcelQuoteFollowUp T-Contractor 
        '04-24-15 JTC Create ExcelQuoteFollowUp Public Sub from PrintSESCOJobListRealReportQutTO() "Excel Quote FollowUp"
        '04-22-15 JTC add "Excel Quote FollowUp" cboSortRealization.Items(12).ToString = "Excel Quote FollowUp"  when not "SESCO Job List Report" to "Excel Quote FollowUp" Realization
        '04-22-15 JTC Fix RealQuoteToAmtON = True  one MFG Use Realization QuoteTo Sell Amount on Report CompAmt = drQToRow("TMPSellQ")
        '04-20-15 JTC Fix Realization Salesman/Name Code/QuoteCode eliminate Duplicates
        '04-16-15 JTC Fix Real Customers to one MFG See Added QuotesToOneMFG = RealWithOneMfgCustCode
        '04-17-15 JTC Fix "QuoteCodeQ" error Export LU to Excel
        '02-04-15 JTC on Realization ExportTgLookupToExcel with one Mfg and AllSpecifiers skip filling TgLookup again
        '02-04-15 JTC Must be Upper case  BrandReportMfg = BrandReportMfg.ToUpper
        '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero" If Me.chkMfgBreakdown.Text = "T" Customer Only option ALS Print Quote Hdr Amt when QuoteTo is Zero" And Me.cboTypeCustomer.Text.Trim.ToUpper = "T" And drQToRow.BusinessType <> "T" And Me.chkMfgBreakdown.Checked = True Then ' "Add MFG Total Breakdown to Reports" 
        '01-12-15 JTC If run Realization B/4 Quote Summary then Me.ChkDetailTotalsOnly.Visible is False 
        '01-09-15 JH  ALS - ADD - DefLnCodes.Tables(0).Rows.Add("C", "Distributor/Customer") - ONLY WANT TO SEE THE DISTRIBUTORS (marked in NA) ON THE BELOW REPORT AFTER THEY GET THE CONTRACTORS
        '01-07-14 JH  ALS - REALIZATION - FIX BUSINSESS TYPE JOIN FROM NA TO GET CONTRACTORS, ETC QUOTED ON REALIZATION TAB (SAVED AS CUSTOMER ON QUOTE TO), THEY WANT QUOTE TO SLS (REALIZATION, BY CUSTOMER, MAJOR SORT SLS, SECOND JOB NAME)
        '12-31-14 JTC Fix Realization run after Quote Summary
        '12-31-14 JTC If OneMfg or OneCust Don't put FirmName in NCode
        '12-29-14 JTC 12-18-14 JTC Fix SLS1-4 Specifier Rpt If NCode is blank the SQL needs to Fix it 12-31-14 IF(PC.NCODE = '', Left(PC.FIRMNAME,6), pc.ncode) as SLSCODE2,
        '12-18-14 JTC Fix Else was deleting others If US = "M" Then If drQToRow.Typec = "M" And US = "M" Th
        '12-18-14 JH  FORECASTING REPORT - COMMENT OUT THE SKIP IF 0 IN SELL AND EXTENDED - LOT PRICE QUOTES GET SKIPPED, EST DEL DATE ON TAB B NOT RIGHT BECAUSE WE AREN'T CLEARING RT BEFORE THE SECOND LOOP
        '12-15-14 JTC drQToRow.Typec = "M" And US = "M" must be the same Realization one MFG or one Cust
        '11-21-14 12-01-14 JH  ADD JOB NAME TO LINE ITEM REPORT, FIX EXPORT TO EXCEL FOR THAT REPORT
        '             DON'T SAVE SHOW HIDE FILE EVERYTIME THEY RUN A REPORT, ONLY WHEN THEY CLICK SAVE.  MESSES UP WHEN YOU ADD CUSTOMER CODE TO LINE ITEM REPORT AND THEN RE-RUN IT W/O CUSTOMERS SELECTED WHERE THERE IS NO FIRMNAME OR CODE FIELD IN TRUEGRID       
        '11-20-14 JTC Blank out after SelectionText = ""  Product Sales History - Line Items" Status = " & rcode'11-20-14 JTC add SelectionText to Product Lines Report
        '11-20-14 JTC Fix add SLSQ selection to Product Sales History - Line Items show on Report Heading Selection
        '11-17-14 JTC Fix Realization Just (Mfg Or Just Cust) selected  If drQToRow.Typec = "M" And US = "M" Then Hit = 0 : drQToRow.Delete() : RowCnt -= 1 : Continue For '10-30-14 Don't Show Mfg Record Per Jaci DataRowState.Deleted
        '11-05-14 JTC Complie x86 and '11-05-14 JTC in FileGetUserID Added , OpenAccess.Read, OpenShare.Shared
        '11-04-14 JTC Add Quote Summary Rept (Job Name or Quote Code) Seconary Sort Option Salesman 1-4 Split Show Dollars at split percent FollowBy = Sls * percent
        '11-04-14 JTC If RealQuoteToAmtON = True For Mfg Delete C=Customer Records DataRowState.Deleted If drQToRow.Typec = "C" And US = "M" Then Hit = 0 : drQToRow.Delete() : RowCnt -= 1 : Continue For (Same if Mfg Records)
        '11-04-14 JTC Duplicate Major Sort on Job Name? Question GoTo SKipJobQuestion:  
        '10-28-14 JH  EXTERNAL COMPONENT ERROR (TERMINAL SERVICES ONLY, SERVER 2012) NRG - REMOVE C1 DATEPICKER
        '10-23-14 JTC Fix UserPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) 
        '10-14-14 JTC Added DoEvents to Disable CmdRunReport & CmdOK
        '10-13-14 JTC Disable CmdRunReport & CmdOK to Fix DblClick event GoTo EndExit 
        '10-13-14 JTC Realization By JobName Sequence first RealWithOneMfgCustSortJobName = True on one MFG cust All Quotes to One Mgf/Cust With Specifiers in Job Name Seq
        '09-30-14 JTC Fix Dates FillQutRealLUDataSet( Relization with one MFG Bad Date  
        '08-20-14 JTC Fix Realization Spreadsheet by Month for one MFG If (Row = MaxRow And SingleMFG <> "") Then TotYTDQ = 9999999 : GoTo PrintLoop2 ' = frmQuoteRpt.txtQutRealCode.Text '08-20-14 JTC If Not "ALL"
        '07-31-14 JTC Fix ExportTGLookuptoExcel not to reload Layout files If RealTgLookupExcel = False Then This is run After show hide and Loading Layout wipes out showhide column visible '07-31-14JTC Fix Header 
        '07-31-14 JTC Forecasting only for PHIL requires BrandReporting
        '07-31-14 JTC Fix Sesco Job List If cboSortRealization.GetItemCheckState(12) = CheckState.Checked Then '12 is "SESCO Job List Report"
        '07-24-14 JTC Fix SESCO Job Listing Excel Report to not show SecondarySeq
        '07-24-14 JH Fix Error on Col 26 ExportTGLoolupToExcel Dim ExcelCol As Short = 1 '= Chr(64 + PrtCols)
        '07-24-14 JTC Forecasting Report If SESCO = True then Sell=0 ExtSell= 0 & Unit Sell GtExtQlSell = 0 : ExtQlSell = 0
        '07-22-14 JTC Fix Header Rept-B Quotes By Brand the Header was Commented out in Error
        '07-22-14 JH  SPECIFIER, ARCHITECT, CONTRACTOR, ENGINEER, NOT SHOWING ON QUOTE TO, SHOW ALL HEADER DATA REPORT
        '07-17-14 JTC Fix One code MsgBox
        '07-16-14 JH Fix Dist VqrtRealQTOShowHide grid & Layout to show SLSCode
        '07-15-14 JTC RealWithOneMfgCustCode = "XXXXXX' OneNCodeOnly If Mfg QuoteTO show RealQuoteToAmtON = True
        '07-08-14 JTC Skip Tilda on Forecasting Quote Lines & Order Lines Delete All text after "~" Tilda  ZV = InStr(drQutLn("Description"), "~") 
        '06-28-14 JTC Skip SubTotal if Real & "Descending Sales Dollars" only Sequence If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.txtPrimarySortSeq.Text = "Descending Sales Dollars" And SortSeq = "projectcust.Sell" Then GoTo SkipSubTotal '06-28-14 JTC Skip SubTotal if "Descending Sales Dollars" only Sequence
        '06-13-14 JTC Fix Excel 2013 Version added s(1) need Sheet numberto first Sheet to fix Heading not showing on Forecast report
        '05-29-14 - HOLD ORDER MESSING UP EST DELIVERY DATE
        '05-28-14 JTC/JH Fix Estimated Delivery date and Report Header Date
        '05-15-14 JTC Error if Me.pnlTypeOfRpt.Text <> "Terr Spec Credit Report" Then Don't do Call frmShowHideGrid.ShowHideGridCol IE: Fixed Columns
        '05-09-14 JH Chg SpanCols No - 1 Quote Reports  RTS.Width = "auto" '05-09-14 JH Added
        '05-09-14 JTC Bold Company name RT.Cells(0, 1).Style.FontBold = True 
        '05-09-14 JTC Eliminate Dollars on QuoteTo for dist Report Quote summary if chkIncludeCommDolPer.Checked = CheckState.Unchecked
        '03-27-14 JH Fixed Grid Layouts and Deleted need for Adding Col If InStrColNam("FIRMNAME") Then Call AddTgColumns("FirmName", "Near", tgln)
        '03-24-14 JTC delete Time display msg Forecasting Brand
        '03-19-14 JH More Layout Changes & Relization Report
        '03-12-14 JH Fix ShowHideGridCol  
        '02-28-14 JTC Use TMPREPORTS1 on HOLDORDERS to speed up
        '02-28-14 JTC Added SourceQuote # to Forecast Report A in MFG Column If IsDBNull(drQutLn("SourceQuote")) Then drQutLn("SourceQuote") = "" 
        '02-28-14 JTC Added Job Name to Forecast Report HoldOrders
        '02-27-14 JTC Added Time MsgBox("Step1= " & Step1 & "  Step2= " & Step2 & "  Step3= " & Step3 & " Minutes by Step" & vbCrLf & "PLease view the report file. ") 
        '02-25-14 JTC Added mnuSupport"ENABLEBRANDREPORTING"
        '02-25-14 JTC Q Line Items Eliminate Dups on PC table
        '02-25-14 JTC If Q Line Items if txtCustomerCodeLine then check Me.chkSlsFromHeader.CheckState = CheckState.Checked
        '02-25-14 JTC Fix Quote Lines with One Customer code No PC.NCode in Where Clause Add PC.NCode as NCode, PC.FirmName as FirmName (add INNER JOIN PROJECTCUST AS PC ON Q.QuoteID = PC.QuoteID and (PC.Typec <> 'C' And PC.Typec <> 'M') )
        '02-24-14 JTC Added Shell(UserPath & "VNAME.EXE " & " " & "/OpenBrandTable", 1)  
        '02-24-14 JTC Move "BRANDREPORT-PHIL" From UserPath to UserSysDir
        '02-17-14 JTC Added Public DefChgEstDelDateCodes To "QUTDEFAU.DAT" Set in VQUT forecasting Private Sub SetDefaultEstDelDate() 
        '02-06-14 JTC Fix Changing MFG on "BrandReport-" & BrandReportMfg & ".DAT" from Menu, File
        '02-05-14 JTC Philips Forecasting Report Better Msg "This will only include Philips Brands and Dollars.
        '02-05-14 JTC Realization Added Get Many Mfg or Cust JONES,HERRY Realization FillQutRealLUDataSet
        '02-04-14 JTC Real to MFG add projectcust.NCODE = 'DAYB' and if one Mfg from frmQuoteRpt.txtQutRealCode.Text.ToUpper in FillQutRealLUDataSet
        '02-04-14 JTC OpenSQL set myConnectionString timeout=1800" '30 min for Reports ALR
        '02-03-14 JTC Can't Select SLSsplit and use SLS from Header
        '02-01-14 JTC Fix Realization When One Code selected  And OrderBy OS.SLSCode on Salesman Sequence
        '02-01-14 JTC OpenSQL set myConnectionString timeout=900" '15 min for Reports
        '01-30-14 JTC Fix Realization Null on drQToRow.SLSCode SubTotChk9360
        '01-27-14 JTC Set Excel Forecast WookSheets.count to 3 Need three to find out thecurrent number of worksheets and then call Workbook.Worksheets.Add() 
        '01-22-14 JTC Add REAL txtSecondarySort.Text = "Status" and txtSecondarySort.Text = "Bid Date"
        '01-21-14 JTC Realization with All Specifiers to one MFG REALWithOneMFGCust Has Export Lookup to Excel Option
        '01-21-14 JTC Realization by Salesmam by Jobname or NCode ALS   If frmQuoteRpt.txtSecondarySort.Text = "Job Name" Then
        '01-17-14 JTC Put Date Range in File Name Forecast Report 
        '01-14-14 JTC Off if No "BRANDREPORT-PHIL" file  Me.chkBrandReport.Visible = False 
        '01-14-14 JTC Turn Off Brand Except on Forecast Report Me.chkBrandReport.CheckState = CheckState.Unchecked 
        '01-06-14 JTC RealQuoteToAmtON = True  Use Realization QuoteTo Sell Amount to Select Quotes Question If Quote Amt entered Or Quote Header Amount
        '12-12-13 JTC Bypass Bid Date test in SelectHit9500 Forecasting If frmQuoteRpt.txtPrimarySortSeq.Text = "Forecasting" Then GoTo 9510'12-12-13 JTC Bypass Bid Date test in SelectHit9500
        '12-12-13 JTC PHIL Forecast Report Fix Status GOT or SUBMIT 
        '12-02-13 JTC Compile 
        '11-19-13 JTC Fix Missing BidDate Check error added If Me.chkBidJobsOnly.Text = "Delivery Date Jobs Only" Then 
        '11-17-13 JTC Align PO Number Left
        '11-15-13 JTC Added Philips Forecasting Message Added Two More Sheets to excel file
        '11-14-13 JTC Add Status Column 
        '11-07-13 JTC Fix Duplicate mnuBrandMfgChg_Click BrandReportMfg = "XXXX"
        '11-06-13 JTC added Menu options Me.mnuBrandReport.Enabled = False : Me.mnuBrandMfgChg.Enabled = False : Me.mnuBrandReport.Text = "Brand Reporting - Off"
        '11-03-13 JTC Include HoldOrder within the last 3 years
        '11-01-13 JTC No Forecast 
        '10-28-13 JTC Fix SecuritySubNew   110:    ToDayDT = Now.ToString("yyyyMMdd")  '10-28-13
        '10-23-13 JTC PhilipsForecast Report Quotes no hold orders yet A=total each quote B=total each brand/Desc on each quote
        '10-16-13 JTC If (No Security or SYSTEM=Admin) then allow any Branch Codes to be entered '10-16-13 JTC Added QL.BranchCode to Line Items
        '10-15-13 JTC  If A = "SecurityGroup" 'Admin = SYSTEM, BRANCH, REGIONAL, "SYSTEM" Then SecurityAdministrator = True = All Branches IE: Ignore BRANCH 
        '10-15-13 JTC "BRANCH" If BRANCH, GetBranchCode(UserID) 'REGIONAL, Then SecurityBrancheCodes = dsadmin.adminuser.Rows(0).Item("AdminBranches").ToString
        '10-14-13 JTC If A = "ReportQut" Then SecuritySubNew
        '08-13-13 JTC/JH Testing Only added Sub SecuritySubNew() & Public Sub FillAdminTables()
        '08-03-13 JTC Added CustCode and FirmName to Product Line Item Reports and chkShowCustomer chkUseSpecifierCode 08-06-13 & 08-07-13
        '08-02-13 JTC 07-23-13 JTC If frmOrderRpt.chkUseSpecifierCode.Checked = True Change ColHeader  "CustCode" to  "SpecCode"
        '07-25-13 JTC IsDBNull If IsDBNull(drQToRow("SellQ")) Then drQToRow.SellQ = 0 '
        '07-15-13 JTC Change Error Logic in ShowHide & TrueGridLayoutFiles
        '06-30-13 JTC Added File In Use to save TGLookup to Excel
        '06-28-13 JTC Realization If Turn All Specifiers On or off Leave other settings as they were
        '06-26-13 JTC Added Show Book Total TotPrt9250 On Product Sales History - Line Items" 
        '06-24-13 JTC Add Book to Rep Reports tg.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" 
        '06-17-13 JTC Add "1 SLS/Page" on "Salesman/Customer" Realization
        '05-21-13 JTC Fix Blank Bid Dates If frmQuoteRpt.chkBlankBidDates.CheckState = CheckState.Checked Then '05-21-13 deleted frmQuoteRpt.chkBlankBidDates.Visible = True And 
        '05-20-13 JTC Added Realization 12 Month Spreadsheet
        '05-16-13 JTC Added Realization When sub Sort is salesman they can change tobe Salesman Major Sequence
        '05-16-13 JTC Add PHIL Brand Reports Get Brands from N&A Notes Record Call GetPHILBrands(A)  Public BrandReport
        '05-14-13 JTC Added Quote Summary by RptEnteredBy = 17 'EnteredBy '05-14-13
        '05-01-13 JTC If Sls 1-4 then turn Me.chkSlsFromHeader.CheckState on .RptSalesman
        '05-01-13 JTC Chg >= to = Moved Down SortCode = Me.txtSlsSplit.Text VQRT2.RptMajorType.RptSalesman And Me.chkSlsFromHeader.CheckState = CheckState.Checked 
        '03-25-13 JTC Me.chkShowLatestCust.Visible on RealWithOneCode
        '03-24-13 JTC Me.chkShowLatestCust.Visible = False on Quote Summary Report
        '03-22-13 JTC Go back to Select Criteria Tab From Grid cmdBackViewGrid
        '03-22-13 JTC "Only Quotes with One MFG/Cust" 'If RealWithOneMfgCust = True Then RealWithOneMfgCust = False
        '03-20-13 JTC Fix Salesman Report chkSlsFromHeader SQL Order By 'Major NCode Sequence
        '03-20-13 JTC Fix Salesman Report chkSlsFromHeader SQL Order By
        '03-08-13 JTC Added Summary Rpt Salesman = "Use Quote SLS 1 Split for Salesman" Me.chkSlsFromHeader.Text Me.chkSlsFromHeader.CheckState = CheckState.Checked Realization uses "Use Salesman From Quote Header on Report"
        '03-05-13 JTC Added Search Desc Logic TxtSearchString on Product Lines reporting
        '02-22-13 JTC Add QS.SLSCode as SLS1 from Quote QUTSLSSPLIT Table Position 1 Realization 
        '02-13-13 JTC Fix Realization MarketSegment = drQToRow.MarketSegment  SpecCross = drQToRow.SpecCross Fix cboSpecCross to default file values
        '01-29-13 JTC & 01-28-13 JTC Realization subtotals & OrderBY Fix Descending Dollars
        '01-24-13 JTC Fix Blank Preview
        '01-24-13 JTC More Fixes tostop printing on frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked on DIST or Rep
        '01-23-13 JTC Fix No Cost Dollars on DIST Report If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" 
        '01-19-13 JTC Added PrtGrayScale to eliminate Color Printing  Added chkPrintGrayScale 01-18-13
        '01-06-13 JTC Added My.Settings.WholeDollars = Me.chkWholeDollars.Checked  My.Settings.AddCommas = Me.chkAddCommas.Checked My.Settings.AddDollarSign = Me.chkAddDollarSign.Checked  
        '01-06-13 Added 01-01-13 and above today
        '01-03-13 JTC Changed all "########0.00" to DecFormat  Didn't do anything with Hard Coded "########0"
        '01-02-13 DecFormat As String = "########0.00" '01-01-12 NoCents "########0") DecFormat for Nocents Whole Dollars Commas = "###,###,##0"
        '01-02-13 chkWholeDollars.Checked = True Then DecFormat = "########0" Else DecFormat = "########0.00" '01-02-12 DecFormat As String = "########0.00" '01-01-12 NoCents "########0") DecFormat for Nocents Whole Dollars
        '01-02-12 If chkAddCommas.Checked = True Then If chkWholeDollars.Checked = True Then DecFormat = "###,###,##0" Else DecFormat = "###,###,##0.00" '01-02-12 
        '01-02-13 JTC Fix DTPicker1EndEntry.Value = "Last Month" showing 2013
        '01-01-13 JTC Public DecFormat As String = "########0.00"=Normal, NoCents="########0") DecFormat for Nocents Whole Dollars
        '12-13-12 JTC on RenderTable add RT.CellStyle.Padding.Left = "1mm"  RT.CellStyle.Padding.Right = "1mm" 
        '12-03-12 JH - DELTED BINDING FROM CBOTYPEOFJOB (WAS QUOTEBINDINGSOURCE-TYPEOFJOB) AND DELETED QUOTEBINDINGSOURCE FROM FORM
        '11-28-12 JTC 1-Fix All Option on realization 2-Add Book Column to Realization
        '11-01-12 JTC Better Error Msg on SESCU JOB Report if open in Excel
        '10-31-12 JTC Change P.MarketSegment to Q.MarketSegment
        '10-31-12 Turn Off Unless  If Me.txtPrimarySortSeq.Text = "SESCO Job List Report" Then  Else SESCO = False 
        '10-30-12 JTC No Secondary Selection on Salesman / Customer Realization
        '10-30-12 JTC Fix Branch Selection BranchReporting If Trim(BranchCodeRpt) <> "" And Trim(BranchCodeRpt) <> "ALL" Then
        '10-29-12 JTC Fix Retrieval Code Sequence Select Error the second time
        '09-28-12 JH New Icon
        '09-25-12 JTC fix garbage in background on  Me.fraSortPrimarySeq.BackColor = Color.White '09-25-12
        '09-23-12 JTC Pai& Unpaid Totals 
        '09-21-12 JTC Addeed ProjectName anr Paid Unpaid totals
        '09-20-12 JTC If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Fix  "Salesman Follow-Up Report")
        '09-19-12 JTC If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then '09-19-12  GoTo QutLineHistoryRpt
        '09-10-12 JH Fix Excel from converting Quote Code to Date format 12-0123 in ExportTgLookupToExcel objSheet.Columns(ExcelCol).NumberFormat = "@" 
        '09-09-12 JTC Fix Help F1 Me.HelpProvider1.HelpNamespace = UserPathHelp & "HLPQuoteReports.pdf" '09-09-12 
        '09-06-12 JTC Fix SLS/Desc Dollar frmQuoteRpt.ChkSpecifiers.Text used for Descending Dollars also ChkSpecifiersSkip: If frmQuoteRpt.ChkSpecifiers.Text = "Sort Report by Descending Dollar" Then GoTo ChkSpecifiersSkip 
        '08-28-12 JTC Shut off Realization choices when you click on summary report and other Non Realization Reports
        '08-28-12 JTC Fix Null on Error Realization"drQToRow.SellQ = 0 
        '08-28-12 JTC Add , quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, to SQL on Descending Dollar fix SellQ problem
        '08-27-12 JTC Select Realization on Total Quote Amount not sell of each row drQToRow.Sell CompAmt = drQToRow.SellQ
        '08-14-12 JTC Report Quotes with no Specifiers If frmQuoteRpt.ChkQuoteNoSpecifiers.CheckState = CheckState.Checked Then 
        '08-13-12 JTC Fix ChkSpecifiers.Text = Blank once in awhile
        '08-08-12 JTC Realization Moved Up From Below and fixed pnlQutRealCode.Text too long (Fix SalesMan/Customer Report)
        '08-03-12 JH  VS 2010
        '08-01-12 JTC If they load Original then delete the old ShowHide in order to see new columns if any
        '07-27-12 JTC Fixed Blank Name Code, Put Left(Firmname,8)
        '07-27-12 JTC Added Specifier totals only ib Descending Dollar Group BY Added and quote.Status <> 'NOREPT' 
        '07-26-12 JTC Added Secondary selection on Realization reports
        '07-06-12 use A,E,L,S,T,X  If RealExtByInfluencePercent = True then Extend by % on Realization Reports by specifier
        '07-06-12 JTC Better Error messages.
        '06-29-12 JTC No Specifiers on Quote Summary they are on RealizationfrmQuoteRpt.cboSortPrimarySeq.Items.Add("Specifier Credit")
        '06-22-12 JTC Fix Spec Credit left this label wrong Should be Me.lblJobName.Text = "Job Name Search String" '06-22-12 
        '06-14-12 JTC Better about
        '06-05-12 06-06-12 JTC Fix If frmQuoteRpt.chkBlankBidDates.CheckState = CheckState.Checked Then SESCOJobListRealReportQutTO Sesco = True 
        '05-21-12 Fix Object error & Select DidDatesDebug.Print(drQRow.QuoteCode) SelectHit9500
        '05-11-12 JTC Moved BidTime Next to BidDate in TG
        '05-04-12 JTC Again Fix Hit Routine Only Check Bid Date
        '05-04-12 JTC OpenShare.Shared FileClose(3) : FileOpen(3, FileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared) : If Not EOF(3) Then myConnectionString = LineInput(3) : FileClose(3) 
        '04-30-12 JTC Fix Hit Routine Only Check Bid Date If frmQuoteRpt.ChkCheckBidDates.CheckState = CheckState.Checked 
        '04-30-12 Set TG.AllowAddNew to False
        '04-26-12 Fix object Reference ErrorIf frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then '04-26-12 Fix object Reference Error
        '04-25-12 Increase size of Quote# RT.Cols(PC).Width = ".7in" 'FollowBy .7 field on RptMajorType.RptFollowBy And SESCO = True
        '04-20-12 WNA fix ExportTgLookupToExcel - was exporting only the last line in the grid. added Market Segment column to header grid.  make sure LPCost will not display in rep version
        '04-15-12 JTC Better ToolTop QuoteTo Salesman
        '04-13-12 JTC Fix Deleted Row Problem on SlsStatusFoll If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ" : GoTo 9365 
        '04-04-12 JTC Fix Select EntryDate&Biddate&InvludeBlankBid If drQRow.IsBidDateNull = True Then GoTo 9510 'OK on BidDate  and chkBlankBidDates.CheckState = CheckState.Checked
        '03-29-12 JTC Added Show just One-Latest Cudtomer QuoteTo .chkShowLatestCust.CheckState = CheckState.Checked Then 
        '03-29-12 JTC Don't Print Specifiers if NCode & FirmName Blank If drQPCRow.NCode.Trim = "" And drQPCRow.FirmName.Trim = "" Then GoTo NoHitNext 
        '03-23-12 JTC Added Width to Export to Excel
        '03-22-12 JTC Added Status to Sesco Report
        '03-19-12 JTC Fix Wording Exclude Quotes With Status = NOREPT
        '03-19-12 JTC Added Export Lookup Grid To Excel Call ExportTgLookupToExcel(Me.tgItem)  TByRef tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid) 
        '03-14-12 JTC Salesman Secondary total on EOF Fix Added $ signs and % signs to report
        '03-13-12 JTC Heading Width Fixed Regular Salesman Seq tof Fixed
        '03-12-12 JTC SESCO Bold & Centered 
        '03-09-12 JTC SESCO Special Report 03-10-12 03-11-12 
        '03-06-12 Always print SLSCode on QuoteTO  Only One Didn't per Job Seq Follow/SelectCode/BidDate
        '03-03-12 JTC Added Sesco Report FollowBy Seq, One per page, Ability to select One FollowBy
        '03-01-12 JTC Added on For Each dr If drQNRow.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
        '02-26-12 JTC Add SESCO Job List to Realization 
        '02-22-12 to '02-25-12 PrintSESCOJobListRealReportQutTO() '02-22-12
        '02-11-12 JTC ReUse ChkSpecifiers.Text = "Sort Report by Descending Dollar 
        'If Me.ChkSpecifiers.Text = "Sort Report by Descending Dollar" And Me.ChkSpecifiers.CheckState = CheckState.Checked Then '02-11-12 " "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
        '    '"Yes = Sort by Name Code / Descending Sales Dollars or No = Just Descending Sales Dollars"
        '    Resp = MsgBox("Yes = Sort by Name Code / Descending Sales Dollars or " & vbCrLf & "No = Just Descending Sales Dollars", MsgBoxStyle.Question + MsgBoxSty
        '02-11-12 JTC Add Notes to Realization 'If Me.pnlTypeOfRpt.Text = "Realization" Or Me.chkNotes.CheckState = CheckState.Checked Then
        '02-08-12 JTC No Column On Realization Fix ("Comm-$").Visible = True error
        '02-07-12 JTC Fix Summary Rpt Include Customer or Mfg Comm Amt was wrong
        '02-07-12 JTC Set Form.AutoScaleMode to None to fix lower resolutions problems
        '02-06-12 JTC Fix Dist Ver
        '02-05-12 JTC Extend by Specifier Inflence
        '02-04-12 JTC Fix Header on First Page Col hdg on All Pages
        '02-04-12 JTC RealExtByPercent
        '02-03-12 JTC Turn off Cost on Rep Realization
        '01-31-12 JTC Fix Dist QuoteTo  added - Click Run Report to see final report selection to Grid Caption
        '01-30-12 JTC/JH Fix Primary Sort Seq after running Realization Report
        '01-20-12 JTC Fix Tips & Decimal Places
        '01-20-12 JTC Added ChkBoxes for selection
        '01-18-12 JTC Realization Format n2 and Fix subtotals
        '11-29-11 JTC Chg Typec From X to O=Other on Specifier Typec
        '11-23-11 JTC Added SpecCross to Report & Chg quote.JobName Not project.ProjectName
        '11-17-11 JTC Fix Sub totals not Printing MsgBox("The Sell column needs to be in the sixth column" & vbCrLf & "or greater to Print totals correctly." & vbCrLf & "Click Yes tostop and move the Sell to the right (or add columns on the left)." & vbCrLf & "Click No to run the report as is.", MsgBoxStyle.YesNoCancel, "Sell column totals print") 
        '11-10-11 JTC/JH Turn Off Filter Bar tg.FilterBar = False Because I need all records for Report
        '09-16-11 JTC Tool tip For Multiple specifiers Chg From Single Code to Select Codes
        '09-01-11 JTC New C1 Controls
        '08-22-11 Fix Realization O = Other Report
        '08-21-11 Fix Realization to Let user Add Arch,Eng,Etc to report Fix height so Run Report Command would Show
        '08-20-11 added quote.Sell as QSell, quote.Cost as QCost, quote.Comm as QComm, to put on Specifier Value
        '07-21-11 JTC Report Only Bid Jobs If chkBidJobsOnly.Checked = True Then 
        '07-09-11 JTC Fix Seperate to Separate Spelling
        '06-13-11 JTC Quote Summary with Arch Etc and had Deleted quotes they are already deleted from truegrid Don't need a movenext
        '06-06-11 JTC On Product Lines Added chkPrtNTElines and chkHaveMFGCode.Checked = True Then  Must Have MFG Code
        '06-06-11 JTC On Realization Dim ShowAllQuoteHeader As String = ""  "Show All Quote Header Fields"
        '05-30-11 JTC VQRTSESCOJOBLIST.DAT REPORT 
        '05-26-11 JTC .C1SuperTooltip1. replaces Tooltip1. Fixed edits on txtStartAmt
        '05-19-11 JTC Had SelectHit9500(Hit, multsrtrvs) turned off by mistake????
        '05-18-11 JTC Realization Other TypeC
        '05-17-11 JTC Added Dates to Realization Report Added ProjectCust typeC 
        '04-20-11 JTC Fix Cmd Run Report Line Items
        '04-18-11 Turn Off SplashScreen '04-19-11 Date Tool Tip Press Tab
        '04-15-11 Run Report(Print/Preview/Export) Added
        '03-09-11 Jtc Error if Sell = -0.01 + 0.0001 to 0.01 Decimal MarginOrCommCalc(decimal)
        '02-06-11 If .xml not in UserDir & CurrXMLFile copy from UserSysDir also ShowHideFile
        '12-23-10 Fix TrueGridLayoutFiles If ORIG Get From UserSysDir not UserDir
        '11-30-10 JTC Quote Shortage Fix and Sort By Job Name
        '11-20-10 JTC Added SLS OptionEnterDate Descending Sell & EntryDate
        '11-19-10 JTC Fix PageBreak
        '11-11-10 JTC Fix Too Many Blank lines When Hit = 0 
        '11-04-10 JTC jh Fix Heading extend to right and RTS Sub RT object
        '11-01-10 JTC Set B/4  InitializeComponent() Report Run in "fr-FR" Frank Canada System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-us") '11-01-10 JTC Culture
        '10-25-10 JTC Added Chk Box To Print Blank line after each Quote
        '10-18-10 JTC If StartBidDate = "#1/1/1900#" then ignore BidDate  'Grid date Dropdown to False
        '10-17-10 JTC Bold Totals print  Fix Rep Comm-% calulation
        '10-15-10 JH  Bid date fix for Canada
        '10-12-10 JTC Error Msg Go to the Left Tab to start Reports Process. Select a Type of Report First") : Me.tabQrt.SelectedIndex = 0 : Exit Sub '10-12-10 
        '09-15-10 Fix Dist Reports Added JobNameF
        '09-10-10 Blank Bid Date If ColName = "BidDate" And ColText = "01/01/00" Then RT.Cells(RC, PC).Text = " " '09-10-10 Blank Bid Date
        '09-09-10 Don't Print Dollars on Specifiers  'QuoteTo Records are M,C,O  Specifiers are A,E,S,T,X
        '09-09-10 JTC Deleted SlsSplit in SQL beCause the report will show job twice on Me.txtSlsSplit.Text
        '09-07-10 JTC Fix ShowHide when you move columns around  Don't Show Password
        '09-04-10 UserID on Title
        '09-03-10 Show hide to front Sub CmdShowColstoPrt
        '05-06-10 JTC Shortage Only on Q type and add project Name
        '02-06-10 JTC Quote Summary was too slow fixed lookup other tables
        '12-10-09 JTC Fixed Various Sub totaling Etc.
        '12-04-09 JTC Dist and REP report layouts Added Planned Project & Spec Credit
        '11-12-09 JH Fix TrackBar Zoom for grids 
        '11-03-09 JTC Fixed Screen for high Resolutions
        '11-03-09 "1-1-1900" to "#1/1/1900#"
        '10-12-09 JTC added Extend by Probability 
        '07-28-09 JTC Delete file if Exist If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09

    End Sub
    Public Sub SetPrtHeadings(ByRef SlsHdg As String, ByRef StatHdg As String, ByRef AmtHdg As String, ByRef CommHdg As String, ByRef CommPerHdg As String)
        'Dim CheckState As Short '08-31-01
        If frmQuoteRpt.chkIncludeCommDolPer.CheckState = Windows.Forms.CheckState.Checked Then '12-09-02 WNA
            SlsHdg = " SLS  " : StatHdg = "STATUS" : AmtHdg = "AMOUNT" : CommHdg = "COMM$" : CommPerHdg = "COMM%" ' REP 08-31-01
        Else
            SlsHdg = " SLS  " : StatHdg = "STATUS" : AmtHdg = "AMOUNT" : CommHdg = "" : CommPerHdg = "" ' REP 08-31-01
        End If
        If MFG Then '05-09-01 WNA
            SlsHdg = " STAT " : StatHdg = " REP  " : AmtHdg = " SELL " : CommHdg = "COST " : CommPerHdg = "MARG%" 'MFG 08-31-01
            If DAYB Then '08-28-01
                SlsHdg = " REP  " : StatHdg = "STATUS" : AmtHdg = " SELL " : CommHdg = "COST " : CommPerHdg = "MARG%" ' DAYB 12-06-01 Reversed SELL & COST Hdg
            End If
        End If
    End Sub

    Public Sub TotalsCalc(ByRef A As String, ByRef B As String, ByRef C As TotalLevels)
        On Error GoTo 99997
        Dim I As Short = 0 '09-10-09 
        'Public A As String, FixSell As Decimal, FixProfit As Decimal, FixProfitPer As Decimal, LampSell As Decimal, LampCost As Decimal, Amt As Decimal, CommAmt As Decimal, Commpct As Decimal
        '01-25-09 TotalLevels.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
        If A = "AddAllLevels" Then 'Add to all levels FixSell FixCost LnQuantityA
            'Only Two levels in this progran'SellFixtureA(I),CostFixtureA(I),SellFixtureAExt(I)
            For I = 0 To 2 : SellFixtureA(I) = SellFixtureA(I) + FixSell : Next
            For I = 0 To 2 : CostFixtureA(I) = CostFixtureA(I) + FixCost : Next
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then '09-21-12  GoTo QutLineHistoryRpt
            Else
                For I = 0 To 2 : CommAmtA(I) = CommAmtA(I) + CommAmt : Next '12-09-09
            End If

            For I = 0 To 2 : RealizSellAExt(I) = RealizSellAExt(I) + RealizSellExt : Next '12-10-09
            ' Public RealizSellExt As Decimal '12-10-09    Public RealizSellAExt(5) As Decimal
            'SellFixtureAExt, CostFixtureAExt
            For I = 0 To 2 : SellFixtureAExt(I) = SellFixtureAExt(I) + FixSellExt : Next '09-11-09
            For I = 0 To 2 : CostFixtureAExt(I) = CostFixtureAExt(I) + FixCostExt : Next ' 'Rep Ext Comm to FixCostExt '02-11-10 
            For I = 0 To 2 : ProfitFixtureA(I) = ProfitFixtureA(I) + FixProfit : Next
            For I = 0 To 2 : QuantityA(I) = QuantityA(I) + LnQuantityA : Next
            '12-09-09 Below ON
            For I = 0 To 2 : LampSellA(I) = LampSellA(I) + LampSell : Next '12-09-09 
            TOTDLampSell = TOTDLampSell + LampSell
            For I = 0 To 2 : LampCostA(I) = LampCostA(I) + LampCost : Next
            TOTDLampCost = TOTDLampCost + LampCost
            For I = 0 To 2 : ProfitLampA(I) = ProfitLampA(I) + FixProfit : Next
            '12-09-09 For I = 0 To 1 : CommAmtA(I) = CommAmtA(I) + LampCost : Next
        End If
        If A = "ZeroLevels" Then 'Add to all levels
            F = C 'Zero These totals and all lower totals
            For I = F To 5
                QuantityA(I) = 0 : ProfitFixtureA(I) = 0
                SellFixtureA(I) = 0 : CostFixtureA(I) = 0
                SellFixtureAExt(I) = 0 : CostFixtureAExt(I) = 0
                LampSellA(I) = 0 : LampCostA(I) = 0 : ProfitLampA(I) = 0
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then '09-21-12  GoTo QutLineHistoryRpt
                Else
                    CommAmtA(I) = 0 '12-09-09
                End If
                RealizSellAExt(I) = 0 '12-10-09 
                RealizSellExt = 0 '12-10-09 
                TOTDLampSell = 0 '12-09-09 TOTDLampSell + LampSell
                TOTDLampCost = 0 '12-09-09 TOTDLampCost + LampCost
                LnQuantityA = 0 '06-14-10 

            Next

        End If
        'If DAYB Then '11-26-01
        '    GTCommPct = ((GT - GTCommAmt) / (GT + 0.0001)) * 100 '11-26-01
        '    TOTDCommPct = ((TOTD - TOTDCommAmt) / (TOTD + 0.0001)) * 100 '11-26-01
        GoTo 99998
99997:  'On Error GoTo 99997:
ErrBox: Dim Msg As String
        Msg = "VB Error # = " & Str(Err.Number) & "  ERROR AT " & Str(Erl()) & vbCrLf & "TotalsCalc - PLEASE READ BACK OF MANUAL" & vbCrLf & "ERROR MESSAGE SECTION Z-2"
        Dim resp As Integer
        resp = MsgBox(Msg & vbCrLf & ErrorToString(Err.Number), MsgBoxStyle.OkCancel, US)
        If resp = MsgBoxResult.Cancel Then Resume 99998 ' GoTo Msub_Exit: 'Cancel
        Resume 99998 'Return
99998:
    End Sub


    Public Sub SetYearEnd()
        Dim YearToUse As String
        Dim CurrentYear As String
        Dim CurrentDate As String
        Dim EDate As String
        Dim SDate As String
        Dim VB6 As Short '01-10-06 JH
        'New YEAR END FORMAT Sub.  Gets the Current Date & Year and Compares it
        'to 01-16 of the current year.  If it's past 01-16 of the current year
        'update the Date Values for the New Year.  It not, leave them at the Previous year
        'Invoice Reports = 02-01
        'Other   Reports = 01-16
        On Error Resume Next
        SDate = "0101" & Format(Now, "yyyy")
        EDate = "1201" & Format(Now, "yyyy")
        CurrentDate = Format(Now, "MM-dd-yyyy") 'yyyyMMdd"
        CurrentYear = Format(Now, "yyyy")
        If CurrentDate >= "01-16-" & CurrentYear Then
            YearToUse = Format(Now, "yyyy")
        Else
            YearToUse = CStr(Format(Now, "yyyy") - 1)
        End If
    End Sub
    Public Function GetRepNumber(ByVal MFG As String) As String
        Dim FatalExceptionCnt As Boolean = False
Start:  Try
            Dim tmpds As dsSaw8 : tmpds = New dsSaw8 : tmpds.EnforceConstraints = False
            Dim mdaND As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            mdaND.SelectCommand = New MySqlCommand("Select * from namedefaults WHERE NCode = " & "'" & SafeSQL(MFG) & "' and RecType = '" & "REP#" & "'", myConnection)
            mdaND.Fill(tmpds, "namedefaults")
            If tmpds.namedefaults.Rows.Count <> 0 Then
                Return tmpds.namedefaults.Rows(0)("CATEGORY")
            Else
                Return ""
            End If
        Catch ex As Exception
            If ex.Message.Contains("Fatal error encountered during command execution") = True Then
                'Leaving order system open for a long time and then coming back later on and trying to LU
                If FatalExceptionCnt = False Then Call OpenSQL(myConnection) : FatalExceptionCnt = True : GoTo Start Else GoTo DFLTMSG '05-31-12
            Else
DFLTMSG:        MessageBox.Show("Error in GetNameDefaults" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return Nothing
        End Try

    End Function
    Public Function GetEDIOrderRecord(ByVal NameCode As String) As dsSaw8.namedefaultsRow
        Dim FatalExceptionCnt As Boolean = False
Start:  Try
            'STORED PROCEDURES / VIEW NAME DEFAULTS
            Dim myds As New dsSAW8
            myds.EnforceConstraints = False
            NameMySqlParameter = New MySqlParameter 'Reset the SQL Parameter
            mMySqlCommand = New MySqlCommand 'Reset the Command Object
            With mMySqlCommand
                .Connection = myConnection
                .CommandType = CommandType.StoredProcedure
                .CommandText = "GetNameDefaultEDIOrders"
                'SQL Parameter info - Look at the stored Procedure for reference
                With NameMySqlParameter
                    .ParameterName = "parCode"              ' 
                    .Direction = ParameterDirection.Input
                    .MySqlDbType = MySqlDbType.VarChar
                    .Size = 8
                    .Value = SafeSQL(NameCode)
                End With
                .Parameters.Add(NameMySqlParameter)
            End With

            Dim mydataadapterND As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter(mMySqlCommand)
            mydataadapterND.Fill(myds, "namedefaults")
            If myds.namedefaults.Rows.Count > 0 Then
                Return myds.namedefaults.Rows(0)
            Else
                Return Nothing
            End If

        Catch ex As Exception
            If ex.Message.Contains("Fatal error encountered during command execution") = True Then
                'Leaving order system open for a long time and then coming back later on and trying to LU
                If FatalExceptionCnt = False Then Call OpenSQL(myConnection) : FatalExceptionCnt = True : GoTo Start Else GoTo DFLTMSG '05-31-12
            Else
DFLTMSG:        MessageBox.Show("Error in GetEDIOrderRecord" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            Return Nothing
        End Try

    End Function
    Public Sub GetPHILBrands(ByRef A As String) '05-16-13 
        '05-16-13 JTC Add PHIL Brand Reports Get Brands from Notes Record Call GetPHILBrands(A)
        'SELECT * FROM saw8sesco.namedefaults where NCode = '999999' and RecType = 'EDI' and Category = 'ORDERS'
        'Comments = 'ZZ|MULTIMICRO SYST|ZZ|606449916|THMI=LAM=MORL=DAYB=MCPH=OMEG=CAPR=THMO=CHLO=GUTH'
7988:   Try
            Dim dtND999 As dsSaw8.namedefaultsRow = GetEDIOrderRecord("999999")
            If Not dtND999 Is Nothing Then '10-05-11
                B = InStrRev(dtND999.Comments, "|")
                Dim EDI1 As String = UCase(Mid(dtND999.Comments, B + 1))
                A = UCase(Mid(dtND999.Comments, B + 1))
                A = Replace(A, "=", ",")
            Else

            End If
            GoTo Exit_Done
ValidMfg_End:
        Catch ex As Exception
            MessageBox.Show("Error in GetPHILBrands" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
Exit_Done:
    End Sub
    Sub CheckMKDir(ByRef FilePath As String)
        Static A As String '01-15-02 WNA
        'Use Dir function to check if directory exists.
        On Error GoTo ErrNoDir
        A = Dir(FilePath, FileAttribute.Directory) 'Directory does not exist.
        If Len(A) = 0 Then
            MkDir(FilePath)
            Exit Sub
        End If
        Exit Sub
ErrNoDir:
        On Error Resume Next
        MkDir(FilePath)
    End Sub


#Region "User Info"
    Public Sub GetUserInfo()
        Dim FDI As Short
        Dim A As String '07-28-06 WNA
        Dim AGMSTR(20) As String
        Dim fn As Integer = FreeFile()
        Try
            Dim drnd As dsSaw8.namedetailRow = GetCustomer("999999")
            If drnd Is Nothing Then
                A = UserSysDir & "AMASTER.DAT"
                Dim fileExists As Boolean = CheckForFile(A, False)
                If fileExists = True Then
                    'FileOpen(5, FileNM, OpenMode.Input, OpenAccess.Read, OpenShare.Shared) '11-05-14 JTC in FileGetUserID Added , OpenAccess.Read, OpenShare.Shared
                    FileClose(fn) : FileOpen(fn, A, OpenMode.Input, OpenAccess.Read, OpenShare.Shared) ''11-05-14 JTC in FileGetUserID Added , OpenAccess.Read, OpenShare.Shared
                    For FDI = 1 To 20
                        If EOF(fn) Then Exit For
                        AGMSTR(FDI) = LineInput(fn)
                    Next
                    AGnam = AGMSTR(1) : FileClose(fn)
                Else
                    AGnam = ""
                End If
            Else
                AGnam = drnd.FirmName
            End If

            Call FileGetUserID(UserDocDir & "UserID.DAT", UserID, "Read") 'Read Write '05-27-09
            If Trim(UserID) = "" Or UserID = "none" Then UserID = "ZZZ"

        Catch exc As Exception
            MsgBox(exc.Message, , "Get User Information")
            If Trim(UserID) = "" Or UserID = "none" Then UserID = "ZZZ"
        End Try


    End Sub
    Public Sub FileGetUserID(ByVal FileNM As String, ByRef Str As String, ByVal Func As String)
        '11-06-09 JH
        Try '07-01-09 UserDocDir & UserID.DAT  FileNM = UserDocDir & FileNM

            Dim UserIDin As String = "" '= LineInput(5)
            Dim ToDayDT As String = VB.Format(Now, "yyyyMMdd")
            Dim ToDayDtin As String = "" '07-01-09

            If Func = "Read" Then
                FileClose(5) : FileOpen(5, FileNM, OpenMode.Input, OpenAccess.Read, OpenShare.Shared) '11-05-14 JTC in FileGetUserID Added , OpenAccess.Read, OpenShare.Shared
                If EOF(5) Then FileClose(5) : GoTo FileGetUseIDExit
                If EOF(5) = False Then UserID = LineInput(5)
                FileClose(5)
            End If
            If Func = "Write" Then
                FileClose(5) : FileOpen(5, FileNM, OpenMode.Output)
                PrintLine(5, UserID)
                'PrintLine(5, UserName)
                FileClose(5)
            End If
FileGetUseIDExit:
            FileClose(5)
        Catch e As Exception
            MsgBox(e.Message, , "FileGetUserID")
        End Try
    End Sub
#End Region

    Public Function GetCustomer(ByVal Code As String) As dsSaw8.namedetailRow
        Dim strsql As String = ""
        Dim myds As dsSaw8 : myds = New dsSaw8
        myds.EnforceConstraints = False
        Dim drn As dsSaw8.namedetailRow
        Try
            strsql = "Select * from NameDetail where Code = '" & SafeSQL(Code.Trim) & "'"
            Dim mydataadapterProject As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            mydataadapterProject.SelectCommand = New MySqlCommand(strsql, myConnection)
            Dim cbP As MySql.Data.MySqlClient.MySqlCommandBuilder = New MySqlCommandBuilder(mydataadapterProject)
            mydataadapterProject.Fill(myds, "NameDetail")
            If myds.Tables("NameDetail").Rows.Count <> 0 Then
                drn = myds.Tables("NameDetail").Rows(0)
                Return drn
            Else
                Return Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message & " in GetCustomer")
            Return Nothing
        End Try
    End Function
    Public Function GetBranchCodes(ByVal datatable As String, Optional AndOrWhere As String = " and ") As String

        Dim strsql As String = ""
        Try   '10-15-13
            'Admin = SYSTEM, BRANCH, REGIONAL,
            '"SYSTEM" Then SecurityAdministrator = True = All Branches IE: Ignore BRANCH
            '"BRANCH" If BRANCH, GetBranchCode(UserID) 
            'REGIONAL, Then SecurityBrancheCodes = dsadmin.adminuser.Rows(0).Item("AdminBranches").ToString
            If SecurityAdministrator = True Or SecurityLevel = "SYSTEM" Then BranchSql = "" : frmQuoteRpt._fdBranchCode.Text = "ALL" : GoTo GetBranchCodeExit '10-20-13 Exit Function

            'If dsadmin.adminuser(0).Item("Admin").ToString <> "NORMAL" And dsadmin.adminuser(0).Item("Admin").ToString <> "SYSTEM" Then
            '    If IsPriceLU = False Or dsadmin.admingroup(0).Item("ViewPricesAllBranches").ToString = "N" Then
            Dim BranchCode As String
            If SecurityLevel = "BRANCH" Then ' dsadmin.adminuser(0).Item("Admin").ToString = "BRANCH" Then
                BranchCode = GetBranchCode(UserID)
                BranchSql = AndOrWhere & " (" & datatable & ".BranchCode = '" & BranchCode & "' or " & datatable & ".BranchCode = '') "
            Else
                BranchCode = dsadmin.adminuser(0).Item("AdminBranches").ToString
                If BranchCode <> "ALL" Then
                    Dim BC As String = ""
                    If SecurityBrancheCodes.Contains(",") = True Then
                        BranchSql = AndOrWhere & " ( " & datatable & ".BranchCode = '" & SecurityBrancheCodes.Replace(",", "' or " & datatable & ".BranchCode = '") & "' or " & datatable & ".BranchCode = ''  )"
                    Else
                        BranchSql = AndOrWhere & " (" & datatable & ".BranchCode = '" & SecurityBrancheCodes & "'" & " or " & datatable & ".BranchCode = '')"
                    End If
                End If
            End If


        Catch

        End Try
GetBranchCodeExit:
        Return BranchSql

    End Function
    Public Function GetBranchCode(ByVal SLSCode As String) As String '08-30-13
        Dim FatalExceptionCnt As Boolean = False
Start:  Try '05-19-10 JH
            Dim ds As dsSaw8 = New dsSaw8
            ds.EnforceConstraints = False
            Dim strsql As String = "SELECT * FROM namecontact where code = '999999' and category = 'slsman' and EmpCode = '" & SafeSQL(SLSCode) & "'"
            Dim daLookup As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            daLookup.SelectCommand = New MySqlCommand(strsql, myConnection)
            daLookup.Fill(ds, "namecontact")
            If ds.namecontact.Rows.Count <> 0 Then
                Return ds.namecontact.Rows(0)("BranchCode")
            Else
                Return ""
            End If
        Catch ex As Exception
            If ex.Message.Contains("Fatal error encountered during command execution") = True Then
                If FatalExceptionCnt = False Then Call OpenSQL(myConnection) : FatalExceptionCnt = True : GoTo Start Else GoTo DFLTMSG 'Leaving order system open for a long time and then coming back later on and trying to LU
            Else
DFLTMSG:        MessageBox.Show("Error in GetBranchCode" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return ""
        End Try

    End Function

End Module