<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmQuoteRpt
#Region "Windows Form Designer generated code "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        '11-01-10 JTC Set B/4  InitializeComponent(
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-us") '11-01-10 JTC Culture
        System.Threading.Thread.CurrentThread.CurrentUICulture = New System.Globalization.CultureInfo("en-us")

        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public WithEvents mnuline As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFileGoTo As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileSeparator As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFileExit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuTime As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    'Public WithEvents vsPrinter1 As AxvsViewLib.AxvsPrinter
    Public CMDialog1Open As System.Windows.Forms.OpenFileDialog
    Public CMDialog1Save As System.Windows.Forms.SaveFileDialog
    Public CMDialog1Font As System.Windows.Forms.FontDialog
    Public CMDialog1Color As System.Windows.Forms.ColorDialog
    Public CMDialog1Print As System.Windows.Forms.PrintDialog
    Public WithEvents SSPanel1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents cboQuoteRptPrt As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
    Public WithEvents cmdSortPrimarySeq As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Public WithEvents cmdSortSecondarySeq As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Public WithEvents fraOutputOptions As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
    Public WithEvents optOutputOptions As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    Public WithEvents txtCurrentPage As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmQuoteRpt))
        Me.txtSpecifierCode = New System.Windows.Forms.TextBox()
        Me.txtSalesman = New System.Windows.Forms.TextBox()
        Me.txtRetrieval = New System.Windows.Forms.TextBox()
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.txtEndQuoteAmt = New System.Windows.Forms.TextBox()
        Me.txtStartQuoteAmt = New System.Windows.Forms.TextBox()
        Me.txtJobNameSS = New System.Windows.Forms.TextBox()
        Me.txtCSR = New System.Windows.Forms.TextBox()
        Me.txtLastChgBy = New System.Windows.Forms.TextBox()
        Me.txtSlsSplit = New System.Windows.Forms.TextBox()
        Me.txtState = New System.Windows.Forms.TextBox()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.txtMktSegment = New System.Windows.Forms.TextBox()
        Me.chkBranchReport = New System.Windows.Forms.CheckBox()
        Me.chkIncludeCommDolPer = New System.Windows.Forms.CheckBox()
        Me.chkExcludeDuplicates = New System.Windows.Forms.CheckBox()
        Me.txtQuoteToSls = New System.Windows.Forms.TextBox()
        Me.txtCSRofCust = New System.Windows.Forms.TextBox()
        Me.txtSelectCode = New System.Windows.Forms.TextBox()
        Me.txtSlsTerr = New System.Windows.Forms.TextBox()
        Me.TxtSingleCatNum = New System.Windows.Forms.TextBox()
        Me.TxtSearchString = New System.Windows.Forms.TextBox()
        Me.txtMfgLine = New System.Windows.Forms.TextBox()
        Me.txtSpecCross = New System.Windows.Forms.TextBox()
        Me.ChkExtendByProb = New System.Windows.Forms.CheckBox()
        Me.cboLinesInclude = New System.Windows.Forms.ComboBox()
        Me.chkPrtPlanLines = New System.Windows.Forms.CheckBox()
        Me.cboTypeofJob = New C1.Win.C1List.C1Combo()
        Me.cboLotUnit = New System.Windows.Forms.ComboBox()
        Me.cboStockJob = New System.Windows.Forms.ComboBox()
        Me.txtQutRealCode = New System.Windows.Forms.TextBox()
        Me.txtPrcCode = New System.Windows.Forms.TextBox()
        Me.ChkTotalsOnly = New System.Windows.Forms.CheckBox()
        Me.txtLastChgByLine = New System.Windows.Forms.TextBox()
        Me.txtRetr = New System.Windows.Forms.TextBox()
        Me.txtStat = New System.Windows.Forms.TextBox()
        Me.optUnitOrExtended_Unit = New System.Windows.Forms.RadioButton()
        Me.optUnitOrExtended_Extd = New System.Windows.Forms.RadioButton()
        Me.optSalesorCost_Cost = New System.Windows.Forms.RadioButton()
        Me.optoptSalesorCost_Sales = New System.Windows.Forms.RadioButton()
        Me.chkIncludeCommDoll = New System.Windows.Forms.CheckBox()
        Me.chkIncludeCommPer = New System.Windows.Forms.CheckBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.chkNotes = New System.Windows.Forms.CheckBox()
        Me.ChkSpecifiers = New System.Windows.Forms.CheckBox()
        Me.chkMfgBreakdown = New System.Windows.Forms.CheckBox()
        Me.chkCustomerBreakdown = New System.Windows.Forms.CheckBox()
        Me.chkDetailTotal = New System.Windows.Forms.CheckBox()
        Me.chkSlsFromHeader = New System.Windows.Forms.CheckBox()
        Me.chkExportAllExcel = New System.Windows.Forms.CheckBox()
        Me.chkBlankBidDates = New System.Windows.Forms.CheckBox()
        Me.cmdReportQuote = New C1.Win.C1Input.C1Button()
        Me.cmdReportTerrSpecCredit = New C1.Win.C1Input.C1Button()
        Me.cmdReportRealization = New C1.Win.C1Input.C1Button()
        Me.cmdReportOtherTypes = New C1.Win.C1Input.C1Button()
        Me.cmdReportLineItems = New C1.Win.C1Input.C1Button()
        Me.chkBlankLine = New System.Windows.Forms.CheckBox()
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuline = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFileGoTo = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.ProgramDateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuBrandReport = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuBrandMfgChg = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuJump = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSupport = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuBrandListLoad = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuBrandExclude = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTime = New System.Windows.Forms.ToolStripMenuItem()
        Me.CMDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.CMDialog1Save = New System.Windows.Forms.SaveFileDialog()
        Me.CMDialog1Font = New System.Windows.Forms.FontDialog()
        Me.CMDialog1Color = New System.Windows.Forms.ColorDialog()
        Me.CMDialog1Print = New System.Windows.Forms.PrintDialog()
        Me.SSPanel1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.cboQuoteRptPrt = New Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray(Me.components)
        Me.cmdSortPrimarySeq = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.cmdSortSecondarySeq = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.fraOutputOptions = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(Me.components)
        Me.optOutputOptions = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.txtCurrentPage = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.ExitMainRibbon = New C1.Win.C1Ribbon.RibbonButton()
        Me.ExitRibbon = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonListItem1 = New C1.Win.C1Ribbon.RibbonListItem()
        Me.RibbonLabel2 = New C1.Win.C1Ribbon.RibbonLabel()
        Me.CloseButton = New C1.Win.C1Ribbon.RibbonSplitButton()
        Me.PrintDocumentButton = New C1.Win.C1Ribbon.RibbonSplitButton()
        Me.SaveAs1DocumentButton = New C1.Win.C1Ribbon.RibbonSplitButton()
        Me.SaveDocumentButton = New C1.Win.C1Ribbon.RibbonButton()
        Me.LookupRibbon = New C1.Win.C1Ribbon.RibbonButton()
        Me.OpenDocumentButton = New C1.Win.C1Ribbon.RibbonButton()
        Me.NewDocumentButton = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonButton32 = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonButton23 = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonButton13 = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonButton3 = New C1.Win.C1Ribbon.RibbonButton()
        Me.PrintPreviewRibbon = New C1.Win.C1Ribbon.RibbonButton()
        Me.PrintQuickRibbon = New C1.Win.C1Ribbon.RibbonButton()
        Me.PrintRibbon = New C1.Win.C1Ribbon.RibbonButton()
        Me.C1Ribbon1 = New C1.Win.C1Ribbon.C1Ribbon()
        Me.RibbonApplicationMenu1 = New C1.Win.C1Ribbon.RibbonApplicationMenu()
        Me.rbnExitMainMenu = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonButton2 = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonListItem2 = New C1.Win.C1Ribbon.RibbonListItem()
        Me.RibbonLabel1 = New C1.Win.C1Ribbon.RibbonLabel()
        Me.RibbonBottomToolBar1 = New C1.Win.C1Ribbon.RibbonBottomToolBar()
        Me.RibbonConfigToolBar1 = New C1.Win.C1Ribbon.RibbonConfigToolBar()
        Me.RibbonStyleMenu = New C1.Win.C1Ribbon.RibbonMenu()
        Me.RibbonToggleGroup1 = New C1.Win.C1Ribbon.RibbonToggleGroup()
        Me.Office2007BlueStyleButton = New C1.Win.C1Ribbon.RibbonToggleButton()
        Me.Office2007SilverStyleButton = New C1.Win.C1Ribbon.RibbonToggleButton()
        Me.Office2007BlackStyleButton = New C1.Win.C1Ribbon.RibbonToggleButton()
        Me.F1HelpButton = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonQat1 = New C1.Win.C1Ribbon.RibbonQat()
        Me.RbnBtnHelp = New C1.Win.C1Ribbon.RibbonButton()
        Me.rtCustomize = New C1.Win.C1Ribbon.RibbonTab()
        Me.rgColors = New C1.Win.C1Ribbon.RibbonGroup()
        Me.RibbonColorPicker1 = New C1.Win.C1Ribbon.RibbonColorPicker()
        Me.RibbonColorPicker2 = New C1.Win.C1Ribbon.RibbonColorPicker()
        Me.rgThemes = New C1.Win.C1Ribbon.RibbonGroup()
        Me.RibbonGallery1 = New C1.Win.C1Ribbon.RibbonGallery()
        Me.RibbonGalleryItem1 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem2 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem3 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem4 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem5 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem6 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem7 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem8 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem9 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem10 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem11 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem12 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem13 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem14 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem15 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem16 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.RibbonGalleryItem17 = New C1.Win.C1Ribbon.RibbonGalleryItem()
        Me.rgFont = New C1.Win.C1Ribbon.RibbonGroup()
        Me.RibbonToolBar1 = New C1.Win.C1Ribbon.RibbonToolBar()
        Me.RibbonFontComboBox2 = New C1.Win.C1Ribbon.RibbonFontComboBox()
        Me.FontSizeComboBox = New C1.Win.C1Ribbon.RibbonComboBox()
        Me.size8Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size9Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size10Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size11Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size12Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size14Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size16Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size18Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size20Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size22Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size24Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size26Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size28Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size36Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size48Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.size72Button = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonToolBar2 = New C1.Win.C1Ribbon.RibbonToolBar()
        Me.FontBoldButton = New C1.Win.C1Ribbon.RibbonButton()
        Me.FontItalicButton = New C1.Win.C1Ribbon.RibbonButton()
        Me.FontUnderlineButton = New C1.Win.C1Ribbon.RibbonButton()
        Me.FontStrikethroughButton = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonSeparator7 = New C1.Win.C1Ribbon.RibbonSeparator()
        Me.FontColorPicker = New C1.Win.C1Ribbon.RibbonColorPicker()
        Me.RibbonSeparator8 = New C1.Win.C1Ribbon.RibbonSeparator()
        Me.AutoFit = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonTab1 = New C1.Win.C1Ribbon.RibbonTab()
        Me.RbnBtnGridViewInverted = New C1.Win.C1Ribbon.RibbonGroup()
        Me.RbnBtnGridViewNormal = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnGridViewGroupBy = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnGridViewExpandGp = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnGridViewCollapse = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnGridViewSplit = New C1.Win.C1Ribbon.RibbonButton()
        Me.rbnAddFilterBar = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonTab2 = New C1.Win.C1Ribbon.RibbonTab()
        Me.RibbonGroup3 = New C1.Win.C1Ribbon.RibbonGroup()
        Me.RbnBtnExportExcel = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnExportPDF = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnExportRTF = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnExportCSVTab = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnExportCSVComma = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnExportHTML = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnExportOptions = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnBtnExportPrint = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonGroup2 = New C1.Win.C1Ribbon.RibbonGroup()
        Me.RbnTgToExcel = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonTab3 = New C1.Win.C1Ribbon.RibbonTab()
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem1 = New C1.Win.C1Ribbon.RibbonGroup()
        Me.RbnSaveCurrentGridLayoutSettingsToolStripMenuItem = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnResetToOriginalGridLayoutToolStripMenuItem = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1 = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1 = New C1.Win.C1Ribbon.RibbonButton()
        Me.RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1 = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonGroup1 = New C1.Win.C1Ribbon.RibbonGroup()
        Me.rbnDeleteFiles = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonTab4 = New C1.Win.C1Ribbon.RibbonTab()
        Me.rbnMaxNameLength = New C1.Win.C1Ribbon.RibbonGroup()
        Me.rbnMaxNameTxt = New C1.Win.C1Ribbon.RibbonTextBox()
        Me.rbnMaxJobTxt = New C1.Win.C1Ribbon.RibbonTextBox()
        Me.RibbonTab5 = New C1.Win.C1Ribbon.RibbonTab()
        Me.rbnWholeDollars = New C1.Win.C1Ribbon.RibbonGroup()
        Me.chkWholeDollars = New C1.Win.C1Ribbon.RibbonCheckBox()
        Me.chkAddCommas = New C1.Win.C1Ribbon.RibbonCheckBox()
        Me.chkAddDollarSign = New C1.Win.C1Ribbon.RibbonCheckBox()
        Me.RibbonTab6 = New C1.Win.C1Ribbon.RibbonTab()
        Me.rbnPrintColor = New C1.Win.C1Ribbon.RibbonGroup()
        Me.chkPrintGrayScale = New C1.Win.C1Ribbon.RibbonCheckBox()
        Me.RibbonTab7 = New C1.Win.C1Ribbon.RibbonTab()
        Me.RibbonGroup26 = New C1.Win.C1Ribbon.RibbonGroup()
        Me.rbnJoinGoToMeeting = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonGroup18 = New C1.Win.C1Ribbon.RibbonGroup()
        Me.rbnHelpAbout = New C1.Win.C1Ribbon.RibbonButton()
        Me.rbnHelpAboutDirectory = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonGroup5 = New C1.Win.C1Ribbon.RibbonGroup()
        Me.rbnHelpMaster = New C1.Win.C1Ribbon.RibbonButton()
        Me.RibbonTopToolBar1 = New C1.Win.C1Ribbon.RibbonTopToolBar()
        Me.tabQrt = New System.Windows.Forms.TabControl()
        Me.TabPage0 = New System.Windows.Forms.TabPage()
        Me.fraDisplaySortSeq = New System.Windows.Forms.GroupBox()
        Me.txtPrimarySortSeq = New System.Windows.Forms.TextBox()
        Me.txtSecondarySort = New System.Windows.Forms.TextBox()
        Me.pnlPrimarySortSeq = New System.Windows.Forms.Label()
        Me.pnlTypeOfRpt = New System.Windows.Forms.Label()
        Me.pnlSecondarySort = New System.Windows.Forms.Label()
        Me.fraReport = New System.Windows.Forms.GroupBox()
        Me.cboSortRealization = New System.Windows.Forms.CheckedListBox()
        Me.fraReportCmdSelection = New System.Windows.Forms.GroupBox()
        Me.cmdReportProjShortage = New C1.Win.C1Input.C1Button()
        Me.fraSortSecondarySeq = New System.Windows.Forms.GroupBox()
        Me.txtSortSecondarySeq = New System.Windows.Forms.TextBox()
        Me.cmdSecondarySeqCancel = New C1.Win.C1Input.C1Button()
        Me.cboSortSecondarySeq = New System.Windows.Forms.ListBox()
        Me.cmdSecondarySeqContinue = New C1.Win.C1Input.C1Button()
        Me.SSPanel4 = New System.Windows.Forms.Label()
        Me.fraSortPrimarySeq = New System.Windows.Forms.GroupBox()
        Me.txtPrimarySort = New System.Windows.Forms.TextBox()
        Me.cboSortPrimarySeq = New System.Windows.Forms.ListBox()
        Me.cmdPrimarySeqCancel1 = New C1.Win.C1Input.C1Button()
        Me.cmdPrimarySeqContinue1 = New C1.Win.C1Input.C1Button()
        Me.SSPanel3 = New System.Windows.Forms.Label()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me._fdBranchCode = New C1.Win.C1List.C1Combo()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.chkShowLatestCust = New System.Windows.Forms.CheckBox()
        Me.chkBidJobsOnly = New System.Windows.Forms.CheckBox()
        Me.fraQuoteReports = New System.Windows.Forms.GroupBox()
        Me.chkIncludeNotesLineItems = New System.Windows.Forms.CheckBox()
        Me.chkIncludeSLSSPlit = New System.Windows.Forms.CheckBox()
        Me.chkIncludeSpecifiers = New System.Windows.Forms.CheckBox()
        Me.cboTypeCustomer = New C1.Win.C1List.C1Combo()
        Me.lblTypeCustomer = New System.Windows.Forms.Label()
        Me.chkBrandReport = New System.Windows.Forms.CheckBox()
        Me.cbospeccross = New C1.Win.C1List.C1Combo()
        Me.cmdok1 = New C1.Win.C1Input.C1Button()
        Me.cmdCancel1 = New C1.Win.C1Input.C1Button()
        Me.cmdResetDefaults1 = New C1.Win.C1Input.C1Button()
        Me.pnlQutRealCode = New System.Windows.Forms.Label()
        Me.pnlSpecifierCode = New System.Windows.Forms.Label()
        Me.fraFinishReports = New System.Windows.Forms.GroupBox()
        Me.ChkQuoteNoSpecifiers = New System.Windows.Forms.CheckBox()
        Me.chkSalesmanPerPage = New System.Windows.Forms.CheckBox()
        Me.pnlSpecCross = New System.Windows.Forms.Label()
        Me.lblRetrieval = New System.Windows.Forms.Label()
        Me.pnlLotUnit = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblSalesman = New System.Windows.Forms.Label()
        Me.lblStartQuote = New System.Windows.Forms.Label()
        Me.pnlSlsSplits = New System.Windows.Forms.Label()
        Me.lblEndQuote = New System.Windows.Forms.Label()
        Me.pnlCSR = New System.Windows.Forms.Label()
        Me.pnlSltCode = New System.Windows.Forms.Label()
        Me.pnlStkJob = New System.Windows.Forms.Label()
        Me.pnlCSRdist = New System.Windows.Forms.Label()
        Me.pnlQuoteToSls = New System.Windows.Forms.Label()
        Me.lblJobName = New System.Windows.Forms.Label()
        Me.PnlLastChgBy = New System.Windows.Forms.Label()
        Me.pnlMktSeg = New System.Windows.Forms.Label()
        Me.pnlCity = New System.Windows.Forms.Label()
        Me.pnlState = New System.Windows.Forms.Label()
        Me.fraSelectDate = New System.Windows.Forms.GroupBox()
        Me.DTPicker1EndBid = New System.Windows.Forms.DateTimePicker()
        Me.DTPicker1EndEntry = New System.Windows.Forms.DateTimePicker()
        Me.DTPicker1StartBid = New System.Windows.Forms.DateTimePicker()
        Me.ChkCheckBidDates = New System.Windows.Forms.CheckBox()
        Me.lblEndBid = New System.Windows.Forms.Label()
        Me.lblStartBid = New System.Windows.Forms.Label()
        Me.lblEndEntry = New System.Windows.Forms.Label()
        Me.txtStartEntry = New System.Windows.Forms.TextBox()
        Me.DTPickerStartEntry = New System.Windows.Forms.DateTimePicker()
        Me.lblStartEntry = New System.Windows.Forms.Label()
        Me.txtEndEntry = New System.Windows.Forms.TextBox()
        Me.txtEndBid = New System.Windows.Forms.TextBox()
        Me.txtStartBid = New System.Windows.Forms.TextBox()
        Me.fraQuoteLineReports = New System.Windows.Forms.GroupBox()
        Me.chkHaveMFGCode = New System.Windows.Forms.CheckBox()
        Me.chkPrtNTElines = New System.Windows.Forms.CheckBox()
        Me.cmdOK2 = New C1.Win.C1Input.C1Button()
        Me.cmdCancel2 = New C1.Win.C1Input.C1Button()
        Me.cmdResetDefaults2 = New C1.Win.C1Input.C1Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkShowCustomers = New System.Windows.Forms.CheckBox()
        Me.chkUseSpecifierCode = New System.Windows.Forms.CheckBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtCustomerCodeLine = New System.Windows.Forms.TextBox()
        Me.pnlPrcCode = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PnlCatNum = New System.Windows.Forms.Label()
        Me.PnlCatSrch = New System.Windows.Forms.Label()
        Me.PnlMfg = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.PnlSls = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PnlRet = New System.Windows.Forms.Label()
        Me.PnlStatus = New System.Windows.Forms.Label()
        Me.fraUnitorExtended = New System.Windows.Forms.GroupBox()
        Me.fraSalesorCost = New System.Windows.Forms.GroupBox()
        Me.fraQtIncludeCommission = New System.Windows.Forms.GroupBox()
        Me.gbxSortSeq = New System.Windows.Forms.GroupBox()
        Me.txtSortSeq = New System.Windows.Forms.TextBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.cmdBackViewGrid = New System.Windows.Forms.Button()
        Me.gbxSortSeqV = New System.Windows.Forms.GroupBox()
        Me.txtSortSeqCriteria = New System.Windows.Forms.TextBox()
        Me.txtSortSeqV = New System.Windows.Forms.TextBox()
        Me.CmdShowColstoPrt = New System.Windows.Forms.Button()
        Me.CmdRunReport = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.tgln = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.QuoteLinesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DsSaw8 = New VQRT.dsSaw8()
        Me.tgSpecReg = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.SpecRegFollowUpBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tglnDIST = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.tgrDIST = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.QuoteRealLUBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.tgQhDIST = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.QutNotesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.tgr = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.tgQh = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.QUTLU1BindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider()
        Me.C1StatusBar2 = New C1.Win.C1Ribbon.C1StatusBar()
        Me.DocumentModifiedLabel = New C1.Win.C1Ribbon.RibbonLabel()
        Me.RibbonSeparator9 = New C1.Win.C1Ribbon.RibbonSeparator()
        Me.pbProgress = New C1.Win.C1Ribbon.RibbonProgressBar()
        Me.cmdPercent = New C1.Win.C1Ribbon.RibbonButton()
        Me.trackbar = New C1.Win.C1Ribbon.RibbonTrackBar()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.C1SuperTooltip1 = New C1.Win.C1SuperTooltip.C1SuperTooltip(Me.components)
        Me.QuoteTableAdapter = New VQRT.dsSaw8TableAdapters.quoteTableAdapter()
        Me.QuoteprojectcustBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.QuotelinesTableAdapter = New VQRT.dsSaw8TableAdapters.quotelinesTableAdapter()
        Me.QutSlsSplitBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.QutslssplitTableAdapter = New VQRT.dsSaw8TableAdapters.qutslssplitTableAdapter()
        Me.QutnotesTableAdapter = New VQRT.dsSaw8TableAdapters.qutnotesTableAdapter()
        Me.ProjectcustTableAdapter = New VQRT.dsSaw8TableAdapters.projectcustTableAdapter()
        Me.QutLU1TableAdapter = New VQRT.dsSaw8TableAdapters.QUTLU1TableAdapter()
        Me.QuoteRealLUTableAdapter1 = New VQRT.dsSaw8TableAdapters.QuoteRealLUTableAdapter()
        Me.QuoteRealNDULBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.QuoteRealNDULTableAdapter = New VQRT.dsSaw8TableAdapters.QuoteRealNDULTableAdapter()
        Me.DsSaw8BindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SpecRegFollowUpTableAdapter = New VQRT.dsSaw8TableAdapters.SpecRegFollowUpTableAdapter()
        Me.ChkSpecifiersCustInCols = New System.Windows.Forms.CheckBox()
        CType(Me.cboTypeofJob, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdReportQuote, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdReportTerrSpecCredit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdReportRealization, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdReportOtherTypes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdReportLineItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainMenu1.SuspendLayout()
        CType(Me.SSPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboQuoteRptPrt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdSortPrimarySeq, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdSortSecondarySeq, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fraOutputOptions, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optOutputOptions, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCurrentPage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.C1Ribbon1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabQrt.SuspendLayout()
        Me.TabPage0.SuspendLayout()
        Me.fraDisplaySortSeq.SuspendLayout()
        Me.fraReport.SuspendLayout()
        Me.fraReportCmdSelection.SuspendLayout()
        CType(Me.cmdReportProjShortage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraSortSecondarySeq.SuspendLayout()
        CType(Me.cmdSecondarySeqCancel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdSecondarySeqContinue, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraSortPrimarySeq.SuspendLayout()
        CType(Me.cmdPrimarySeqCancel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdPrimarySeqContinue1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        CType(Me._fdBranchCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraQuoteReports.SuspendLayout()
        CType(Me.cboTypeCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbospeccross, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdok1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdCancel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdResetDefaults1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraFinishReports.SuspendLayout()
        Me.fraSelectDate.SuspendLayout()
        Me.fraQuoteLineReports.SuspendLayout()
        CType(Me.cmdOK2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdCancel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdResetDefaults2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.fraUnitorExtended.SuspendLayout()
        Me.fraSalesorCost.SuspendLayout()
        Me.fraQtIncludeCommission.SuspendLayout()
        Me.gbxSortSeq.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gbxSortSeqV.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.tgln, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QuoteLinesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DsSaw8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tgSpecReg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SpecRegFollowUpBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tglnDIST, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tgrDIST, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QuoteRealLUBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tgQhDIST, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QutNotesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tgr, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tgQh, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QUTLU1BindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.C1StatusBar2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QuoteprojectcustBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QutSlsSplitBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.QuoteRealNDULBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DsSaw8BindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtSpecifierCode
        '
        Me.txtSpecifierCode.AcceptsReturn = True
        Me.txtSpecifierCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtSpecifierCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSpecifierCode.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpecifierCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSpecifierCode.Location = New System.Drawing.Point(152, 127)
        Me.txtSpecifierCode.MaxLength = 0
        Me.txtSpecifierCode.Name = "txtSpecifierCode"
        Me.txtSpecifierCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSpecifierCode.Size = New System.Drawing.Size(81, 19)
        Me.txtSpecifierCode.TabIndex = 5
        Me.txtSpecifierCode.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtSpecifierCode, "Enter A Customer Code. Separate With Commas for Multiple Selections.")
        '
        'txtSalesman
        '
        Me.txtSalesman.AcceptsReturn = True
        Me.txtSalesman.BackColor = System.Drawing.SystemColors.Window
        Me.txtSalesman.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSalesman.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSalesman.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSalesman.Location = New System.Drawing.Point(152, 184)
        Me.txtSalesman.MaxLength = 0
        Me.txtSalesman.Name = "txtSalesman"
        Me.txtSalesman.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSalesman.Size = New System.Drawing.Size(81, 19)
        Me.txtSalesman.TabIndex = 7
        Me.txtSalesman.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtSalesman, "Salesman Codes-Header.  Separate With Commas for Multiple Selections.")
        '
        'txtRetrieval
        '
        Me.txtRetrieval.AcceptsReturn = True
        Me.txtRetrieval.BackColor = System.Drawing.SystemColors.Window
        Me.txtRetrieval.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRetrieval.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRetrieval.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRetrieval.Location = New System.Drawing.Point(152, 155)
        Me.txtRetrieval.MaxLength = 0
        Me.txtRetrieval.Name = "txtRetrieval"
        Me.txtRetrieval.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRetrieval.Size = New System.Drawing.Size(81, 19)
        Me.txtRetrieval.TabIndex = 6
        Me.txtRetrieval.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtRetrieval, "Enter a Retrieval Codes.  Separate With Commas for Multiple Selections.")
        '
        'txtStatus
        '
        Me.txtStatus.AcceptsReturn = True
        Me.txtStatus.BackColor = System.Drawing.SystemColors.Window
        Me.txtStatus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStatus.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStatus.Location = New System.Drawing.Point(153, 99)
        Me.txtStatus.MaxLength = 40
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStatus.Size = New System.Drawing.Size(81, 19)
        Me.txtStatus.TabIndex = 4
        Me.txtStatus.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtStatus, "Quote Status Like OPEN, GOT, Etc. Separate With Commas for Multiple Selections.")
        '
        'txtEndQuoteAmt
        '
        Me.txtEndQuoteAmt.AcceptsReturn = True
        Me.txtEndQuoteAmt.BackColor = System.Drawing.SystemColors.Window
        Me.txtEndQuoteAmt.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEndQuoteAmt.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEndQuoteAmt.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEndQuoteAmt.Location = New System.Drawing.Point(153, 71)
        Me.txtEndQuoteAmt.MaxLength = 12
        Me.txtEndQuoteAmt.Name = "txtEndQuoteAmt"
        Me.txtEndQuoteAmt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEndQuoteAmt.Size = New System.Drawing.Size(81, 19)
        Me.txtEndQuoteAmt.TabIndex = 3
        Me.txtEndQuoteAmt.Text = "999999999"
        Me.C1SuperTooltip1.SetToolTip(Me.txtEndQuoteAmt, "Ending Quote Dollar Dollar Amount (No $ or , )")
        '
        'txtStartQuoteAmt
        '
        Me.txtStartQuoteAmt.AcceptsReturn = True
        Me.txtStartQuoteAmt.BackColor = System.Drawing.SystemColors.Window
        Me.txtStartQuoteAmt.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStartQuoteAmt.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStartQuoteAmt.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStartQuoteAmt.Location = New System.Drawing.Point(152, 43)
        Me.txtStartQuoteAmt.MaxLength = 12
        Me.txtStartQuoteAmt.Name = "txtStartQuoteAmt"
        Me.txtStartQuoteAmt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStartQuoteAmt.Size = New System.Drawing.Size(81, 19)
        Me.txtStartQuoteAmt.TabIndex = 2
        Me.txtStartQuoteAmt.Text = "0"
        Me.C1SuperTooltip1.SetToolTip(Me.txtStartQuoteAmt, "Starting Quote Dollar Amount (No $ or , )")
        '
        'txtJobNameSS
        '
        Me.txtJobNameSS.AcceptsReturn = True
        Me.txtJobNameSS.BackColor = System.Drawing.SystemColors.Window
        Me.txtJobNameSS.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtJobNameSS.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtJobNameSS.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtJobNameSS.Location = New System.Drawing.Point(240, 290)
        Me.txtJobNameSS.MaxLength = 0
        Me.txtJobNameSS.Name = "txtJobNameSS"
        Me.txtJobNameSS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtJobNameSS.Size = New System.Drawing.Size(177, 19)
        Me.txtJobNameSS.TabIndex = 21
        Me.C1SuperTooltip1.SetToolTip(Me.txtJobNameSS, "Search For Job Names Containing Selected Search String.")
        '
        'txtCSR
        '
        Me.txtCSR.AcceptsReturn = True
        Me.txtCSR.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSR.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCSR.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCSR.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCSR.Location = New System.Drawing.Point(350, 98)
        Me.txtCSR.MaxLength = 40
        Me.txtCSR.Name = "txtCSR"
        Me.txtCSR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCSR.Size = New System.Drawing.Size(65, 19)
        Me.txtCSR.TabIndex = 14
        Me.txtCSR.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtCSR, "Customer Service Codes.  Separate With Commas for Multiple Selections. Separate W" &
        "ith Commas for Multiple Selections.")
        '
        'txtLastChgBy
        '
        Me.txtLastChgBy.AcceptsReturn = True
        Me.txtLastChgBy.BackColor = System.Drawing.SystemColors.Window
        Me.txtLastChgBy.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLastChgBy.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastChgBy.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLastChgBy.Location = New System.Drawing.Point(152, 212)
        Me.txtLastChgBy.MaxLength = 40
        Me.txtLastChgBy.Name = "txtLastChgBy"
        Me.txtLastChgBy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLastChgBy.Size = New System.Drawing.Size(81, 19)
        Me.txtLastChgBy.TabIndex = 8
        Me.txtLastChgBy.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtLastChgBy, "Salesman Codes.  Separate With Commas for Multiple Selections.")
        '
        'txtSlsSplit
        '
        Me.txtSlsSplit.AcceptsReturn = True
        Me.txtSlsSplit.BackColor = System.Drawing.SystemColors.Window
        Me.txtSlsSplit.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSlsSplit.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSlsSplit.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSlsSplit.Location = New System.Drawing.Point(350, 216)
        Me.txtSlsSplit.MaxLength = 40
        Me.txtSlsSplit.Name = "txtSlsSplit"
        Me.txtSlsSplit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSlsSplit.Size = New System.Drawing.Size(67, 19)
        Me.txtSlsSplit.TabIndex = 20
        Me.txtSlsSplit.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtSlsSplit, "Salesman Codes.  Separate With Commas for Multiple Selections.")
        '
        'txtState
        '
        Me.txtState.AcceptsReturn = True
        Me.txtState.BackColor = System.Drawing.SystemColors.Window
        Me.txtState.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtState.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtState.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtState.Location = New System.Drawing.Point(152, 240)
        Me.txtState.MaxLength = 10
        Me.txtState.Name = "txtState"
        Me.txtState.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtState.Size = New System.Drawing.Size(81, 19)
        Me.txtState.TabIndex = 9
        Me.txtState.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtState, "State Codes.  Separate With Commas for Multiple Selections.")
        '
        'txtCity
        '
        Me.txtCity.AcceptsReturn = True
        Me.txtCity.BackColor = System.Drawing.SystemColors.Window
        Me.txtCity.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCity.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCity.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCity.Location = New System.Drawing.Point(152, 266)
        Me.txtCity.MaxLength = 20
        Me.txtCity.Name = "txtCity"
        Me.txtCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCity.Size = New System.Drawing.Size(81, 19)
        Me.txtCity.TabIndex = 10
        Me.txtCity.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtCity, "City Name")
        '
        'txtMktSegment
        '
        Me.txtMktSegment.AcceptsReturn = True
        Me.txtMktSegment.BackColor = System.Drawing.SystemColors.Window
        Me.txtMktSegment.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMktSegment.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMktSegment.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMktSegment.Location = New System.Drawing.Point(153, 294)
        Me.txtMktSegment.MaxLength = 20
        Me.txtMktSegment.Name = "txtMktSegment"
        Me.txtMktSegment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMktSegment.Size = New System.Drawing.Size(81, 19)
        Me.txtMktSegment.TabIndex = 11
        Me.txtMktSegment.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtMktSegment, "Market Segment")
        '
        'chkBranchReport
        '
        Me.chkBranchReport.AutoSize = True
        Me.chkBranchReport.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBranchReport.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBranchReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBranchReport.Location = New System.Drawing.Point(6, 279)
        Me.chkBranchReport.Name = "chkBranchReport"
        Me.chkBranchReport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBranchReport.Size = New System.Drawing.Size(110, 18)
        Me.chkBranchReport.TabIndex = 32
        Me.chkBranchReport.Text = "Branch Reporting"
        Me.C1SuperTooltip1.SetToolTip(Me.chkBranchReport, "Use Branch Codes instead of SLS Codes")
        Me.chkBranchReport.UseVisualStyleBackColor = False
        '
        'chkIncludeCommDolPer
        '
        Me.chkIncludeCommDolPer.AutoSize = True
        Me.chkIncludeCommDolPer.Checked = True
        Me.chkIncludeCommDolPer.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIncludeCommDolPer.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncludeCommDolPer.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIncludeCommDolPer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIncludeCommDolPer.Location = New System.Drawing.Point(6, 40)
        Me.chkIncludeCommDolPer.Name = "chkIncludeCommDolPer"
        Me.chkIncludeCommDolPer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncludeCommDolPer.Size = New System.Drawing.Size(213, 18)
        Me.chkIncludeCommDolPer.TabIndex = 23
        Me.chkIncludeCommDolPer.Text = "Include Commission $ and % on Report"
        Me.C1SuperTooltip1.SetToolTip(Me.chkIncludeCommDolPer, "Include Commission $ and % On Report")
        Me.chkIncludeCommDolPer.UseVisualStyleBackColor = False
        '
        'chkExcludeDuplicates
        '
        Me.chkExcludeDuplicates.AutoSize = True
        Me.chkExcludeDuplicates.Checked = True
        Me.chkExcludeDuplicates.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkExcludeDuplicates.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkExcludeDuplicates.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkExcludeDuplicates.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkExcludeDuplicates.Location = New System.Drawing.Point(6, 14)
        Me.chkExcludeDuplicates.Name = "chkExcludeDuplicates"
        Me.chkExcludeDuplicates.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkExcludeDuplicates.Size = New System.Drawing.Size(212, 18)
        Me.chkExcludeDuplicates.TabIndex = 22
        Me.chkExcludeDuplicates.Text = "Exclude Quotes With Status = NOREPT"
        Me.C1SuperTooltip1.SetToolTip(Me.chkExcludeDuplicates, "Do not print quotes marked with a Status of NOREPT")
        Me.chkExcludeDuplicates.UseVisualStyleBackColor = False
        '
        'txtQuoteToSls
        '
        Me.txtQuoteToSls.AcceptsReturn = True
        Me.txtQuoteToSls.BackColor = System.Drawing.SystemColors.Window
        Me.txtQuoteToSls.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtQuoteToSls.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuoteToSls.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtQuoteToSls.Location = New System.Drawing.Point(350, 12)
        Me.txtQuoteToSls.MaxLength = 40
        Me.txtQuoteToSls.Name = "txtQuoteToSls"
        Me.txtQuoteToSls.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtQuoteToSls.Size = New System.Drawing.Size(65, 19)
        Me.txtQuoteToSls.TabIndex = 13
        Me.txtQuoteToSls.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtQuoteToSls, "QuoteTo Salesman Codes.  Separate With Commas for Multiple Selections.")
        '
        'txtCSRofCust
        '
        Me.txtCSRofCust.AcceptsReturn = True
        Me.txtCSRofCust.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSRofCust.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCSRofCust.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCSRofCust.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCSRofCust.Location = New System.Drawing.Point(350, 188)
        Me.txtCSRofCust.MaxLength = 20
        Me.txtCSRofCust.Name = "txtCSRofCust"
        Me.txtCSRofCust.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCSRofCust.Size = New System.Drawing.Size(67, 19)
        Me.txtCSRofCust.TabIndex = 12
        Me.txtCSRofCust.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtCSRofCust, "Customr Service Codes.  Separate With Commas for Multiple Selections.")
        Me.txtCSRofCust.Visible = False
        '
        'txtSelectCode
        '
        Me.txtSelectCode.AcceptsReturn = True
        Me.txtSelectCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtSelectCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSelectCode.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSelectCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSelectCode.Location = New System.Drawing.Point(350, 70)
        Me.txtSelectCode.MaxLength = 40
        Me.txtSelectCode.Name = "txtSelectCode"
        Me.txtSelectCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSelectCode.Size = New System.Drawing.Size(68, 19)
        Me.txtSelectCode.TabIndex = 17
        Me.txtSelectCode.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtSelectCode, "Enter your Selectcode. Separate With Commas for Multiple Selections.")
        '
        'txtSlsTerr
        '
        Me.txtSlsTerr.AcceptsReturn = True
        Me.txtSlsTerr.BackColor = System.Drawing.SystemColors.Window
        Me.txtSlsTerr.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSlsTerr.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSlsTerr.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSlsTerr.Location = New System.Drawing.Point(138, 102)
        Me.txtSlsTerr.MaxLength = 30
        Me.txtSlsTerr.Name = "txtSlsTerr"
        Me.txtSlsTerr.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSlsTerr.Size = New System.Drawing.Size(79, 19)
        Me.txtSlsTerr.TabIndex = 3
        Me.txtSlsTerr.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtSlsTerr, "Enter one SLS Code or ALL")
        '
        'TxtSingleCatNum
        '
        Me.TxtSingleCatNum.AcceptsReturn = True
        Me.TxtSingleCatNum.BackColor = System.Drawing.SystemColors.Window
        Me.TxtSingleCatNum.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtSingleCatNum.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtSingleCatNum.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtSingleCatNum.Location = New System.Drawing.Point(176, 16)
        Me.TxtSingleCatNum.MaxLength = 45
        Me.TxtSingleCatNum.Name = "TxtSingleCatNum"
        Me.TxtSingleCatNum.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtSingleCatNum.Size = New System.Drawing.Size(185, 19)
        Me.TxtSingleCatNum.TabIndex = 0
        Me.TxtSingleCatNum.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.TxtSingleCatNum, "Exact Catalog Match or ALL")
        '
        'TxtSearchString
        '
        Me.TxtSearchString.AcceptsReturn = True
        Me.TxtSearchString.BackColor = System.Drawing.SystemColors.Window
        Me.TxtSearchString.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtSearchString.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtSearchString.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtSearchString.Location = New System.Drawing.Point(176, 46)
        Me.TxtSearchString.MaxLength = 45
        Me.TxtSearchString.Name = "TxtSearchString"
        Me.TxtSearchString.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtSearchString.Size = New System.Drawing.Size(185, 19)
        Me.TxtSearchString.TabIndex = 1
        Me.C1SuperTooltip1.SetToolTip(Me.TxtSearchString, "Search for any character string in each Catalog Desc.")
        '
        'txtMfgLine
        '
        Me.txtMfgLine.AcceptsReturn = True
        Me.txtMfgLine.BackColor = System.Drawing.SystemColors.Window
        Me.txtMfgLine.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMfgLine.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMfgLine.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMfgLine.Location = New System.Drawing.Point(176, 74)
        Me.txtMfgLine.MaxLength = 300
        Me.txtMfgLine.Name = "txtMfgLine"
        Me.txtMfgLine.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMfgLine.Size = New System.Drawing.Size(185, 19)
        Me.txtMfgLine.TabIndex = 2
        Me.txtMfgLine.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtMfgLine, "For Multiple MFG's - Separate each by a Comma")
        '
        'txtSpecCross
        '
        Me.txtSpecCross.AcceptsReturn = True
        Me.txtSpecCross.BackColor = System.Drawing.SystemColors.Window
        Me.txtSpecCross.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSpecCross.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSpecCross.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSpecCross.Location = New System.Drawing.Point(176, 102)
        Me.txtSpecCross.MaxLength = 6
        Me.txtSpecCross.Name = "txtSpecCross"
        Me.txtSpecCross.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSpecCross.Size = New System.Drawing.Size(79, 19)
        Me.txtSpecCross.TabIndex = 3
        Me.txtSpecCross.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtSpecCross, "S = Specified:  C = Crossed")
        '
        'ChkExtendByProb
        '
        Me.ChkExtendByProb.AutoSize = True
        Me.ChkExtendByProb.Cursor = System.Windows.Forms.Cursors.Default
        Me.ChkExtendByProb.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkExtendByProb.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ChkExtendByProb.Location = New System.Drawing.Point(6, 253)
        Me.ChkExtendByProb.Name = "ChkExtendByProb"
        Me.ChkExtendByProb.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ChkExtendByProb.Size = New System.Drawing.Size(159, 18)
        Me.ChkExtendByProb.TabIndex = 33
        Me.ChkExtendByProb.Text = "Extend By Quote Probability"
        Me.C1SuperTooltip1.SetToolTip(Me.ChkExtendByProb, "Extend The Sell Amout By Probability")
        Me.ChkExtendByProb.UseVisualStyleBackColor = False
        '
        'cboLinesInclude
        '
        Me.cboLinesInclude.BackColor = System.Drawing.SystemColors.Window
        Me.cboLinesInclude.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLinesInclude.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLinesInclude.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLinesInclude.Items.AddRange(New Object() {"Include All Lines on Job", "Include Only Paid Items on the Job", "Include Only UnPaid Items on the Job"})
        Me.cboLinesInclude.Location = New System.Drawing.Point(227, 318)
        Me.cboLinesInclude.Name = "cboLinesInclude"
        Me.cboLinesInclude.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLinesInclude.Size = New System.Drawing.Size(190, 21)
        Me.cboLinesInclude.TabIndex = 425
        Me.cboLinesInclude.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.cboLinesInclude, "Include All, Paid or Only UnPaid Items on the Job")
        Me.cboLinesInclude.Visible = False
        '
        'chkPrtPlanLines
        '
        Me.chkPrtPlanLines.AutoSize = True
        Me.chkPrtPlanLines.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPrtPlanLines.Location = New System.Drawing.Point(20, 347)
        Me.chkPrtPlanLines.Name = "chkPrtPlanLines"
        Me.chkPrtPlanLines.Size = New System.Drawing.Size(153, 18)
        Me.chkPrtPlanLines.TabIndex = 424
        Me.chkPrtPlanLines.Text = "Print Planned Project Lines"
        Me.C1SuperTooltip1.SetToolTip(Me.chkPrtPlanLines, "Print Planned Project Lines")
        Me.chkPrtPlanLines.UseVisualStyleBackColor = True
        '
        'cboTypeofJob
        '
        Me.cboTypeofJob.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.cboTypeofJob.AutoCompletion = True
        Me.cboTypeofJob.Caption = ""
        Me.cboTypeofJob.CaptionHeight = 17
        Me.cboTypeofJob.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboTypeofJob.ColumnCaptionHeight = 17
        Me.cboTypeofJob.ColumnFooterHeight = 17
        Me.cboTypeofJob.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.cboTypeofJob.DropDownWidth = 200
        Me.cboTypeofJob.EditorBackColor = System.Drawing.SystemColors.Window
        Me.cboTypeofJob.EditorFont = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTypeofJob.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTypeofJob.ExtendRightColumn = True
        Me.cboTypeofJob.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTypeofJob.Images.Add(CType(resources.GetObject("cboTypeofJob.Images"), System.Drawing.Image))
        Me.cboTypeofJob.ItemHeight = 15
        Me.cboTypeofJob.LimitToList = True
        Me.cboTypeofJob.Location = New System.Drawing.Point(350, 244)
        Me.cboTypeofJob.MatchEntryTimeout = CType(2000, Long)
        Me.cboTypeofJob.MaxDropDownItems = CType(30, Short)
        Me.cboTypeofJob.MaxLength = 10
        Me.cboTypeofJob.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.cboTypeofJob.Name = "cboTypeofJob"
        Me.cboTypeofJob.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.cboTypeofJob.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.cboTypeofJob.Size = New System.Drawing.Size(68, 21)
        Me.cboTypeofJob.SuperBack = True
        Me.cboTypeofJob.TabIndex = 423
        Me.cboTypeofJob.Text = "Q"
        Me.C1SuperTooltip1.SetToolTip(Me.cboTypeofJob, "A = All Jobs, Q = Quotes,  S = Spec Credit,  P = Planned Proj, T = Submittal Proj" &
        ",  O = Other")
        Me.cboTypeofJob.PropBag = resources.GetString("cboTypeofJob.PropBag")
        '
        'cboLotUnit
        '
        Me.cboLotUnit.BackColor = System.Drawing.SystemColors.Window
        Me.cboLotUnit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLotUnit.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLotUnit.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLotUnit.Items.AddRange(New Object() {"Lot", "Unit", "ALL"})
        Me.cboLotUnit.Location = New System.Drawing.Point(350, 126)
        Me.cboLotUnit.Name = "cboLotUnit"
        Me.cboLotUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLotUnit.Size = New System.Drawing.Size(65, 21)
        Me.cboLotUnit.TabIndex = 16
        Me.cboLotUnit.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.cboLotUnit, "Lot Or Unit Pricing")
        '
        'cboStockJob
        '
        Me.cboStockJob.BackColor = System.Drawing.SystemColors.Window
        Me.cboStockJob.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboStockJob.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboStockJob.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboStockJob.Items.AddRange(New Object() {"Stock", "Job", "ALL"})
        Me.cboStockJob.Location = New System.Drawing.Point(350, 157)
        Me.cboStockJob.Name = "cboStockJob"
        Me.cboStockJob.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboStockJob.Size = New System.Drawing.Size(65, 21)
        Me.cboStockJob.TabIndex = 19
        Me.cboStockJob.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.cboStockJob, "Enter S = Stock, J = Job ")
        '
        'txtQutRealCode
        '
        Me.txtQutRealCode.AcceptsReturn = True
        Me.txtQutRealCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtQutRealCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtQutRealCode.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQutRealCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtQutRealCode.Location = New System.Drawing.Point(152, 15)
        Me.txtQutRealCode.MaxLength = 0
        Me.txtQutRealCode.Name = "txtQutRealCode"
        Me.txtQutRealCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtQutRealCode.Size = New System.Drawing.Size(81, 19)
        Me.txtQutRealCode.TabIndex = 1
        Me.txtQutRealCode.Text = " "
        Me.C1SuperTooltip1.SetToolTip(Me.txtQutRealCode, "Enter Code for Report (ABC, HERRY,CONTR1) Separate With Commas for Multiple Selec" &
        "tions.")
        Me.txtQutRealCode.Visible = False
        '
        'txtPrcCode
        '
        Me.txtPrcCode.AcceptsReturn = True
        Me.txtPrcCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrcCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrcCode.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrcCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrcCode.Location = New System.Drawing.Point(176, 130)
        Me.txtPrcCode.MaxLength = 6
        Me.txtPrcCode.Name = "txtPrcCode"
        Me.txtPrcCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrcCode.Size = New System.Drawing.Size(79, 19)
        Me.txtPrcCode.TabIndex = 4
        Me.txtPrcCode.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtPrcCode, "Enter Specific Price Code (IE: AA, PF, Etc.)")
        '
        'ChkTotalsOnly
        '
        Me.ChkTotalsOnly.AutoSize = True
        Me.ChkTotalsOnly.Cursor = System.Windows.Forms.Cursors.Default
        Me.ChkTotalsOnly.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkTotalsOnly.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ChkTotalsOnly.Location = New System.Drawing.Point(9, 150)
        Me.ChkTotalsOnly.Name = "ChkTotalsOnly"
        Me.ChkTotalsOnly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ChkTotalsOnly.Size = New System.Drawing.Size(114, 18)
        Me.ChkTotalsOnly.TabIndex = 60
        Me.ChkTotalsOnly.Text = "Totals Only Report"
        Me.C1SuperTooltip1.SetToolTip(Me.ChkTotalsOnly, "Totals Only Report")
        Me.ChkTotalsOnly.UseVisualStyleBackColor = False
        '
        'txtLastChgByLine
        '
        Me.txtLastChgByLine.AcceptsReturn = True
        Me.txtLastChgByLine.BackColor = System.Drawing.SystemColors.Window
        Me.txtLastChgByLine.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLastChgByLine.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastChgByLine.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLastChgByLine.Location = New System.Drawing.Point(138, 74)
        Me.txtLastChgByLine.MaxLength = 6
        Me.txtLastChgByLine.Name = "txtLastChgByLine"
        Me.txtLastChgByLine.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLastChgByLine.Size = New System.Drawing.Size(79, 19)
        Me.txtLastChgByLine.TabIndex = 2
        Me.txtLastChgByLine.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtLastChgByLine, "Enter User Codes. Separate by Commas")
        '
        'txtRetr
        '
        Me.txtRetr.AcceptsReturn = True
        Me.txtRetr.BackColor = System.Drawing.SystemColors.Window
        Me.txtRetr.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRetr.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRetr.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRetr.Location = New System.Drawing.Point(138, 46)
        Me.txtRetr.MaxLength = 10
        Me.txtRetr.Name = "txtRetr"
        Me.txtRetr.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRetr.Size = New System.Drawing.Size(121, 19)
        Me.txtRetr.TabIndex = 1
        Me.txtRetr.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtRetr, "Enter Your Retrieval Codes. Separate by Commas")
        '
        'txtStat
        '
        Me.txtStat.AcceptsReturn = True
        Me.txtStat.BackColor = System.Drawing.SystemColors.Window
        Me.txtStat.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStat.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStat.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStat.Location = New System.Drawing.Point(138, 18)
        Me.txtStat.MaxLength = 40
        Me.txtStat.Name = "txtStat"
        Me.txtStat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStat.Size = New System.Drawing.Size(79, 19)
        Me.txtStat.TabIndex = 0
        Me.txtStat.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtStat, "Enter Status Codes Like: OPEN, GOT, Etc.")
        '
        'optUnitOrExtended_Unit
        '
        Me.optUnitOrExtended_Unit.AutoSize = True
        Me.optUnitOrExtended_Unit.Cursor = System.Windows.Forms.Cursors.Default
        Me.optUnitOrExtended_Unit.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUnitOrExtended_Unit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUnitOrExtended_Unit.Location = New System.Drawing.Point(5, 37)
        Me.optUnitOrExtended_Unit.Name = "optUnitOrExtended_Unit"
        Me.optUnitOrExtended_Unit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optUnitOrExtended_Unit.Size = New System.Drawing.Size(76, 18)
        Me.optUnitOrExtended_Unit.TabIndex = 66
        Me.optUnitOrExtended_Unit.TabStop = True
        Me.optUnitOrExtended_Unit.Text = "Unit Prices"
        Me.C1SuperTooltip1.SetToolTip(Me.optUnitOrExtended_Unit, "Print Unit Prices")
        Me.optUnitOrExtended_Unit.UseVisualStyleBackColor = False
        '
        'optUnitOrExtended_Extd
        '
        Me.optUnitOrExtended_Extd.AutoSize = True
        Me.optUnitOrExtended_Extd.Checked = True
        Me.optUnitOrExtended_Extd.Cursor = System.Windows.Forms.Cursors.Default
        Me.optUnitOrExtended_Extd.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optUnitOrExtended_Extd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optUnitOrExtended_Extd.Location = New System.Drawing.Point(5, 11)
        Me.optUnitOrExtended_Extd.Name = "optUnitOrExtended_Extd"
        Me.optUnitOrExtended_Extd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optUnitOrExtended_Extd.Size = New System.Drawing.Size(103, 18)
        Me.optUnitOrExtended_Extd.TabIndex = 67
        Me.optUnitOrExtended_Extd.TabStop = True
        Me.optUnitOrExtended_Extd.Text = "Extended Prices"
        Me.C1SuperTooltip1.SetToolTip(Me.optUnitOrExtended_Extd, "Print Extended Prices")
        Me.optUnitOrExtended_Extd.UseVisualStyleBackColor = False
        '
        'optSalesorCost_Cost
        '
        Me.optSalesorCost_Cost.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSalesorCost_Cost.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optSalesorCost_Cost.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSalesorCost_Cost.Location = New System.Drawing.Point(5, 42)
        Me.optSalesorCost_Cost.Name = "optSalesorCost_Cost"
        Me.optSalesorCost_Cost.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSalesorCost_Cost.Size = New System.Drawing.Size(97, 17)
        Me.optSalesorCost_Cost.TabIndex = 1
        Me.optSalesorCost_Cost.TabStop = True
        Me.optSalesorCost_Cost.Text = "Cost Dollars"
        Me.C1SuperTooltip1.SetToolTip(Me.optSalesorCost_Cost, "Select All Orders")
        Me.optSalesorCost_Cost.UseVisualStyleBackColor = False
        '
        'optoptSalesorCost_Sales
        '
        Me.optoptSalesorCost_Sales.Checked = True
        Me.optoptSalesorCost_Sales.Cursor = System.Windows.Forms.Cursors.Default
        Me.optoptSalesorCost_Sales.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optoptSalesorCost_Sales.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optoptSalesorCost_Sales.Location = New System.Drawing.Point(5, 14)
        Me.optoptSalesorCost_Sales.Name = "optoptSalesorCost_Sales"
        Me.optoptSalesorCost_Sales.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optoptSalesorCost_Sales.Size = New System.Drawing.Size(121, 17)
        Me.optoptSalesorCost_Sales.TabIndex = 0
        Me.optoptSalesorCost_Sales.TabStop = True
        Me.optoptSalesorCost_Sales.Text = "Sales Dollars"
        Me.C1SuperTooltip1.SetToolTip(Me.optoptSalesorCost_Sales, "Select Just Open Orders")
        Me.optoptSalesorCost_Sales.UseVisualStyleBackColor = False
        '
        'chkIncludeCommDoll
        '
        Me.chkIncludeCommDoll.AutoSize = True
        Me.chkIncludeCommDoll.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncludeCommDoll.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIncludeCommDoll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIncludeCommDoll.Location = New System.Drawing.Point(5, 11)
        Me.chkIncludeCommDoll.Name = "chkIncludeCommDoll"
        Me.chkIncludeCommDoll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncludeCommDoll.Size = New System.Drawing.Size(101, 18)
        Me.chkIncludeCommDoll.TabIndex = 83
        Me.chkIncludeCommDoll.Text = "Include Comm $"
        Me.C1SuperTooltip1.SetToolTip(Me.chkIncludeCommDoll, "Include Comm $")
        Me.chkIncludeCommDoll.UseVisualStyleBackColor = False
        '
        'chkIncludeCommPer
        '
        Me.chkIncludeCommPer.AutoSize = True
        Me.chkIncludeCommPer.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIncludeCommPer.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIncludeCommPer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIncludeCommPer.Location = New System.Drawing.Point(5, 46)
        Me.chkIncludeCommPer.Name = "chkIncludeCommPer"
        Me.chkIncludeCommPer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIncludeCommPer.Size = New System.Drawing.Size(105, 18)
        Me.chkIncludeCommPer.TabIndex = 82
        Me.chkIncludeCommPer.Text = "Include Comm %"
        Me.C1SuperTooltip1.SetToolTip(Me.chkIncludeCommPer, "Include Comm %")
        Me.chkIncludeCommPer.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(238, 246)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(63, 14)
        Me.Label4.TabIndex = 421
        Me.Label4.Text = "Type of Job"
        Me.C1SuperTooltip1.SetToolTip(Me.Label4, "A = All Jobs, Q = Quotes,  S = Spec Credit,  P = Planned Proj, T = Submittal Proj" &
        ",  O = Other")
        '
        'chkNotes
        '
        Me.chkNotes.AutoSize = True
        Me.chkNotes.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkNotes.Location = New System.Drawing.Point(6, 227)
        Me.chkNotes.Name = "chkNotes"
        Me.chkNotes.Size = New System.Drawing.Size(124, 18)
        Me.chkNotes.TabIndex = 31
        Me.chkNotes.Text = "Add Notes to Report"
        Me.C1SuperTooltip1.SetToolTip(Me.chkNotes, "Add Notes to Report")
        Me.chkNotes.UseVisualStyleBackColor = True
        '
        'ChkSpecifiers
        '
        Me.ChkSpecifiers.AutoSize = True
        Me.ChkSpecifiers.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkSpecifiers.Location = New System.Drawing.Point(6, 177)
        Me.ChkSpecifiers.Name = "ChkSpecifiers"
        Me.ChkSpecifiers.Size = New System.Drawing.Size(231, 18)
        Me.ChkSpecifiers.TabIndex = 30
        Me.ChkSpecifiers.Text = "Add Specifiers (Arch, Eng, Etc) to Reports"
        Me.C1SuperTooltip1.SetToolTip(Me.ChkSpecifiers, "Add Specifiers (Arch, Eng, Etc) to Reports")
        Me.ChkSpecifiers.UseVisualStyleBackColor = True
        '
        'chkMfgBreakdown
        '
        Me.chkMfgBreakdown.AutoSize = True
        Me.chkMfgBreakdown.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMfgBreakdown.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkMfgBreakdown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMfgBreakdown.Location = New System.Drawing.Point(6, 149)
        Me.chkMfgBreakdown.Name = "chkMfgBreakdown"
        Me.chkMfgBreakdown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMfgBreakdown.Size = New System.Drawing.Size(208, 18)
        Me.chkMfgBreakdown.TabIndex = 29
        Me.chkMfgBreakdown.Text = "Add MFG Total Breakdown to Reports"
        Me.C1SuperTooltip1.SetToolTip(Me.chkMfgBreakdown, "Add MFG Total Breakdown To Reports")
        Me.chkMfgBreakdown.UseVisualStyleBackColor = False
        '
        'chkCustomerBreakdown
        '
        Me.chkCustomerBreakdown.AutoSize = True
        Me.chkCustomerBreakdown.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCustomerBreakdown.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCustomerBreakdown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCustomerBreakdown.Location = New System.Drawing.Point(6, 123)
        Me.chkCustomerBreakdown.Name = "chkCustomerBreakdown"
        Me.chkCustomerBreakdown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCustomerBreakdown.Size = New System.Drawing.Size(220, 18)
        Me.chkCustomerBreakdown.TabIndex = 28
        Me.chkCustomerBreakdown.Text = "Add Cust QuoteTo Breakdown to Report"
        Me.C1SuperTooltip1.SetToolTip(Me.chkCustomerBreakdown, "Add Cust QuoteTo Breakdown To Reports")
        Me.chkCustomerBreakdown.UseVisualStyleBackColor = False
        '
        'chkDetailTotal
        '
        Me.chkDetailTotal.AutoSize = True
        Me.chkDetailTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDetailTotal.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDetailTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDetailTotal.Location = New System.Drawing.Point(6, 97)
        Me.chkDetailTotal.Name = "chkDetailTotal"
        Me.chkDetailTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDetailTotal.Size = New System.Drawing.Size(79, 18)
        Me.chkDetailTotal.TabIndex = 26
        Me.chkDetailTotal.Text = "Totals Only"
        Me.C1SuperTooltip1.SetToolTip(Me.chkDetailTotal, "Totals only on Reports")
        Me.chkDetailTotal.UseVisualStyleBackColor = False
        '
        'chkSlsFromHeader
        '
        Me.chkSlsFromHeader.AutoSize = True
        Me.chkSlsFromHeader.Checked = True
        Me.chkSlsFromHeader.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSlsFromHeader.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSlsFromHeader.Enabled = False
        Me.chkSlsFromHeader.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSlsFromHeader.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSlsFromHeader.Location = New System.Drawing.Point(6, 68)
        Me.chkSlsFromHeader.Name = "chkSlsFromHeader"
        Me.chkSlsFromHeader.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSlsFromHeader.Size = New System.Drawing.Size(242, 18)
        Me.chkSlsFromHeader.TabIndex = 24
        Me.chkSlsFromHeader.Text = "Use Salesman From Quote Header on Report"
        Me.C1SuperTooltip1.SetToolTip(Me.chkSlsFromHeader, "Use Salesman From Quote Header on Report ")
        Me.chkSlsFromHeader.UseVisualStyleBackColor = False
        '
        'chkExportAllExcel
        '
        Me.chkExportAllExcel.AutoSize = True
        Me.chkExportAllExcel.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkExportAllExcel.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkExportAllExcel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkExportAllExcel.Location = New System.Drawing.Point(122, 277)
        Me.chkExportAllExcel.Name = "chkExportAllExcel"
        Me.chkExportAllExcel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkExportAllExcel.Size = New System.Drawing.Size(169, 18)
        Me.chkExportAllExcel.TabIndex = 25
        Me.chkExportAllExcel.Text = "Export All Quote Data to Excel"
        Me.C1SuperTooltip1.SetToolTip(Me.chkExportAllExcel, "Export All Quote Data to Excel")
        Me.chkExportAllExcel.UseVisualStyleBackColor = False
        Me.chkExportAllExcel.Visible = False
        '
        'chkBlankBidDates
        '
        Me.chkBlankBidDates.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBlankBidDates.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBlankBidDates.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBlankBidDates.Location = New System.Drawing.Point(9, 92)
        Me.chkBlankBidDates.Name = "chkBlankBidDates"
        Me.chkBlankBidDates.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBlankBidDates.Size = New System.Drawing.Size(259, 22)
        Me.chkBlankBidDates.TabIndex = 68
        Me.chkBlankBidDates.Text = "Include Quotes with Blank Bid Dates"
        Me.C1SuperTooltip1.SetToolTip(Me.chkBlankBidDates, "Include Quotes with Blank Bid Dates")
        Me.chkBlankBidDates.UseVisualStyleBackColor = False
        '
        'cmdReportQuote
        '
        Me.cmdReportQuote.Location = New System.Drawing.Point(12, 16)
        Me.cmdReportQuote.Name = "cmdReportQuote"
        Me.cmdReportQuote.Size = New System.Drawing.Size(183, 35)
        Me.cmdReportQuote.TabIndex = 37
        Me.cmdReportQuote.Text = "Quote Summary Report"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdReportQuote, "Report each Quote  in a variety of Sequences")
        Me.cmdReportQuote.UseVisualStyleBackColor = True
        Me.cmdReportQuote.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cmdReportTerrSpecCredit
        '
        Me.cmdReportTerrSpecCredit.Location = New System.Drawing.Point(12, 210)
        Me.cmdReportTerrSpecCredit.Name = "cmdReportTerrSpecCredit"
        Me.cmdReportTerrSpecCredit.Size = New System.Drawing.Size(183, 35)
        Me.cmdReportTerrSpecCredit.TabIndex = 41
        Me.cmdReportTerrSpecCredit.Text = "Terr Spec Credit Report"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdReportTerrSpecCredit, "Out of Terr Spec Credit Report By Mfg or Salesman")
        Me.cmdReportTerrSpecCredit.UseVisualStyleBackColor = True
        Me.cmdReportTerrSpecCredit.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cmdReportRealization
        '
        Me.cmdReportRealization.Location = New System.Drawing.Point(12, 57)
        Me.cmdReportRealization.Name = "cmdReportRealization"
        Me.cmdReportRealization.Size = New System.Drawing.Size(183, 50)
        Me.cmdReportRealization.TabIndex = 38
        Me.cmdReportRealization.Text = "Customer/MFG's/Specifiers" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Quote to Realization Report"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdReportRealization, "Customer/Manufacturer/Specifier Quote volumes and Attainment Rates")
        Me.cmdReportRealization.UseVisualStyleBackColor = True
        Me.cmdReportRealization.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cmdReportOtherTypes
        '
        Me.cmdReportOtherTypes.Location = New System.Drawing.Point(12, 169)
        Me.cmdReportOtherTypes.Name = "cmdReportOtherTypes"
        Me.cmdReportOtherTypes.Size = New System.Drawing.Size(183, 35)
        Me.cmdReportOtherTypes.TabIndex = 40
        Me.cmdReportOtherTypes.Text = "Other Quote Types"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdReportOtherTypes, "Other Types - Planned Project, Submittals, Other Types")
        Me.cmdReportOtherTypes.UseVisualStyleBackColor = True
        Me.cmdReportOtherTypes.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cmdReportLineItems
        '
        Me.cmdReportLineItems.Location = New System.Drawing.Point(12, 113)
        Me.cmdReportLineItems.Name = "cmdReportLineItems"
        Me.cmdReportLineItems.Size = New System.Drawing.Size(183, 50)
        Me.cmdReportLineItems.TabIndex = 39
        Me.cmdReportLineItems.Text = "Product Sales History" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "  Quote Line Item Reporting"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdReportLineItems, "Product Sales History by Mfg ang Catalog Number in Quotes")
        Me.cmdReportLineItems.UseVisualStyleBackColor = True
        Me.cmdReportLineItems.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'chkBlankLine
        '
        Me.chkBlankLine.AutoSize = True
        Me.chkBlankLine.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBlankLine.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBlankLine.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBlankLine.Location = New System.Drawing.Point(6, 302)
        Me.chkBlankLine.Name = "chkBlankLine"
        Me.chkBlankLine.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBlankLine.Size = New System.Drawing.Size(160, 18)
        Me.chkBlankLine.TabIndex = 34
        Me.chkBlankLine.Text = "Add Blank Line After  Quote"
        Me.C1SuperTooltip1.SetToolTip(Me.chkBlankLine, "Add a Blabk line After the Primary Sort Sequence.")
        Me.chkBlankLine.UseVisualStyleBackColor = False
        '
        'MainMenu1
        '
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuExit, Me.mnuTime})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(1056, 24)
        Me.MainMenu1.TabIndex = 114
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuline, Me.mnuFileGoTo, Me.mnuFileSeparator, Me.ProgramDateToolStripMenuItem, Me.mnuFileExit, Me.mnuBrandReport, Me.mnuBrandMfgChg, Me.mnuJump, Me.mnuSupport, Me.mnuBrandListLoad, Me.mnuBrandExclude})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(37, 20)
        Me.mnuFile.Text = "File"
        '
        'mnuline
        '
        Me.mnuline.Name = "mnuline"
        Me.mnuline.Size = New System.Drawing.Size(261, 6)
        '
        'mnuFileGoTo
        '
        Me.mnuFileGoTo.Name = "mnuFileGoTo"
        Me.mnuFileGoTo.Size = New System.Drawing.Size(264, 22)
        Me.mnuFileGoTo.Text = "Go To Quote Menu"
        '
        'mnuFileSeparator
        '
        Me.mnuFileSeparator.Name = "mnuFileSeparator"
        Me.mnuFileSeparator.Size = New System.Drawing.Size(261, 6)
        '
        'ProgramDateToolStripMenuItem
        '
        Me.ProgramDateToolStripMenuItem.Name = "ProgramDateToolStripMenuItem"
        Me.ProgramDateToolStripMenuItem.Size = New System.Drawing.Size(264, 22)
        Me.ProgramDateToolStripMenuItem.Text = "Program Date"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(264, 22)
        Me.mnuFileExit.Text = "E&xit to Windows"
        '
        'mnuBrandReport
        '
        Me.mnuBrandReport.Name = "mnuBrandReport"
        Me.mnuBrandReport.Size = New System.Drawing.Size(264, 22)
        Me.mnuBrandReport.Text = "Brand Reporting - Off"
        Me.mnuBrandReport.ToolTipText = "Brand Line Item Reporting Feature On/Off"
        '
        'mnuBrandMfgChg
        '
        Me.mnuBrandMfgChg.Name = "mnuBrandMfgChg"
        Me.mnuBrandMfgChg.Size = New System.Drawing.Size(264, 22)
        Me.mnuBrandMfgChg.Text = "Brand Mfg Code - XXXX"
        Me.mnuBrandMfgChg.ToolTipText = "Change Brand Mfg Code (PHIL, COOP, LITH) for Line Item Reporting"
        '
        'mnuJump
        '
        Me.mnuJump.Name = "mnuJump"
        Me.mnuJump.Size = New System.Drawing.Size(264, 22)
        Me.mnuJump.Text = "Jump to Name-Address Brand Table"
        Me.mnuJump.ToolTipText = "Jump to Name-Address Brand Table"
        '
        'mnuSupport
        '
        Me.mnuSupport.Name = "mnuSupport"
        Me.mnuSupport.Size = New System.Drawing.Size(264, 22)
        Me.mnuSupport.Text = "Support Functions"
        Me.mnuSupport.ToolTipText = "For use by Support Personnel"
        '
        'mnuBrandListLoad
        '
        Me.mnuBrandListLoad.Name = "mnuBrandListLoad"
        Me.mnuBrandListLoad.Size = New System.Drawing.Size(264, 22)
        Me.mnuBrandListLoad.Text = "Brand Listing Load"
        '
        'mnuBrandExclude
        '
        Me.mnuBrandExclude.Name = "mnuBrandExclude"
        Me.mnuBrandExclude.Size = New System.Drawing.Size(264, 22)
        Me.mnuBrandExclude.Text = "Exclude Brands"
        '
        'mnuExit
        '
        Me.mnuExit.Name = "mnuExit"
        Me.mnuExit.Size = New System.Drawing.Size(115, 20)
        Me.mnuExit.Text = "E&xit to Main Menu"
        '
        'mnuTime
        '
        Me.mnuTime.Name = "mnuTime"
        Me.mnuTime.Size = New System.Drawing.Size(87, 20)
        Me.mnuTime.Text = "Time Display"
        '
        'CMDialog1Font
        '
        Me.CMDialog1Font.Color = System.Drawing.SystemColors.ControlText
        '
        'cboQuoteRptPrt
        '
        '
        'txtCurrentPage
        '
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'ExitMainRibbon
        '
        Me.ExitMainRibbon.Name = "ExitMainRibbon"
        Me.ExitMainRibbon.SmallImage = CType(resources.GetObject("ExitMainRibbon.SmallImage"), System.Drawing.Image)
        Me.ExitMainRibbon.Text = "Exit to Main Menu"
        '
        'ExitRibbon
        '
        Me.ExitRibbon.Name = "ExitRibbon"
        Me.ExitRibbon.SmallImage = CType(resources.GetObject("ExitRibbon.SmallImage"), System.Drawing.Image)
        Me.ExitRibbon.Text = "Exit"
        '
        'RibbonListItem1
        '
        Me.RibbonListItem1.Name = "RibbonListItem1"
        '
        'RibbonLabel2
        '
        Me.RibbonLabel2.Name = "RibbonLabel2"
        Me.RibbonLabel2.Text = "Recent Documents"
        '
        'CloseButton
        '
        Me.CloseButton.LargeImage = CType(resources.GetObject("CloseButton.LargeImage"), System.Drawing.Image)
        Me.CloseButton.Name = "CloseButton"
        Me.CloseButton.Text = "Close"
        '
        'PrintDocumentButton
        '
        Me.PrintDocumentButton.LargeImage = CType(resources.GetObject("PrintDocumentButton.LargeImage"), System.Drawing.Image)
        Me.PrintDocumentButton.Name = "PrintDocumentButton"
        Me.PrintDocumentButton.Text = "Print"
        '
        'SaveAs1DocumentButton
        '
        Me.SaveAs1DocumentButton.LargeImage = CType(resources.GetObject("SaveAs1DocumentButton.LargeImage"), System.Drawing.Image)
        Me.SaveAs1DocumentButton.Name = "SaveAs1DocumentButton"
        Me.SaveAs1DocumentButton.Text = "Save As"
        '
        'SaveDocumentButton
        '
        Me.SaveDocumentButton.LargeImage = CType(resources.GetObject("SaveDocumentButton.LargeImage"), System.Drawing.Image)
        Me.SaveDocumentButton.Name = "SaveDocumentButton"
        Me.SaveDocumentButton.Text = "Save"
        '
        'LookupRibbon
        '
        Me.LookupRibbon.LargeImage = CType(resources.GetObject("LookupRibbon.LargeImage"), System.Drawing.Image)
        Me.LookupRibbon.Name = "LookupRibbon"
        Me.LookupRibbon.SmallImage = CType(resources.GetObject("LookupRibbon.SmallImage"), System.Drawing.Image)
        Me.LookupRibbon.Text = "Lookup"
        '
        'OpenDocumentButton
        '
        Me.OpenDocumentButton.LargeImage = CType(resources.GetObject("OpenDocumentButton.LargeImage"), System.Drawing.Image)
        Me.OpenDocumentButton.Name = "OpenDocumentButton"
        Me.OpenDocumentButton.Text = "Open"
        '
        'NewDocumentButton
        '
        Me.NewDocumentButton.LargeImage = CType(resources.GetObject("NewDocumentButton.LargeImage"), System.Drawing.Image)
        Me.NewDocumentButton.Name = "NewDocumentButton"
        Me.NewDocumentButton.Text = "New"
        '
        'RibbonButton32
        '
        Me.RibbonButton32.Description = "Close all open documents"
        Me.RibbonButton32.LargeImage = CType(resources.GetObject("RibbonButton32.LargeImage"), System.Drawing.Image)
        Me.RibbonButton32.Name = "RibbonButton32"
        Me.RibbonButton32.Text = "Close All"
        '
        'RibbonButton23
        '
        Me.RibbonButton23.Description = "Close the current document"
        Me.RibbonButton23.LargeImage = CType(resources.GetObject("RibbonButton23.LargeImage"), System.Drawing.Image)
        Me.RibbonButton23.Name = "RibbonButton23"
        Me.RibbonButton23.Text = "Close"
        '
        'RibbonButton13
        '
        Me.RibbonButton13.LargeImage = CType(resources.GetObject("RibbonButton13.LargeImage"), System.Drawing.Image)
        Me.RibbonButton13.Name = "RibbonButton13"
        Me.RibbonButton13.Text = "Word Document"
        '
        'RibbonButton3
        '
        Me.RibbonButton3.LargeImage = CType(resources.GetObject("RibbonButton3.LargeImage"), System.Drawing.Image)
        Me.RibbonButton3.Name = "RibbonButton3"
        Me.RibbonButton3.Text = "Save as PDF"
        '
        'PrintPreviewRibbon
        '
        Me.PrintPreviewRibbon.Description = "Preview the document before printing"
        Me.PrintPreviewRibbon.LargeImage = CType(resources.GetObject("PrintPreviewRibbon.LargeImage"), System.Drawing.Image)
        Me.PrintPreviewRibbon.Name = "PrintPreviewRibbon"
        Me.PrintPreviewRibbon.Text = "Print Preview"
        '
        'PrintQuickRibbon
        '
        Me.PrintQuickRibbon.Description = "Send the document directly to the default printer without making any changes"
        Me.PrintQuickRibbon.LargeImage = CType(resources.GetObject("PrintQuickRibbon.LargeImage"), System.Drawing.Image)
        Me.PrintQuickRibbon.Name = "PrintQuickRibbon"
        Me.PrintQuickRibbon.Text = "Button"
        '
        'PrintRibbon
        '
        Me.PrintRibbon.Description = "Select a printer, number of copies and other print options before printing"
        Me.PrintRibbon.LargeImage = CType(resources.GetObject("PrintRibbon.LargeImage"), System.Drawing.Image)
        Me.PrintRibbon.Name = "PrintRibbon"
        Me.PrintRibbon.Text = "Print"
        '
        'C1Ribbon1
        '
        Me.C1Ribbon1.ApplicationMenuHolder = Me.RibbonApplicationMenu1
        Me.C1Ribbon1.AutoSizeElement = C1.Framework.AutoSizeElement.Width
        Me.C1Ribbon1.BottomToolBarHolder = Me.RibbonBottomToolBar1
        Me.C1Ribbon1.ConfigToolBarHolder = Me.RibbonConfigToolBar1
        Me.C1Ribbon1.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.C1Ribbon1.Location = New System.Drawing.Point(0, 24)
        Me.C1Ribbon1.Name = "C1Ribbon1"
        Me.C1Ribbon1.QatHolder = Me.RibbonQat1
        Me.C1Ribbon1.QatItemsHolder.Add(Me.RbnBtnHelp)
        Me.C1Ribbon1.Size = New System.Drawing.Size(1056, 155)
        Me.C1Ribbon1.Tabs.Add(Me.rtCustomize)
        Me.C1Ribbon1.Tabs.Add(Me.RibbonTab1)
        Me.C1Ribbon1.Tabs.Add(Me.RibbonTab2)
        Me.C1Ribbon1.Tabs.Add(Me.RibbonTab3)
        Me.C1Ribbon1.Tabs.Add(Me.RibbonTab4)
        Me.C1Ribbon1.Tabs.Add(Me.RibbonTab5)
        Me.C1Ribbon1.Tabs.Add(Me.RibbonTab6)
        Me.C1Ribbon1.Tabs.Add(Me.RibbonTab7)
        Me.C1Ribbon1.TopToolBarHolder = Me.RibbonTopToolBar1
        '
        'RibbonApplicationMenu1
        '
        Me.RibbonApplicationMenu1.BottomPaneItems.Add(Me.rbnExitMainMenu)
        Me.RibbonApplicationMenu1.BottomPaneItems.Add(Me.RibbonButton2)
        Me.RibbonApplicationMenu1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RibbonApplicationMenu1.LargeImage = CType(resources.GetObject("RibbonApplicationMenu1.LargeImage"), System.Drawing.Image)
        Me.RibbonApplicationMenu1.Name = "RibbonApplicationMenu1"
        Me.RibbonApplicationMenu1.RightPaneItems.Add(Me.RibbonListItem2)
        '
        'rbnExitMainMenu
        '
        Me.rbnExitMainMenu.Name = "rbnExitMainMenu"
        Me.rbnExitMainMenu.SmallImage = CType(resources.GetObject("rbnExitMainMenu.SmallImage"), System.Drawing.Image)
        Me.rbnExitMainMenu.Text = "Exit"
        '
        'RibbonButton2
        '
        Me.RibbonButton2.Name = "RibbonButton2"
        Me.RibbonButton2.SmallImage = CType(resources.GetObject("RibbonButton2.SmallImage"), System.Drawing.Image)
        Me.RibbonButton2.Text = "Exit to Main Menu"
        '
        'RibbonListItem2
        '
        Me.RibbonListItem2.Items.Add(Me.RibbonLabel1)
        Me.RibbonListItem2.Name = "RibbonListItem2"
        '
        'RibbonLabel1
        '
        Me.RibbonLabel1.Name = "RibbonLabel1"
        Me.RibbonLabel1.Text = "Recent Documents"
        '
        'RibbonBottomToolBar1
        '
        Me.RibbonBottomToolBar1.Name = "RibbonBottomToolBar1"
        '
        'RibbonConfigToolBar1
        '
        Me.RibbonConfigToolBar1.Items.Add(Me.RibbonStyleMenu)
        Me.RibbonConfigToolBar1.Items.Add(Me.F1HelpButton)
        Me.RibbonConfigToolBar1.Name = "RibbonConfigToolBar1"
        '
        'RibbonStyleMenu
        '
        Me.RibbonStyleMenu.Items.Add(Me.RibbonToggleGroup1)
        Me.RibbonStyleMenu.Name = "RibbonStyleMenu"
        Me.RibbonStyleMenu.Text = "Ribbon Style"
        '
        'RibbonToggleGroup1
        '
        Me.RibbonToggleGroup1.Items.Add(Me.Office2007BlueStyleButton)
        Me.RibbonToggleGroup1.Items.Add(Me.Office2007SilverStyleButton)
        Me.RibbonToggleGroup1.Items.Add(Me.Office2007BlackStyleButton)
        Me.RibbonToggleGroup1.Name = "RibbonToggleGroup1"
        '
        'Office2007BlueStyleButton
        '
        Me.Office2007BlueStyleButton.Name = "Office2007BlueStyleButton"
        Me.Office2007BlueStyleButton.Text = "Blue"
        '
        'Office2007SilverStyleButton
        '
        Me.Office2007SilverStyleButton.Name = "Office2007SilverStyleButton"
        Me.Office2007SilverStyleButton.Text = "Silver"
        '
        'Office2007BlackStyleButton
        '
        Me.Office2007BlackStyleButton.Name = "Office2007BlackStyleButton"
        Me.Office2007BlackStyleButton.Text = "Black"
        '
        'F1HelpButton
        '
        Me.F1HelpButton.Name = "F1HelpButton"
        Me.F1HelpButton.SmallImage = CType(resources.GetObject("F1HelpButton.SmallImage"), System.Drawing.Image)
        '
        'RibbonQat1
        '
        Me.RibbonQat1.ItemLinks.Add(Me.rbnExitMainMenu)
        Me.RibbonQat1.Name = "RibbonQat1"
        '
        'RbnBtnHelp
        '
        Me.RbnBtnHelp.Description = "Help"
        Me.RbnBtnHelp.Name = "RbnBtnHelp"
        Me.RbnBtnHelp.SmallImage = CType(resources.GetObject("RbnBtnHelp.SmallImage"), System.Drawing.Image)
        Me.RbnBtnHelp.Text = "Help"
        '
        'rtCustomize
        '
        Me.rtCustomize.Groups.Add(Me.rgColors)
        Me.rtCustomize.Groups.Add(Me.rgThemes)
        Me.rtCustomize.Groups.Add(Me.rgFont)
        Me.rtCustomize.Name = "rtCustomize"
        Me.rtCustomize.Text = "Colors"
        '
        'rgColors
        '
        Me.rgColors.Items.Add(Me.RibbonColorPicker1)
        Me.rgColors.Items.Add(Me.RibbonColorPicker2)
        Me.rgColors.Name = "rgColors"
        Me.rgColors.Text = "Colors"
        '
        'RibbonColorPicker1
        '
        Me.RibbonColorPicker1.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.RibbonColorPicker1.Name = "RibbonColorPicker1"
        Me.RibbonColorPicker1.SmallImage = CType(resources.GetObject("RibbonColorPicker1.SmallImage"), System.Drawing.Image)
        Me.RibbonColorPicker1.Text = "Foreground"
        '
        'RibbonColorPicker2
        '
        Me.RibbonColorPicker2.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.RibbonColorPicker2.Name = "RibbonColorPicker2"
        Me.RibbonColorPicker2.SmallImage = CType(resources.GetObject("RibbonColorPicker2.SmallImage"), System.Drawing.Image)
        Me.RibbonColorPicker2.Text = "Background"
        '
        'rgThemes
        '
        Me.rgThemes.Items.Add(Me.RibbonGallery1)
        Me.rgThemes.Name = "rgThemes"
        Me.rgThemes.Text = "Themes"
        '
        'RibbonGallery1
        '
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem1)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem2)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem3)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem4)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem5)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem6)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem7)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem8)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem9)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem10)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem11)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem12)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem13)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem14)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem15)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem16)
        Me.RibbonGallery1.Items.Add(Me.RibbonGalleryItem17)
        Me.RibbonGallery1.Name = "RibbonGallery1"
        Me.RibbonGallery1.VisibleItems = 2
        '
        'RibbonGalleryItem1
        '
        Me.RibbonGalleryItem1.LargeImage = CType(resources.GetObject("RibbonGalleryItem1.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem1.Name = "RibbonGalleryItem1"
        '
        'RibbonGalleryItem2
        '
        Me.RibbonGalleryItem2.LargeImage = CType(resources.GetObject("RibbonGalleryItem2.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem2.Name = "RibbonGalleryItem2"
        '
        'RibbonGalleryItem3
        '
        Me.RibbonGalleryItem3.LargeImage = CType(resources.GetObject("RibbonGalleryItem3.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem3.Name = "RibbonGalleryItem3"
        '
        'RibbonGalleryItem4
        '
        Me.RibbonGalleryItem4.LargeImage = CType(resources.GetObject("RibbonGalleryItem4.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem4.Name = "RibbonGalleryItem4"
        '
        'RibbonGalleryItem5
        '
        Me.RibbonGalleryItem5.LargeImage = CType(resources.GetObject("RibbonGalleryItem5.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem5.Name = "RibbonGalleryItem5"
        '
        'RibbonGalleryItem6
        '
        Me.RibbonGalleryItem6.LargeImage = CType(resources.GetObject("RibbonGalleryItem6.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem6.Name = "RibbonGalleryItem6"
        '
        'RibbonGalleryItem7
        '
        Me.RibbonGalleryItem7.LargeImage = CType(resources.GetObject("RibbonGalleryItem7.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem7.Name = "RibbonGalleryItem7"
        '
        'RibbonGalleryItem8
        '
        Me.RibbonGalleryItem8.LargeImage = CType(resources.GetObject("RibbonGalleryItem8.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem8.Name = "RibbonGalleryItem8"
        '
        'RibbonGalleryItem9
        '
        Me.RibbonGalleryItem9.LargeImage = CType(resources.GetObject("RibbonGalleryItem9.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem9.Name = "RibbonGalleryItem9"
        '
        'RibbonGalleryItem10
        '
        Me.RibbonGalleryItem10.LargeImage = CType(resources.GetObject("RibbonGalleryItem10.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem10.Name = "RibbonGalleryItem10"
        '
        'RibbonGalleryItem11
        '
        Me.RibbonGalleryItem11.LargeImage = CType(resources.GetObject("RibbonGalleryItem11.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem11.Name = "RibbonGalleryItem11"
        '
        'RibbonGalleryItem12
        '
        Me.RibbonGalleryItem12.LargeImage = CType(resources.GetObject("RibbonGalleryItem12.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem12.Name = "RibbonGalleryItem12"
        '
        'RibbonGalleryItem13
        '
        Me.RibbonGalleryItem13.LargeImage = CType(resources.GetObject("RibbonGalleryItem13.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem13.Name = "RibbonGalleryItem13"
        '
        'RibbonGalleryItem14
        '
        Me.RibbonGalleryItem14.LargeImage = CType(resources.GetObject("RibbonGalleryItem14.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem14.Name = "RibbonGalleryItem14"
        '
        'RibbonGalleryItem15
        '
        Me.RibbonGalleryItem15.LargeImage = CType(resources.GetObject("RibbonGalleryItem15.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem15.Name = "RibbonGalleryItem15"
        '
        'RibbonGalleryItem16
        '
        Me.RibbonGalleryItem16.LargeImage = CType(resources.GetObject("RibbonGalleryItem16.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem16.Name = "RibbonGalleryItem16"
        '
        'RibbonGalleryItem17
        '
        Me.RibbonGalleryItem17.LargeImage = CType(resources.GetObject("RibbonGalleryItem17.LargeImage"), System.Drawing.Image)
        Me.RibbonGalleryItem17.Name = "RibbonGalleryItem17"
        '
        'rgFont
        '
        Me.rgFont.Items.Add(Me.RibbonToolBar1)
        Me.rgFont.Items.Add(Me.RibbonToolBar2)
        Me.rgFont.Name = "rgFont"
        Me.rgFont.Text = "Font"
        '
        'RibbonToolBar1
        '
        Me.RibbonToolBar1.Items.Add(Me.RibbonFontComboBox2)
        Me.RibbonToolBar1.Items.Add(Me.FontSizeComboBox)
        Me.RibbonToolBar1.Name = "RibbonToolBar1"
        '
        'RibbonFontComboBox2
        '
        Me.RibbonFontComboBox2.Name = "RibbonFontComboBox2"
        '
        'FontSizeComboBox
        '
        Me.FontSizeComboBox.Items.Add(Me.size8Button)
        Me.FontSizeComboBox.Items.Add(Me.size9Button)
        Me.FontSizeComboBox.Items.Add(Me.size10Button)
        Me.FontSizeComboBox.Items.Add(Me.size11Button)
        Me.FontSizeComboBox.Items.Add(Me.size12Button)
        Me.FontSizeComboBox.Items.Add(Me.size14Button)
        Me.FontSizeComboBox.Items.Add(Me.size16Button)
        Me.FontSizeComboBox.Items.Add(Me.size18Button)
        Me.FontSizeComboBox.Items.Add(Me.size20Button)
        Me.FontSizeComboBox.Items.Add(Me.size22Button)
        Me.FontSizeComboBox.Items.Add(Me.size24Button)
        Me.FontSizeComboBox.Items.Add(Me.size26Button)
        Me.FontSizeComboBox.Items.Add(Me.size28Button)
        Me.FontSizeComboBox.Items.Add(Me.size36Button)
        Me.FontSizeComboBox.Items.Add(Me.size48Button)
        Me.FontSizeComboBox.Items.Add(Me.size72Button)
        Me.FontSizeComboBox.Name = "FontSizeComboBox"
        Me.FontSizeComboBox.TextAreaWidth = 40
        '
        'size8Button
        '
        Me.size8Button.Name = "size8Button"
        Me.size8Button.Text = "8"
        '
        'size9Button
        '
        Me.size9Button.Name = "size9Button"
        Me.size9Button.Text = "9"
        '
        'size10Button
        '
        Me.size10Button.Name = "size10Button"
        Me.size10Button.Text = "10"
        '
        'size11Button
        '
        Me.size11Button.Name = "size11Button"
        Me.size11Button.Text = "11"
        '
        'size12Button
        '
        Me.size12Button.Name = "size12Button"
        Me.size12Button.Text = "12"
        '
        'size14Button
        '
        Me.size14Button.Name = "size14Button"
        Me.size14Button.Text = "14"
        '
        'size16Button
        '
        Me.size16Button.Name = "size16Button"
        Me.size16Button.Text = "16"
        '
        'size18Button
        '
        Me.size18Button.Name = "size18Button"
        Me.size18Button.Text = "18"
        '
        'size20Button
        '
        Me.size20Button.Name = "size20Button"
        Me.size20Button.Text = "20"
        '
        'size22Button
        '
        Me.size22Button.Name = "size22Button"
        Me.size22Button.Text = "22"
        '
        'size24Button
        '
        Me.size24Button.Name = "size24Button"
        Me.size24Button.Text = "24"
        '
        'size26Button
        '
        Me.size26Button.Name = "size26Button"
        Me.size26Button.Text = "26"
        '
        'size28Button
        '
        Me.size28Button.Name = "size28Button"
        Me.size28Button.Text = "28"
        '
        'size36Button
        '
        Me.size36Button.Name = "size36Button"
        Me.size36Button.Text = "36"
        '
        'size48Button
        '
        Me.size48Button.Name = "size48Button"
        Me.size48Button.Text = "48"
        '
        'size72Button
        '
        Me.size72Button.Name = "size72Button"
        Me.size72Button.Text = "72"
        '
        'RibbonToolBar2
        '
        Me.RibbonToolBar2.Items.Add(Me.FontBoldButton)
        Me.RibbonToolBar2.Items.Add(Me.FontItalicButton)
        Me.RibbonToolBar2.Items.Add(Me.FontUnderlineButton)
        Me.RibbonToolBar2.Items.Add(Me.FontStrikethroughButton)
        Me.RibbonToolBar2.Items.Add(Me.RibbonSeparator7)
        Me.RibbonToolBar2.Items.Add(Me.FontColorPicker)
        Me.RibbonToolBar2.Items.Add(Me.RibbonSeparator8)
        Me.RibbonToolBar2.Items.Add(Me.AutoFit)
        Me.RibbonToolBar2.Name = "RibbonToolBar2"
        '
        'FontBoldButton
        '
        Me.FontBoldButton.Name = "FontBoldButton"
        Me.FontBoldButton.SmallImage = CType(resources.GetObject("FontBoldButton.SmallImage"), System.Drawing.Image)
        '
        'FontItalicButton
        '
        Me.FontItalicButton.Name = "FontItalicButton"
        Me.FontItalicButton.SmallImage = CType(resources.GetObject("FontItalicButton.SmallImage"), System.Drawing.Image)
        '
        'FontUnderlineButton
        '
        Me.FontUnderlineButton.Name = "FontUnderlineButton"
        Me.FontUnderlineButton.SmallImage = CType(resources.GetObject("FontUnderlineButton.SmallImage"), System.Drawing.Image)
        '
        'FontStrikethroughButton
        '
        Me.FontStrikethroughButton.Name = "FontStrikethroughButton"
        Me.FontStrikethroughButton.SmallImage = CType(resources.GetObject("FontStrikethroughButton.SmallImage"), System.Drawing.Image)
        '
        'RibbonSeparator7
        '
        Me.RibbonSeparator7.Name = "RibbonSeparator7"
        '
        'FontColorPicker
        '
        Me.FontColorPicker.Color = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.FontColorPicker.Name = "FontColorPicker"
        Me.FontColorPicker.SmallImage = CType(resources.GetObject("FontColorPicker.SmallImage"), System.Drawing.Image)
        '
        'RibbonSeparator8
        '
        Me.RibbonSeparator8.Name = "RibbonSeparator8"
        '
        'AutoFit
        '
        Me.AutoFit.Name = "AutoFit"
        Me.AutoFit.SmallImage = CType(resources.GetObject("AutoFit.SmallImage"), System.Drawing.Image)
        '
        'RibbonTab1
        '
        Me.RibbonTab1.Groups.Add(Me.RbnBtnGridViewInverted)
        Me.RibbonTab1.Name = "RibbonTab1"
        Me.RibbonTab1.Text = "Grid Views"
        Me.RibbonTab1.ToolTip = "See Various Types of Grid Views"
        Me.RibbonTab1.Visible = False
        '
        'RbnBtnGridViewInverted
        '
        Me.RbnBtnGridViewInverted.Items.Add(Me.RbnBtnGridViewNormal)
        Me.RbnBtnGridViewInverted.Items.Add(Me.RbnBtnGridViewGroupBy)
        Me.RbnBtnGridViewInverted.Items.Add(Me.RbnBtnGridViewExpandGp)
        Me.RbnBtnGridViewInverted.Items.Add(Me.RbnBtnGridViewCollapse)
        Me.RbnBtnGridViewInverted.Items.Add(Me.RbnBtnGridViewSplit)
        Me.RbnBtnGridViewInverted.Items.Add(Me.rbnAddFilterBar)
        Me.RbnBtnGridViewInverted.Name = "RbnBtnGridViewInverted"
        Me.RbnBtnGridViewInverted.Text = "RbnBtnGridViewInverted"
        '
        'RbnBtnGridViewNormal
        '
        Me.RbnBtnGridViewNormal.LargeImage = Global.VQRT.Resources.NormalViewLAR
        Me.RbnBtnGridViewNormal.Name = "RbnBtnGridViewNormal"
        Me.RbnBtnGridViewNormal.SmallImage = CType(resources.GetObject("RbnBtnGridViewNormal.SmallImage"), System.Drawing.Image)
        Me.RbnBtnGridViewNormal.Text = "Normal Grid View"
        '
        'RbnBtnGridViewGroupBy
        '
        Me.RbnBtnGridViewGroupBy.Name = "RbnBtnGridViewGroupBy"
        Me.RbnBtnGridViewGroupBy.SmallImage = CType(resources.GetObject("RbnBtnGridViewGroupBy.SmallImage"), System.Drawing.Image)
        Me.RbnBtnGridViewGroupBy.Text = "GroupBy Grid View"
        '
        'RbnBtnGridViewExpandGp
        '
        Me.RbnBtnGridViewExpandGp.Name = "RbnBtnGridViewExpandGp"
        Me.RbnBtnGridViewExpandGp.SmallImage = CType(resources.GetObject("RbnBtnGridViewExpandGp.SmallImage"), System.Drawing.Image)
        Me.RbnBtnGridViewExpandGp.Text = "Expand All Grouped Rows"
        '
        'RbnBtnGridViewCollapse
        '
        Me.RbnBtnGridViewCollapse.Name = "RbnBtnGridViewCollapse"
        Me.RbnBtnGridViewCollapse.SmallImage = CType(resources.GetObject("RbnBtnGridViewCollapse.SmallImage"), System.Drawing.Image)
        Me.RbnBtnGridViewCollapse.Text = "Collapse All Grouped Rows"
        '
        'RbnBtnGridViewSplit
        '
        Me.RbnBtnGridViewSplit.Name = "RbnBtnGridViewSplit"
        Me.RbnBtnGridViewSplit.SmallImage = CType(resources.GetObject("RbnBtnGridViewSplit.SmallImage"), System.Drawing.Image)
        Me.RbnBtnGridViewSplit.Text = "Add Split Grid View"
        '
        'rbnAddFilterBar
        '
        Me.rbnAddFilterBar.Name = "rbnAddFilterBar"
        Me.rbnAddFilterBar.SmallImage = CType(resources.GetObject("rbnAddFilterBar.SmallImage"), System.Drawing.Image)
        Me.rbnAddFilterBar.Text = "Add Filter Bar to Grid"
        '
        'RibbonTab2
        '
        Me.RibbonTab2.Groups.Add(Me.RibbonGroup3)
        Me.RibbonTab2.Groups.Add(Me.RibbonGroup2)
        Me.RibbonTab2.Name = "RibbonTab2"
        Me.RibbonTab2.Text = "Export Options"
        Me.RibbonTab2.ToolTip = "Export the Grid to Excel, Csv, Html, Pdf"
        '
        'RibbonGroup3
        '
        Me.RibbonGroup3.Items.Add(Me.RbnBtnExportExcel)
        Me.RibbonGroup3.Items.Add(Me.RbnBtnExportPDF)
        Me.RibbonGroup3.Items.Add(Me.RbnBtnExportRTF)
        Me.RibbonGroup3.Items.Add(Me.RbnBtnExportCSVTab)
        Me.RibbonGroup3.Items.Add(Me.RbnBtnExportCSVComma)
        Me.RibbonGroup3.Items.Add(Me.RbnBtnExportHTML)
        Me.RibbonGroup3.Items.Add(Me.RbnBtnExportOptions)
        Me.RibbonGroup3.Items.Add(Me.RbnBtnExportPrint)
        Me.RibbonGroup3.Name = "RibbonGroup3"
        Me.RibbonGroup3.Text = "Group"
        '
        'RbnBtnExportExcel
        '
        Me.RbnBtnExportExcel.LargeImage = Global.VQRT.Resources.excel4
        Me.RbnBtnExportExcel.Name = "RbnBtnExportExcel"
        Me.RbnBtnExportExcel.SmallImage = CType(resources.GetObject("RbnBtnExportExcel.SmallImage"), System.Drawing.Image)
        Me.RbnBtnExportExcel.Text = "Export To Excel"
        '
        'RbnBtnExportPDF
        '
        Me.RbnBtnExportPDF.LargeImage = Global.VQRT.Resources.PDF6
        Me.RbnBtnExportPDF.Name = "RbnBtnExportPDF"
        Me.RbnBtnExportPDF.SmallImage = CType(resources.GetObject("RbnBtnExportPDF.SmallImage"), System.Drawing.Image)
        Me.RbnBtnExportPDF.Text = "Export To PDF"
        '
        'RbnBtnExportRTF
        '
        Me.RbnBtnExportRTF.LargeImage = CType(resources.GetObject("RbnBtnExportRTF.LargeImage"), System.Drawing.Image)
        Me.RbnBtnExportRTF.Name = "RbnBtnExportRTF"
        Me.RbnBtnExportRTF.SmallImage = CType(resources.GetObject("RbnBtnExportRTF.SmallImage"), System.Drawing.Image)
        Me.RbnBtnExportRTF.Text = "Export To RTF"
        '
        'RbnBtnExportCSVTab
        '
        Me.RbnBtnExportCSVTab.Name = "RbnBtnExportCSVTab"
        Me.RbnBtnExportCSVTab.SmallImage = CType(resources.GetObject("RbnBtnExportCSVTab.SmallImage"), System.Drawing.Image)
        Me.RbnBtnExportCSVTab.Text = "Export To CSV Tab DeLimited"
        '
        'RbnBtnExportCSVComma
        '
        Me.RbnBtnExportCSVComma.Name = "RbnBtnExportCSVComma"
        Me.RbnBtnExportCSVComma.SmallImage = CType(resources.GetObject("RbnBtnExportCSVComma.SmallImage"), System.Drawing.Image)
        Me.RbnBtnExportCSVComma.Text = "Export To CSV Comma DeLimited"
        '
        'RbnBtnExportHTML
        '
        Me.RbnBtnExportHTML.Name = "RbnBtnExportHTML"
        Me.RbnBtnExportHTML.SmallImage = CType(resources.GetObject("RbnBtnExportHTML.SmallImage"), System.Drawing.Image)
        Me.RbnBtnExportHTML.Text = "Export To HTML"
        '
        'RbnBtnExportOptions
        '
        Me.RbnBtnExportOptions.Name = "RbnBtnExportOptions"
        Me.RbnBtnExportOptions.Text = "Export To Other Options"
        '
        'RbnBtnExportPrint
        '
        Me.RbnBtnExportPrint.Name = "RbnBtnExportPrint"
        Me.RbnBtnExportPrint.Text = "Print Reports"
        '
        'RibbonGroup2
        '
        Me.RibbonGroup2.Items.Add(Me.RbnTgToExcel)
        Me.RibbonGroup2.Name = "RibbonGroup2"
        Me.RibbonGroup2.Text = "Export Lookup Grid "
        '
        'RbnTgToExcel
        '
        Me.RbnTgToExcel.LargeImage = CType(resources.GetObject("RbnTgToExcel.LargeImage"), System.Drawing.Image)
        Me.RbnTgToExcel.Name = "RbnTgToExcel"
        Me.RbnTgToExcel.SmallImage = CType(resources.GetObject("RbnTgToExcel.SmallImage"), System.Drawing.Image)
        Me.RbnTgToExcel.Text = "Export Lookup Grid To Excel"
        '
        'RibbonTab3
        '
        Me.RibbonTab3.Groups.Add(Me.RbnResetToCurrentGridLayoutToolStripMenuItem1)
        Me.RibbonTab3.Groups.Add(Me.RibbonGroup1)
        Me.RibbonTab3.Name = "RibbonTab3"
        Me.RibbonTab3.Text = "Grid Layout"
        Me.RibbonTab3.ToolTip = "Save and Change Grid Columns and Setting"
        '
        'RbnResetToCurrentGridLayoutToolStripMenuItem1
        '
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem1.Items.Add(Me.RbnSaveCurrentGridLayoutSettingsToolStripMenuItem)
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem1.Items.Add(Me.RbnResetToCurrentGridLayoutToolStripMenuItem)
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem1.Items.Add(Me.RbnResetToOriginalGridLayoutToolStripMenuItem)
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem1.Items.Add(Me.RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1)
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem1.Items.Add(Me.RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1)
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem1.Items.Add(Me.RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1)
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem1.Name = "RbnResetToCurrentGridLayoutToolStripMenuItem1"
        '
        'RbnSaveCurrentGridLayoutSettingsToolStripMenuItem
        '
        Me.RbnSaveCurrentGridLayoutSettingsToolStripMenuItem.Name = "RbnSaveCurrentGridLayoutSettingsToolStripMenuItem"
        Me.RbnSaveCurrentGridLayoutSettingsToolStripMenuItem.SmallImage = CType(resources.GetObject("RbnSaveCurrentGridLayoutSettingsToolStripMenuItem.SmallImage"), System.Drawing.Image)
        Me.RbnSaveCurrentGridLayoutSettingsToolStripMenuItem.Text = "Save Current  Grid Layout Settings"
        '
        'RbnResetToCurrentGridLayoutToolStripMenuItem
        '
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem.Name = "RbnResetToCurrentGridLayoutToolStripMenuItem"
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem.SmallImage = CType(resources.GetObject("RbnResetToCurrentGridLayoutToolStripMenuItem.SmallImage"), System.Drawing.Image)
        Me.RbnResetToCurrentGridLayoutToolStripMenuItem.Text = "Reset To Current Grid Layout Settings"
        '
        'RbnResetToOriginalGridLayoutToolStripMenuItem
        '
        Me.RbnResetToOriginalGridLayoutToolStripMenuItem.Name = "RbnResetToOriginalGridLayoutToolStripMenuItem"
        Me.RbnResetToOriginalGridLayoutToolStripMenuItem.SmallImage = CType(resources.GetObject("RbnResetToOriginalGridLayoutToolStripMenuItem.SmallImage"), System.Drawing.Image)
        Me.RbnResetToOriginalGridLayoutToolStripMenuItem.Text = "Reset To Original  Grid Layout Settings"
        '
        'RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1
        '
        Me.RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.Name = "RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1"
        Me.RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.SmallImage = CType(resources.GetObject("RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.SmallImage"), System.Drawing.Image)
        Me.RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.Text = "Save Current QuoteTo Grid Layout Settings"
        Me.RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.Visible = False
        '
        'RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1
        '
        Me.RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.Name = "RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1"
        Me.RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.SmallImage = CType(resources.GetObject("RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.SmallImage"), System.Drawing.Image)
        Me.RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.Text = "Reset To Current QuoteTo Grid Layout Settings"
        Me.RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1.Visible = False
        '
        'RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1
        '
        Me.RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1.Name = "RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1"
        Me.RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1.SmallImage = CType(resources.GetObject("RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1.SmallImage"), System.Drawing.Image)
        Me.RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1.Text = "Reset To Original QuoteTo Grid Layout Settings"
        Me.RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1.Visible = False
        '
        'RibbonGroup1
        '
        Me.RibbonGroup1.Items.Add(Me.rbnDeleteFiles)
        Me.RibbonGroup1.Name = "RibbonGroup1"
        Me.RibbonGroup1.Text = "Delete Layout Files"
        '
        'rbnDeleteFiles
        '
        Me.rbnDeleteFiles.Name = "rbnDeleteFiles"
        Me.rbnDeleteFiles.Text = "Delete Layout Files"
        '
        'RibbonTab4
        '
        Me.RibbonTab4.Groups.Add(Me.rbnMaxNameLength)
        Me.RibbonTab4.Name = "RibbonTab4"
        Me.RibbonTab4.Text = "Width Settings"
        '
        'rbnMaxNameLength
        '
        Me.rbnMaxNameLength.Items.Add(Me.rbnMaxNameTxt)
        Me.rbnMaxNameLength.Items.Add(Me.rbnMaxJobTxt)
        Me.rbnMaxNameLength.Name = "rbnMaxNameLength"
        Me.rbnMaxNameLength.Text = "Max Column Width Settings"
        '
        'rbnMaxNameTxt
        '
        Me.rbnMaxNameTxt.Label = "Max Width for Customer Name Column"
        Me.rbnMaxNameTxt.Name = "rbnMaxNameTxt"
        Me.rbnMaxNameTxt.Text = "45"
        Me.rbnMaxNameTxt.TextAreaWidth = 20
        Me.rbnMaxNameTxt.ToolTip = "Shorten the Max Width for Custmer Name Column get more lines per Print Page. (10/" &
    "45)"
        '
        'rbnMaxJobTxt
        '
        Me.rbnMaxJobTxt.Label = "Max Width for Job Name Column"
        Me.rbnMaxJobTxt.MaxLength = 2
        Me.rbnMaxJobTxt.Name = "rbnMaxJobTxt"
        Me.rbnMaxJobTxt.Text = "40"
        Me.rbnMaxJobTxt.TextAreaWidth = 20
        Me.rbnMaxJobTxt.ToolTip = "Shorten the Max Width for Job Name Column to get more lines per Print Page. (10/4" &
    "0)"
        '
        'RibbonTab5
        '
        Me.RibbonTab5.Groups.Add(Me.rbnWholeDollars)
        Me.RibbonTab5.Name = "RibbonTab5"
        Me.RibbonTab5.Text = "Format Dollars"
        Me.RibbonTab5.ToolTip = "Format Dollars to whole Dollars and with Commas"
        '
        'rbnWholeDollars
        '
        Me.rbnWholeDollars.Items.Add(Me.chkWholeDollars)
        Me.rbnWholeDollars.Items.Add(Me.chkAddCommas)
        Me.rbnWholeDollars.Items.Add(Me.chkAddDollarSign)
        Me.rbnWholeDollars.Name = "rbnWholeDollars"
        Me.rbnWholeDollars.Text = "Format Dollar Amounts"
        '
        'chkWholeDollars
        '
        Me.chkWholeDollars.Name = "chkWholeDollars"
        Me.chkWholeDollars.Text = "Round to Whole Dollars - No Cents"
        Me.chkWholeDollars.ToolTip = "Report in Whole Dollars - No Cents"
        '
        'chkAddCommas
        '
        Me.chkAddCommas.Name = "chkAddCommas"
        Me.chkAddCommas.Text = "Add Commas to Amounts 1,275,320"
        Me.chkAddCommas.ToolTip = "Add Commas to Amounts 1,275,320"
        '
        'chkAddDollarSign
        '
        Me.chkAddDollarSign.Name = "chkAddDollarSign"
        Me.chkAddDollarSign.Text = "Add $ Dollar Sign to Amounts"
        Me.chkAddDollarSign.ToolTip = "Add $ Dollar Sign to Amounts $350"
        '
        'RibbonTab6
        '
        Me.RibbonTab6.Groups.Add(Me.rbnPrintColor)
        Me.RibbonTab6.Name = "RibbonTab6"
        Me.RibbonTab6.Text = "Print Color"
        Me.RibbonTab6.ToolTip = "Print in Color or GrayScale"
        '
        'rbnPrintColor
        '
        Me.rbnPrintColor.Items.Add(Me.chkPrintGrayScale)
        Me.rbnPrintColor.Name = "rbnPrintColor"
        Me.rbnPrintColor.Text = "Print Color/GrayScale"
        '
        'chkPrintGrayScale
        '
        Me.chkPrintGrayScale.Name = "chkPrintGrayScale"
        Me.chkPrintGrayScale.Text = "Print In GrayScale"
        Me.chkPrintGrayScale.ToolTip = "Print in Color or GrayScale"
        '
        'RibbonTab7
        '
        Me.RibbonTab7.Groups.Add(Me.RibbonGroup26)
        Me.RibbonTab7.Groups.Add(Me.RibbonGroup18)
        Me.RibbonTab7.Groups.Add(Me.RibbonGroup5)
        Me.RibbonTab7.Name = "RibbonTab7"
        Me.RibbonTab7.Text = "Help - F1"
        '
        'RibbonGroup26
        '
        Me.RibbonGroup26.Items.Add(Me.rbnJoinGoToMeeting)
        Me.RibbonGroup26.Name = "RibbonGroup26"
        Me.RibbonGroup26.Text = "Goto Meeting"
        '
        'rbnJoinGoToMeeting
        '
        Me.rbnJoinGoToMeeting.LargeImage = CType(resources.GetObject("rbnJoinGoToMeeting.LargeImage"), System.Drawing.Image)
        Me.rbnJoinGoToMeeting.Name = "rbnJoinGoToMeeting"
        Me.rbnJoinGoToMeeting.SmallImage = CType(resources.GetObject("rbnJoinGoToMeeting.SmallImage"), System.Drawing.Image)
        Me.rbnJoinGoToMeeting.Text = "Join Goto Meeting"
        '
        'RibbonGroup18
        '
        Me.RibbonGroup18.Items.Add(Me.rbnHelpAbout)
        Me.RibbonGroup18.Items.Add(Me.rbnHelpAboutDirectory)
        Me.RibbonGroup18.Name = "RibbonGroup18"
        Me.RibbonGroup18.Text = "About"
        '
        'rbnHelpAbout
        '
        Me.rbnHelpAbout.Name = "rbnHelpAbout"
        Me.rbnHelpAbout.Text = "About SAW8"
        '
        'rbnHelpAboutDirectory
        '
        Me.rbnHelpAboutDirectory.Name = "rbnHelpAboutDirectory"
        Me.rbnHelpAboutDirectory.Text = "Directory = "
        '
        'RibbonGroup5
        '
        Me.RibbonGroup5.Items.Add(Me.rbnHelpMaster)
        Me.RibbonGroup5.Name = "RibbonGroup5"
        Me.RibbonGroup5.Text = "Help - F1"
        '
        'rbnHelpMaster
        '
        Me.rbnHelpMaster.LargeImage = CType(resources.GetObject("rbnHelpMaster.LargeImage"), System.Drawing.Image)
        Me.rbnHelpMaster.Name = "rbnHelpMaster"
        Me.rbnHelpMaster.SmallImage = CType(resources.GetObject("rbnHelpMaster.SmallImage"), System.Drawing.Image)
        Me.rbnHelpMaster.Text = "Quote Reports Help System"
        '
        'RibbonTopToolBar1
        '
        Me.RibbonTopToolBar1.Name = "RibbonTopToolBar1"
        '
        'tabQrt
        '
        Me.tabQrt.Controls.Add(Me.TabPage0)
        Me.tabQrt.Controls.Add(Me.TabPage1)
        Me.tabQrt.Controls.Add(Me.TabPage2)
        Me.tabQrt.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabQrt.ItemSize = New System.Drawing.Size(87, 19)
        Me.tabQrt.Location = New System.Drawing.Point(0, 179)
        Me.tabQrt.Name = "tabQrt"
        Me.tabQrt.Padding = New System.Drawing.Point(45, 3)
        Me.tabQrt.SelectedIndex = 0
        Me.tabQrt.Size = New System.Drawing.Size(1056, 683)
        Me.tabQrt.TabIndex = 118
        '
        'TabPage0
        '
        Me.TabPage0.Controls.Add(Me.fraDisplaySortSeq)
        Me.TabPage0.Controls.Add(Me.fraReport)
        Me.TabPage0.Location = New System.Drawing.Point(4, 23)
        Me.TabPage0.Name = "TabPage0"
        Me.TabPage0.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage0.Size = New System.Drawing.Size(1048, 656)
        Me.TabPage0.TabIndex = 0
        Me.TabPage0.Text = "Sort Sequence"
        Me.TabPage0.UseVisualStyleBackColor = True
        '
        'fraDisplaySortSeq
        '
        Me.fraDisplaySortSeq.Controls.Add(Me.txtPrimarySortSeq)
        Me.fraDisplaySortSeq.Controls.Add(Me.txtSecondarySort)
        Me.fraDisplaySortSeq.Controls.Add(Me.pnlPrimarySortSeq)
        Me.fraDisplaySortSeq.Controls.Add(Me.pnlTypeOfRpt)
        Me.fraDisplaySortSeq.Controls.Add(Me.pnlSecondarySort)
        Me.fraDisplaySortSeq.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraDisplaySortSeq.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraDisplaySortSeq.Location = New System.Drawing.Point(8, 6)
        Me.fraDisplaySortSeq.Name = "fraDisplaySortSeq"
        Me.fraDisplaySortSeq.Padding = New System.Windows.Forms.Padding(0)
        Me.fraDisplaySortSeq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDisplaySortSeq.Size = New System.Drawing.Size(825, 38)
        Me.fraDisplaySortSeq.TabIndex = 1
        Me.fraDisplaySortSeq.TabStop = False
        '
        'txtPrimarySortSeq
        '
        Me.txtPrimarySortSeq.AcceptsReturn = True
        Me.txtPrimarySortSeq.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrimarySortSeq.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrimarySortSeq.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrimarySortSeq.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrimarySortSeq.HideSelection = False
        Me.txtPrimarySortSeq.Location = New System.Drawing.Point(376, 13)
        Me.txtPrimarySortSeq.MaxLength = 0
        Me.txtPrimarySortSeq.Multiline = True
        Me.txtPrimarySortSeq.Name = "txtPrimarySortSeq"
        Me.txtPrimarySortSeq.ReadOnly = True
        Me.txtPrimarySortSeq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrimarySortSeq.Size = New System.Drawing.Size(129, 19)
        Me.txtPrimarySortSeq.TabIndex = 3
        Me.txtPrimarySortSeq.TabStop = False
        Me.txtPrimarySortSeq.Text = " "
        Me.txtPrimarySortSeq.Visible = False
        '
        'txtSecondarySort
        '
        Me.txtSecondarySort.AcceptsReturn = True
        Me.txtSecondarySort.BackColor = System.Drawing.SystemColors.Window
        Me.txtSecondarySort.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSecondarySort.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSecondarySort.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSecondarySort.HideSelection = False
        Me.txtSecondarySort.Location = New System.Drawing.Point(650, 13)
        Me.txtSecondarySort.MaxLength = 0
        Me.txtSecondarySort.Multiline = True
        Me.txtSecondarySort.Name = "txtSecondarySort"
        Me.txtSecondarySort.ReadOnly = True
        Me.txtSecondarySort.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSecondarySort.Size = New System.Drawing.Size(139, 19)
        Me.txtSecondarySort.TabIndex = 2
        Me.txtSecondarySort.TabStop = False
        Me.txtSecondarySort.Text = " "
        Me.txtSecondarySort.Visible = False
        '
        'pnlPrimarySortSeq
        '
        Me.pnlPrimarySortSeq.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlPrimarySortSeq.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlPrimarySortSeq.ForeColor = System.Drawing.Color.Black
        Me.pnlPrimarySortSeq.Location = New System.Drawing.Point(249, 13)
        Me.pnlPrimarySortSeq.Name = "pnlPrimarySortSeq"
        Me.pnlPrimarySortSeq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlPrimarySortSeq.Size = New System.Drawing.Size(121, 17)
        Me.pnlPrimarySortSeq.TabIndex = 6
        Me.pnlPrimarySortSeq.Text = "Primary Sort Sequence"
        Me.pnlPrimarySortSeq.Visible = False
        '
        'pnlTypeOfRpt
        '
        Me.pnlTypeOfRpt.BackColor = System.Drawing.Color.LemonChiffon
        Me.pnlTypeOfRpt.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlTypeOfRpt.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlTypeOfRpt.ForeColor = System.Drawing.Color.Black
        Me.pnlTypeOfRpt.Location = New System.Drawing.Point(16, 13)
        Me.pnlTypeOfRpt.Name = "pnlTypeOfRpt"
        Me.pnlTypeOfRpt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlTypeOfRpt.Size = New System.Drawing.Size(227, 17)
        Me.pnlTypeOfRpt.TabIndex = 5
        Me.pnlTypeOfRpt.Visible = False
        '
        'pnlSecondarySort
        '
        Me.pnlSecondarySort.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlSecondarySort.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlSecondarySort.ForeColor = System.Drawing.Color.Black
        Me.pnlSecondarySort.Location = New System.Drawing.Point(511, 13)
        Me.pnlSecondarySort.Name = "pnlSecondarySort"
        Me.pnlSecondarySort.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlSecondarySort.Size = New System.Drawing.Size(141, 17)
        Me.pnlSecondarySort.TabIndex = 4
        Me.pnlSecondarySort.Text = "Secondary Sort Sequence"
        Me.pnlSecondarySort.Visible = False
        '
        'fraReport
        '
        Me.fraReport.Controls.Add(Me.cboSortRealization)
        Me.fraReport.Controls.Add(Me.fraReportCmdSelection)
        Me.fraReport.Controls.Add(Me.fraSortSecondarySeq)
        Me.fraReport.Controls.Add(Me.fraSortPrimarySeq)
        Me.fraReport.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraReport.Location = New System.Drawing.Point(8, 52)
        Me.fraReport.Name = "fraReport"
        Me.fraReport.Padding = New System.Windows.Forms.Padding(0)
        Me.fraReport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraReport.Size = New System.Drawing.Size(982, 467)
        Me.fraReport.TabIndex = 7
        Me.fraReport.TabStop = False
        Me.fraReport.Text = "Report Options"
        '
        'cboSortRealization
        '
        Me.cboSortRealization.CheckOnClick = True
        Me.cboSortRealization.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSortRealization.FormattingEnabled = True
        Me.cboSortRealization.Items.AddRange(New Object() {"QuoteTO: Customer -C", "QuoteTO: Manufacturer - M", "QuoteTO: Other - O", "QuoteTO: Salesman/Customer - C", "Architect - A", "Engineer - E", "Ltg Designer - L", "Specifier - S", "Contractor - T", "Other - X", "ALL Specifiers", "Only Quotes with One MFG/Cust", "SESCO Job List Report"})
        Me.cboSortRealization.Location = New System.Drawing.Point(12, 85)
        Me.cboSortRealization.Name = "cboSortRealization"
        Me.cboSortRealization.Size = New System.Drawing.Size(273, 276)
        Me.cboSortRealization.TabIndex = 38
        Me.cboSortRealization.Visible = False
        '
        'fraReportCmdSelection
        '
        Me.fraReportCmdSelection.Controls.Add(Me.cmdReportProjShortage)
        Me.fraReportCmdSelection.Controls.Add(Me.cmdReportQuote)
        Me.fraReportCmdSelection.Controls.Add(Me.cmdReportTerrSpecCredit)
        Me.fraReportCmdSelection.Controls.Add(Me.cmdReportRealization)
        Me.fraReportCmdSelection.Controls.Add(Me.cmdReportOtherTypes)
        Me.fraReportCmdSelection.Controls.Add(Me.cmdReportLineItems)
        Me.fraReportCmdSelection.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraReportCmdSelection.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraReportCmdSelection.Location = New System.Drawing.Point(10, 14)
        Me.fraReportCmdSelection.Name = "fraReportCmdSelection"
        Me.fraReportCmdSelection.Padding = New System.Windows.Forms.Padding(0)
        Me.fraReportCmdSelection.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraReportCmdSelection.Size = New System.Drawing.Size(209, 350)
        Me.fraReportCmdSelection.TabIndex = 32
        Me.fraReportCmdSelection.TabStop = False
        Me.fraReportCmdSelection.Text = "Type of Report"
        '
        'cmdReportProjShortage
        '
        Me.cmdReportProjShortage.Location = New System.Drawing.Point(12, 251)
        Me.cmdReportProjShortage.Name = "cmdReportProjShortage"
        Me.cmdReportProjShortage.Size = New System.Drawing.Size(183, 50)
        Me.cmdReportProjShortage.TabIndex = 42
        Me.cmdReportProjShortage.Text = "Job Quote Commission Shortage Report"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdReportProjShortage, "Report on Orders not Shipped and Invoices not paid on Quotes. Matches Quote Code " &
        "to the Order Rep Quote Number Field.")
        Me.cmdReportProjShortage.UseVisualStyleBackColor = True
        Me.cmdReportProjShortage.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'fraSortSecondarySeq
        '
        Me.fraSortSecondarySeq.Controls.Add(Me.txtSortSecondarySeq)
        Me.fraSortSecondarySeq.Controls.Add(Me.cmdSecondarySeqCancel)
        Me.fraSortSecondarySeq.Controls.Add(Me.cboSortSecondarySeq)
        Me.fraSortSecondarySeq.Controls.Add(Me.cmdSecondarySeqContinue)
        Me.fraSortSecondarySeq.Controls.Add(Me.SSPanel4)
        Me.fraSortSecondarySeq.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraSortSecondarySeq.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraSortSecondarySeq.Location = New System.Drawing.Point(294, 16)
        Me.fraSortSecondarySeq.Name = "fraSortSecondarySeq"
        Me.fraSortSecondarySeq.Padding = New System.Windows.Forms.Padding(0)
        Me.fraSortSecondarySeq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSortSecondarySeq.Size = New System.Drawing.Size(275, 384)
        Me.fraSortSecondarySeq.TabIndex = 36
        Me.fraSortSecondarySeq.TabStop = False
        Me.fraSortSecondarySeq.Text = "Secondary Sort"
        Me.fraSortSecondarySeq.Visible = False
        '
        'txtSortSecondarySeq
        '
        Me.txtSortSecondarySeq.AcceptsReturn = True
        Me.txtSortSecondarySeq.BackColor = System.Drawing.SystemColors.Window
        Me.txtSortSecondarySeq.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSortSecondarySeq.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSortSecondarySeq.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSortSecondarySeq.Location = New System.Drawing.Point(11, 81)
        Me.txtSortSecondarySeq.MaxLength = 0
        Me.txtSortSecondarySeq.Name = "txtSortSecondarySeq"
        Me.txtSortSecondarySeq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSortSecondarySeq.Size = New System.Drawing.Size(239, 20)
        Me.txtSortSecondarySeq.TabIndex = 122
        Me.txtSortSecondarySeq.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtSortSecondarySeq, "KEEN or KEEN,GLOB,BULB to exclude -KEEN,-GLOB (If one is minus all must be minus." &
        ")")
        '
        'cmdSecondarySeqCancel
        '
        Me.cmdSecondarySeqCancel.Location = New System.Drawing.Point(89, 354)
        Me.cmdSecondarySeqCancel.Name = "cmdSecondarySeqCancel"
        Me.cmdSecondarySeqCancel.Size = New System.Drawing.Size(97, 23)
        Me.cmdSecondarySeqCancel.TabIndex = 42
        Me.cmdSecondarySeqCancel.Text = "Cancel<Back"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdSecondarySeqCancel, "Return to Previous Screen")
        Me.cmdSecondarySeqCancel.UseVisualStyleBackColor = True
        Me.cmdSecondarySeqCancel.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cboSortSecondarySeq
        '
        Me.cboSortSecondarySeq.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cboSortSecondarySeq.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSortSecondarySeq.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSortSecondarySeq.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSortSecondarySeq.ItemHeight = 14
        Me.cboSortSecondarySeq.Location = New System.Drawing.Point(11, 106)
        Me.cboSortSecondarySeq.Name = "cboSortSecondarySeq"
        Me.cboSortSecondarySeq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSortSecondarySeq.Size = New System.Drawing.Size(239, 228)
        Me.cboSortSecondarySeq.TabIndex = 121
        '
        'cmdSecondarySeqContinue
        '
        Me.cmdSecondarySeqContinue.Location = New System.Drawing.Point(8, 354)
        Me.cmdSecondarySeqContinue.Name = "cmdSecondarySeqContinue"
        Me.cmdSecondarySeqContinue.Size = New System.Drawing.Size(75, 23)
        Me.cmdSecondarySeqContinue.TabIndex = 41
        Me.cmdSecondarySeqContinue.Text = "Continue"
        Me.cmdSecondarySeqContinue.UseVisualStyleBackColor = True
        Me.cmdSecondarySeqContinue.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'SSPanel4
        '
        Me.SSPanel4.BackColor = System.Drawing.Color.LemonChiffon
        Me.SSPanel4.Cursor = System.Windows.Forms.Cursors.Default
        Me.SSPanel4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SSPanel4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SSPanel4.Location = New System.Drawing.Point(8, 16)
        Me.SSPanel4.Name = "SSPanel4"
        Me.SSPanel4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SSPanel4.Size = New System.Drawing.Size(267, 50)
        Me.SSPanel4.TabIndex = 40
        Me.SSPanel4.Text = "Select Secondary Sort sequence from the list and click Continue."
        '
        'fraSortPrimarySeq
        '
        Me.fraSortPrimarySeq.Controls.Add(Me.txtPrimarySort)
        Me.fraSortPrimarySeq.Controls.Add(Me.cboSortPrimarySeq)
        Me.fraSortPrimarySeq.Controls.Add(Me.cmdPrimarySeqCancel1)
        Me.fraSortPrimarySeq.Controls.Add(Me.cmdPrimarySeqContinue1)
        Me.fraSortPrimarySeq.Controls.Add(Me.SSPanel3)
        Me.fraSortPrimarySeq.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraSortPrimarySeq.ForeColor = System.Drawing.Color.Black
        Me.fraSortPrimarySeq.Location = New System.Drawing.Point(8, 16)
        Me.fraSortPrimarySeq.Name = "fraSortPrimarySeq"
        Me.fraSortPrimarySeq.Padding = New System.Windows.Forms.Padding(0)
        Me.fraSortPrimarySeq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSortPrimarySeq.Size = New System.Drawing.Size(270, 385)
        Me.fraSortPrimarySeq.TabIndex = 8
        Me.fraSortPrimarySeq.TabStop = False
        Me.fraSortPrimarySeq.Text = "Primary Sort Sequence"
        Me.fraSortPrimarySeq.Visible = False
        '
        'txtPrimarySort
        '
        Me.txtPrimarySort.AcceptsReturn = True
        Me.txtPrimarySort.BackColor = System.Drawing.SystemColors.Window
        Me.txtPrimarySort.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrimarySort.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrimarySort.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrimarySort.Location = New System.Drawing.Point(11, 78)
        Me.txtPrimarySort.MaxLength = 0
        Me.txtPrimarySort.Name = "txtPrimarySort"
        Me.txtPrimarySort.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrimarySort.Size = New System.Drawing.Size(239, 20)
        Me.txtPrimarySort.TabIndex = 124
        Me.txtPrimarySort.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtPrimarySort, "KEEN or KEEN,GLOB,BULB to exclude -KEEN,-GLOB (If one is minus all must be minus." &
        ")")
        '
        'cboSortPrimarySeq
        '
        Me.cboSortPrimarySeq.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.cboSortPrimarySeq.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSortPrimarySeq.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSortPrimarySeq.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSortPrimarySeq.ItemHeight = 14
        Me.cboSortPrimarySeq.Location = New System.Drawing.Point(11, 103)
        Me.cboSortPrimarySeq.Name = "cboSortPrimarySeq"
        Me.cboSortPrimarySeq.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSortPrimarySeq.Size = New System.Drawing.Size(239, 228)
        Me.cboSortPrimarySeq.TabIndex = 123
        '
        'cmdPrimarySeqCancel1
        '
        Me.cmdPrimarySeqCancel1.Location = New System.Drawing.Point(89, 354)
        Me.cmdPrimarySeqCancel1.Name = "cmdPrimarySeqCancel1"
        Me.cmdPrimarySeqCancel1.Size = New System.Drawing.Size(100, 23)
        Me.cmdPrimarySeqCancel1.TabIndex = 43
        Me.cmdPrimarySeqCancel1.Text = "Cancel<Back"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdPrimarySeqCancel1, "Return to Previous Screen")
        Me.cmdPrimarySeqCancel1.UseVisualStyleBackColor = True
        Me.cmdPrimarySeqCancel1.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cmdPrimarySeqContinue1
        '
        Me.cmdPrimarySeqContinue1.Location = New System.Drawing.Point(8, 354)
        Me.cmdPrimarySeqContinue1.Name = "cmdPrimarySeqContinue1"
        Me.cmdPrimarySeqContinue1.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrimarySeqContinue1.TabIndex = 13
        Me.cmdPrimarySeqContinue1.Text = "Continue"
        Me.cmdPrimarySeqContinue1.UseVisualStyleBackColor = True
        Me.cmdPrimarySeqContinue1.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'SSPanel3
        '
        Me.SSPanel3.BackColor = System.Drawing.Color.LemonChiffon
        Me.SSPanel3.Cursor = System.Windows.Forms.Cursors.Default
        Me.SSPanel3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SSPanel3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SSPanel3.Location = New System.Drawing.Point(8, 16)
        Me.SSPanel3.Name = "SSPanel3"
        Me.SSPanel3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SSPanel3.Size = New System.Drawing.Size(272, 50)
        Me.SSPanel3.TabIndex = 12
        Me.SSPanel3.Text = "Select a Primary sort sequence from the list and click Continue. Some options wil" &
    "l also offer a Secondary sort sequence."
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me._fdBranchCode)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.chkShowLatestCust)
        Me.TabPage1.Controls.Add(Me.chkBidJobsOnly)
        Me.TabPage1.Controls.Add(Me.fraQuoteReports)
        Me.TabPage1.Controls.Add(Me.fraSelectDate)
        Me.TabPage1.Controls.Add(Me.fraQuoteLineReports)
        Me.TabPage1.Controls.Add(Me.gbxSortSeq)
        Me.TabPage1.Location = New System.Drawing.Point(4, 23)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1048, 656)
        Me.TabPage1.TabIndex = 1
        Me.TabPage1.Text = "Select Criteria"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        '_fdBranchCode
        '
        Me._fdBranchCode.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me._fdBranchCode.AutoCompletion = True
        Me._fdBranchCode.Caption = ""
        Me._fdBranchCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me._fdBranchCode.ColumnWidth = 100
        Me._fdBranchCode.DeadAreaBackColor = System.Drawing.Color.Empty
        Me._fdBranchCode.DisplayMember = "BranchCode"
        Me._fdBranchCode.DropDownWidth = 200
        Me._fdBranchCode.EditorBackColor = System.Drawing.SystemColors.Window
        Me._fdBranchCode.EditorFont = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._fdBranchCode.EditorForeColor = System.Drawing.Color.Blue
        Me._fdBranchCode.ExtendRightColumn = True
        Me._fdBranchCode.Images.Add(CType(resources.GetObject("_fdBranchCode.Images"), System.Drawing.Image))
        Me._fdBranchCode.ItemHeight = 20
        Me._fdBranchCode.Location = New System.Drawing.Point(704, 64)
        Me._fdBranchCode.MatchEntryTimeout = CType(2000, Long)
        Me._fdBranchCode.MaxDropDownItems = CType(30, Short)
        Me._fdBranchCode.MaxLength = 50
        Me._fdBranchCode.MouseCursor = System.Windows.Forms.Cursors.Default
        Me._fdBranchCode.Name = "_fdBranchCode"
        Me._fdBranchCode.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me._fdBranchCode.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me._fdBranchCode.Size = New System.Drawing.Size(194, 21)
        Me._fdBranchCode.SuperBack = True
        Me._fdBranchCode.TabIndex = 182
        Me._fdBranchCode.Text = "ALL"
        Me._fdBranchCode.ToolTip = "Enter Branch Code or Codes seperated by a Comma (NOR,EST,WST)"
        Me._fdBranchCode.ValueMember = "BranchCode"
        Me._fdBranchCode.PropBag = resources.GetString("_fdBranchCode.PropBag")
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(586, 67)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(80, 13)
        Me.Label8.TabIndex = 183
        Me.Label8.Text = "Branch Code(s)"
        '
        'chkShowLatestCust
        '
        Me.chkShowLatestCust.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkShowLatestCust.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowLatestCust.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkShowLatestCust.Location = New System.Drawing.Point(370, 80)
        Me.chkShowLatestCust.Name = "chkShowLatestCust"
        Me.chkShowLatestCust.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkShowLatestCust.Size = New System.Drawing.Size(358, 27)
        Me.chkShowLatestCust.TabIndex = 119
        Me.chkShowLatestCust.Text = "Only Show the Latest Quote per Customer Quoted"
        Me.C1SuperTooltip1.SetToolTip(Me.chkShowLatestCust, "Only Show the Latest Quote To per Customer")
        Me.chkShowLatestCust.UseVisualStyleBackColor = False
        Me.chkShowLatestCust.Visible = False
        '
        'chkBidJobsOnly
        '
        Me.chkBidJobsOnly.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBidJobsOnly.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBidJobsOnly.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBidJobsOnly.Location = New System.Drawing.Point(369, 60)
        Me.chkBidJobsOnly.Name = "chkBidJobsOnly"
        Me.chkBidJobsOnly.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBidJobsOnly.Size = New System.Drawing.Size(212, 22)
        Me.chkBidJobsOnly.TabIndex = 117
        Me.chkBidJobsOnly.Text = "Bid Board Jobs Only "
        Me.C1SuperTooltip1.SetToolTip(Me.chkBidJobsOnly, "Only Select Quotes with BidBoard Checked.")
        Me.chkBidJobsOnly.UseVisualStyleBackColor = False
        '
        'fraQuoteReports
        '
        Me.fraQuoteReports.AutoSize = True
        Me.fraQuoteReports.Controls.Add(Me.chkIncludeNotesLineItems)
        Me.fraQuoteReports.Controls.Add(Me.chkIncludeSLSSPlit)
        Me.fraQuoteReports.Controls.Add(Me.chkIncludeSpecifiers)
        Me.fraQuoteReports.Controls.Add(Me.cboTypeCustomer)
        Me.fraQuoteReports.Controls.Add(Me.lblTypeCustomer)
        Me.fraQuoteReports.Controls.Add(Me.chkBrandReport)
        Me.fraQuoteReports.Controls.Add(Me.cbospeccross)
        Me.fraQuoteReports.Controls.Add(Me.cmdok1)
        Me.fraQuoteReports.Controls.Add(Me.cboLinesInclude)
        Me.fraQuoteReports.Controls.Add(Me.cmdCancel1)
        Me.fraQuoteReports.Controls.Add(Me.cmdResetDefaults1)
        Me.fraQuoteReports.Controls.Add(Me.chkPrtPlanLines)
        Me.fraQuoteReports.Controls.Add(Me.pnlQutRealCode)
        Me.fraQuoteReports.Controls.Add(Me.cboTypeofJob)
        Me.fraQuoteReports.Controls.Add(Me.Label4)
        Me.fraQuoteReports.Controls.Add(Me.txtSelectCode)
        Me.fraQuoteReports.Controls.Add(Me.pnlSpecifierCode)
        Me.fraQuoteReports.Controls.Add(Me.fraFinishReports)
        Me.fraQuoteReports.Controls.Add(Me.pnlSpecCross)
        Me.fraQuoteReports.Controls.Add(Me.txtCSRofCust)
        Me.fraQuoteReports.Controls.Add(Me.lblRetrieval)
        Me.fraQuoteReports.Controls.Add(Me.pnlLotUnit)
        Me.fraQuoteReports.Controls.Add(Me.txtSpecifierCode)
        Me.fraQuoteReports.Controls.Add(Me.txtJobNameSS)
        Me.fraQuoteReports.Controls.Add(Me.lblStatus)
        Me.fraQuoteReports.Controls.Add(Me.txtSalesman)
        Me.fraQuoteReports.Controls.Add(Me.lblSalesman)
        Me.fraQuoteReports.Controls.Add(Me.lblStartQuote)
        Me.fraQuoteReports.Controls.Add(Me.cboLotUnit)
        Me.fraQuoteReports.Controls.Add(Me.txtMktSegment)
        Me.fraQuoteReports.Controls.Add(Me.pnlSlsSplits)
        Me.fraQuoteReports.Controls.Add(Me.txtRetrieval)
        Me.fraQuoteReports.Controls.Add(Me.cboStockJob)
        Me.fraQuoteReports.Controls.Add(Me.lblEndQuote)
        Me.fraQuoteReports.Controls.Add(Me.pnlCSR)
        Me.fraQuoteReports.Controls.Add(Me.txtCity)
        Me.fraQuoteReports.Controls.Add(Me.pnlSltCode)
        Me.fraQuoteReports.Controls.Add(Me.txtStatus)
        Me.fraQuoteReports.Controls.Add(Me.txtCSR)
        Me.fraQuoteReports.Controls.Add(Me.txtState)
        Me.fraQuoteReports.Controls.Add(Me.pnlStkJob)
        Me.fraQuoteReports.Controls.Add(Me.pnlCSRdist)
        Me.fraQuoteReports.Controls.Add(Me.pnlQuoteToSls)
        Me.fraQuoteReports.Controls.Add(Me.txtSlsSplit)
        Me.fraQuoteReports.Controls.Add(Me.txtEndQuoteAmt)
        Me.fraQuoteReports.Controls.Add(Me.lblJobName)
        Me.fraQuoteReports.Controls.Add(Me.txtLastChgBy)
        Me.fraQuoteReports.Controls.Add(Me.txtQuoteToSls)
        Me.fraQuoteReports.Controls.Add(Me.PnlLastChgBy)
        Me.fraQuoteReports.Controls.Add(Me.txtStartQuoteAmt)
        Me.fraQuoteReports.Controls.Add(Me.pnlMktSeg)
        Me.fraQuoteReports.Controls.Add(Me.pnlCity)
        Me.fraQuoteReports.Controls.Add(Me.txtQutRealCode)
        Me.fraQuoteReports.Controls.Add(Me.pnlState)
        Me.fraQuoteReports.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraQuoteReports.Location = New System.Drawing.Point(8, 113)
        Me.fraQuoteReports.Name = "fraQuoteReports"
        Me.fraQuoteReports.Size = New System.Drawing.Size(850, 508)
        Me.fraQuoteReports.TabIndex = 115
        Me.fraQuoteReports.TabStop = False
        '
        'chkIncludeNotesLineItems
        '
        Me.chkIncludeNotesLineItems.AutoSize = True
        Me.chkIncludeNotesLineItems.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIncludeNotesLineItems.Location = New System.Drawing.Point(228, 391)
        Me.chkIncludeNotesLineItems.Name = "chkIncludeNotesLineItems"
        Me.chkIncludeNotesLineItems.Size = New System.Drawing.Size(91, 18)
        Me.chkIncludeNotesLineItems.TabIndex = 432
        Me.chkIncludeNotesLineItems.Text = "Include Notes"
        Me.chkIncludeNotesLineItems.UseVisualStyleBackColor = True
        Me.chkIncludeNotesLineItems.Visible = False
        '
        'chkIncludeSLSSPlit
        '
        Me.chkIncludeSLSSPlit.AutoSize = True
        Me.chkIncludeSLSSPlit.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIncludeSLSSPlit.Location = New System.Drawing.Point(228, 367)
        Me.chkIncludeSLSSPlit.Name = "chkIncludeSLSSPlit"
        Me.chkIncludeSLSSPlit.Size = New System.Drawing.Size(112, 18)
        Me.chkIncludeSLSSPlit.TabIndex = 431
        Me.chkIncludeSLSSPlit.Text = "Include SLS Splits"
        Me.chkIncludeSLSSPlit.UseVisualStyleBackColor = True
        Me.chkIncludeSLSSPlit.Visible = False
        '
        'chkIncludeSpecifiers
        '
        Me.chkIncludeSpecifiers.AutoSize = True
        Me.chkIncludeSpecifiers.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIncludeSpecifiers.Location = New System.Drawing.Point(228, 343)
        Me.chkIncludeSpecifiers.Name = "chkIncludeSpecifiers"
        Me.chkIncludeSpecifiers.Size = New System.Drawing.Size(112, 18)
        Me.chkIncludeSpecifiers.TabIndex = 430
        Me.chkIncludeSpecifiers.Text = "Include Specifiers"
        Me.chkIncludeSpecifiers.UseVisualStyleBackColor = True
        Me.chkIncludeSpecifiers.Visible = False
        '
        'cboTypeCustomer
        '
        Me.cboTypeCustomer.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.cboTypeCustomer.AutoCompletion = True
        Me.cboTypeCustomer.Caption = ""
        Me.cboTypeCustomer.CaptionHeight = 17
        Me.cboTypeCustomer.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboTypeCustomer.ColumnCaptionHeight = 17
        Me.cboTypeCustomer.ColumnFooterHeight = 17
        Me.cboTypeCustomer.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.cboTypeCustomer.DropDownWidth = 200
        Me.cboTypeCustomer.EditorBackColor = System.Drawing.SystemColors.Window
        Me.cboTypeCustomer.EditorFont = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTypeCustomer.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTypeCustomer.ExtendRightColumn = True
        Me.cboTypeCustomer.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTypeCustomer.Images.Add(CType(resources.GetObject("cboTypeCustomer.Images"), System.Drawing.Image))
        Me.cboTypeCustomer.ItemHeight = 15
        Me.cboTypeCustomer.LimitToList = True
        Me.cboTypeCustomer.Location = New System.Drawing.Point(152, 319)
        Me.cboTypeCustomer.MatchEntryTimeout = CType(2000, Long)
        Me.cboTypeCustomer.MaxDropDownItems = CType(30, Short)
        Me.cboTypeCustomer.MaxLength = 10
        Me.cboTypeCustomer.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.cboTypeCustomer.Name = "cboTypeCustomer"
        Me.cboTypeCustomer.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.cboTypeCustomer.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.cboTypeCustomer.Size = New System.Drawing.Size(68, 21)
        Me.cboTypeCustomer.SuperBack = True
        Me.cboTypeCustomer.TabIndex = 429
        Me.cboTypeCustomer.Text = "A"
        Me.cboTypeCustomer.Visible = False
        Me.cboTypeCustomer.PropBag = resources.GetString("cboTypeCustomer.PropBag")
        '
        'lblTypeCustomer
        '
        Me.lblTypeCustomer.AutoSize = True
        Me.lblTypeCustomer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTypeCustomer.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTypeCustomer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTypeCustomer.Location = New System.Drawing.Point(17, 321)
        Me.lblTypeCustomer.Name = "lblTypeCustomer"
        Me.lblTypeCustomer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTypeCustomer.Size = New System.Drawing.Size(92, 14)
        Me.lblTypeCustomer.TabIndex = 428
        Me.lblTypeCustomer.Text = "Type of Customer"
        Me.C1SuperTooltip1.SetToolTip(Me.lblTypeCustomer, "A = All Jobs, Q = Quotes,  S = Spec Credit,  P = Planned Proj, T = Submittal Proj" &
        ",  O = Other")
        Me.lblTypeCustomer.Visible = False
        '
        'chkBrandReport
        '
        Me.chkBrandReport.AutoSize = True
        Me.chkBrandReport.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBrandReport.Location = New System.Drawing.Point(20, 376)
        Me.chkBrandReport.Name = "chkBrandReport"
        Me.chkBrandReport.Size = New System.Drawing.Size(125, 18)
        Me.chkBrandReport.TabIndex = 427
        Me.chkBrandReport.Text = "Brand Mfg Reporting"
        Me.C1SuperTooltip1.SetToolTip(Me.chkBrandReport, "Brand/Mfg Line Item Reporting")
        Me.chkBrandReport.UseVisualStyleBackColor = True
        Me.chkBrandReport.Visible = False
        '
        'cbospeccross
        '
        Me.cbospeccross.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.cbospeccross.AutoCompletion = True
        Me.cbospeccross.AutoDropDown = True
        Me.cbospeccross.AutoSelect = True
        Me.cbospeccross.Caption = ""
        Me.cbospeccross.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cbospeccross.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.cbospeccross.DropDownWidth = 150
        Me.cbospeccross.EditorBackColor = System.Drawing.SystemColors.Window
        Me.cbospeccross.EditorFont = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbospeccross.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.cbospeccross.ExtendRightColumn = True
        Me.cbospeccross.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbospeccross.Images.Add(CType(resources.GetObject("cbospeccross.Images"), System.Drawing.Image))
        Me.cbospeccross.ItemHeight = 20
        Me.cbospeccross.Location = New System.Drawing.Point(348, 38)
        Me.cbospeccross.MatchEntryTimeout = CType(2000, Long)
        Me.cbospeccross.MaxDropDownItems = CType(30, Short)
        Me.cbospeccross.MaxLength = 3
        Me.cbospeccross.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.cbospeccross.Name = "cbospeccross"
        Me.cbospeccross.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.cbospeccross.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.cbospeccross.Size = New System.Drawing.Size(73, 21)
        Me.cbospeccross.SuperBack = True
        Me.cbospeccross.TabIndex = 426
        Me.cbospeccross.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.cbospeccross, "Select a Code from your Dropdown Defaults. Enter ALL for all codes.")
        Me.cbospeccross.PropBag = resources.GetString("cbospeccross.PropBag")
        '
        'cmdok1
        '
        Me.cmdok1.Location = New System.Drawing.Point(437, 374)
        Me.cmdok1.Name = "cmdok1"
        Me.cmdok1.Size = New System.Drawing.Size(140, 23)
        Me.cmdok1.TabIndex = 120
        Me.cmdok1.Text = "Run Report/Select Data"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdok1, "This will Select and Sort the Records for this Report.")
        Me.cmdok1.UseVisualStyleBackColor = True
        Me.cmdok1.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cmdCancel1
        '
        Me.cmdCancel1.Location = New System.Drawing.Point(583, 374)
        Me.cmdCancel1.Name = "cmdCancel1"
        Me.cmdCancel1.Size = New System.Drawing.Size(109, 23)
        Me.cmdCancel1.TabIndex = 118
        Me.cmdCancel1.Text = "Cancel<Back"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdCancel1, "Return to Previous Screen")
        Me.cmdCancel1.UseVisualStyleBackColor = True
        Me.cmdCancel1.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cmdResetDefaults1
        '
        Me.cmdResetDefaults1.Location = New System.Drawing.Point(698, 374)
        Me.cmdResetDefaults1.Name = "cmdResetDefaults1"
        Me.cmdResetDefaults1.Size = New System.Drawing.Size(90, 23)
        Me.cmdResetDefaults1.TabIndex = 119
        Me.cmdResetDefaults1.Text = "Reset Defaults"
        Me.cmdResetDefaults1.UseVisualStyleBackColor = True
        Me.cmdResetDefaults1.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'pnlQutRealCode
        '
        Me.pnlQutRealCode.AutoSize = True
        Me.pnlQutRealCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlQutRealCode.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlQutRealCode.ForeColor = System.Drawing.Color.Black
        Me.pnlQutRealCode.Location = New System.Drawing.Point(17, 15)
        Me.pnlQutRealCode.Name = "pnlQutRealCode"
        Me.pnlQutRealCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlQutRealCode.Size = New System.Drawing.Size(32, 14)
        Me.pnlQutRealCode.TabIndex = 92
        Me.pnlQutRealCode.Text = "Code"
        Me.pnlQutRealCode.Visible = False
        '
        'pnlSpecifierCode
        '
        Me.pnlSpecifierCode.AutoSize = True
        Me.pnlSpecifierCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlSpecifierCode.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlSpecifierCode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.pnlSpecifierCode.Location = New System.Drawing.Point(19, 127)
        Me.pnlSpecifierCode.Name = "pnlSpecifierCode"
        Me.pnlSpecifierCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlSpecifierCode.Size = New System.Drawing.Size(78, 14)
        Me.pnlSpecifierCode.TabIndex = 112
        Me.pnlSpecifierCode.Text = "Specifier Code"
        '
        'fraFinishReports
        '
        Me.fraFinishReports.Controls.Add(Me.ChkSpecifiersCustInCols)
        Me.fraFinishReports.Controls.Add(Me.ChkQuoteNoSpecifiers)
        Me.fraFinishReports.Controls.Add(Me.chkBlankLine)
        Me.fraFinishReports.Controls.Add(Me.ChkExtendByProb)
        Me.fraFinishReports.Controls.Add(Me.chkNotes)
        Me.fraFinishReports.Controls.Add(Me.ChkSpecifiers)
        Me.fraFinishReports.Controls.Add(Me.chkMfgBreakdown)
        Me.fraFinishReports.Controls.Add(Me.chkCustomerBreakdown)
        Me.fraFinishReports.Controls.Add(Me.chkBranchReport)
        Me.fraFinishReports.Controls.Add(Me.chkSalesmanPerPage)
        Me.fraFinishReports.Controls.Add(Me.chkDetailTotal)
        Me.fraFinishReports.Controls.Add(Me.chkIncludeCommDolPer)
        Me.fraFinishReports.Controls.Add(Me.chkExcludeDuplicates)
        Me.fraFinishReports.Controls.Add(Me.chkSlsFromHeader)
        Me.fraFinishReports.Controls.Add(Me.chkExportAllExcel)
        Me.fraFinishReports.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraFinishReports.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraFinishReports.Location = New System.Drawing.Point(431, 4)
        Me.fraFinishReports.Name = "fraFinishReports"
        Me.fraFinishReports.Padding = New System.Windows.Forms.Padding(0)
        Me.fraFinishReports.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraFinishReports.Size = New System.Drawing.Size(338, 361)
        Me.fraFinishReports.TabIndex = 45
        Me.fraFinishReports.TabStop = False
        '
        'ChkQuoteNoSpecifiers
        '
        Me.ChkQuoteNoSpecifiers.AutoSize = True
        Me.ChkQuoteNoSpecifiers.Cursor = System.Windows.Forms.Cursors.Default
        Me.ChkQuoteNoSpecifiers.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkQuoteNoSpecifiers.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ChkQuoteNoSpecifiers.Location = New System.Drawing.Point(6, 328)
        Me.ChkQuoteNoSpecifiers.Name = "ChkQuoteNoSpecifiers"
        Me.ChkQuoteNoSpecifiers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ChkQuoteNoSpecifiers.Size = New System.Drawing.Size(177, 18)
        Me.ChkQuoteNoSpecifiers.TabIndex = 35
        Me.ChkQuoteNoSpecifiers.Text = "Only Quotes with no Specifiers"
        Me.C1SuperTooltip1.SetToolTip(Me.ChkQuoteNoSpecifiers, "Only Report Quotes with no Specifiers ")
        Me.ChkQuoteNoSpecifiers.UseVisualStyleBackColor = False
        '
        'chkSalesmanPerPage
        '
        Me.chkSalesmanPerPage.AutoSize = True
        Me.chkSalesmanPerPage.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSalesmanPerPage.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSalesmanPerPage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSalesmanPerPage.Location = New System.Drawing.Point(118, 97)
        Me.chkSalesmanPerPage.Name = "chkSalesmanPerPage"
        Me.chkSalesmanPerPage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSalesmanPerPage.Size = New System.Drawing.Size(87, 18)
        Me.chkSalesmanPerPage.TabIndex = 27
        Me.chkSalesmanPerPage.Text = "1 Code/Page"
        Me.chkSalesmanPerPage.UseVisualStyleBackColor = False
        '
        'pnlSpecCross
        '
        Me.pnlSpecCross.AutoSize = True
        Me.pnlSpecCross.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlSpecCross.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlSpecCross.ForeColor = System.Drawing.Color.Black
        Me.pnlSpecCross.Location = New System.Drawing.Point(239, 43)
        Me.pnlSpecCross.Name = "pnlSpecCross"
        Me.pnlSpecCross.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlSpecCross.Size = New System.Drawing.Size(64, 14)
        Me.pnlSpecCross.TabIndex = 93
        Me.pnlSpecCross.Text = "Spec/Cross"
        '
        'lblRetrieval
        '
        Me.lblRetrieval.AutoSize = True
        Me.lblRetrieval.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRetrieval.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRetrieval.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRetrieval.Location = New System.Drawing.Point(19, 159)
        Me.lblRetrieval.Name = "lblRetrieval"
        Me.lblRetrieval.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRetrieval.Size = New System.Drawing.Size(77, 14)
        Me.lblRetrieval.TabIndex = 111
        Me.lblRetrieval.Text = "Retrieval Code"
        '
        'pnlLotUnit
        '
        Me.pnlLotUnit.AutoSize = True
        Me.pnlLotUnit.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlLotUnit.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlLotUnit.ForeColor = System.Drawing.Color.Black
        Me.pnlLotUnit.Location = New System.Drawing.Point(239, 127)
        Me.pnlLotUnit.Name = "pnlLotUnit"
        Me.pnlLotUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlLotUnit.Size = New System.Drawing.Size(43, 14)
        Me.pnlLotUnit.TabIndex = 94
        Me.pnlLotUnit.Text = "Lot/Unit"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblStatus.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStatus.Location = New System.Drawing.Point(16, 101)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStatus.Size = New System.Drawing.Size(66, 14)
        Me.lblStatus.TabIndex = 110
        Me.lblStatus.Text = "Status Code"
        '
        'lblSalesman
        '
        Me.lblSalesman.AutoSize = True
        Me.lblSalesman.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSalesman.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSalesman.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSalesman.Location = New System.Drawing.Point(19, 184)
        Me.lblSalesman.Name = "lblSalesman"
        Me.lblSalesman.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSalesman.Size = New System.Drawing.Size(94, 14)
        Me.lblSalesman.TabIndex = 109
        Me.lblSalesman.Text = "SLS Code-Header"
        '
        'lblStartQuote
        '
        Me.lblStartQuote.AutoSize = True
        Me.lblStartQuote.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblStartQuote.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartQuote.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStartQuote.Location = New System.Drawing.Point(17, 46)
        Me.lblStartQuote.Name = "lblStartQuote"
        Me.lblStartQuote.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStartQuote.Size = New System.Drawing.Size(97, 14)
        Me.lblStartQuote.TabIndex = 108
        Me.lblStartQuote.Text = "Starting Quote Amt"
        '
        'pnlSlsSplits
        '
        Me.pnlSlsSplits.AutoSize = True
        Me.pnlSlsSplits.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlSlsSplits.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlSlsSplits.ForeColor = System.Drawing.Color.Black
        Me.pnlSlsSplits.Location = New System.Drawing.Point(239, 218)
        Me.pnlSlsSplits.Name = "pnlSlsSplits"
        Me.pnlSlsSplits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlSlsSplits.Size = New System.Drawing.Size(69, 14)
        Me.pnlSlsSplits.TabIndex = 99
        Me.pnlSlsSplits.Text = "SLS 1-4 Split"
        Me.C1SuperTooltip1.SetToolTip(Me.pnlSlsSplits, "All SLS on SLS 1-4 Split")
        '
        'lblEndQuote
        '
        Me.lblEndQuote.AutoSize = True
        Me.lblEndQuote.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEndQuote.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndQuote.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEndQuote.Location = New System.Drawing.Point(18, 71)
        Me.lblEndQuote.Name = "lblEndQuote"
        Me.lblEndQuote.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEndQuote.Size = New System.Drawing.Size(92, 14)
        Me.lblEndQuote.TabIndex = 107
        Me.lblEndQuote.Text = "Ending Quote Amt"
        '
        'pnlCSR
        '
        Me.pnlCSR.AutoSize = True
        Me.pnlCSR.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlCSR.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlCSR.ForeColor = System.Drawing.Color.Black
        Me.pnlCSR.Location = New System.Drawing.Point(239, 98)
        Me.pnlCSR.Name = "pnlCSR"
        Me.pnlCSR.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlCSR.Size = New System.Drawing.Size(67, 14)
        Me.pnlCSR.TabIndex = 100
        Me.pnlCSR.Text = "CSR - Quote"
        '
        'pnlSltCode
        '
        Me.pnlSltCode.AutoSize = True
        Me.pnlSltCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlSltCode.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlSltCode.ForeColor = System.Drawing.Color.Black
        Me.pnlSltCode.Location = New System.Drawing.Point(239, 71)
        Me.pnlSltCode.Name = "pnlSltCode"
        Me.pnlSltCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlSltCode.Size = New System.Drawing.Size(65, 14)
        Me.pnlSltCode.TabIndex = 101
        Me.pnlSltCode.Text = "Select Code"
        '
        'pnlStkJob
        '
        Me.pnlStkJob.AutoSize = True
        Me.pnlStkJob.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlStkJob.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlStkJob.ForeColor = System.Drawing.Color.Black
        Me.pnlStkJob.Location = New System.Drawing.Point(239, 164)
        Me.pnlStkJob.Name = "pnlStkJob"
        Me.pnlStkJob.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlStkJob.Size = New System.Drawing.Size(54, 14)
        Me.pnlStkJob.TabIndex = 102
        Me.pnlStkJob.Text = "Stock/Job"
        '
        'pnlCSRdist
        '
        Me.pnlCSRdist.AutoSize = True
        Me.pnlCSRdist.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlCSRdist.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlCSRdist.ForeColor = System.Drawing.Color.Black
        Me.pnlCSRdist.Location = New System.Drawing.Point(239, 191)
        Me.pnlCSRdist.Name = "pnlCSRdist"
        Me.pnlCSRdist.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlCSRdist.Size = New System.Drawing.Size(84, 14)
        Me.pnlCSRdist.TabIndex = 105
        Me.pnlCSRdist.Text = "CSR - Customer"
        Me.pnlCSRdist.Visible = False
        '
        'pnlQuoteToSls
        '
        Me.pnlQuoteToSls.AutoSize = True
        Me.pnlQuoteToSls.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlQuoteToSls.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlQuoteToSls.ForeColor = System.Drawing.Color.Black
        Me.pnlQuoteToSls.Location = New System.Drawing.Point(239, 18)
        Me.pnlQuoteToSls.Name = "pnlQuoteToSls"
        Me.pnlQuoteToSls.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlQuoteToSls.Size = New System.Drawing.Size(68, 14)
        Me.pnlQuoteToSls.TabIndex = 103
        Me.pnlQuoteToSls.Text = "Quote To Sls"
        '
        'lblJobName
        '
        Me.lblJobName.AutoSize = True
        Me.lblJobName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblJobName.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblJobName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblJobName.Location = New System.Drawing.Point(240, 273)
        Me.lblJobName.Name = "lblJobName"
        Me.lblJobName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblJobName.Size = New System.Drawing.Size(123, 14)
        Me.lblJobName.TabIndex = 106
        Me.lblJobName.Text = "Job Name Search String"
        '
        'PnlLastChgBy
        '
        Me.PnlLastChgBy.AutoSize = True
        Me.PnlLastChgBy.Cursor = System.Windows.Forms.Cursors.Default
        Me.PnlLastChgBy.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlLastChgBy.ForeColor = System.Drawing.Color.Black
        Me.PnlLastChgBy.Location = New System.Drawing.Point(19, 208)
        Me.PnlLastChgBy.Name = "PnlLastChgBy"
        Me.PnlLastChgBy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PnlLastChgBy.Size = New System.Drawing.Size(84, 14)
        Me.PnlLastChgBy.TabIndex = 98
        Me.PnlLastChgBy.Text = "Last Change By"
        '
        'pnlMktSeg
        '
        Me.pnlMktSeg.AutoSize = True
        Me.pnlMktSeg.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlMktSeg.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlMktSeg.ForeColor = System.Drawing.Color.Black
        Me.pnlMktSeg.Location = New System.Drawing.Point(17, 297)
        Me.pnlMktSeg.Name = "pnlMktSeg"
        Me.pnlMktSeg.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlMktSeg.Size = New System.Drawing.Size(84, 14)
        Me.pnlMktSeg.TabIndex = 97
        Me.pnlMktSeg.Text = "Market Segment"
        '
        'pnlCity
        '
        Me.pnlCity.AutoSize = True
        Me.pnlCity.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlCity.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlCity.ForeColor = System.Drawing.Color.Black
        Me.pnlCity.Location = New System.Drawing.Point(17, 269)
        Me.pnlCity.Name = "pnlCity"
        Me.pnlCity.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlCity.Size = New System.Drawing.Size(25, 14)
        Me.pnlCity.TabIndex = 95
        Me.pnlCity.Text = "City"
        '
        'pnlState
        '
        Me.pnlState.AutoSize = True
        Me.pnlState.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlState.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlState.ForeColor = System.Drawing.Color.Black
        Me.pnlState.Location = New System.Drawing.Point(19, 240)
        Me.pnlState.Name = "pnlState"
        Me.pnlState.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlState.Size = New System.Drawing.Size(32, 14)
        Me.pnlState.TabIndex = 96
        Me.pnlState.Text = "State"
        '
        'fraSelectDate
        '
        Me.fraSelectDate.Controls.Add(Me.DTPicker1EndBid)
        Me.fraSelectDate.Controls.Add(Me.DTPicker1EndEntry)
        Me.fraSelectDate.Controls.Add(Me.DTPicker1StartBid)
        Me.fraSelectDate.Controls.Add(Me.ChkCheckBidDates)
        Me.fraSelectDate.Controls.Add(Me.chkBlankBidDates)
        Me.fraSelectDate.Controls.Add(Me.lblEndBid)
        Me.fraSelectDate.Controls.Add(Me.lblStartBid)
        Me.fraSelectDate.Controls.Add(Me.lblEndEntry)
        Me.fraSelectDate.Controls.Add(Me.txtStartEntry)
        Me.fraSelectDate.Controls.Add(Me.DTPickerStartEntry)
        Me.fraSelectDate.Controls.Add(Me.lblStartEntry)
        Me.fraSelectDate.Controls.Add(Me.txtEndEntry)
        Me.fraSelectDate.Controls.Add(Me.txtEndBid)
        Me.fraSelectDate.Controls.Add(Me.txtStartBid)
        Me.fraSelectDate.Location = New System.Drawing.Point(11, 4)
        Me.fraSelectDate.Name = "fraSelectDate"
        Me.fraSelectDate.Size = New System.Drawing.Size(355, 115)
        Me.fraSelectDate.TabIndex = 113
        Me.fraSelectDate.TabStop = False
        Me.fraSelectDate.Text = "Date Range - MMDDYY"
        '
        'DTPicker1EndBid
        '
        Me.DTPicker1EndBid.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPicker1EndBid.Location = New System.Drawing.Point(216, 46)
        Me.DTPicker1EndBid.Name = "DTPicker1EndBid"
        Me.DTPicker1EndBid.Size = New System.Drawing.Size(129, 20)
        Me.DTPicker1EndBid.TabIndex = 436
        '
        'DTPicker1EndEntry
        '
        Me.DTPicker1EndEntry.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPicker1EndEntry.Location = New System.Drawing.Point(218, 17)
        Me.DTPicker1EndEntry.Name = "DTPicker1EndEntry"
        Me.DTPicker1EndEntry.Size = New System.Drawing.Size(129, 20)
        Me.DTPicker1EndEntry.TabIndex = 435
        '
        'DTPicker1StartBid
        '
        Me.DTPicker1StartBid.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPicker1StartBid.Location = New System.Drawing.Point(58, 46)
        Me.DTPicker1StartBid.Name = "DTPicker1StartBid"
        Me.DTPicker1StartBid.Size = New System.Drawing.Size(129, 20)
        Me.DTPicker1StartBid.TabIndex = 434
        '
        'ChkCheckBidDates
        '
        Me.ChkCheckBidDates.Cursor = System.Windows.Forms.Cursors.Default
        Me.ChkCheckBidDates.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkCheckBidDates.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ChkCheckBidDates.Location = New System.Drawing.Point(10, 73)
        Me.ChkCheckBidDates.Name = "ChkCheckBidDates"
        Me.ChkCheckBidDates.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ChkCheckBidDates.Size = New System.Drawing.Size(298, 22)
        Me.ChkCheckBidDates.TabIndex = 82
        Me.ChkCheckBidDates.Text = "Check Bid Dates when Selecting Quotes"
        Me.C1SuperTooltip1.SetToolTip(Me.ChkCheckBidDates, "Include Quotes with Blank Bid Dates")
        Me.ChkCheckBidDates.UseVisualStyleBackColor = False
        '
        'lblEndBid
        '
        Me.lblEndBid.AutoSize = True
        Me.lblEndBid.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEndBid.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndBid.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEndBid.Location = New System.Drawing.Point(193, 46)
        Me.lblEndBid.Name = "lblEndBid"
        Me.lblEndBid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEndBid.Size = New System.Drawing.Size(16, 14)
        Me.lblEndBid.TabIndex = 77
        Me.lblEndBid.Text = "to"
        '
        'lblStartBid
        '
        Me.lblStartBid.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblStartBid.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartBid.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStartBid.Location = New System.Drawing.Point(4, 40)
        Me.lblStartBid.Name = "lblStartBid"
        Me.lblStartBid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStartBid.Size = New System.Drawing.Size(73, 20)
        Me.lblStartBid.TabIndex = 78
        Me.lblStartBid.Text = "Bid      From"
        '
        'lblEndEntry
        '
        Me.lblEndEntry.AutoSize = True
        Me.lblEndEntry.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEndEntry.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndEntry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEndEntry.Location = New System.Drawing.Point(193, 22)
        Me.lblEndEntry.Name = "lblEndEntry"
        Me.lblEndEntry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEndEntry.Size = New System.Drawing.Size(16, 14)
        Me.lblEndEntry.TabIndex = 79
        Me.lblEndEntry.Text = "to"
        '
        'txtStartEntry
        '
        Me.txtStartEntry.AcceptsReturn = True
        Me.txtStartEntry.BackColor = System.Drawing.SystemColors.Window
        Me.txtStartEntry.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStartEntry.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStartEntry.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStartEntry.Location = New System.Drawing.Point(93, 22)
        Me.txtStartEntry.MaxLength = 6
        Me.txtStartEntry.Name = "txtStartEntry"
        Me.txtStartEntry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStartEntry.Size = New System.Drawing.Size(73, 20)
        Me.txtStartEntry.TabIndex = 72
        Me.txtStartEntry.Text = "ALL"
        Me.txtStartEntry.Visible = False
        '
        'DTPickerStartEntry
        '
        Me.DTPickerStartEntry.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTPickerStartEntry.Location = New System.Drawing.Point(58, 17)
        Me.DTPickerStartEntry.Name = "DTPickerStartEntry"
        Me.DTPickerStartEntry.Size = New System.Drawing.Size(129, 20)
        Me.DTPickerStartEntry.TabIndex = 433
        '
        'lblStartEntry
        '
        Me.lblStartEntry.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblStartEntry.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartEntry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStartEntry.Location = New System.Drawing.Point(4, 16)
        Me.lblStartEntry.Name = "lblStartEntry"
        Me.lblStartEntry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStartEntry.Size = New System.Drawing.Size(73, 20)
        Me.lblStartEntry.TabIndex = 80
        Me.lblStartEntry.Text = "Entry    From"
        '
        'txtEndEntry
        '
        Me.txtEndEntry.AcceptsReturn = True
        Me.txtEndEntry.BackColor = System.Drawing.SystemColors.Window
        Me.txtEndEntry.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEndEntry.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEndEntry.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEndEntry.Location = New System.Drawing.Point(225, 22)
        Me.txtEndEntry.MaxLength = 6
        Me.txtEndEntry.Name = "txtEndEntry"
        Me.txtEndEntry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEndEntry.Size = New System.Drawing.Size(73, 20)
        Me.txtEndEntry.TabIndex = 71
        Me.txtEndEntry.Text = "ALL"
        Me.txtEndEntry.Visible = False
        '
        'txtEndBid
        '
        Me.txtEndBid.AcceptsReturn = True
        Me.txtEndBid.BackColor = System.Drawing.SystemColors.Window
        Me.txtEndBid.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtEndBid.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEndBid.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtEndBid.Location = New System.Drawing.Point(225, 46)
        Me.txtEndBid.MaxLength = 6
        Me.txtEndBid.Name = "txtEndBid"
        Me.txtEndBid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtEndBid.Size = New System.Drawing.Size(73, 20)
        Me.txtEndBid.TabIndex = 69
        Me.txtEndBid.Text = "ALL"
        Me.txtEndBid.Visible = False
        '
        'txtStartBid
        '
        Me.txtStartBid.AcceptsReturn = True
        Me.txtStartBid.BackColor = System.Drawing.SystemColors.Window
        Me.txtStartBid.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStartBid.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStartBid.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStartBid.Location = New System.Drawing.Point(93, 46)
        Me.txtStartBid.MaxLength = 6
        Me.txtStartBid.Name = "txtStartBid"
        Me.txtStartBid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStartBid.Size = New System.Drawing.Size(73, 20)
        Me.txtStartBid.TabIndex = 70
        Me.txtStartBid.Text = "ALL"
        Me.txtStartBid.Visible = False
        '
        'fraQuoteLineReports
        '
        Me.fraQuoteLineReports.Controls.Add(Me.chkHaveMFGCode)
        Me.fraQuoteLineReports.Controls.Add(Me.chkPrtNTElines)
        Me.fraQuoteLineReports.Controls.Add(Me.cmdOK2)
        Me.fraQuoteLineReports.Controls.Add(Me.cmdCancel2)
        Me.fraQuoteLineReports.Controls.Add(Me.cmdResetDefaults2)
        Me.fraQuoteLineReports.Controls.Add(Me.GroupBox2)
        Me.fraQuoteLineReports.Controls.Add(Me.ChkTotalsOnly)
        Me.fraQuoteLineReports.Controls.Add(Me.GroupBox1)
        Me.fraQuoteLineReports.Controls.Add(Me.fraUnitorExtended)
        Me.fraQuoteLineReports.Controls.Add(Me.fraSalesorCost)
        Me.fraQuoteLineReports.Controls.Add(Me.fraQtIncludeCommission)
        Me.fraQuoteLineReports.Location = New System.Drawing.Point(11, 123)
        Me.fraQuoteLineReports.Name = "fraQuoteLineReports"
        Me.fraQuoteLineReports.Size = New System.Drawing.Size(717, 367)
        Me.fraQuoteLineReports.TabIndex = 116
        Me.fraQuoteLineReports.TabStop = False
        '
        'chkHaveMFGCode
        '
        Me.chkHaveMFGCode.Checked = True
        Me.chkHaveMFGCode.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHaveMFGCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkHaveMFGCode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkHaveMFGCode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkHaveMFGCode.Location = New System.Drawing.Point(9, 190)
        Me.chkHaveMFGCode.Name = "chkHaveMFGCode"
        Me.chkHaveMFGCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkHaveMFGCode.Size = New System.Drawing.Size(223, 19)
        Me.chkHaveMFGCode.TabIndex = 435
        Me.chkHaveMFGCode.Text = "Only Lines With MFG Code."
        Me.C1SuperTooltip1.SetToolTip(Me.chkHaveMFGCode, "Lines must have a MFG code to Print on Report.")
        Me.chkHaveMFGCode.UseVisualStyleBackColor = False
        '
        'chkPrtNTElines
        '
        Me.chkPrtNTElines.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPrtNTElines.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPrtNTElines.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPrtNTElines.Location = New System.Drawing.Point(9, 170)
        Me.chkPrtNTElines.Name = "chkPrtNTElines"
        Me.chkPrtNTElines.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPrtNTElines.Size = New System.Drawing.Size(161, 19)
        Me.chkPrtNTElines.TabIndex = 434
        Me.chkPrtNTElines.Text = "Add NTE, Etc (Note) Lines."
        Me.C1SuperTooltip1.SetToolTip(Me.chkPrtNTElines, "Check to add NTE, NPE, SUB, TAX, Etc. Lines to Report.")
        Me.chkPrtNTElines.UseVisualStyleBackColor = False
        '
        'cmdOK2
        '
        Me.cmdOK2.Location = New System.Drawing.Point(356, 254)
        Me.cmdOK2.Name = "cmdOK2"
        Me.cmdOK2.Size = New System.Drawing.Size(140, 23)
        Me.cmdOK2.TabIndex = 123
        Me.cmdOK2.Text = "Run Report/Select Data"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdOK2, "This will Select and Sort the Records for this Report.")
        Me.cmdOK2.UseVisualStyleBackColor = True
        Me.cmdOK2.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cmdCancel2
        '
        Me.cmdCancel2.Location = New System.Drawing.Point(502, 254)
        Me.cmdCancel2.Name = "cmdCancel2"
        Me.cmdCancel2.Size = New System.Drawing.Size(112, 23)
        Me.cmdCancel2.TabIndex = 121
        Me.cmdCancel2.Text = "Cancel<Back"
        Me.cmdCancel2.UseVisualStyleBackColor = True
        Me.cmdCancel2.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'cmdResetDefaults2
        '
        Me.cmdResetDefaults2.Location = New System.Drawing.Point(620, 254)
        Me.cmdResetDefaults2.Name = "cmdResetDefaults2"
        Me.cmdResetDefaults2.Size = New System.Drawing.Size(90, 23)
        Me.cmdResetDefaults2.TabIndex = 122
        Me.cmdResetDefaults2.Text = "Reset Defaults"
        Me.cmdResetDefaults2.UseVisualStyleBackColor = True
        Me.cmdResetDefaults2.VisualStyleBaseStyle = C1.Win.C1Input.VisualStyle.Office2007Blue
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkShowCustomers)
        Me.GroupBox2.Controls.Add(Me.chkUseSpecifierCode)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.txtCustomerCodeLine)
        Me.GroupBox2.Controls.Add(Me.txtPrcCode)
        Me.GroupBox2.Controls.Add(Me.pnlPrcCode)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.txtSpecCross)
        Me.GroupBox2.Controls.Add(Me.TxtSingleCatNum)
        Me.GroupBox2.Controls.Add(Me.TxtSearchString)
        Me.GroupBox2.Controls.Add(Me.txtMfgLine)
        Me.GroupBox2.Controls.Add(Me.PnlCatNum)
        Me.GroupBox2.Controls.Add(Me.PnlCatSrch)
        Me.GroupBox2.Controls.Add(Me.PnlMfg)
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(286, 14)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(386, 236)
        Me.GroupBox2.TabIndex = 86
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Line Item Selection"
        '
        'chkShowCustomers
        '
        Me.chkShowCustomers.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkShowCustomers.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowCustomers.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkShowCustomers.Location = New System.Drawing.Point(6, 184)
        Me.chkShowCustomers.Name = "chkShowCustomers"
        Me.chkShowCustomers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkShowCustomers.Size = New System.Drawing.Size(276, 20)
        Me.chkShowCustomers.TabIndex = 438
        Me.chkShowCustomers.Text = "Show Customers Quoted "
        Me.C1SuperTooltip1.SetToolTip(Me.chkShowCustomers, "Show Each Customer Quoted  on Each Line Item")
        Me.chkShowCustomers.UseVisualStyleBackColor = False
        '
        'chkUseSpecifierCode
        '
        Me.chkUseSpecifierCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkUseSpecifierCode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseSpecifierCode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkUseSpecifierCode.Location = New System.Drawing.Point(8, 211)
        Me.chkUseSpecifierCode.Name = "chkUseSpecifierCode"
        Me.chkUseSpecifierCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkUseSpecifierCode.Size = New System.Drawing.Size(276, 20)
        Me.chkUseSpecifierCode.TabIndex = 437
        Me.chkUseSpecifierCode.Text = "Show Specifier Codes"
        Me.C1SuperTooltip1.SetToolTip(Me.chkUseSpecifierCode, "Substitute Specifier Code for Customer Code.")
        Me.chkUseSpecifierCode.UseVisualStyleBackColor = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(7, 158)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(93, 14)
        Me.Label7.TabIndex = 436
        Me.Label7.Text = "Customer Code    "
        '
        'txtCustomerCodeLine
        '
        Me.txtCustomerCodeLine.AcceptsReturn = True
        Me.txtCustomerCodeLine.BackColor = System.Drawing.SystemColors.Window
        Me.txtCustomerCodeLine.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCustomerCodeLine.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustomerCodeLine.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCustomerCodeLine.Location = New System.Drawing.Point(177, 155)
        Me.txtCustomerCodeLine.MaxLength = 50
        Me.txtCustomerCodeLine.Name = "txtCustomerCodeLine"
        Me.txtCustomerCodeLine.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCustomerCodeLine.Size = New System.Drawing.Size(79, 20)
        Me.txtCustomerCodeLine.TabIndex = 435
        Me.txtCustomerCodeLine.Text = "ALL"
        Me.C1SuperTooltip1.SetToolTip(Me.txtCustomerCodeLine, "For Multiple Customers - Separate each Code by a Comma. Select Specifier or Custo" &
        "mer below.")
        '
        'pnlPrcCode
        '
        Me.pnlPrcCode.AutoSize = True
        Me.pnlPrcCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.pnlPrcCode.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlPrcCode.ForeColor = System.Drawing.Color.Black
        Me.pnlPrcCode.Location = New System.Drawing.Point(6, 136)
        Me.pnlPrcCode.Name = "pnlPrcCode"
        Me.pnlPrcCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.pnlPrcCode.Size = New System.Drawing.Size(89, 14)
        Me.pnlPrcCode.TabIndex = 87
        Me.pnlPrcCode.Text = "Price Code          "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(6, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(91, 14)
        Me.Label3.TabIndex = 65
        Me.Label3.Text = "Spec/Cross         "
        '
        'PnlCatNum
        '
        Me.PnlCatNum.AutoSize = True
        Me.PnlCatNum.Cursor = System.Windows.Forms.Cursors.Default
        Me.PnlCatNum.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlCatNum.ForeColor = System.Drawing.Color.Black
        Me.PnlCatNum.Location = New System.Drawing.Point(6, 23)
        Me.PnlCatNum.Name = "PnlCatNum"
        Me.PnlCatNum.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PnlCatNum.Size = New System.Drawing.Size(127, 14)
        Me.PnlCatNum.TabIndex = 60
        Me.PnlCatNum.Text = "Single Catalog Number    "
        '
        'PnlCatSrch
        '
        Me.PnlCatSrch.AutoSize = True
        Me.PnlCatSrch.Cursor = System.Windows.Forms.Cursors.Default
        Me.PnlCatSrch.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlCatSrch.ForeColor = System.Drawing.Color.Black
        Me.PnlCatSrch.Location = New System.Drawing.Point(6, 52)
        Me.PnlCatSrch.Name = "PnlCatSrch"
        Me.PnlCatSrch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PnlCatSrch.Size = New System.Drawing.Size(172, 14)
        Me.PnlCatSrch.TabIndex = 61
        Me.PnlCatSrch.Text = "Catalog # Search String                 "
        '
        'PnlMfg
        '
        Me.PnlMfg.AutoSize = True
        Me.PnlMfg.Cursor = System.Windows.Forms.Cursors.Default
        Me.PnlMfg.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlMfg.ForeColor = System.Drawing.Color.Black
        Me.PnlMfg.Location = New System.Drawing.Point(6, 80)
        Me.PnlMfg.Name = "PnlMfg"
        Me.PnlMfg.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PnlMfg.Size = New System.Drawing.Size(75, 14)
        Me.PnlMfg.TabIndex = 62
        Me.PnlMfg.Text = "MFG Code      "
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtSlsTerr)
        Me.GroupBox1.Controls.Add(Me.PnlSls)
        Me.GroupBox1.Controls.Add(Me.txtLastChgByLine)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtRetr)
        Me.GroupBox1.Controls.Add(Me.PnlRet)
        Me.GroupBox1.Controls.Add(Me.txtStat)
        Me.GroupBox1.Controls.Add(Me.PnlStatus)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(9, 14)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(271, 135)
        Me.GroupBox1.TabIndex = 82
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Quote Selection"
        '
        'PnlSls
        '
        Me.PnlSls.AutoSize = True
        Me.PnlSls.Cursor = System.Windows.Forms.Cursors.Default
        Me.PnlSls.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlSls.ForeColor = System.Drawing.Color.Black
        Me.PnlSls.Location = New System.Drawing.Point(6, 108)
        Me.PnlSls.Name = "PnlSls"
        Me.PnlSls.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PnlSls.Size = New System.Drawing.Size(90, 14)
        Me.PnlSls.TabIndex = 90
        Me.PnlSls.Text = "SLSQ Quote SLS"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(6, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(90, 14)
        Me.Label1.TabIndex = 88
        Me.Label1.Text = "Last Change By  "
        '
        'PnlRet
        '
        Me.PnlRet.AutoSize = True
        Me.PnlRet.Cursor = System.Windows.Forms.Cursors.Default
        Me.PnlRet.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlRet.ForeColor = System.Drawing.Color.Black
        Me.PnlRet.Location = New System.Drawing.Point(7, 52)
        Me.PnlRet.Name = "PnlRet"
        Me.PnlRet.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PnlRet.Size = New System.Drawing.Size(113, 14)
        Me.PnlRet.TabIndex = 43
        Me.PnlRet.Text = "Retrieval Code            "
        '
        'PnlStatus
        '
        Me.PnlStatus.AutoSize = True
        Me.PnlStatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.PnlStatus.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlStatus.ForeColor = System.Drawing.Color.Black
        Me.PnlStatus.Location = New System.Drawing.Point(7, 24)
        Me.PnlStatus.Name = "PnlStatus"
        Me.PnlStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.PnlStatus.Size = New System.Drawing.Size(96, 14)
        Me.PnlStatus.TabIndex = 41
        Me.PnlStatus.Text = "Status Code          "
        '
        'fraUnitorExtended
        '
        Me.fraUnitorExtended.Controls.Add(Me.optUnitOrExtended_Unit)
        Me.fraUnitorExtended.Controls.Add(Me.optUnitOrExtended_Extd)
        Me.fraUnitorExtended.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraUnitorExtended.ForeColor = System.Drawing.Color.Black
        Me.fraUnitorExtended.Location = New System.Drawing.Point(38, 211)
        Me.fraUnitorExtended.Name = "fraUnitorExtended"
        Me.fraUnitorExtended.Padding = New System.Windows.Forms.Padding(0)
        Me.fraUnitorExtended.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraUnitorExtended.Size = New System.Drawing.Size(147, 65)
        Me.fraUnitorExtended.TabIndex = 65
        Me.fraUnitorExtended.TabStop = False
        Me.fraUnitorExtended.Visible = False
        '
        'fraSalesorCost
        '
        Me.fraSalesorCost.Controls.Add(Me.optSalesorCost_Cost)
        Me.fraSalesorCost.Controls.Add(Me.optoptSalesorCost_Sales)
        Me.fraSalesorCost.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraSalesorCost.ForeColor = System.Drawing.Color.Black
        Me.fraSalesorCost.Location = New System.Drawing.Point(238, 283)
        Me.fraSalesorCost.Name = "fraSalesorCost"
        Me.fraSalesorCost.Padding = New System.Windows.Forms.Padding(0)
        Me.fraSalesorCost.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSalesorCost.Size = New System.Drawing.Size(147, 65)
        Me.fraSalesorCost.TabIndex = 68
        Me.fraSalesorCost.TabStop = False
        Me.fraSalesorCost.Visible = False
        '
        'fraQtIncludeCommission
        '
        Me.fraQtIncludeCommission.Controls.Add(Me.chkIncludeCommDoll)
        Me.fraQtIncludeCommission.Controls.Add(Me.chkIncludeCommPer)
        Me.fraQtIncludeCommission.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraQtIncludeCommission.ForeColor = System.Drawing.Color.Black
        Me.fraQtIncludeCommission.Location = New System.Drawing.Point(389, 284)
        Me.fraQtIncludeCommission.Name = "fraQtIncludeCommission"
        Me.fraQtIncludeCommission.Padding = New System.Windows.Forms.Padding(0)
        Me.fraQtIncludeCommission.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraQtIncludeCommission.Size = New System.Drawing.Size(147, 65)
        Me.fraQtIncludeCommission.TabIndex = 81
        Me.fraQtIncludeCommission.TabStop = False
        Me.fraQtIncludeCommission.Visible = False
        '
        'gbxSortSeq
        '
        Me.gbxSortSeq.Controls.Add(Me.txtSortSeq)
        Me.gbxSortSeq.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxSortSeq.Location = New System.Drawing.Point(363, 6)
        Me.gbxSortSeq.Name = "gbxSortSeq"
        Me.gbxSortSeq.Size = New System.Drawing.Size(464, 54)
        Me.gbxSortSeq.TabIndex = 114
        Me.gbxSortSeq.TabStop = False
        Me.gbxSortSeq.Text = "Report Sort Sequence"
        '
        'txtSortSeq
        '
        Me.txtSortSeq.Location = New System.Drawing.Point(9, 19)
        Me.txtSortSeq.Name = "txtSortSeq"
        Me.txtSortSeq.Size = New System.Drawing.Size(458, 19)
        Me.txtSortSeq.TabIndex = 0
        Me.txtSortSeq.Text = "Report Sort Sequence"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.SplitContainer1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 23)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(1048, 656)
        Me.TabPage2.TabIndex = 2
        Me.TabPage2.Text = "View Grid"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.cmdBackViewGrid)
        Me.SplitContainer1.Panel1.Controls.Add(Me.gbxSortSeqV)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CmdShowColstoPrt)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CmdRunReport)
        Me.SplitContainer1.Panel1.Padding = New System.Windows.Forms.Padding(20)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Panel1)
        Me.SplitContainer1.Panel2.Margin = New System.Windows.Forms.Padding(20)
        Me.SplitContainer1.Panel2.Padding = New System.Windows.Forms.Padding(15)
        Me.SplitContainer1.Size = New System.Drawing.Size(1048, 656)
        Me.SplitContainer1.SplitterDistance = 91
        Me.SplitContainer1.TabIndex = 0
        '
        'cmdBackViewGrid
        '
        Me.cmdBackViewGrid.Location = New System.Drawing.Point(779, 16)
        Me.cmdBackViewGrid.Name = "cmdBackViewGrid"
        Me.cmdBackViewGrid.Size = New System.Drawing.Size(125, 47)
        Me.cmdBackViewGrid.TabIndex = 116
        Me.cmdBackViewGrid.Text = "Cancel < Back"
        Me.C1SuperTooltip1.SetToolTip(Me.cmdBackViewGrid, "Go back to Select Criteria")
        Me.cmdBackViewGrid.UseVisualStyleBackColor = True
        '
        'gbxSortSeqV
        '
        Me.gbxSortSeqV.Controls.Add(Me.txtSortSeqCriteria)
        Me.gbxSortSeqV.Controls.Add(Me.txtSortSeqV)
        Me.gbxSortSeqV.Location = New System.Drawing.Point(216, 4)
        Me.gbxSortSeqV.Name = "gbxSortSeqV"
        Me.gbxSortSeqV.Size = New System.Drawing.Size(557, 74)
        Me.gbxSortSeqV.TabIndex = 115
        Me.gbxSortSeqV.TabStop = False
        Me.gbxSortSeqV.Text = "Report Sort Sequence"
        '
        'txtSortSeqCriteria
        '
        Me.txtSortSeqCriteria.Location = New System.Drawing.Point(6, 41)
        Me.txtSortSeqCriteria.Name = "txtSortSeqCriteria"
        Me.txtSortSeqCriteria.Size = New System.Drawing.Size(551, 20)
        Me.txtSortSeqCriteria.TabIndex = 1
        Me.txtSortSeqCriteria.Text = "Report Selection Criteria"
        '
        'txtSortSeqV
        '
        Me.txtSortSeqV.Location = New System.Drawing.Point(6, 19)
        Me.txtSortSeqV.Name = "txtSortSeqV"
        Me.txtSortSeqV.Size = New System.Drawing.Size(552, 20)
        Me.txtSortSeqV.TabIndex = 0
        Me.txtSortSeqV.Text = "Report Sort Sequence"
        '
        'CmdShowColstoPrt
        '
        Me.CmdShowColstoPrt.Location = New System.Drawing.Point(33, 8)
        Me.CmdShowColstoPrt.Name = "CmdShowColstoPrt"
        Me.CmdShowColstoPrt.Size = New System.Drawing.Size(177, 27)
        Me.CmdShowColstoPrt.TabIndex = 1
        Me.CmdShowColstoPrt.Text = "Show Columns to Print"
        Me.CmdShowColstoPrt.UseVisualStyleBackColor = True
        '
        'CmdRunReport
        '
        Me.CmdRunReport.Location = New System.Drawing.Point(33, 35)
        Me.CmdRunReport.Name = "CmdRunReport"
        Me.CmdRunReport.Size = New System.Drawing.Size(177, 47)
        Me.CmdRunReport.TabIndex = 0
        Me.CmdRunReport.Text = "Run Report (Print/Preview/Export)"
        Me.C1SuperTooltip1.SetToolTip(Me.CmdRunReport, "Click Here to Generate the Report for Preview, Print and Export.")
        Me.CmdRunReport.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.tgln)
        Me.Panel1.Controls.Add(Me.tgSpecReg)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.tglnDIST)
        Me.Panel1.Controls.Add(Me.tgrDIST)
        Me.Panel1.Controls.Add(Me.tgQhDIST)
        Me.Panel1.Controls.Add(Me.tgr)
        Me.Panel1.Controls.Add(Me.tgQh)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(15, 15)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Padding = New System.Windows.Forms.Padding(5)
        Me.Panel1.Size = New System.Drawing.Size(1018, 531)
        Me.Panel1.TabIndex = 401
        '
        'tgln
        '
        Me.tgln.AllowArrows = False
        Me.tgln.AllowFilter = False
        Me.tgln.AllowRowSelect = False
        Me.tgln.AllowSort = False
        Me.tgln.Caption = "Line Items - Click Run Report to see final report selection"
        Me.tgln.CaptionHeight = 19
        Me.tgln.ColumnFooters = True
        Me.tgln.DataSource = Me.QuoteLinesBindingSource
        Me.tgln.GroupByCaption = "Drag a column header here to group by that column"
        Me.tgln.Images.Add(CType(resources.GetObject("tgln.Images"), System.Drawing.Image))
        Me.tgln.Location = New System.Drawing.Point(7, 99)
        Me.tgln.Name = "tgln"
        Me.tgln.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tgln.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tgln.PreviewInfo.ZoomFactor = 75.0R
        Me.tgln.PrintInfo.PageSettings = CType(resources.GetObject("tgln.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tgln.RowHeight = 13
        Me.tgln.Size = New System.Drawing.Size(996, 71)
        Me.tgln.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.GridNavigation
        Me.tgln.TabIndex = 400
        Me.tgln.Text = "C1TrueDBGrid1"
        Me.tgln.UseCompatibleTextRendering = False
        Me.tgln.Visible = False
        Me.tgln.PropBag = resources.GetString("tgln.PropBag")
        '
        'QuoteLinesBindingSource
        '
        Me.QuoteLinesBindingSource.DataMember = "quotelines"
        Me.QuoteLinesBindingSource.DataSource = Me.DsSaw8
        '
        'DsSaw8
        '
        Me.DsSaw8.DataSetName = "dsSaw8"
        Me.DsSaw8.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'tgSpecReg
        '
        Me.tgSpecReg.AllowArrows = False
        Me.tgSpecReg.AllowFilter = False
        Me.tgSpecReg.AllowRowSelect = False
        Me.tgSpecReg.AllowSort = False
        Me.tgSpecReg.Caption = "Spec Credit - Click Run Report to see final report selection"
        Me.tgSpecReg.CaptionHeight = 19
        Me.tgSpecReg.ColumnFooters = True
        Me.tgSpecReg.DataSource = Me.SpecRegFollowUpBindingSource
        Me.tgSpecReg.GroupByCaption = "Drag a column header here to group by that column"
        Me.tgSpecReg.Images.Add(CType(resources.GetObject("tgSpecReg.Images"), System.Drawing.Image))
        Me.tgSpecReg.Location = New System.Drawing.Point(9, 177)
        Me.tgSpecReg.Name = "tgSpecReg"
        Me.tgSpecReg.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tgSpecReg.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tgSpecReg.PreviewInfo.ZoomFactor = 75.0R
        Me.tgSpecReg.PrintInfo.PageSettings = CType(resources.GetObject("tgSpecReg.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tgSpecReg.RowHeight = 13
        Me.tgSpecReg.Size = New System.Drawing.Size(884, 70)
        Me.tgSpecReg.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.GridNavigation
        Me.tgSpecReg.TabIndex = 405
        Me.tgSpecReg.Text = "C1TrueDBGrid1"
        Me.tgSpecReg.UseCompatibleTextRendering = False
        Me.tgSpecReg.Visible = False
        Me.tgSpecReg.PropBag = resources.GetString("tgSpecReg.PropBag")
        '
        'SpecRegFollowUpBindingSource
        '
        Me.SpecRegFollowUpBindingSource.DataMember = "SpecRegFollowUp"
        Me.SpecRegFollowUpBindingSource.DataSource = Me.DsSaw8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(15, 254)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(433, 22)
        Me.Label5.TabIndex = 404
        Me.Label5.Text = "DIST GRIDS - USE TO CREATE LAYOUT FILES"
        Me.Label5.Visible = False
        '
        'tglnDIST
        '
        Me.tglnDIST.AllowArrows = False
        Me.tglnDIST.AllowFilter = False
        Me.tglnDIST.AllowRowSelect = False
        Me.tglnDIST.AllowSort = False
        Me.tglnDIST.Caption = "Line Items - Click Run Report to see final report selection"
        Me.tglnDIST.CaptionHeight = 19
        Me.tglnDIST.ColumnFooters = True
        Me.tglnDIST.DataSource = Me.QuoteLinesBindingSource
        Me.tglnDIST.GroupByCaption = "Drag a column header here to group by that column"
        Me.tglnDIST.Images.Add(CType(resources.GetObject("tglnDIST.Images"), System.Drawing.Image))
        Me.tglnDIST.Location = New System.Drawing.Point(14, 438)
        Me.tglnDIST.Name = "tglnDIST"
        Me.tglnDIST.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tglnDIST.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tglnDIST.PreviewInfo.ZoomFactor = 75.0R
        Me.tglnDIST.PrintInfo.PageSettings = CType(resources.GetObject("tglnDIST.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tglnDIST.RowHeight = 13
        Me.tglnDIST.Size = New System.Drawing.Size(996, 50)
        Me.tglnDIST.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.GridNavigation
        Me.tglnDIST.TabIndex = 403
        Me.tglnDIST.Text = "C1TrueDBGrid1"
        Me.tglnDIST.UseCompatibleTextRendering = False
        Me.tglnDIST.Visible = False
        Me.tglnDIST.PropBag = resources.GetString("tglnDIST.PropBag")
        '
        'tgrDIST
        '
        Me.tgrDIST.AllowSort = False
        Me.tgrDIST.Caption = "Quote To Grid - Click Run Report to see final report selection"
        Me.tgrDIST.CaptionHeight = 19
        Me.tgrDIST.DataSource = Me.QuoteRealLUBindingSource
        Me.tgrDIST.GroupByCaption = "Drag a column header here to group by that column"
        Me.tgrDIST.Images.Add(CType(resources.GetObject("tgrDIST.Images"), System.Drawing.Image))
        Me.tgrDIST.Location = New System.Drawing.Point(11, 362)
        Me.tgrDIST.Name = "tgrDIST"
        Me.tgrDIST.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tgrDIST.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tgrDIST.PreviewInfo.ZoomFactor = 75.0R
        Me.tgrDIST.PrintInfo.PageSettings = CType(resources.GetObject("tgrDIST.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tgrDIST.RowHeight = 17
        Me.tgrDIST.Size = New System.Drawing.Size(996, 55)
        Me.tgrDIST.TabIndex = 402
        Me.tgrDIST.Text = "C1TrueDBGrid1"
        Me.tgrDIST.UseCompatibleTextRendering = False
        Me.tgrDIST.Visible = False
        Me.tgrDIST.PropBag = resources.GetString("tgrDIST.PropBag")
        '
        'QuoteRealLUBindingSource
        '
        Me.QuoteRealLUBindingSource.DataMember = "QuoteRealLU"
        Me.QuoteRealLUBindingSource.DataSource = Me.DsSaw8
        '
        'tgQhDIST
        '
        Me.tgQhDIST.AllowSort = False
        Me.tgQhDIST.Caption = "Quote Master Grid - Click Run Report to see final report selection"
        Me.tgQhDIST.CaptionHeight = 19
        Me.tgQhDIST.DataSource = Me.QutNotesBindingSource
        Me.tgQhDIST.GroupByAreaVisible = False
        Me.tgQhDIST.GroupByCaption = "Drag a column header here to group by that column"
        Me.tgQhDIST.Images.Add(CType(resources.GetObject("tgQhDIST.Images"), System.Drawing.Image))
        Me.tgQhDIST.Location = New System.Drawing.Point(14, 282)
        Me.tgQhDIST.Margin = New System.Windows.Forms.Padding(30)
        Me.tgQhDIST.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.FloatingEditor
        Me.tgQhDIST.Name = "tgQhDIST"
        Me.tgQhDIST.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tgQhDIST.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tgQhDIST.PreviewInfo.ZoomFactor = 75.0R
        Me.tgQhDIST.PrintInfo.PageSettings = CType(resources.GetObject("tgQhDIST.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tgQhDIST.RowHeight = 17
        Me.tgQhDIST.Size = New System.Drawing.Size(993, 55)
        Me.tgQhDIST.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.GridNavigation
        Me.tgQhDIST.TabIndex = 401
        Me.tgQhDIST.Text = "C1TrueDBGrid1"
        Me.tgQhDIST.UseCompatibleTextRendering = False
        Me.tgQhDIST.Visible = False
        Me.tgQhDIST.PropBag = resources.GetString("tgQhDIST.PropBag")
        '
        'QutNotesBindingSource
        '
        Me.QutNotesBindingSource.DataMember = "qutnotes"
        Me.QutNotesBindingSource.DataSource = Me.DsSaw8
        '
        'tgr
        '
        Me.tgr.AllowSort = False
        Me.tgr.Caption = "Quote To Grid - Click Run Report to see final report selection"
        Me.tgr.CaptionHeight = 19
        Me.tgr.DataSource = Me.QuoteRealLUBindingSource
        Me.tgr.GroupByCaption = "Drag a column header here to group by that column"
        Me.tgr.Images.Add(CType(resources.GetObject("tgr.Images"), System.Drawing.Image))
        Me.tgr.Location = New System.Drawing.Point(7, 70)
        Me.tgr.Name = "tgr"
        Me.tgr.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tgr.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tgr.PreviewInfo.ZoomFactor = 75.0R
        Me.tgr.PrintInfo.PageSettings = CType(resources.GetObject("tgr.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tgr.RowHeight = 17
        Me.tgr.Size = New System.Drawing.Size(996, 50)
        Me.tgr.TabIndex = 1
        Me.tgr.Text = "C1TrueDBGrid1"
        Me.tgr.UseCompatibleTextRendering = False
        Me.tgr.Visible = False
        Me.tgr.PropBag = resources.GetString("tgr.PropBag")
        '
        'tgQh
        '
        Me.tgQh.AllowSort = False
        Me.tgQh.Caption = "Quote Master Grid - Click Run Report to see final report selection"
        Me.tgQh.CaptionHeight = 19
        Me.tgQh.DataSource = Me.QUTLU1BindingSource
        Me.tgQh.GroupByAreaVisible = False
        Me.tgQh.GroupByCaption = "Drag a column header here to group by that column"
        Me.tgQh.Images.Add(CType(resources.GetObject("tgQh.Images"), System.Drawing.Image))
        Me.tgQh.Location = New System.Drawing.Point(7, 6)
        Me.tgQh.Margin = New System.Windows.Forms.Padding(30)
        Me.tgQh.MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.FloatingEditor
        Me.tgQh.Name = "tgQh"
        Me.tgQh.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tgQh.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tgQh.PreviewInfo.ZoomFactor = 75.0R
        Me.tgQh.PrintInfo.PageSettings = CType(resources.GetObject("tgQh.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tgQh.RowHeight = 17
        Me.tgQh.Size = New System.Drawing.Size(996, 60)
        Me.tgQh.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.GridNavigation
        Me.tgQh.TabIndex = 0
        Me.tgQh.Text = "C1TrueDBGrid1"
        Me.tgQh.UseCompatibleTextRendering = False
        Me.tgQh.Visible = False
        Me.tgQh.PropBag = resources.GetString("tgQh.PropBag")
        '
        'QUTLU1BindingSource
        '
        Me.QUTLU1BindingSource.DataMember = "QUTLU1"
        Me.QUTLU1BindingSource.DataSource = Me.DsSaw8
        '
        'C1StatusBar2
        '
        Me.C1StatusBar2.AutoSizeElement = C1.Framework.AutoSizeElement.Width
        Me.C1StatusBar2.LeftPaneItems.Add(Me.DocumentModifiedLabel)
        Me.C1StatusBar2.LeftPaneItems.Add(Me.RibbonSeparator9)
        Me.C1StatusBar2.LeftPaneItems.Add(Me.pbProgress)
        Me.C1StatusBar2.Location = New System.Drawing.Point(0, 840)
        Me.C1StatusBar2.Name = "C1StatusBar2"
        Me.C1StatusBar2.RightPaneItems.Add(Me.cmdPercent)
        Me.C1StatusBar2.RightPaneItems.Add(Me.trackbar)
        Me.C1StatusBar2.RightPaneWidth = 175
        Me.C1StatusBar2.Size = New System.Drawing.Size(1056, 22)
        '
        'DocumentModifiedLabel
        '
        Me.DocumentModifiedLabel.Enabled = False
        Me.DocumentModifiedLabel.Name = "DocumentModifiedLabel"
        Me.DocumentModifiedLabel.SmallImage = CType(resources.GetObject("DocumentModifiedLabel.SmallImage"), System.Drawing.Image)
        '
        'RibbonSeparator9
        '
        Me.RibbonSeparator9.Name = "RibbonSeparator9"
        '
        'pbProgress
        '
        Me.pbProgress.Name = "pbProgress"
        Me.pbProgress.Visible = False
        '
        'cmdPercent
        '
        Me.cmdPercent.Name = "cmdPercent"
        Me.cmdPercent.Text = "50%"
        '
        'trackbar
        '
        Me.trackbar.Name = "trackbar"
        Me.trackbar.Value = 50
        '
        'C1SuperTooltip1
        '
        Me.C1SuperTooltip1.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.C1SuperTooltip1.RightToLeft = System.Windows.Forms.RightToLeft.Inherit
        '
        'QuoteTableAdapter
        '
        Me.QuoteTableAdapter.ClearBeforeFill = True
        '
        'QuoteprojectcustBindingSource
        '
        Me.QuoteprojectcustBindingSource.DataMember = "projectcust"
        Me.QuoteprojectcustBindingSource.DataSource = Me.DsSaw8
        '
        'QuotelinesTableAdapter
        '
        Me.QuotelinesTableAdapter.ClearBeforeFill = True
        '
        'QutSlsSplitBindingSource
        '
        Me.QutSlsSplitBindingSource.DataMember = "qutslssplit"
        Me.QutSlsSplitBindingSource.DataSource = Me.DsSaw8
        '
        'QutslssplitTableAdapter
        '
        Me.QutslssplitTableAdapter.ClearBeforeFill = True
        '
        'QutnotesTableAdapter
        '
        Me.QutnotesTableAdapter.ClearBeforeFill = True
        '
        'ProjectcustTableAdapter
        '
        Me.ProjectcustTableAdapter.ClearBeforeFill = True
        '
        'QutLU1TableAdapter
        '
        Me.QutLU1TableAdapter.ClearBeforeFill = True
        '
        'QuoteRealLUTableAdapter1
        '
        Me.QuoteRealLUTableAdapter1.ClearBeforeFill = True
        '
        'QuoteRealNDULBindingSource
        '
        Me.QuoteRealNDULBindingSource.DataMember = "QuoteRealNDUL"
        Me.QuoteRealNDULBindingSource.DataSource = Me.DsSaw8
        '
        'QuoteRealNDULTableAdapter
        '
        Me.QuoteRealNDULTableAdapter.ClearBeforeFill = True
        '
        'DsSaw8BindingSource
        '
        Me.DsSaw8BindingSource.DataSource = Me.DsSaw8
        Me.DsSaw8BindingSource.Position = 0
        '
        'SpecRegFollowUpTableAdapter
        '
        Me.SpecRegFollowUpTableAdapter.ClearBeforeFill = True
        '
        'ChkSpecifiersCustInCols
        '
        Me.ChkSpecifiersCustInCols.AutoSize = True
        Me.ChkSpecifiersCustInCols.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkSpecifiersCustInCols.Location = New System.Drawing.Point(7, 202)
        Me.ChkSpecifiersCustInCols.Name = "ChkSpecifiersCustInCols"
        Me.ChkSpecifiersCustInCols.Size = New System.Drawing.Size(245, 18)
        Me.ChkSpecifiersCustInCols.TabIndex = 36
        Me.ChkSpecifiersCustInCols.Text = "Add Specifiers/Customers in Columns (Excel)"
        Me.C1SuperTooltip1.SetToolTip(Me.ChkSpecifiersCustInCols, "Add Specifiers (Arch, Eng, Etc) to Reports")
        Me.ChkSpecifiersCustInCols.UseVisualStyleBackColor = True
        '
        'frmQuoteRpt
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1056, 862)
        Me.Controls.Add(Me.C1StatusBar2)
        Me.Controls.Add(Me.tabQrt)
        Me.Controls.Add(Me.C1Ribbon1)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 42)
        Me.Name = "frmQuoteRpt"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Quote Reports"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.cboTypeofJob, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdReportQuote, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdReportTerrSpecCredit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdReportRealization, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdReportOtherTypes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdReportLineItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        CType(Me.SSPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboQuoteRptPrt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdSortPrimarySeq, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdSortSecondarySeq, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fraOutputOptions, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optOutputOptions, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCurrentPage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.C1Ribbon1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabQrt.ResumeLayout(False)
        Me.TabPage0.ResumeLayout(False)
        Me.fraDisplaySortSeq.ResumeLayout(False)
        Me.fraDisplaySortSeq.PerformLayout()
        Me.fraReport.ResumeLayout(False)
        Me.fraReportCmdSelection.ResumeLayout(False)
        CType(Me.cmdReportProjShortage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraSortSecondarySeq.ResumeLayout(False)
        Me.fraSortSecondarySeq.PerformLayout()
        CType(Me.cmdSecondarySeqCancel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdSecondarySeqContinue, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraSortPrimarySeq.ResumeLayout(False)
        Me.fraSortPrimarySeq.PerformLayout()
        CType(Me.cmdPrimarySeqCancel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdPrimarySeqContinue1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me._fdBranchCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraQuoteReports.ResumeLayout(False)
        Me.fraQuoteReports.PerformLayout()
        CType(Me.cboTypeCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbospeccross, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdok1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdCancel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdResetDefaults1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraFinishReports.ResumeLayout(False)
        Me.fraFinishReports.PerformLayout()
        Me.fraSelectDate.ResumeLayout(False)
        Me.fraSelectDate.PerformLayout()
        Me.fraQuoteLineReports.ResumeLayout(False)
        Me.fraQuoteLineReports.PerformLayout()
        CType(Me.cmdOK2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdCancel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdResetDefaults2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.fraUnitorExtended.ResumeLayout(False)
        Me.fraUnitorExtended.PerformLayout()
        Me.fraSalesorCost.ResumeLayout(False)
        Me.fraQtIncludeCommission.ResumeLayout(False)
        Me.fraQtIncludeCommission.PerformLayout()
        Me.gbxSortSeq.ResumeLayout(False)
        Me.gbxSortSeq.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.gbxSortSeqV.ResumeLayout(False)
        Me.gbxSortSeqV.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.tgln, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QuoteLinesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DsSaw8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tgSpecReg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SpecRegFollowUpBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tglnDIST, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tgrDIST, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QuoteRealLUBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tgQhDIST, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QutNotesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tgr, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tgQh, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QUTLU1BindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.C1StatusBar2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QuoteprojectcustBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QutSlsSplitBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.QuoteRealNDULBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DsSaw8BindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents DsSaw8 As VQRT.dsSaw8
    Friend WithEvents QuoteTableAdapter As VQRT.dsSaw8TableAdapters.quoteTableAdapter
    Friend WithEvents QuoteprojectcustBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents QuoteLinesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents QutSlsSplitBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents QutNotesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents QUTLU1BindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents ExitMainRibbon As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents ExitRibbon As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonListItem1 As C1.Win.C1Ribbon.RibbonListItem
    Friend WithEvents RibbonLabel2 As C1.Win.C1Ribbon.RibbonLabel
    Friend WithEvents CloseButton As C1.Win.C1Ribbon.RibbonSplitButton
    Friend WithEvents PrintDocumentButton As C1.Win.C1Ribbon.RibbonSplitButton
    Friend WithEvents SaveAs1DocumentButton As C1.Win.C1Ribbon.RibbonSplitButton
    Friend WithEvents SaveDocumentButton As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents LookupRibbon As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents OpenDocumentButton As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents NewDocumentButton As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonButton32 As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonButton23 As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonButton13 As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonButton3 As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents PrintPreviewRibbon As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents PrintQuickRibbon As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents PrintRibbon As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents C1Ribbon1 As C1.Win.C1Ribbon.C1Ribbon
    Friend WithEvents RibbonApplicationMenu1 As C1.Win.C1Ribbon.RibbonApplicationMenu
    Friend WithEvents rbnExitMainMenu As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonButton2 As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonListItem2 As C1.Win.C1Ribbon.RibbonListItem
    Friend WithEvents RibbonLabel1 As C1.Win.C1Ribbon.RibbonLabel
    Friend WithEvents RibbonConfigToolBar1 As C1.Win.C1Ribbon.RibbonConfigToolBar
    Friend WithEvents RibbonStyleMenu As C1.Win.C1Ribbon.RibbonMenu
    Friend WithEvents RibbonToggleGroup1 As C1.Win.C1Ribbon.RibbonToggleGroup
    Friend WithEvents Office2007BlueStyleButton As C1.Win.C1Ribbon.RibbonToggleButton
    Friend WithEvents Office2007SilverStyleButton As C1.Win.C1Ribbon.RibbonToggleButton
    Friend WithEvents Office2007BlackStyleButton As C1.Win.C1Ribbon.RibbonToggleButton
    Friend WithEvents F1HelpButton As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonQat1 As C1.Win.C1Ribbon.RibbonQat
    Friend WithEvents RbnBtnGridViewInverted As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents RbnBtnGridViewNormal As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnGridViewGroupBy As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnGridViewSplit As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnGridViewExpandGp As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnGridViewCollapse As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnHelp As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents rtCustomize As C1.Win.C1Ribbon.RibbonTab
    Friend WithEvents rgColors As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents RibbonColorPicker1 As C1.Win.C1Ribbon.RibbonColorPicker
    Friend WithEvents RibbonColorPicker2 As C1.Win.C1Ribbon.RibbonColorPicker
    Friend WithEvents rgThemes As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents RibbonGallery1 As C1.Win.C1Ribbon.RibbonGallery
    Friend WithEvents RibbonGalleryItem1 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem2 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem3 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem4 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem5 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem6 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem7 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem8 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem9 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem10 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem11 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem12 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem13 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem14 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem15 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem16 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents RibbonGalleryItem17 As C1.Win.C1Ribbon.RibbonGalleryItem
    Friend WithEvents rgFont As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents RibbonToolBar1 As C1.Win.C1Ribbon.RibbonToolBar
    Friend WithEvents RibbonFontComboBox2 As C1.Win.C1Ribbon.RibbonFontComboBox
    Friend WithEvents FontSizeComboBox As C1.Win.C1Ribbon.RibbonComboBox
    Friend WithEvents size8Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size9Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size10Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size11Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size12Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size14Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size16Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size18Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size20Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size22Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size24Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size26Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size28Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size36Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size48Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents size72Button As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonToolBar2 As C1.Win.C1Ribbon.RibbonToolBar
    Friend WithEvents FontBoldButton As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents FontItalicButton As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents FontUnderlineButton As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents FontStrikethroughButton As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonSeparator7 As C1.Win.C1Ribbon.RibbonSeparator
    Friend WithEvents FontColorPicker As C1.Win.C1Ribbon.RibbonColorPicker
    Friend WithEvents RibbonSeparator8 As C1.Win.C1Ribbon.RibbonSeparator
    Friend WithEvents AutoFit As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonTab1 As C1.Win.C1Ribbon.RibbonTab
    Friend WithEvents RibbonTab2 As C1.Win.C1Ribbon.RibbonTab
    Friend WithEvents RibbonGroup3 As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents RbnBtnExportExcel As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnExportCSVTab As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnExportCSVComma As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnExportHTML As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnExportPDF As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnExportRTF As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnExportOptions As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnBtnExportPrint As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonTab3 As C1.Win.C1Ribbon.RibbonTab
    Friend WithEvents RbnResetToCurrentGridLayoutToolStripMenuItem1 As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents RbnSaveCurrentGridLayoutSettingsToolStripMenuItem As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnResetToOriginalGridLayoutToolStripMenuItem As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnResetToCurrentGridLayoutToolStripMenuItem As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents tabQrt As System.Windows.Forms.TabControl
    Friend WithEvents TabPage0 As System.Windows.Forms.TabPage
    Public WithEvents fraDisplaySortSeq As System.Windows.Forms.GroupBox
    Public WithEvents txtPrimarySortSeq As System.Windows.Forms.TextBox
    Public WithEvents txtSecondarySort As System.Windows.Forms.TextBox
    Public WithEvents pnlPrimarySortSeq As System.Windows.Forms.Label
    Public WithEvents pnlTypeOfRpt As System.Windows.Forms.Label
    Public WithEvents pnlSecondarySort As System.Windows.Forms.Label
    Public WithEvents fraReport As System.Windows.Forms.GroupBox
    Public WithEvents fraReportCmdSelection As System.Windows.Forms.GroupBox
    Public WithEvents fraSortSecondarySeq As System.Windows.Forms.GroupBox
    Public WithEvents SSPanel4 As System.Windows.Forms.Label
    Public WithEvents fraSortPrimarySeq As System.Windows.Forms.GroupBox
    Public WithEvents SSPanel3 As System.Windows.Forms.Label
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents fraSelectDate As System.Windows.Forms.GroupBox
    Public WithEvents chkBlankBidDates As System.Windows.Forms.CheckBox
    Public WithEvents lblEndBid As System.Windows.Forms.Label
    Public WithEvents lblStartBid As System.Windows.Forms.Label
    Public WithEvents lblEndEntry As System.Windows.Forms.Label
    Public WithEvents txtStartEntry As System.Windows.Forms.TextBox
    Public WithEvents lblStartEntry As System.Windows.Forms.Label
    Public WithEvents txtEndEntry As System.Windows.Forms.TextBox
    Public WithEvents txtEndBid As System.Windows.Forms.TextBox
    Public WithEvents txtStartBid As System.Windows.Forms.TextBox
    Public WithEvents txtSpecifierCode As System.Windows.Forms.TextBox
    Public WithEvents txtSalesman As System.Windows.Forms.TextBox
    Public WithEvents lblStartQuote As System.Windows.Forms.Label
    Public WithEvents txtRetrieval As System.Windows.Forms.TextBox
    Public WithEvents pnlQutRealCode As System.Windows.Forms.Label
    Public WithEvents txtStatus As System.Windows.Forms.TextBox
    Public WithEvents pnlSpecCross As System.Windows.Forms.Label
    Public WithEvents txtEndQuoteAmt As System.Windows.Forms.TextBox
    Public WithEvents pnlLotUnit As System.Windows.Forms.Label
    Public WithEvents txtStartQuoteAmt As System.Windows.Forms.TextBox
    Public WithEvents pnlCity As System.Windows.Forms.Label
    Public WithEvents txtJobNameSS As System.Windows.Forms.TextBox
    Public WithEvents pnlState As System.Windows.Forms.Label
    Public WithEvents txtQutRealCode As System.Windows.Forms.TextBox
    Public WithEvents pnlMktSeg As System.Windows.Forms.Label
    Public WithEvents PnlLastChgBy As System.Windows.Forms.Label
    Public WithEvents cboLotUnit As System.Windows.Forms.ComboBox
    Public WithEvents pnlSlsSplits As System.Windows.Forms.Label
    Public WithEvents cboStockJob As System.Windows.Forms.ComboBox
    Public WithEvents pnlCSR As System.Windows.Forms.Label
    Public WithEvents pnlSltCode As System.Windows.Forms.Label
    Public WithEvents txtCSR As System.Windows.Forms.TextBox
    Public WithEvents pnlStkJob As System.Windows.Forms.Label
    Public WithEvents txtLastChgBy As System.Windows.Forms.TextBox
    Public WithEvents pnlQuoteToSls As System.Windows.Forms.Label
    Public WithEvents txtSlsSplit As System.Windows.Forms.TextBox
    Public WithEvents pnlCSRdist As System.Windows.Forms.Label
    Public WithEvents txtState As System.Windows.Forms.TextBox
    Public WithEvents lblJobName As System.Windows.Forms.Label
    Public WithEvents txtCity As System.Windows.Forms.TextBox
    Public WithEvents lblEndQuote As System.Windows.Forms.Label
    Public WithEvents txtMktSegment As System.Windows.Forms.TextBox
    Public WithEvents lblSalesman As System.Windows.Forms.Label
    Public WithEvents fraFinishReports As System.Windows.Forms.GroupBox
    Public WithEvents chkBranchReport As System.Windows.Forms.CheckBox
    Public WithEvents chkSalesmanPerPage As System.Windows.Forms.CheckBox
    Public WithEvents chkDetailTotal As System.Windows.Forms.CheckBox
    Public WithEvents chkIncludeCommDolPer As System.Windows.Forms.CheckBox
    Public WithEvents chkExcludeDuplicates As System.Windows.Forms.CheckBox
    Public WithEvents chkSlsFromHeader As System.Windows.Forms.CheckBox
    Public WithEvents chkExportAllExcel As System.Windows.Forms.CheckBox
    Public WithEvents lblStatus As System.Windows.Forms.Label
    Public WithEvents txtQuoteToSls As System.Windows.Forms.TextBox
    Public WithEvents lblRetrieval As System.Windows.Forms.Label
    Public WithEvents txtCSRofCust As System.Windows.Forms.TextBox
    Public WithEvents pnlSpecifierCode As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents tgQh As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Public WithEvents chkMfgBreakdown As System.Windows.Forms.CheckBox
    Public WithEvents chkCustomerBreakdown As System.Windows.Forms.CheckBox
    Friend WithEvents CmdShowColstoPrt As System.Windows.Forms.Button
    Friend WithEvents CmdRunReport As System.Windows.Forms.Button
    Friend WithEvents chkNotes As System.Windows.Forms.CheckBox
    Friend WithEvents ChkSpecifiers As System.Windows.Forms.CheckBox
    Friend WithEvents gbxSortSeq As System.Windows.Forms.GroupBox
    Friend WithEvents txtSortSeq As System.Windows.Forms.TextBox
    Friend WithEvents gbxSortSeqV As System.Windows.Forms.GroupBox
    Friend WithEvents txtSortSeqV As System.Windows.Forms.TextBox
    Friend WithEvents QuoteRealLUTableAdapter1 As VQRT.dsSaw8TableAdapters.QuoteRealLUTableAdapter
    Friend WithEvents QuoteRealLUBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents tgr As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents txtSortSeqCriteria As System.Windows.Forms.TextBox
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Friend WithEvents rbnAddFilterBar As C1.Win.C1Ribbon.RibbonButton
#End Region

    Private Sub tabQrt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabQrt.Click
        On Error Resume Next

        Select Case Me.tabQrt.SelectedIndex '08-20-08
            Case 1 'Selection '03-22-13 added , MsgBoxStyle.OkCancel
                If Me.pnlTypeOfRpt.Text.Trim = "" Then
                    Resp = MsgBox("Go to the Left Tab to start Reports Process. Select a Type of Report First", MsgBoxStyle.OkCancel)
                    If Resp = vbCancel Then Exit Sub '03-22-13
                    Me.tabQrt.SelectedIndex = 0 : Exit Sub '10-12-10 
                End If
                '09-04-10 If tgQh.RowCount < 1 Then ' Is Nothing Then 'If tg Then
                If Me.pnlTypeOfRpt.Text.StartsWith("Product Sales History") Then
                    If tgln.RowCount < 1 Then GoTo AskMsg
                ElseIf Me.pnlTypeOfRpt.Text.StartsWith("Realization") Then
                    If tgr.RowCount < 1 Then GoTo AskMsg
                Else
                    If tgQh.RowCount < 1 Then GoTo AskMsg
                End If
                Exit Sub
AskMsg:         '09-04-10 
                Resp = MsgBox("Go to the Left Tab Sort Sequence to start Reports Process", MsgBoxStyle.OkCancel)
                If Resp = vbOK Then
                    '09-04-10  Me.tabQrt.SelectedIndex = 0 '06-08-09
                Else
                    Exit Sub '06-30-09
                End If
                Me.tabQrt.SelectedIndex = 0 '06-08-09
                '06-06-09 Show hide
                'Call frmShowHideGrid.ShowHideGridColBidB("Show", UserDir & "ShowHidePrintBidB.xml") '06-06-09
                'Me.tg.DataView = C1.Win.C1TrueDBGrid.DataViewEnum.GroupBy
                'Me.tg.FilterBar = True
                'Me.TabPage2.Text = 
        End Select
    End Sub
    Public WithEvents txtSelectCode As System.Windows.Forms.TextBox

    Private Sub txtSelectCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSelectCode.Enter
        On Error Resume Next
        Me.txtSelectCode.SelectionStart = 0
        Me.txtSelectCode.SelectionLength = Len(Me.txtSelectCode.Text)
        If txtSelectCode.Text = "" Then txtSelectCode.Text = "ALL" '09-25-07 JH
    End Sub
    Friend WithEvents C1StatusBar2 As C1.Win.C1Ribbon.C1StatusBar
    Friend WithEvents DocumentModifiedLabel As C1.Win.C1Ribbon.RibbonLabel
    Friend WithEvents RibbonSeparator9 As C1.Win.C1Ribbon.RibbonSeparator
    Friend WithEvents pbProgress As C1.Win.C1Ribbon.RibbonProgressBar
    Friend WithEvents cmdPercent As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents trackbar As C1.Win.C1Ribbon.RibbonTrackBar
    Friend WithEvents RbnSaveCurrentQuoteToGridLayoutSettingsToolStripMenuItem1 As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnResetCurrentQuoteToGridLayoutSettingsToolStripMenuItem1 As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RbnResetToOriginalQuoteToGridLayoutToolStripMenuItem1 As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents fraQuoteReports As System.Windows.Forms.GroupBox
    Friend WithEvents fraQuoteLineReports As System.Windows.Forms.GroupBox
    Public WithEvents ChkTotalsOnly As System.Windows.Forms.CheckBox
    Public WithEvents fraUnitorExtended As System.Windows.Forms.GroupBox
    Public WithEvents optUnitOrExtended_Unit As System.Windows.Forms.RadioButton
    Public WithEvents optUnitOrExtended_Extd As System.Windows.Forms.RadioButton
    Public WithEvents fraSalesorCost As System.Windows.Forms.GroupBox
    Public WithEvents optSalesorCost_Cost As System.Windows.Forms.RadioButton
    Public WithEvents optoptSalesorCost_Sales As System.Windows.Forms.RadioButton
    Public WithEvents fraQtIncludeCommission As System.Windows.Forms.GroupBox
    Public WithEvents chkIncludeCommDoll As System.Windows.Forms.CheckBox
    Public WithEvents chkIncludeCommPer As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents tgln As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Public WithEvents txtStat As System.Windows.Forms.TextBox
    Public WithEvents PnlStatus As System.Windows.Forms.Label
    Public WithEvents txtSlsTerr As System.Windows.Forms.TextBox
    Public WithEvents PnlSls As System.Windows.Forms.Label
    Public WithEvents txtLastChgByLine As System.Windows.Forms.TextBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents txtRetr As System.Windows.Forms.TextBox
    Public WithEvents PnlRet As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents TxtSingleCatNum As System.Windows.Forms.TextBox
    Public WithEvents TxtSearchString As System.Windows.Forms.TextBox
    Public WithEvents txtMfgLine As System.Windows.Forms.TextBox
    Public WithEvents PnlCatNum As System.Windows.Forms.Label
    Public WithEvents PnlCatSrch As System.Windows.Forms.Label
    Public WithEvents PnlMfg As System.Windows.Forms.Label
    Public WithEvents txtPrcCode As System.Windows.Forms.TextBox
    Public WithEvents pnlPrcCode As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents txtSpecCross As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ProgramDateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents ChkExtendByProb As System.Windows.Forms.CheckBox
    Public WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboTypeofJob As C1.Win.C1List.C1Combo
    Friend WithEvents chkPrtPlanLines As System.Windows.Forms.CheckBox
    Public WithEvents cboLinesInclude As System.Windows.Forms.ComboBox
    Friend WithEvents cmdReportRealization As C1.Win.C1Input.C1Button
    Friend WithEvents cmdReportQuote As C1.Win.C1Input.C1Button
    Friend WithEvents cmdReportProjShortage As C1.Win.C1Input.C1Button
    Friend WithEvents cmdReportTerrSpecCredit As C1.Win.C1Input.C1Button
    Friend WithEvents cmdReportOtherTypes As C1.Win.C1Input.C1Button
    Friend WithEvents cmdReportLineItems As C1.Win.C1Input.C1Button
    Friend WithEvents cmdSecondarySeqContinue As C1.Win.C1Input.C1Button
    Friend WithEvents cmdPrimarySeqContinue1 As C1.Win.C1Input.C1Button
    Friend WithEvents cmdSecondarySeqCancel As C1.Win.C1Input.C1Button
    Friend WithEvents cmdPrimarySeqCancel1 As C1.Win.C1Input.C1Button
    Friend WithEvents cmdResetDefaults1 As C1.Win.C1Input.C1Button
    Friend WithEvents cmdCancel1 As C1.Win.C1Input.C1Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Public WithEvents chkBlankLine As System.Windows.Forms.CheckBox
    Friend WithEvents cmdok1 As C1.Win.C1Input.C1Button
    Friend WithEvents cmdOK2 As C1.Win.C1Input.C1Button
    Friend WithEvents cmdCancel2 As C1.Win.C1Input.C1Button
    Friend WithEvents cmdResetDefaults2 As C1.Win.C1Input.C1Button
    Friend WithEvents QuotelinesTableAdapter As VQRT.dsSaw8TableAdapters.quotelinesTableAdapter
    Friend WithEvents QutslssplitTableAdapter As VQRT.dsSaw8TableAdapters.qutslssplitTableAdapter
    Friend WithEvents QutnotesTableAdapter As VQRT.dsSaw8TableAdapters.qutnotesTableAdapter
    Friend WithEvents ProjectcustTableAdapter As VQRT.dsSaw8TableAdapters.projectcustTableAdapter
    Friend WithEvents QutLU1TableAdapter As VQRT.dsSaw8TableAdapters.QUTLU1TableAdapter
    Public WithEvents chkHaveMFGCode As System.Windows.Forms.CheckBox
    Public WithEvents chkPrtNTElines As System.Windows.Forms.CheckBox
    Friend WithEvents cboSortRealization As System.Windows.Forms.CheckedListBox
    Public WithEvents ChkCheckBidDates As System.Windows.Forms.CheckBox
    Friend WithEvents RibbonGroup2 As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents RbnTgToExcel As C1.Win.C1Ribbon.RibbonButton
    Public WithEvents chkBidJobsOnly As System.Windows.Forms.CheckBox
    Public WithEvents chkShowLatestCust As System.Windows.Forms.CheckBox
    Public WithEvents ChkQuoteNoSpecifiers As System.Windows.Forms.CheckBox
    Friend WithEvents RibbonTab4 As C1.Win.C1Ribbon.RibbonTab
    Friend WithEvents rbnMaxNameLength As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents rbnMaxNameTxt As C1.Win.C1Ribbon.RibbonTextBox
    Friend WithEvents rbnMaxJobTxt As C1.Win.C1Ribbon.RibbonTextBox
    Friend WithEvents RibbonTab5 As C1.Win.C1Ribbon.RibbonTab
    Friend WithEvents rbnWholeDollars As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents chkWholeDollars As C1.Win.C1Ribbon.RibbonCheckBox
    Friend WithEvents chkAddCommas As C1.Win.C1Ribbon.RibbonCheckBox
    Friend WithEvents chkAddDollarSign As C1.Win.C1Ribbon.RibbonCheckBox
    Friend WithEvents RibbonTab6 As C1.Win.C1Ribbon.RibbonTab
    Friend WithEvents rbnPrintColor As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents chkPrintGrayScale As C1.Win.C1Ribbon.RibbonCheckBox
    Friend WithEvents cbospeccross As C1.Win.C1List.C1Combo
    Friend WithEvents cmdBackViewGrid As System.Windows.Forms.Button
    Friend WithEvents chkBrandReport As System.Windows.Forms.CheckBox
    Public WithEvents chkShowCustomers As System.Windows.Forms.CheckBox
    Public WithEvents chkUseSpecifierCode As System.Windows.Forms.CheckBox
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents txtCustomerCodeLine As System.Windows.Forms.TextBox
    Friend WithEvents _fdBranchCode As C1.Win.C1List.C1Combo
    Public WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents mnuBrandReport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuBrandMfgChg As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuJump As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSupport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cboTypeCustomer As C1.Win.C1List.C1Combo
    Public WithEvents lblTypeCustomer As System.Windows.Forms.Label
    Public WithEvents txtSortSecondarySeq As System.Windows.Forms.TextBox
    Public WithEvents cboSortSecondarySeq As System.Windows.Forms.ListBox
    Friend WithEvents tgQhDIST As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents tgrDIST As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents QuoteRealNDULBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents QuoteRealNDULTableAdapter As VQRT.dsSaw8TableAdapters.QuoteRealNDULTableAdapter
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tglnDIST As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Public WithEvents txtPrimarySort As System.Windows.Forms.TextBox
    Public WithEvents cboSortPrimarySeq As System.Windows.Forms.ListBox
    Friend WithEvents chkIncludeSLSSPlit As System.Windows.Forms.CheckBox
    Friend WithEvents chkIncludeSpecifiers As System.Windows.Forms.CheckBox
    Friend WithEvents DsSaw8BindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents chkIncludeNotesLineItems As System.Windows.Forms.CheckBox
    Friend WithEvents SpecRegFollowUpBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents SpecRegFollowUpTableAdapter As VQRT.dsSaw8TableAdapters.SpecRegFollowUpTableAdapter
    Friend WithEvents tgSpecReg As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Friend WithEvents RibbonGroup1 As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents rbnDeleteFiles As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents mnuBrandListLoad As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuBrandExclude As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DTPicker1EndBid As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTPicker1EndEntry As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTPicker1StartBid As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTPickerStartEntry As System.Windows.Forms.DateTimePicker
    Friend WithEvents RibbonTab7 As C1.Win.C1Ribbon.RibbonTab
    Friend WithEvents RibbonGroup26 As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents rbnJoinGoToMeeting As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonGroup18 As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents rbnHelpAbout As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents rbnHelpAboutDirectory As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents RibbonGroup5 As C1.Win.C1Ribbon.RibbonGroup
    Friend WithEvents rbnHelpMaster As C1.Win.C1Ribbon.RibbonButton
    Friend WithEvents C1SuperTooltip1 As C1.Win.C1SuperTooltip.C1SuperTooltip
    Friend WithEvents RibbonBottomToolBar1 As C1.Win.C1Ribbon.RibbonBottomToolBar
    Friend WithEvents RibbonTopToolBar1 As C1.Win.C1Ribbon.RibbonTopToolBar
    Friend WithEvents ChkSpecifiersCustInCols As CheckBox
End Class