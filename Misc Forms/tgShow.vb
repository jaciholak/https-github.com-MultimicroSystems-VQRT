Imports System.Windows.Forms
Imports C1.C1Preview
Imports C1.C1Preview.DataBinding '09-20-08
Imports C1.Win.C1Preview '06-18-08 
Imports VB = Microsoft.VisualBasic

Public Class frmShowHideGrid
    'Public dsGrid As DataSet = New DataSet("dsGrid") '12-07-08
    'Public table As DataTable = dsGrid.Tables.Add("Items") '12-12'08
    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
        ' Create a new instance of the child form.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Window " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs)
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: Add code here to open the file.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: Add code here to save the current contents of the form to a file.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    End Sub

    Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Me.ToolStrip.Visible = Me.ToolBarToolStripMenuItem.Checked
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Me.StatusStrip.Visible = Me.StatusBarToolStripMenuItem.Checked
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub btnShowSaveExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowSaveExit.Click
        '"Show Hide Quote Hdr Printing Columns" = "VQrtShowHideRepPrtHdr.xml"   tgQh
        '"Show Hide Quote Line Items" = "VQrtLineItemsShowHide.xml"  tg
        '"Show Hide Realization Columns" = "VQrtRealQTOShowHidePrint.xml"   tgr  
        '09-02-09 
        If Me.Text = "Show Hide Quote Hdr Printing Columns" Then '= "VQrtShowHideRepPrtHdr.xml" 
            Dim tmpName As String = ""
            If DIST Then '01-19-10
                tmpName = "VQrtShowHideDistPrtHdr.xml" '
            Else
                tmpName = "VQrtShowHideRepPrtHdr.xml" '
            End If
            Call ShowHideGridCol("Save", tmpName, frmQuoteRpt.tgQh)

        ElseIf Me.Text = "Show Hide Quote Line Items" Then
            'text = "Show Hide Quote Hdr Printing Columns"

            If DIST Then 'tmpName = "VQrtLineItemsDistShowHide.xml" else tmpName = "VQrtLineItemsRepShowHide.xml" '05-05-10
                Call Me.ShowHideGridCol("Save", "VQrtLineItemsDistShowHide.xml", frmQuoteRpt.tgln)
            Else '05-05-10
                Call Me.ShowHideGridCol("Save", "VQrtLineItemsRepShowHide.xml", frmQuoteRpt.tgln)
            End If
        ElseIf Me.Text = "Show Hide Realization Columns" Then
            Dim ShowAllQuoteHeader As String = "" '06-06-11 "Show All Quote Header Fields"
            If frmQuoteRpt.chkCustomerBreakdown.CheckState = CheckState.Checked Then ShowAllQuoteHeader = "ShowAll" '06-06-11 = "Show All Quote Header Fields" Then '06-06-11 "Add Cust QuoteTo Breakdown to Report"
            If DIST Then '05-07-10 VQrtRealQTOShowHideDistPrint.xml else VQrtRealQTOShowHideRepPrint.xml
                Call Me.ShowHideGridCol("Save", "VQrtRealQTOShowHideDistPrint" & ShowAllQuoteHeader & ".xml", frmQuoteRpt.tgr)
            Else
                Call Me.ShowHideGridCol("Save", "VQrtRealQTOShowHideRepPrint" & ShowAllQuoteHeader & ".xml", frmQuoteRpt.tgr)
            End If
        End If

        SplitContainer2.Panel2Collapsed = True
        Me.Close()

    End Sub

    '12-04-10 Dup CodePublic Sub SetupPrintPreview(ByVal doc As C1PrintDocument, ByVal FirmName As String)
    '    Try
    '' make the document:
    '        doc.Clear()
    '        RT = New C1.C1Preview.RenderTable ' RT  Table is Public
    '' for tables, "auto" width means that the width of the table
    ''        ' will be equal to the widths of all columns, so we MUST also
    ''        ' set the columns' widths.
    ''        rt.Width = "auto"
    '        RT.Width = "auto" ' Wide 
    '        RT.CellStyle.Padding.All = ".5mm"
    '        RT.Style.GridLines.All = LineDef.Default

    '        ppv.Doc.PageLayout.PageSettings.Landscape = True
    ''Not Used NowRTotals = New C1.C1Preview.RenderText
    '' define PageLayout for the first page
    'Dim pl As New PageLayout()
    '        pl.PageSettings = New C1PageSettings()
    '        pl.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.Letter
    '        pl.PageSettings.Landscape = True
    '        pl.PageSettings.LeftMargin = ".4in" '".25cm"
    '        pl.PageSettings.RightMargin = ".4in" '".25cm"
    '        pl.PageSettings.TopMargin = ".5in"
    '        pl.PageSettings.BottomMargin = ".5in"
    ''FrmMenu.C1PrintDocument1 = doc '09-18-08
    ''C1.C1Preview.C1PrintDocument.
    ''FrmMenu.C1PrintDocument1.PageLayouts.FirstPage = pl
    ''FrmMenu.C1PrintDocument1.PageLayout = pl
    ''FrmMenu.C1PrintDocument1.PageLayout.PageHeader = New RenderText("header")
    ''FrmMenu.C1PrintDocument1.PageLayout.PageHeader.Style.TextAlignHorz = AlignHorzEnum.Right
    ''FrmMenu.C1PrintDocument1.PageLayout.PageHeader.Style.Spacing.Bottom = ".2in" '".5cm"
    ''FrmMenu.C1PrintDocument1.PageLayout.PageHeader.Style.Borders.Bottom = LineDef.Default
    ''09-18-08
    '        doc.PageLayouts.FirstPage = pl
    '        doc.PageLayout = pl
    '        doc.PageLayout.PageHeader = New RenderText("header")
    '        doc.PageLayout.PageHeader.Style.TextAlignHorz = AlignHorzEnum.Right
    '        doc.PageLayout.PageHeader.Style.Spacing.Bottom = ".3in" '".5cm"
    '        doc.PageLayout.PageHeader.Style.Borders.Bottom = LineDef.Default

    'Dim RTa As New C1.C1Preview.RenderTable
    '        RTa.Cells(0, 0).Text = "Task Follow Up Report    Report Date = " & VB6.Format(Now, "Short Date") & Space(8) & FirmName '06-23-08
    '        RTa.Cells(0, 0).Style.TextAlignHorz = AlignHorzEnum.Left
    '        RTa.Cells(0, 1).Text = "Page [PageNo] of [PageCount]     *" '  & "     " '06-18-08 Spacing
    '        RTa.Cells(0, 1).Style.TextAlignHorz = AlignHorzEnum.Right
    ''FrmMenu.C1PrintDocument1.PageLayout.PageHeader = RTa '06-20-08
    ''FrmMenu.C1PrintDocument1.PageLayout.PageFooter = New RenderText("Page [PageNo] of [PageCount]     *") '  & "      ") ' & "      " '06-18-08 Spacing
    ''FrmMenu.C1PrintDocument1.PageLayout.PageFooter.Style.TextAlignHorz = AlignHorzEnum.Right
    ''FrmMenu.C1PrintDocument1.PageLayout.PageFooter.Style.Spacing.Top = ".2in" '"0.5cm"
    ''FrmMenu.C1PrintDocument1.PageLayout.PageFooter.Style.Borders.Top = LineDef.Default
    ''FrmMenu.C1PrintDocument1.PageLayout.PageSettings.Landscape = True
    ''09-18-08
    '        doc.PageLayout.PageHeader = RTa '06-20-08
    '        doc.PageLayout.PageFooter = New RenderText("Page [PageNo] of [PageCount]     *") '  & "      ") ' & "      " '06-18-08 Spacing
    '        doc.PageLayout.PageFooter.Style.TextAlignHorz = AlignHorzEnum.Right
    '        doc.PageLayout.PageFooter.Style.Spacing.Top = ".2in" '"0.5cm"
    '        doc.PageLayout.PageFooter.Style.Borders.Top = LineDef.Default
    '        doc.PageLayout.PageSettings.Landscape = True
    ''FrmMenu.C1PrintDocument1.Body.Children.Add(New C1.C1Preview.RenderC1Printable(doc)) '09-18-08

    'Dim widths(RTa.Cols.Count) As Double
    'Dim row, col As Integer
    '        For row = 0 To RTa.Rows.Count - 1
    '            For col = 0 To RTa.Cols.Count - 1
    '                If RTa.Cells(row, col).RenderObject IsNot Nothing Then
    'Dim s As SizeD = RTa.Cells(row, col).RenderObject.CalcSize(Unit.Auto, Unit.Auto)
    '                    widths(col) = Math.Max(widths(col), s.Width)
    '                End If
    '            Next col
    '        Next row
    '        RTa.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-15-08
    ''RTa.SplitHorzBehavior = True '12-12-08 
    ''Ng doc.Body.Children.Add(RTa) '12-15-08
    '        doc.PageLayout.PageFooter.Style.Spacing.Top = "0.5cm"
    '        doc.PageLayout.PageFooter.Style.Borders.Top = LineDef.Default
    '        doc.Style.Font = New Font("Arial", 8)
    '' create the title of the document
    ''Dim title As New RenderParagraph()
    ''title.Content.AddText("The new version of C1PrintDocumet provides the ")
    ''title.Content.AddText("PageLayouts", Color.Blue)
    ''title.Content.AddText(" property allowing to define separate page layouts for the first page, even pages, and odd pages.")
    ''title.Style.TextAlignHorz = AlignHorzEnum.Justify
    ''title.Style.Borders.Bottom = New LineDef("1mm", Color.Black)
    ''doc.Body.Children.Add(title)
    '    Catch myException As Exception
    '        MsgBox(myException.Message & vbCrLf & "SetupPrintPreview" & vbCrLf)
    '        If DebugOn ThenStop'CatchStop
    '    End Try
    'End Sub



    Private Sub frmShowHideGrid_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.Visible = True
    End Sub
    Public Sub ShowDefaultFiles()
        cboSavePrintOption.Visible = True
        Dim ShowHideFile As String = UserDir
        cboSavePrintOption.Items.Clear()
        'VPCDprintSave
        Dim Response As Integer
        Dim Diskfile As String

        Diskfile = Dir$(UserDir, 4)
        If Len(Diskfile) = 0 Then
            Response = MsgBox("No Files Currently on Diskette", vbOKOnly, "")
        End If

        Do While Diskfile <> ""  'List files from disk on form
            If (VB.Left(Diskfile, 13)) = "VPCDprintSave" Then
                Diskfile = Replace(Diskfile, ".xml", "")                'Debug.Print(Mid(Diskfile, 14, Diskfile.Length))
                Diskfile = Replace(Diskfile, ".XML", "")
                cboSavePrintOption.Items.Add(Mid(Diskfile, 14, Diskfile.Length))
            End If
            Diskfile = Dir$()
        Loop

        cboSavePrintOption.Items.Add("Create New...")
    End Sub
    Private Sub btnShowAllOn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowAllOn.Click, cmdShowAllOff.Click
        'Call ShowHideGridCol("AllOn", "ShowHidePrintXOVER.xml") '12-07-08 ByVal ShowHide As String)
        Dim I As Integer = 0
        Dim dt As DataTable = tgShow.DataSource : Dim Rows As Integer = dt.Rows.Count - 1
        If CType(sender, Windows.Forms.Button).Text = "All On" Then
            For I = 0 To Rows : tgShow(I, 1) = True : Next
        Else
            For I = 0 To Rows : tgShow(I, 1) = False : Next
        End If
    End Sub
    Public Sub ShowHideGridCol(ByVal ShowHide As String, ByVal ShowHideFile As String, ByVal tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid) '12-05-08   TrueGridCreateInCode.doc   12-05-08
        'On Error Resume Next
        'call frmShowHideGrid.ShowHideGridCol("Load", "ShowHideGridVPCDLI.xml", tgLines)
        'Show / Save / AllOn 'Call ShowHideGridCol("Show") / "Save"'12-05-08 ByVal ShowHide As String)
        'ShowHideFile = "C:\SAW8\ShowHideGridFoll.xml" or ShowHideFile = "C:\SAW8\ShowHidePrintFoll.xml"
        'Each Program needs to chg Foll in ShowHideGridFoll.xml & ShowHidePrintFoll.xml
        'Public ShowHideFileName As String '12-12-08 Copy All Three to Module
        'Public dsGrid As DataSet = New DataSet("dsGrid") '12-12-08   'Public table As DataTable = dsGrid.Tables.Add("Items") '12-12'08
        '"Show Hide Quote Hdr Printing Columns" = "VQrtShowHidePrtHdr.xml"   tgQh
        '"Show Hide Quote Line Items" = "VQrtLineItemsDistShowHide.xml"  tg  or Rep
        '"Show Hide Realization Columns" = "VQrtRealQTOShowHidePrint.xml"   tgr 
        Dim ColName As String = "" '12-31-14 JTC added
        Try
            Dim ShowHideFileSave As String = ShowHideFile '02-06-11
            ShowHideFile = UserDir & ShowHideFile
            If ShowHide = "Save" Or ShowHide = "AllOn" Then GoTo 130 '12-07-08
            If ShowHide = "Load" Then  Else Me.Show()
            tgShow.Visible = True
            Dim I As Integer
            dsGrid = New System.Data.DataSet("dsGrid") '12-12-08
            table = New DataTable '12-12-08
            table = dsGrid.Tables.Add("Items")
            table.Columns.Add("Column Name", GetType(String))
            table.Columns.Add("Show", GetType(Boolean))
            '02-06-11 If not in UserDir & CurrXMLFile copy from UserSysDir also ShowHideFile
            '03-19-14
            'If My.Computer.FileSystem.FileExists(UserDir & ShowHideFileSave) = False Then '02-06-11
            '    My.Computer.FileSystem.CopyFile(UserSysDir & ShowHideFileSave, UserDir & ShowHideFileSave, True) '02-06-11 True = Overwrite Overwrite no error)
            'End If
            If My.Computer.FileSystem.FileExists(ShowHideFile) Then '02-14-10        Dim fileExists As Boolean = CheckForFile(ShowHideFile, False) If fileExists = True Then
                dsGrid.ReadXml(ShowHideFile, XmlReadMode.ReadSchema)
                If dsGrid Is Nothing Then GoTo 110
                If RealCustomerOnly = True And frmQuoteRpt.cboTypeCustomer.Text.Trim.ToUpper <> "ALL" Then '03-12-14
                    'GoTo 110 'tgShow.Splits(0).DisplayColumns("Business Type").Visible = True
                    Dim HIT As Boolean = False
                    For J As Integer = 0 To dsGrid.Tables(0).Rows.Count - 1
                        If dsGrid.Tables(0).Rows(J)(0) = "Business Type" Then
                            HIT = True : Exit For
                        End If
                    Next
                    If HIT = False Then dsGrid.Clear() : GoTo 110
                Else
                    For J As Integer = 0 To dsGrid.Tables(0).Rows.Count - 1
                        If dsGrid.Tables(0).Rows(J)(0) = "Business Type" Then
                            dsGrid.Clear()
                            GoTo 110
                        End If
                    Next
                End If

                tgShow.SetDataBinding(table, "")
                GoTo 120
            End If
            ' Create a DataSet with one table containing two columns
110:        Dim show As Boolean = True ' 12-12-08 0 '' Create a DataSet with one table containing two columns
            Dim row As DataRow ' Add  rows.
            ColName = "" '12-31-14 JTC Dim ColName As String = "" '09-07-10
            'Loads All Column Names
            For I = 0 To tg.Splits(0).DisplayColumns.Count - 1
                ColName = tg.Splits(0).DisplayColumns(I).Name '09-11-10 ColName = tgShow(I, 0).ToString '09-07-10 Change ColName
                If tg.Splits(0).DisplayColumns(I).Visible = True Then show = True Else show = False '03-12-14 jh
                If ColName = "" Then Continue For '09-07-10      To                          ColName '09-07-10  
                'Debug.Print(ColName)
                'These are the invisible Columns for all three grids  Or ColName = "OverageSplit"
                If ColName = "LineID" Or ColName = "ProjectID" Or ColName = "ProdID" Or ColName = "QuoteID" Or ColName = "ProdID" Or ColName = "LPProdID" Or ColName = "ProjectCustID" Or ColName = "LnSeq" Or ColName = "Password" Or ColName = "BkSell" Or ColName = "QUOTEID1" Then GoTo 123 '08-06-13 BkSell 09-07-10 Password Skip
                If DIST Then ' Skip 
                    If ColName = "BkComm" Or ColName = "Comm-$" Or ColName = "Comm-%" Or ColName = "BKSell" Or ColName = "UOverage" Or ColName = "OverageSplit" Or ColName = "SpecCredit" Or ColName = "LPComm" Or ColName = "Overage" Then GoTo 123 '01-19-10
                Else 'Rep      'Chg LpCost to LPCost -Caps P
                    '06-24-13 If tg.Name.ToString = "tgQh" And ColName = "Cost" Then tg.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" '06-24-13
                    '03-19-14 If tg.Name.ToString = "tg" And ColName = "Cost" Or ColName = "Ext Cost" Or ColName = "LPCost" Then GoTo 123 '09-14-10 12-17-09
                    If tg.Name.ToString <> "tg" Then '12-17-09 Don't do on Line Items
                        '11-28-12 If ColName = "Cost" Or ColName = "LPCost" Then GoTo 123 '12-04-09'11-28-12
                        '03-19-14 If ColName = "LPCost" Or ColName = "QUOTEID1" Then GoTo 123 '08-06-13 12-04-09'11-28-12
                    End If
                End If
                'System.Diagnostics.Debug.Print(frmQuoteRpt.cboTypeCustomer.Text.Trim.ToUpper)
                'System.Diagnostics.Debug.Print(frmQuoteRpt.cboTypeCustomer.Visible)
                If RealCustomerOnly = True And frmQuoteRpt.cboTypeCustomer.Text.Trim.ToUpper <> "ALL" Then '03-12-14
                Else
                    If ColName = "Business Type" Then GoTo 123
                End If

                '06-25-18
                'If ShowHideFile.ToUpper.Contains("SHOWALL") = False Then '07-22-14
                '    If ColName = "Specifier" Or ColName = "Engineer" Or ColName = "Contractor" Or ColName = "Other" Or ColName = "Architect" Or ColName = "LightingDesigner" Then '07-06-17 Lighting Designer
                '        GoTo 123
                '    End If
                'End If
                '06-25-18


                row = table.NewRow()
                row("Column Name") = ColName
                row("Show") = show 'True Set to 1 = on
                table.Rows.Add(row)
123:        Next
            'Debug.Print(I.ToString)
            tgShow.SetDataBinding(table, "")
120:        tgShow.EmptyRows = True ' no dead area in the grid
            tgShow.ExtendRightColumn = True
            tgShow.Refresh()
130:        'For Save & AllOn  Set The Grids ans Save
            If dsGrid Is Nothing Then Exit Sub
            '09-07-10Dim SubCol As Integer = 0 'For the ID Columns - subtract 1 off I to keep index right in tgShow
            '09-07-10 Moved UP Dim ColName As String = "" '09-07-10 tgShow(I, 0) '"EstDelivDate")
            Dim RC As Integer = 0 '09-14-10 Row Cnt
140:        '09-12-10 tgShow.MoveLast() : HighCnt = tgShow.Row : tgShow.MoveFirst() '02-10-10
            For I = 0 To tg.Splits(0).DisplayColumns.Count - 1
                'If ShowHide = "AllOn" Then ColName = tg.Splits(0).DisplayColumns(I).Name Else ColName = tgShow(RC, 0).ToString '09-07-10 Change ColName
                'If I = 48 ThenStop
                ColName = tgShow(I, 0).ToString
                If ColName = "" Then Continue For '09-07-10      To                          ColName '09-07-10  
                'If ColName = "LPProdID" ThenStop '06-24-13 
                'If tg.Name.ToString = "tgQh" And I > 43 Then GoTo GetNxt '02-06-10
                'If DebugOn ThenDebug.PRINT(ColName) ' ColName & tg.Splits(0).DisplayColumns.Count - 1)

                Dim tf As Boolean = tgShow(I, 1)
                If ColName = "LineID" Or ColName = "ProjectID" Or ColName = "ProdID" Or ColName = "QuoteID" Or ColName = "ProdID" Or ColName = "LPProdID" Or ColName = "ProjectCustID" Or ColName = "LnSeq" Or ColName = "Password" Or ColName = "BkSell" Or ColName = "QUOTEID1" Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt '08-05-12 QUOTEID1  qtln Delete "BkSell"  09-15-10  Password Skip
                '11-04-14 JTC
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) And frmQuoteRpt.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then '11-04-14 JTC
                    If ColName = "Follow By" Then tg.Splits(0).DisplayColumns(ColName).Visible = True : GoTo GetNxt '11-04-14 
                End If
                'tg.Splits(0).DisplayColumns("Password").Visible = False
                If ShowHide = "AllOn" Then
                    '09-07-09 Skip Columns
                    If DIST Then
                        If ColName = "BkComm" Or ColName = "Comm-$" Or ColName = "Comm-%" Or ColName = "BKSell" Or ColName = "UOverage" Or ColName = "OverageSplit" Or ColName = "SpecCredit" Or ColName = "LPComm" Or ColName = "Overage" Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt '01-19-10
                    Else 'rep   'tgr tgQH tgLn
                        '09-04-10 If tg.Name.ToString = "tg" And ColName = "Cost" Or ColName = "Ext Cost" Then tg.Splits(0).DisplayColumns(I).Visible = False : GoTo GetNxt '12-17-09
                        If tg.Name.ToString = "tgQh" And ColName = "Cost" Then tg.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" : tgShow(I, 0) = ColName : GoTo 145 '06-24-13
                        '03-19-14 If ColName = "Cost" Or ColName = "Ext Cost" Or ColName = "LPCost" Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt '09-04-10 
                        If tg.Name.ToString <> "tg" Then '12-17-09 Don't do on Line Items
                            '06-24-13 JTC Delete BkSell Me.tgln
                            If ColName = "Cost" Or ColName = "LPCost" Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt '12-03-09 
                        End If
                        'or ColName = "TmpMFG") Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt
                    End If
145:
                    If ColName = "LineID" Or ColName = "ProjectID" Or ColName = "ProdID" Or ColName = "QuoteID" Or ColName = "ProdID" Or ColName = "LPProdID" Or ColName = "ProjectCustID" Or ColName = "LnSeq" Or ColName = "Password" Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt '09-07-10 added "Password"
                    '09-14-10 If I >= table.Rows.Count Then Exit For '05-04-10 Deleted .Count -1  -1
                    'Debug.Print(I.ToString)
                    tg.Splits(0).DisplayColumns(ColName).Visible = True : tgShow(RC, 1) = True '09-14-10 
                    'Debug.Print(tg.Splits(0).DisplayColumns(ColName).Visible.ToString & "" & tgShow(I, 0) & "  ColName = " & ColName)
                    RC += 1 '09-14-10 
                Else '   Show
                    '09-07-09 Skip Columns
                    If DIST Then
                        If ColName = "BkComm" Or ColName = "Comm-$" Or ColName = "Comm-%" Or ColName = "BKSell" Or ColName = "UOverage" Or ColName = "OverageSplit" Or ColName = "SpecCredit" Or ColName = "LPComm" Or ColName = "Overage" Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt '01-19-10
                    Else 'Rep
                        '03-19-14 If tg.Name.ToString = "tgQh" And ColName = "Cost" Then tg.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" : tgShow(I, 0) = ColName : GoTo 150 '06-24-13
                        '03-19-14 If tg.Name.ToString = "tg" And ColName = "Cost" Or ColName = "Ext Cost" Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt '09-07-10
                        '03-19-14 If tg.Name.ToString <> "tg" Then '12-17-09 Don't do on Line Items tgr tgh
                        '03-12-14 If ColName = "Cost" Or ColName = "LPCost" Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt ' '09-07-10 Added False
                        '03-19-14 End If
                    End If
150:                'If DebugOn ThenDebug.Print(ColName & "  " & I.ToString)
                    If ColName = "LineID" Or ColName = "ProjectID" Or ColName = "ProdID" Or ColName = "QuoteID" Or ColName = "ProdID" Or ColName = "LPProdID" Or ColName = "ProjectCustID" Or ColName = "LnSeq" Or ColName = "Password" Or ColName = "QUOTEID1" Then tg.Splits(0).DisplayColumns(ColName).Visible = False : GoTo GetNxt ''08-05-12 QUOTEID1 09-07-10 added "Password"
                    If frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) And frmQuoteRpt.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then '11-04-14 JTC
                        If ColName = "Follow By" Then tg.Splits(0).DisplayColumns(ColName).Visible = True : GoTo GetNxt '11-04-14 
                    End If
                    'If I >= table.Rows.Count ThenStop : Exit For '05-04-10 Deleted .Count -1  -1
                    '09-07-10 If tgShow(I - SubCol, 1).ToString = "" Then tgShow(I - SubCol, 1) = False

                    'LINE ITEM REPORT, ADD CUSTOMER, THEN RUN AGAIN W/O CUST - ERROR ON THESE TWO FIELDS B/C THEY AREN'T IN TRUEGRID BUT SAVED IN SHOW HIDE FILE 12-01-14 JH
                    If frmQuoteRpt.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Then
                        If (ShowHideFileName.EndsWith("VQrtLineItemsRepShowHide.xml") = True Or ShowHideFileName.EndsWith("VQrtLineItemsDistShowHide.xml") = True) And (ColName.ToUpper = "FIRMNAME" Or ColName.ToUpper = "NCODE") Then '12-01-14 
                            If frmQuoteRpt.chkUseSpecifierCode.Checked = True Or frmQuoteRpt.chkShowCustomers.Checked = True Then
                            Else
                                GoTo GetNxt
                            End If
                        End If
                    End If

                    tg.Splits(0).DisplayColumns(ColName).Visible = IIf(tgShow(I, 1).ToString = "", False, tgShow(I, 1)) '09-07-10 tg.Splits(0).DisplayColumns(I).Visible = tgShow(I 1)
                    If ColName = "SpecCross" Then  '11-11-15
                        '11-30-15 JH Stop
                        tg.Columns(ColName).DataField = "SpecCrossH"
                    End If
                    '06-24-13 Back to Original tg.Splits(0).DisplayColumns(ColName).Visible = CBool(tgShow(I, 1))
                    'Debug.Print(tg.Splits(0).DisplayColumns(ColName).Visible.ToString) '& "" & tgShow(I, 0) & "  ColName = " & ColName)
                    'Debug.Print(tgShow(I, 1).ToString & "  ColName = " & ColName)
                    RC += 1 '09-14-10 
                End If
GetNxt:     Next
            Try
                If DIST = False Then
                    If RealCustomerOnly = True And frmQuoteRpt.cboTypeCustomer.Text.Trim.ToUpper <> "ALL" Then '03-12-14
                    Else
                        tg.Splits(0).DisplayColumns("Business Type").Visible = False
                    End If
                End If
            Catch ex As Exception
            End Try

            If ShowHide = "Save" Then '12-01-14 JH (if) ONLY SAVE THIS WHEN THEY CLICK ON SAVE, NOT EVERY TIME THEY GO THROUGH HERE
                dsGrid.WriteXml(ShowHideFile, XmlWriteMode.WriteSchema)
            End If

            Me.Left = 600 '02-21-10
            Me.SetTopLevel(True)
ExitCode:

        Catch ex As Exception
            MessageBox.Show("Error in ShowHideGridCol - ColName = " & ColName & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT ", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub cboSavePrintOption_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSavePrintOption.SelectedIndexChanged

        Dim NM As String = ""
        If cboSavePrintOption.Text = "Create New..." Then
            NM = InputBox("Enter a Template Name", "MyPrint")
            If NM = "" Then Exit Sub

            Dim fileExists As Boolean = CheckForFile(UserDir & NM, False)
            If fileExists = True Then

            Else
                'Dim stm As New System.IO.FileStream(UserDir & "VPCDprintSave" & NM & ".xml", System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
                Me.cboSavePrintOption.Items.Insert(Me.cboSavePrintOption.Items.Count - 1, NM)
                Me.cboSavePrintOption.SelectedIndex = Me.cboSavePrintOption.Items.Count - 1


                Call Me.ShowHideGridCol("Save", "VQRTprintSave" & NM & ".xml", frmQuoteRpt.tgln)

                Me.cboSavePrintOption.Text = NM

            End If

        Else
            NM = Me.cboSavePrintOption.Text

            Call Me.ShowHideGridCol("Show", "VQRTprintSave" & NM & ".xml", frmQuoteRpt.tgln)

            'FillPrintCombo()


        End If
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If cboSavePrintOption.Text = "Create New..." Then Exit Sub
        If CheckForFile(UserDir & "VPCDprintSave" & Trim(cboSavePrintOption.Text) & ".xml") = True Then
            Kill(UserDir & "VPCDprintSave" & cboSavePrintOption.Text & ".xml")
        End If
        'Reset Order Form Combo
        'FillPrintCombo()
        'Reset Default Combo
        ShowDefaultFiles()
        cboSavePrintOption.Text = ""
    End Sub
End Class
