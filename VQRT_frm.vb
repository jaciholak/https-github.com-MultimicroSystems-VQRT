Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Strings
'Imports Microsoft.VisualBasic.PowerPacks
Imports C1.C1Preview
Imports MySql.Data.MySqlClient
Imports System.ComponentModel
Imports C1.Win.C1Ribbon                    '11-25-08
Imports System.Collections.Specialized     '11-25-08
Imports System.Threading '11-01-10 JTC
Imports System.Globalization '11-01-10
Imports System.Diagnostics '07-27-12
Imports System '02-24-14
Imports System.IO '02-24-14

' Imports VQRT.dsSaw8TableAdapter


Friend Class frmQuoteRpt
    Inherits System.Windows.Forms.Form
    'Attribute VB_PublicNameSpace = False
    'Private Sub cboLotUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    '    On Error Resume Next
    '    Me.cboLotUnit.SelectionStart = 0
    '    Me.cboLotUnit.SelectionLength = Len(Me.cboLotUnit.Text)
    'End Sub
    Private Sub c1Ribbon1_VisualStyleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles C1Ribbon1.VisualStyleChanged
        Select Case C1Ribbon1.VisualStyle '11-25-08
            Case C1.Win.C1Ribbon.VisualStyle.Office2007Blue
                Office2007BlueStyleButton.Pressed = True
            Case C1.Win.C1Ribbon.VisualStyle.Office2007Silver
                Office2007SilverStyleButton.Pressed = True
            Case C1.Win.C1Ribbon.VisualStyle.Office2007Black
                Office2007BlackStyleButton.Pressed = True
        End Select
        UpdateRibbonStyleMenuCheckMark()
    End Sub

    Private Sub UpdateRibbonStyleMenuCheckMark()

        Select Case C1Ribbon1.VisualStyle '11-25-08
            Case C1.Win.C1Ribbon.VisualStyle.Office2007Blue
                Office2007BlueStyleButton.Pressed = True
            Case C1.Win.C1Ribbon.VisualStyle.Office2007Silver
                Office2007SilverStyleButton.Pressed = True
            Case C1.Win.C1Ribbon.VisualStyle.Office2007Black
                Office2007BlackStyleButton.Pressed = True
        End Select
    End Sub
    Private Sub StyleButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Office2007BlackStyleButton.Click, Office2007SilverStyleButton.Click, Office2007BlueStyleButton.Click
        Dim B As C1.Win.C1Ribbon.RibbonToggleButton = CType(sender, C1.Win.C1Ribbon.RibbonToggleButton)
        If B.ID = "Office2007BlueStyleButton" Then
            C1Ribbon1.VisualStyle = C1.Win.C1Ribbon.VisualStyle.Office2007Blue
        ElseIf B.ID = "Office2007SilverStyleButton" Then
            C1Ribbon1.VisualStyle = C1.Win.C1Ribbon.VisualStyle.Office2007Silver
        ElseIf B.ID = "Office2007BlackStyleButton" Then
            C1Ribbon1.VisualStyle = C1.Win.C1Ribbon.VisualStyle.Office2007Black
        End If
    End Sub
    Private Sub RibbonColorPicker1_SelectedColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonColorPicker1.SelectedColorChanged
        Dim col As Color = RibbonColorPicker1.Color
        If col.Name = "0" Then Exit Sub 'Exit on startup
        Me.tgQh.Splits(0).AlternatingRowStyle = True
        Me.tgQh.Splits(0).OddRowStyle.BackColor = col
        'Me.tgln.Refresh()
        Me.tgQh.Rebind(True) 'true on rebind perserves column layout
        'ChangeColors(RibbonColorPicker1.Color, RibbonColorPicker2.Color, RibbonColorPicker1.Color, RibbonColorPicker2.Color, Me)
        'Dim i As Integer
        'For i = 1 To 29
        '    If i <> 20 Then Me.Fd(i).BackColor = col
        'Next
        'For i = 43 To 48
        '    If i <> 46 Then Me.Fd(i).BackColor = col
        'Next
        'txtProdID.BackColor = col
        'TextBox12.BackColor = col
        'TextBox10.BackColor = col
        'TextBox11.BackColor = col

    End Sub

    Private Sub RibbonColorPicker2_SelectedColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonColorPicker2.SelectedColorChanged
        Dim col As Color = RibbonColorPicker2.Color
        If col.Name = "0" Then Exit Sub 'Exit on startup
        Me.tgQh.Splits(0).AlternatingRowStyle = True
        Me.tgQh.Splits(0).EvenRowStyle.BackColor = col
        'Me.tgln.Refresh()
        tgQh.Rebind(True) 'true on rebind perserves column layout

    End Sub
    Private Sub RbnBtnIncreaseZoom10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '11-03-08
        Dim tmpbutton As C1.Win.C1Ribbon.RibbonButton
        tmpbutton = sender
        'MsgBox(tmpbutton.Text.ToString())
        Select Case tmpbutton.Text.ToString()
            Case "Increase Grid Zoom By 10%"
                Call zoomTG(1.1)
            Case "Increase Grid Zoom By 20%"
                Call zoomTG(1.2)
            Case "Increase Grid Zoom By 30%"
                Call zoomTG(1.3)
            Case "Decrease Grid Zoom By 10%"
                Call zoomTG(0.9)
            Case "Decrease Grid Zoom By 20%"
                Call zoomTG(0.8)
            Case "Decrease Grid Zoom By 30%"
                Call zoomTG(0.7)
        End Select

    End Sub
    Private Sub RbnTgToExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbnTgToExcel.Click
        'Me.tgr) Me.tgQh Me.tgln "Quote Summary"  "Realization") "Product Sales History - Line Items"
        If Me.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Then '03-19-12 Added
            Call ExportTgLookupToExcel(Me.tgln) '03-17-12 TByRef tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid) 
        ElseIf Me.pnlTypeOfRpt.Text = "Quote Summary" Then '03-19-12 Or Me.pnlTypeOfRpt.Text = "Total Commission Due" Or Me.pnlTypeOfRpt.Text = "Order Backlog Report" Then
            Call ExportTgLookupToExcel(Me.tgQh) '03-17-12 TByRef tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid
        ElseIf Me.pnlTypeOfRpt.Text = "Realization" Then
            Call ExportTgLookupToExcel(Me.tgr) '03-17-12 TByRef tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        Else
            MsgBox("No Grid Excel Export Option for this Report." & vbCrLf & Me.pnlTypeOfRpt.Text & vbCrLf & "Use The Preview Export Option.")
        End If

    End Sub
    Private Sub zoomTG(ByVal pcnt As Single)

        'Call zoom(CSng(0.5)) '= 50% (-0.1 = Decrease)
        ' sizes the grid to the given passed in percentage of original size
        'This sample shows how to change the size of the grid using various 
        'properties and styles to create a zoom effect. Form size stays 
        ' adjust row height
        'pcnt = 1.1 'Larger   0.9=Smaller
        tgQh.RowHeight = CInt(CSng(tgQh.RowHeight) * pcnt)
        ' and recordselector width
        tgQh.RecordSelectorWidth = CInt(CSng(tgQh.RecordSelectorWidth) * pcnt)

        ' adjust font sizes.  Normal is the root style so changing its sizes adjust all
        ' other styles
        tgQh.Styles("Normal").Font = New Font(tgQh.Styles("Normal").Font.FontFamily, tgQh.Font.Size * pcnt)

        ' now adjust the column widths
        Dim i As Integer
        For i = 0 To (tgQh.Splits(0).DisplayColumns.Count) - 1
            tgQh.Splits(0).DisplayColumns(i).Width = CInt(CSng(tgQh.Splits(0).DisplayColumns(i).Width) * pcnt)
        Next i
        'tgln.Rebind()
        'tgln.Refresh()
        tgQh.Rebind(True) 'true on rebind perserves column layout
        System.Windows.Forms.Application.DoEvents()
    End Sub 'zoom
    Private Sub RibbonGalleryItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonGalleryItem1.Click, RibbonGalleryItem2.Click, RibbonGalleryItem3.Click, RibbonGalleryItem4.Click, RibbonGalleryItem5.Click, RibbonGalleryItem6.Click, RibbonGalleryItem7.Click, RibbonGalleryItem8.Click, RibbonGalleryItem9.Click, RibbonGalleryItem10.Click, RibbonGalleryItem11.Click, RibbonGalleryItem12.Click, RibbonGalleryItem13.Click, RibbonGalleryItem14.Click, RibbonGalleryItem15.Click, RibbonGalleryItem16.Click, RibbonGalleryItem17.Click

        '
        'VNAME.My.Resources.Resources.Bullywood_AntiqueWhiteSMALL
        '10-28-08 JTC
        'Dim x As Integer
        'For x = 0 To (Me.MdiChildren.Length) - 1
        '    If Me.MdiChildren(x).Name <> "FrmQuote_Backup" Then

        Dim tempChild As frmQuoteRpt = CType(Me, frmQuoteRpt) '11-24-08
        If sender.Equals(RibbonGalleryItem1) Then
            ChangeColors(Color.Gainsboro, Color.AntiqueWhite, Color.BurlyWood, Color.AntiqueWhite, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem2) Then
            ChangeColors(Color.Gainsboro, Color.LightCyan, Color.DarkCyan, Color.LightCyan, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem3) Then
            ChangeColors(Color.Gainsboro, Color.Honeydew, Color.DarkSeaGreen, Color.Honeydew, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem4) Then
            ChangeColors(Color.Gainsboro, Color.Honeydew, Color.Honeydew, Color.PaleGreen, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem5) Then
            ChangeColors(Color.Gainsboro, Color.AliceBlue, Color.LightBlue, Color.AliceBlue, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem6) Then
            ChangeColors(Color.Gainsboro, Color.SeaShell, Color.SeaShell, Color.LightCoral, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem7) Then
            ChangeColors(Color.Gainsboro, Color.SeaShell, Color.LightSalmon, Color.SeaShell, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem8) Then
            ChangeColors(Color.Gainsboro, Color.Plum, Color.MediumOrchid, Color.Plum, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem9) Then
            ChangeColors(Color.Gainsboro, Color.MistyRose, Color.RosyBrown, Color.MistyRose, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem10) Then
            ChangeColors(Color.Gainsboro, Color.OldLace, Color.Moccasin, Color.OldLace, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem11) Then
            ChangeColors(Color.Gainsboro, Color.LightGoldenrodYellow, Color.Khaki, Color.LightGoldenrodYellow, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem12) Then
            ChangeColors(Color.Gainsboro, Color.Linen, Color.PeachPuff, Color.Linen, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem13) Then
            ChangeColors(Color.Gainsboro, Color.PapayaWhip, Color.PapayaWhip, Color.Pink, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem14) Then
            ChangeColors(Color.Gainsboro, Color.LightPink, Color.Thistle, Color.LightPink, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem15) Then
            ChangeColors(Color.Gainsboro, Color.Lavender, Color.RoyalBlue, Color.Lavender, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem16) Then
            ChangeColors(Color.Silver, Color.WhiteSmoke, Color.Silver, Color.WhiteSmoke, tempChild)
        ElseIf sender.Equals(RibbonGalleryItem17) Then
            ChangeColors(Color.Gainsboro, Color.GhostWhite, Color.SlateBlue, Color.GhostWhite, tempChild)
            'ChangeColorsReports(Color.Gainsboro, Color.AntiqueWhite, Color.BurlyWood, Color.AntiqueWhite)
        End If
        'Next

    End Sub
    Public Sub ExportTgLookupToExcel(ByRef tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid) '03-19-12
        Try
            'NewMethodExcel:
            '06-29-13 Dim FileName As String = UserPath & "DATA\QuoteReport001.Xls"
            RealTgLookupExcel = True ' As Boolean = 0 '11-27-13
            '11-27-13 strSql = "and (projectcust.TypeC = 'C'  or projectcust.TypeC = 'C' )
            '11-27-13"SELECT projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID where Quote.TypeOfJob = 'Q'  and  Quote.EntryDate >= '2009-11-01' and Quote.EntryDate <= '2013-11-30'  and (projectcust.TypeC = 'C'  or projectcust.TypeC = 'C'  or projectcust.TypeC = 'M'  or projectcust.TypeC = 'C' )  order by projectcust.Ncode, JobName "
            '11-27-13
            'If strSql.Trim = "" then 'If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then sort = sort + " All Specifiers for MFG " & RealWithOneMfgCustCode 'Stop'01-21-14 
            If tg.Name = "tgln" Then '11-21-14 - Line Item Reporting wiping out grid
                GoTo SkipFill
            End If
            '02-04-15 JTC on Realization ExportTgLookupToExcel with one Mfg and AllSpecifiers skip filling TgLookup again
            If Me.pnlTypeOfRpt.Text = "Realization" And RealTgLookupExcel = True And RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then
                GoTo SkipFill
            End If

            '02-20-17 JH why are we doing this?? Call cmdOK1_Click(cmdReportQuote, New System.EventArgs()) '01-21-14 Have to go thru here to get specifiers
            '02-20-17 - will group - export real to excel, hangs on this trying to get spec
SkipFill:
            Dim ExcelCol As Short = 1

            Dim R As Integer = 0 ' tg.RowCount - 1 '
            Dim MaxRow As Integer = tg.RowCount - 1 '
            Dim C As Integer = 0 'tg.Splits(0).DisplayColumns.Count - 1
            Dim MaxCol As Integer = tg.Splits(0).DisplayColumns.Count - 1
            'Dim LineData As String
            Static EndingLine As Short
            Static StartingLine As Short
            'Dim iCol As Int32 '        'For iRow = 0 To 1 'ProdDataArray
            Dim objExcel As Object = Nothing
            Dim objBooks As Object = Nothing
            Dim objSheets As Object = Nothing
            Dim objSheet As Object = Nothing
            Dim objBook As Object = Nothing
            Dim A As String = ""
            objExcel = CreateObject("Excel.Application")
            objExcel.DisplayAlerts = False
            objBook = objExcel.Workbooks.Add()
            objBooks = objExcel.Workbooks
            objBook = objBooks(1)
            objSheet = objBook.ActiveSheet()
            objSheet.Name = "Lookup Grid to Excel" '03-23-12
            objSheets = objBook.Worksheets
            StartingLine = 0 : EndingLine = 0 '10-19-11
5038:
            tg.MoveFirst()
            R = 0 'Excel Starts with 1 not Zero !@#$%^&* Rows & Columns
            Dim PrtCols As Integer = 1
            If RealWithOneMfgCustCode.Trim <> "" Then '07-24-14 JTC
                objSheet.Cells(1, 1) = "All Specifiers for Code = " & RealWithOneMfgCustCode '01-21-14If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then sort = sort + "  'Stop'01-21-14    objSheet.Cells(R + 2, PrtCols) = tg.Splits(0)rtCols) = tg.Splits(0)
            End If
            ExcelCol = 1
            R = 1 '01-21-14 
            For I = 0 To MaxCol 'Header
                'Debug.Print(tg.Splits(0).DisplayColumns(I).Name)
                If tg.Splits(0).DisplayColumns(I).Visible = False Then Continue For
                If (tg.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
                '07-31-14JTC Fix Header no Plus 1 objSheet.Cells(R + 1, PrtCols) = tg.Splits(0).DisplayColumns(I).Name
                objSheet.Cells(R, PrtCols) = tg.Splits(0).DisplayColumns(I).Name
                Dim ColName As String = tg.Splits(0).DisplayColumns(I).Name
                'Debug.Print(ColName)
                If ColName = "JobName" Or ColName = "QuoteCode" Or ColName = "SourceQuote" Or ColName = "SLSQ" Or ColName = "NCode" Then '11-30-15 NCode
                    objSheet.Columns(ExcelCol).NumberFormat = "@" '09-10-12 JH Fix Excel from converting Quote Code to Date format 12-0123 in ExportTgLookupToExcel
                End If
                Dim Width As Integer = tg.Splits(0).DisplayColumns(I).Width '03-23-12
                If PrtCols < 27 And Width > 50 Then objSheet.Columns(ExcelCol).ColumnWidth = Width / 7 '03-23-12 
                PrtCols += 1 : ExcelCol += 1
            Next I
            PrtCols = 1

            For R = 0 To MaxRow
                ExcelCol = 1
                For C = 0 To MaxCol ' tgLookup.RowCount - 1
                    If tg.Splits(0).DisplayColumns(C).Visible = False Then Continue For
                    If (tg.Splits(0).DisplayColumns(C).Width / 100) < 0.1 Then Continue For
                    objSheet.Cells(R + 2, PrtCols) = tg.Splits(0).DisplayColumns(C).DataColumn.Text
                    'ExcelCol = 1 '= Chr(64 + PrtCols)'07-24-14 JH Fix Error on Col 26 ExportTGLoolupToExcel
                    'Debug.Print(tg.Splits(0).DisplayColumns(C).DataColumn.Text)
                    'objSheet.Columns(ExcelCol).NumberFormat = "@" '09-10-12 JH Fix Excel from converting Quote Code to Date format 12-0123 in ExportTgLookupToExcel
                    Dim Width As Integer = tg.Splits(0).DisplayColumns(C).Width '03-23-12
                    If PrtCols < 27 And Width > 50 Then objSheet.Columns(ExcelCol).ColumnWidth = Width / 7 '03-23-12 
                    'objSheet.Columns("B").ColumnWidth = 12 : objSheet.Columns("C").ColumnWidth = 18
                    PrtCols += 1 : ExcelCol += 1
                Next C
                '03-29-12 Start Could copy cells and put QuoteTo Info At end of the line
                'objSheet = objExcel.Sheets.Item(2)
                'objSheet.Cells.CurrentRegion.Select()
                'objSheet.Cells.Copy()
                'copy the A3-E5 to A20-E22
                'objSheet.Copy(objSheet.Range("A:Z"), objSheet.Range("A20:E22"), True)
                'delete cell data A3 but not its format
                'objSheet.Range("A3").Clear(ExcelClearOptions.ClearContent)
                'delete the format of A2
                'objsheet.Range["A2"].Clear(ExcelClearOptions.ClearFormat)
                'delete cell data B11 including its format
                'objsheet.Range["B11"].Clear(ExcelClearOptions.ClearAll)
                'insert "Island" to B11
                'objsheet.Range["B11"].Text = "Island"
                'delete the sixth and seventh rows from worksheet
                'objSheet.DeleteRow(6, 2)
                'delete the third column of worksheet
                'objSheet.DeleteColumn(3)
                'autofit the fifth column
                'objSheet.AutoFitColumn(5)
                'Create an array with 3 columns and 100 rows.
                'Dim DataArray(99, 2) As Object
                'Dim r As Integer
                'For r = 0 To 99
                '    DataArray(r, 0) = "ORD" & Format(r + 1, "0000")
                '    DataArray(r, 1) = Rnd() * 1000
                '    DataArray(r, 2) = DataArray(r, 1) * 0.07
                'Next

                ''Add headers to the worksheet on row 1.
                'oSheet = oBook.Worksheets(1)
                'oSheet.Range("A1").Value = "Order ID"
                'oSheet.Range("B1").Value = "Amount"
                'oSheet.Range("C1").Value = "Tax"

                ''Transfer the array to the worksheet starting at cell A2.
                'oSheet.Range("A2").Resize(100, 3).Value = DataArray


                '03-29-12 End test Could copy cells and put QuoteTo Info At end of the line
                PrtCols = 1
                tg.MoveNext()
            Next R 'Row
            'FileName += ".csv" '06-17-10 
            'FileName = SaveDialog(FileName, "Export to CSV", "csv Files (*.csv)|*.csv")
            'If FileNaTrim = "" Then GoTo NoFileName '06-16-10 
            'Done Above FileName = UserPath & "DATA\QuoteReport001.Xls"
            '04-10-12 Test*************************************************************************
            '04-20-12 not needed twice????R = 0 'Excel Starts with 1 not Zero !@#$%^&*
            'PrtCols = 1 'A-Z,AA-AZ,BA
            'For I = 0 To MaxCol 'Header
            '    If tg.Splits(0).DisplayColumns(I).Visible = False Then Continue For
            '    If (tg.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
            '    objSheet.Cells(R + 1, PrtCols) = tg.Splits(0).DisplayColumns(I).Name
            '    Dim ExcelCol As String = Chr(64 + PrtCols)
            '    Dim Width As Integer = tg.Splits(0).DisplayColumns(I).Width '03-23-12
            '    If PrtCols < 27 And Width > 50 Then objSheet.Columns(ExcelCol).ColumnWidth = Width / 7 '03-23-12 
            '    PrtCols += 1
            'Next I
            'PrtCols = 1
            'For R = 0 To MaxRow
            '    For C = 0 To MaxCol ' tgLookup.RowCount - 1
            '        If tg.Splits(0).DisplayColumns(C).Visible = False Then Continue For
            '        If (tg.Splits(0).DisplayColumns(C).Width / 100) < 0.1 Then Continue For
            '        objSheet.Cells(R + 2, PrtCols) = tg.Splits(0).DisplayColumns(C).DataColumn.Text
            '        Dim ExcelCol As String = Chr(64 + PrtCols)
            '        Dim Width As Integer = tg.Splits(0).DisplayColumns(C).Width '03-23-12
            '        If PrtCols < 27 And Width > 50 Then objSheet.Columns(ExcelCol).ColumnWidth = Width / 7 '03-23-12 
            '        'objSheet.Columns("B").ColumnWidth = 12 : objSheet.Columns("C").ColumnWidth = 18
            '        PrtCols += 1
            '    Next C
            '      PrtCols = 1
            '    tg.MoveNext()
            'Next R 'Row

            '06-30-13 Start File In Use *************************************************************************
            '11-30-15 JH - LET THEM SAVE IT ANYWHERE
            '            Dim FileNum As Short = 1 '06-30-13
            '            Dim FileNumStr As String = "001"
            '            Dim FileName As String = UserPath & "DATA\QuoteReport001" & UserID & ".Xls"
            'StartFileN:
            '            Dim fileInUse2 As Boolean = False
            '            Dim F As Short = FreeFile()
            '            If My.Computer.FileSystem.FileExists(FileName) Then
            '                Try             'fs = System.IO.File.Open(FileName, IO.FileMode.Open, IO.FileAccess.Write, IO.FileShare.None)
            '                    FileOpen(F, FileName, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
            '                    fileInUse2 = False
            '                    FileClose(F) 'Stop' Return False
            '                Catch ex As Exception
            '                    FileNum += 1
            '                    FileName = UserPath & "DATA\QuoteReport" & Format(FileNum, "000") & UserID & ".Xls"
            '                    fileInUse2 = True 'Stop' Return True
            '                End Try
            '                If FileNum > 15 Then MessageBox.Show("Error in ExportTgLookup (VQRT)" & vbCrLf & FileName & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12 
            '                If fileInUse2 = True Then GoTo StartFileN
            '            End If '06-30-13 End File In Use *************************************************************************

            '            If My.Computer.FileSystem.FileExists(FileName) Then ' , False)
            '                Kill(FileName)
            '            End If '01-06-09
            '            objBook.SaveAs(FileName, 39, , , False, False, , False, False, , , ) 'MessageBox.Show("All Done   Created " & FileName)
            '11-30-15 JH - LET THEM SAVE IT ANYWHERE

            objExcel.VISIBLE = True 'Show User the Excel Report
NoFileName:

        Catch ex As Exception
            MessageBox.Show("Error in ExportTgLookup (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12 
        End Try

        RealTgLookupExcel = False ' As Boolean = 0 '11-27-13
    End Sub
    Public Sub ChangeColors(ByVal Back As Color, ByVal Fore As Color, ByVal GridBack As Color, ByVal GridFore As Color, ByVal frm As frmQuoteRpt) '11-24-08

        Me.RibbonColorPicker1.Color = Back
        Me.RibbonColorPicker2.Color = Fore
        Dim n As Integer
        For n = 0 To tgQh.Splits.Count - 1 '10-30-08
            Me.tgQh.Splits(n).AlternatingRowStyle = True
            Me.tgQh.Splits(n).OddRowStyle.BackColor = GridBack
            Me.tgQh.Splits(n).EvenRowStyle.BackColor = GridFore
        Next n

        'frm.rtNotes.BackColor = Fore 'NAME DETAIL
        'frm.FldPhone.BackColor = Fore
        'frm.FldFax.BackColor = Fore

        'frm.dgNameDefaults.RowsDefaultCellStyle.BackColor = GridFore 'NAME DETAIL
        'frm.dgNameDefaults.AlternatingRowsDefaultCellStyle.BackColor = GridBack 'NAME DETAIL

        'frm.tgCustomerNumbers.Splits(0).AlternatingRowStyle = True
        'frm.tgCustomerNumbers.Splits(0).OddRowStyle.BackColor = GridBack
        'frm.tgCustomerNumbers.Splits(0).EvenRowStyle.BackColor = GridFore

        'frm.lstNDCategory.BackColor = Fore
        'frm.tbNDefaultsCustomerCode.BackColor = Fore
        'frm.txtNDefaultsFirmname.BackColor = Fore
        'frm._FldC_0.BackColor = Fore
        'frm.cboCategory.BackColor = Fore
        'frm._FldC_2.BackColor = Fore
        'frm._FldC_3.BackColor = Fore
        'frm._FldC_4.BackColor = Fore
        'frm._FldC_5.BackColor = Fore
        'frm._FldC_6.BackColor = Fore
        'frm.cboRecordType.BackColor = Fore

        'FrmLookup.tgln.Splits(0).AlternatingRowStyle = True
        'FrmLookup.tgln.Splits(0).OddRowStyle.BackColor = GridBack
        'FrmLookup.tgln.Splits(0).EvenRowStyle.BackColor = GridFore

        'Dim i As Integer
        'For i = 0 To 36
        '    frm.Fld(i).BackColor = Fore
        'Next

        'FrmLookup.txtRC1.BackColor = Fore
        'FrmLookup.txtRC2.BackColor = Fore
        'FrmLookup.txtRC3.BackColor = Fore
        'FrmLookup.txtRC4.BackColor = Fore
        'FrmLookup.txtRC5.BackColor = Fore
        'FrmLookup.txtRC6.BackColor = Fore
        'FrmLookup.txtRC7.BackColor = Fore
        'FrmLookup.txtRC8.BackColor = Fore
        'FrmLookup.txtRC9.BackColor = Fore
        'FrmLookup.txtRC10.BackColor = Fore

    End Sub
    Private Sub rbnExitMainMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbnExitMainMenu.Click
        Call Jump("VMENU.EXE") : Me.Close() '04-18-10
    End Sub
    Private Sub RbnBtnGridViewNormal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbnBtnGridViewNormal.Click, RbnBtnGridViewGroupBy.Click, RbnBtnGridViewSplit.Click, RbnBtnGridViewExpandGp.Click, RbnBtnGridViewCollapse.Click, rbnAddFilterBar.Click
        '10-27-08 Normal View Inverted View  Multiple Line View Group By View
        'Select Case RbnBtnGridViewNormal_Click.text
        '10-03-08 If sender.Equals(RibbonGalleryItem1) Then
        Dim tmpbutton As C1.Win.C1Ribbon.RibbonButton
        tmpbutton = sender
        'MsgBox(tmpbutton.Text.ToString())

        Select Case tmpbutton.Text.ToString()
            Case "Expand All Grouped Rows" 'RbnBtnGridViewExpandGp
                'See if You have groupBy set
                Dim rtype As C1.Win.C1TrueDBGrid.RowTypeEnum = Me.tgQh.Splits(0).Rows(Me.tgQh.Row).RowType
                If rtype <> C1.Win.C1TrueDBGrid.RowTypeEnum.CollapsedGroupRow Then
                    MsgBox("Their are no Collapsed Rows to Expand" & vbCrLf & "Put the Mfg or CustCode Header in the GroupBy Area")
                End If
                'Select Case rtype
                '    Case C1.Win.C1TrueDBGrid.RowTypeEnum.CollapsedGroupRow
                '         'Me.radioButton2.Checked = CheckState.Checked
                '    Case C1.Win.C1TrueDBGrid.RowTypeEnum.ExpandedGroupRow
                '         'Me.radioButton1.Checked = CheckState.Checked
                'End Select

                ''''''''''''''''''''''''''''
                Dim i As Integer = 0
                ' can't use for...next, doesn't re-evaluate the upper bound and the Rows collection does change as you expand/collapse
                While (i < Me.tgQh.Splits(0).Rows.Count)
                    Me.tgQh.ExpandGroupRow(i)
                    i = i + 1
                End While
            Case "Collapse All Grouped Rows"  'RbnBtnGridViewCollapse
                Dim rtype As C1.Win.C1TrueDBGrid.RowTypeEnum = Me.tgQh.Splits(0).Rows(Me.tgQh.Row).RowType
                If rtype <> C1.Win.C1TrueDBGrid.RowTypeEnum.ExpandedGroupRow Then
                    MsgBox("Their are no Expanded Rows to Collapse" & vbCrLf & "Expand Rows Before you can Collapse" & vbCrLf & "Put the Mfg or CustCode Header in the GroupBy Area") '11-04-08
                End If
                Dim i As Integer = 0
                ' can't use for...next, doesn't re-evaluate the upper bound and the Rows collection does change as you expand/collapse
                While (i < Me.tgQh.Splits(0).Rows.Count)
                    'If Me.tgln.Splits(0).Rows(i).RowType <> C1.Win.C1TrueDBGrid.RowTypeEnum.DataRow Then
                    Me.tgQh.CollapseGroupRow(i)
                    'End If
                    i = i + 1
                End While

            Case "Add Split Grid View"
                If tgQh.Splits.Count > 2 Then
                    Dim i As Integer = 0
                    For i = 2 To tgQh.Splits.Count : tgQh.RemoveHorizontalSplit(0) : Next '  tgln.RemoveHorizontalSplit(0) ': tgln.RemoveHorizontalSplit(0) '11-24-08
                Else
                    tgQh.InsertHorizontalSplit(1) '10-30-08 Color Does not appear on Split 2
                End If

            Case "Normal Grid View"
                Me.tgQh.DataView = C1.Win.C1TrueDBGrid.DataViewEnum.Normal
                ' This is OK also Me.tgln.GroupByAreaVisible = False 'Turn Off GroupBy
            Case "Add Filter Bar to Grid"
                Me.tgQh.FilterBar = True '06-08-09
                'Case "Form Grid View"
                '    Me.tgln.DataView = C1.Win.C1TrueDBGrid.DataViewEnum.Form
            Case "GroupBy Grid View" 'set dataview to outlook-style
                Me.tgQh.DataView = C1.Win.C1TrueDBGrid.DataViewEnum.GroupBy
                Me.tgQh.GroupStyle.GradientMode = C1.Win.C1TrueDBGrid.GradientModeEnum.BackwardDiagonal
                'Case "MultiLine Grid View"
                '    Me.tgln.DataView = C1.Win.C1TrueDBGrid.DataViewEnum.MultipleLines
                'Case "Hierarchical"
                '    MsgBox("This only has an affect if the grid is displaying a hierarchical data set")
                '    Me.tgln.DataView = C1.Win.C1TrueDBGrid.DataViewEnum.Hierarchical
        End Select
        ' MsgBox(e.ToString)

    End Sub

    Private Sub RbnBtnExportExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbnBtnExportExcel.Click, RbnBtnExportCSVTab.Click, RbnBtnExportCSVComma.Click, RbnBtnExportHTML.Click, RbnBtnExportPDF.Click, RbnBtnExportRTF.Click, RbnBtnExportOptions.Click, RbnBtnExportPrint.Click
        'Maybe check for overwrite?
        ExportExcelProductLines = False '02-15-10
        '06-11-10 Me.tgQh.DataView = C1.Win.C1TrueDBGrid.DataViewEnum.Normal 'No Groupby Bar
        Dim FileName As String = "VqrtQuoteReport"
        Dim tmpbutton As C1.Win.C1Ribbon.RibbonButton
        tmpbutton = sender
        If tmpbutton.Text.ToString = "Print Reports" Then '06-11-10 
            If ppv.CanSelect Then '06-11-10   'Debug.Print(ppv.Text)
                Resp = MsgBox("Go to Print Preview form on your task bar and Click on Print.", MsgBoxStyle.OkOnly, "Print Quote Report")
                'ppv.Show()'Doesn't Seem to work ???
                ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
                ppv.BringToFront()
                ppv.Focus()
                Exit Sub
            Else
                Resp = MsgBox("You must run the report to fill the Print Preview form, then Print.", MsgBoxStyle.OkOnly, "Print Quote Report")
                Exit Sub
            End If
        End If
        If ppv.CanSelect Then '06-01-10   'Debug.Print(ppv.Text)
            Resp = MsgBox("The current Report in the Print Preview form Will be Exported OK?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Export to Excel")
            If Resp <> vbYes Then Exit Sub
        Else
            Resp = MsgBox("Their is no Report in the Print Preview form to Export." & vbCrLf & " You must run the report to fill the Print Preview form first.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Export to Excel")
            Exit Sub
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' 

        'Case "Print Reports" ' 12-29-08 "Export To Printer" '11-11-08
        FileName = "VQrtQuoteReport" '06-11-10 ShowHideRepPrtHdr.xml"
        'If DIST Then '01-19-10
        '    FileName = "VQrtQuoteReport" '06-11-10 ShowHideDistPrtHdr.xml" '
        If Me.pnlTypeOfRpt.Text = "Realization" Then
            FileName = "VQrtQuoteReport" '06-11-10 ShowHideRepPrtHdr.xml" '
        End If
        If Me.pnlTypeOfRpt.Text = "Realization" Then
            '06-11-10 If DIST Then FileName = "VQrtRealQTOShowHideDistPrint.xml" Else FileName = "VQrtRealQTOShowHideRepPrint.xml" '05-08-10 
            FileName = "VQrtQuoteRealization" '06-11-10 
        End If
        If Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then '06-17-10
            FileName = "VQrtProductHistory" '06-17-10 
        End If
        'ShowHideFileName = FileName '09-02-09
        'frmShowHideGrid.Text = "Show Hide Printing Columns" '12-22-08
        'frmShowHideGrid.Show() '12-15-08
        'ShowHideFileName = FileName '09-02-09
        'Call frmShowHideGrid.ShowHideGridCol("Show", FileName, Me.tgr) Me.tgQh
        'Me.tabQrt.SelectTab(2) '12-29-08
        'Exit Sub         '"",""
        'GoTo 120 ' tg doc method Not Used 
115:    'Doc Example
        'ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)
        'Me.Size = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        'IfDebugOn Then Sto 'Fix24
        ''MakeDoc1(ppv.Doc) 'testing Doc Options
        'ppv.Doc.Generate()
        'ppv.Show()
        ''ppv.Focus()
        'ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        'ppv.BringToFront()
120:    'Exit Sub
140:    Try
            Select Case tmpbutton.Text.ToString()
                Case "Export To Excel"
                    FileName += ".xls" '06-02-10 
                    FileName = SaveDialog(FileName, "Export to Excel", "Excel Files (*.xls)|*.xls")
                    If FileName.Trim = "" Then GoTo NoFileName '06-16-10 
                    If FileName.Trim = "" Then GoTo NoFileName '06-16-10 
                    Dim xl As New C1.C1Preview.Export.XlsExporter
                    xl.Document = ppv.C1PrintPreviewControl1.PreviewPane.Document '   'xl.Document = Myppv.C1PrintPreviewControl1.PreviewPane.Document
                    If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    xl.Preview = True 'If You Preview you don't need Shell
                    xl.Export(FileName)
                    Cursor = Cursors.Default
                    GoTo PrintFollowReportExit '01-06-09
                    'If Me.pnlTypeOfRpt.Text = "Realization" Then
                    '    Call PrintRealizationReportQutTO() '02-18-09
                    'ElseIf Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then
                    '    ExportExcelProductLines = True '02-15-10
                    '    Call PrintReportQuoteLines() '09-13-09 
                    '    ExportExcelProductLines = False '02-15-10
                    'Else
                    '    Call PrintReportQuotes()
                    'End If

                    'Dim xl As New C1.C1Preview.Export.XlsExporter
                    'xl.Document = ppv.C1PrintPreviewControl1.PreviewPane.Document
                    'FileName = UserDocDir & "QuoteReportTmp.xls"
                    'If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    'xl.Export(FileName)
                    'Cursor = Cursors.Default
                    'Dim objShell '12-24-08 
                    'objShell = CreateObject("Shell.Application")
                    'objShell.ShellExecute("excel.exe", UserDocDir & "QuoteReportTmp.xls", "", "open", 1)
                    'objShell = Nothing
                    GoTo PrintFollowReportExit '01-06-09
180:            Case "Export To CSV Tab DeLimited"
                    FileName += ".txt" '06-17-10                                              (*.csv)|*.csv")
                    FileName = Me.SaveDialog(FileName, "Export to TXT", "txt Files (*.txt)|*.txt")
                    If FileName.Trim = "" Then GoTo NoFileName '06-16-10 
                    If My.Computer.FileSystem.FileExists(FileName) Then ' , False)
                        Kill(FileName)
                    End If '01-06-09
                    FileClose(3) : FileOpen(3, FileName, OpenMode.Output)
                    Dim R As Integer = RT.Rows.Count - 1
                    Dim C As Integer = RT.Cols.Count - 1
                    Dim LineData As String
                    For R = 1 To RT.Rows.Count - 1
                        LineData = ""
                        For C = 0 To RT.Rows.Count - 1
                            If RT.Cols(C).Visible = True Then LineData = LineData & RT.Cells(R, C).Text & Microsoft.VisualBasic.Constants.vbTab '01-06-09
                            'If BeginCode = "CSV Comma" And RT.Cols(C).Visible = True Then LineData = LineData & RT.Cells(R, C).Text & "," '01-06-09
                        Next C
                        PrintLine(3, LineData) 'eol
                    Next R 'Row
                    FileClose(3)
                    Cursor = Cursors.Default
                    MsgBox(FileName & vbCrLf & "File Created") '07-20-10
                    GoTo PrintFollowReportExit '01-02-09
                    '****************************************************************************
                    'BeginCode = "CSV Tab" '01-06-09 Public Used in Print Report
                    'If Me.pnlTypeOfRpt.Text = "Realization" Then
                    '    Call PrintRealizationReportQutTO() '02-18-09
                    'ElseIf Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then
                    '    Call PrintReportQuoteLines() '09-13-09 
                    'Else
                    '    Call PrintReportQuotes()
                    'End If
                    'BeginCode = "" 'Off ' "CSV Tab" '01-06-09 Public Used in Print Report
                    'Cursor = Cursors.Default
                    'Dim xl As New C1.C1Preview.Export.XlsExporter
                    'xl.Document = ppv.C1PrintPreviewControl1.PreviewPane.Document
                    'FileName = UserDocDir & "QuoteReportTmp.csv"
                    'If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    'xl.Preview = True
                    'xl.Export(FileName)
                    ''System.Diagnostics.Process.Start(FileName)
                    ''BeginCode = "" '01-06-09 Public
                    'GoTo PrintFollowReportExit '01-02-09

200:            Case "Export To CSV Comma DeLimited"
                    FileName += ".csv" '06-17-10 
                    FileName = SaveDialog(FileName, "Export to CSV", "csv Files (*.csv)|*.csv")
                    If FileName.Trim = "" Then GoTo NoFileName '06-16-10 
                    'UserPathData
                    If My.Computer.FileSystem.FileExists("C:\Cahill\" & FileName) Then ' , False)
                        Kill(FileName)
                    End If '01-06-09
                    FileClose(3) : FileOpen(3, FileName, OpenMode.Output)
                    '03-11-12 JTC Grid To Excel Start******************************************
                    Dim R As Integer = 0 'RT.Rows.Count - 1
                    Dim C As Integer = 0 ' RT.Cols.Count - 1
                    Dim LineData As String = ""
                    Dim Firstloop As Int16 = 0
                    'For R = 0 To tgQh.Splits(0).Rows.Count - 1
                    drQRow = dsQutLU.QUTLU1.Rows(0)
                    For Each drQRow In dsQutLU.QUTLU1.Rows 'dsQutLU
                        If drQRow.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                        If Firstloop = 0 Then tgQh.MoveFirst() : LineData = "" : Firstloop = 1
                        For I = 0 To Me.tgQh.Splits(0).DisplayColumns.Count - 1
                            If Me.tgQh.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                            If (Me.tgQh.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
                            LineData += Me.tgQh.Splits(0).DisplayColumns(I).DataColumn.Text & "," 'vbTab ' ","
                        Next
                        PrintLine(3, LineData) 'eol
GetNext:                LineData = ""
                        tgQh.MoveNext()
                    Next 'Row
                    FileClose(3)
                    Cursor = Cursors.Default
                    MsgBox(FileName & vbCrLf & "File Created") '07-20-10
                    GoTo PrintFollowReportExit '01-02-09
                    'End **********************************************************************
                    R = RT.Rows.Count - 1
                    C = RT.Cols.Count - 1
                    LineData = ""
                    For R = 1 To RT.Rows.Count - 1
                        LineData = ""
                        For C = 0 To RT.Rows.Count - 1
                            '06-16-10 If RT.Cols(C).Visible = True Then LineData = LineData & RT.Cells(R, C).Text & Microsoft.VisualBasic.Constants.vbTab '01-06-09
                            If RT.Cols(C).Visible = True Then LineData = LineData & RT.Cells(R, C).Text & "," '06-16-10
                        Next C
                        PrintLine(3, LineData) 'eol
                    Next R 'Row
                    FileClose(3)
                    Cursor = Cursors.Default
                    MsgBox(FileName & vbCrLf & "File Created") '07-20-10
                    GoTo PrintFollowReportExit '01-02-09
                    'BeginCode = "CSV Comma" '01-06-09 Public Used in Print Report
                    'If Me.pnlTypeOfRpt.Text = "Realization" Then
                    '    Call PrintRealizationReportQutTO() '02-18-09
                    'ElseIf Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then
                    '    Call PrintReportQuoteLines() '09-13-09 
                    'Else
                    '    Call PrintReportQuotes()
                    'End If
                    ''FileName = UserDocDir & "QuoteReportTmp.csv" 'BeginCode = "CSV Tab" '01-06-09 Public
                    ''If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    'Cursor = Cursors.Default
                    'FileName = UserDocDir & "QuoteReportTmp.csv"
                    'Dim objShell '01-07-09
                    'objShell = CreateObject("Shell.Application")
                    'objShell.ShellExecute("excel.exe", FileName, "", "open", 1) '09-13-09 
                    'objShell = Nothing
                    'BeginCode = "" '01-06-09 Public
                    GoTo PrintFollowReportExit '01-02-09Me.tgr) Me.tgQh Me.tgln
                    'Me.tgln.ExportToDelimitedFile(TmpDir & TmpFileName, C1.Win.C1TrueDBGrid.RowSelectorEnum.AllRows, ",")

220:            Case "Export To HTML"
                    'Ugly Format Job Name Spacing is wrong (IE: Last Column on Wide Reports)
                    FileName += ".html" '06-02-10 
                    FileName = Me.SaveDialog(FileName, "Export to HTML", "HTML Files (*.html)|*.html")
                    If FileName.Trim = "" Then GoTo NoFileName '06-16-10 
                    Dim xl As New C1.C1Preview.Export.HtmlExporter
                    xl.Document = ppv.C1PrintPreviewControl1.PreviewPane.Document
                    If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    xl.Preview = True
                    xl.Export(FileName)
                    Cursor = Cursors.Default
                    GoTo PrintFollowReportExit '01-02-09
                    'BeginCode = "Export To HTML" '01-07-09
                    'If Me.pnlTypeOfRpt.Text = "Realization" Then
                    '    Call PrintRealizationReportQutTO() '02-18-09
                    'ElseIf Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then
                    '    Call PrintReportQuoteLines() '09-13-09 
                    'Else
                    '    Call PrintReportQuotes()
                    'End If
                    'FileName = UserDocDir & "QuoteReportTmp.HTML" 'BeginCode = "CSV Tab" '01-06-09 Public
                    'If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    'Dim exp As New C1.C1Preview.Export.HtmlExporter
                    'exp.Preview = True '01-07-09 Set Dialog Parameters User Clicks OK
                    'exp.Paginated = False '01-07-09
                    'exp.Range = New C1.C1Preview.OutputRange(1, ppv.C1PrintPreviewControl1.PreviewPane.Pages.Count) '((1, 3)
                    'exp.Document = ppv.C1PrintPreviewControl1.PreviewPane.Document
                    'exp.Export(FileName)
                    ''System.Diagnostics.Process.Start(FileName)
                    'BeginCode = ""
                    'Cursor = Cursors.Default
                    'GoTo PrintFollowReportExit '01-02-09
                    'TmpFileName = TmpFileName & ".html"
                    'fileExists = CheckForFile(TmpDir & TmpFileName, False)
                    'If fileExists = True Then
                    '    Kill(TmpDir & TmpFileName)
                    'End If
                    'Me.tgln.ExportToHTML(TmpDir & TmpFileName)
240:            Case "Export To PDF"
                    FileName += ".pdf" '06-02-10 
                    FileName = SaveDialog(FileName, "Export to PDF", "PDF Files (*.pdf)|*.pdf") '"Export to PDF", "Adobe PDF *.pdf)|*.pdf")
                    If FileName.Trim = "" Then GoTo NoFileName '06-16-10 
                    Dim xl As New C1.C1Preview.Export.PdfExporter
                    xl.Document = ppv.C1PrintPreviewControl1.PreviewPane.Document
                    If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    xl.Preview = True
                    xl.Export(FileName)
                    Cursor = Cursors.Default
                    GoTo PrintFollowReportExit '01-02-09

                Case "Export To RTF"
                    FileName += ".rtf" '06-02-10 
                    FileName = SaveDialog(FileName, "Export to RTF", "RTF Files (*.rtf)|*.rtf")
                    If FileName.Trim = "" Then GoTo NoFileName '06-16-10 
                    Dim xl As New C1.C1Preview.Export.RtfExporter
                    xl.Document = ppv.C1PrintPreviewControl1.PreviewPane.Document
                    If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    xl.Preview = True
                    xl.Export(FileName)
                    Cursor = Cursors.Default
                    GoTo PrintFollowReportExit '01-02-09
                Case "Export To Other Options"
                    MsgBox("Their are no other options at this time.")
                    GoTo PrintFollowReportExit
                    'If Me.pnlTypeOfRpt.Text = "Realization" Then
                    '    Call PrintRealizationReportQutTO() '02-18-09
                    'ElseIf Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then
                    '    Call PrintReportQuoteLines() '09-13-09 
                    'Else
                    '    Call PrintReportQuotes()
                    'End If
                    'FileName = UserDocDir & "QuoteReportTmp.Pdf"
                    'If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    'Dim exp As New C1.C1Preview.Export.PdfExporter
                    'exp.Range = New C1.C1Preview.OutputRange(1, ppv.C1PrintPreviewControl1.PreviewPane.Pages.Count) '((1, 3)
                    'exp.Document = ppv.C1PrintPreviewControl1.PreviewPane.Document
                    'exp.Export(FileName)
                    ''Gives a dialog box to user Click OK
                    'Cursor = Cursors.Default
                    ''System.Diagnostics.Process.Start(FileName)
                    'BeginCode = "" '01-06-09 Public
                    'MsgBox("Document Created = " & FileName) '02-18-09
                    'GoTo PrintFollowReportExit '01-02-09

260:            Case "Export To RTF"
                    If Me.pnlTypeOfRpt.Text = "Realization" Then
                        Call PrintRealizationReportQutTO() '02-18-09
                    ElseIf Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then
                        Call PrintReportQuoteLines() '09-13-09 
                    Else
                        Call PrintReportQuotes()
                    End If
                    FileName = UserDocDir & "QuoteReportTmp.rtf" '06-21-09
                    If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName) '07-28-09
                    Dim rt As New C1.C1Preview.Export.RtfExporter
                    rt.Preview = True '01-07-08
                    rt.OpenXmlDocument = True
                    rt.Paginated = True
                    rt.Document = ppv.C1PrintPreviewControl1.PreviewPane.Document
                    rt.Export(FileName)
                    'Gives a dialog box to user Click OK
                    Cursor = Cursors.Default
                    GoTo PrintFollowReportExit

280:            Case "Export To Other Options"
                    MsgBox("Their are no other options at this time.")
                    GoTo PrintFollowReportExit
            End Select
        Catch ex As Exception '06-16-10 
            MessageBox.Show("Error in RbnBtnExportCSVComma (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12  MsgBox("Error during the Export To Excel Procedure" & vbCrLf & ex.Message.ToString)
        End Try
        '07-20-10 MsgBox(UserDocDir & FileName & vbCrLf & "File Created") '11-04-06
        GoTo PrintFollowReportExit
NoFileName:
        MsgBox("No File Name, Please try again.") '06-16-10 
        GoTo PrintFollowReportExit
        '' we want the last column on page to stretch to the right edge of the page,
        '' so that there is no white space left before the margin
        'rt.Stretch = StretchTableEnum.LastColumnOnPage
        '' for the rightmost column, we turn stretching off:
        'rt.Cols(rt.Cols.Count - 1).Stretch = StretchColumnEnum.No
        '' tell the table that it can split horizontally,
        '' otherwise the right part of the table will be clipped:
        'rt.CanSplitHorz = True
        'If FrmMenu.FontDialog1.ShowDialog() = DialogResult.OK Then
        '    Me.C1PrintDocumentRPT.Style.Font = FrmMenu.FontDialog1.Font
        '    RT.Style.Font = FrmMenu.FontDialog1.Font
        '    RT.RowGroups.Table.Style.Font = FrmMenu.FontDialog1.Font
        '    Me.C1PrintDocumentRPT.Reflow() 'Call AutoSizeTable(RT)
        'End If
        'Me.CMDialog1Open.Filter = ".XML files (*.xml)|*.xml"  '|All files (*.*)|*.*"
        'Me.CMDialog1Open.ShowDialog()

        '120: '           Try
        '                    ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)
        '                    Me.Size = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        '                    'Test Print
        '                    '  sets the height of printed rows to fit all the data:
        '                    'ppv.C1PrintPreviewControl1.Visible = True '11-11-08
        '                    Me.tgQh.PrintInfo.VarRowHeight = C1.Win.C1TrueDBGrid.PrintInfo.RowHeightEnum.StretchToFit

        '                    'The following example translates the grid's color scheme to the print page:
        '                    IfDebugOn Then Sto
        '                    Me.tgQh.PrintInfo.UseGridColors = True
        '                    'Me.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)

        '                    'Wrap text
        '                    'Me.tgln.PrintInfo.WrapText = C1.Win.C1TrueDBGrid.PrintInfo.WrapTextEnum.Wrap
        '                    'The following example sets the style for the page header:

        '                    Dim fntFont As Font
        '                    fntFont = New Font(Me.tgQh.PrintInfo.PageHeaderStyle.Font.Name, Me.tgQh.PrintInfo.PageHeaderStyle.Font.Size, FontStyle.Italic)
        '                    Me.tgQh.PrintInfo.PageHeaderStyle.Font = fntFont
        '                    Me.tgQh.PrintInfo.PageHeader = "Follow Up - Expedite"
        '                    '  sets the pages' margins for printing
        '                    Me.tgQh.PrintInfo.PageSettings.Landscape = True '11-07-08
        '                    Me.tgQh.PrintInfo.PageSettings.Margins.Top = 50
        '                    Me.tgQh.PrintInfo.PageSettings.Margins.Bottom = 25
        '                    Me.tgQh.PrintInfo.PageSettings.Margins.Left = 50
        '                    Me.tgQh.PrintInfo.PageSettings.Margins.Right = 50
        '                    '  previews and prints the grid if it has no data rows:
        '                    Me.tgQh.PrintInfo.PrintEmptyGrid = True
        '                    '  previews and prints horizontal splits in the grid:

        '                    Me.tgQh.PrintInfo.PrintHorizontalSplits = True
        '                    'The following example sets the caption of the Progress form to read "Generating document...":
        '                    Me.tgQh.PrintInfo.ProgressCaption = "Generating document..."
        '                    'The following example prints column footers at the bottom of each page:
        '                    Me.tgQh.PrintInfo.RepeatColumnFooters = True
        '                    ' Column headers will be on every page.
        '                    Me.tgQh.PrintInfo.RepeatColumnHeaders = True
        '                    'The following example prints the grid caption on each page:
        '                    ' Header for the entire grid.
        '                    Me.tgQh.Caption = "Follow Up - Expedite"
        '                    ' Grid caption will be on every page.
        '                    Me.tgQh.PrintInfo.RepeatGridHeader = True
        '                    ' Splits caption will be on every page.
        '                    Me.tgQh.PrintInfo.RepeatSplitHeaders = True
        '                    ' Show Options dialog box.
        '                    Me.tgQh.PrintInfo.ShowOptionsDialog = True
        '                    '  does not show highlighted cells when previewing or printing:
        '                    Me.tgQh.PrintInfo.ShowSelection = True ' False
        '                    'PrintInfo.PageBreaksEnum Enumeration 
        '                    'RT.BreakAfter = BreakEnum.Page '06-01-10 PageBreak
        '                    'ng Me.tgln.PrintInfo.PageBreaksEnum.ClipInArea() ' Clip columns. 
        '                    'ngMe.tgln.PrintInfo.PageBreaksEnum.FitIntoArea() ' Fit all columns in one page. 
        '                    'ngMe.tgln.PrintInfo.PageBreaksEnum.OnColumn() ' Breaks on any column that doesn't fit. 
        '                    'ngMe.tgln.PrintInfo.PageBreaksEnum.OnSplit() ' Breaks on a new split or any column that doesn't fit. 
        '                    ' Me.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)
        '                    'ppv.Show()
        '                    'Stretch = StretchTableEnum.LastColumnOnPage
        '                    Me.tgQh.PrintInfo.FillAreaWidth = True '11-12-08
        '                    '' Invoke print preview.
        '                    Me.tgQh.PrintInfo.FillAreaWidth = C1.Win.C1TrueDBGrid.PrintInfo.FillEmptyEnum.ExtendAll
        '                    IfDebugOn Then Sto 'Get PaperSourceKind Error
        '                    Me.tgQh.PrintInfo.PrintPreview()
        '                    IfDebugOn Then Sto
        '                    Exit Sub

        '                Catch ex As Exception
        '                    MsgBox(ex.Message.ToString)
        '                End Try

PrintFollowReportExit:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow

    End Sub

    Private Sub RbnSaveCurrentGridLayoutSettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbnSaveCurrentGridLayoutSettingsToolStripMenuItem.Click, RbnResetToCurrentGridLayoutToolStripMenuItem.Click, RbnResetToOriginalGridLayoutToolStripMenuItem.Click
        'RbnResetToOriginalGridLayoutToolStripMenuItem_Click
        'Dist & Rep use same files 
        Me.tgQh.FilterBar = False
        Me.tgln.FilterBar = False
        Me.tgr.FilterBar = False
        Dim CurrOrig As String = "Curr" '05-04-10 
        Dim LoadSave As String = "Load"
        If sender = RbnSaveCurrentGridLayoutSettingsToolStripMenuItem Then CurrOrig = "Curr" : LoadSave = "Save" '05-04-10
        If sender = RbnResetToOriginalGridLayoutToolStripMenuItem Then CurrOrig = "Orig" : LoadSave = "Load" '05-04-10
        If sender = RbnResetToCurrentGridLayoutToolStripMenuItem Then CurrOrig = "Curr" : LoadSave = "Load" '05-04-10
        On Error Resume Next                    'Rep
        '01-22-10"VQrtHdrDistTGLayout"=tgQh,"VQrtHdrTGLayout"=tgQh,"VQrtQuoteToTGLayout"=tgr-Realization,"VQrtLinesTGLayout"=tg
        '"Project Shortage","Product Sales History","Realization","Terr Spec Credit Report","Quote Summary","Planned Projects"'If Me.pnlTypeOfRpt.Text.StartsWith("
        If Me.pnlTypeOfRpt.Text.Trim = "" Then Exit Sub 'Sto ' Exit Sub
        If Me.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Then '01-22-10
            If DIST Then
                Call TrueGridLayoutFiles(LoadSave, CurrOrig, "VQrtLinesDistTGLayout", tgln) '05-06-10 If DIST Then
            Else
                Call TrueGridLayoutFiles(LoadSave, CurrOrig, "VQrtLinesTGLayout", tgln) '
            End If
        ElseIf Me.pnlTypeOfRpt.Text.StartsWith("Realization") Then '01-22-10 Quote To "Product Sales History - Line Items"
            '****************************************
            Dim ShowAllQuoteHeader As String = "" '06-06-11
            If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then ShowAllQuoteHeader = "ShowAll" '06-06-11 = "Show All Quote Header Fields" Then '06-06-11 "Add Cust QuoteTo Breakdown to Report"
            If DIST Then '09-04-10 "VQrtRealQTOShowHideDistPrint.xml"
                Call TrueGridLayoutFiles(LoadSave, CurrOrig, "VQrtQuoteToDistTGLayout" & ShowAllQuoteHeader, tgr) '09-15-10 VQrtRealQTOShowHideDistPrint.xml to "VQrtQuoteToDistTGLayout"
            Else
                Call TrueGridLayoutFiles(LoadSave, CurrOrig, "VQrtQuoteToTGLayout" & ShowAllQuoteHeader, tgr) '
            End If
        Else
            If DIST Then '01-19-10
                Call TrueGridLayoutFiles(LoadSave, CurrOrig, "VQrtHdrDistTGLayout", tgQh) '12-18-09 
                '08-01-12 JTC If they load Original then delete the old ShowHide
                If CurrOrig = "Orig" And My.Computer.FileSystem.FileExists(UserDir & "VQrtLineItemsDistShowHide.xml") = True Then
                    My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLineItemsDistShowHide.xml")
                End If
                frmShowHideGrid.Close() '08-01-12 
            Else
                Call TrueGridLayoutFiles(LoadSave, CurrOrig, "VQrtHdrTGLayout", tgQh) '12-18-09 
                '08-01-12 JTC If they load Original then delete the old ShowHide
                If CurrOrig = "Orig" And My.Computer.FileSystem.FileExists(UserDir & "VQrtLineItemsRepShowHide.xml") = True Then
                    My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLineItemsRepShowHide.xml")
                End If
                frmShowHideGrid.Close() '08-01-12 
            End If
        End If

    End Sub


    Private Sub AddTgColumns(ByVal ColName As String, ByVal NearFarCenter As String, ByVal tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        Dim AlignTmp As Integer 'AddColumns ColumnAdd
        Try '02-06-10 Call AddTgColumns(ColName = "Comm-%", NearFarCenter = "Far", tgLookup)
            AlignTmp = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            If NearFarCenter = "Center" Then AlignTmp = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
            If NearFarCenter = "Near" Then AlignTmp = C1.Win.C1TrueDBGrid.AlignHorzEnum.Near
            Dim Temp As String = Me.tgln.Splits(0).DisplayColumns(ColName).Name '02-03-10 If Me.tgLookup.Splits(0).DisplayColumns("Comm-%").Name = "Comm-%" Then
            tg.Splits(0).DisplayColumns(ColName).Visible = True '08-05-13
            'Already Their Don't ADD
        Catch myException As Exception
            'Add Columnn '01-21-10
            Dim Col As New C1.Win.C1TrueDBGrid.C1DataColumn
            Dim dc As C1.Win.C1TrueDBGrid.C1DisplayColumn
            With Me.tgln                    '.Columns.Insert(.Columns.Count - 1, Col)
                .Columns.Add(Col)
                Col.Caption = ColName '"Comm-%"
                
                dc = .Splits(0).DisplayColumns.Item(ColName) ' Move the newly added column to leftmost position in the grid.
                '.Splits(0).DisplayColumns.RemoveAt(.Splits(0).DisplayColumns.IndexOf(dc))'.Splits(0).DisplayColumns.Insert(.Splits(0).DisplayColumns.Count - 1, dc)
                dc.Visible = True
                dc.Style.HorizontalAlignment = AlignTmp 'C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
                dc.HeadingStyle.HorizontalAlignment = AlignTmp 'C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
                If ColName = "FirmName" Then
                    dc.Width = 200
                End If
                Col.DataField = ColName '03-24-14
                .Rebind(True) 'true on rebind perserves column layout
            End With
        End Try
    End Sub
    Public Sub TrueGridLayoutFiles(ByRef SaveLoad As String, ByRef OrigCurr As String, ByRef XMLFileName As String, ByRef tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        'Parameters: SaveLoad = either SAVE or LOAD, OrigCurr = either ORIG or CURR, Base Layout Filename, Tg
        'SaveLoad = either SAVE or LOAD
        'OrigCurr = either ORIG or CURR
        'XMLFileName  = the Base Layout Filename. ex: VFollTGLayout  The routine adds Original.xml or Current.xml
        'tg = C1.Win.C1TrueDBGrid.C1TrueDBGrid Name      ex: FrmMenu.tg
        'If TrueGridError = True Then Exit Sub '10-16-12
        Dim ErrorString As String = ""
        Dim A As String = ""
        Dim ErrorCnt As Integer = 0
Start:
        If tg.Name = "" Then Exit Sub '09-14-11
        'Dim OrigXMLFile As String = XMLFileName & tg.Name.ToUpper & "Original.xml" '09-15-08"
        'Dim CurrXMLFile As String = XMLFileName & tg.Name.ToUpper & "Current.xml" '09-15-08"

        Dim OrigXMLFile As String = XMLFileName & "Original.xml" '09-15-08"
        Dim CurrXMLFile As String = XMLFileName & "Current.xml" '09-15-08"


        Dim AOrig As String = ""
        Try
            '02-18-14 tg.FetchRowStyles = True '11-02-10
            SaveLoad = UCase(SaveLoad) : OrigCurr = UCase(OrigCurr)

            If SaveLoad = "SAVE" Then
                tg.Splits(0).ColumnCaptionHeight = 0 '07-03-13
                Try
                    For Each tDC As C1.Win.C1TrueDBGrid.C1DataColumn In tg.Columns '03-02-11
                        If tDC.FilterText <> "" Then
                            tDC.FilterText = ""
                        End If
                        tDC.FilterClearText = True
                    Next
                Catch ex As Exception
                End Try

                If OrigCurr = "ORIG" Then
                    Dim stm As New System.IO.FileStream(UserDir & OrigXMLFile, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
                    tg.SaveLayout(stm)
                    stm.Close()
                Else
                    Dim stm As New System.IO.FileStream(UserDir & CurrXMLFile, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
                    tg.SaveLayout(stm)
                    stm.Close()
                End If
            End If

            If SaveLoad = "LOAD" Or SaveLoad = "FORMLOAD" Then '09-15-09 
                '    Throw New Exception
                AOrig = UserSysDir & OrigXMLFile '10-21-09- Get from SYS
                Dim fileExists As Boolean = CheckForFile(AOrig, False)
                If fileExists = False Then 'If No Orig Then Save as Orig
                    ErrorString = "Saving Original in SYS"
                    Dim stm As New System.IO.FileStream(AOrig, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
                    tg.SaveLayout(stm)
                    stm.Close()
                    Application.DoEvents()
                End If  'This assumes if You don't have an Original their is no Current

                A = UserDir & CurrXMLFile '08-07-08 Load if they have previously saved
                fileExists = CheckForFile(A, False)
                If fileExists = True Then
                    ErrorString = "Loading User Saved layout"
                    tg.LoadLayout(A)
                    Application.DoEvents()
                    '11-21-11 InitZoom(tg) '07-15-10
                Else
                    Dim stm As New System.IO.FileStream(A, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite) '03-10-11 Save the Original into the user directory to avoid sharing violations when multiple users try to access the file
                    ErrorString = "Saving User From SYS"
                    tg.SaveLayout(stm) : stm.Close() '03-10-11
                    Application.DoEvents()
                    tg.LoadLayout(A) '03-10-11 was Aorig
                End If

            End If

            If SaveLoad = "RESET" Then '03-30-11
                'LOAD ORIGINAL AND THEN SAVE IT TO THE CURRENT
                AOrig = UserSysDir & OrigXMLFile
                Dim fileExists As Boolean = CheckForFile(AOrig, False)
                If fileExists = False Then MessageBox.Show("Couldn't find Layout File.  Exit out of the system and come back in") : Exit Sub
                tg.LoadLayout(AOrig)
                Dim stm As New System.IO.FileStream(UserDir & CurrXMLFile, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
                tg.SaveLayout(stm) : stm.Close() '03-10-11
            End If

            If tg.Name = "tgQh" Then tg.Columns("EntryDate").EnableDateTimeEditor = False : tg.Columns("BidDate").EnableDateTimeEditor = False
            If tg.Name = "tgln" Then
                tg.Columns("Comm").NumberFormat = "n2"
                tg.RowHeight = 0 '12-01-14
            End If

            tg.FilterBar = False '11-10-11
            If OrigXMLFile.Contains("VQrtHdrTGLayout") Then
                tg.Splits(0).DisplayColumns("Comm-%").FetchStyle = True
            End If

        Catch ex As Exception
            If ex.Message.StartsWith("The Process cannot access the file") Then
            Else 'If ex.Message.StartsWith("Exception has been thrown by the target of an invocation") Then

                'Else
                'DftErrMsg:      MessageBox.Show("Error in TrueGridLayoutFiles " & tg.Name & " " & ErrorString & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VORDER " & ProgramDate, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try

    End Sub
    'Private Sub TrueGridLayoutFiles(ByVal SaveLoad As String, ByVal OrigCurr As String, ByVal XMLFileName As String, ByVal tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
    '    'Parameters: SaveLoad = either SAVE or LOAD, OrigCurr = either ORIG or CURR, Base Layout Filename, Tg
    '    'SaveLoad = either SAVE or LOAD
    '    'OrigCurr = either ORIG or CURR                   Rep             Dist
    '    'XMLFileName  = the Base Layout Filename. ex: VQrtHdrTGLayout VQrtHdrDistTGLayout (The routine adds Original.xml or Current.xml
    '    'tg = C1.Win.C1TrueDBGrid.C1TrueDBGrid Name      ex: FrmFoll.tg

    '    Dim OrigXMLFile As String = XMLFileName & "Original.xml" '09-15-08"
    '    Dim CurrXMLFile As String = XMLFileName & "Current.xml" '09-15-08"
    '    Try
    '        'Debug.Print(OrigXMLFile)
    '        'QUOTE SUMMARY REPORT
    '        '   VQRTHDRTGLayoutOriginal.xml
    '        '   VQrtShowHideRepPrtHdr.xml

    '        'VQrtQuoteToTGLayoutOriginal.xml
    '        'VQrtLinesTGLayoutOriginal.xml

    '        SaveLoad = UCase(SaveLoad) : OrigCurr = UCase(OrigCurr)
    '        Dim A As String = ""
    '        'tg.FetchRowStyles = True '08-02-13
    '        'For I As Integer = 0 To tg.Columns.Count - 1 '07-08-13
    '        '    tg.Columns(I).Tag = I
    '        '    tg.Splits(0).DisplayColumns(I).FetchStyle = True '09-11-10
    '        'Next
    '        '01-24-12 If not in UserDir & CurrXMLFile copy from UserSysDir Orig to UserDir & CurrXMLFile
    '        If My.Computer.FileSystem.FileExists(UserDir & CurrXMLFile) = False Then '02-06-11
    '            If My.Computer.FileSystem.FileExists(UserSysDir & OrigXMLFile) = True Then '01-25-12
    '                My.Computer.FileSystem.CopyFile(UserSysDir & OrigXMLFile, UserDir & CurrXMLFile, True) '02-06-11 True = Overwrite Overwrite no error)
    '            End If                                     '01-25-12 Chg to Orig
    '        End If
    '        If SaveLoad = "SAVE" Then 'Use for Development Only

    '            If OrigCurr = "ORIG" Then  'OrigXMLFile'12-23-10 to UserSysDir
    '                Dim stm As New System.IO.FileStream(UserSysDir & OrigXMLFile, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
    '                tg.SaveLayout(stm)
    '                stm.Close()
    '            Else    'So Save CurrXMLFile         
    '                Dim stm As New System.IO.FileStream(UserDir & CurrXMLFile, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
    '                tg.SaveLayout(stm)
    '                stm.Close()
    '            End If
    '            '02-08-09 NO Call frmShowHideGrid.ShowHideGridCol("Reset", UserDir & "ShowHideGridQrt.xml") 'Moved Columns Around
    '        End If
    '        If SaveLoad = "LOAD" Then 'Load - Read 
    '            If OrigCurr = "ORIG" Then
    '                '12-23-10 to UserSysDir
    '                A = UserSysDir & OrigXMLFile
    '                'Dim fileExists As Boolean = CheckForFile(A, False)
    '                If My.Computer.FileSystem.FileExists(A) Then '06-20-10 If fileExists = True Then
    '                    tg.LoadLayout(A)
    '                End If
    '            Else
    '                'Load Current if they have previously saved     
    '                A = UserDir & CurrXMLFile
    '                'Dim fileExists As Boolean = CheckForFile(A, False)
    '                If My.Computer.FileSystem.FileExists(A) Then '06-20-10 If fileExists = True Then
    '                    tg.LoadLayout(A)
    '                    'If tg.Name = "tgQh" And DIST = False And tg.Splits(0).DisplayColumns("LPCost").Visible = True Then tg.Splits(0).DisplayColumns("LPCost").Visible = False '07-27-12     
    '                    'If tg.Name = "tgQh" And DIST = False Then tg.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" '06-24-13 And tg.Splits(0).DisplayColumns("Cost").Visible = True
    '                    If tg.Name = "tgln" Then '08-06-13
    '                        '03-12-14
    '                        'tg.Splits(0).DisplayColumns("LineID").Visible = False '08-06-13
    '                        'tg.Splits(0).DisplayColumns("LPProdID").Visible = False '06-24-13
    '                        'tg.Splits(0).DisplayColumns("QuoteID").Visible = False '06-24-13
    '                        'tg.Splits(0).DisplayColumns("ProdID").Visible = False '06-24-13
    '                        'Call AddTgColumns("Ext Sell", "Far", tgln) '08-06-13
    '                        '03-12-14

    '                        '03-12-14 - fix later
    '                        'If Me.chkUseSpecifierCode.Checked = True Or Me.chkShowCustomers.Checked = True Then '08-02-13 
    '                        '    If InStrColNam("FIRMNAME") Then Call AddTgColumns("FirmName", "Near", tgln) '08-02-13'UseFirmName = True Else Call TgColumnsVisible("FirmName", "False", tgInvoiceMaster)
    '                        '    If InStrColNam("NCODE") Then Call AddTgColumns("NCode", "Near", tgln) '08-05-13
    '                        'End If

    '                    End If
    '                    If tg.Name = "tgln" And DIST = False Then
    '                        '03-12-14tg.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" '06-24-13
    '                        '03-12-14tg.Splits(0).DisplayColumns("LPCost").Visible = False '06-24-13
    '                        '03-12-14Call AddTgColumns("Ext Comm", "Far", tgln) '08-06-13
    '                        'Moved Uptg.Splits(0).DisplayColumns("LPProdID").Visible = False '06-24-13
    '                        'tg.Splits(0).DisplayColumns("ProdID").Visible = False '06-24-13
    '                        'tg.Splits(0).DisplayColumns("ProdID").Visible = False '06-24-13
    '                        '08-06-13 tg.Splits(0).DisplayColumns("Ext Cost").Visible = False '06-24-13
    '                        'If Me.chkUseSpecifierCode.Checked = True Or Me.chkShowCustomers.Checked = True Then '08-02-13 
    '                        '    If InStrColNam("FIRMNAME") Then Call AddTgColumns("FirmName", "Near", tgln) '08-02-13'UseFirmName = True Else Call TgColumnsVisible("FirmName", "False", tgInvoiceMaster)
    '                        '    If InStrColNam("NCODE") Then Call AddTgColumns("NCode", "Near", tgln) '08-05-13
    '                        'End If
    '                    End If
    '                    '06-24-13If tg.Name = "tgln" And DIST = False Then tg.Splits(0).DisplayColumns("BkSell").Visible = False '06-24-13
    '                    '03-12-14If tg.Name = "tgr" And DIST = False And tg.Splits(0).DisplayColumns("Cost").Visible = False Then tg.Splits(0).DisplayColumns("Cost").Visible = True : tg.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" '11-28-12 JTC Cost to Book show Book

    '                Else
    '                    MsgBox("You must save the Current Grid Settings First" & vbCrLf & A) '02-19-20
    '                End If
    '            End If
    '        End If
    '        If SaveLoad = "FORMLOAD" Then '09-15-09 
    '            '12-23-10 to UserSysDir
    '            A = UserSysDir & OrigXMLFile '08-07-08 Save an Original Setting if not done
    '            'Dim fileExists As Boolean = CheckForFile(A, False)
    '            If My.Computer.FileSystem.FileExists(A) = False Then '06-20-10 If fileExists = False Then 'If No Orig Then Save as Orig
    '                'Call TrueGridLayoutFiles("Save", "Orig") 
    '                Dim stm As New System.IO.FileStream(A, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite) '10-28-08
    '                '01-31-12 Me.tgln.SaveLayout(stm) '12-18-09
    '                tg.SaveLayout(stm) '01-31-12
    '                stm.Close()
    '            End If  'This assumes if You don't have an Original their is no Current
    '            A = UserDir & CurrXMLFile '08-07-08 Load if they have previously saved
    '            If My.Computer.FileSystem.FileExists(A) Then '09-07-10 fileExists = CheckForFile(A, False)
    '                '09-07-10 If fileExists = True Then
    '                '01-31-12 Me.tgln.LoadLayout(A) 'UserDir & "\" & "VQrtHdrTGLayoutOriginal.xml")
    '                tg.LoadLayout(A) '01-31-12 
    '                '03-12-14If tg.Name = "tgQh" And DIST = False And tg.Splits(0).DisplayColumns("LPCost").Visible = True Then tg.Splits(0).DisplayColumns("LPCost").Visible = False '06-25-13
    '                '03-12-14If tg.Name = "tgQh" And DIST = False Then tg.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" '06-24-13 And tg.Splits(0).DisplayColumns("Cost").Visible = True
    '                '03-12-14If tg.Name = "tgln" And DIST = False Then
    '                '03-12-14
    '                'tg.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" '06-25-13
    '                'tg.Splits(0).DisplayColumns("LPCost").Visible = False '06-24-13
    '                'tg.Splits(0).DisplayColumns("LPProdID").Visible = False '06-24-13
    '                'tg.Splits(0).DisplayColumns("ProdID").Visible = False '06-24-13
    '                'tg.Splits(0).DisplayColumns("ProdID").Visible = False '06-24-13
    '                '08-06-13 tg.Splits(0).DisplayColumns("Ext Cost").Visible = False '06-24-13
    '                '03-12-14End If
    '            End If
    '        End If

    '        If tg.Name = "tgQh" Then tg.Columns("EntryDate").EnableDateTimeEditor = False : tg.Columns("BidDate").EnableDateTimeEditor = False
    '        tg.FilterBar = False '11-10-11

    '        If OrigXMLFile.Contains("VQrtHdrTGLayout") Then
    '            tg.Splits(0).DisplayColumns("Comm-%").FetchStyle = True
    '        End If

    '    Catch ex As Exception
    '        If OrigCurr = "ORIG" Then '07-15-13
    '            MsgBox(ex.Message.ToString & vbCrLf & UserDir & OrigXMLFile & vbCrLf & "If the problem persists call Multimicro for support", , "VQRT TruGridLayoutFiles") '07-15-13
    '        Else
    '            '07-15-13 MessageBox.Show("Error in TruGridLayoutFiles VQRT" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12 
    '            MsgBox(ex.Message.ToString & vbCrLf & UserDir & CurrXMLFile & vbCrLf & "If the problem persists call Multimicro for support", , "VQRT TruGridLayoutFiles") '07-15-13
    '        End If
    '    End Try ''Step thru below on New program to Build the first time
    '    'Call frmShowHideGrid.ShowHideGridCol("AllOn", UserDir & "ShowHideGridQrt.xml") '12-07-08 ByVal ShowHide As String)
    'End Sub

    Private Sub cboQuoteRptPrt_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboQuoteRptPrt.SelectedIndexChanged
        Dim index As Short = cboQuoteRptPrt.GetIndex(eventSender)
        On Error Resume Next '02-07-00 WNA
    End Sub
    Private Sub cboSpecCross_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbospeccross.Enter '02-13-13
        'Private Sub cboSpecCross_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSpecCross.SelectedIndexChanged
        On Error Resume Next
        Me.cbospeccross.SelectionStart = 0
        Me.cbospeccross.SelectionLength = Len(Me.cbospeccross.Text)
    End Sub
    Private Sub cboSpecCross_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cbospeccross.KeyPress '02-13-13
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error Resume Next
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    'Private Sub cboSpecCross_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSpecCross.SelectedIndexChanged
    '    Dim Cancel As Boolean = System.EventArgs.Cancel
    '    Dim Resp As Short
    '    On Error Resume Next
    '    If Trim(Me.cboSpecCross.Text) <> "ALL" And Trim(Me.cboSpecCross.Text) <> "Spec" And Trim(Me.cboSpecCross.Text) <> "Cross" Then
    '        Resp = MsgBox("Invalid entry - Select an option from the drop down menu", MsgBoxStyle.Information, "Select Spec/Cross")
    '        Me.cboSpecCross.Text = "ALL"
    '        Cancel = True
    '    End If
    '    System.EventArgs.Cancel = Cancel
    'End Sub
    Private Sub cboSpecCross_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cbospeccross.Validating, cbospeccross.LostFocus '02-13-13
        'Private Sub cboSpecCross_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSpecCross.SelectedIndexChanged
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim Resp As Short
        On Error Resume Next
        Me.cbospeccross.Text = Me.cbospeccross.Text.ToUpper '02-13-13
        If Trim(Me.cbospeccross.Text) <> "ALL" Then '02-13-13 And Trim(Me.cbospeccross.Text) <> "Spec" And Trim(Me.cbospeccross.Text) <> "Cross" Then
            'Resp = MsgBox("Invalid entry - Select an option from the drop down menu", MsgBoxStyle.Information, "Select Spec/Cross")
            If Len(Me.cbospeccross.Text.Trim) > 1 Then '02-13-13 
                Me.cbospeccross.Text = "ALL"
                Cancel = True
            End If
        End If
        eventArgs.Cancel = Cancel
    End Sub
    'Private Sub cboStockJob_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    '    On Error Resume Next
    '    Me.cboStockJob.SelectionStart = 0
    '    Me.cboStockJob.SelectionLength = Len(Me.cboStockJob.Text)
    'End Sub

    'Private Sub cboStockJob_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboStockJob.KeyPress
    '    Dim KeyAscii As Short = Asc(EventArgs.KeyChar)
    '    If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
    '    EventArgs.KeyChar = Chr(KeyAscii)
    '    If KeyAscii = 0 Then
    '        EventArgs.Handled = True
    '    End If
    'End Sub


    Private Sub cboStockJob_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cboStockJob.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim Resp As Short
        On Error Resume Next
        If Trim(Me.cboStockJob.Text) <> "ALL" And Trim(Me.cboStockJob.Text) <> "Stock" And Trim(Me.cboStockJob.Text) <> "Job" Then
            Resp = MsgBox("Invalid entry - Select an option from the drop down menu", MsgBoxStyle.Information, "Select Stock/Job")
            Me.cboStockJob.Text = "ALL"
            Cancel = True
        End If
        eventArgs.Cancel = Cancel
    End Sub
    'Private Sub chkBlankBidDates_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkBlankBidDates.CheckStateChanged
    Private Sub chkBlankBidDates_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBlankBidDates.CheckedChanged

        'If this is checked then on a report if a bid date range is entered
        'all quotes with blank bid dates will be on the report.
        '10-08-15 JTC Added New Code to this SUB
        If chkBlankBidDates.CheckState = CheckState.Checked Then
            'If Me.ChkCheckBidDates.Text = "Check Deliver Dates when Selecting Quotes" ThenStop
            Me.ChkCheckBidDates.CheckState = CheckState.Checked
            'Me.chkBidJobsOnly.CheckState = CheckState.Checked '= True ' CheckState =
            Me.DTPicker1StartBid.Enabled = True '02-04-12
            Me.DTPicker1EndBid.Enabled = True '02-04-12
            'ChkIgnoreBidDates.Text = "Use Bid Dates when Selecting Quotes"
            'End If
        Else
            '06-05-12 Me.DTPicker1StartBid.Value = CDate("01/01/1900")
            'chkBlankBidDates.CheckState = CheckState.Unchecked
            'chkBlankBidDates.Visible = False ' CheckState =
            'Me.DTPicker1StartBid.Enabled = False '02-04-12
            'Me.DTPicker1EndBid.Enabled = False '02-04-12 
        End If

    End Sub
    'Private Sub chkCustomerBreakdown_Click(ByRef Value As Short) Handles chkCustomerBreakdown.Enter
    '    '07-07-09
    '    Dim CheckState As Object
    '    'rptl% = 5
    '    If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then
    '        RptCust = 1 '12-06-02 WNA
    '        If RptMFG = 1 Then RptMFGCust = 1 '12-06-02 WNA Else RptCust% = 1
    '    Else
    '        RptCust = 0
    '        RptMFGCust = 0
    '    End If
    'nd Sub

    Private Sub cboSortPrimarySeq_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboLotUnit.KeyPress, cboStockJob.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    'Private Sub tgColumnsAdd(ByVal ColName As String, ByVal NearFarCenter As String, ByVal tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
    '    'SomeTimes it Passes wrong tgLookup Grid Call tgColumnsAdd("Total Amt", "Far", tgInvoiceMaster) '06-10-10 .Width = 1.5
    '    Dim AlignTmp As Integer 'AddColumns, tgColumnsAdd, tgColumnsDelete(ColName = "Comm-%", NearFarCenter = "Far", tgLookup)
    '    Try '02-05-10 Call tgColumnsDelete(ColName = "Comm-%", NearFarCenter = "Far", tgLookup)
    '        AlignTmp = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
    '        If NearFarCenter = "Center" Then AlignTmp = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
    '        If NearFarCenter = "Near" Then AlignTmp = C1.Win.C1TrueDBGrid.AlignHorzEnum.Near
    '        Dim Temp As String = tg.Splits(0).DisplayColumns(ColName).Name '02-03-10 If Me.tgLookup.Splits(0).DisplayColumns("Comm-%").Name = "Comm-%" Then
    '        tg.Splits(0).DisplayColumns(ColName).Visible = True '04-21-10 
    '        'Debug.Print(Me.tg.Splits(0).DisplayColumns(ColName).Width)
    '        'Already Their Don't ADD
    '    Catch myException As Exception
    '        'Add Columnn '01-21-10    Goes Here if Col Does Not Exist
    '        Dim Col As New C1.Win.C1TrueDBGrid.C1DataColumn
    '        Dim dc As C1.Win.C1TrueDBGrid.C1DisplayColumn
    '        With tg                    '.Columns.Insert(.Columns.Count - 1, Col)
    '            .Columns.Add(Col)
    '            Col.Caption = ColName '"Comm-%"
    '            dc = .Splits(0).DisplayColumns.Item(ColName) ' Move the newly added column to leftmost position in the grid.
    '            '.Splits(0).DisplayColumns.RemoveAt(.Splits(0).DisplayColumns.IndexOf(dc))'.Splits(0).DisplayColumns.Insert(.Splits(0).DisplayColumns.Count - 1, dc)
    '            dc.Visible = True
    '            Col.DataField = ColName '"Comm-%"'04-21-10 
    '            dc.Style.HorizontalAlignment = AlignTmp 'C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
    '            dc.HeadingStyle.HorizontalAlignment = AlignTmp 'C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
    '            .Rebind(True) 'true on rebind perserves column layout
    '        End With
    '    End Try
    'End Sub

    Public Sub TgColumnsVisible(ByVal ColName As String, ByRef TrueFalse As String, ByRef tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid) '10-13-11 was byval
        Try '02-21-10 Call tgColumnsVisible("Comm-%", "True", tgLookup)'02-23-10          tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid
            tg.Splits(0).DisplayColumns(ColName).Visible = TrueFalse '03-29-11 From True to TrueFalse 
        Catch myException As Exception
            TrueFalse = "NoCol" '02-08-12
            'IfDebugOn Then MsgBox("Visible Error " & ColName) '04-12-10 
            'Do Nothing It doesn't exist
        End Try
    End Sub
    Public Function InStrColNam(ByVal FileName As String) As Boolean
        Dim NameExists As Boolean = False '09-01-11
        If InStr(strSql.ToUpper, FileName) Then NameExists = True : GoTo getout
        If InStr(strSql.ToUpper, FileName) = True Then NameExists = True : GoTo getout
GetOut:
        Return NameExists
    End Function
    Private Sub cboSortSecondarySeq_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    'Private Sub cmdAbort_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAbort.Click
    '    '07-07-09
    '    AbortPrtFlag = True
    '    'prtAbort.Visible = False '06-29-96
    'End Sub
    Private Sub cmdCancel1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel1.Click, cmdCancel2.Click '04-20-11
        On Error Resume Next
        SESCO = False '02-25-12
        Me.lblJobName.Text = "Job Name Search String" '06-22-12 
        ' Me.pnlTypeOfRpt.Text = "" '02-25-12
        'Me.ChkSpecifiers.Text = "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
        '02-11-12 Use ChkSpecifiers.Text = "Sort Report by Descending Dollar 
        Me.ChkSpecifiers.Text = "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 "Sort Report by Descending Dollar" '02-11-12 " "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
        'Me.ChkSpecifiers.Visible = True '02-11-12 
        'Me.fraRptSel.Visible = False
        If Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then Me.fraSortSecondarySeq.Visible = False '09-21-12
        Me.fraSortPrimarySeq.Visible = True
        'Me.pnlSecondarySort.Visible = False
        'Me.txtSecondarySort.Visible = False
        Me.tabQrt.SelectedIndex = 0
        Call tabQRT_TabActivate(0)
    End Sub

    Private Sub cmdOK1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdok1.Click, cmdOK2.Click '04-20-11
        'Ok Run Report'
        Me.cmdok1.Enabled = False '10-13-14 JTC Fix Dbl Click on Reports
        SortCode = "" '06-02-15 JTC
        ForecastAllMfg = False '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
        If BrandReportMfg = "PHIL" Or BrandReportMfg = "DAYB" Or BrandReportMfg = "DAY" Or SESCO = True Then
            ForecastAllMfg = False '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
        Else 'Not PHIL Not SESCO
            If Me.txtPrimarySortSeq.Text = "Forecasting" Then '05-28-15 JTC
                ForecastAllMfg = True '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
            Else
                ForecastAllMfg = False
            End If
        End If
        System.Windows.Forms.Application.DoEvents() '10-14-14 JTC
        If Me.pnlTypeOfRpt.Text.Trim = "" Then
            MsgBox("Go to the Left Tab to start Reports Process. Select a Type of Report First")
            Me.tabQrt.SelectedIndex = 0 : GoTo EndExit '10-13-14 JTC Exit Sub '10-12-10 
        End If
        '02-03-14 JTC Can't Select SLSsplit and use SLS from Header
        If Me.txtSlsSplit.Text.ToUpper <> "ALL" Then If Me.chkSlsFromHeader.CheckState = CheckState.Checked Then Me.chkSlsFromHeader.CheckState = CheckState.Unchecked '02-03-14
        US = "" '08-07-13 
        MaxNameLength = Val(Me.rbnMaxNameTxt.Text) '12-23-12 45   Public As Int16
        MaxJobLength = Val(Me.rbnMaxJobTxt.Text) '12-23-12 40   Public As Int16
        If MaxNameLength < 10 Or MaxNameLength > 45 Then MaxNameLength = 45 : Me.rbnMaxNameTxt.Text = "45" '12-23-12
        If MaxJobLength < 10 Or MaxJobLength > 40 Then MaxJobLength = 40 : Me.rbnMaxJobTxt.Text = "40" '12-23-12
        '01-02-12 DecFormat As String = "########0.00" '01-01-12 NoCents "########0") DecFormat for Nocents Whole Dollars
        If chkWholeDollars.Checked = True Then DecFormat = "########0" Else DecFormat = "########0.00" '01-02-12 DecFormat As String = "########0.00" '01-01-12 NoCents "########0") DecFormat for Nocents Whole Dollars
        If chkAddCommas.Checked = True Then
            If chkWholeDollars.Checked = True Then DecFormat = "###,###,##0" Else DecFormat = "###,###,##0.00" '01-02-12 
        End If
        '02-05-12 Specifiers with an Influence % Greater Than Zero
        BranchReporting = False '10-30-12 
        If Me.pnlTypeOfRpt.Text = "" Then Me.pnlTypeOfRpt.Text = "Quote Summary" '02-25-12 "Realization") Then '01-22-10 Quote To "Product Sales History - Line Items"
        Dim OnlyInFluenceGtZero As Boolean = False ''02-05-12 Specifiers with an Influence % Greater Than Zero
        Dim SavePath As String
        SavePath = CurDir() '07-20-05 JH 'Printing to Adobe Changes the Current Directory
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Me.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Then GoTo QutLineHistoryRpt '08-31-09
        Call SetSelectionHeader() 'Set up selection Criteria for report header
        'If Me.optOutputOptions(1).CheckState = CheckState.Checked Then Call CopytoExcel() 'Open excel file
        If AbortPrtFlag = True Then GoTo EndExit '10-13-14 JTC  Exit Sub
        Dim RowCnt As Single = 0 '02-26-09
        If chkBranchReport.CheckState = CheckState.Checked Then '06-15-10
            BranchReporting = True ' As Boolean '05-31-07 JH  Public BranchCodeRpt As String '06-15-10
            BranchCodeRpt = InputBox("Enter One or More Branch Codes. (NOR) or" & vbCrLf & "NOR,WST,EST", "Branch Code", "ALL") '10-30-12
            BranchCodeRpt = BranchCodeRpt.ToUpper
            'If BranchReporting Then If InStr("," & Trim(BranchCodeRpt) & ",", "," & Trim(BranchCode) & ",") = 0 Then Hit = 0 : GoTo GetNext_5540 '06-07-10     'End If
        End If
        If SESCO = True And Me.pnlTypeOfRpt.Text = "Realization" Then '02-26-12 And My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") Then '02-22-12
            RealCustomer = True : RealArchitect = True : RealEngineer = True : RealLtgDesigner = True : RealSpecifier = True : RealContractor = True '02-22-12
            Me.chkCustomerBreakdown.CheckState = CheckState.Checked
        End If
        If Me.chkBidJobsOnly.Text = "Delivery Date Jobs Only" And Me.txtPrimarySortSeq.Text = "Forecasting" Then '01-27-14 JTC Forecast Zero records  
            If BrandReport = False Then '05-16-13
                Me.ChkCheckBidDates.CheckState = CheckState.Unchecked '02-24-14 
                Me.ChkCheckBidDates.Text = "Check Bid Dates when Selecting Quotes" '02-24-14 
                MsgBox("Forecasting options are not set correctly." & vbCrLf & "Use Cancel<Back and Reset Report Options") : GoTo EndExit '10-13-14 JTC Exit Sub '02-24-14 
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''
        If Me.pnlTypeOfRpt.Text = "Realization" Then 'not Quote Summary"
            'If All true
            If RealCustomer = True Or RealManufacturer = True Or RealOther = True Or RealArchitect = True Or RealEngineer = True Or RealLtgDesigner = True Or RealSpecifier = True Or RealContractor = True Or RealOther = True Then '02-04-12
                'Check Dollar Range Amount
                'If Trim(Me.txtStartQuoteAmt.Text) <> "" Or Trim(Me.txtStartQuoteAmt.Text) <> "0" Or Trim(Me.txtEndQuoteAmt.Text) <> "" Or Trim(Me.txtEndQuoteAmt.Text) <> "999999999" Then '"999999999" '01-06-14 JTC Added 9 "999,999,999"
                If Trim(Me.txtStartQuoteAmt.Text) <> "0" Or Trim(Me.txtEndQuoteAmt.Text) <> "999999999" Then '"999999999" '01-06-14 JTC Added 9 "999,999,999"
                    Resp = MsgBox("Yes = Use Quote Header Sell Amoumt to Select Quotes or " & vbCrLf & "No = Use Realization QuoteTo Sell Amount to Select Quotes", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Realization Amount") '01-06-13
                    If Resp = vbYes Then
                        RealQuoteToAmtON = False
                    Else 'Public RealQuoteToAmtON As Boolean = 0 '01-06-14
                        RealQuoteToAmtON = True '01-06-14 JTC RealQuoteToAmtON = True  Use Realization QuoteTo Sell Amount Select Quotes
                    End If
                Else 'No selection so default is use Quote Header Amt 
                    RealQuoteToAmtON = False
                End If
                'Me.ChkExtendByProb.Text = "Extend By Quote Probability" '02-04-12
                'Me.ChkExtendByProb.Text = "Extend Specifiers By Influence %" '02-04-12
                If Me.ChkExtendByProb.CheckState = CheckState.Checked Then '02
                    RealExtByInfluencePercent = True '02-05-12 RealExtByInfluencePercent
                    'Resp = MsgBox("Do You only want Specifiers with an Influence % Greater Than Zero?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, "Influence %") '02-04-12 
                    Resp = MsgBox("If Influence Percent is zero we will use 100%", MsgBoxStyle.Information, "Realization Select") '07-06-12
                    '07-06-12 If Resp = vbYes Then RealExtByInfluencePercent
                    'Dim OnlyInFluenceGtZero As Boolean = False ''02-05-12 Specifiers with an Influence % Greater Than Zero
                    'If Resp = vbYes Then OnlyInFluenceGtZero = True Else OnlyInFluenceGtZero = False '02-04-12'02-04-12
                    'If Resp = vbCancel Then Exit Sub
                Else
                    RealExtByInfluencePercent = False '02-04-12
                End If
            End If 'End if 'If All true
            OrderBy = "" 'P'Q.MarketSegment"
            Select Case txtPrimarySortSeq.Text 'sender.ToString
                Case "Customer"
                    SortSeq = "projectcust.Ncode"
                Case "Manufacturer"
                    SortSeq = "projectcust.Ncode"
                Case "Salesman/Customer"
                    SortSeq = "projectcust.SLSCode, projectcust.NCode" '02-25-09 
                Case "Salesman" '05-16-13 JTC Added Realization When sub Sort is salesman they can change tobe Salesman Major Sequence
                    If Me.txtSecondarySort.Text = "Job Name" Then
                        SortSeq = "projectcust.SLSCode, projectcust.Jobname " '01-21-14
                    Else
                        SortSeq = "projectcust.SLSCode, projectcust.NCode" '05-16-13
                    End If
                Case "Architect"
                    SortSeq = "projectcust.Ncode"
                Case "Engineer"
                    SortSeq = "projectcust.Ncode"
                Case "Specifier"
                    SortSeq = "projectcust.Ncode"
                Case "Other"
                    SortSeq = "projectcust.NCode"
                Case "All"
                    SortSeq = "projectcust.Ncode"
                Case "Bid Date"
                    SortSeq = "projectcust.BidDate" '04-28-15 
                Case Else
                    SortSeq = "projectcust.NCode" 'Name Code

            End Select
            If SESCO = True Then SortSeq = "Quote.Jobname" '02-22-12
            ''Me.ChkSpecifiers.Text = "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
            ''02-11-12 Use ChkSpecifiers.Text = "Sort Report by Descending Dollar 
            'If Me.ChkSpecifiers.Text = "Sort Report by Descending Dollar" And Me.ChkSpecifiers.CheckState = CheckState.Checked Then '02-11-12 " "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
            '    '"Yes = Sort by Name Code / Descending Sales Dollars or No = Just Descending Sales Dollars"
            '    Resp = MsgBox("Yes = Sort by Name Code / Descending Sales Dollars or " & vbCrLf & "No = Just Descending Sales Dollars", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Realization Sort") '02-11-12
            '    If Resp = vbYes Then
            '        SortSeq = "projectcust.Sell, projectcust.NCode" ''02-11-12"
            '    Else
            '        SortSeq = "projectcust.Sell" ''02-11-12"
            '    End If
            'End If
            OrderBy = SortSeq
            If BranchReporting = True Then '10-30-12
                If OrderBy <> "" Then OrderBy = "ORDER BY " & " projectcust.BranchCode, " & OrderBy '10-30-12
            Else
                If OrderBy <> "" Then OrderBy = "ORDER BY " & OrderBy
            End If
            strSql = ""
            Dim MyDefaultsSQL As String = ""
            Dim MyQuotesSQL As String = ""
            Dim MySLSSQL As String = ""
            Try
                'SORT ORDER CODE
            Catch myException As Exception
                MsgBox(myException.Message & vbCrLf & "Print Task" & vbCrLf)
                ' IfDebugOn ThenStop 'CatchStop
            End Try
            If Me.pnlTypeOfRpt.Text = "Project Shortage Report" Then '03-15-10 Then '10-23-02 WNA
                FileClose(4) '
                If My.Computer.FileSystem.FileExists(UserDocDir & "TEMPOSR.DAT") Then Kill(UserDocDir & "TEMPOSR.DAT")
                FileOpen(4, UserDocDir & "TEMPOSR.DAT", OpenMode.Output)
                'If Me.optOutputOptions(1).Checked = CheckState.Checked Then ' Excel
                '    FileClose(2)
                '    If My.Computer.FileSystem.FileExists(UserDocDir & "TEMPEXC.TXT") Then Kill(UserDocDir & "TEMPEXC.TXT")
                '    FileOpen(2, UserDocDir & "TEMPEXC.TXT", OpenMode.Output)
                'End If
            End If
            '02-09-10
            'If DIST = False Then 'IE REP Realization
            '    For I As Int16 = 0 To Me.tgr.Splits(0).DisplayColumns.Count - 1
            '        If Me.tgr.Splits(0).DisplayColumns(I).Name = "LPMarg" Then
            '            '09-05-10 Me.tgr.Splits(0).DisplayColumns("LPMarg").DataColumn.Caption = "LPComm" ' tgQh = Quote Tabl
            '            Me.tgr.Splits(0).DisplayColumns("LPMarg").Visible = False '09-05-10 DataColumn.Caption = "LPComm" ' tgQh = Quote Tabl
            '        ElseIf Me.tgr.Splits(0).DisplayColumns(I).Name = "Margin" Then
            '            Me.tgr.Splits(0).DisplayColumns("Margin").DataColumn.Caption = "Comm-%" '01-22-10"
            '        ElseIf Me.tgr.Splits(0).DisplayColumns(I).Name = "LPCost" Then
            '            Me.tgr.Splits(0).DisplayColumns("LPCost").Visible = False ' DataColumn.Caption = "Comm-%" '01-22-10"
            '        End If
            '    Next
            'End If
            If Me.pnlTypeOfRpt.Text = "Realization" And (Me.txtSecondarySort.Text = "Spread Sheet by Month" Or Me.txtSecondarySort.Text = "Spread Sheet by Year") Then '06-22-15 JTC 05-17-13 
                '06-30-15 JTC Test Start and End Year Test 
                Dim StartYear As Date = VB6.Format(Me.DTPickerStartEntry.Text, "yyyy/MM")
                Dim EndYear As Date = VB6.Format(Me.DTPicker1EndEntry.Text, "yyyy/MM") 'StartYear.AddYears(1) ' VB6.Format(frmQuoteRpt.DTPicker1EndEntry.Text, "yyyy")
                If Me.pnlTypeOfRpt.Text = "Realization" And Me.txtSecondarySort.Text = "Spread Sheet by Month" Then
                    If Format(EndYear, "yyyy") <> Format(StartYear, "yyyy") Then 'Format(StartYear.AddYears(I), "yyyy")
                        MsgBox("Start Year and End Year must be Equal for the twelve month spreadsheet Report! ****** Fix Dates", MsgBoxStyle.Critical + MsgBoxStyle.RetryCancel, "Spread Sheet by Month") : GoTo EndExit
                    End If
                ElseIf Me.pnlTypeOfRpt.Text = "Realization" And Me.txtSecondarySort.Text = "Spread Sheet by Year" Then
                    If Format(EndYear, "yyyy") <> Format(StartYear.AddYears(3), "yyyy") Then
                        MsgBox("End Year must be 4 years after Start Year for the 4 Year Spread Sheet by Year Report! ***** Fix Dates", MsgBoxStyle.Critical + MsgBoxStyle.RetryCancel, "Spread Sheet by Year Report") : GoTo EndExit
                    End If
                End If
            End If
            Call FillQutRealLUDataSet(SortSeq, SortDir, OnlyInFluenceGtZero) '02-05-12 
            'Debug.Print(SortSeq) 'Wrong 
            RowCnt = dsQuoteRealLU.QuoteRealLU.Rows.Count - 1 '05-17-11
            'If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then '01-21-14
            '    RealALL = True '01-21-14 JTC Select Specifiers
            'End If
            If RowCnt = -1 Then RowCnt = 0 : GoTo RowCntZero '05-17-11
            If SESCO = True Or ExcelQuoteFU = True Then GoTo RowCntMsg '04-28-15 JTC 02-22-12      If SESCO = True Then GoTo RowCntMsg '04-28-15 JTC 02-22-12
            PQTCUST = "" '11-29-13 Public "HITNONE"
            RowCnt = 0
            Dim LastRecordDup As String = "" '06-14-10 Eliminate Duplicates
            drQToRow = dsQuoteRealLU.QuoteRealLU.Rows(0) '02-11-09 
            Dim Hit As Short = 0 '11-27-13 moved up 
            'Debug.Print(Format(Now, "hh:MM"))
            For Each drQToRow In dsQuoteRealLU.QuoteRealLU.Rows 'dsQutLU
                If drQToRow.RowState = DataRowState.Deleted Then Continue For ' GoTo 235 '06-19-08
                If RealTgLookupExcel = True Then '11-27-13 and
                    Hit = 0  ' Hit = drQToRow.Item.count
                    If IsDBNull(drQToRow("Typec")) Then
                        '11-29-30 0        1            2       3       4          5          6       7     8       9             10    11    12     13       14           15        16     17        18         19         20       21       22     23      24         25    26   27        28          29         30          31           32       33        34          35   36     37     38         39            40         41      41     43      44         45        46    47      48       49     50       51        52        
                        'ProjectCustID, ProjectID, QuoteCode, NCode, FirmName, ContactName, SLSCode, Got, Typec, MFGQuoteNumber, Cost, Sell, Comm, Overage, QuoteToDate, OrdDate, NotGot, Comments, SPANumber, SpecCross, LotUnit, LPCost, LPSell, LPComm, LampsIncl, Terms, FOB, QuoteID, BranchCode, LeadTime, LastChgDate, LastChgBy, Requested, FileName, QuoteCode, SellQ, CostQ, CommQ, JobName, MarketSegment, EntryDate, BidDate, SLSQ, Status, RetrCode, SelectCode, City, State, lastChgBy, CSR, LotUnit, StockJob, TypeOfJob
                        For I = 0 To 51
                            If I = 2 Then drQToRow(I) = "" : drQToRow("QuoteCode") = drQToRow("QuoteCode") : Continue For '04-17-15 JTC Fix "QuoteCodeQ" error
                            If I = 3 Then drQToRow(I) = "HITNONE" : Continue For 'NCode = "HITNONE"
                            If I = 14 Or I = 14 Or I = 41 Or I = 51 Then drQToRow(I) = "1900-01-01" : Continue For 'EntryDate & BidDate '41 = LastChange
                            If I = 15 Then drQToRow(I) = False : Continue For 'NotGot Boolean
                            If I = 30 Or I = 31 Then Continue For 'EntryDate & BidDate
                            'Debug.Print(drQToRow(I).GetType.ToString)
                            drQToRow(I) = 0
                            drQToRow("SLSCode") = drQToRow("SLSQ")
                            drQToRow("Sell") = drQToRow("SellQ")
                        Next
                        drQToRow("Typec") = "C" '
                    End If
                    If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then '01-21-14
                        ' RealALL = False ' True '01-21-14 JTC Select Specifiers
                        If drQToRow.Typec = "A" Then Hit = 1 'eC = 'A' " 'Arch
                        If drQToRow.Typec = "E" Then Hit = 1 ' 'E' " 'Eng
                        If drQToRow.Typec = "L" Then Hit = 1 ' = 'L' " 'LtgDesigner
                        If drQToRow.Typec = "S" Then Hit = 1 ' 'S' " ' Specifier
                    End If
                    If PQTCUST = "HITNONE" Then GoTo SkipRealTest '11-29-13 : Continue For 'NCode = "HITNONE" PQTCUST = ""'11-29-13 Public
                    If RealCustomer = True And drQToRow.Typec = "C" Then Hit = 1
                    If RealManufacturer = True And drQToRow.Typec = "M" Then Hit = 1 ' 'MFG
                    If RealQuoteTOOther = True And drQToRow.Typec = "O" Then Hit = 1 ' = 'O' " 'Other
                    If RealSLSCustomer = True And drQToRow.Typec = "C" Then Hit = 1 ' 'C' " 'SLS/Customer
                    If (RealArchitect = True Or RealALL = True) And drQToRow.Typec = "A" Then Hit = 1 'eC = 'A' " 'Arch
                    If (RealEngineer = True Or RealALL = True) And drQToRow.Typec = "E" Then Hit = 1 ' 'E' " 'Eng
                    If (RealLtgDesigner = True Or RealALL = True) And drQToRow.Typec = "L" Then Hit = 1 ' = 'L' " 'LtgDesigner
                    If (RealSpecifier = True Or RealALL = True) And drQToRow.Typec = "S" Then Hit = 1 ' 'S' " ' Specifier
                    If (RealContractor = True Or RealALL = True) And drQToRow.Typec = "T" Then Hit = 1 ' = 'T' " 'Contractor
                    If (RealOther = True Or RealALL = True) And drQToRow.Typec = "Y" Then Hit = 1 '= 'X' " 'Other 01-31-12
                    'If RealTgLookupExcel = True And Me.pnlTypeOfRpt.Text = "Realization" Then '01-20-14
                    '    If drQToRow.Typec = "A" Or drQToRow.Typec = "E" Or drQToRow.Typec = "L" Or drQToRow.Typec = "S" Then
                    '        Hit = 1 'eC = 'A' " 'Arch
                    '    End If
                    'End If
                    If Hit = 0 Then drQToRow.Delete() : Continue For '11-27-13
                End If 'Not Excel
SkipRealTest:
                '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero" BUsinessType = T 'Debug.Print(drQToRow.BusinessType) '  = "T" ThenStop
                If Me.chkMfgBreakdown.Text = " Print Quote Hdr Amt when QuoteTo is Zero" And Me.cboTypeCustomer.Text.Trim.ToUpper = "T" And drQToRow.BusinessType <> "T" And Me.chkMfgBreakdown.Checked = True Then ' "Add MFG Total Breakdown to Reports" '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero" 
                    drQToRow.Delete() : Continue For
                End If

                If RealTgLookupExcel = True Then '11-27-13 
                Else  'Don't Fill if Excel
                    If IsDBNull(drQToRow("NCode")) Then drQToRow.NCode = "" '12-31-14 JTC Null Fix
                    If IsDBNull(drQToRow("BranchCode")) Then drQToRow.BranchCode = "" '12-31-14 JTC Null Fix
                    If IsDBNull(drQToRow("QuoteCode")) Then drQToRow.QuoteCode = "" '12-31-14 JTC Null Fix
                    If drQToRow.NCode = "" Then drQToRow.NCode = VB.Left(drQToRow.FirmName, 8) ' Blank" '07-27-12
                End If
                'If drQToRow("Other") Is Nothing ThenStop
                'If IsDBNull(drQToRow("Other")) ThenStop
                'If drQToRow.Other Is Nothing ThenStop
                '11-27-13Dim Hit As Short = 0 '05-19-11
                Hit = 1 '01-03-14 Turned on in SelectHit9500    was 0 '11-27-13 
                Call SelectHit9500(Hit, multsrtrvs) '01-25-09
                'Debug.Print(RowCnt)
                If drQToRow.FirmName.Length > MaxNameLength Then drQToRow.FirmName = drQToRow.FirmName.Substring(0, MaxNameLength) '01-06-12  Trim CustName
                If drQToRow.JobName.Length > MaxJobLength Then drQToRow.JobName = drQToRow.JobName.Substring(0, MaxJobLength) '12-22-12 Trim CustName

                '04-20-15 JTC Fix Realization Salesman/Name Code/QuoteCode Duplicates
                If Me.txtPrimarySortSeq.Text = "Salesman" And txtSecondarySort.Text = "Name Code" Then 'Realization  Sort By = Salesman / Name Code
                    If LastRecordDup.Trim = "" Then
                        LastRecordDup = drQToRow.SLSCode & drQToRow.NCode & drQToRow.QuoteCode
                    Else
                        If LastRecordDup = drQToRow.SLSCode & drQToRow.NCode & drQToRow.QuoteCode Then Hit = 0 ' No dups for now
                    End If
                    GoTo SkipLastDup
                End If
                '???JTC I don't see how below works because I don,t reset LastRecordDup anywhere below???????????
                If LastRecordDup.Trim = "" Then
                    LastRecordDup = drQToRow.NCode & drQToRow.QuoteCode '06-14-10
                Else
                    If LastRecordDup = drQToRow.NCode & drQToRow.QuoteCode Then Hit = 0 '06-14-10 No dups for now
                End If
SkipLastDup:

                'Test If drQToRow("NCode") = "GES/AT" ThenStop '01-03-14
                If drQToRow("NCode") = "HITNONE" Then Hit = 1 '11-29-13 'NCode = "HITNONE" PQTCUST = ""'11-29-13 Public
                If Hit = 0 Then drQToRow.Delete() : Continue For Else RowCnt += 1 '02-06-09  RowState = DataRowState.Deleted
                If drQToRow.Typec = "C" Or drQToRow.Typec = "M" Then 'don't do on Customer or Mfg
                Else 'If Sell on QTO = 0 then use Job Sell
                    'Debug.Print(drQToRow.Sell)
                    If IsDBNull(drQToRow("SellQ")) Then drQToRow.SellQ = 0 '01-20-12
                    If drQToRow.Sell = 0 Then drQToRow.Sell = drQToRow.SellQ
                    If IsDBNull(drQToRow("CostQ")) Then drQToRow.CostQ = 0 '01-20-12
                    If drQToRow.Cost = 0 Then drQToRow.Cost = drQToRow.CostQ
                    If IsDBNull(drQToRow("CommQ")) Then drQToRow.CommQ = 0 '01-20-12
                    If drQToRow.Comm = 0 Then drQToRow.Comm = drQToRow.CommQ
                End If
                If RealQuoteToAmtON = True Then '07-15-14  Public RealQuoteToAmtON As Boolean = 0 
                    If Me.chkMfgBreakdown.Text = " Print Quote Hdr Amt when QuoteTo is Zero" And Me.cboTypeCustomer.Text.Trim.ToUpper = "T" And drQToRow.BusinessType = "T" And Me.chkMfgBreakdown.Checked = True Then GoTo CustTypeTOnly '01-27-15 JTC
                    '10-30-14 JTC If drQToRow.Typec = "M" Then 'Mfg Only don't do on Customer out drQToRow.Typec = "C" Or 
                    'RealCustomer = True  SelTypec = "C" '07-15-14  = 'M'Me.pnlQutRealCode.Text = "Select Cust Code"
                    US = "" '10-30-14 JTC
                    If RealManufacturer = False And RealCustomer = True And RealQuoteTOOther = False And RealSLSCustomer = False Then US = "C" '10-30-14 JTC07-15-14  = 'M'Me.pnlQutRealCode.Text = "Select Cust Code"
                    If RealManufacturer = True And RealCustomer = False And RealQuoteTOOther = False And RealSLSCustomer = False Then US = "M" '10-30-14 JTC 07-15-14 Me.pnlQutRealCode.Text = "Select MFG Code" '"Select MFG Code"
                    '12-15-14 JTC If drQToRow.Typec = "C" And US = "M" Then Hit = 0 : drQToRow.Delete() : RowCnt -= 1 : Continue For '10-30-14 If RealQuoteToAmtON = True For Mfg Delete c=Customer Records DataRowState.Deleted
                    '12-15-14 JTC drQToRow.Typec = "M" And US = "M" must be the same Realization one MFG or one Cust
                    If US <> "" Then '12-15-14 JTC skip Delete Logic
                        If US = "M" Then '12-18-14 JTC Fix was deleting others 
                            If drQToRow.Typec = "M" And US = "M" Then Continue For Else Hit = 0 : drQToRow.Delete() : RowCnt -= 1 : Continue For '12-15-14 JTC10-30-14 If RealQuoteToAmtON = True For Mfg Delete c=Customer Records DataRowState.Deleted
                        End If
                        '11-17-14 JTC Fix Realization Just Mfg Or Just Cust If drQToRow.Typec = "M" And US = "M" Then Hit = 0 : drQToRow.Delete() : RowCnt -= 1 : Continue For '10-30-14 Don't Show Mfg Record Per Jaci DataRowState.Deleted
                        '12-15-14 JTC If drQToRow.Typec = "M" And US = "C" Then Hit = 0 : drQToRow.Delete() : RowCnt -= 1 : Continue For '10-30-14 DataRowState.Deleted
                        If US = "C" Then '12-18-14 JTC Fix was deleting others
                            If drQToRow.Typec = "C" And US = "C" Then Continue For Else Hit = 0 : drQToRow.Delete() : RowCnt -= 1 : Continue For '12-15=14 JTc '10-30-14 DataRowState.Deleted
                        End If
                    End If
                    ''11-17-14 JTC Fix Realization Just Mfg Or Just Cust If drQToRow.Typec = "C" And US = "C" Then Hit = 0 : drQToRow.Delete() : RowCnt -= 1 : Continue For '10-30-14 Don't Show Customer DataRowState.Deleted
                    If drQToRow.Typec <> US Then '10-30-14 JTC Use RealQuoteToAmtON on all Typec Except "C" Cust 'Mfg Only don't do on Customer out drQToRow.Typec = "C" Or 
                        If IsDBNull(drQToRow("TMPSellQ")) Then drQToRow.SellQ = 0 Else drQToRow.Sell = drQToRow("TMPSellQ")
                        If IsDBNull(drQToRow("TMPCostQ")) Then drQToRow.CostQ = 0 Else drQToRow.Cost = drQToRow("TMPCostQ")
                        If IsDBNull(drQToRow("TMPCommQ")) Then drQToRow.CommQ = 0 Else drQToRow.Comm = drQToRow("TMPCommQ")
                    End If
                End If
CustTypeTOnly:  '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero" Contractor drQToRow.BusinessType <> "T"
                If Me.chkMfgBreakdown.Text = " Print Quote Hdr Amt when QuoteTo is Zero" And Me.cboTypeCustomer.Text.Trim.ToUpper = "T" And drQToRow.BusinessType = "T" And Me.chkMfgBreakdown.Checked = True Then ' "Add MFG Total Breakdown to Reports" '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero"
                    If IsDBNull(drQToRow("SellQ")) Then drQToRow.SellQ = 0 '01-20-12
                    drQToRow.Sell = drQToRow.SellQ
                    If IsDBNull(drQToRow("CostQ")) Then drQToRow.CostQ = 0 '01-20-12
                    drQToRow.Cost = drQToRow.CostQ
                    If IsDBNull(drQToRow("CommQ")) Then drQToRow.CommQ = 0 '01-20-12
                    drQToRow.Comm = drQToRow.CommQ
                End If
                '04-20-15 JTC Fix Realization Salesman/Name Code/QuoteCode Duplicates
                If Me.txtPrimarySortSeq.Text = "Salesman" And txtSecondarySort.Text = "Name Code" Then  'Realization  Sort By = Salesman / Name Code
                    LastRecordDup = drQToRow.SLSCode & drQToRow.NCode & drQToRow.QuoteCode
                End If
            Next
RowCntMsg:
            'Debug.Print(SortSeq)
            If RowCnt = -1 Then RowCnt = 0 '03-03-11
            US = "Records Selected Before Filtering = " & RowCnt.ToString & vbCrLf & "Click Run Report Button." '01-27-15 JTC 
            If SESCO = True Then US += "  SESCO Job List to Excel"
            If ExcelQuoteFU = True Then US = "Records Selected Before Filtering = " & RowCnt.ToString & vbCrLf & " View Excel Quote FollowUp Report" '04-28-15 JTC 
            If RealTgLookupExcel = True And SESCO = False Then GoTo RowCntZero '12-02-13
            MsgBox(US) '02-26-12"Records Selected Before Filtering = " & RowCnt.ToString) '10-17-10 
            If SESCO = True Then Call PrintSESCOJobListRealReportQutTO() '04-22-15 JTC 02-22-12
            If ExcelQuoteFU = True Then Me.chkShowLatestCust.Checked = True '04-22-15 JTC 
            If ExcelQuoteFU = True Then Call ExcelQuoteFollowUp() '04-28-15 JTC Old 
            'Me.cboSortPrimarySeq.Text = "Excel Quote FollowUp"
RowCntZero:
            If RowCnt = 0 Then
                MsgBox("Record Count is Zero. Please check your selection again") '05-18-11
                Me.tabQrt.SelectedIndex = 1
                Call tabQRT_TabActivate(1) '02-01-09
                Me.Focus()
                GoTo EndExit '10-13-14 JTC Exit Sub
            End If
            Me.tgQh.Visible = False ' Quote Grid
            Me.tgr.Visible = True ' Prof Cust Grid
            Me.tgr.Dock = DockStyle.Fill '08-18-09 JH
            Me.tgln.Visible = False '08-18-09 JH
            Me.tgr.Rebind(True) '02-07-09
            If SESCO = True Then GoTo EndExit '10-13-14 JTC Exit Sub '02-22-12
            Dim ShowAllQuoteHeader As String = "" '06-06-11
            If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then ShowAllQuoteHeader = "ShowAll" '06-06-11 = "Show All Quote Header Fields" 

            If DIST Then 'VQrtLinesDistTGLayoutCurrent
                'If RealTgLookupExcel = False Then '07-31-14 JTC This is run After show hide and Loading Layout wipes out showhide column visible '07-31-14JTC Fix Header 
                If RealTgLookupExcel = False Then Call TrueGridLayoutFiles("Load", "Curr", "VQrtQuoteToDistTGLayout" & ShowAllQuoteHeader, tgr) '06-06-11 = "Show All Quote Header Fields"
            Else
                '07-31-14 JTC Don't do if True 
                If RealTgLookupExcel = False Then Call TrueGridLayoutFiles("Load", "Curr", "VQrtQuoteToTGLayout" & ShowAllQuoteHeader, tgr) '06-06-11 = "Show All Quote Header Fields"
                'tgr.Rebind(False)
                '02-08-12 JTC Not On Realization Fix ("Comm-$").Visible = True error
                If Me.pnlTypeOfRpt.Text = "Realization" Then  Else Me.tgr.Splits(0).DisplayColumns("Comm-$").Visible = True '10-04-10
                '03-19-14 Me.tgr.Splits(0).DisplayColumns("LPComm").Visible = True '12-08-09
                '03-19-14 Me.tgr.Splits(0).DisplayColumns("Overage").Visible = True '12-08-09
                '02-03-12 Me.tgr.Splits(0).DisplayColumns("Margin").Visible = False '11-04-10  
                '11-27-12 JTC Book OKMe.tgr.Splits(0).DisplayColumns("Cost").Visible = False '12-08-09
                '03-19-14 Me.tgr.Splits(0).DisplayColumns("LPCost").Visible = False '12-08-09
            End If
            If Me.pnlTypeOfRpt.Text = "Realization" Then 'If txtPrimarySortSeq.Text = "Specifier Credit" Then 'sender.ToString '"Quote Summary  Sort By = "Specifier Credit
                'Me.tgr.Splits(0).DisplayColumns("Comm-$").Visible = True '10-04-10
                'Me.tgr.Splits(0).DisplayColumns("LPComm").Visible = True '12-08-09
                'Me.tgr.Splits(0).DisplayColumns("Overage").Visible = True '12-08-09
                '03-19-14 If DIST Then Me.tgr.Splits(0).DisplayColumns("Margin").Visible = True '02-03-12 
                '03-19-14 If DIST Then Me.tgr.Splits(0).DisplayColumns("Cost").Visible = True '02-03-12
                'Call AddTgColumns("Reference", "Near", tgQh) '02-06-10 NearFarCenter =
                '03-19-14 Call AddTgColumns("MarketSegment", "Near", tgr) '06-29-12
                '03-19-14 Me.tgr.Columns("Comm").NumberFormat = "n2"
                '03-19-14 Me.tgr.Splits(0).DisplayColumns("Comm").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far  ' HorizontalAlignment.Far '02-09-10
                '03-19-14 Me.tgr.Columns("Sell").NumberFormat = "n2"
                '03-19-14 Me.tgr.Splits(0).DisplayColumns("Sell").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far  ' HorizontalAlignment.Far '02-09-10
            Else
                'If Me.tgQh.Visible = True Then Call AddTgColumns("MarketSegment", "Near", tgQh) '06-29-12
            End If
            'Dim tmpText As String = txtPrimarySortSeq.Text ' & "/" & txtSecondarySort
            'Dim tmpText2 As String =  "/" & txtSecondarySort:tmpText = 
            Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
            Me.txtSortSeqCriteria.Text = SelectionText '02-24-09
            '02-24-09 Me.txtSortSeq.Text = txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
            Me.txtSortSeqV.Text = Me.txtSortSeq.Text ' txtPrimarySortSeq.Text & "/" & txtSecondarySort.Text
            If VB.Right(Me.txtSortSeq.Text, 1) = "/" Then Me.txtSortSeq.Text = Replace(Me.txtSortSeq.Text, "/", "")
            Me.tabQrt.SelectedIndex = 2
            Call tabQRT_TabActivate(2) '02-01-09
            Me.Focus()
            frmShowHideGrid.BringToFront() '02-10-09
            GoTo EndExit '10-13-14 JTC Exit Sub
        End If 'End of REALIZATION %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And Me.txtSlsSplit.Text <> "ALL" Then '05-01-13 Me.chkSlsFromHeader.CheckState = CheckState.Checked Then
            If Me.txtSlsSplit.Text <> "" Then Me.chkSlsFromHeader.CheckState = CheckState.Checked '05-01-13 
        End If
        Select Case txtPrimarySortSeq.Text 'sender.ToString '"Quote Summary  Sort By = Specifier Credit
            Case "Entry Date"
                SortSeq = "quote.EntryDate"
            Case "Bid Date"
                SortSeq = "quote.BidDate"
            Case "Job Name"
                '11-23-11 quote.JobName Not project.ProjectName
                SortSeq = "quote.JobName" '11-23-11 quote.JobName Not project.ProjectName
            Case "Quote Code"
                SortSeq = "quote.QuoteCode"
            Case "Followed By" '03-03-12
                SortSeq = "quote.FollowedBy"
            Case "Entered By" '05-14-13
                SortSeq = "quote.EnteredBy"
            Case "Salesman"
                SortSeq = "quote.SLSQ"
                If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And Me.chkSlsFromHeader.CheckState = CheckState.Checked Then '03-08-13 JTC Add QS.SLSCode as SLS1 from Quote QUTSLSSPLIT Table Position 1
                    SortSeq = "QS.SLSCode" 'Salesman
                End If
            Case "Status"
                SortSeq = "quote.Status"
            Case "Location"
                SortSeq = "quote.Location" '11-19-10 City"
            Case "Retrieval Code"
                SortSeq = "quote.RetrCode"
            Case "Descending Dollar"
                SortSeq = "quote.Sell"
            Case "Market Segment"
                SortSeq = "Project.MarketSegment"
            Case Else
                SortSeq = "quote.QuoteCode"
        End Select
        Try
            'SORT ORDER CODE
            OrderBy = BuildSQLOrderBY() 'From SortSeq = "Q.Reference" 01-30-09
            If OrderBy = Nothing Then  Else OrderBy = Replace(OrderBy.ToUpper, "ORDER BY ORDER BY", "ORDER BY") '11-15-13 JTC Fix = Replace(strSql, "ORDER BY ORDER BY", "ORDER BY")
            If OrderBy = Nothing Then  Else OrderBy = Replace(OrderBy.ToUpper, "ORDER BY  ORDER BY", "ORDER BY") '01-16-214
            'SQL SECTION *********************************************************************************************
            strSql = ""
            'SELECT project.ProjectName, Quote * FROM quote LEFT OUTER JOIN project ON quote.ProjectID = project.ProjectID
            'strsql = "Select project.ProjectName, project.ProjectID, quote.QuoteID, quote.QuoteCode, quote.EntryDate, quote.EndDate, quote.BidDate, quote.Status, quote.SLSQ, quote.EnteredBy, quote.Sell from quote left join project on quote.ProjectID = project.ProjectID "
            '11-23-11 quote.JobName Not project.ProjectName
            '11-23-11strSql = "Select Q.*, P.projectname, P.MarketSegment from Quote Q left join Project P on P.ProjectID = Q.ProjectID " 'where Q.EntryDate > " & "'" & "2008-05-08" & "'" ' " '01-25-09 "'" & "'" & "M%" & "'"  07-08-10 LEFT
            '10-31-12 Del Market from Q not P strSql = "Select Q.*, P.MarketSegment from Quote Q left join Project P on P.ProjectID = Q.ProjectID "
            strSql = "Select Q.* from Quote Q " '10-31-12 left join Project P on P.ProjectID = Q.ProjectID "
            'If Trim(Me.txtSlsSplit.Text) <> "" Then '05-05-10 And Trim(frmQuoteRpt.txtSalesman.Text) <> "ALL" Then
            '    'with SlsCode
            If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And Me.chkSlsFromHeader.CheckState = CheckState.Checked Then '03-08-13
                '03-08-13 JTC Add QS.SLSCode as SLS1 from Quote QUTSLSSPLIT Table Position 1
                strSql = "Select Q.*, QS.SLSCode as SLSCode from Quote Q  LEFT JOIN QUTSLSSPLIT QS ON Q.QuoteID = QS.QuoteID  " '03-08-13 OrderBy = "QS.SLSCode" 'Salesman
                'Select Q.*, QS.SLSCode as SLSCode from Quote Q  LEFT JOIN QUTSLSSPLIT QS ON Q.QuoteID = QS.QuoteID 
                ' End If  ' Me.chkSlsFromHeader.CheckState = CheckState.Unchecked  ' Me.chkSlsFromHeader.Text = "Use Quote SLS 1 Split for Salesman" '03-08-13 "Use Salesman From Quote Header on Report"
                'If frmQuoteRpt.txtSlsSplit.Text.ToUpper <> "ALL" Then 'add the following
                '02-22-13 JTC Add QS.SLSCode as SLSCode from Quote QUTSLSSPLIT Table All Sls's from table Realization 
                'strSql = "SELECT QS.SLSCode as SLSCode, projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID " '02-22-13
                'strSql += " LEFT JOIN QUTSLSSPLIT QS ON Quote.QuoteID = QS.QuoteID " '02-22-13 Sls 1-4 AND QS.slsnumber = 1 "
            End If
            '11-04-14 JTC Quote Summary with Sales Splits
            If Me.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) And Me.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then '11-04-14 JTC
                strSql = "Select Q.*, QS.SLSCode as SLSCode, QS.SLSSplit as QSSpecCredit from Quote Q  LEFT JOIN QUTSLSSPLIT QS ON Q.QuoteID = QS.QuoteID  " '03-08-13 OrderBy = "QS.SLSCode" 'Salesman
            End If
            'Specifier Credit was moved to Realization ******************************************************************************************
            'If txtPrimarySortSeq.Text = "Specifier Credit" Then 'sender.ToString '"Quote Summary  Sort By = "Specifier Credit
            '    'strSql = "SELECT PC.NCode as SpecCredit, P.ProjectName, P.MarketSegment, Q.EntryDate, Q.BidDate, Q.SLSQ, Q.Status, Q.RetrCode, Q.SelectCode, Q.City, Q.State, Q.lastChgBy, Q.CSR, Q.LotUnit, Q.StockJob, Q.TypeOfJob FROM project P INNER JOIN projectcust PC ON P.ProjectID = PC.ProjectID INNER JOIN quote Q ON PC.QuoteID = Q.QuoteID "
            '    'Carefull I Put Specifier in Reference '11-23-11 quote.JobName Not project.ProjectName
            '    '10-31-12 No P.Mar strSql = "SELECT PC.NCode as Reference, P.MarketSegment, Q.* FROM project P INNER JOIN projectcust PC ON P.ProjectID = PC.ProjectID INNER JOIN quote Q ON PC.QuoteID = Q.QuoteID "
            '    strSql = "SELECT PC.NCode as Reference, Q.* FROM quote Q INNER JOIN projectcust PC ON PC.QuoteID = Q.QuoteID " '10-31-12 
            '    SortSeq = "Q.Reference" : OrderBy = "and PC.NCode <> '' and PC.Typec <> 'M' and PC.Typec <> 'C' and PC.Typec <> 'O' ORDER BY Reference" '02-06=10
            'End If
            If Me.pnlTypeOfRpt.Text = "Project Shortage Report" Then '05-17-10 
                '11-30-10 strSql = "Select Q.*, projectname, P.MarketSegment from Quote Q LEFT join Project P on P.ProjectID = Q.ProjectID " ' 05-17-10  07-08-10 LEFT
                strSql = "Select Q.* from Quote Q " '11-30-10 and Q.EntryDate >= '2010-1-1' and Q.Entrydate <= '2010-11-30' ORDER BY Q.JobName 'Select Q.*, from Quote " '11-30-10  
                OrderBy = "ORDER BY Q.JobName" '11-30-10
            End If
            '09-10-12 
            If Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '09-10-12 txtPrimarySortSeq.Text = "MFG Follow-Up Report"

                'TerrSpecRegAllPaidUnPaid As String = "A"'09-10-12 A=All,P=Paid, U=Unpaid 
                '10-31-12 Del Market from Q not P strSql = "Select Q.*, P.MarketSegment from Quote Q left join Project P on P.ProjectID = Q.ProjectID "
                strSql = "Select Q.* from Quote Q " '10-31-12 left join Project P on P.ProjectID = Q.ProjectID "

                If txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Then
                    Resp = MsgBox("Do you want One Salesman per page?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "MFG Follow-Up Report") '09-19-12
                    If Resp = vbYes Then Me.chkSalesmanPerPage.CheckState = CheckState.Checked Else Me.chkSalesmanPerPage.CheckState = CheckState.Unchecked
                    If Resp = vbCancel Then GoTo EndExit '10-13-14 Exit Sub
                    GoTo QutLineHistoryRpt '09-19-12
                ElseIf txtPrimarySortSeq.Text = "MFG Follow-Up Report" Then
                    Resp = MsgBox("Do you want One MFG per page?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "MFG Follow-Up Report") '09-19-12
                    If Resp = vbYes Then Me.chkSalesmanPerPage.CheckState = CheckState.Checked Else Me.chkSalesmanPerPage.CheckState = CheckState.Unchecked
                    If Resp = vbCancel Then GoTo EndExit '10-13-14 Exit Sub
                    GoTo QutLineHistoryRpt '09-19-12
                    'strSql = "Select Q.*, P.MarketSegment, PL.* from Quote Q left join Project P on P.ProjectID = Q.ProjectID left JOIN projectLines PL ON Q.QuoteID = PL.QuoteID " '09-19-12
                End If
            End If
            ''05-01-13 Moved Down
            'If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And Me.chkSlsFromHeader.CheckState = CheckState.Checked Then
            '    If Me.txtSlsSplit.Text <> "ALL" Then SortCode = Me.txtSlsSplit.Text '05-01-13 
            'End If
            'If SortCode <> "" Then
            '    If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And Me.chkSlsFromHeader.CheckState = CheckState.Checked Then
            '        strSql += "where QS.slsnumber = 1 and " & SortSeq & " >= '" & SortCode & "' "
            '    Else
            '        strSql += "where " & SortSeq & " >= '" & SortCode & "' "
            '    End If
            'End If
            '11-01-10
            Dim sStartDate As Date = Me.DTPickerStartEntry.Value '], 	yyyy'-'MM'-'dd HH':'mm':'ss'Z'Date = Me.DTPickerStartEntry.Value
            Dim month As String = sStartDate.Month
            Dim day As String = sStartDate.Day
            Dim Year As String = sStartDate.Year
            Dim SD As String = Year & "-" & month & "-" & day
            'Test OK SD = sStartDate.ToString("yyyy-MM-dd") '11-01-10 This works
            Dim sEndDate2 As Date = Me.DTPicker1EndEntry.Text 'Dim sStartDate As Date = Me.DTPickerStartEntry.Value 
            month = sEndDate2.Month
            day = sEndDate2.Day
            Year = sEndDate2.Year
            Dim ED As String = Year & "-" & month & "-" & day
            Dim BidsOnly As String = "" '07-20-11
            If Me.chkBidJobsOnly.Text = "Delivery Date Jobs Only" Then strSql += " where Q.TypeOfJob = '" & "Q" & "' " : GoTo SkipEntryDate '10-22-13"
            If chkBidJobsOnly.Checked = True Then '07-21-11 Report Only Bid Jobs
                BidsOnly = " and Q.BidBoard = 'Y' " '07-21-11
            End If
            Dim JT As String = Me.cboTypeofJob.Text
            If JT = "A" Or JT = "" Or JT = "*" Then '=ALL 05-08-10 No Select on Q.TypeOfJob JT = "*"
                '11-13-09 strsql += "where Q.EntryDate >= '" & sStartDate & "' and Q.BidDate <= '" & sEndDate & "' " & OrderBy '01-30-09"order by Q.QuoteCode " ' and Q.EndDate =< '20091231' " & "order by Q.QuoteCode " '01-25-09
                strSql += "where Q.EntryDate >= '" & SD & "' and Q.Entrydate <= '" & ED & "' " & BidsOnly '02-03-12 & OrderBy '07-21-11 BidsOnly11-13-09 01-30-09"order by Q.QuoteCode " ' and Q.EndDate =< '20091231' " & "order by Q.QuoteCode " '01-25-09
            Else
                '11-13-09 strsql += "where Q.EntryDate >= '" & sStartDate & "' and Q.BidDate <= '" & sEndDate & "' " & OrderBy '01-30-09"order by Q.QuoteCode " ' and Q.EndDate =< '20091231' " & "order by Q.QuoteCode " '01-25-09
                strSql += "where Q.TypeOfJob = '" & JT & "' and Q.EntryDate >= '" & SD & "' and Q.Entrydate <= '" & ED & "' " & BidsOnly '02-03-12 & OrderBy ''07-21-11 BidsOnly 11-13-09 01-30-09"order by Q.QuoteCode " ' and Q.EndDate =< '20091231' " & "order by Q.QuoteCode " '01-25-09 11-01-10 startdate to SD
            End If
SkipEntryDate:  '10-22-13
            If Me.ChkCheckBidDates.CheckState = CheckState.Checked Then '10-22-13
                Dim sEndBidDate As String = VB6.Format(Me.DTPicker1EndBid.Value, "yyyy-MM-dd") ''02-03-12 - not /
                Dim sStartBidDate As String = VB6.Format(Me.DTPicker1StartBid.Value, "yyyy-MM-dd") '
                '10-08-15 JTC '10-08-15 JTC Include Blank Bid Dates 
                If Me.chkBlankBidDates.CheckState = CheckState.Checked Then '10-08-15 JTC Include Blank Bid Dates
                    strSql += " and (Q.BidDate >= '" & sStartBidDate & "' and Q.BidDate <= '" & sEndBidDate & "'  or  Q.BidDate is null or Q.biddate = '" & "1900-01-01" & "' ) " '10-08-15 04-04-12 added or Q.biddate = '" & "1900-01-01" & "'
                Else
                    '10-08-15 JTC 
                    strSql += " and (Q.BidDate >= '" & sStartBidDate & "' and Q.BidDate <= '" & sEndBidDate & "' ) " '10-23-13 or  Q.biddate is null or Q.biddate = '" & "1900-01-01" & "') " '04-04-12 added or Q.biddate = '" & "1900-01-01" & "'
                End If
                If Me.chkBidJobsOnly.Text = "Delivery Date Jobs Only" Then strSql = Replace(strSql, "Q.BidDate", "Q.EstDelivDate") '10-22-13 '11-19-13 JTC Forecast Fix BidDate error added If Me.chkBidJobsOnly.Text = "Delivery Date Jobs Only" Then 
                '10-08-15 JTC
            End If
            '05-01-13 Chg >= to = Moved Down SortCode = Me.txtSlsSplit.Text VQRT2.RptMajorType.RptSalesman And Me.chkSlsFromHeader.CheckState = CheckState.Checked 
            If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And Me.chkSlsFromHeader.CheckState = CheckState.Checked Then
                If Me.txtSlsSplit.Text <> "ALL" Then SortCode = Me.txtSlsSplit.Text '05-01-13 
            End If
            If Me.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) And Me.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then '06-02-15  JTC
                If Me.txtSlsSplit.Text <> "ALL" Then
                    SortCode = Me.txtSlsSplit.Text
                    strSql += " and QS.slsnumber = 1 and QS.SLSCode = '" & SortCode & "' "
                    GoTo SkipSortCode '06-02-15
                End If
            End If
            If SortCode <> "" Then
                If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And Me.chkSlsFromHeader.CheckState = CheckState.Checked Then
                    strSql += " and QS.slsnumber = 1 and " & SortSeq & " = '" & SortCode & "' " '05-01-13 use = not >= SortSeq & " >= '" & SortCode & "' "
                Else '05-01-13 Use and not where
                    strSql += "and " & SortSeq & " >= '" & SortCode & "' "
                End If
            End If
SkipSortCode:  '10-16-13 JTC 
            If Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And SortSeq = "quote.QuoteCode" Then OrderBy = " ORDER BY Q.QuoteCode " '10=16-13 
            '11-04-14 JTC Quote Summary with Sales Splits
            If Me.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) And Me.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then '11-04-14 JTC
                If VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Then OrderBy = " ORDER BY Q.QuoteCode, SLSCode " Else OrderBy = " ORDER BY Q.JobName, SLSCode "
                If Me.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then '06-02-15 JTC
                    Resp = MsgBox("Do you want Report in SLS-1-4 Sequence with One Salesman per page?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "SLS-1-4 Sequence with One Salesman per page option") '06-02-15 JTC
                    If Resp = vbYes Then Me.chkSalesmanPerPage.CheckState = CheckState.Checked Else Me.chkSalesmanPerPage.CheckState = CheckState.Unchecked
                    If Resp = vbYes And VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Then OrderBy = " ORDER BY SLSCode, Q.QuoteCode " Else OrderBy = " ORDER BY SLSCode, Q.JobName " '06-02-15 JTC 
                End If
                'Call AddTgColumns("Follow By", "Near", tgQh) '11-04-14 JTC 
                'Call AddTgColumns("QSSpecCredit", "Near", tgQh) '11-04-14 JTC Me.tgln.Splits(0).DisplayColumns("Reference").Visible = True '02-06-10
            End If
            strSql += OrderBy '02-03-12 & OrderBy
            strSql = Replace(strSql, "ORDER BY ORDER BY", "ORDER BY") '02-25-12 ORDER BY ORDER BY 
            'Debug.Print(strsql)  'strsql += "order by " & SortSeq & " " & SortDir
            'END SQL SECTION *****************************************************************************************
            'Don't Need Use SpecCredit Field
            If txtPrimarySortSeq.Text = "Specifier Credit" Then 'sender.ToString '"Quote Summary  Sort By = "Specifier Credit
                Call AddTgColumns("Reference", "Near", tgQh) '02-06-10 NearFarCenter =
            End If
        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "Print Task" & vbCrLf)
            ' IfDebugOn ThenStop 'CatchStop
        End Try
        RowCnt = 0
        '02-06-10 Put StrSql Logic here
        Call FillQutLUDataSet(SortSeq, SortDir)
        If txtPrimarySortSeq.Text = "Specifier Credit" Then
            Me.tgln.Splits(0).DisplayColumns("Reference").Visible = True '02-06-10
        End If
        If Me.pnlTypeOfRpt.Text = "Quote Summary" And Me.ChkSpecifiersCustInCols.Checked = True Then '06-25-18 06-27-18
            Dim PrintFirmName As Boolean = vbYes
            Resp = MessageBox.Show("Do you want the Code or Firm Name listed?  Yes = Code, No = Firm Name", "Code/Firm Name", MessageBoxButtons.YesNo)
            If Resp = vbYes Then PrintFirmName = False
            Call AddTgColumns("Architect", "Near", tgQh)
            Call AddTgColumns("Engineer", "Near", tgQh)
            Call AddTgColumns("Specifier", "Near", tgQh)
            Call AddTgColumns("Contractor", "Near", tgQh)
            Call AddTgColumns("Customer", "Near", tgQh)

            Dim lblTemp As New System.Windows.Forms.Label '02-20-17
            lblTemp.Size = New Size(290, 60)
            lblTemp.Name = "Progress"
            Me.Controls.Add(lblTemp)
            lblTemp.Text = "Loading Specifiers:"
            lblTemp.TextAlign = ContentAlignment.MiddleCenter
            lblTemp.Location = New Point(100, 100)
            lblTemp.BackColor = Color.LightGreen
            lblTemp.BringToFront()
            Application.DoEvents()
            Dim Cnt As Integer = 1
            For Each dr As dsSaw8.QUTLU1Row In dsQutLU.QUTLU1.Rows
                Dim tmpspec As dsSaw8 = New dsSaw8 : tmpspec.EnforceConstraints = False
                Dim tmpda As MySqlDataAdapter = New MySqlDataAdapter
                tmpda.SelectCommand = New MySqlCommand("SELECT * FROM PROJECTCUST WHERE QUOTECODE = '" & SafeSQL(dr.QuoteCode) & "' and (projectcust.TypeC = 'C'  or projectcust.TypeC = 'X'  or projectcust.TypeC = 'A'  or projectcust.TypeC = 'E'  or projectcust.TypeC = 'L'  or projectcust.TypeC = 'S'  or projectcust.TypeC = 'T')", myConnection)
                tmpda.Fill(tmpspec, "projectcust")
                Dim Architect As String = ""
                Dim Engineer As String = ""
                Dim Contractor As String = ""
                Dim Specifier As String = ""
                Dim Customer As String = ""
                For Each drSpec As dsSaw8.projectcustRow In tmpspec.projectcust.Rows
                    lblTemp.Text = "Loading Specifiers: " & Cnt & "  of " & dsQutLU.QUTLU1.Rows.Count : Application.DoEvents() '02-21-17
                    If drSpec.Typec = "A" Then
                        If PrintFirmName = True Then
                            If Architect = "" Then Architect = drSpec.FirmName Else Architect += ", " & drSpec.FirmName
                        Else
                            If Architect = "" Then Architect = drSpec.NCode Else Architect += ", " & drSpec.NCode
                        End If
                    ElseIf drSpec.Typec = "E" Then
                        If PrintFirmName = True Then
                            If Engineer = "" Then Engineer = drSpec.FirmName Else Engineer += ", " & drSpec.FirmName
                        Else
                            If Engineer = "" Then Engineer = drSpec.NCode Else Engineer += ", " & drSpec.NCode
                        End If
                    ElseIf drSpec.Typec = "T" Then
                        If PrintFirmName = True Then
                            If Contractor = "" Then Contractor = drSpec.FirmName Else Contractor += ", " & drSpec.FirmName
                        Else
                            If Contractor = "" Then Contractor = drSpec.NCode Else Contractor += ", " & drSpec.NCode
                        End If
                    ElseIf drSpec.Typec = "S" Or drSpec.Typec = "X" Then
                        If PrintFirmName = True Then
                            If Specifier = "" Then Specifier = drSpec.FirmName Else Specifier += ", " & drSpec.FirmName
                        Else
                            If Specifier = "" Then Specifier = drSpec.NCode Else Specifier += ", " & drSpec.NCode
                        End If
                    ElseIf drSpec.Typec = "C" Then
                        If PrintFirmName = True Then
                            If Customer = "" Then Customer = drSpec.FirmName Else Customer += ", " & drSpec.FirmName
                        Else
                            If Customer = "" Then Customer = drSpec.NCode Else Customer += ", " & drSpec.NCode
                        End If
                    End If
                Next
                dr("Architect") = Architect
                dr("Engineer") = Engineer
                dr("Contractor") = Contractor
                dr("Specifier") = Specifier
                dr("Customer") = Customer
                Cnt += 1
            Next

            lblTemp.Visible = False : Application.DoEvents() '02-21-17
        End If




        If Me.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) And Me.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then '11-04-14 JTC
            If VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Then OrderBy = " ORDER BY Q.QuoteCode, SLSCode " Else OrderBy = " ORDER BY Q.JobName, SLSCode "
            Call AddTgColumns("Follow By", "Near", tgQh) '11-04-14 JTC 
            Call AddTgColumns("QSSpecCredit", "Near", tgQh) '11-04-14 JTC
            Me.tgQh.Splits(0).DisplayColumns("Follow By").Visible = True '11-04-14 JTC
            'Me.tgQh.Splits(0).DisplayColumns("QSSpecCredit").Visible = True '11-04-14 JTC
        End If
        'Dim RowCnt As Integer = 0
        If dsQutLU.QUTLU1.Rows.Count = 0 Then GoTo NoRecords '11-03-09
        If Me.pnlTypeOfRpt.Text = "Project Shortage Report" Then '05-08-10 
            If My.Computer.FileSystem.FileExists(UserDocDir & "QuoteShortage.DAT") Then Kill(UserDocDir & "QuoteShortage.DAT")
            FileClose(4) : FileOpen(4, UserDocDir & "QuoteShortage.DAT", OpenMode.Output)
        End If
        'Call AddTgColumns("SLSCode", "Center", tgQh) '03-08-13
        drQRow = dsQutLU.QUTLU1.Rows(0)
        'Debug.Print(dsQutLU.QUTLU1.Rows.Count)
        For Each drQRow In dsQutLU.QUTLU1.Rows 'dsQutLU
            If drQRow.RowState = DataRowState.Deleted Then Continue For ' GoTo 235 '06-19-08
            'Could Set A Default for followed by
            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And Trim(drQRow.FollowBy) = "" Then drQRow.Delete() : Continue For '03-06-12 No Blank FollowBy drQRow.FollowBy = "000" '03-03-12 

            If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And Me.chkSlsFromHeader.CheckState = CheckState.Checked Then '03-08-13
                If IsDBNull(drQRow("SLSCode")) Then drQRow("SLSCode") = "000"
                drQRow.SLSQ = drQRow("SLSCode")
            End If
            '11-04-14 JTC 
            If Me.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) And Me.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then '11-04-14 JTC
                If IsDBNull(drQRow("SLSCode")) Then drQRow("SLSCode") = "000"
                If IsDBNull(drQRow("FollowBy")) Then drQRow("FollowBy") = "000"
                If IsDBNull(drQRow("QSSpecCredit")) Then drQRow("QSSpecCredit") = "1"

                'Leave SLSQ the way it wasdrQRow.SLSQ = drQRow("SLSCode")
                drQRow.FollowBy = drQRow("SLSCode") & "*" & drQRow("QSSpecCredit")
                drQRow.Sell = drQRow("Sell") * drQRow("QSSpecCredit")
                drQRow.Comm = drQRow("Comm") * drQRow("QSSpecCredit")
                drQRow.Cost = drQRow("Cost") * drQRow("QSSpecCredit")
                Call AddTgColumns("Follow By", "Near", tgQh) '11-04-14 JTC 
                'Call AddTgColumns("QSSpecCredit", "Near", tgQh) '11-04-14 JTC
                Me.tgQh.Splits(0).DisplayColumns("Follow By").Visible = True '11-04-14 JTC
                ' Me.tgQh.Splits(0).DisplayColumns("QSSpecCredit").Visible = True '11-04-14 JTC
            End If
            'NoGoodIf Me.pnlTypeOfRpt.Text = "Quote Summary" Then
            '    If IsDBNull(drQRow("BidDate")) Then drQRow("BidDate") = "" '10-08-15 JTC Blank Bid Date
            '    If drQRow("BidDate") = "1900-01-01" Then drQRow("BidDate") = "" '10-08-15 JTC Blank Bid Date
            'End If
            Dim Hit As Short = 0 '06-20-10 

            If Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '12-01-09 
                Dim QutID As String, ProjID As String
                QutID = drQRow.QuoteID
                ProjID = drQRow.ProjectID
                Call DataBaseToScreen(Me, QutID, ProjID) 'to get other Tables for this quote
            End If

            'Debug.Print(drQRow.QuoteCode)
            Call SelectHit9500(Hit, multsrtrvs) '05-19-11 JTC Hit = 1 

            If Hit = 0 Then drQRow.Delete() : GoTo GetMExit Else RowCnt += 1 '05-17-10 GoTo GetMExit
            '03-03-12 If Trim(drQRow.JobName) = "" Then GoTo GetMExit '11-30-10 
            '05-08-10
            If Me.pnlTypeOfRpt.Text = "Project Shortage Report" Then '05-08-10 
                If Trim(drQRow.JobName) = "" Then GoTo GetMExit '11-30-10 
                PrintLine(4, VB.Left(drQRow.QuoteCode & "          ", 10) & "," & drQRow.JobName) '09-14-10) '05-16-10 PrintLine(1, "Hello", "World")   ' Separate strings with a Comma
            End If
GetMExit:
        Next
NoRecords:  '11-03-09
        If Me.pnlTypeOfRpt.Text = "Project Shortage Report" Then '03-15-10 
            FileClose(4) '05-08-10         UserDocDir & "QuoteShortage.DAT"
            FileClose() : Zarg = Zarg & " /QuoteShortage=YES" & "|" '04-26-10 & "/Nam=" & AGnam & "|" & "/User=" & UserID & "|" & "/Col=" & MBackCol & "|" '07-11-05 Shell to Main menu if not Authorized under Security System
            Dim Taskid As Single = Shell("VORT.EXE " & Zarg, 1)
            Me.Close() '
            End
        End If
        If RowCnt = -1 Then RowCnt = 0 '03-03-11
        If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Then US = "Only Quotes With Data in Followed By Field are included!" Else US = "" '03-22-12
        If RowCnt = 0 And Me.chkBidJobsOnly.Text = "Delivery Date Jobs Only" And Me.txtPrimarySortSeq.Text = "Forecasting" Then '01-27-14 JTC Forecast Zero records  
            MsgBox("Zero Records Selected." & vbCrLf & "Estimated Deliv Dates not yyyy-mm-dd" & vbCrLf & "Order Status not SUBMIT,GOT for Forecast" & vbCrLf & "Please check your selection again." & vbCrLf & US) '03-22-12
            Me.tabQrt.SelectedIndex = 1
            Call tabQRT_TabActivate(1) '02-01-09
            Me.Focus()
            GoTo EndExit '10-13-14 Exit Sub

        ElseIf RowCnt = 0 Then
            MsgBox("Zero Records Selected." & vbCrLf & "Please check your selection again." & vbCrLf & US) '03-22-12
            Me.tabQrt.SelectedIndex = 1
            Call tabQRT_TabActivate(1) '02-01-09
            Me.Focus()
            GoTo EndExit '10-13-14 Exit Sub
        Else
            If Me.chkBidJobsOnly.Text = "Delivery Date Jobs Only" And Me.txtPrimarySortSeq.Text = "Forecasting" Then '02-05-14 -27-14 JTC Forecast Zero records  
                MsgBox("Records Selected Before Filtering = " & RowCnt.ToString & vbCrLf & "The Lookup Grid shows total job dollars" & vbCrLf & "The report will only include Philips Brands and Philips dollars.") '04-26-12  
            ElseIf ForecastAllMfg = True Then 'BrandReportMfg = "PHIL" Or BrandReportMfg = "DAYB" Or BrandReportMfg = "DAY" Or SESCO = True Then ForecastAllMfg = False else  ForecastAllMfg = True '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
                MsgBox("Forecasting Records Selected Before Filtering = " & RowCnt.ToString & vbCrLf & "The Lookup Grid shows total job dollars" & vbCrLf & "The report will only include MFG/Codes/Brands dollars.") '04-26-12  
            Else
                MsgBox("Records Selected Before Filtering = " & RowCnt.ToString & vbCrLf & US & vbCrLf & "Click Run Report Button.") '01-27-15 JTC ) '04-26-12  
                'US = "Records Selected Before Filtering = " & RowCnt.ToString & vbCrLf & "Click Run Report Button." '01-27-15 JTC 
            End If
        End If
        Me.tgQh.Visible = True ' Quote Grid
        Me.tgQh.Dock = DockStyle.Fill '08-18-09 JH
        Me.tgr.Visible = False ' Prof Cust Grid
        Me.tgln.Visible = False '08-18-09 JH
        Me.tgQh.Rebind(True) '02-07-09

        If RealTgLookupExcel = False Then '07-31-14 JTC This is run After show hide and Loading Layout wipes out showhide column visible '07-31-14JTC Fix Header 
            If DIST Then 'VQrtLinesDistTGLayoutCurrent
                Call TrueGridLayoutFiles("Load", "Curr", "VQrtHdrDistTGLayout", tgQh) '09-15-10 
            Else
                Call TrueGridLayoutFiles("Load", "Curr", "VQrtHdrTGLayout", tgQh) '09-07-10 
            End If
        End If
        Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
        Me.txtSortSeqCriteria.Text = SelectionText '02-24-09
        Me.txtSortSeqV.Text = Me.txtSortSeq.Text ' txtPrimarySortSeq.Text & "/" & txtSecondarySort.Text
        If VB.Right(Me.txtSortSeq.Text, 1) = "/" Then Me.txtSortSeq.Text = Replace(Me.txtSortSeq.Text, "/", "")
        Me.tabQrt.SelectedIndex = 2
        Call tabQRT_TabActivate(2) '02-01-09
        Me.Focus()
        frmShowHideGrid.BringToFront() '02-10-09
        '07-26-12 
        If Me.pnlTypeOfRpt.Text = "Quote Summary" Then '07-27-12
            If txtPrimarySortSeq.Text = "Market Segment" Then If InStrColNam("MARKETSEGMENT") Then Call TgColumnsVisible("MarketSegment", "True", Me.tgQh) '07-27-12 f If not in Sq
            If txtPrimarySortSeq.Text = "Location" Then If InStrColNam("LOCATION") Then Call TgColumnsVisible("Location", "True", Me.tgQh) '07-27-12 f If not in Sq
        End If
        GoTo EndExit '10-13-14 Exit Sub 'Exit Sub'Exit Sub'Exit Sub'Exit Sub'Exit Sub'Exit Sub'Exit Sub'Exit Sub'Exit Sub'Exit Sub
        '**********************************************************************************************
QutLineHistoryRpt:
        If Me.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Or (Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (Me.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or Me.txtPrimarySortSeq.Text = "Salesman Follow-Up Report")) Then '09-19-12
            ' If Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '09-10-12 txtPrimarySortSeq.Text = "MFG Follow-Up Report"  GoTo QutLineHistoryRpt
            'Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" then If DefTypeOfJob = "Spec Credit" Then JT = "S"
            'Jaci Sql Select  Q.QuoteCode, QL.* from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  
            If (txtCustomerCodeLine.Text.Trim = "" Or txtCustomerCodeLine.Text = "ALL") Then  Else chkShowCustomers.CheckState = CheckState.Checked '02-25-14 JTC If Q Line Items if txtCustomerCodeLine then check Me.chkSlsFromHeader.CheckState = CheckState.Checked
            If txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Then '09-19-12
                '10-16-13 JTC added Q.BranchCode, Line items Sql
                '09-21-12 Add ProjectName strSql = "Select  QL.*, Q.QuoteCode, Q.SLSQ as VendorCode, Q.SLSQ as Vendor from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  " '09-19-12 
                '03-19-14 strSql = "Select  QL.*, Q.QuoteCode, P.ProjectName, Q.SLSQ as VendorCode, Q.SLSQ as Vendor from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID left join Project P on P.ProjectID = Q.ProjectID " '09-21-12 
                strSql = "SELECT QL.LineID, QL.QuoteID, QL.Qty, QL.Type, QL.MFG, QL.Description, Q.QuoteCode, Q.SLSQ ,Q.EntryDate, Q.City, Q.State, Q.JobName, Q.Status, QL.Comm, QL.Sell, QL.Paid FROM quote Q INNER JOIN quotelines QL ON Q.QuoteID = QL.QuoteID " '03-19-14 
            ElseIf txtPrimarySortSeq.Text = "Quote Summary" Then '09-21-12
                strSql = "Select  QL.*, Q.QuoteCode, P.ProjectName from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID left join Project P on P.ProjectID = Q.ProjectID " '09-21-12 
            ElseIf Me.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Then '09-23-12
                '08-02-13 JTC Add FirmName  * OL.Qty  O.MFG as TmpMFG,    '07-15-13 Chg O.CustCode to as CustCode
                '08-02-13strSql As String = "Select  OL.*, O.MFG as TmpMFG, O.CustCode as CustCode, O.CustName as Firm"
                If chkShowCustomers.Checked = True Then '08-02-13  
                    strSql = "Select  QL.*, Q.QuoteCode, Q.QUOTEID,Q.JobName, PC.NCode as NCode, PC.FirmName as FirmName from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID INNER JOIN PROJECTCUST AS PC ON Q.QuoteID = PC.QuoteID and PC.Typec = 'C' " '08-07-13 NG Limit 1
                    'Need to Eliminate Duplicate customers  and Quote.QuoteID = projectcust.QuoteID Limit 1 
                    ' (Select projectcust.FirmName from projectcust where projectcust.Typec = 'T' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Contractor, "
                    'testStop : strSql = "Select  QL.*, Q.QuoteCode, Q.QUOTEID as QUOTEID1,  (Select PC.NCode from projectcust PC where PC.Typec = 'C' and Q.QuoteID = PC.QuoteID Limit 1) as NCode  from QUOTE Q INNER JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID "
                    'Worked ? "Select  QL.*, Q.QuoteCode, Q.QUOTEID as QUOTEID1,  (Select PC.NCode from projectcust PC where PC.Typec = 'C' and Q.QuoteID = PC.QuoteID Limit 1) as NCode  from QUOTE Q INNER JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID where Q.EntryDate >= '2011-08-01' and Q.EntryDate <= '2013-08-31' and Q.TypeOfJob = 'Q'  and QL.Active = '1' and QL.LnCode <> 'NTE' and QL.LnCode <> 'NPE' and QL.LnCode <> 'SUB'  and QL.LnCode <> 'BTX' and QL.LnCode <> 'TXL'  and QL.LnCode <> 'TXS'  and QL.LnCode <> 'TXF'  and QL.LnCode <> 'TAX'  and QL.Description <> ''  and QL.MFG <> '' order by QL.MFG, QL.Description, QL.EntryDate DESC "
                    'No FirmName Etc
                ElseIf chkUseSpecifierCode.Checked = True Then
                    strSql = " Select  QL.*, Q.QuoteCode, Q.QUOTEID,Q.JobName, PC.NCode as NCode, PC.FirmName as FirmName from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID INNER JOIN PROJECTCUST AS PC ON Q.QuoteID = PC.QuoteID and (PC.Typec <> 'C' And PC.Typec <> 'M') " '08-02-13"
                Else
                    If chkShowCustomers.Checked = True Or chkUseSpecifierCode.Checked = True Then  '03-19-14
                        strSql = "Select  QL.*, Q.QuoteCode, Q.QUOTEID,Q.JobName, PC.NCode as NCode, PC.FirmName as FirmName  from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  INNER JOIN PROJECTCUST AS PC ON Q.QuoteID = PC.QuoteID and (PC.Typec <> 'C' And PC.Typec <> 'M') " '02-25-14 JTC Fix Quote Lines with One Customer code No PC.NCode in Where Clause Add PC.NCode as NCode, PC.FirmName as FirmName (add INNER JOIN PROJECTCUST AS PC ON Q.QuoteID = PC.QuoteID and (PC.Typec <> 'C' And PC.Typec <> 'M') )
                    Else
                        strSql = "Select  QL.*, Q.QuoteCode, Q.QUOTEID,Q.JobName  from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID " '03-19-14
                    End If
                End If
            Else      'Me.txtPrimarySortSeq.Text = "MFG Follow-Up Report"
                '09-21-12 add ProjectName strSql = "Select  QL.*, Q.QuoteCode from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  " '02-06-10
                '03-19-14strSql = "Select  QL.*, Q.QuoteCode, P.ProjectName from QUOTE Q LEFT JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID left join Project P on P.ProjectID = Q.ProjectID " '09-21-12 
                '"MFG Follow-Up Report"
                strSql = "SELECT QL.LineID, QL.QuoteID, QL.Qty, QL.Type, QL.MFG, QL.Description, Q.QuoteCode, Q.SLSQ, Q.EntryDate, Q.City, Q.State, Q.JobName, Q.Status, QL.Comm, QL.Sell, QL.Paid FROM quote Q INNER JOIN quotelines QL ON Q.QuoteID = QL.QuoteID " '03-19-14
            End If
            'If chkShowCustomers.CheckState = CheckState.Unchecked And chkUseSpecifierCode.CheckState = CheckState.Unchecked Then '02-25-14 JTC Eliminate Dups on PC table
            '    strSql = Replace(strSql, "INNER JOIN PROJECTCUST AS PC ON Q.QuoteID = PC.QuoteID and (PC.Typec <> 'C' And PC.Typec <> 'M')", " ")
            '    strSql = Replace(strSql, ", PC.NCode as NCode, PC.FirmName as FirmName", " ")
            'End If
            strSql = BuildSQLDetail("Q.", False, strSql) '08-31-09 
            'strSql = "Select  Q.QuoteCode, QL.* from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID   Select  Q.QuoteCode, QL.* from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  "
            If chkPrtNTElines.Checked = True Then '06-06-11 JTC On Product Lines Added chkPrtNTElines and chkHaveMFGCode.Checked = True Then '06-06-11 Must Have MFG Code
                'Print All Lines
            Else
                '06-06-11  and OL.Active = '1' and OL.LnCode <> 'NTE' and OL.LnCode <> 'NPN'  and OL.LnCode <> 'NPE' and OL.LnCode <> 'SUB'  and OL.LnCode <> 'BTX' and OL.LnCode <> 'TXL'  and OL.LnCode <> 'TXS'  and OL.LnCode <> 'TXF'  and OL.LnCode <> 'TAX'  and OL.Description <> ''
                strSql = strSql & " and QL.Active = '1' and QL.LnCode <> 'NTE' and QL.LnCode <> 'NPE' and QL.LnCode <> 'SUB'  and QL.LnCode <> 'BTX' and QL.LnCode <> 'TXL'  and QL.LnCode <> 'TXS'  and QL.LnCode <> 'TXF'  and QL.LnCode <> 'TAX'  and QL.Description <> '' " '06-06-11
            End If
            If chkHaveMFGCode.Checked = True Then '06-06-11 Must Have MFG Code
                strSql = strSql & " and QL.MFG <> '' " 'Must Have MFG Code 
            End If
            If txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Then '09-19-12
                strSql = strSql & "order by Q.SLSQ, Q.QUOTECODE, QL.Description, QL.EntryDate DESC " '09-20-12
            ElseIf txtPrimarySortSeq.Text = "Quote Summary" Then '09-21-12
                strSql = strSql & "order by P.ProjectName, QL.MFG, QL.Description " '09-21-12
            Else  'Me.txtPrimarySortSeq.Text = "MFG Follow-Up Report"
                strSql = strSql & "order by QL.MFG, QL.Description, QL.EntryDate DESC " '09-07-09  and Q.EndDate =< '20091231' " & "order by Q.QuoteCode " '01-25-09

            End If 'SortSeq = "QuoteLines.MFG, QuoteLines.Description, QuoteLines.EnterDate" '08-31-09 

            'VQUT Code %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            '10-15-13 Added Below
            If My.Computer.FileSystem.FileExists(UserSysDir & "VADMINNET.INI") = True Or SecurityLevel = "" Then '10-16-13 JTC Added No Security
                If SecurityLevel = "BRANCH" Or SecurityLevel = "REGIONAL" Or SecurityBrancheCodes.Trim.ToUpper <> "ALL" Then '10-16-13 JTC Added No Security Or SecurityBrancheCodes.Trim.ToUpper <> "ALL"
                    Dim STR1 As String = strSql.Substring(0, strSql.IndexOf("order by"))
                    Dim STR2 As String = " " & strSql.Substring(strSql.IndexOf("order by"))
                    Dim BC As String = ""
                    If SecurityBrancheCodes.Trim.ToUpper <> "ALL" Then
                        If SecurityBrancheCodes.Contains(",") = True Then  '10-16-13 JTC Added QL.BranchCode to Line Items
                            BC = " and ( QL.BranchCode = '" & SecurityBrancheCodes.Replace(",", "' or QL.BranchCode = '") & "' or QL.BranchCode = ''  )"
                        Else
                            BC = " and (QL.BranchCode = '" & SecurityBrancheCodes & "'" & " or QL.BranchCode = '')"
                        End If
                    End If
                    strSql = STR1 & BC & STR2
                End If
            End If

            Dim mysqlcmd As New MySqlCommand
            Dim daQuoteline As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            Dim dsquotelinelu As dsSaw8
            dsquotelinelu = New dsSaw8
            'strsql = "Select * from quotelines"
            daQuoteline = New MySqlDataAdapter
            daQuoteline.SelectCommand = New MySqlCommand(strSql, myConnection)
            Dim cbQutLin As MySql.Data.MySqlClient.MySqlCommandBuilder
            cbQutLin = New MySqlCommandBuilder(daQuoteline)
            'daQuoteline.Fill(dsQuoteRealLU, "quotelines")
            dsquotelinelu.EnforceConstraints = False '09-07-09 
            daQuoteline.Fill(dsquotelinelu, "quotelines")
            Me.QuoteLinesBindingSource.DataSource = dsquotelinelu.quotelines
            Me.QuoteLinesBindingSource.Filter = "MFG <> '' and Description <> ''" '09-07-09 
            '11-20-14 JTC Count short 1 RowCnt = dsquotelinelu.quotelines.Rows.Count - 1 '06-25-13
            RowCnt = dsquotelinelu.quotelines.Rows.Count '11-20-14 JTC Count short 1 RowCnt
            If RowCnt < 1 Then RowCnt = 0 : GoTo 5000 '06-25-13
            Dim Hit As Short = 1 '03-05-13 JTC Added Search Desc Logic TxtSearchString on Product Lines
            drQline = dsquotelinelu.quotelines.Rows(0)

            '03-19-14 moved to BuildSQLDetail and put in the sql
            'If Me.TxtSearchString.Text.Trim <> "" Then '03-19-14
            '    For Each drQline In dsquotelinelu.Tables("quotelines").Rows
            '        If drQline.RowState = DataRowState.Deleted Then Continue For
            '        Hit = 1
            '        If Trim(Me.TxtSearchString.Text) <> "" Then
            '            If InStr(Trim(drQline.Description), Trim(Me.TxtSearchString.Text)) Then Hit = 1 Else Hit = 0 ' GoTo SelExit9530 '05-05-10 
            '        End If
            '        'Debug.Print(drQline.Description)
            '        If Hit = 0 Then drQline.Delete() : RowCnt -= 1 : Continue For '10-16-13 JTC Added -RowCntElse RowCnt += 1 '05-17-10 GoTo GetMExit
            '    Next
            'End If
            'RowCnt = dsquotelinelu.quotelines.Count  ' Tables("quotelines").Rows
5000:
            If RowCnt = 0 Then
                MsgBox("Zero Records Selected." & vbCrLf & "Please check your selection again." & vbCrLf & US) '03-22-12
                Me.tabQrt.SelectedIndex = 1
                Call tabQRT_TabActivate(1) '02-01-09
                Me.Focus()
                GoTo EndExit '10-13-14 Exit Sub
            Else
                If US.Trim <> "" Then
                    MsgBox("Records Selected After Filtering = " & RowCnt.ToString & vbCrLf & US & vbCrLf & "Click Run Report Button.") '01-27-15 JTC )
                Else
                    MsgBox("Records Selected Before Filtering = " & RowCnt.ToString & vbCrLf & US & vbCrLf & "Click Run Report Button.") '01-27-15 JTC) '04-26-12 
                    '& vbCrLf & "Click Run Report Button." )'01-27-15 JTC 
                End If
            End If

            'If Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '09-21-12And Me.txtPrimarySortSeq.Text = "Quote Summary" Then '09-21-12  GoTo QutLineHistoryRpt
            '    Call AddTgColumns("ProjectName", "Left", tgln)
            'End If

            '03-19-14Call AddTgColumns("UM", "Center", tgln) '09-05-10 NearFarCenter =
            Me.tgQh.Visible = False ' Quote Grid
            Me.tgr.Visible = False ' Proj Cust Grid
            Me.tgln.Dock = DockStyle.Fill
            Me.tgln.Visible = True
            Me.tgln.Rebind(True)
            'Me.tgln.Rebind(False)
            'If RealTgLookupExcel = False Then '07-31-14 JTC This is run After show hide and Loading Layout wipes out showhide column visible '07-31-14JTC Fix Header 
            If DIST Then 'VQrtLinesDistTGLayoutCurrent
                If RealTgLookupExcel = False Then Call TrueGridLayoutFiles("Load", "Curr", "VQrtLinesDistTGLayout", tgln) '07-31-14 JTC 09-15-10 
            Else
                If Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '03-19-14
                    If RealTgLookupExcel = False Then Call TrueGridLayoutFiles("Load", "Curr", "VQrtSpecCredit", tgln) '07-31-14 JTC 09-07-10 
                Else
                    If RealTgLookupExcel = False Then Call TrueGridLayoutFiles("Load", "Curr", "VQrtLinesTGLayout", tgln) '07-31-14 JTC 09-07-10 
                End If
            End If
            '03-19-14
            'If Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '09-21-12 And Me.txtPrimarySortSeq.Text = "Quote Summary" Then '09-21-12  GoTo QutLineHistoryRpt
            '    Me.tgln.Rebind(False) '09-23-12 Me.tgln.Rebind(False) 
            '    Call AddTgColumns("ProjectName", "Left", tgln)
            'End If
            '03-19-14 Call AddTgColumns("UM", "Center", tgln) '09-05-10 NearFarCenter =

            If chkUseSpecifierCode.Checked = True Or chkShowCustomers.Checked = True Then '03-19-14 (if)
                If InStrColNam("FIRMNAME") Then Call AddTgColumns("FirmName", "Near", tgln) '08-05-13'UseFirmName = True Else Call TgColumnsVisible("FirmName", "False", tgInvoiceMaster)
                If InStrColNam("NCODE") Then Call AddTgColumns("NCode", "Near", tgln) '08-05-13'UseFirmName = True Else Call TgColumnsVisible("FirmName", "False", tgInvoiceMaster)
            End If


            '03-19-14 If DIST Then Call AddTgColumns("Ext Sell", "Far", tgln)
            '03-19-14 If DIST Then Call AddTgColumns("Ext Cost", "Far", tgln) '08-06-13
            '03-19-14
            'For I = 0 To tgln.Splits(0).DisplayColumns.Count - 1 '08-06-13
            '    Dim ColName As String = tgln.Splits(0).DisplayColumns(I).Name '08-06-13 .ToUpper
            '    If ColName = "QUOTEID1" Or ColName = "LineID" Or ColName = "QuoteID" Or ColName = "ProdID" Or ColName = "LPProdID" Then
            '        tgln.Splits(0).DisplayColumns(I).Visible = False
            '    End If
            '    If DIST = False And ColName = "Cost" Then
            '        tgln.Splits(0).DisplayColumns(I).DataColumn.Caption = "Book" '08-03-13
            '    End If
            '    If DIST And ColName = "Comm" Then
            '        tgln.Splits(0).DisplayColumns(I).DataColumn.Caption = "Margin" '08-03-13
            '    End If
            '    If DIST And ColName = "BkComm" Then
            '        tgln.Splits(0).DisplayColumns(I).Visible = False '08-06-13
            '    End If
            'Next I
            '03-19-14
            '09-21-12 ******************************************************************
            '03-19-14
            'If (Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (Me.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or Me.txtPrimarySortSeq.Text = "Salesman Follow-Up Report")) Then '09-19-12
            '    For I = 0 To tgln.Splits(0).DisplayColumns.Count - 1 '09-21-12
            '        'If tgln.Splits(0).DisplayColumns(I).Visible = False Then Continue For
            '        ' If (tgln.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
            '        Dim ColName As String = tgln.Splits(0).DisplayColumns(I).Name.ToUpper
            '        If ColName = "QTY" Or ColName = "TYPE" Or ColName = "MFG" Or ColName = "DESCRIPTION" Or ColName = "COMM" Or ColName = "PAID" Or ColName = "PROJECTNAME" Or ColName = "LNCODE" Then
            '            tgln.Splits(0).DisplayColumns(I).Visible = True
            '        Else
            '            tgln.Splits(0).DisplayColumns(I).Visible = False
            '        End If
            '    Next I
            'End If
            '03-19-14
            '***********************************************************************************
            Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
            Me.txtSortSeqCriteria.Text = SelectionText
            Me.txtSortSeqV.Text = Me.txtSortSeq.Text ' txtPrimarySortSeq.Text & "/" & txtSecondarySort.Text
            If VB.Right(Me.txtSortSeq.Text, 1) = "/" Then Me.txtSortSeq.Text = Replace(Me.txtSortSeq.Text, "/", "")
            Me.tabQrt.SelectedIndex = 2
            Call tabQRT_TabActivate(2) '02-01-09
            Me.Focus()             ' IfDebugOn then Sto
            frmShowHideGrid.BringToFront() '02-10-09
        End If
EndExit:  '10-13-14 JTC
        Me.cmdok1.Enabled = True '10-13-14 JTC Fix Dbl Click on Reports
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow 'HourGlass
        System.Windows.Forms.Application.DoEvents() '10-14-14 JTC
    End Sub

    Private Sub cmdReportProjShortage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error Resume Next
        Me.cboSortRealization.Visible = False '01-18-12
        Me.pnlTypeOfRpt.Text = "Project Shortage"
        Me.chkIncludeCommDolPer.Visible = False '12-09-02 WNA
        Call cmdReportQuote_Click(cmdReportQuote, New System.EventArgs())
    End Sub
    Private Sub cmdReportQuote_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReportProjShortage.Click, cmdReportQuote.Click
        On Error Resume Next
        ExcelQuoteFU = False '04-28-14 JTC
        SESCO = False '04-28-14 JTC
        Me.ChkSpecifiersCustInCols.Visible = True '06-28-18
        Me.chkBrandReport.Visible = False '05-16-13 
        Me.chkShowLatestCust.Visible = False '03-24-13
        Me.chkShowLatestCust.CheckState = CheckState.Unchecked '03-24-13
        Me.ChkTotalsOnly.Checked = CheckState.Unchecked '07-26-12
        Me.chkDetailTotal.Checked = CheckState.Unchecked '07-26-12
        Me.cboSortRealization.Visible = False '01-18-12
        Me.cboSortPrimarySeq.Visible = True '01-30-12
        Me.fraFinishReports.Visible = True '09-23-12
        Me.ChkSpecifiers.CheckState = CheckState.Unchecked '02-11-12 Reused for Realization If Me.pnlTypeOfRpt.Text = "Realization" Or Me.chkNotes.CheckState = CheckState.Checked Then
        If Me.chkSlsFromHeader.Text = "Use Quote SLS 1 Split for Salesman" Then Me.chkSlsFromHeader.Text = "Use Salesman From Quote Header on Report" '03-08-13
        '08-28-12 JTC Shut off Realization choices when you click on summary report
        For I = 0 To 10
            cboSortRealization.SetItemCheckState(I, CheckState.Unchecked) '01-21-12
        Next
        'cboSortRealization.SetItemCheckState(11, CheckState.Unchecked) '02-26-12 SESCO
        RealALL = False '10-13-14 JTC Fix Check box
        RealCustomer = False
        RealManufacturer = False
        RealQuoteTOOther = False '01-31-12
        RealSLSCustomer = False
        RealArchitect = False
        RealEngineer = False
        RealLtgDesigner = False
        RealSpecifier = False
        RealContractor = False
        RealOther = False ' 08-28-12 
        RealCustomerOnly = False '03-11-14
        lblTypeCustomer.Visible = False : cboTypeCustomer.Visible = False '03-11-14
        '02-11-12 JTC Add Notes to Realization
        ''Me.ChkSpecifiers.Text = "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
        ''02-11-12 Use ChkSpecifiers.Text = "Sort Report by Descending Dollar 
        'Me.ChkSpecifiers.Text = "Sort Report by Descending Dollar" '02-11-12 " "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
        'Me.ChkSpecifiers.Visible = True '02-11-12 
        If eventSender Is cmdReportProjShortage Then '03-15-10 
            '.cmdReportProjShortage.Click '
            Me.pnlTypeOfRpt.Text = "Project Shortage Report" '03-15-10
            Me.Text = "Project Shortage Report" & "  " & AGnam & "  UserID =" & UserID '09-04-10
        Else
            'If Me.pnlTypeOfRpt.Text = "" Then '10-23-02 WNA
            Me.pnlTypeOfRpt.Text = "Quote Summary" '02-24-09 
            Me.Text = "Quote Summary" & "  " & AGnam & "  UserID =" & UserID '09-04-10'06-19-10 
        End If
        fraReportCmdSelection.Visible = True
        'Me.fraLines.Enabled = True
        ' Else
        'Me.fraLines.Enabled = False '05-17-02 WNA
        ' End If
        'fraSortPrimarySeq.Visible = True

        Me.cboLinesInclude.Visible = False '12-01-09
        Me.pnlTypeOfRpt.Visible = True
        Me.fraReportCmdSelection.Visible = False

        Me.fraSortPrimarySeq.Visible = True
        Me.pnlPrimarySortSeq.Visible = True
        Me.txtPrimarySortSeq.Text = ""
        Me.txtPrimarySortSeq.Visible = True
        Me.pnlQutRealCode.Visible = False
        Me.txtQutRealCode.Visible = False
        Me.chkSlsFromHeader.Visible = False
        Me.chkSlsFromHeader.Enabled = False
        Me.pnlQuoteToSls.Visible = True '01-27-09
        Me.txtQuoteToSls.Visible = True
        If MFG = 0 And DAYB = 0 And DIST = 0 Then '12-09-02 WNAMe.pnlTypeOfRpt.Text = "Project Shortage Report" '03-15-10
            If Me.pnlTypeOfRpt.Text <> "Project Shortage Report" Then Me.chkIncludeCommDolPer.Visible = True
        End If
        Call FillPrimarySortCombo()
        Me.cboSortPrimarySeq.Focus()
    End Sub
    Private Sub cmdReportRealization_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReportRealization.Click
        On Error Resume Next
        Me.Text = "Realization Reports" & "  " & AGnam & "  UserID =" & UserID '09-04-10'06-14-10 
        Me.pnlTypeOfRpt.Text = "Realization" 'frmQuoteRpt.txtPrimarySortSeq.Text = "Descending Sales Dollars" ' Name Code" '02-11-12
        'Me.txtPrimarySortSeq.Text = Trim(Me.cboSortPrimarySeq.Text)
        ' Me.cboSortPrimarySeq.Text = "" '03-20-13 Can't Make my mind Up
        'For Each obj As Object In Me.cboSortRealization.Items '03-22-13 
        '    'Me.cboSortRealization.SetItemCheckState = CheckState.Unchecked
        '    'cboSortRealization.SetItemCheckState(3, CheckState.Unchecked)
        '    obj.SetItemCheckState = CheckState.Unchecked
        'Next
        Me.lblJobName.Text = "Job Name Search String" '06-22-12 
        Me.ChkSpecifiersCustInCols.Visible = False '06-28-18
        Me.fraFinishReports.Visible = True '09-23-12
        Me.cboLinesInclude.Visible = False '12-01-09
        Me.pnlTypeOfRpt.Visible = True
        Me.fraReportCmdSelection.Visible = False
        Me.fraSortPrimarySeq.Visible = True
        Me.cboSortRealization.Visible = True '01-18-12
        ' Me.cboSortRealization.BringToFront() '09-21-12
        ''Me.ChkSpecifiers.Text = "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
        ''02-11-12 Use ChkSpecifiers.Text = "Sort Report by Descending Dollar 
        'Me.ChkSpecifiers.Text = "Sort Report by Descending Dollar" '02-11-12 " "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
        'Me.ChkSpecifiers.Visible = True '02-11-12 
        Me.pnlPrimarySortSeq.Visible = True
        Me.txtPrimarySortSeq.Text = ""
        Me.txtPrimarySortSeq.Visible = True
        Me.pnlQutRealCode.Text = "Code"
        Me.txtQutRealCode.Text = "ALL"
        Me.pnlQutRealCode.Visible = True
        Me.txtQutRealCode.Visible = True
        Me.chkSlsFromHeader.Visible = True '02-01-99 WNA
        Me.chkSlsFromHeader.Enabled = True '02-01-99 WNA
        Me.pnlQuoteToSls.Visible = True '07-16-02 WNA
        Me.txtQuoteToSls.Visible = True '07-16-02 WNA
        'Me.fraLines.Enabled = False
        If MFG = 0 And DAYB = 0 Then '02-21-03 WNA
            Me.chkIncludeCommDolPer.Visible = True
        End If
        '04-22-15 JTC Use cboSortRealization not cboSortPrimarySeq for "SESCO Job List Report" or Excel Quote FollowUp Me.cboSortRealization.Items(12) = "SESCO Job List Report" or "Excel Quote FollowUp"
        If My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") Then '02-25-12
            ' Dim LastItem As Int16 = Me.cboSortRealization.Items.Count - 1 '04-22-15  'dim LastItem As Int16 = Me.cboSortPrimarySeq.Items.Count - 1 '03-29-12 
            If Me.cboSortRealization.Items(12) = "SESCO Job List Report" Then ' 'If Me.cboSortPrimarySeq.Items(LastItem).text = "SESCO Job List Report" Then
                SESCO = True
                ExcelQuoteFU = False '04-28-15 JTC
            Else
                Me.cboSortRealization.Items(12) = "Excel Quote FollowUp" '04-22-15 JTC 
                SESCO = False
            End If
        Else
            '04-22-15 JTC Wrong cbo cboSortRealization.Items.Remove("SESCO Job List Report") '02-26-12(11) = False '02-26-12
            Me.cboSortRealization.Items(12) = "Excel Quote FollowUp" '04-22-15" Then
            SESCO = False    'Me.cboSortPrimarySeq.Items.Add("Excel Quote FollowUp") '04-22-15 JTC Chg "SESCO Job List Report" to "Excel Quote FollowUp" Realization
        End If
        Me.cboSortPrimarySeq.Visible = False
    End Sub
    Private Sub cmdResetDefaults1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdResetDefaults1.Click, cmdResetDefaults2.Click '04-20-11
        On Error Resume Next
        Me.txtStartEntry.Text = "ALL"
        Me.txtEndEntry.Text = "ALL"
        Me.txtStartBid.Text = "ALL"
        Me.txtEndBid.Text = "ALL"
        Me.txtStartQuoteAmt.Text = "0"
        Me.txtEndQuoteAmt.Text = "999999999" '03-24-08 JTC Added 9 "999,999,999"
        Me.txtStatus.Text = "ALL"
        Me.txtSalesman.Text = "ALL"
        Me.txtRetrieval.Text = "ALL"
        Me.txtSpecifierCode.Text = "ALL"
        Me.txtLastChgBy.Text = "ALL"
        Me.txtState.Text = "ALL"
        Me.txtCity.Text = "ALL"
        Me.txtMktSegment.Text = "ALL"
        Me.txtJobNameSS.Text = ""
        Me.txtQuoteToSls.Text = "ALL" '07-16-02 WNA
        Me.chkBlankBidDates.Visible = False '02-19-04 WNA
        If DIST <> 1 Then
            Me.txtCSR.Text = "ALL"
            Me.txtSelectCode.Text = "ALL"
            Me.cboLotUnit.Text = "ALL"
            Me.cbospeccross.Text = "ALL"
            Me.cboStockJob.Text = "ALL"
            Me.txtSlsSplit.Text = "ALL"
        Else
            Me.txtCSRofCust.Text = "ALL" '09-25-07 JH
        End If

        System.Windows.Forms.Application.DoEvents()
    End Sub
    Private Sub cmdSecondarySeqCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSecondarySeqCancel.Click
        '07-24-10 Case 1 'cancel back to primary sort seq
        Me.pnlSecondarySort.Visible = False
        Me.txtSecondarySort.Visible = False
        Me.fraSortSecondarySeq.Visible = False '07-24-10
        Me.fraSortPrimarySeq.Visible = True
        Me.cboSortPrimarySeq.Focus()
        SortNeeded = ""

    End Sub
    Private Sub cmdSecondarySeqContinue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSecondarySeqContinue.Click, cboSortSecondarySeq.DoubleClick '09-24-10 , cmdSecondarySeqCancel.Click '01-30-09
        Dim index As Short = 0 '01-30-09cmdSortSecondarySeq.GetIndex(eventSender)
        Dim Enable As Short
        Dim Resp As Short
        On Error Resume Next
        Me.fraSortSecondarySeq.Visible = False
        Me.chkSlsFromHeader.Text = "Use Salesman From Quote Header on Report" ' "Use Quote SLS 1 Split for Salesman" '03-08-13
        Me.chkSlsFromHeader.CheckState = CheckState.Unchecked '03-08-13
        Select Case index
            Case 0 'continue run report
                If Trim(Me.cboSortSecondarySeq.Text) = "" Then
                    Resp = MsgBox("You must select a Secondary Sort Sequence before you Continue", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Secondary Sort Continue")
                    Exit Sub
                End If
                Me.txtSecondarySort.Text = Trim(Me.cboSortSecondarySeq.Text)
                'Call Main("SECONDSEL")

                Call SetSecondarySortValues()
                If Me.pnlTypeOfRpt.Text = "Project Shortage" Then '10-23-02 WNA
                    Enable = 0
                Else
                    Enable = 1
                End If
                Call EnableOrDisable2(Enable)
                If RealCustomerOnly = True Then '03-11-14
                    cboTypeCustomer.Visible = True : lblTypeCustomer.Visible = True
                Else
                    cboTypeCustomer.Visible = False : lblTypeCustomer.Visible = False
                End If
                '07-26-12 Me.chkDetailTotal.Visible = False

            Case 1 'cancel back to primary sort seq
                Me.pnlSecondarySort.Visible = False
                Me.txtSecondarySort.Visible = False
                'Me.fraRptSel.Visible = False
                Me.fraSortPrimarySeq.Visible = True
                Me.cboSortPrimarySeq.Focus()
                SortNeeded = ""
        End Select
        If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Then '03-03-12 Chg SpecifierCode to FollowedBY
            Me.pnlSpecifierCode.Text = "FollowedBy Code"
            Me.pnlSpecifierCode.Visible = True
            Me.txtSpecifierCode.Text = "ALL"
            Me.txtSpecifierCode.Visible = True
            Me.ChkSpecifiers.Checked = True
            Me.chkCustomerBreakdown.Enabled = True '03-06-12
            Me.chkCustomerBreakdown.Checked = True '03-06-12 
            Me.chkNotes.Checked = True '03-06-12 
            Me.chkBlankLine.Checked = True '03-09-12 
        Else
            Me.pnlSpecifierCode.Text = "Specifier Code"
            Me.pnlSpecifierCode.Visible = False
            Me.txtSpecifierCode.Visible = False

        End If



        If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Then '03-08-13
            Me.chkSlsFromHeader.Visible = True
            Me.chkSlsFromHeader.Enabled = True
            Me.chkSlsFromHeader.CheckState = CheckState.Unchecked
            Me.chkSlsFromHeader.Text = "Use Quote SLS 1 Split for Salesman" '03-08-13 "Use Salesman From Quote Header on Report"
            If Me.pnlTypeOfRpt.Text = "Realization" Then '01-06-14 
                'VB6 DOESN'T USE SLS 1 ON THE QUOTE TO - BY SLS, CUST.  REALIZATION - CUSTOMER THEN SLS GOES THOUGH HERE TOO
                'JH - I THINK WARD WANTED THE SLS FROM THE SPLIT ON THIS REPORT BUT I CAN'T FIND ANY NOTES ABOUT IT
                Me.chkSlsFromHeader.Text = "Print SLS 1 Split for SLSCode on Report"
            End If
            'Debug.Print(Me.chkSlsFromHeader.Text)
        End If
        '11-04-14 JTCStop
        If Me.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) Then '11-04-14 JTC
            'Me.cboSortSecondarySeq.Items.Clear() '11-04-14 JTC  Me.cboSortSecondarySeq.Text = "None"       Me.cboSortSecondarySeq.Items.Add("None")
            If Me.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then
                Me.chkSlsFromHeader.Visible = True
                Me.chkSlsFromHeader.Enabled = True
                Me.chkSlsFromHeader.CheckState = CheckState.Checked
                Me.chkSlsFromHeader.Text = "Use Quote SLS 1-4 Splits for Salesman"
            End If
        End If
        '01-28-09 Set Tab 
        Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
        Me.txtSortSeqCriteria.Text = SelectionText '02-24-09
        Me.txtSortSeqV.Text = Me.txtSortSeq.Text 'txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
        If VB.Right(Me.txtSortSeq.Text, 1) = "/" Then Me.txtSortSeq.Text = Replace(Me.txtSortSeq.Text, "/", "")
        If ExcelQuoteFU = True Then '04-28-15 JTC
            Resp = MsgBox("Excel Report to show who you Quoted." & vbCrLf & "Shows Specifiers and Contractors on the Job." & vbCrLf & "Selects on Both Entry Dates and Bid Dates", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Excel Quote FollowUp Report") '04-29-25 JTC
        End If
        Me.tabQrt.SelectedIndex = 1
        Call tabQRT_TabActivate(1)


    End Sub

    Private Sub lstSortThirdSeq_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboSortSecondarySeq.SelectedIndexChanged
        txtSortSecondarySeq.Text = cboSortSecondarySeq.Text
    End Sub

    'Private Sub DTPicker1EndBid_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DTPicker1EndBid.Change
    '    Dim A As String
    '    A = DTPicker1EndBid.Value 'MM/dd/yyyy
    '    txtEndBid.Text = VB6.Format(A, "m/d/yyyy") '01-27-01 Left$(A$, 2) & Mid$(A$, 4, 2) & Right$(A$, 2) '06-13-00
    '    'EndBidDate$ = Format$(A$, "yyyymmdd")    '01-24-01 Right$(A$, 4) + Left$(A$, 2) & Mid$(A$, 4, 2)
    '    Me.txtEndBid.Focus()
    'End Sub
    'Private Sub DTPicker1EndBid_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    '    Call DTPicker1EndBid_Change(DTPicker1EndBid, New System.EventArgs())
    'End Sub
    'Private Sub DTPicker1EndBid_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent)
    '    Dim KeyAscii As Object
    '    If eventArgs.keyCode = 13 Then
    '        System.Windows.Forms.SendKeys.Send("{tab}")
    '        KeyAscii = 0
    '    End If '06-13-00
    '    If eventArgs.keyCode = 189 Then
    '        System.Windows.Forms.SendKeys.Send("{RIGHT}")
    '        KeyAscii = 0
    '    End If '06-13-00
    'End Sub
    Private Sub DTPicker1EndEntry_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim A As String
        A = DTPicker1EndEntry.Value 'MM/dd/yyyy
        txtEndEntry.Text = VB6.Format(A, "mmddyy") '01-24-01 Left$(A$, 2) & Mid$(A$, 4, 2) & Right$(A$, 2) '06-13-00
        'EndEntryDate$ = Format$(A$, "yyyymmdd")    '01-24-01 Right$(A$, 4) + Left$(A$, 2) & Mid$(A$, 4, 2)
        Me.txtEndEntry.Focus()
    End Sub
    Private Sub DTPicker1EndEntry_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Call DTPicker1EndEntry_Change(DTPicker1EndEntry, New System.EventArgs())
    End Sub
    'Private Sub DTPicker1EndEntry_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent)
    '    Dim KeyAscii As Object
    '    If eventArgs.keyCode = 13 Then
    '        System.Windows.Forms.SendKeys.Send("{tab}")
    '        KeyAscii = 0
    '    End If '06-13-00
    '    If eventArgs.keyCode = 189 Then
    '        System.Windows.Forms.SendKeys.Send("{RIGHT}")
    '        KeyAscii = 0
    '    End If '06-13-00
    'End Sub
    Private Sub DTPicker1StartBid_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim A As String
        A = DTPicker1StartBid.Value 'MM/dd/yyyy
        txtStartBid.Text = VB6.Format(A, "mmddyy") '01-24-01  Left$(A$, 2) & Mid$(A$, 4, 2) & Right$(A$, 2) '06-13-00
        'StartBidDate$ = Format$(A$, "yyyymmdd")    '01-24-01 Right$(A$, 4) + Left$(A$, 2) & Mid$(A$, 4, 2)
        Me.txtStartBid.Focus()

    End Sub
    Private Sub DTPicker1StartBid_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Call DTPicker1StartBid_Change(DTPicker1StartBid, New System.EventArgs())
    End Sub

    Sub trackbar_Scroll(ByVal sender As Object, ByVal e As EventArgs)
        Dim val As Integer = trackbar.Value
        cmdPercent.Text = val.ToString & "%"
        zoom(val / 100, tgQh) '11-12-09 
        zoom(val / 100, tgln)
        zoom(val / 100, tgr)
    End Sub
    Private Sub zoom(ByVal pcnt As Single, ByVal mytg As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        Try
            If pcnt = 0 Then Exit Sub
            ' adjust row height
            mytg.RowHeight = CInt(CSng(_rowHeight) * pcnt)
            ' and recordselector width
            mytg.RecordSelectorWidth = CInt(CSng(_recSelWidth) * pcnt)
            ' adjust font sizes.  Normal is the root style so changing its sizes adjust all other styles
            mytg.Styles("Normal").Font = New Font(mytg.Styles("Normal").Font.FontFamily, _fontSize * pcnt, FontStyle.Bold)
            ' now adjust the column widths
            Dim i As Integer
            Select Case mytg.Name.ToUpper
                Case "TGQH" ' 05-17-10 Upper
                    For i = 0 To (mytg.Splits(0).DisplayColumns.Count) - 1 : mytg.Splits(0).DisplayColumns(i).Width = CInt(CSng(_colWidthstgQh(i)) * pcnt) : Next i
                Case "TG"
                    For i = 0 To (mytg.Splits(0).DisplayColumns.Count) - 1 : mytg.Splits(0).DisplayColumns(i).Width = CInt(CSng(_colWidthstg(i)) * pcnt) : Next i
                Case "TGR"
                    For i = 0 To (mytg.Splits(0).DisplayColumns.Count) - 1 : mytg.Splits(0).DisplayColumns(i).Width = CInt(CSng(_colWidthstgr(i)) * pcnt) : Next i
            End Select
        Catch ex As Exception
            MessageBox.Show("Error in Zoom VQRT" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "Zoom VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12 MsgBox(exc.Message, , "Zoom")
        End Try
    End Sub 'zoom
    Private Sub frmQuoteRpt_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'DsSaw8.SpecRegFollowUp' table. You can move, or remove it, as needed.
        'Me.SpecRegFollowUpTableAdapter.Fill(Me.DsSaw8.SpecRegFollowUp)
        'TODO: This line of code loads data into the 'DsSaw8.SpecRegFollowUp' table. You can move, or remove it, as needed.
        ' Me.SpecRegFollowUpTableAdapter.Fill(Me.DsSaw8.SpecRegFollowUp)
        'TODO: This line of code loads data into the 'DsSaw8.QuoteRealNDUL' table. You can move, or remove it, as needed.
        'Me.QuoteRealNDULTableAdapter.Fill(Me.DsSaw8.QuoteRealNDUL)

        '11-01-10 JTC Add Imports System.Threading and Imports System.Globalization
        '11-01-10 JTC Set Culture Here and B/4  InitializeComponent()"fr-CA")French Canada ' not (France= "fr-FR")
        '08-19-11System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-us") '11-01-10 JTC Culture
        '08-19-11System.Threading.Thread.CurrentThread.CurrentUICulture = New System.Globalization.CultureInfo("en-us") ''11-01-10 JTC Culture
        '??DTPickerStartEntry.CultureInfo = New System.Globalization.CultureInfo("en-US")
        '??DTPickerStartEntry.CultureName = New System.Globalization.CultureInfo("en-US")
        'ALLS(-Allscape)
        'ALKC(-Alkco)
        'ARDE -Ardee Lighting
        'BRON(-Bronzelite)
        'CAPR(-Capri)
        'CHLO -Chloride				Core Brand             
        'CMTP -CMT-Composite Poles
        'COLO -Color Kinetics		Core Brand
        'CSA  -Crescent Stonco
        'DAYB -Day-Brite			Core Brand
        'ELCN –Electo/Connect
        'EXCL(-Exceline)
        'FORE(-Forecast)
        'GARD -Gardco/EMCO			Core Brand
        'EMCO(-EMCO)
        'GUTH(-Guth)
        'HADC -Hadco				Core Brand
        'HANO -Hanover Lantern
        'HILT(-High - Lites)
        'LAM(-LAM)
        'LEDA -Ledalite				Core Brand
        'LITG(-Lightguard)
        'LOL  -Lightolier			Core Brand
        'LUMC -Lumec				Core Brand
        'MORL(-Morlite)
        'OMEG(-Omega)
        'LOLC -Philips (Lightolier)Controls
        'PHRD -Philips Roadway
        'QUAL -Quality Lighting
        'SHAK –Shakespeare Composit Structures
        'SELE(-Selecon)
        'STON –Stonco				Core Brand
        'SPRT(-Sportlite)
        'STRA -Strand Lighting
        'THMI -Thomas Lighting
        'TRAN -Translite Systems
        'USSM –USS Manufacturing Inc.
        'WIDE(-WideLite)
        Dim Taskid As Integer 'Form Load Settings FormLoad Load Form LoadForm tabQrt_DrawItem
        'tabQrt_DrawItem on after testingDebugOn
        ProgramDateToolStripMenuItem.Text = "Program Date = 06-28-18" ' sub TODOList  Turn On Private Sub tabQrt_DrawItem
        Dim A As String
        Me.mnuTime.Text = "Date: " & Now.ToString("MM/dd/yyyy") & "  Time: " & Now.ToString("HH:mm:ss") '11-04-10 09
        Me.Show()
        Me.tabQrt.DrawMode = TabDrawMode.OwnerDrawFixed
        ' save some state information of TG
        _rowHeight = Me.tgQh.RowHeight
        _recSelWidth = Me.tgQh.RecordSelectorWidth
        _fontSize = Me.tgQh.Styles("Normal").Font.Size
        trackbar.SmallChange = 1
        trackbar.LargeChange = 10
        trackbar.Minimum = 30
        trackbar.Maximum = 300
        trackbar.Value = 100
        AddHandler trackbar.Scroll, AddressOf trackbar_Scroll
        RibbonFontComboBox2.Text = Me.tgQh.Font.Name
        FontSizeComboBox.Text = 10
        '11-12-09 
        Dim dc As C1.Win.C1TrueDBGrid.C1DisplayColumn
        For Each dc In tgQh.Splits(0).DisplayColumns
            _colWidthstgQh.Add(CSng(dc.Width))
        Next dc
        For Each dc In tgr.Splits(0).DisplayColumns
            _colWidthstgr.Add(CSng(dc.Width))
        Next dc
        For Each dc In tgln.Splits(0).DisplayColumns
            _colWidthstg.Add(CSng(dc.Width))
        Next dc

        fraQuoteLineReports.Visible = False
        '_cmdSortPrimarySeq_0, _cmdSortPrimarySeq_1, _cmdSortSecondarySeq_0, _cmdSortSecondarySeq_1, _fraOutputOptions_0, _fraOutputOptions_1, _optOutputOptions_0, _optOutputOptions_1, _optOutputOptions_2, _optOutputOptions_3, _optOutputOptions_4, _optOutputOptions_5, C1Ribbon1, C1StatusBar1, cboLotUnit, cboSortPrimarySeq, cboSortSecondarySeq, cboSpecCross, cboStockJob, chkBlankBidDates, chkBranchReport, chkCustomerBreakdown, chkDetailTotal, chkExcludeDuplicates, chkExportAllExcel, chkIncludeCommDolPer, chkMfgBreakdown, chkSalesmanPerPage, chkSlsFromHeader, cmdCancel, cmdCancelRpt, cmdFmtOK, cmdOK, cmdReportProjShortage, cmdReportQuote, cmdReportRealization, cmdResetDefaults, DTPicker1EndBid, DTPicker1EndEntry, DTPicker1StartBid, DTPickerStartEntry, frabreakdown, fraDisplaySortSeq, fraFinishReports, fraLines, fraRecordSelect, fraReport, fraReportCmdSelection, fraRptSel, fraSelectDate, fraSortPrimarySeq, fraSortSecondarySeq, Label2, lblEndBid, lblEndEntry, lblEndQuote, lblJobName, lblRetrieval, lblSalesman,
        'lblStartBid, lblStartEntry, lblStartQuote, lblStatus, MainMenu1, optOne, optSelectNewRecords, optThree, optTwo, optUsePrevSelection, pnlCity, pnlCSR, pnlCSRdist, PnlLastChgBy, pnlLotUnit, pnlMktSeg, pnlPrimarySortSeq, pnlQuoteToSls, pnlQutRealCode, pnlSecondarySort, pnlSlsSplits, pnlSltCode, pnlSpecCross, pnlState, pnlStkJob, pnlTypeOfRpt, SplitContainer1, SSPanel3, SSPanel4, TabPage0, TabPage1, TabPage2, tabQrt, tg, txtCity, txtCSR, txtCSRDist, txtEndBid, txtEndEntry, txtEndQuote, txtJobNameSS, txtLastChgBy, txtMktSegment, txtPrimarySortSeq, txtQuoteToSls, txtQutRealCode, txtRetrieval, txtSalesman, txtSecondarySort, txtSelectCode, txtSlsSplit, txtSpecifierCode, txtStartBid, txtStartEntry, txtStartQuote, txtState, txtStatus, 

        Zarg = VB.Command()
        AGnam = ReturnZarg("/Nam=") '10-08-09  Call function 'Zarg$ = "/Nam=|/User=JKH|/Col=|/ExpediteQAdd|07-0007|"     '10-08-09 B = InStr(Zarg, "/Nam=") : If B Then E = InStr(B, Zarg, "|") : If E Then Agnam = Mid(Zarg, B + 5, E - B - 5)
        UserID = ReturnZarg("/User=") '10-08-09
        '10-23-14 JTC UserPath = My.Computer.FileSystem.CurrentDirectory.ToString '04-17-10 & "\" '06-02-09
        UserPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) '10-23-14 JTC UserPath
        If My.User.Name = "JTCVist3-PC\JTCVist3" Or My.User.Name = "JTC7-PC\JTC7" Then '07-06-12
            If UserPath = "C:\VBNET\VQRT\bin" Then UserPath = "C:\Users\Public\SAW8_5.5\" : UserID = "JTC"
            If UserPath = "C:\VBNET\VQRT\bin" Then UserPath = "C:\SAW8SESCOATL\" : UserID = "JTC"
            If UserPath = "C:\VBNET\VQRT\bin" Then UserPath = "C:\SAW855\" : UserID = "JTC"
            If UserPath = "C:\VBNET\VQRT\bin" Then UserPath = "C:\SAW8Folders\SAW8_Rep\" : UserID = "JKH"
            If UserPath = "C:\VBNET\VMENU\bin\Debug" Then UserPath = "C:\SAW8_GEXPROREGIONAL\" : DebugOn = True : UserID = "JTC" '11-02-10 
            'Stop : UserPath = "C:\SAW8_5.5\" : DebugOn = True '11-02-10 
            'Stop : UserPath = "C:\SAW8_Dist\" : DebugOn = True '11-02-10 
        ElseIf My.User.Name = "Jaci2010\jacitemp" Then
            If UserPath = "C:\VBNET\VQRT\bin" Then UserPath = "C:\SAW8Folders\SAW8_Rep\" : DebugOn = True : UserID = "JKH"
            If UserPath = "C:\VBNET\VQRT\bin" Then UserPath = "C:\SAW8\" : DebugOn = True : UserID = "JKH"
        End If
        'test System.IO.Directory.SetCurrentDirectory("C:\SAW7\") 'Set Current Directory
        'test A = System.IO.Path.GetFullPath("VMENU.EXE") 'test
        'test System.IO.Directory.SetCurrentDirectory(UserPath & "VMENU.EXE") 'Set Current Directory
        If UserPath.EndsWith("\") = False Then UserPath = UserPath & "\" '04-17-10
        System.IO.Directory.SetCurrentDirectory(UserPath) '12-07-10 Set Current Directory
        Call Me.FormSetting("Load") '10-31-10 
        rbnHelpAboutDirectory.Text = "Directory = " & UserPath '08-31-10
        C1Ribbon1.Minimized = True '11-01-10 Moved dowm
        'DebugOn = True : UserPath = "C:\SAW8\" '10-12-09 'Set to "C:\SAW8" for testing *****************************************
        ''If DebugOn = True Then UserPath = "C:\USERS\PUBLIC\SAW8\" 'testing only 06-02-09***************************
        '09-03-10 Moved this up for GetUserInfo

        If My.Computer.FileSystem.FileExists(UserPath & "ServerPath.ini") = True Then '01-03-17
            Dim StrLen As Integer
            Dim iSize As Integer
            Dim DtaFile As String = ""
            iSize = 255 : DtaFile = Space$(iSize)
            StrLen = GetPrivateProfileString("PATH", "PATH", "none", DtaFile, iSize, UserPath & "ServerPath.ini")
            If DtaFile.Substring(0, Math.Min(DtaFile.Length, StrLen)) <> "none" Then
                Serverpath = VB.Left(DtaFile, StrLen)
            Else
                Serverpath = ""
            End If
            If Serverpath <> "" Then If My.Computer.FileSystem.DirectoryExists(Serverpath) = False Then Serverpath = ""
            If Serverpath <> "" Then If Serverpath.EndsWith("\") = False Then Serverpath += "\"
        End If

        Call OpenSQL(myConnection) '02-19-10 UserPath & & FormSetting("Load") must be set before OpenSQL Call OpenSQL(myConnection)*
        '09-03-10 Moved up
        UserDocDir = Environment.GetFolderPath(Environment.SpecialFolder.Personal) '04-17-10 & "\" '05-27-09 "C:\Users\JTCVist3\Documents\"
        If UserDocDir.EndsWith("\") = False Then UserDocDir = UserDocDir & "\" '04-17-10 
        If Trim(UserID) = "" Then GetUserInfo() '04-30-10 JH
        If Trim(UserID) = "" Then UserID = "ZXX" '12-18-09
        UserSysDir = UserPath & "USER\SYS\" '06-02-09 "C:\SAW8\USER\SYS\" 'Create SYS dir
        UserPathImages = UserPath & "\IMAGES\" '02-15-10
        '03-19-14 Call FillDataSet(Me)
        Me.chkBlankBidDates.CheckState = CheckState.Unchecked '02-03-12
        Me.ChkCheckBidDates.CheckState = CheckState.Unchecked '02-03-12
        Me.chkBlankBidDates.Visible = False '02-03-12
        Me.DTPicker1StartBid.Enabled = False '02-04-12
        Me.DTPicker1EndBid.Enabled = False '02-04-12
        '08-26-09 moved up Call FormSetting("Load") '12-05-08 in 
        Me.optUnitOrExtended_Extd.Checked = True '09-10-09 Extended Prices
        UserDir = UserPath & "USER\" & UserID & "\" '06-02-09 "C:\SAW8\USER\ABC\"
        'Public UserPathHelp As String '08-27-10 = UserPath & "HELP\" '08-27-10
        UserPathHelp = UserPath & "HELP\" '08-27-10 JTC Added to Make SAWDirectories 
        Me.HelpProvider1.HelpNamespace = UserPathHelp & "HLPQuoteReports.pdf" '09-09-12 
        HelpProvider1.SetShowHelp(Me, True) '08-27-10
        Me.Text = "Quote Reports " & AGnam & "  UserID =" & UserID ' 04-26-10 
        Zarg = "/Nam=" & AGnam & "|/User=" & UserID & "|" '04-26-10
        'Setup Directories
        '"C:\Users\MyName\Documents\" for Docs & temp storage on Each PC
        'Create "C:\SAW8\USER\SYS" Should be done on installation
        If My.Computer.FileSystem.DirectoryExists(UserSysDir) Then
        Else
            My.Computer.FileSystem.CreateDirectory(UserSysDir)
        End If
        If My.Computer.FileSystem.DirectoryExists(UserDir) Then
        Else ' Create "C:\SAW8\USER\WES\"
            My.Computer.FileSystem.CreateDirectory(UserDir)
            'copy files from \SYS to \New User
            My.Computer.FileSystem.CopyDirectory(UserSysDir, UserDir)
        End If

        'Drop Down for Type of Job - 11-13-09 JH
        Dim tmpJT As New DataSet : Dim dtJT As New DataTable
        dtJT.Columns.Add("Code", GetType(String)) : dtJT.Columns.Add("Description", GetType(String))
        tmpJT.Tables.Add(dtJT)
        tmpJT.Tables(0).Rows.Add(New Object() {"Q", "Quotes Only"})
        tmpJT.Tables(0).Rows.Add(New Object() {"S", "Out of Terr Spec Credit"})
        tmpJT.Tables(0).Rows.Add(New Object() {"P", "Planned Project"})
        tmpJT.Tables(0).Rows.Add(New Object() {"T", "PDF Submittal"})
        tmpJT.Tables(0).Rows.Add(New Object() {"O", "Other"})
        tmpJT.Tables(0).Rows.Add(New Object() {"A", "A=All Jobs"}) '11-13-09 
        cboTypeofJob.DataSource = tmpJT.Tables(0)
        cboTypeofJob.Splits(0).DisplayColumns(0).Width = 30
        cboTypeofJob.Splits(0).DisplayColumns(1).Width = 50
        cboTypeofJob.LimitToList = True

        Dim DefLnCodes As New DataSet : Dim tbDefLnCodes As New DataTable '03-11-14
        DefLnCodes.Tables.Add(tbDefLnCodes)
        DefLnCodes.Tables(0).Columns.Add("Code") : DefLnCodes.Tables(0).Columns.Add("Description")
        DefLnCodes.Tables(0).Rows.Add("ALL", "All Types")
        DefLnCodes.Tables(0).Rows.Add("A", "Architect")
        DefLnCodes.Tables(0).Rows.Add("E", "Engineer")
        DefLnCodes.Tables(0).Rows.Add("L", "Ltg Designer")  '01-12-12
        DefLnCodes.Tables(0).Rows.Add("S", "Specifier")
        DefLnCodes.Tables(0).Rows.Add("T", "Contractor")
        DefLnCodes.Tables(0).Rows.Add("C", "Distributor/Customer") '01-09-14 (only want to see the distributors (marked in NA) on the specifier report
        DefLnCodes.Tables(0).Rows.Add("X", "Other") '09-20-10
        cboTypeCustomer.DataSource = DefLnCodes.Tables(0).Copy
        cboTypeCustomer.LimitToList = False
        cboTypeCustomer.DropDownWidth = 200 '01-09-14 JH
        cboTypeCustomer.Splits(0).DisplayColumns(0).Width = 50 '01-09-14 JH
        cboTypeCustomer.Text = "ALL"
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
        If UCase(My.Computer.FileSystem.FileExists(UserPath & "DISTQUOTE.DAT")) Then '12-04-09
            DIST = True 'Public DIST As Boolean'07-11-09 
        End If
        Call Main_Renamed("INIT")
        If DIST Then
            Call SetupSelectCriteria()
            Me.cmdReportTerrSpecCredit.Visible = False '12-04-09 
            Me.cmdReportProjShortage.Visible = False '10-23-02 WNA
            Me.chkIncludeCommDolPer.Text = "Include Cost and Margin" '02-21-03 WNA
            Me.chkBrandReport.Visible = False '05-16-13 
        End If
        If MFG Or DAYB Then
            Me.cmdReportProjShortage.Visible = False '10-23-02 WNA
            Me.chkIncludeCommDolPer.Visible = False '12-09-02 WNA
        End If
        'Dim TcName As String = ""
        'For I As Int16 = 0 To Me.tg.Splits(0).DisplayColumns.Count - 1
        '    TcName = TcName & Me.tg.Splits(0).DisplayColumns(I).Name & ","
        'Next
        '03-19-14 Me.tgln.Splits(0).DisplayColumns("UM").Visible = True '09-04 -10
        '08-13-13 JTC Testing OnlyStop :  A = "FirstLogOn" : Call SecuritySubNew(A) '08-13-13 JTC Testing Only
        A = "FirstLogOn" : Call SecuritySub(A) '08-05-05 WNA
        If A = "SecurityNG" Or A = "SecurityHD" Then GoTo 900 '04-19-06 WNA
        '08-13-13 JTC Testing OnlyStop :   A = "ReportQut" : Call SecuritySubNew(A) '08-13-13 JTC
        A = "ReportQut" : Call SecuritySub(A) '08-05-05 WNA
        If A = "SecurityNG" Or A = "SecurityHD" Then GoTo 900 '04-19-06 WNA
        _fdBranchCode.Text = "ALL" '10-15-13 If no secority
        If My.Computer.FileSystem.FileExists(UserSysDir & "VADMINNET.INI") = True Then
            A = "SecurityGroup" : Call SecuritySub(A)
            If A = "SecurityNG" Or A = "SecurityHD" Then
                GoTo 900 '04-19-06 WNA
            End If
            'SecurityBrancheCodes = GetBranchCode(UserID)
            'Admin = SYSTEM, BRANCH, REGIONAL,
            '"SYSTEM" Then SecurityAdministrator = True = All Branches IE: Ignore BRANCH
            '"BRANCH" If BRANCH, GetBranchCode(UserID) 
            'REGIONAL, Then SecurityBrancheCodes = dsadmin.adminuser.Rows(0).Item("AdminBranches").ToString
            '10-15-13 JTC Put Branch Code or Codes in Text Box
            If SecurityBrancheCodes.Trim = "" Then SecurityBrancheCodes = "ALL"
            'If SecurityBrancheCodes <> "" And SecurityBrancheCodes.ToUpper <> "ALL" Then '10-15-13 JTC Put Branch Code or Codes in Text Box
            : _fdBranchCode.Enabled = True '10-28-13
            If SecurityLevel = "SYSTEM" Or SecurityLevel = "" Then '10-16-13 JTC If Admin allow any Branch Codes
                _fdBranchCode.Text = SecurityBrancheCodes
            ElseIf SecurityLevel = "BRANCH" Then
                _fdBranchCode.Text = SecurityBrancheCodes '10-28-13 
                : _fdBranchCode.Enabled = False
            Else
                _fdBranchCode.LimitToList = False : _fdBranchCode.Enabled = False
                _fdBranchCode.Text = SecurityBrancheCodes
            End If
        Else
            'No Security
            If SecurityBrancheCodes.Trim = "" Then SecurityBrancheCodes = "ALL"
            _fdBranchCode.Text = "ALL" '10-16-13 
        End If
        'From Vquote ''''''''''''''''''''''''''''''''''''''''''''''''
        'If CheckForFile(UserSysDir & "VADMINNET.INI") = True Then '08-30-13
        '    SecuritySubNew(A)
        '    Exit Sub
        'End If
        'Check for branch settings turned on.
        'Set which branch records a user can view
        'LUBranchCode =   'set as public
        'SYSTEM = View all   REGIONAL = View listed branches from vadminuser table   BRANCH = Only branch listed in N&A   SLSPERSON = Only records listed as SLSMAN on
        If My.Computer.FileSystem.FileExists(UserSysDir & "VADMINNET.INI") = True Then
            If SecurityBrancheCodes <> "" Then
                If My.Computer.FileSystem.FileExists(UserSysDir & "VADMINNET.INI") = True Then
                    'BranchSql is Global and holds and 'EST' and 'WST' etc
                    BranchSql = GetBranchCodes("quotelines", True)
                End If
            End If
        End If
        '10-20-13 JTC $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'FileName2 = Dir$(UserPath & "BrandReport*.*") '"BrandReport-*.DAT")
        'If My.Computer.FileSystem.FileExists(UserPath & FileName2) = False Then '11-07-13 JTC Fix Duplicate mnuBrandMfgChg_Click BrandReportMfg = "XXXX"
        '    My.Computer.FileSystem.CopyFile(UserPath & FileName2, UserPath & "BrandReport-" & BrandReportMfg & ".DAT")
        '    My.Computer.FileSystem.DeleteFile(UserPath & FileName2)
        'End If
        For Each fName As String In Directory.GetFiles(UserPath, "BrandReport*.*")
            fName = Replace(fName, UserPath, "")
            If fName.ToUpper.StartsWith("BRANDREPORT-") Then '02-24-14 JTC Move "BRANDREPORT-" From UserPath to UserSys
                My.Computer.FileSystem.CopyFile(UserPath & fName, UserSysDir & fName, True)
                My.Computer.FileSystem.DeleteFile(UserPath & fName)
            End If
        Next
        Dim FileName2 As String '03-12-13 
        '02-24-14 JTC Move "BRANDREPORT-" From UserPath to UserSysDir
        FileName2 = Dir$(UserSysDir & "BrandReport-*.*") '"BrandReport-*.DAT")
        If FileName2.ToUpper.StartsWith("BRANDREPORT-") Then '03-12-13
            BrandReportMfg = Mid(FileName2, 13, 4)
            F = InStr(BrandReportMfg, ".") 'In case less than 4 
            If F <> 0 Then BrandReportMfg = VB.Left(BrandReportMfg, F - 1) ' If F <> 0 Then strSql = VB.Left(strSql, F - 1)
            If BrandReportMfg.Length > 2 Then 'If mnuBrandReport.Text = "Brand Reporting - Off" Then
                'BrandReportMfg = "PHIL" : Me.ChkBrandMfgRpt.CheckState = CheckState.Checked '10-20-13
                '02-24-14 JTC Brand Reporting is Off Until Forecasting Report
                Me.mnuBrandReport.Text = "Brand Reporting - Off" '11-06-13
                Me.mnuBrandMfgChg.Text = "Brand Mfg Code - " & BrandReportMfg '11-06-13 XXXX Brand Mfg Code - XXXX
                Me.chkBrandReport.Text = "Brand Reporting - " & BrandReportMfg '11-06-13 XXXX Brand Mfg Code - XXXX"
                Me.chkBrandReport.Visible = True '03-14-13
                Me.chkBrandReport.CheckState = CheckState.Unchecked '02-24-14 
                'Me.txtQutRealCode.Text = PhilBrands 
            End If
        Else
            Me.chkBrandReport.Visible = False '01-14-14 JTC Off if No "BRANDREPORT-PHIL" file
            Me.mnuBrandReport.Visible = False '01-14-14 JTC Text = "Brand Reporting - On" '11-06-13
            Me.mnuBrandMfgChg.Visible = False '01-14-14 JTCText = "Brand Mfg Code - " & BrandReportMfg '11-06-13 XXXX Brand Mfg Code - XXXX
        End If
        BrandReportMfg = BrandReportMfg.ToUpper '02-04-15 JTC Must be Upper case  BrandReportMfg = BrandReportMfg.ToUpper
        '10-20-13 JTC $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        '04-20-10Me.DTPicker1EndEntry.Value = CDate("12/31/2010")
        '10-29-10 JTC Chged Back 10-15-10 jh '***********************************************************************
        Me.DTPickerStartEntry.Value = CDate(Format(Now, "yyyy-") & Format(Now, "MM-") & "01") '04-20-10 Start Of Month 'Dim NewDate As New DateTime ' 'NewDate = CDate(Format(Now, "yyyy-") & Format(Now, "MM-") & "01") '04-20-10 Start Of Month CDate(Format(Now, "MM") & "01" & Format(Now, "yyyy"))   'NewDate = NewDate.AddMonths(1) : NewDate = NewDate.AddDays(-1)
        Me.DTPicker1EndEntry.Value = GetLastDayInMonth(CDate(Format(Now, "MM-") & "01-" & Format(Now, "yyyy"))) '04-20-10 JTC NewDate ' CDate("12/31/" & Format(Now, "yyyy")) '02-02-10
        Me.DTPicker1StartBid.Value = CDate(Format(Now, "yyyy-") & Format(Now, "MM-") & "01") '06-05-12 CDate("01/01/1900") 
        Me.DTPicker1EndBid.Value = GetLastDayInMonth(CDate(Format(Now, "MM-") & "01-" & Format(Now, "yyyy"))) '06-05-12 CDate("12/31/2020")
        '06-05-12 Dim NewDate As Date = Now '02-03-12
        '06-05-12 NewDate = NewDate.AddMonths(6) 'AddYears : NewDate = NewDate.AddDays(-1)
        '06-05-12 Me.DTPicker1EndBid.Value = NewDate '02-03-12 'A = Format(Now.YearaDDyear(1).Year + 1, "yyyy")
        '10-29-10 JTC Test 10-15-10 jh************************************************************
        Call SetupSpec("LOAD", False)  '02-13-13 WNA
        Me.cbospeccross.ClearItems()
        Me.cbospeccross.DataBindings.Clear()
        Me.cbospeccross.DataSource = dsSCross.Tables(0)
        Me.cbospeccross.Text = "ALL"
        '"Project Shortage","Product Sales History","Realization","Terr Spec Credit Report","Quote Summary","Planned Projects"
        '12-18-09 QuoteTo Realization "VQrtQuoteToTGLayout", tgr,,"VQrtHdrTGLayout", tgQh,,"VQrtLinesTGLayout", tgln
        '01-22-10"VQrtHdrDistTGLayout"=tgQh,"VQrtHdrTGLayout"=tgQh,"VQrtQuoteToTGLayout"=tgr-Realization,"VQrtLinesTGLayout"=tgln
        '03-19-14 - JH - ALL GRIDS SHOULD BE ABLE TO CREATE THE XML FILES FOR THEM.  IF THEY HAVE THE OLD ONES THEY MAY GET ERRORS - TRY TO RESET
        If DIST Then
            If File.Exists(UserSysDir & "VQrtHdrDistTGLayoutOriginal.xml") = True Then
                Dim FDate As Date = IO.File.GetLastWriteTime(UserSysDir & "VQrtHdrDistTGLayoutOriginal.xml")
                If FDate.ToShortDateString < #3/27/2014# Then
                    If File.Exists(UserSysDir & "VQrtHdrDistTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtHdrDistTGLayoutOriginal.xml")
                    If File.Exists(UserSysDir & "VQrtLinesDistTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtLinesDistTGLayoutOriginal.xml")
                    If File.Exists(UserSysDir & "VQrtQuoteToDistTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtQuoteToDistTGLayoutOriginal.xml")
                    If File.Exists(UserDir & "VQrtHdrDistTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtHdrDistTGLayoutCurrent.xml")
                    If File.Exists(UserDir & "VQrtLineItemsDistShowHide.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLineItemsDistShowHide.xml")
                    If File.Exists(UserDir & "VQrtLinesDistTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLinesDistTGLayoutCurrent.xml")
                    If File.Exists(UserDir & "VQrtQuoteToDistTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtQuoteToDistTGLayoutCurrent.xml")
                    If File.Exists(UserDir & "VQrtRealQTOShowHideDistPrint.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtRealQTOShowHideDistPrint.xml")
                    If File.Exists(UserDir & "VQrtShowHideDistPrtHdr.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtShowHideDistPrtHdr.xml")
                End If
            End If
        Else
            If File.Exists(UserSysDir & "VQrtHdrTGLayoutOriginal.xml") = True Then
                Dim FDate As Date = IO.File.GetLastWriteTime(UserSysDir & "VQrtHdrTGLayoutOriginal.xml")
                If FDate.ToShortDateString < #3/27/2014# Then
                    If File.Exists(UserSysDir & "VQrtHdrTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtHdrTGLayoutOriginal.xml")
                    If File.Exists(UserSysDir & "VQrtLinesTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtLinesTGLayoutOriginal.xml")
                    If File.Exists(UserSysDir & "VQrtQuoteToTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtQuoteToTGLayoutOriginal.xml")
                    If File.Exists(UserSysDir & "VQrtSpecCreditOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtSpecCreditOriginal.xml")
                    If File.Exists(UserDir & "VQrtHdrTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtHdrTGLayoutCurrent.xml")
                    If File.Exists(UserDir & "VQrtLineItemsRepShowHide.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLineItemsRepShowHide.xml")
                    If File.Exists(UserDir & "VQrtLinesTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLinesTGLayoutCurrent.xml")
                    If File.Exists(UserDir & "VQrtQuoteToTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtQuoteToTGLayoutCurrent.xml")
                    If File.Exists(UserDir & "VQrtRealQTOShowHideRepPrint.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtRealQTOShowHideRepPrint.xml")
                    If File.Exists(UserDir & "VQrtShowHideRepPrtHdr.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtShowHideRepPrtHdr.xml")
                    If File.Exists(UserDir & "VQrtSpecCreditCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtSpecCreditCurrent.xml")
                    If File.Exists(UserDir & "VQrtQuoteToTGLayoutShowAllCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtQuoteToTGLayoutShowAllCurrent.xml") '07-22-14
                    If File.Exists(UserDir & "VQrtRealQTOShowHideRepPrintShowAll.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtRealQTOShowHideRepPrintShowAll.xml") '07-22-14
                End If
            End If
        End If

        If DIST Then '01-19-10
            If File.Exists(UserSysDir & "VQrtHdrDistTGLayoutOriginal.xml") = False Then '03-19-14
                'IF THE FILE ISN'T THERE WE CAN CREATE IT FROM THE tgQhDIST grid, it's setup with the correct captions
                Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtHdrDistTGLayout", tgQhDIST) '12-18-09 
            Else
                Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtHdrDistTGLayout", tgQh) '12-18-09 
            End If
        Else
            Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtHdrTGLayout", tgQh) '12-18-09 
        End If

        If DIST = False Then
            Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtSpecCredit", tgSpecReg) '03-19-14
        End If


        'Me.tgQh.Caption = "Quote Master Grid - Click Run Report to see final report selection" '01-31-12
        '06-06-11 ****************************************
        'Dim ShowAllQuoteHeader As String = "" '06-06-11
        'If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then ShowAllQuoteHeader = "ShowAll" '06-06-11 = "Show All Quote Header Fields" Then '06-06-11 "Add Cust QuoteTo Breakdown to Report"
        If DIST Then '09-05-10 
            If File.Exists(UserSysDir & "VQrtQuoteToDistTGLayoutOriginal.xml") = False Then '03-19-14
                'IF THE FILE ISN'T THERE WE CAN CREATE IT FROM THE tgrDIST grid, it's setup with the correct captions
                Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtQuoteToDistTGLayout", tgrDIST) '06-06-11
            Else
                Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtQuoteToDistTGLayout", tgr) '06-06-11
            End If

        Else
            Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtQuoteToTGLayout", tgr) '06-06-11
        End If
        'Me.tgr.Splits(0).DisplayColumns("Cost").Visible = False '12-08-09
        If DIST Then '05-06-10 
            If File.Exists(UserSysDir & "VQrtLinesDistTGLayoutOriginal.xml") = False Then '03-19-14
                'IF THE FILE ISN'T THERE WE CAN CREATE IT FROM THE tglnDIST grid, it's setup with the correct captions
                Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtLinesDistTGLayout", tglnDIST) '03-19-14
            Else
                Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtLinesDistTGLayout", tgln)
            End If
        Else
            Call TrueGridLayoutFiles("FormLoad", "Curr", "VQrtLinesTGLayout", tgln)
        End If

        '03-12-14 JH - these two are for the checkbox for Show Customer Quoted and Use Specifier for Cust Code - both do not work fix later
        '03-12-14 If InStrColNam("FIRMNAME") Then Call AddTgColumns("FirmName", "Near", tgln) '08-05-13
        '03-12-14 If InStrColNam("NCODE") Then Call AddTgColumns("NCode", "Near", tgln) '08-05-13

        '02-03-12 Moved Grid Changes to here 
        If DIST Then 'Files Setup is DIST 
            ' tgln = Quote Lines
            '03-19-14
            'Me.tgln.Splits(0).DisplayColumns("UOverage").Visible = False ' tgln= Quote Lines
            'Me.tgln.Splits(0).DisplayColumns("BkComm").Visible = False    ''08-06-13 tgln= Quote Lines
            'Me.tgln.Splits(0).DisplayColumns("Comm").Visible = False '08-06-13
            'Me.tgln.Splits(0).DisplayColumns("ProdID").Visible = False '07-26-09
            ''06-24-13 JTC Delete BkSell Me.tgln Me.tgln.Splits(0).DisplayColumns("BKSell").Visible = False '12-18-09
            'Me.tgln.Splits(0).DisplayColumns("Ext Sell").Visible = True '05-08-10
            ''tgln(e.Row, "UM").ToString()
            ''Quote Header
            'Me.tgQh.Splits(0).DisplayColumns("SpecCredit").Visible = False '12-03-09
            'Me.tgQh.Splits(0).DisplayColumns("Comm-$").Visible = False '12-04-09
            ''02-06-12 Me.tgQh.Columns("Comm-%").NumberFormat = "n2"
            ''02-06-12 Me.tgQh.Splits(0).DisplayColumns("Comm-%").DataColumn.Caption = "Margin"
            ''tgr Realization Quote to
            'Me.tgr.Splits(0).DisplayColumns("Comm-$").Visible = False '12-04-09
            'Me.tgr.Splits(0).DisplayColumns("LPComm").Visible = False '12-08-09
            'Me.tgr.Splits(0).DisplayColumns("Overage").Visible = False '12-08-09
            'Me.tgr.Splits(0).DisplayColumns("Margin").Visible = True '09-04-10 

        Else  'Rep 

            'tgln= Quote Lines 'Me.tgln.Splits(0).DisplayColumns("Margin").DataColumn.Caption = "Comm-%" ' tgln = Quote Lines
            '03-12-14
            'Me.tgln.Splits(0).DisplayColumns("Cost").Visible = True : tgln.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" '06-24-13 False '12-17-09 DataColumn.Caption = "BKSell"    ' tgln = Quote Lines
            'Me.tgln.Splits(0).DisplayColumns("Comm").Visible = True
            'Me.tgln.Splits(0).DisplayColumns("UOverage").Visible = True '12-03-09 tgln = Quote Lines
            'Me.tgln.Splits(0).DisplayColumns("BkComm").Visible = True '08-06-13   ' tgln = Quote Lines
            'Me.tgln.Splits(0).DisplayColumns("Ext Sell").Visible = True '05-08-10 
            'Me.tgln.Splits(0).DisplayColumns("LPCost").Visible = False '05-08-10
            ''tgQh = Quote Header Table  Rep
            'For I As Int16 = 0 To Me.tgQh.Splits(0).DisplayColumns.Count - 1
            '    If Me.tgQh.Splits(0).DisplayColumns(I).Name = "LPMarg" Then
            '        Me.tgQh.Splits(0).DisplayColumns("LPMarg").DataColumn.Caption = "LPComm" ' tgQh = Quote Tabl
            '    ElseIf Me.tgQh.Splits(0).DisplayColumns(I).Name = "Margin" Then
            '        Me.tgQh.Splits(0).DisplayColumns("Margin").DataColumn.Caption = "Comm-%" '05-07-10
            '        Exit For
            '    End If
            'Next
            'Me.tgQh.Columns("Comm-%").NumberFormat = "n2"
            'Me.tgQh.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" '06-24-13 JTC Rep = Display and Print Book 
            'Me.tgQh.Splits(0).DisplayColumns("Book").Visible = True '06-24-13
            'Me.tgQh.Splits(0).DisplayColumns("LPCost").Visible = False '12-04-09
            ''tgr Realization Quote to
            'Me.tgr.Splits(0).DisplayColumns("Comm-$").Visible = True '10-04-10
            'Me.tgr.Splits(0).DisplayColumns("LPComm").Visible = True '12-08-09
            'Me.tgr.Splits(0).DisplayColumns("Overage").Visible = True '12-08-09
            ''02-03-12 Me.tgr.Splits(0).DisplayColumns("Margin").Visible = False '11-04-10  
            ''11-27-12 Me.tgr.Splits(0).DisplayColumns("Cost").Visible = False '12-08-09
            'Me.tgr.Splits(0).DisplayColumns("Cost").DataColumn.Caption = "Book" '11-27-12 JTC Cost to Book show Book
            '03-12-14

        End If

        Me.cboTypeofJob.Text = "Q" '04-26-12

        GoTo 999
900:
        If splashscreen.Visible = True Then splashscreen.Hide() '04-18-11
        MsgBox("See Administrator for Security Rights to use this Quote Reports Program" & vbCrLf & " User=" & UserID & "  Exiting to the Main Menu")
        FileClose() '04-26-10  : Zarg = " /Nam=" & AGnam & "|" & "/User=" & UserID & "|" & "/Col=" & MBackCol & "|" '07-11-05 Shell to Main menu if not Authorized under Security System
        Taskid = Shell("vmenu.exe " & Zarg, 1)
        Me.Close() '
        End
999:
        'Debug.Print(Me.Height.ToString) ' > 600 Then Me.Height = 665 '08-21-11
    End Sub

    Private Sub frmQuoteRpt_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Call Main_Renamed("ENDWIN")
        '06-10-09
        '03-19-12 Call FormSetting("Save") 'During FormClosing Event to Save Settings
    End Sub
    Public Sub mnuSupport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSupport.Click
        '02-24-14 JTC Added mnuSupport"ENABLEBRANDREPORTING"  '    Public Sub mnuBrandMfg2(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBrandMfgChg.Click
        Dim Msg As String = UCase(InputBox("Enter Your Support Function.", "Support Personnel.", "")) '02-24-14
        If Msg = "ENABLEBRANDREPORTING" Then
            ' Call mnuBrandMfg2(eventSender, New System.EventArgs())
            If BrandReportMfg.Trim = "" Then BrandReportMfg = "XXXX"
            BrandReportMfg = UCase(InputBox("Enter Major Brand to breakdown PHIL, COOP, LITH", "Enter Top Level MFG Code.", BrandReportMfg)) '03-12-13
            BrandReportMfg = VB.Left(BrandReportMfg, 4) '04-04-13
            Dim FileName2 As String
            '02-24-14 JTC Move "BRANDREPORT-" From UserPath to UserSysDir
            FileName2 = Dir$(UserSysDir & "BrandReport*.*") '"BrandReport-*.DAT")

            'IF FILENAME IS BLANK - THERE ISN'T A BRANDREPORT FILE IN USERSYS - CREATE ONE
            If FileName2 = "" Then '02-24-14 jh
                Dim FN As Integer = FreeFile() : FileClose(FN)
                FileOpen(FN, UserSysDir & "BrandReport-" & BrandReportMfg & ".DAT", OpenMode.Output)
                FileClose(FN)
                Me.mnuBrandMfgChg.Text = "Brand Mfg Code - " & BrandReportMfg '02-27-14  XXXX Brand Mfg Code - XXXX
                Me.chkBrandReport.Text = "Brand Reporting - " & BrandReportMfg '02-27-14  XXXX Brand Mfg Code - XXX
                MessageBox.Show("Brand Reporting Enabled")
            Else
                If My.Computer.FileSystem.FileExists(UserSysDir & "BrandReport-" & BrandReportMfg & ".DAT") = False Then '02-27-14
                    My.Computer.FileSystem.DeleteFile(UserSysDir & FileName2)
                    Dim FN As Integer = FreeFile() : FileClose(FN)
                    FileOpen(FN, UserSysDir & "BrandReport-" & BrandReportMfg & ".DAT", OpenMode.Output)
                    FileClose(FN)
                    Me.mnuBrandMfgChg.Text = "Brand Mfg Code - " & BrandReportMfg '02-27-14  XXXX Brand Mfg Code - XXXX
                    Me.chkBrandReport.Text = "Brand Reporting - " & BrandReportMfg '02-27-14  XXXX Brand Mfg Code - XXX
                End If
                MessageBox.Show("Brand Reporting Already Enabled")
            End If
            '02-24-14 JH
            BrandReportMfg = BrandReportMfg.ToUpper '02-04-15 JTC Must be Upper case  BrandReportMfg = BrandReportMfg.ToUpper
        End If
    End Sub
    Public Sub mnuJump_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuJump.Click
        '02-24-14 JTC Added Shell(UserPath & "VNAME.EXE " & " " & "/OpenBrandTable", 1)  
        Dim taskid As Integer
        taskid = Shell(UserPath & "VNAME.EXE " & "/OpenBrandTable", 1) '
    End Sub
    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Call Main_Renamed("SHELL255") '09-14-09 to main menu
    End Sub
    Public Function SaveDialog(Optional ByVal FileName As String = "", Optional ByVal Title As String = "", Optional ByVal Filter As String = "") As String
        Me.SaveFileDialog1.FileName = FileName
        Me.SaveFileDialog1.Title = Title
        Me.SaveFileDialog1.Filter = Filter '"Tab Delimeted Files (*.txt)|*.txt"  '|All files (*.*)|*.*"
        Me.SaveFileDialog1.RestoreDirectory = True '06-01-09
        If Me.SaveFileDialog1.ShowDialog = DialogResult.OK Then
            SaveDialog = SaveFileDialog1.FileName
        Else
            SaveDialog = ""
        End If
    End Function
    Public Sub FormSetting(ByVal LoadSave As String)
        'My.Settings.Reset() This will reset the MySettingsFile
        'Call FormSetting("Load") 'During Form Loa
        'Call FormSetting("Save") 'During FormClosing Event to Save Settings
        ''The user.config file is created in the 
        ''<c:\Documents and Settings>\<username>\[Local Settings\]Application Data\<companyname>\<appdomainname>_<eid>_<hash>\<verison>.
        Try  'Old Code Dim Settings As System.Configuration.ApplicationSettingsBase 'Settings = My.Settings
            'Settings.SettingsKey = Me.Name  'Allows Multiple Form Names 'Dim theSettings As My.MySettings 'theSettings = DirectCast(Settings, My.MySettings)
            If LoadSave = "Load" Then '06-10-09
                If My.Settings.FormSize <> Size.Empty Then Me.Size = My.Settings.FormSize
                Me.WindowState = My.Settings.WindowState '06-30-09 
                If My.Settings.FormLocation <> Point.Empty Then Me.Location = My.Settings.FormLocation '09-12-08
                Me.BackColor = My.Settings.BackGroundColor '09-12-08
                Me.ForeColor = My.Settings.ForegroundColor '09-12-08
                Me.Font = My.Settings.Font
                If Me.Height > 600 Then Me.Height = 640 '08-21-11
                If Me.Height < 640 Then Me.Height = 640 '08-21-11
                Me.rbnMaxNameTxt.Text = My.Settings.RibbonMaxNameTxt '01-06-13
                Me.rbnMaxJobTxt.Text = My.Settings.RibbonMaxJobTxt '12-23-12
                Me.chkWholeDollars.Checked = My.Settings.WholeDollars '01-06-13
                Me.chkPrintGrayScale.Checked = My.Settings.PrintGrayScale '01-18-13
                If Me.chkPrintGrayScale.Checked = True Then Me.RibbonTab6.Text = "Print GrayScale" Else Me.RibbonTab6.Text = "Print Color" '01-18-13
                Me.chkAddCommas.Checked = My.Settings.AddCommas '01-06-13
                Me.chkAddDollarSign.Checked = My.Settings.AddDollarSign '01-06-13
            End If
            If LoadSave = "Save" Then
                My.Settings.WindowState = Me.WindowState
                If Me.WindowState = FormWindowState.Normal Then
                    My.Settings.FormSize = Me.Size
                    My.Settings.FormLocation = Me.Location
                Else
                    My.Settings.FormSize = Me.RestoreBounds.Size
                End If
                My.Settings.Font = Me.Font '09-12-08
                My.Settings.BackGroundColor = Me.BackColor '09-12-08
                My.Settings.ForegroundColor = Me.ForeColor '09-12-08
                My.Settings.RibbonMaxNameTxt = Me.rbnMaxNameTxt.Text '01-06-12 My.Settings
                My.Settings.RibbonMaxJobTxt = Me.rbnMaxJobTxt.Text  '12-23-12
                My.Settings.WholeDollars = Me.chkWholeDollars.Checked '01-06-13
                My.Settings.PrintGrayScale = Me.chkPrintGrayScale.Checked '01-18-13
                My.Settings.AddCommas = Me.chkAddCommas.Checked '01-06-13
                My.Settings.AddDollarSign = Me.chkAddDollarSign.Checked  '01-06-13
                My.Settings.Save()
            End If
        Catch ex As Exception
            MessageBox.Show("Error in Form Setting (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12 
        End Try
    End Sub
    'Public Sub mnuFileAbort_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    '    Dim response As Integer

    '    On Error Resume Next
    'End Sub
    Public Sub mnuFileExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileExit.Click
        'Exit to windows
        'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        ZE = "88"
        Call Main_Renamed("ENDWIN") '09-14-09
    End Sub
    Public Sub mnuFileGoTo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileGoTo.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Arrow
        ZE = "88"
        Call Main_Renamed("VQUT.EXE")
    End Sub
    Public Sub mnuHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)  '09-17-09 , Handles mnuHelp2.Click
        '08-27-10 use  Help.ShowHelpIndex(Me, UserPathHelp & "hlpVqrt.chm") not  System.Windows.Forms.SendKeys.Send("{F1}"))
        Call HelpSub("HelpToolStripMenuItem") '08-27-10 JTC 
    End Sub
    Public Sub OpenAttachment(ByVal FileToOpen As String)
        Dim oProcess As Process
        Try '11-24-09 Call OpenAttachment(FileName)
            If My.Computer.FileSystem.FileExists(FileToOpen) Then '11-24-09 
                oProcess = Process.Start(FileToOpen)
            End If
        Catch ex As Exception
            MessageBox.Show("Error in Open Attachment" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub HelpSub(ByRef A As String)
        'send HelpToolStripMenuItem Click Event to this sub to Load PDF
        'Call HelpSub("HelpToolStripMenuItem") '08-27-10 JTC send HelpToolStripMenuItem Click Event to this sub
        Dim HelpPath As String = "" : Dim FileName As String = ""
        UserPathHelp = UserPath & "HELP\" '08-27-10 JTC Added to Make SAWDirectories
        If My.Computer.FileSystem.DirectoryExists(UserPathHelp) = False Then
            My.Computer.FileSystem.CreateDirectory(UserPathHelp) '08-27-10 
        End If
        'Check for PDF first then show it 
        'Me.HelpProvider1.HelpNamespace = UserPathHelp & "HLPQuoteReports.pdf" '09-09-12 
        If My.Computer.FileSystem.FileExists(UserPathHelp & "HLPQuoteReports.pdf") = True Then
            HelpPath = UserPathHelp : FileName = "HLPQuoteReports.pdf"
            If A = "HelpToolStripMenuItem" Then  '08-27-10 JTC
                Call OpenAttachment(HelpPath & FileName)
                Exit Sub
            End If
        Else
            'Me.HelpProvider1.HelpNamespace = UserPathHelp & "HLPQuoteReports.pdf" '09-09-12
            MessageBox.Show(UserPathHelp & "HLPQuoteReports.pdf" & vbCrLf & "  File missing. If the problem persists call Multimicro for support", "VQRT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        'Check for chm second but "hlpUtil.chm" is set in formload
        '09-09-12 Out If My.Computer.FileSystem.FileExists(UserPathHelp & "hlpvqrt.chm") = True Then
        'HelpPath = UserPathHelp : FileName = "hlpvqrt.chm"
        'If A = "HelpToolStripMenuItem" Then  '08-27-10 JTC 
        '    Help.ShowHelpIndex(Me, UserPathHelp & "hlpVqrt.chm") 'Old was System.Windows.Forms.SendKeys.Send("{F1}")
        '    Exit Sub
        'End If
        'End If
        Exit Sub
GotHelpFile:
        'Do In FormLoadfrmAgencyMaster.HelpProvider1.HelpNamespace = HelpPath & FileName '08-27-10 'frmAgencyMaster.SetShowHelp(frmAgencyMaster, True) '06-04-09
        'frmAgencyMaster.SetShowHelp(frmAgencyMaster, True) '06-04-09
    End Sub
    Private Sub chkMfgBreakdown_Click(ByRef Value As Short)
        'rptl% = 4
        If chkMfgBreakdown.CheckState = CheckState.Checked Then 'Threed.Constants_CheckBoxValue.ssCBChecked Then
            RptMFG = 1 '12-06-02 WNA
            If RptCust = 1 Then RptMFGCust = 1 '12-06-02 WNA Else RptMFG% = 1
        Else
            RptMFG = 0
            RptMFGCust = 0
        End If
    End Sub
    Private Sub chkBrandReport_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBrandReport.CheckedChanged
        If Me.chkBrandReport.Checked = True Then
            'If MajSel = RptMaj.RptSpecCredit Or MajSel = RptMaj.RptRefCredit Then
            If Me.pnlTypeOfRpt.Text <> "Quote Summary" And mnuBrandReport.Text <> "Brand Reporting - On" Then '11-06-13
                MsgBox("Brand Reporting not allowed on this option." & vbCrLf & "Turning Brand Reporting Off.")
                mnuBrandReport.Text = "Brand Reporting - Off"
                Me.chkBrandReport.CheckState = CheckState.Unchecked '03-14-13
                : Exit Sub '03-14-13
            End If
            mnuBrandReport.Text = "Brand Reporting - On"
        Else
            ' Me.chkBrandReport.Enabled = False
            mnuBrandReport.Text = "Brand Reporting - Off"
        End If
    End Sub
    'Public Sub mnuBrandMfg2(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) '10-20-13 Handles mnuBrandMfgChg.Click
    '    '10-20-13 Change Brand Name Me.mnuBrandMfgChg.Text = "Brand Mfg Code - " & BrandReportMfg '03-12-13 XXXX Brand Mfg Code - XXXX
    '    If BrandReportMfg.Trim = "" Then BrandReportMfg = "XXXX"
    '    BrandReportMfg = UCase(InputBox("Enter Major Brand to breakdown PHIL, COOP, LITH", "Enter Top Level MFG Code.", BrandReportMfg)) '03-12-13
    '    BrandReportMfg = VB.Left(BrandReportMfg, 4) '04-04-13
    '    Dim FileName2 As String
    '    FileName2 = Dir$(UserPath & "BrandReport*.*") '"BrandReport-*.DAT")
    '    My.Computer.FileSystem.CopyFile(UserPath & FileName2, UserPath & "BrandReport-" & BrandReportMfg & ".DAT")
    '    My.Computer.FileSystem.DeleteFile(UserPath & FileName2)
    '    '03-12-13 Change Brand Name 
    '    ' Me.mnuBrandMfgChg.Text = "Brand Mfg Code - " & BrandReportMfg '03-12-13 XXXX Brand Mfg Code - XXXX
    '    ' Me.chkBrandReport.Text = "Brand Reporting - " & BrandReportMfg '03-14-13 XXXX Brand Mfg Code - XXXX"
    '    '03-12-13 mnuBrandReport
    '    'If mnuBrandReport.Text = "Brand Reporting - Off" Then
    '    '    mnuBrandReport.Text = "Brand Reporting - On"
    '    'End If
    'End Sub
    Public Sub mnuQuoteRealization_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error Resume Next
        Call cmdReportRealization_Click(cmdReportRealization, New System.EventArgs())
    End Sub
    Public Sub mnuQuoteReports_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error Resume Next
        Call cmdReportQuote_Click(cmdReportQuote, New System.EventArgs())
    End Sub
    Private Sub optOne_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If eventSender.Checked Then
            RptL = 1 ' One Line
            'Me.chkExportAllExcel.Visible = False '01-08-04 WNA
        End If
    End Sub
    Private Sub optOutputOptions_Click(ByRef index As Short, ByRef Value As Short)
        '09-06-02 WNA
        Select Case index
            Case 0
                Me.optOutputOptions(3).Checked = CheckState.Checked
                'Me.chkExportAllExcel.Visible = False '01-08-04 WNA
                'Me.chkExportAllExcel.CheckState = False
            Case 1
                Me.optOutputOptions(4).Checked = CheckState.Checked
                If Me.pnlTypeOfRpt.Text = "Realization" Or Me.chkNotes.CheckState = CheckState.Checked Then
                    'Me.chkExportAllExcel.Visible = True '01-08-04 WNA
                End If
            Case 2
                Me.optOutputOptions(5).Checked = CheckState.Checked
                'Me.chkExportAllExcel.Visible = False '01-08-04 WNA
                'Me.chkExportAllExcel.CheckState = False
            Case 3
                Me.optOutputOptions(0).Checked = CheckState.Checked
                Me.chkExportAllExcel.Visible = False '01-08-04 WNA
                Me.chkExportAllExcel.CheckState = False
            Case 4
                Me.optOutputOptions(1).Checked = CheckState.Checked
                If Me.pnlTypeOfRpt.Text = "Realization" Or Me.chkNotes.CheckState = CheckState.Checked Then
                    Me.chkExportAllExcel.Visible = True '01-08-04 WNA
                End If
            Case 5
                Me.optOutputOptions(2).Checked = CheckState.Checked
                Me.chkExportAllExcel.Visible = False '01-08-04 WNA
                Me.chkExportAllExcel.CheckState = False
        End Select
    End Sub
    Private Sub optSelectNewRecords_Click(ByRef Value As Short)
        'Me.cmdFmtOK.Text = "Select New Records"
    End Sub
    Private Sub optThree_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If eventSender.Checked Then
            RptL = 3 ' 3 Line
            If Me.optOutputOptions(1).Checked = CheckState.Checked Then ' Excel   '01-08-04 WNA
                'Me.chkExportAllExcel.Visible = True
            End If
        End If
    End Sub
    Private Sub optTwo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If eventSender.Checked Then
            RptL = 2 ' 2 line
            'Me.chkExportAllExcel.Visible = False '01-08-04 WNA
        End If
    End Sub
    Private Sub optUsePrevSelection_Click(ByRef Value As Short)
        'Me.cmdFmtOK.Text = "Run Report"
    End Sub
    Private Sub tabQrt_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles tabQrt.SelectedIndexChanged
        On Error Resume Next
        'Moved cboTypeofJob.Text = "Q" '04-26-12
        ' Me.Refresh() 'Me.InvokePaint(Me, New EventArgs)
        'Application.DoEvents()

        Select Case tabQrt.SelectedIndex
            Case 0
                Me.fraDisplaySortSeq.Visible = True
                Me.fraSortPrimarySeq.Refresh() '09-21-12
                Me.fraSortPrimarySeq.BackColor = Color.White '09-25-12
            Case 1
                If Me.pnlTypeOfRpt.Text <> "Product Sales History - Line Items" Then '12-15-09 
                    Me.fraDisplaySortSeq.Visible = True
                Else
                    Me.fraQuoteReports.Visible = False
                    Me.fraQuoteLineReports.Visible = True

                End If
                '      If MARK Then   'MARK = Swap Status & Salesman
                '         lblStatus.Caption = "Salesman Code   (REP10)"
                '         lblSalesman.Caption = "Status Code      (OPN) "
                '      End If
            Case 2
                'Me.Panel1.Visible = True  '09-25-12
                '"Show Hide Quote Hdr Printing Columns" = "VQrtShowHidePrtHdr.xml"   tgQh
                '"Show Hide Quote Line Items" = "VQrtLineItemsDistShowHide.xml"  tgln or Rep
                '"Show Hide Realization Columns" = "VQrtRealQTOShowHidePrint.xml"   tgr
                Dim tmpName As String = "" ' howHidePrintQrt.xml" '09-02-09
                'If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Quote Summary") Then '09-19-12  GoTo QutLineHistoryRpt
                If Me.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Then '09-21-12
                    If tgln.Visible = False Then Exit Sub '03-21-14
                    If tgln.Caption = "Spec Credit - Click Run Report to see final report selection" Then Exit Sub '03-21-14
                    frmShowHideGrid.Text = "Show Hide Quote Line Items"
                    Me.Show() '02-06-09
                    If DIST Then tmpName = "VQrtLineItemsDistShowHide.xml" Else tmpName = "VQrtLineItemsRepShowHide.xml" '05-05-10 tmpName = "VQrtLineItemsShowHide.xml" '
                    ShowHideFileName = tmpName '09-02-09
                    Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgln) '09-02-09
                    Me.Panel1.Visible = True '09-25-12
                    Exit Sub 'Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgr) '09-02-09
                ElseIf (Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (Me.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or Me.txtPrimarySortSeq.Text = "Salesman Follow-Up Report")) Then '03-19-14 was above
                    '03-19-14 FIXED FORMAT FOR NOW
                    'frmShowHideGrid.Text = "Show Hide Spec Credit"
                    'Me.Show()
                    'tmpName = "VQrtSpecCreditShowHide.xml"
                    'ShowHideFileName = tmpName
                    'Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgln)
                    'Me.Panel1.Visible = True
                    Exit Sub
                End If
                'Me.fraDisplaySortSeq.Visible = False ShowHideFileName = UserDir & "ShowHidePrintFoll.xml" '12-11-08
                If Me.pnlTypeOfRpt.Text = "Realization" Then
                    If tgr.Visible = False Then Exit Sub '03-21-14
                    frmShowHideGrid.Text = "Show Hide Realization Columns" '12-22-08
                    Me.Show() '02-06-09
                    '05-07-10 VQrtRealQTOShowHideDistPrint.xml else VQrtRealQTOShowHideRepPrint.xml
                    Dim ShowAllQuoteHeader As String = "" '06-06-11 "Show All Quote Header Fields"
                    If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then ShowAllQuoteHeader = "ShowAll" '06-06-11 = "Show All Quote Header Fields" Then '06-06-11 "Add Cust QuoteTo Breakdown to Report"
                    If DIST Then tmpName = "VQrtRealQTOShowHideDistPrint" & ShowAllQuoteHeader & ".xml" Else tmpName = "VQrtRealQTOShowHideRepPrint" & ShowAllQuoteHeader & ".xml" '06-06-11 
                    'If Me.pnlTypeOfRpt.Text = "Realization" Then tmpName = "VQrtRealQTOShowHidePrint.xml"
                    ShowHideFileName = tmpName '09-02-09
                    Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgr) '09-02-09) '12-05-08 ByVal ShowHide As String)
                Else '  Reqular Quote Report & Planned Project
                    If tgQh.Visible = False Then Exit Sub '03-19-14
                    frmShowHideGrid.Text = "Show Hide Quote Hdr Printing Columns" '"VQrtShowHidePrtHdr.xml"
                    Me.Show() '02-06-09
                    If DIST Then '01-19-10
                        tmpName = "VQrtShowHideDistPrtHdr.xml" '
                    Else
                        tmpName = "VQrtShowHideRepPrtHdr.xml" '
                    End If
                    If Me.pnlTypeOfRpt.Text = "Realization" Then 'tmpName = "VQrtRealQTOShowHidePrint.xml"
                        Dim ShowAllQuoteHeader As String = "" '06-06-11 "Show All Quote Header Fields"
                        If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then ShowAllQuoteHeader = "ShowAll" '06-06-11 = "Show All Quote Header Fields" Then '06-06-11 "Add Cust QuoteTo Breakdown to Report"
                        If DIST Then tmpName = "VQrtRealQTOShowHideDistPrint" & ShowAllQuoteHeader & ".xml" Else tmpName = "VQrtRealQTOShowHideRepPrint" & ShowAllQuoteHeader & ".xml" '06-06-11 
                    End If
                    ShowHideFileName = tmpName '09-02-09
                    '05-15-14 JTC Error if Me.pnlTypeOfRpt.Text <> "Terr Spec Credit Report" Then Don't do Call frmShowHideGrid.ShowHideGridCol(" IE: Fixed Columns
                    If Me.pnlTypeOfRpt.Text <> "Terr Spec Credit Report" Then Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgQh) '09-02-09)
                End If

        End Select
    End Sub

    Private Sub tabQRT_TabActivate(ByRef TabToActivate As Short)
        'On Error Resume Next
        ''Moved cboTypeofJob.Text = "Q" '04-26-12
        '' Me.Refresh() 'Me.InvokePaint(Me, New EventArgs)
        ''Application.DoEvents()

        'Select Case TabToActivate
        '    Case 0
        '        Me.fraDisplaySortSeq.Visible = True
        '        'Me.Panel1.Visible = False '09-25-12

        '    Case 1
        '        If Me.pnlTypeOfRpt.Text <> "Product Sales History - Line Items" Then '12-15-09 
        '            Me.fraDisplaySortSeq.Visible = True
        '        Else
        '            Me.fraQuoteReports.Visible = False
        '            Me.fraQuoteLineReports.Visible = True
        '        End If
        '        '      If MARK Then   'MARK = Swap Status & Salesman
        '        '         lblStatus.Caption = "Salesman Code   (REP10)"
        '        '         lblSalesman.Caption = "Status Code      (OPN) "
        '        '      End If
        '    Case 2
        '        'Me.Panel1.Visible = True  '09-25-12
        '        '"Show Hide Quote Hdr Printing Columns" = "VQrtShowHidePrtHdr.xml"   tgQh
        '        '"Show Hide Quote Line Items" = "VQrtLineItemsDistShowHide.xml"  tgln or Rep
        '        '"Show Hide Realization Columns" = "VQrtRealQTOShowHidePrint.xml"   tgr
        '        Dim tmpName As String = "" ' howHidePrintQrt.xml" '09-02-09
        '        'If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Quote Summary") Then '09-19-12  GoTo QutLineHistoryRpt
        '        If Me.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Or (Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (Me.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or Me.txtPrimarySortSeq.Text = "Salesman Follow-Up Report")) Then '09-21-12
        '            frmShowHideGrid.Text = "Show Hide Quote Line Items"
        '            Me.Show() '02-06-09
        '            If DIST Then tmpName = "VQrtLineItemsDistShowHide.xml" Else tmpName = "VQrtLineItemsRepShowHide.xml" '05-05-10 tmpName = "VQrtLineItemsShowHide.xml" '
        '            ShowHideFileName = tmpName '09-02-09
        '            Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgln) '09-02-09
        '            Exit Sub 'Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgr) '09-02-09
        '        End If
        '        'Me.fraDisplaySortSeq.Visible = False ShowHideFileName = UserDir & "ShowHidePrintFoll.xml" '12-11-08
        '        If Me.pnlTypeOfRpt.Text = "Realization" Then
        '            frmShowHideGrid.Text = "Show Hide Realization Columns" '12-22-08
        '            Me.Show() '02-06-09
        '            '05-07-10 VQrtRealQTOShowHideDistPrint.xml else VQrtRealQTOShowHideRepPrint.xml
        '            Dim ShowAllQuoteHeader As String = "" '06-06-11 "Show All Quote Header Fields"
        '            If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then ShowAllQuoteHeader = "ShowAll" '06-06-11 = "Show All Quote Header Fields" Then '06-06-11 "Add Cust QuoteTo Breakdown to Report"
        '            If DIST Then tmpName = "VQrtRealQTOShowHideDistPrint" & ShowAllQuoteHeader & ".xml" Else tmpName = "VQrtRealQTOShowHideRepPrint" & ShowAllQuoteHeader & ".xml" '06-06-11 
        '            'If Me.pnlTypeOfRpt.Text = "Realization" Then tmpName = "VQrtRealQTOShowHidePrint.xml"
        '            ShowHideFileName = tmpName '09-02-09
        '            Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgr) '09-02-09) '12-05-08 ByVal ShowHide As String)
        '        Else '  Reqular Quote Report & Planned Project
        '            frmShowHideGrid.Text = "Show Hide Quote Hdr Printing Columns" '"VQrtShowHidePrtHdr.xml"
        '            Me.Show() '02-06-09
        '            If DIST Then '01-19-10
        '                tmpName = "VQrtShowHideDistPrtHdr.xml" '
        '            Else
        '                tmpName = "VQrtShowHideRepPrtHdr.xml" '
        '            End If
        '            If Me.pnlTypeOfRpt.Text = "Realization" Then 'tmpName = "VQrtRealQTOShowHidePrint.xml"
        '                Dim ShowAllQuoteHeader As String = "" '06-06-11 "Show All Quote Header Fields"
        '                If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then ShowAllQuoteHeader = "ShowAll" '06-06-11 = "Show All Quote Header Fields" Then '06-06-11 "Add Cust QuoteTo Breakdown to Report"
        '                If DIST Then tmpName = "VQrtRealQTOShowHideDistPrint" & ShowAllQuoteHeader & ".xml" Else tmpName = "VQrtRealQTOShowHideRepPrint" & ShowAllQuoteHeader & ".xml" '06-06-11 
        '            End If
        '            ShowHideFileName = tmpName '09-02-09
        '            Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgQh) '09-02-09)
        '        End If

        'End Select
    End Sub
    Private Sub txtCity_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCity.Enter
        On Error Resume Next
        Me.txtCity.SelectionStart = 0
        Me.txtCity.SelectionLength = Len(Me.txtCity.Text)
        If txtCity.Text = "" Then txtCity.Text = "ALL" '09-25-07 JH
    End Sub
    Private Sub txtCity_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCity.KeyPress, txtSelectCode.KeyPress, txtSlsSplit.KeyPress, txtJobNameSS.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub


    Private Sub txtCity_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCity.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtCity.Text = Trim(txtCity.Text)
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCSR_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCSR.Enter
        On Error Resume Next
        Me.txtCSR.SelectionStart = 0
        Me.txtCSR.SelectionLength = Len(Me.txtCSR.Text)
    End Sub
    Private Sub txtCSR_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCSR.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCSR_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCSR.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtCSR.Text = Trim(UCase(txtCSR.Text))
        If txtCSR.Text = "" Then txtCSR.Text = "ALL" '09-25-07 JH
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCSRDist_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCSRofCust.Enter
        On Error Resume Next '09-25-07 JH
        Me.txtCSRofCust.SelectionStart = 0
        Me.txtCSRofCust.SelectionLength = Len(Me.txtCSRofCust.Text)
    End Sub
    Private Sub txtCSRDist_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCSRofCust.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error Resume Next
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0 '09-25-07 JH
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCSRDist_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCSRofCust.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtCSRofCust.Text = Trim(UCase(txtCSRofCust.Text)) '09-25-07 JH
        If txtCSRofCust.Text = "" Then txtCSRofCust.Text = "ALL"
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCurrentPage_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCurrentPage.Enter
        Dim index As Short = txtCurrentPage.GetIndex(eventSender)
        On Error Resume Next
        Me.txtCurrentPage(0).SelectionStart = 0
        Me.txtCurrentPage(0).SelectionLength = Len(Me.txtCurrentPage(0).Text)
    End Sub
    Private Sub txtCurrentPage_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCurrentPage.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim index As Short = txtCurrentPage.GetIndex(eventSender)
        On Error Resume Next
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    '    Private Sub txtCurrentPage_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCurrentPage.Validating
    '        Dim Cancel As Boolean = eventArgs.Cancel
    '        Dim index As Short = txtCurrentPage.GetIndex(eventSender)
    '        On Error Resume Next
    '        If index = 0 Then
    '            If IsNumeric(Val(Me.txtCurrentPage(0).Text)) = True Then
    '                'If Val(Me.txtCurrentPage(0).Text) >= 1 And Val(Me.txtCurrentPage(0).Text) <= Me.vsPrinter1.PageCount Then
    '                '	Me.vsPrinter1.PreviewPage = Val(Me.txtCurrentPage(0).Text)
    '                '	GoTo EventExitSub
    '                'End If
    '            End If
    '        Else
    '            GoTo EventExitSub
    '        End If
    '        MsgBox("Invalid Page Number Entered" & vbCrLf & "Number Must be Between 1 and " & Str(CDbl(Me.txtCurrentPage(1).Text)), , "Select Current Page")
    '        Me.txtCurrentPage(0).Text = CStr(Page) 'Me.vsPrinter1.PreviewPage)
    '        Cancel = True
    'EventExitSub:
    '        eventArgs.Cancel = Cancel
    '    End Sub
    Private Sub txtEndBid_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        txtEndBid.SelectionStart = 0
        txtEndBid.SelectionLength = Len(txtEndBid.Text)
    End Sub
    Private Sub txtEndBid_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEndBid_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs)
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim Resp As Short
        Dim Msg As String
        Dim D As String
        Dim A As String
        txtEndBid.Text = UCase(txtEndBid.Text)
        A = Trim(txtEndBid.Text)
        D = ""
        If A <> "ALL" Then
            If Len(A) <> 6 Then Msg = "Date Must Be Entered in MMDDYY Format -- Please Reenter" : GoTo EndBidError
            If Val(VB.Left(A, 2)) < 1 Or Val(VB.Left(A, 2)) > 12 Then Msg = "Month Error -- Please Reenter" : GoTo EndBidError
            If Val(Mid(A, 3, 2)) < 1 Or Val(Mid(A, 3, 2)) > 31 Then Msg = "Day Error -- Please Reenter" : GoTo EndBidError
            Me.chkBlankBidDates.Visible = True '02-19-04 WNA
        Else '02-19-04 WNA
            If Trim(Me.txtStartBid.Text) = "ALL" Then Me.chkBlankBidDates.Visible = False
        End If
        If Me.chkBlankBidDates.Visible = True Then '02-26-04 WNA
            If Trim(Me.txtStartEntry.Text) = "ALL" And Trim(Me.txtEndEntry.Text) = "ALL" Then '02-26-04 WNA
                Me.chkBlankBidDates.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                Me.chkBlankBidDates.CheckState = System.Windows.Forms.CheckState.Checked
            End If
        End If
        GoTo EventExitSub

EndBidError:
        Msg = txtEndBid.Text & " is invalid." & vbCrLf & Msg & vbCrLf & "Click on Down Arrow to select a valid date" '06-13-00
        Resp = MsgBox(Msg, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Ending Bid Date") : Cancel = True
        txtEndBid.Text = "ALL"

EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtEndEntry_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        txtEndEntry.SelectionStart = 0
        txtEndEntry.SelectionLength = Len(txtEndEntry.Text)
    End Sub
    Private Sub txtEndEntry_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEndEntry_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs)
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim Resp As Short
        Dim Msg As String
        Dim A As String
        Dim D As String
        D = ""
        txtEndEntry.Text = UCase(txtEndEntry.Text)
        A = UCase(txtEndEntry.Text)
        If Trim(A) <> "ALL" Then
            If Len(A) <> 6 Then Msg = "Date Must Be Entered in MMDDYY Format -- Please Reenter" : GoTo EndError
            If Val(VB.Left(A, 2)) < 1 Or Val(VB.Left(A, 2)) > 12 Then Msg = "Month Error -- Please Reenter" : GoTo EndError
            If Val(Mid(A, 3, 2)) < 1 Or Val(Mid(A, 3, 2)) > 31 Then Msg = "Day Error -- Please Reenter" : GoTo EndError
            If Me.chkBlankBidDates.Visible = True Then '02-26-04 WNA
                Me.chkBlankBidDates.CheckState = System.Windows.Forms.CheckState.Checked
            End If
        Else
            If Me.chkBlankBidDates.Visible = True Then '02-26-04 WNA
                If Trim(Me.txtStartEntry.Text) <> "ALL" Then
                    Me.chkBlankBidDates.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    Me.chkBlankBidDates.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
            End If
        End If

        GoTo EventExitSub

EndError:
        Msg = txtEndEntry.Text & " is invalid." & vbCrLf & Msg & vbCrLf & "Click on Down Arrow to select a valid date" '06-13-00
        Resp = MsgBox(Msg, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Ending Entry Date") : Cancel = True
        txtEndEntry.Text = "ALL"

EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtEndQuote_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEndQuoteAmt.Enter
        txtEndQuoteAmt.SelectionStart = 0
        txtEndQuoteAmt.SelectionLength = Len(txtEndQuoteAmt.Text)
    End Sub
    Private Sub txtEndQuote_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEndQuoteAmt.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEndQuote_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtEndQuoteAmt.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim A As String = ""
        txtEndQuoteAmt.Text = Trim(UCase(txtEndQuoteAmt.Text))
        If A <> "999999999" Then '03-24-08 JTC Added 9 "999,999,999"
            If IsNumeric(txtEndQuoteAmt.Text) = False Then MsgBox("Entry Must Be Numeric -- Please Reenter") : txtEndQuoteAmt.Text = "999999999" : Cancel = True : GoTo EventExitSub
        End If
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtJobNameSS_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtJobNameSS.Enter
        txtJobNameSS.SelectionStart = 0
        txtJobNameSS.SelectionLength = Len(txtJobNameSS.Text)
    End Sub
    Private Sub txtJobNameSS_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtJobNameSS.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtJobNameSS_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtJobNameSS.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtJobNameSS.Text = Trim(UCase(txtJobNameSS.Text))
        'If Trim$(txtJobNameSS.Text) = "" Then
        '    SelJobSS$ = "ALL"
        'Else
        '    SelJobSS$ = UCase$(txtJobNameSS.Text)
        'End If
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLastChgBy_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLastChgBy.Enter
        On Error Resume Next
        Me.txtLastChgBy.SelectionStart = 0
        Me.txtLastChgBy.SelectionLength = Len(Me.txtLastChgBy.Text)
    End Sub
    Private Sub txtLastChgBy_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLastChgBy.KeyPress, txtSlsTerr.KeyPress, txtSpecCross.KeyPress, TxtSingleCatNum.KeyPress, txtMfgLine.KeyPress, txtPrcCode.KeyPress, txtRetr.KeyPress, txtStat.KeyPress, TxtSearchString.KeyPress, txtLastChgByLine.KeyPress '09-03-09
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtLastChgBy_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLastChgBy.Validating, txtSlsTerr.Validating, txtSpecCross.Validating, TxtSingleCatNum.Validating, txtMfgLine.Validating, txtPrcCode.Validating, txtRetr.Validating, txtStat.Validating, TxtSearchString.Validating, txtLastChgByLine.Validating '09-03-09

        CType(eventSender, TextBox).Text = CType(eventSender, TextBox).Text.ToUpper
        'txtLastChgBy.Text = Trim(UCase(txtLastChgBy.Text))
        'If txtLastChgBy.Text = "" Then txtLastChgBy.Text = "ALL" '09-25-07 JH
        'eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtMktSegment_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMktSegment.Enter
        On Error Resume Next
        Me.txtMktSegment.SelectionStart = 0
        Me.txtMktSegment.SelectionLength = Len(Me.txtMktSegment.Text)
    End Sub
    Private Sub txtMktSegment_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMktSegment.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtMktSegment_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtMktSegment.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtMktSegment.Text = Trim(txtMktSegment.Text)
        If txtMktSegment.Text = "" Then txtMktSegment.Text = "ALL" '09-25-07 JH
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtQuoteToSls_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtQuoteToSls.Enter
        On Error Resume Next
        Me.txtQuoteToSls.SelectionStart = 0
        Me.txtQuoteToSls.SelectionLength = Len(Me.txtQuoteToSls.Text)
    End Sub
    Private Sub txtQuoteToSls_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtQuoteToSls.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    'Private Sub txtSelectCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSelectCode.KeyPress
    '    Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
    '    On Error Resume Next
    '    'Enter Key = Tab Key
    '    If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
    '    eventArgs.KeyChar = Chr(KeyAscii)
    '    If KeyAscii = 0 Then
    '        eventArgs.Handled = True
    '    End If
    'End Sub
    Private Sub txtQuoteToSls_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtQuoteToSls.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtQuoteToSls.Text = UCase(Trim(txtQuoteToSls.Text))
        If txtQuoteToSls.Text = "" Then txtQuoteToSls.Text = "ALL" '09-25-07 JH
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtQutRealCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtQutRealCode.Enter
        txtQutRealCode.SelectionStart = 0
        txtQutRealCode.SelectionLength = Len(txtQutRealCode.Text)

    End Sub
    Private Sub txtQutRealCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtQutRealCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtQutRealCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtQutRealCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtQutRealCode.Text = UCase(Trim(txtQutRealCode.Text))
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtRetrieval_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRetrieval.Enter
        txtRetrieval.SelectionStart = 0
        txtRetrieval.SelectionLength = Len(txtRetrieval.Text)
    End Sub
    Private Sub txtRetrieval_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRetrieval.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtRetrieval_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtRetrieval.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtRetrieval.Text = UCase(Trim(txtRetrieval.Text))
        If txtRetrieval.Text = "" Then txtRetrieval.Text = "ALL" '09-25-07 JH
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSalesman_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSalesman.Enter
        txtSalesman.SelectionStart = 0
        txtSalesman.SelectionLength = Len(txtSalesman.Text)
    End Sub
    Private Sub txtSalesman_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSalesman.KeyPress
        Dim keyAscii As Short = Asc(eventArgs.KeyChar)
        If keyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : keyAscii = 0
        eventArgs.KeyChar = Chr(keyAscii)
        If keyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSalesman_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSalesman.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtSalesman.Text = UCase(Trim(txtSalesman.Text))
        If txtSalesman.Text = "" Then txtSalesman.Text = "ALL" '09-25-07 JH
        'If MARK Then 'MARK = Swap Status & Salesman
        '    If Len(txtSalesman.Text) > 3 Then MsgBox "Status Must Be 3 Characters Long -- Please Reenter", vbOKOnly, US$: txtSalesman.Focus: Exit Sub 'MARK
        'Else
        '05-19-11 JTC,JKH,WES If Len(txtSalesman.Text) > 3 Then MsgBox("Salesman/Territory Must Be 3 Characters Long -- Please Reenter", vbOKOnly, "Salesman Code") : txtSalesman.Focus()
        'End If
        eventArgs.Cancel = Cancel
    End Sub
    'Private Sub txtSelectCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSelectCode.Enter
    '    On Error Resume Next
    '    Me.txtSelectCode.SelectionStart = 0
    '    Me.txtSelectCode.SelectionLength = Len(Me.txtSelectCode.Text)
    'End Sub

    'Private Sub txtSelectCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSelectCode.Validating
    '    Dim Cancel As Boolean = eventArgs.Cancel
    '    Dim Resp As Short
    '    On Error Resume Next
    '    txtSelectCode.Text = Trim(UCase(txtSelectCode.Text))
    '    If Me.txtSelectCode.Text <> "ALL" Then
    '        If Len(Trim(Me.txtSelectCode.Text)) > 1 Then
    '            Resp = MsgBox("Invalid Entry - Select Code can only be one character", MsgBoxStyle.Information, "Single Select Code")
    '            Me.txtSelectCode.Text = "ALL"
    '            Cancel = True
    '        End If
    '    End If

    '    eventArgs.Cancel = Cancel
    'End Sub
    Private Sub txtSlsSplit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSlsSplit.Enter
        On Error Resume Next
        Me.txtSlsSplit.SelectionStart = 0
        Me.txtSlsSplit.SelectionLength = Len(Me.txtSlsSplit.Text)
    End Sub
    Private Sub txtSlsSplit_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSlsSplit.KeyPress '07-07-09
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSlsSplit_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSlsSplit.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtSlsSplit.Text = Trim(UCase(txtSlsSplit.Text))
        If txtSlsSplit.Text = "" Then txtSlsSplit.Text = "ALL" '09-25-07 JH
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSpecifierCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSpecifierCode.Enter
        txtSpecifierCode.SelectionStart = 0
        txtSpecifierCode.SelectionLength = Len(txtSpecifierCode.Text)
    End Sub
    Private Sub txtSpecifierCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSpecifierCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSpecifierCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSpecifierCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtSpecifierCode.Text = UCase(Trim(txtSpecifierCode.Text)) '02-27-01 WNA
        If txtSpecifierCode.Text = "" Then txtSpecifierCode.Text = "ALL" '09-25-07 JH
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtStartBid_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        txtStartBid.SelectionStart = 0
        txtStartBid.SelectionLength = Len(txtStartBid.Text)
    End Sub
    Private Sub txtStartBid_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtStartBid_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs)
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim Resp As Short
        Dim Msg As String
        Dim D As String
        Dim A As String
        txtStartBid.Text = UCase(txtStartBid.Text)
        A = Trim(txtStartBid.Text)
        D = ""
        If A <> "ALL" Then
            If Len(A) <> 6 Then Msg = "Date Must Be Entered in MMDDYY Format -- Please Reenter" : GoTo StartBidError
            If Val(VB.Left(A, 2)) < 1 Or Val(VB.Left(A, 2)) > 12 Then Msg = "Month Error -- Please Reenter" : GoTo StartBidError
            If Val(Mid(A, 3, 2)) < 1 Or Val(Mid(A, 3, 2)) > 31 Then Msg = "Day Error -- Please Reenter" : GoTo StartBidError
            Me.chkBlankBidDates.Visible = True '02-19-04 WNA
        Else '02-19-04 WNA
            If Trim(Me.txtEndBid.Text) = "ALL" Then Me.chkBlankBidDates.Visible = False
        End If
        If Me.chkBlankBidDates.Visible = True Then '02-26-04 WNA
            If Trim(Me.txtStartEntry.Text) = "ALL" And Trim(Me.txtEndEntry.Text) = "ALL" Then '02-26-04 WNA
                Me.chkBlankBidDates.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                Me.chkBlankBidDates.CheckState = System.Windows.Forms.CheckState.Checked
            End If
        End If
        GoTo EventExitSub

StartBidError:
        Msg = txtStartBid.Text & " is invalid." & vbCrLf & Msg & vbCrLf & "Click on Down Arrow to select a valid date" '06-13-00
        Resp = MsgBox(Msg, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Starting Bid Date") : Cancel = True
        txtStartBid.Text = "ALL"

EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub txtStartQuote_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStartQuoteAmt.Enter
        txtStartQuoteAmt.SelectionStart = 0
        txtStartQuoteAmt.SelectionLength = Len(txtStartQuoteAmt.Text)
    End Sub
    Private Sub txtStartQuote_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtStartQuoteAmt.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtStartQuote_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtStartQuoteAmt.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtStartQuoteAmt.Text = Trim(UCase(txtStartQuoteAmt.Text))
        If txtStartQuoteAmt.Text <> "0" Then
            If IsNumeric(txtStartQuoteAmt.Text) = False Then MsgBox("Entry Must Be Numeric -- Please Reenter") : txtStartQuoteAmt.Text = "0" : Cancel = True : GoTo EventExitSub 'txtStartQuote.Focus: Exit Sub
        End If
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtState_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtState.Enter
        On Error Resume Next
        Me.txtState.SelectionStart = 0
        Me.txtState.SelectionLength = Len(Me.txtState.Text)
    End Sub
    Private Sub txtState_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtState.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtState_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtState.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtState.Text = Trim(UCase(txtState.Text))
        If txtState.Text = "" Then txtState.Text = "ALL" '09-25-07 JH
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtStatus_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStatus.Enter
        '03-06-09 Private Sub txtStatus_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStatus.TextChanged
        txtStatus.SelectionStart = 0
        txtStatus.SelectionLength = Len(txtStatus.Text)
    End Sub
    Private Sub txtStatus_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtStatus.KeyPress
        '02-06-09 Private Sub txtStatus_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStatus.TextChanged
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}") : KeyAscii = 0
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtStatus_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtStatus.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        txtStatus.Text = UCase(Trim(txtStatus.Text))
        If txtStatus.Text = "" Then txtStatus.Text = "ALL" '09-25-07 JH
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub JumpBidBoard_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VBIDB.EXE") : Me.Close()
    End Sub

    Private Sub JumpCatalog_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VCAT.EXE") : Me.Close()
    End Sub

    Private Sub JumpCommissionRec_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VCOM.EXE") : Me.Close()
    End Sub

    Private Sub JumpCrossover_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VXOV.EXE") : Me.Close()
    End Sub

    Private Sub JumpFactoryStatus_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VSTAT.EXE") : Me.Close()
    End Sub

    Private Sub JumpFollowUp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VFOLLOW.EXE") : Me.Close()
    End Sub

    Private Sub JumpInvoicing_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VOIN.EXE") : Me.Close()
    End Sub

    Private Sub JumpLetter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VLTR.EXE") : Me.Close()
    End Sub

    Private Sub JumpOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VORDER.EXE") : Me.Close()
    End Sub

    '09-10-12Private Sub JumpOutofTerritory_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '     Call Jump("VSPC.EXE") : Me.Close()
    ' End Sub

    Private Sub JumpPDF_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VSUBM.EXE") : Me.Close()
    End Sub

    '09-10-12Private Sub JumpPlannedProjects_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call Jump("VPLAN.EXE") : Me.Close()
    'End Sub

    Private Sub JumpPrice_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VPRICE.EXE") : Me.Close()
    End Sub

    Private Sub JumpProjects_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VPROJ.EXE") : Me.Close()
    End Sub

    Private Sub JumpQuote_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VQUT.EXE") : Me.Close()
    End Sub

    Private Sub JumpSample_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Call Jump("VSAMP.EXE") : Me.Close()
    End Sub

    Private Sub JumpSubmittalLabel_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Call Jump("VTRANS.EXE") : Me.Close()
    End Sub

    Private Sub JumpUtility_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Call Jump("VUTIL.EXE") : Me.Close()
    End Sub

    Private Sub JumpWarehouseInventory_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VWHS.EXE") : Me.Close()
    End Sub

    Private Sub JumpWhseOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Jump("VORDERW.EXE") : Me.Close()
    End Sub
    Public Sub Jump(ByRef NM As String)
        Dim taskid As Integer
        Dim Resp As Short
        Dim DFile As String
        On Error Resume Next
        FileClose()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor 'HourGlass
        DFile = Dir(NM)
        If Len(DFile) = 0 Then Resp = MsgBox(NM & " Program not on your system", 16, "Sales Assistant") : Exit Sub 'GoTo 999
        taskid = Shell(NM & " " & Zarg, 1)
        FileClose() : Call Me.FormSetting("Save") 'During FormClosing Event to Save Settings
        Call CloseSQL(myConnection) '06-10-09  myConnection.Close() '06-10-09
        Me.Close() '03-19-12 
        End ' 
    End Sub


    Private Sub cmdPrimarySeqContinue1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrimarySeqContinue1.Click, cboSortPrimarySeq.DoubleClick
        ' Private Sub cmdSotPrimarySeq_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSortPrimarySeq.Click, cmdSortPrimarySeqContinue
        'Dim index As Short = cmdSortPrimarySeq.GetIndex(sender) '01-28-09
        'SESCO = False '02-25-12 
        OrderBy = "" '01-16-14
        Me.chkBrandReport.CheckState = CheckState.Unchecked '01-14-14
        Me.mnuBrandReport.Enabled = False '11-06-13 
        Me.mnuBrandMfgChg.Enabled = False '11-06-13
        Me.chkBlankBidDates.Visible = True '11-06-13
        Me.mnuBrandReport.Text = "Brand Reporting - Off"
        Me.chkBrandReport.Visible = True '

        Me.lblStartEntry.Enabled = True '11-06-13
        Me.DTPickerStartEntry.Enabled = True '11-06-13
        Me.DTPicker1EndEntry.Enabled = True '11-06-13
        If Me.txtPrimarySortSeq.Text = "SESCO Job List Report" Then  Else SESCO = False '10-31-12 Turn Off Unless Last Item Checked 
        If Me.pnlTypeOfRpt.Text = "Realization" Then '04-28-15 Me.pnlTypeOfRpt.Text = "Realization"
            If cboSortRealization.Items(12) = "Excel Quote FollowUp" And cboSortRealization.GetItemCheckState(12) = CheckState.Checked Then
                ExcelQuoteFU = True '04-28-15 JTC
                'Me.txtPrimarySortSeq.Text = "Excel Quote FollowUp" '04-28-15 JTC
                'Me.cboSortPrimarySeq.SelectedItem(12) = "Excel Quote FollowUp"
                'Me.cboSortPrimarySeq.Text = "Excel Quote FollowUp" '04-28-15 JTC
                'Me.cboSortPrimarySeq.SelectedItem = "Excel Quote FollowUp"
            End If
        Else   'Not Realization
            ' Me.cboSortRealization.Items(12) = CheckState.Unchecked '04-29-15
            ExcelQuoteFU = False
        End If
        Me.chkDetailTotal.CheckState = CheckState.Unchecked '07-26-12
        Dim Enable As Short
        Dim Resp As Short
        On Error Resume Next
        Me.chkSlsFromHeader.Text = "Use Salesman From Quote Header on Report" ' "Use Quote SLS 1 Split for Salesman" '03-08-13
        Me.chkSlsFromHeader.CheckState = CheckState.Unchecked '03-08-13 Realization uses "Use Salesman From Quote Header on Report"
        If Me.pnlTypeOfRpt.Text = "Quote Summary" Then '03-08-13
            Me.chkSlsFromHeader.Visible = False
            Me.chkSlsFromHeader.Enabled = False
            'Me.ChkTotalsOnly.Visible = True '01-12-15 JTC If run Realization B/4 Quote Summary then Me.ChkTotalsOnly.Visible is False
            Me.chkDetailTotal.Visible = True '01-12-15 JTC If run Realization B/4 Quote Summary then Me.ChkDetailTotalsOnly.Visible is False
        End If

        If Me.txtPrimarySort.Text = "MFG Follow-Up Report" Or Me.txtPrimarySort.Text = "Salesman Follow-Up Report" Then '03-19-14
            chkIncludeSpecifiers.Visible = True
            chkIncludeSLSSPlit.Visible = True
            chkIncludeNotesLineItems.Visible = True
        Else
            chkIncludeSpecifiers.Visible = False
            chkIncludeSLSSPlit.Visible = False
            chkIncludeNotesLineItems.Visible = False
        End If
        '02-01-14 JTC Don't Use Me.ChkSpecifiers.Text = "Sort Report by Descending Dollar" anymore
        'If Me.ChkSpecifiers.Text = "Sort Report by Descending Dollar" Then '09-06-12
        '    Me.ChkSpecifiers.Text = "Add Specifiers (Arch, Eng, Etc) to Reports" '
        '    Me.ChkSpecifiers.Checked = False  '09-06-12 
        'End If
        fraFinishReports.Enabled = True '04-17-12
        Me.ChkExtendByProb.Text = "Extend By Quote Probability" '02-04-12 Reset
        '01-27-15 JTC  Print Quote Hdr Amt when QuoteTo is Zero" Set Back
        Me.chkMfgBreakdown.Text = "Add MFG Total Breakdown to Reports" '01-27-15 JTC  Print Quote Hdr Amt when QuoteTo is Zero"
        Me.chkMfgBreakdown.CheckState = CheckState.Unchecked '01-27-15 JTC 
        Me.C1SuperTooltip1.SetToolTip(Me.chkMfgBreakdown, "Add MFG Total Breakdown To Reports") '01-27-15 JTC 
        Me.chkCustomerBreakdown.Text = "Add Cust QuoteTo Breakdown to Report" '06-06-11 "Show All Quote Header Fields" '06-06-11 "Add Cust QuoteTo Breakdown to Report"
        Me.chkCustomerBreakdown.CheckState = CheckState.Unchecked '06-06-11 
        Me.C1SuperTooltip1.SetToolTip(Me.chkCustomerBreakdown, "Add Cust QuoteTo Breakdown To Reports") '06-08-11 
        'Debug.Print(Cmd)
        If Cmd.StartsWith("Other Quote Types") Then
            'Cmd = ""
        Else
            Me.cboTypeofJob.Text = "Q" '06-14-10 Was set to P after Planned projects
        End If
        Me.fraQuoteReports.Visible = True '04-20-11
        txtSecondarySort.Text = "" '02-10-09 
        Me.chkPrtPlanLines.Visible = False '11-24-09
        If DIST Then
            ' Me.chkSpecifiers.Text = "2 Lines (Arch + Eng + Spec 3)"
        End If
        Me.txtSpecifierCode.Text = "" : Me.txtQutRealCode.Text = "" '05-07-10 
        Me.chkSalesmanPerPage.Visible = False '03-25-03 WNA
        Me.ChkSpecifiers.Visible = True '05-04-10
        Me.chkNotes.Visible = True
        Me.chkCustomerBreakdown.Visible = True
        If Me.pnlTypeOfRpt.Text <> "Realization" Then Me.chkMfgBreakdown.Visible = True '01-27-15 JTC
        Me.pnlQutRealCode.Visible = False '05-05-10 
        Me.txtQutRealCode.Visible = False '05-05-10 
        Me.pnlSpecifierCode.Visible = False '05-05-10 
        Me.txtSpecifierCode.Visible = False '05-05-10 
        Me.pnlQuoteToSls.Visible = False '05-07-10
        Me.txtQuoteToSls.Visible = False '05-07-10
        ' Case 0 'continue run reports

        If Trim(Me.cboSortPrimarySeq.Text) = "" And Me.pnlTypeOfRpt.Text <> "Realization" Then '02-04-14
            Resp = MsgBox("You must select a Primary Sort Sequence before you Continue", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Primary Sort Continue")
            Exit Sub
        End If
        '12-23-09 ))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))
        SubSeq = 0 '10-31-12 "" '01-14-10
        '12-31-14 JTC Fix Realization run after Quote Summary
        If Trim(Me.txtPrimarySortSeq.Text) = "" And Me.pnlTypeOfRpt.Text = "Realization" And Me.cboSortRealization.GetItemCheckState(3) = CheckState.Checked Then
            Me.txtPrimarySortSeq.Text = "Salesman/Customer" '12-31-14 JTC Fix Realization run after Quote Summary
        ElseIf ExcelQuoteFU = True Then '04-28-15 JTCMe.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-28-15 JTC
            Call FillSecondarySortCombo() '04-29-15 JTC 
            Me.cboSortSecondarySeq.Focus() '11-04-14 JTC
            Me.fraSortSecondarySeq.Visible = True
            Me.pnlSecondarySort.Visible = True
            Me.txtSecondarySort.Visible = True
        Else
            Me.txtPrimarySortSeq.Text = Trim(Me.cboSortPrimarySeq.Text)
        End If
        Me.lblStartBid.Text = "Bid" '10-22-13
        Me.ChkCheckBidDates.CheckState = CheckState.Unchecked
        Me.ChkCheckBidDates.Text = "Check Bid Dates when Selecting Quotes"
        Me.chkBidJobsOnly.Text = "Bid Board Jobs Only"
        Me.chkBidJobsOnly.CheckState = CheckState.Unchecked '10-23-13 Report Only Bid Jobs
        If Me.txtPrimarySortSeq.Text = "Forecasting" And Me.pnlTypeOfRpt.Text = "Quote Summary" Then
            Dim TmpMsg As String = "Brand Mfg Code - " & BrandReportMfg '02-24-14
            If ForecastAllMfg = False Then 'BrandReportMfg = "PHIL" Or BrandReportMfg = "DAYB" Or BrandReportMfg = "DAY" Or SESCO = True Then ForecastAllMfg = False else  ForecastAllMfg = True '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
                Resp = MsgBox("This will generate an Excel file of Quotes and Hold Orders." & vbCrLf & "Send this forecasting file to Philips as requested." & vbCrLf & "This will only include Philips Brands and Dollars." & vbCrLf & TmpMsg & vbCrLf & "Usually Set the Quote Status box to SUBMIT,BID or your choice", MsgBoxStyle.OkCancel, "Forecasting Report") '05-28-15 JTC02-24-14 -17-1411-15-13
                If Resp = vbCancel Then Exit Sub '11-15-13
            Else ' ForcastAllMfg
                Resp = MsgBox("This will generate an Excel file of Quotes for Forecasting to MFGs." & vbCrLf & "Send this forecasting file to your MFG as requested." & vbCrLf & "This will only include MFG Codes/Brands entered and their Dollars." & vbCrLf & TmpMsg & vbCrLf & "Usually Set the Quote Status box to SUBMIT,BID or your choice", MsgBoxStyle.OkCancel, "Forecasting Report") '02-28-15 JTC 02-24-14 -17-1411-15-13
                If Resp = vbCancel Then Exit Sub '11-15-13
                If BrandList.Trim = "" Then
                    Me.chkBrandReport.CheckState = CheckState.Checked '05-15-15 JTC Loads BrandList & lets them Change on ForecastAllMFG
                    If BrandList.Trim <> "" Then '05-15-15 JTC 
                        BrandList = InputBox(BrandList.ToUpper & vbCrLf & "Modify this Brand List as needed. (KEEN,GLOB,ABC)." & vbCrLf & vbCrLf & "Forecasting Brands", "Forecasting Brands", BrandList.ToUpper)
                    Else
                        BrandList = InputBox("This requires Mfg Brand codes on Quote Lineitems." & vbCrLf & "Enter ALL or List Codes (KEEN,GLOB,ABC)." & vbCrLf & vbCrLf & "Forecasting Brands", "Forecasting Brands", "ALL")
                    End If
                End If
                End If
                VQRT2.RepType = 0 ' VQRT2.RptMajorType.RptFollowBy '11-06-13
                '05-14-15 JTC Allow Quote Forecasting for all MFGs not just Philips
                '05-14-15 JTC If BrandReportMfg = "COOP" Then GoTo NoForecasting '10-23-13
                '10-22-13  GetRepNumber(cboMfgLookup.Text.Trim)
                Me.mnuBrandReport.Text = "Brand Reporting - On" '11-06-13
                Me.mnuBrandMfgChg.Text = "Brand Mfg Code - " & BrandReportMfg '11-06-13 XXXX Brand Mfg Code - XXXX
                Me.chkBrandReport.Text = "Brand Reporting - " & BrandReportMfg '11-06-13 XXXX Brand Mfg Code - XXXX"
                Me.chkBrandReport.Visible = True '11-06-13
                If ForecastAllMfg = False Then 'BrandReportMfg = "PHIL" Or BrandReportMfg = "DAYB" Or BrandReportMfg = "DAY" Or SESCO = True Then ForecastAllMfg = False else  ForecastAllMfg = True '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
                    Me.chkBrandReport.CheckState = CheckState.Checked '10-20-13
                    Me.fraSortSecondarySeq.Visible = False '11-06-13
                    Me.lblStartEntry.Enabled = False '11-06-13
                    Me.DTPickerStartEntry.Enabled = False '11-06-13
                    Me.DTPicker1EndEntry.Enabled = False '11-06-13
                    Me.chkBlankBidDates.Visible = False '11-06-13
                    Me.mnuBrandReport.Enabled = True '11-06-13 
                    Me.mnuBrandMfgChg.Enabled = True '11-06-13 
                Me.txtStatus.Text = "SUBMIT,BID" '05-28-15 JTC
                    Me.lblStartBid.Text = "Deliver"
                    Me.ChkCheckBidDates.CheckState = CheckState.Checked
                    Dim A As String = Me.ChkCheckBidDates.Text
                    Me.ChkCheckBidDates.Text = "Check Deliver Dates when Selecting Quotes"
                    chkBlankBidDates.Visible = False '11-06-13 
                    Me.chkBidJobsOnly.CheckState = CheckState.Checked '10-22-13 Report Only Bid Jobs
                    Me.chkBidJobsOnly.Text = "Delivery Date Jobs Only"
                    RepCustNumber = GetRepNumber(BrandReportMfg) ' 10-28-13 PHIL or "DAYB") ' BrandReportMfg)
                    Call SetDefaultEstDelDate() '02-17-14 JTC 
                    If DefChgEstDelDateCodes.Trim <> "" Then Me.txtStatus.Text = DefChgEstDelDateCodes.ToUpper '02-17-14 "SUBMIT,GOT"
                Else
                    'means ForecastAllMfg = True Use All Selections Then 'BrandReportMfg = "PHIL" Or BrandReportMfg = "DAYB" Or BrandReportMfg = "DAY" Or SESCO = True Then ForecastAllMfg = False else  ForecastAllMfg = True '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
                End If
        Else
                Me.chkBrandReport.Visible = False '01-14-14 True '11-06-13
        End If
NoForecasting:  '10-23-13If BrandReportMfg = "COOP" then GoTo NoForecasting:'10-23-13
        '05-05-10 Me.txtQutRealCode.Visible = True : Me.pnlQutRealCode.Visible = True '12-11-08
        If Me.pnlTypeOfRpt.Text = "Other Quote Types" Then
            Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
            'Me.TtxtSortSelV.Text = SelectionText '02-24-09
            If Me.cboSortPrimarySeq.Text = "Submittals" Then Me.cboTypeofJob.Text = "T" '06-18-10
            If Me.cboSortPrimarySeq.Text = "Planned Projects" Then Me.cboTypeofJob.Text = "P" '06-18-10
            If Me.cboSortPrimarySeq.Text = "Other Quote Types" Then Me.cboTypeofJob.Text = "O" '06-18-10
            Cmd = "Other Quote Types/" & Me.cboSortPrimarySeq.Text '06-18-10 Me.cboSortPrimarySeq.Text 'public Save
            Me.pnlTypeOfRpt.Text = "Quote Summary" '06-18-10 
            Call cmdReportQuote_Click(cmdReportQuote, New System.EventArgs())

            Exit Sub
        End If
        If SESCO = True Or Me.pnlTypeOfRpt.Text = "Realization" Then

            '08-08-12 Me.txtPrimarySortSeq.Text = "Name Code" '02-11-12 Me.pnlTypeOfRpt.Text = "Realization"'frmQuoteRpt.txtPrimarySortSeq.Text = "Descending Sales Dollars" ' Name Code" '02-11-12
            '08-08-12 JTC Realization Moved Up From Below and fixed pnlQutRealCode.Text too long 
            If Me.txtPrimarySortSeq.Text = "Salesman/Customer" Then '10-30-02 WNA
                Me.pnlQutRealCode.Text = "Select SLS Code" '08-08-12'09-16-11 "Must be set at ALL" '12-11-08 "Single Salesman Code"
                Me.txtQutRealCode.Visible = False : Me.pnlQutRealCode.Visible = False '12-11-08
            ElseIf Me.txtPrimarySortSeq.Text = "Manufacturer" Then  '03-26-03 WNA
                Me.pnlQutRealCode.Text = "Select MFG Codes" '09-16-11"Single MFG Code"
                Me.txtPrimarySortSeq.Text = "Name Code" '08-08-12
            ElseIf ExcelQuoteFU = True Then '04-28-15 JTC Me.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-27-15 JTC
                'Me.txtPrimarySortSeq.Text = "Bid Date" '04-28-2015 JTC
            Else                       '09-16-11 Single
                Me.pnlQutRealCode.Text = "Select " & Trim(Me.txtPrimarySortSeq.Text) '08-08-12 Text too long & " Codes"
                Me.txtPrimarySortSeq.Text = "Name Code" '08-08-12
            End If
            '07-15-14 JTC All Specifiers on so RealALL = True 
            If RealArchitect = True And RealEngineer = True And RealLtgDesigner = True And RealSpecifier = True And RealContractor = True And RealOther = True Then RealALL = True Else RealALL = False '07-15-14  All Specifiers
            '07-15-14 JTC RealCustomer Only
            If RealManufacturer = False And RealCustomer = True And RealQuoteTOOther = False And RealSLSCustomer = False Then Me.pnlQutRealCode.Text = "Select Cust Code"
            '07-15-14 JTC RealManufacturer Only
            If RealManufacturer = True And RealCustomer = False And RealQuoteTOOther = False And RealSLSCustomer = False Then Me.pnlQutRealCode.Text = "Select MFG Code" '"Select MFG Code"
            If RealALL = True And RealManufacturer = False And RealCustomer = False And RealQuoteTOOther = False And RealSLSCustomer = False Then Me.pnlQutRealCode.Text = "Select Specifier Code" '07-16-14  12-05-17 spelling fix on specifier
            Me.txtQutRealCode.Visible = True : Me.pnlQutRealCode.Visible = True '05-05-10
            If Me.txtPrimarySortSeq.Text = "Salesman/Customer" Then '10-30-12 JTC Don't show
                Me.pnlQutRealCode.Text = "Select SLS Code" '08-08-12'09-16-11 "Must be set at ALL" '12-11-08 "Single Salesman Code"
                Me.txtQutRealCode.Visible = False : Me.pnlQutRealCode.Visible = False '12-11-08
                Me.chkSalesmanPerPage.Text = "1 SLS/Page" '06-17-13 JTC Add "1 SLS/Page" on "Salesman/Customer" Realization
                Me.chkSalesmanPerPage.Visible = True '06-17-13 
                '10-30-12 
                Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
                'Me.txtSortSeqCriteria.Text = SelectionText '02-24-09
                '02-24-09  Me.txtSortSeq.Text = txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
                Me.txtSortSeqV.Text = Me.txtSortSeq.Text 'txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
                System.Windows.Forms.Application.DoEvents() '11-25-09 
                Me.tabQrt.SelectedIndex = 1 '11-24-09
                Call tabQRT_TabActivate(1)
                Exit Sub
            End If
            Me.pnlQuoteToSls.Visible = True '05-07-10
            Me.txtQuoteToSls.Visible = True '05-07-10
            Me.fraQuoteLineReports.Visible = False '08-18-09 JH
            Me.fraQuoteReports.Visible = True '08-18-09 JH
            Enable = 0 : Call EnableOrDisable2(Enable)
            Me.ChkSpecifiers.Visible = False '05-04-10
            Me.chkNotes.Visible = True '02-11-12 JTC Add Notes to Realization
            '07-24-14 JTC If SESCO = True And My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") Then '02-25-12
            If My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") Then '02-25-12
                '07-31-14 JTC No Good Dim LastItem As Int16 = Me.cboSortPrimarySeq.Items.Count - 1 '03-29-12 
                'Debug.Print(Me.cboSortPrimarySeq.Items(3).text)
                'If Me.cboSortPrimarySeq.Items(LastItem).text = "SESCO Job List Report" Then
                '    SESCO = True '07-24-14
                'Else
                '    Me.cboSortPrimarySeq.Items.Add("SESCO Job List Report") '02-25-12
                'End If
                ''If Me.cboSortPrimarySeq.ItMe.txtPrimarySortSeq.Text = "SESCO Job List Report" Then
                'Me.txtPrimarySortSeq.Text = "SESCO Job List Report - Job Name Sequence"
                'fraFinishReports.Enabled = False '04-13-12
                'chkShowLatestCust.Visible = True '04-13-12
                '07-31-14 JTC Above Logic No Good It changed txtPrimarySortSeq.text
                'SESCO Job List Report '07-31-14 JTC added Below ' If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True old method 
                If cboSortRealization.GetItemCheckState(12) = CheckState.Checked Then '12 is "SESCO Job List Report"
                    SESCO = True
                    Me.txtPrimarySortSeq.Text = "SESCO Job List Report - Job Name Sequence"
                    fraFinishReports.Enabled = False '04-13-12
                    chkShowLatestCust.Visible = True '04-13-12
                Else
                    SESCO = False
                End If
            End If 'End Sesco
            If RealOther = True Or RealArchitect = True Or RealEngineer = True Or RealLtgDesigner = True Or RealSpecifier = True Or RealContractor = True Or RealOther = True Then '02-04-12
                'Me.ChkExtendByProb.Text = "Extend By Quote Probability" '02-04-12
                Me.ChkExtendByProb.Text = "Extend Specifiers By Influence %" '02-04-12
            End If


            Me.chkCustomerBreakdown.Text = "Show All Quote Header Fields " '06-06-11 "Add Cust QuoteTo Breakdown to Report"
            Me.chkCustomerBreakdown.Visible = True '06-06-11
            Me.C1SuperTooltip1.SetToolTip(Me.chkCustomerBreakdown, "Show All Of The Quote Header Fields (Arch,Eng, Etc.) For Reporting") '06-08-11
            If RealCustomer = True Or RealWithOneMfgCust = True Then '03-25-13               'If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Th
                'Turn on '03-29-12 Need to Add a Check box "Only Show the Latest Quote To Each Customer
                Me.chkShowLatestCust.Visible = True '03-29-12
                Me.chkShowLatestCust.CheckState = CheckState.Checked '03-29-12
            End If
            If RealCustomer = True And RealManufacturer = False And RealQuoteTOOther = False And RealSLSCustomer = False And RealArchitect = False And RealEngineer = False And RealLtgDesigner = False And RealSpecifier = False And RealContractor = False And RealOther = False Then '01-16-15 JTC
                'RealCustomer = True only Change '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero"
                Me.chkMfgBreakdown.Text = " Print Quote Hdr Amt when QuoteTo is Zero" ' "Add MFG Total Breakdown to Reports" '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero"
                Me.chkMfgBreakdown.Visible = False '01-27-15 JTC 
                Me.C1SuperTooltip1.SetToolTip(Me.chkMfgBreakdown, "Print Quote Header Amount when QuoteTo Amount is Zero") ' Add MFG Total Breakdown To Reports") '01-27-15 JTC 
            End If
            '08-08-12 moved up If Me.txtPrimarySortSeq.Text = "Salesman/Customer" Then '10-30-02 WNA
            '    Me.pnlQutRealCode.Text = "Select Salesman Code" ''09-16-11 "Must be set at ALL" '12-11-08 "Single Salesman Code"
            '    Me.txtQutRealCode.Visible = False : Me.pnlQutRealCode.Visible = False '12-11-08
            'ElseIf Me.txtPrimarySortSeq.Text = "Manufacturer" Then  '03-26-03 WNA
            '    Me.pnlQutRealCode.Text = "Select MFG Codes" '09-16-11"Single MFG Code"
            'Else                       '09-16-11 Single
            '    Me.pnlQutRealCode.Text = "Select " & Trim(Me.txtPrimarySortSeq.Text) '08-08-12 Text too long & " Codes"
            'End If
            Me.chkDetailTotal.Visible = True
            Me.chkSalesmanPerPage.Visible = True
            Me.chkSalesmanPerPage.Text = "1 Code/Page" '03-25-03 WNA
            'Me.ChkSpecifiers.Text = "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
            '02-11-12 Use ChkSpecifiers.Text = "Sort Report by Descending Dollar 
            '07-27-12 Me.ChkSpecifiers.Text = "Sort Report by Descending Dollar" '02-11-12 " "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
            Me.ChkSpecifiers.Visible = True '02-11-12 
            '01-28-09 Set Tab 1 Me.txtPrimarySortSeq.Text = "Name Code" '02-11-12 Me.pnlTypeOfRpt.Text = "Realization"'frmQuoteRpt.txtPrimarySortSeq.Text = "Descending Sales Dollars" ' Name Code" '02-11-12
            Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
            Me.txtSortSeqCriteria.Text = SelectionText '02-24-09
            '02-24-09  Me.txtSortSeq.Text = txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
            Me.txtSortSeqV.Text = Me.txtSortSeq.Text 'txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
            If VB.Right(Me.txtSortSeq.Text, 1) = "/" Then Me.txtSortSeq.Text = Replace(Me.txtSortSeq.Text, "/", "")
            '07-24-14 JTC Fix SESCO Job Listing Excel Report to not show SecondarySeq
            If Me.txtPrimarySortSeq.Text = "SESCO Job List Report - Job Name Sequence" Then '07-24-14
                System.Windows.Forms.Application.DoEvents() '11-25-09 
                Me.tabQrt.SelectedIndex = 1 '11-24-09
                Call tabQRT_TabActivate(1)
                Exit Sub
            End If

            Call FillSecondarySortCombo() '07-26-12
            If Me.txtPrimarySortSeq.Text = "Forecasting" Then '11-06-13 
            Else
                Me.fraSortSecondarySeq.Visible = True
                Me.cboSortSecondarySeq.Visible = True
                Me.cboSortSecondarySeq.Focus()
            End If
            'If RealOther = True Or RealArchitect = True Or RealEngineer = True Or RealLtgDesigner = True Or RealSpecifier = True Or RealContractor = True Or RealOther = True Then '02-04-12
            '12-31-14 If SESCO = False And frmQuoteRpt.txtPrimarySortSeq.Text.StartsWith("Salesman") = False And frmQuoteRpt.pnlTypeOfRpt.Text <> "Quote Summary" Then '05-21-13
            'If Me.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-27-15 JTC
            'Debug.Print(Me.cboSortPrimarySeq.Text)
            If SESCO = False And RealCustomer = False And RealManufacturer = False And RealALL = False And RealOther = False And RealArchitect = False And RealEngineer = False And RealLtgDesigner = False And RealSpecifier = False And RealContractor = False And RealOther = False And RealWithOneMfgCust = False And Me.cboSortPrimarySeq.Text <> "Salesman/Customer" And ExcelQuoteFU = False Then '04-28-15 JTC 04-22-15 JTC12-31-14 JTC No Chg And frmQuoteRpt.txtPrimarySortSeq.Text.StartsWith("Salesman") = False
                Resp = MsgBox("You must select some Primary Options before you Continue", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Realization Selections") '03-24-13
                Exit Sub
            End If
            '05-16-13 JTC If Only MFG Turn on ChkBrand
            If RealManufacturer = True And SESCO = False And RealCustomer = False And RealALL = False And RealOther = False And RealArchitect = False And RealEngineer = False And RealLtgDesigner = False And RealSpecifier = False And RealContractor = False And RealOther = False And RealWithOneMfgCust = False Then '06-28-13
                Me.chkBrandReport.Visible = True '05-16-13 
            End If
            '11-04-14 JTC Turn these off on Realization they don,t work
            Me.ChkSpecifiers.Visible = False
            Me.ChkExtendByProb.Visible = False
            Me.chkBlankLine.Visible = False
            Me.chkSlsFromHeader.Visible = False
            Me.ChkQuoteNoSpecifiers.Visible = False
            'If Me.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-27-15 JTC 
            If ExcelQuoteFU = True Then '04-28-15 JTC
                chkShowLatestCust.Visible = True
                Me.chkCustomerBreakdown.Visible = False
                Me.chkDetailTotal.Visible = False
                Me.chkSalesmanPerPage.Visible = False
                Me.ChkCheckBidDates.CheckState = CheckState.Checked
                'Me.cboSortSecondarySeq.Text 
                '06-05-12 If Me.DTPicker1StartBid.Value = CDate("01/01/1900") Then '06-05-12 Dim NewDate As Date = Now '02-03-12 '06-05-12 NewDate = NewDate.AddDays(-7) 'AddYears : NewDate = NewDate.AddDays(-1)
                Me.DTPicker1StartBid.Enabled = True
                Me.DTPicker1EndBid.Enabled = True
                Me.chkShowLatestCust.Visible = True '04-27-15
                Me.chkShowLatestCust.CheckState = CheckState.Checked '04-27-15 JTC
                Me.cboSortSecondarySeq.SelectedIndex = 3
                Me.txtSecondarySort.Text = "Bid Date" : Me.cboSortSecondarySeq.Text = "Bid Date" '04-27-15 JTC
                'Me.txtSortSeqV.Text = "Excel Quote FollowUp" '04-28-15 Me.txtSortSeq.Text 'txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
                'Me.pnlTypeOfRpt.Text = "Excel Quote FollowUp" '04-28-15 JTC
            End If
            Exit Sub '07-26-12 
            'End of If SESCO = True Or Me.pnlTypeOfRpt.Text = "Realization" Then
            '**************************************************************************

        ElseIf Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '11-24-09 
            'No If Me.cboTypeofJob.Text = "Q" Then '06-20-10 No Line Items
            '    Me.chkPrtPlanLines.Visible = False '06-20-10 'Me.chkPrtPlanLines.Text = "Print Spec Credit Lines" '"Print Planned Project Lines"
            '    Me.cboLinesInclude.Visible = False '06-20-10 'Include All Lines on Job
            'End If
            Me.fraQuoteLineReports.Visible = False '08-18-09 JH
            Me.fraQuoteReports.Visible = True '08-18-09 JH
            Call SetPrimarySortValues()
            'If Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Salesman" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Status" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Specifier Credit" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Retrieval Code" Then
            '11-24-09If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Or VQRT2.RepType = VQRT2.RptMajorType.RptStatus Or VQRT2.RepType = VQRT2.RptMajorType.RptSpecif Or VQRT2.RepType = VQRT2.RptMajorType.RptRetrieval Then
            Me.fraSortSecondarySeq.Visible = True
            Me.pnlSecondarySort.Visible = False
            Me.txtSecondarySort.Visible = False
            Me.fraFinishReports.Visible = False '11-24-09
            If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Then '03-25-03 WNA
                Me.chkSalesmanPerPage.Text = "1 SLS/Page"
                Me.chkSalesmanPerPage.Visible = True
            End If
            'Call FillSecondarySortCombo()
            'Me.cboSortSecondarySeq.Focus()
            Me.chkPrtPlanLines.Text = "Print Spec Credit Lines" '"Print Planned Project Lines"
            Me.chkPrtPlanLines.Visible = True '11-24-09
            Me.chkPrtPlanLines.CheckState = CheckState.Checked '11-24-09 
            Me.cboTypeofJob.Text = "S" '11-24-09
            '09-21-12 
            Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
            'Me.txtSortSeqCriteria.Text = SelectionText '02-24-09
            '02-24-09  Me.txtSortSeq.Text = txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
            Me.txtSortSeqV.Text = Me.txtSortSeq.Text 'txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
            System.Windows.Forms.Application.DoEvents() '11-25-09 
            Me.tabQrt.SelectedIndex = 1 '11-24-09
            Call tabQRT_TabActivate(1)
            Me.cboTypeofJob.Text = "S" '11-25-09"Terr Spec Credit Report"
            '11-24-09 End If
            'End Spec Credit Report *****************************************************

            '04-04-13 Never Goes thru below
        ElseIf Me.pnlTypeOfRpt.Text = "Planned Projects" Then '11-23-09    
            Me.fraQuoteLineReports.Visible = False '08-18-09 JH
            Me.fraQuoteReports.Visible = True '08-18-09 JH
            Call SetPrimarySortValues()
            'If Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Salesman" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Status" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Specifier Credit" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Retrieval Code" Then
            '11-24-09If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Or VQRT2.RepType = VQRT2.RptMajorType.RptStatus Or VQRT2.RepType = VQRT2.RptMajorType.RptSpecif Or VQRT2.RepType = VQRT2.RptMajorType.RptRetrieval Then
            Me.fraSortSecondarySeq.Visible = True
            Me.pnlSecondarySort.Visible = True
            Me.txtSecondarySort.Visible = True
            Me.fraFinishReports.Visible = False '11-24-09

            If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Then '03-25-03 WNA
                Me.chkSalesmanPerPage.Text = "1 SLS/Page"
                Me.chkSalesmanPerPage.Visible = True
            End If
            Call FillSecondarySortCombo()
            Me.cboSortSecondarySeq.Focus()
            Me.chkPrtPlanLines.Text = "Print Planned Project Lines"
            Me.chkPrtPlanLines.Visible = True '11-24-09
            Me.chkPrtPlanLines.CheckState = CheckState.Checked '11-24-09 
            Me.cboTypeofJob.Text = "P" '11-24-09
            Me.tabQrt.SelectedIndex = 1 '11-24-09
            Call tabQRT_TabActivate(1)
            '11-24-09 End If
            'End Planned Projects *****************************************************

        ElseIf Me.pnlTypeOfRpt.Text = "Quote Summary" Or Me.pnlTypeOfRpt.Text = "Project Shortage Report" Then '05-16-10       'Not Realization
            Me.fraQuoteLineReports.Visible = False '08-18-09 JH
            Me.fraQuoteReports.Visible = True '08-18-09 JH
            Call SetPrimarySortValues()
            'If Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Salesman" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Status" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Specifier Credit" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Retrieval Code" Then
            If Cmd.StartsWith("Other Quote Types/") Then
                If Me.cboTypeofJob.Text = "T" Then Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "Submittals  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
                If Me.cboTypeofJob.Text = "P" Then Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "Planned Projects  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
                If Me.cboTypeofJob.Text = "O" Then Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "Other Quote Types  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
                ' If Me.pnlTypeOfRpt.Text = "Other Quote Types" Then
                Cmd = "" '06-18-10 "Other Quote Types/" & Me.cboSortPrimarySeq.Text '06-18-10 Me.cboSortPrimarySeq.Text 'public Save
            End If
            '11-04-14 JTC Turn these off on Realization they don,t work
            Me.ChkExtendByProb.Visible = True
            Me.chkBlankLine.Visible = True
            ' Me.chkSlsFromHeader.Visible = False
            Me.ChkQuoteNoSpecifiers.Visible = True

            If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Or VQRT2.RepType = VQRT2.RptMajorType.RptStatus Or VQRT2.RepType = VQRT2.RptMajorType.RptSpecif Or VQRT2.RepType = VQRT2.RptMajorType.RptRetrieval Or VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Or VQRT2.RepType = VQRT2.RptMajorType.RptEnteredBy Then '05-14-13 eNTEREDbY 03-01-12 Followedby
                If Me.txtPrimarySortSeq.Text = "Forecasting" Then '11-06-13 
                Else
                    Me.fraSortSecondarySeq.Visible = True
                    Me.pnlSecondarySort.Visible = True
                    Me.txtSecondarySort.Visible = True
                End If
                If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Or VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Or VQRT2.RepType = VQRT2.RptMajorType.RptEnteredBy Then '05-14-13 03-03-12
                    Me.chkSalesmanPerPage.Text = "1 SLS/Page"
                    If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Then Me.chkSalesmanPerPage.Text = "1 FollowBy/Page" : Me.chkSalesmanPerPage.CheckState = CheckState.Checked '03-03-10
                    If VQRT2.RepType = VQRT2.RptMajorType.RptEnteredBy Then Me.chkSalesmanPerPage.Text = "1 EnteredBy/Page" : Me.chkSalesmanPerPage.CheckState = CheckState.Checked '05-14-13
                    Me.chkSalesmanPerPage.Visible = True
                    Me.chkDetailTotal.Visible = True '01-12-15 JTC If run Realization B/4 Quote Summary then Me.ChkDetailTotalsOnly.Visible is False
                End If
                If My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") And VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Then '03-09-12
                    SESCO = True 'Me.cboSortPrimarySeq.Items.Add("SESCO Job List Report")   Me.txtPrimarySortSeq.Text = "SESCO Job List Report - Job Name Sequence"  'Me.txtPrimarySortSeq.Text = "Name Code"
                Else
                    SESCO = False
                End If
                Call FillSecondarySortCombo()
                Me.cboSortSecondarySeq.Focus()
            Else 'No Secondary
                'Forecast goes Here 
                If Me.pnlTypeOfRpt.Text = "Project Shortage Report" Then '10-23-02 WNAProject Shortage Report
                    Enable = 0
                Else
                    Enable = 1
                End If
                Call EnableOrDisable2(Enable)
                Me.chkDetailTotal.Visible = False
                Me.chkSalesmanPerPage.Visible = False
                '01-28-09 Set Tab 1
                Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
                Me.txtSortSeqCriteria.Text = SelectionText '02-24-09
                '02-24-09 Me.txtSortSeq.Text = txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
                Me.txtSortSeqV.Text = Me.txtSortSeq.Text 'txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
                If VB.Right(Me.txtSortSeq.Text, 1) = "/" Then Me.txtSortSeq.Text = Replace(Me.txtSortSeq.Text, "/", "")
                'If Me.pnlTypeOfRpt.Text = "Quote Summary" And txtPrimarySortSeq.Text = "SESCO Job List Report" Then
                '    'If My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") Then '02-25-12
                '    'frmQuoteRpt.cboSortPrimarySeq.Items.Add("SESCO Job List Report") '02-25-12
                '    SESCO = True
                'Else
                '    SESCO = False
                'End If
                '11-04-14 JTC
                If Me.pnlTypeOfRpt.Text = "Quote Summary" And VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj Then '11-04-14 JTC
                    Call FillSecondarySortCombo()
                    Me.cboSortSecondarySeq.Focus() '11-04-14 JTC
                    Me.fraSortSecondarySeq.Visible = True
                    Me.pnlSecondarySort.Visible = True
                    Me.txtSecondarySort.Visible = True
                Else
                    Dim SaveTypeQ As String = cboTypeofJob.Text ' = "Q" '04-26-12
                    Me.tabQrt.SelectedIndex = 1
                    Call tabQRT_TabActivate(1)
                    cboTypeofJob.Text = SaveTypeQ ' As String = cboTypeofJob.Text ' = "Q" '04-26-12
                    'Me.fraRptSel.Visible = True
                    'Me.optSelectNewRecords.Checked = CheckState.Checked
                    'Me.cmdFmtOK.Focus()
                End If '11-04-14 JTC

            End If
        ElseIf Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then '09-25-12
            'Me.fraQuoteLineReports.Visible = False '08-18-09 JH
            'Me.fraQuoteReports.Visible = True '08-18-09 JH
            'Call SetPrimarySortValues()
            'If Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Salesman" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Status" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Specifier Credit" Or Trim$(frmQuoteRpt.txtPrimarySortSeq.Text) = "Retrieval Code" Then
            '11-24-09If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Or VQRT2.RepType = VQRT2.RptMajorType.RptStatus Or VQRT2.RepType = VQRT2.RptMajorType.RptSpecif Or VQRT2.RepType = VQRT2.RptMajorType.RptRetrieval Then
            'Me.fraSortSecondarySeq.Visible = True
            Me.pnlSecondarySort.Visible = False
            Me.txtSecondarySort.Visible = False
            'Me.fraFinishReports.Visible = False '11-24-09
            If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Then '03-25-03 WNA
                Me.chkSalesmanPerPage.Text = "1 SLS/Page"
                Me.chkSalesmanPerPage.Visible = True
            End If
            'Call FillSecondarySortCombo()
            'Me.cboSortSecondarySeq.Focus()
            'Me.chkPrtPlanLines.Text = "Print Spec Credit Lines" '"Print Planned Project Lines"
            'Me.chkPrtPlanLines.Visible = True '11-24-09
            'Me.chkPrtPlanLines.CheckState = CheckState.Checked '11-24-09 
            'Me.cboTypeofJob.Text = "S" '11-24-09
            '09-21-12 
            Me.txtSortSeq.Text = Me.pnlTypeOfRpt.Text & "  Sort By = " & txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text '02-24-09
            'Me.txtSortSeqCriteria.Text = SelectionText '02-24-09
            '02-24-09  Me.txtSortSeq.Text = txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
            Me.txtSortSeqV.Text = Me.txtSortSeq.Text 'txtPrimarySortSeq.Text & " / " & txtSecondarySort.Text
            System.Windows.Forms.Application.DoEvents() '11-25-09 
            Me.tabQrt.SelectedIndex = 1 '11-24-09
            Call tabQRT_TabActivate(1)
            'End "Product Sales History - Line Items"  Report *****************************************************
        ElseIf Me.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-28-15 JTC
            'Do Nothing
        Else 'Line Item Reporting '08-18-09 JH Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items"
            'Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items"
            Me.fraQuoteReports.Visible = False
            Me.fraQuoteReports.Visible = True '04-15-11

            Me.fraQuoteLineReports.Visible = True
            If DIST Then
                Me.fraSalesorCost.Visible = True '03-29-12
                optSalesorCost_Cost.Text = "Cost Dollars" '02-15-10
            End If
            Me.tabQrt.SelectedIndex = 1
            Call tabQRT_TabActivate(1)
        End If
        '07-24-14 JTC SESCO Forecasting No dollars
        If Me.txtPrimarySortSeq.Text = "Forecasting" And Me.pnlTypeOfRpt.Text = "Quote Summary" Then
            If My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") Then '07-24-14 JTC SESCO Forecasting No dollars
                SESCO = True
            End If
        End If
        'Me.fraSortPrimarySeq.Visible = False

        ' Case 1 'cancel back to report options
        '     Me.pnlTypeOfRpt.Visible = False
        '     Me.pnlTypeOfRpt.Text = "" '10-23-02 WNA
        '     Me.fraSortPrimarySeq.Visible = False
        '     Me.fraReportCmdSelection.Visible = True
        '    Me.pnlPrimarySortSeq.Visible = False
        '    Me.txtPrimarySortSeq.Visible = False
        '   Me.fraRptSel.Visible = False

        'End Select


    End Sub

    Private Sub cmdPrimarySeqCancel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrimarySeqCancel1.Click
        ' Case 1 'cancel back to report options
        SESCO = False '02-25-12
        Me.lblJobName.Text = "Job Name Search String" '06-22-12 
        Me.cboSortRealization.Visible = False '01-18-12
        'Me.cboSortRealization.SendToBack() '01-30-12
        cboSortPrimarySeq.Visible = True '01-30-12
        Me.pnlTypeOfRpt.Visible = False
        Me.pnlTypeOfRpt.Text = "" '10-23-02 WNA

        Me.fraSortPrimarySeq.Visible = False
        Me.fraReportCmdSelection.Visible = True
        Me.pnlPrimarySortSeq.Visible = False
        Me.txtPrimarySortSeq.Visible = False
        '01-30-12

        'Me.fraRptSel.Visible = False

    End Sub


    Private Sub CmdRunReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdRunReport.Click
        Me.CmdRunReport.Enabled = False '10-13-14 JTC Fix Dbl Click on Reports
        System.Windows.Forms.Application.DoEvents() '10-14-14 JTC
        If Me.pnlTypeOfRpt.Text.Trim = "" Then
            Resp = MsgBox("Go to the Left Tab to start Reports Process. Select a Type of Report First", MsgBoxStyle.OkCancel)
            If Resp = vbCancel Then GoTo EndExit '10-13-14 JTC Exit Sub '03-22-13
            Me.tabQrt.SelectedIndex = 0 : GoTo EndExit '10-13-14 JTC Exit Sub '10-12-10 
        End If
        '01-06-12 Added MaxLength Here
        MaxNameLength = Val(Me.rbnMaxNameTxt.Text) '12-23-12 45   Public As Int16
        MaxJobLength = Val(Me.rbnMaxJobTxt.Text) '12-23-12 40   Public As Int16
        If MaxNameLength < 10 Or MaxNameLength > 45 Then MaxNameLength = 45 : Me.rbnMaxNameTxt.Text = "45" '12-23-12
        If MaxJobLength < 10 Or MaxJobLength > 40 Then MaxJobLength = 40 : Me.rbnMaxJobTxt.Text = "40" '12-23-12
        If Me.chkPrintGrayScale.Checked = True Then PrtGrayScale = True Else PrtGrayScale = False 'My.Settings.PrintGrayScale() '01-18-13
        If Me.chkPrintGrayScale.Checked = True Then Me.RibbonTab6.Text = "Print GrayScale" Else Me.RibbonTab6.Text = "Print Color" '01-18-13
        If PrtGrayScale = True Then
            LemonChiffon = Color.LightGray '01-18-13    54 Times   
            AntiqueWhite = Color.DimGray '01-18-13      19 Times   
            LightGray = Color.DarkGray '01-18-13        11 Times
            LightSkyBlue = Color.LightGray '01-18-13    12 Times   
        Else
            LemonChiffon = Color.LemonChiffon '01-18-13
            AntiqueWhite = Color.AntiqueWhite '01-18-13   19 Times   
            LightGray = Color.LightGray '01-18-13         11 Times   
            LightSkyBlue = Color.LightSkyBlue '01-18-13   12 Times   
        End If
        '01-02-12 DecFormat As String = "########0.00" '01-01-12 NoCents "########0") DecFormat for Nocents Whole Dollars
        If chkWholeDollars.Checked = True Then DecFormat = "########0" Else DecFormat = "########0.00" '01-02-12 DecFormat As String = "########0.00" '01-01-12 NoCents "########0") DecFormat for Nocents Whole Dollars
        If chkAddCommas.Checked = True Then
            If chkWholeDollars.Checked = True Then DecFormat = "###,###,##0" Else DecFormat = "###,###,##0.00" '01-06-12 
        End If
        If chkAddDollarSign.Checked = True Then DecFormat = "$" & DecFormat '01-06-12
        '02-01-09
        If SESCO = True And Me.pnlTypeOfRpt.Text = "Realization" Then '02-26-12 
            '02-25-12 And My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") Then '02-22-12
            Me.pnlTypeOfRpt.Text = "Realization"
            RealCustomer = True : RealArchitect = True : RealEngineer = True : RealLtgDesigner = True : RealSpecifier = True : RealContractor = True '02-22-12
            Me.chkCustomerBreakdown.CheckState = CheckState.Checked
            Call PrintSESCOJobListRealReportQutTO()
            GoTo EndExit '04-22-15 JTC
        End If
        '04-22-15 JTC
        If ExcelQuoteFU = True Then '04-28-15 JTC PrimarySeq.Text = "Excel Quote FollowUp" And Me.pnlTypeOfRpt.Text = "Realization" Then
            Me.pnlTypeOfRpt.Text = "Realization"
            RealCustomer = True : RealArchitect = True : RealEngineer = True : RealLtgDesigner = True : RealSpecifier = True : RealContractor = True '02-22-12
            Me.chkCustomerBreakdown.CheckState = CheckState.Checked
            Call ExcelQuoteFollowUp() '04-24-15 JTC Old PrintSESCOJobListRealReportQutTO()
            GoTo EndExit '04-22-15 JTC
        End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor 'HourGlass 12-04-10
        'Update the data to the dataset:

        If Me.pnlTypeOfRpt.Text = "Realization" And (Me.txtSecondarySort.Text = "Spread Sheet by Month" Or Me.txtSecondarySort.Text = "Spread Sheet by Year") Then '06-22-15 JTC 05-17-13 
            '06-30-15 JTC Test Start and End Year test
            Dim StartYear As Date = VB6.Format(Me.DTPickerStartEntry.Text, "yyyy/MM")
            Dim EndYear As Date = VB6.Format(Me.DTPicker1EndEntry.Text, "yyyy/MM") 'StartYear.AddYears(1) ' VB6.Format(frmQuoteRpt.DTPicker1EndEntry.Text, "yyyy")
            If Me.pnlTypeOfRpt.Text = "Realization" And Me.txtSecondarySort.Text = "Spread Sheet by Month" Then
                If Format(EndYear, "yyyy") <> Format(StartYear, "yyyy") Then 'Format(StartYear.AddYears(I), "yyyy")
                    MsgBox("Start Year and End Year must be Equal for the twelve month spreadsheet Report! ****** Fix Dates", MsgBoxStyle.Critical + MsgBoxStyle.RetryCancel, "Spread Sheet by Month") : GoTo EndExit
                End If
            ElseIf Me.pnlTypeOfRpt.Text = "Realization" And Me.txtSecondarySort.Text = "Spread Sheet by Year" Then
                If Format(EndYear, "yyyy") <> Format(StartYear.AddYears(3), "yyyy") Then
                    MsgBox("End Year must be 4 years after Start Year for the 4 Year Spread Sheet by Year Report! ***** Fix Dates", MsgBoxStyle.Critical + MsgBoxStyle.RetryCancel, "Spread Sheet by Year Report") : GoTo EndExit
                End If
            End If
            Me.tgr.UpdateData()
            Call PrtRealizationSpreadSheet(2) '06-23-15 
            '10-20-13 Forecast
        ElseIf Me.pnlTypeOfRpt.Text = "Quote Summary" And txtPrimarySortSeq.Text = "Forecasting" Then '10-20-13
            Me.tgln.UpdateData()
            Call PrintReportQuoteLines()
        ElseIf Me.pnlTypeOfRpt.Text = "Realization" Then
            Me.tgr.UpdateData()
            Call PrintRealizationReportQutTO()
        ElseIf Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then
            Me.tgln.UpdateData()
            Call PrintReportQuoteLines() '09-01-09
            SelectionText = "" '11-20-14 JTC Blank out after Product Sales History - Line Items" Status = " & rcode'11-20-14 JTC add SelectionText to Product Lines Report
        ElseIf Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then '09-19-12
            Me.tgln.UpdateData()
            Call PrintReportQuoteLines()
        Else ' "Terr Spec Credit Report" and others
            Me.tgQh.UpdateData()
            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-11-12
                Call PrintReportQuotesFollowBySesco()
            Else
                Call PrintReportQuotes() 'All Others IE Quote Summary
            End If
        End If
EndExit:    '10-13-14 JTC 
            Me.CmdRunReport.Enabled = True '10-13-14 JTC Fix Dbl Click on Reports
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default 'HourGlass 12-04-10
            Try
                ppv.Focus()  '05-14-15 JTC
            Catch myException As Exception
                'Do Nothing 
            End Try '05-14-15 JTC put on try ppv.Focus()  '02-09-09
            System.Windows.Forms.Application.DoEvents() '10-14-14 JTC

    End Sub
    Private Sub SetDefaultEstDelDate() '02-17-14 JTC Added Public DefChgEstDelDateCodes To "QUTDEFAU.DAT" forecasting
        'Public DefChangeEstDelDate As String = ""     'Y/N     
        'Public DefChgEstDelDateCodes As String = ""   'Status Codes Seperated by commas
        'DefaultstoScreen
        If DefChgEstDelDateCodes = "" Then  '02-14-14 Forecasting
            DefChgEstDelDateCodes = "SUBMIT,GOT"
        End If
        'QutDefauRead
        Dim DefFileName As String = "QUTDEFAU.DAT"
        Dim DefaultName As String = ""
        Dim GlobalDefault As Boolean = False
        Dim FileNumber As Short = FreeFile()
        Dim AA As String
        Dim Cs As String
        Dim DrivePath As String
        Try
            For I As Integer = 0 To 1
                '03-08-12 
                If I = 0 Then
                    DrivePath = UserDir
                Else
                    DrivePath = UserSysDir
                End If
                Call CheckMKDir(DrivePath)
                If My.Computer.FileSystem.FileExists(DrivePath & DefFileName) Then
100:                FileClose(FileNumber) : FileOpen(FileNumber, DrivePath & DefFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared) '"QUTDEFAU.DAT" Or Old ="QUOTEDEF.DAT"


GetRead:
200:                If EOF(FileNumber) Then Continue For Else AA = LineInput(FileNumber)
                    System.Windows.Forms.Application.DoEvents()

                    If AA.IndexOf("|SYS") <> -1 Then  '11-08-11 And Not myform Is Nothing Then
                        GlobalDefault = True
                        AA = AA.Substring(0, AA.IndexOf("|SYS"))
                    Else
                        GlobalDefault = False
                    End If
                    If DrivePath.EndsWith("SYS\") And GlobalDefault = False Then GoTo GetRead
                    F = AA.IndexOf("=")
                    If F = -1 Then
                        GoTo GetRead
                    Else
                        B = AA.Substring(0, F)
                        Cs = AA.Substring(F + 1)
                    End If

                    If InStr(B, "DefChangeEstDelDate") Then DefChangeEstDelDate = Trim(Cs) : DefaultName = "DefChangeEstDelDate" '10-24-13
                    If InStr(B, "DefChgEstDelDateCodes") Then DefChgEstDelDateCodes = Trim(Cs) : If VB.Left(DefChgEstDelDateCodes, 1) = "," Then DefChgEstDelDateCodes = VB.Mid(DefChgEstDelDateCodes, 2) '02-17-14 Public DefChgEstDelDateCodes As String = ""   'Status Codes Seperated by commas

                    GoTo GetRead
                End If
                FileClose(FileNumber)
            Next 'Get 
        Catch ex As Exception
            MessageBox.Show("Error in SetDefaultEstDelDate (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12  MsgBox(exc.Message, , "UnboundColumnFetch")
        End Try
GetExit:
        FileClose(FileNumber)
    End Sub
    Private Sub CmdShowColstoPrt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdShowColstoPrt.Click
        '"Show Hide Quote Hdr Printing Columns" = "VQrtShowHideRepPrtHdr.xml","VQrtShowHideDistPrtHdr.xml"   tgQh
        '"Show Hide Quote Line Items" = "VQrtLineItemsDistShowHide.xml"  tgln
        '"Show Hide Realization Columns" = "VQrtRealQTOShowHideDistPrint.xml"   tgr  or Rep
        '"Project Shortage","Product Sales History","Realization","Terr Spec Credit Report","Quote Summary","Planned Projects"
        Dim tmpName As String = "" ' howHidePrintQrt.xml" '09-02-09
        'If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.txtPrimarySortSeq.Text = "Quote Summary" Then '09-21-12  GoTo QutLineHistoryRpt
        If Me.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Or (Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (Me.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or Me.txtPrimarySortSeq.Text = "Salesman Follow-Up Report")) Then '09-21-09
            frmShowHideGrid.Text = "Show Hide Quote Line Items"
            '05-15-14 JTC Error if Me.pnlTypeOfRpt.Text <> "Terr Spec Credit Report" Then Don't do Call frmShowHideGrid.ShowHideGridCol IE: Fixed Columns
            If Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then Exit Sub '05-15-14 JTC 
            Me.Show() '02-06-09 
            'Dist VQrtLineItemsDistShowHide.xml, VQrtRealQTOShowHideDistPrint.xml,VQrtShowHideDistPrtHdr.xml
            'Rep  VQrtLineItemsRepShowHide.xml,  VQrtRealQTOShowHideRepPrint.xml, VQrtShowHideRepPrtHdr.xml
            If DIST Then tmpName = "VQrtLineItemsDistShowHide.xml" Else tmpName = "VQrtLineItemsRepShowHide.xml" '05-05-10
            ShowHideFileName = tmpName '09-02-09
            Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgln) '09-02-09
            Exit Sub 'Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgr) '09-02-09
        End If
        'Me.fraDisplaySortSeq.Visible = False ShowHideFileName = UserDir & "ShowHidePrintFoll.xml" '12-11-08
        If Me.pnlTypeOfRpt.Text.StartsWith("Realization") Then '01-22-10
            frmShowHideGrid.Text = "Show Hide Realization Columns" '12-22-08
            Me.Show() '02-06-09
            Dim ShowAllQuoteHeader As String = ""
            'Dim ShowAllQuoteHeader As String = "" '06-06-11 "Show All Quote Header Fields"
            If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then ShowAllQuoteHeader = "ShowAll" '06-06-11 = "Show All Quote Header Fields" Then '06-06-11 "Add Cust QuoteTo Breakdown to Report"
            If DIST Then tmpName = "VQrtRealQTOShowHideDistPrint" & ShowAllQuoteHeader & ".xml" Else tmpName = "VQrtRealQTOShowHideRepPrint" & ShowAllQuoteHeader & ".xml" '06-06-11 
            'If Me.pnlTypeOfRpt.Text = "Realization" Then tmpName = "VQrtRealQTOShowHidePrint.xml"
            ShowHideFileName = tmpName '09-02-09
            Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgr) '09-02-09) '12-05-08 ByVal ShowHide As String)
        Else '  Reqular Quote Report & Planned Project & Quote Summary
            frmShowHideGrid.Text = "Show Hide Quote Hdr Printing Columns" '"VQrtShowHidePrtHdr.xml"
            Me.Show() '02-06-09
            If DIST Then '01-19-10
                tmpName = "VQrtShowHideDistPrtHdr.xml" '
            Else
                tmpName = "VQrtShowHideRepPrtHdr.xml" '
            End If
            Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgQh) '05-27-10
        End If

        '09-03-10 Show hide to front Sub CmdShowColstoPrt
        If frmShowHideGrid.Visible = True Then frmShowHideGrid.BringToFront()
        '01-22-10 #######################################################################################################
        'RbnBtnExportExcel_Click()
        'frmShowHideGrid.Text = "Show Hide Realization Columns" '12-22-08
        'frmShowHideGrid.Show() '12-15-08
        'Dim tmpName As String = ""
        'If DIST Then '01-19-10
        '    tmpName = "VQrtShowHideDistPrtHdr.xml" '
        'Else
        '    tmpName = "VQrtShowHideRepPrtHdr.xml" '
        'End If
        ''Dim tmpName As String = "VQrtShowHidePrtHdr.xml"
        'If Me.pnlTypeOfRpt.Text = "Realization" Then tmpName = "VQrtRealQTOShowHidePrint.xml"
        'ShowHideFileName = tmpName '09-02-09
        'Call frmShowHideGrid.ShowHideGridCol("Show", tmpName, Me.tgr) '09-02-09
    End Sub


    Private Sub tg_FormatText(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.FormatTextEventArgs) Handles tgQh.FormatText, tgln.FormatText
        On Error Resume Next
        '02-10-10
        If e.Value.Trim <> "" Then e.Value = VB6.Format(e.Value, "MM/dd/yy") '09-10-09 

    End Sub


    Private Sub tabQrt_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles tabQrt.DrawItem
        'Turn Off toDebug
        Dim g As Graphics = e.Graphics
        Dim tp As TabPage = tabQrt.TabPages(e.Index)
        Dim br As Brush
        Dim sf As New StringFormat
        Dim r As New RectangleF(e.Bounds.X, e.Bounds.Y + 2, e.Bounds.Width, e.Bounds.Height - 2)

        sf.Alignment = StringAlignment.Center
        Dim strTitle As String = tp.Text

        'If the current index is the Selected Index, change the color
        If tabQrt.SelectedIndex = e.Index Then
            'this is the background color of the tab page
            'you could make this a stndard color for the selected page
            br = New SolidBrush(Color.Blue)
            g.DrawString(strTitle, tabQrt.Font, br, r, sf)
        Else
            'these are the standard colors for the unselected tab pages
            br = New SolidBrush(Color.WhiteSmoke)
            g.FillRectangle(br, e.Bounds)
            br = New SolidBrush(Color.Black)
            g.DrawString(strTitle, tabQrt.Font, br, r, sf)
        End If
    End Sub

    Private Sub tg_UnboundColumnFetch(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.UnboundColumnFetchEventArgs) Handles tgQh.UnboundColumnFetch, tgln.UnboundColumnFetch, tgr.UnboundColumnFetch
        Dim colName As String = e.Column.Caption
        Dim tgName As String = CType(sender, C1.Win.C1TrueDBGrid.C1TrueDBGrid).Name
        Try  ' tgln=LineItems tgQh = Quote, tgr = Quote Realization
            'New Public Function MarginOrCommCalc(ByVal Sell As Decimal, ByVal Cost As Decimal) As Decimal
            'tgln = Quote Lines,  ********************************************************************
            'NoIf tgName = "tgln" And colName = "SpecCredit" Then e.Value = tgln(e.Row, "SpecCredit") '02-06-10
            'tgName = ""
            'For I As Int16 = 0 To Me.tgr.Splits(0).DisplayColumns.Count - 1
            '    tgName = tgName & Me.tgr.Splits(0).DisplayColumns(I).Name & ","
            'Next
            '"ProjectCustID,ProjectID,NCode,Got,Typec,QuoteCode,ProjectName,FirmName,ContactName,EntryDate,SLSCode,Status,Cost,Sell,Comm-$,Comm-%,LPCost,LPSell,LPComm,LPMarg,Overage,ChgDate,OrdDate,NotGot,Comments,SPANumber,SpecCross,LotUnit,LampsIncl,Terms,FOB,QuoteID,BranchCode,MarketSegment,MFGQuoteNumber,BidDate,SLSQ,RetrCode,SelectCode,LeadTime,"
            Dim APrice As String = ""
            Dim UnitMeas As Decimal = 1
            'If tgName = "tgln" And colName = "FirmName" Then
            'e.Value = tgln(e.Row, "FirmName").ToString
            'End If

            If tgName = "tgln" Then  'tgln=LineItems
                APrice = tgln(e.Row, "Sell").ToString '02-15-10 
                'If APrice.EndsWith("M") Then Sto
                Dim UnitOfM As String = tgln(e.Row, "UM").ToString
                If UnitOfM = "C" Then UnitMeas = 100
                If UnitOfM = "M" Then UnitMeas = 1000
            End If
            'Caused Cascading event ?? Dim UnitMeaStr As String = UnitMeaSet(APrice, UnitMeas, tgln.Columns("UM").Text) '' C = Hundreds M = Thousands FT =Feet '01-28-04
            'Debug.Print(tgln(e.Row, "Comm-$").ToString)
            If tgName = "tgln" And colName = "Ext Sell" Then
                APrice = tgln(e.Row, "Sell").ToString '02-15-10 
                'If UnitMeas <> 1 Then UnitMeaStr = UnitMeaSet(APrice, UnitMeas, tgln.Columns("UM").Text) '' C = Hundreds M = Thousands FT =Feet '01-28-04
                '02-15-10 e.Value = Format(Val(tgln(e.Row, "Sell")) * Val(tgln(e.Row, "Qty")), "###.00") '12-17-09
                e.Value = Format(Val(APrice) * Val(tgln(e.Row, "Qty")), DecFormat) / UnitMeas '-1-06-13
            End If
            If DIST And (colName = "Margin-%" Or colName = "Margin" Or colName = "Margin-$") Then 'Decimal           Decimal Was Doubl '03-19-14 %
                If tgName = "tgln" Then
                    e.Value = Format(MarginOrCommCalc(Val(tgln(e.Row, "Sell")), Val(tgln(e.Row, "Cost")), False), "###.00")
                End If
                If tgName = "tgQh" Then
                    e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), Val(tgQh(e.Row, "Cost"))), "###.00")
                End If

                If tgName = "tgr" Then '03-19-14
                    e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "Sell")), Val(tgr(e.Row, "Cost"))), "###.00")
                End If

            ElseIf tgName = "tgln" And colName = "Comm-%" Then 'Rep
                Dim CostRep As Decimal = Val(tgln(e.Row, "Sell")) - Val(tgln(e.Row, "Comm")) '08-06-13
                e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), CostRep), "###.00")
            End If   'Rep Below
            If DIST Then '05-08-10 And tgName = "tgln" And colName = "Ext Marg" Then
                If colName = "Ext Marg" Then e.Value = tgln(e.Row, "Margin") '09-07-09 
            Else

            End If
            '06-25-13If DIST And tgName = "tgln" And colName = "Ext Cost" Then
            '    APrice = tgln(e.Row, "Cost").ToString '02-15-10 
            '    e.Value = Format(Val(APrice) * Val(tgln(e.Row, "Qty")), DecFormat) / UnitMeas '02-15-10
            'End If
            If tgName = "tgln" And colName = "Ext Cost" Or colName = "Ext Book" Then '03-19-14
                If DIST Then '06-25-13 
                    APrice = tgln(e.Row, "Cost").ToString '02-15-10 
                    e.Value = Format(Val(APrice) * Val(tgln(e.Row, "Qty")), DecFormat) / UnitMeas '02-15-10
                Else 'Rep Book '06-25-13 JTC Quote Lines Rep Book > FixCost
                    APrice = tgln(e.Row, "Book").ToString '
                    e.Value = Format(Val(APrice) * Val(tgln(e.Row, "Qty")), DecFormat) / UnitMeas '02-15-10
                End If
            End If
            '05-05-10 If Not DIST And tgName = "tgln" And colName = "Ext Comm" Then
            If tgName = "tgln" And colName = "Ext Comm" Or colName = "Ext Comm-$" Then '03-19-14
                APrice = tgln(e.Row, "Comm").ToString '05-08-10 -$
                e.Value = Format(Val(APrice) * Val(tgln(e.Row, "Qty")), DecFormat) / UnitMeas '02-15-10
            End If

            'Below is Quote Grid "tgQh" Header *****************************************************************
            If DIST And tgName = "tgQh" And colName = "Margin" Then                                  'DecimalDoubl                 
                e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), Val(tgQh(e.Row, "Cost"))), "###.00")
            ElseIf tgName = "tgQh" And colName = "Comm-%" Then 'Rep'12-05-09 
                Dim CostRep As Decimal = Val(tgQh(e.Row, "Sell")) - Val(tgQh(e.Row, "Comm-$"))
                e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), CostRep), "###.00")
            End If

            'Below is QuoteTo Realization tgr Grid Fixtures 
            If DIST And tgName = "tgr" And colName = "Margin" Then  'Decimal                 Doubl
                e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "Sell")), Val(tgr(e.Row, "Cost"))), "###.00")
            ElseIf tgName = "tgr" And (colName = "Comm-%" Or colName = "Margin") Then 'Rep '11-28-12 Added or colName = "Margin"
                '12-04-09 e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "Sell")), Val(tgr(e.Row, "Comm-%"))), "###.00")
                Dim CostRep As Decimal = Val(tgr(e.Row, "Sell")) - Val(tgr(e.Row, "Comm-$")) '12-07-09
                e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "Sell")), CostRep), "###.00")
            End If
            'Lamps Section  LAMPS ***************************************************************
            'tgln Below is tgln=LineItems LAMPS LampMargin 'If tgName = "tgln" Then Sto '09-04-09 ***************************
            If DIST And tgName = "tgln" And colName = "LPMarg" Then     'Decimal                 Doubl
                e.Value = Format(MarginOrCommCalc(Val(tgln(e.Row, "LPSell")), Val(tgln(e.Row, "LPCost"))), "###.00")
            ElseIf tgName = "tgln" And colName = "Comm-%" Then 'Rep
                '12-04-09 e.Value = Format(MarginOrCommCalc(Val(tgln(e.Row, "LPSell")), Val(tgln(e.Row, "LPComm"))), "###.00")
                Dim CostRep As Decimal = Val(tgln(e.Row, "Sell")) - Val(tgln(e.Row, "Comm-$")) '12-07-09 
                e.Value = Format(MarginOrCommCalc(Val(tgln(e.Row, "Sell")), CostRep), "###.00")
            End If
            If DIST = False And tgName = "tgln" And colName = "LPComm" Then 'Rep '09-06-10 
                '12-04-09 e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "LPSell")), Val(tgr(e.Row, "LPComm"))), "###.00")
                'Dim CostRep As Decimal = Val(tgr(e.Row, "LPSell")) - Val(tgr(e.Row, "LPCost"))
                e.Value = Format(Val(tgr(e.Row, "LPSell")) - Val(tgr(e.Row, "LPCost")), "###.00")
            End If
            'tgQh Below is Quote Grid LAMPS
            If DIST And tgName = "tgQh" And colName = "LPMarg" Then    'Decimal                 Doubl
                e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "LPSell")), Val(tgQh(e.Row, "LPCost"))), "###.00")
            ElseIf tgName = "tgQh" And colName = "LPComm" Then 'Rep
                '12-04-09 e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "LPSell")), Val(tgQh(e.Row, "LPComm"))), "###.00")
                'e.Value = Format(Decimal.Round(CDec(Val(tgQh(e.Row, "LPSell"))) - CDec(Val(tgQh(e.Row, "LPCost"))), 2), "###.00")
                Dim CommRep As Decimal = Val(tgQh(e.Row, "LPSell")) - Val(tgQh(e.Row, "LPCost"))
                '12-09-09 LPComm Dollarse.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), CostRep), "###.00")
                e.Value = Format(CommRep, "######.00") 'Dollars
            End If

            'tgr Below is QuoteTo Realization Grid Quote To LAMPS
            If DIST And tgName = "tgr" And colName = "LPMarg" Then   'Decimal                 Doubl
                e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "LPSell")), Val(tgr(e.Row, "LPCost"))), "###.00")
            ElseIf tgName = "tgr" And colName = "LPMarg" Then 'Rep '09-05-10 
                '12-04-09 e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "LPSell")), Val(tgr(e.Row, "LPComm"))), "###.00")
                '09-05-10 Dim CostRep As Decimal = Val(tgr(e.Row, "LPSell")) - Val(tgr(e.Row, "LPCost"))
                e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "LPSell")), Val(tgr(e.Row, "LPCost"))), "###.00") '09-05-10
                'tgln(e.Row, "LPComm") = Format(CostRep, "###.00")
            End If
            'If DIST And tgName = "tgr" And colName = "LPComm" Then   'Decimal                 Doubl
            '    e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "LPSell")), Val(tgr(e.Row, "LPCost"))), "###.00")
            If DIST = False And tgName = "tgr" And (colName = "LPComm" Or colName = "LPCost") Then 'Rep '09-05-10 
                '12-04-09 e.Value = Format(MarginOrCommCalc(Val(tgr(e.Row, "LPSell")), Val(tgr(e.Row, "LPComm"))), "###.00")
                'Dim CostRep As Decimal = Val(tgr(e.Row, "LPSell")) - Val(tgr(e.Row, "LPCost"))
                e.Value = Format(Val(tgr(e.Row, "LPSell")) - Val(tgr(e.Row, "LPCost")), "###.00")
            End If

        Catch ex As Exception
            MessageBox.Show("Error in UnboundColumnFetch (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12  MsgBox(exc.Message, , "UnboundColumnFetch")
            ' IfDebugOn ThenStop
        End Try

    End Sub



    Private Sub cmdReportLineItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReportLineItems.Click
        On Error Resume Next
        ExcelQuoteFU = False '04-28-14 JTC
        SESCO = False '04-28-14 JTC
        Me.chkBrandReport.Visible = False '05-16-13 
        '08-28-12 JTC Shut off Realization choices when you click on summary report
        For I = 0 To 10
            cboSortRealization.SetItemCheckState(I, CheckState.Unchecked) '01-21-12
        Next
        'cboSortRealization.SetItemCheckState(11, CheckState.Unchecked) '02-26-12 SESCO
        RealALL = False '10-13-14 JTC Fix Check box
        RealCustomer = False
        RealManufacturer = False
        RealQuoteTOOther = False '01-31-12
        RealSLSCustomer = False
        RealArchitect = False
        RealEngineer = False
        RealLtgDesigner = False
        RealSpecifier = False
        RealContractor = False
        RealOther = False ' 08-28-12 
        RealCustomerOnly = False '03-12-14

        Me.cboSortRealization.Visible = False '01-18-12
        Me.Text = "Quote Product Sales History - Line Items" & "  " & AGnam & "  UserID =" & UserID '09-04-10
        Me.pnlTypeOfRpt.Text = "Product Sales History - Line Items"
        Me.pnlTypeOfRpt.Visible = True
        Me.fraReportCmdSelection.Visible = False
        Me.fraSortPrimarySeq.Visible = True
        Me.pnlPrimarySortSeq.Visible = True
        Me.txtPrimarySortSeq.Text = ""
        Me.txtPrimarySortSeq.Visible = True
        Me.pnlQutRealCode.Text = "Code"
        Me.txtQutRealCode.Text = "ALL"
        Me.pnlQutRealCode.Visible = True
        Me.txtQutRealCode.Visible = True
        Me.chkSlsFromHeader.Visible = True '02-01-99 WNA
        Me.chkSlsFromHeader.Enabled = True '02-01-99 WNA
        Me.pnlQuoteToSls.Visible = True '07-16-02 WNA
        Me.txtQuoteToSls.Visible = True '07-16-02 WNA
        'Me.fraLines.Enabled = False
        If MFG = 0 And DAYB = 0 Then '02-21-03 WNA
            Me.chkIncludeCommDolPer.Visible = True
        End If
        Call FillPrimarySortCombo()
        Me.cboSortPrimarySeq.Focus()
    End Sub

    Private Sub mnuAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Private Sub MnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click
        frmAbout.Show() '09-17-09 
        'AboutBox1.Show()
        'End Sub
    End Sub

    Private Sub cmdReportPlannedProjects_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReportOtherTypes.Click
        On Error Resume Next
        ExcelQuoteFU = False '04-28-14 JTC
        SESCO = False '04-28-14 JTC
        Me.chkBrandReport.Visible = False '05-16-13 

        Me.cboSortRealization.Visible = False '01-18-12
        '08-28-12 JTC Shut off Realization choices when you click on summary report and other Non Realization Reports
        For I = 0 To 10
            cboSortRealization.SetItemCheckState(I, CheckState.Unchecked) '01-21-12
        Next
        'cboSortRealization.SetItemCheckState(11, CheckState.Unchecked) '02-26-12 SESCO
        RealALL = False '10-13-14 JTC Fix Check box
        RealCustomer = False
        RealManufacturer = False
        RealQuoteTOOther = False '01-31-12
        RealSLSCustomer = False
        RealArchitect = False
        RealEngineer = False
        RealLtgDesigner = False
        RealSpecifier = False
        RealContractor = False
        RealOther = False ' 08-28-12 
        RealCustomerOnly = False '03-12-14

        OrderBy = "" '11-04-10
        Me.pnlTypeOfRpt.Text = "Other Quote Types"
        Me.Text = "Other Quote Types" & "  " & AGnam & "  UserID =" & UserID '09-04-10'06-18-10 "Planned Projects Reports" '06-14-10 

        Me.cboLinesInclude.Visible = False '12-01-09
        Me.pnlTypeOfRpt.Visible = True
        Me.fraReportCmdSelection.Visible = False
        Me.fraSortPrimarySeq.Visible = True
        Me.pnlPrimarySortSeq.Visible = True
        Me.txtPrimarySortSeq.Text = ""
        Me.txtPrimarySortSeq.Visible = True
        Me.pnlQutRealCode.Visible = False
        Me.txtQutRealCode.Visible = False
        Me.chkSlsFromHeader.Visible = False
        Me.chkSlsFromHeader.Enabled = False
        Me.pnlQuoteToSls.Visible = True '01-27-09
        Me.txtQuoteToSls.Visible = True
        'If MFG = 0 And DAYB = 0 And DIST = 0 Then '12-09-02 WNA
        '    If Me.pnlTypeOfRpt.Text <> "Project Shortage" Then Me.chkIncludeCommDolPer.Visible = True
        'End If
        Call FillPrimarySortCombo()
        Me.cboSortPrimarySeq.Focus()
        'Me.cboTypeofJob.Text = "P" '06-18-10

    End Sub

    Private Sub cmdReportTerrSpecCredit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReportTerrSpecCredit.Click
        OrderBy = "" '11-04-10
        ExcelQuoteFU = False '04-28-14 JTC
        SESCO = False '04-28-14 JTC
        Me.cboSortRealization.Visible = False '01-18-12
        '08-28-12 JTC Shut off Realization choices when you click on summary report
        For I = 0 To 10
            cboSortRealization.SetItemCheckState(I, CheckState.Unchecked) '01-21-12
        Next
        'cboSortRealization.SetItemCheckState(11, CheckState.Unchecked) '02-26-12 SESCO
        RealALL = False '10-13-14 JTC Fix Check box
        RealCustomer = False
        RealManufacturer = False
        RealQuoteTOOther = False '01-31-12
        RealSLSCustomer = False
        RealArchitect = False
        RealEngineer = False
        RealLtgDesigner = False
        RealSpecifier = False
        RealContractor = False
        RealOther = False ' 08-28-12 
        RealCustomerOnly = False '03-12-14

        'If Me.pnlTypeOfRpt.Text = "" Then '10-23-02 WNA
        Me.pnlTypeOfRpt.Text = "Terr Spec Credit Report" '11-24-09
        'Me.fraLines.Enabled = True
        'Else
        'Me.fraLines.Enabled = False '05-17-02 WNA
        'End If
        Me.chkPrtPlanLines.Visible = True '06-20-10 
        Me.chkPrtPlanLines.Text = "Print Spec Credit Lines" '"Print Planned Project Lines"
        Me.cboLinesInclude.Visible = True '12-01-09'Include All Lines on Job
        Me.cboLinesInclude.Text = "Include All Lines on Job" '12-01-09 
        'Include All Lines on Job, 'Include Only Paid Items on the Job, 'Include Only UnPaid Items on the Job
        Me.lblJobName.Text = "Select One or more MFG's" '12-01-09
        Me.txtJobNameSS.Text = "ALL" '12-01-09
        Me.pnlTypeOfRpt.Visible = True
        Me.fraReportCmdSelection.Visible = False
        Me.fraSortPrimarySeq.Visible = True
        Me.pnlPrimarySortSeq.Visible = True
        Me.txtPrimarySortSeq.Text = ""
        Me.txtPrimarySortSeq.Visible = True
        Me.pnlQutRealCode.Visible = False
        Me.txtQutRealCode.Visible = False
        Me.chkSlsFromHeader.Visible = False
        Me.chkSlsFromHeader.Enabled = False
        Me.pnlQuoteToSls.Visible = True '01-27-09
        Me.txtQuoteToSls.Visible = True
        Call FillPrimarySortCombo()
        Me.cboSortPrimarySeq.Focus()
        Me.cboTypeofJob.Text = "S" '06-18-10
        '    If DefTypeOfJob = "Quotes" Then JT = "Q"
        '    If DefTypeOfJob = "Planned Projects" Then JT = "P"
        '    If DefTypeOfJob = "Spec Credit" Then JT = "S"
        '    If DefTypeOfJob = "Submittals" Then JT = "T"
        '    If DefTypeOfJob = "Other" Then JT = "O"

    End Sub



    Private Sub FontSizeComboBox_ChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles FontSizeComboBox.ChangeCommitted
        If FontSizeComboBox.Text = "" Then
            FontSizeComboBox.Text = "10"
            fs = 10
        End If

        If IsNumeric(FontSizeComboBox.Text) Then
            fs = FontSizeComboBox.Text
        End If

    End Sub



    Private Sub FontSizeComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontSizeComboBox.SelectedIndexChanged

    End Sub

    Private Sub RibbonFontComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RibbonFontComboBox2.SelectedIndexChanged
        If RibbonFontComboBox2.Text = "" Then RibbonFontComboBox2.Text = "Arial"
    End Sub

    Private Sub txtStartQuoteAmt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStartQuoteAmt.Validated, txtEndQuoteAmt.Validated
        '05-26-11 Added         CType(sender, TextBox).Text = CType(eventSender, TextBox).Text.ToUpper
        Dim A As String = sender.text ' txtStartQuoteAmt.Text
        A = Replace(A, ",", "")
        A = Replace(A, "$", "")
        If IsNumeric(A) Then  Else MsgBox("Amount not Numeric. Please Correct.")
        sender.text = A ' txtStartQuoteAmt.Text = A
    End Sub


    Public Sub cboSortRealization_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSortRealization.SelectedIndexChanged
        Static Dim InProcess As Boolean '11-28-12
        Try '01-18-12
            If InProcess = True Then GoTo SuBExit '11-28-12 : InProcess = False
            InProcess = True
            '03-12-14
            RealALL = False '10-13-14 JTC Fix Check box
            RealCustomer = False
            RealManufacturer = False
            RealQuoteToAmtON = False
            RealQuoteTOOther = False
            RealSLSCustomer = False
            RealArchitect = False
            RealEngineer = False
            RealLtgDesigner = False
            RealSpecifier = False
            RealContractor = False
            RealOther = False
            RealWithOneMfgCust = False
            ' RealALL = False
            '03-12-14
            '12-31-14 not Used  Const quote As String = """"
            ' First show the index and check state of all selected items. 
            Dim cnt As Integer = 0
            '07-15-14 JTC If  RealArchitect = False and RealEngineer =  False and RealLtgDesigner =  False and RealSpecifier =  False and RealContractor =  False and RealOther =  False then RealALL = True' All Specifiers
            '07-15-14 JTC If RealManufacturer = False and RealCustomer = True and  RealQuoteTOOther = False and RealSLSCustomer = False then "Select Cust Code"
            '07-15-14 JTC If RealManufacturer = True and RealCustomer = False and  RealQuoteTOOther = False and RealSLSCustomer = False then "Select MFG Code"
            'QuoteTO: Customer 
            If cboSortRealization.GetItemCheckState(0) = CheckState.Checked Then
                RealCustomer = True
                Me.cboSortPrimarySeq.Text = "Customer"
                SortSeq = "projectcust.Ncode"
            Else
                RealCustomer = False '07-15-14 RealALL = False
                'cboSortRealization.SelectedItems.Clear() '07-15-14 
            End If

            'QuoteTO: Manufacturer
            If cboSortRealization.GetItemCheckState(1) = CheckState.Checked Then
                RealManufacturer = True
                Me.cboSortPrimarySeq.Text = "Manufacturer"
                SortSeq = "projectcust.Ncode"
            Else
                RealManufacturer = False '07-15-14 RealALL = False
                'cboSortRealization.SelectedItems(1).Clear() '07-15-14 
            End If

            'QuoteTO: Other 
            If cboSortRealization.GetItemCheckState(2) = CheckState.Checked Then
                RealQuoteTOOther = True
                cboSortPrimarySeq.Text = "Other"
                SortSeq = "projectcust.Ncode"
            Else
                RealQuoteTOOther = False '07-15-14 RealALL = False
            End If

            'QuoteTO: Salesman/Customer
            If cboSortRealization.GetItemCheckState(3) = CheckState.Checked Then
                For I = 0 To 10
                    If I = 3 Then Continue For
                    If cboSortRealization.GetItemCheckState(I) = CheckState.Checked Then
                        MsgBox("No other items can be selected when the QuoteTO: Salesman/Customer option is selected!")
                    End If
                    cboSortRealization.SetItemCheckState(I, CheckState.Unchecked)
                Next
                'Me.cboSortRealization.GetItemCheckState(3) = CheckState.Checked '07-16-14 JTC Keep it Checked
                RealSLSCustomer = True
                Me.cboSortPrimarySeq.Text = "Salesman/Customer"
                SortSeq = "projectcust.SLSCode, projectcust.NCode"
            Else
                RealSLSCustomer = False
            End If

            'Architect
            If cboSortRealization.GetItemCheckState(4) = CheckState.Checked Then
                RealArchitect = True
                Me.cboSortPrimarySeq.Text = "Architect"
                SortSeq = "projectcust.Ncode"
            Else
                RealArchitect = False : RealALL = False
            End If

            'Engineer
            If cboSortRealization.GetItemCheckState(5) = CheckState.Checked Then
                RealEngineer = True
                Me.cboSortPrimarySeq.Text = "Engineer"
                SortSeq = "projectcust.Ncode"
            Else
                RealEngineer = False : RealALL = False
            End If

            'Ltg Designer
            If cboSortRealization.GetItemCheckState(6) = CheckState.Checked Then
                RealLtgDesigner = True
                Me.cboSortPrimarySeq.Text = "Ltg Designer"
                SortSeq = "projectcust.Ncode"
            Else
                RealLtgDesigner = False : RealALL = False
            End If

            'Specifier
            If cboSortRealization.GetItemCheckState(7) = CheckState.Checked Then
                RealSpecifier = True
                Me.cboSortPrimarySeq.Text = "Specifier"
                SortSeq = "projectcust.Ncode"
            Else
                RealSpecifier = False : RealALL = False
            End If

            'Contractor
            If cboSortRealization.GetItemCheckState(8) = CheckState.Checked Then
                RealContractor = True
                Me.cboSortPrimarySeq.Text = "Contractor"
                SortSeq = "projectcust.Ncode"
            Else
                RealContractor = False : RealALL = False
            End If

            'Other
            If cboSortRealization.GetItemCheckState(9) = CheckState.Checked Then
                RealOther = True
                cboSortPrimarySeq.Text = "Other"
                SortSeq = "projectcust.NCode"
            Else
                RealOther = False : RealALL = False
            End If
            'ALL Specifiers
            If cboSortRealization.GetItemCheckState(10) = CheckState.Checked Then
                RealALL = True
                Me.cboSortPrimarySeq.Text = "Other"
                For I = 4 To 9
                    cboSortRealization.SetItemCheckState(I, CheckState.Checked)
                Next
                cboSortRealization.SetItemCheckState(3, CheckState.Unchecked) '01-20-12 Salesman/Customer
                RealArchitect = True
                RealEngineer = True
                RealLtgDesigner = True
                RealSpecifier = True
                RealContractor = True
                RealOther = True
                SortSeq = "projectcust.Ncode"
            Else
                If cboSortRealization.SelectedItem.ToString = "ALL Specifiers" Then
                    RealALL = False
                    For I = 4 To 9
                        cboSortRealization.SetItemCheckState(I, CheckState.Unchecked)
                    Next
                    RealArchitect = False
                    RealEngineer = False
                    RealLtgDesigner = False
                    RealSpecifier = False
                    RealContractor = False
                    RealOther = False
                    SortSeq = "projectcust.Ncode"
                End If
            End If
            'Only Quotes with One MFG/Cust
            If cboSortRealization.GetItemCheckState(11) = CheckState.Checked Then
                RealWithOneMfgCust = True
                Dim StrXX As String = RealWithOneMfgCustCode : If StrXX.Trim = "" Then StrXX = "XXXXXX" ''07-17-14"
                RealWithOneMfgCustCode = UCase(InputBox("Allows you to find all quotes to One Mfg or Cust." & vbCrLf & "Only Select Quotes that have a MFG Code or Customer Code associated with it. (KEEN/GES/AT)", "Mfg/Cust Code", StrXX)) '07-16-14 
                If RealWithOneMfgCustCode = "" Or RealWithOneMfgCustCode = "XXXXXX" Then RealWithOneMfgCust = False : RealWithOneMfgCustCode = "" : cboSortRealization.SetItemCheckState(11, CheckState.Unchecked)
                '04-16-15 JTC Need All Customers with this Factory
                'If RealCustomer = True And RealManufacturer = False And RealQuoteTOOther = False And RealSLSCustomer = False And RealArchitect = False And RealEngineer = False And RealLtgDesigner = False And RealSpecifier = False And RealContractor = False And RealOther = False ThenStop ' 04-16-15
            Else
                RealWithOneMfgCust = False : RealWithOneMfgCustCode = ""
            End If
            If cboSortRealization.GetItemCheckState(12) = CheckState.Unchecked Then '04-29-15 JTC Turn Off
                SESCO = False
                ExcelQuoteFU = False '04-29-15 JTC
            End If
            'SESCO Job List Report & "Excel Quote FollowUp"
            If cboSortRealization.GetItemCheckState(12) = CheckState.Checked Then
                '04-22-15 frmQuoteRpt.cboSortPrimarySeq.Items.Add("Excel Quote FollowUp") '04-22-15 JTC add cboSortRealization.Items(12).ToString = "Excel Quote FollowUp"  when not "SESCO Job List Report" to "Excel Quote FollowUp" Realization
                If cboSortRealization.Items(12).ToString = "SESCO Job List Report" Then '04-28-15  If Me.cboSortRealization.Items(12) = "SESCO Job List Report" Then
                    '04-22-15 JTC Use cboSortRealization not cboSortPrimarySeq for "SESCO Job List Report" or Excel Quote FollowUp Me.cboSortRealization.Items(12) = "SESCO Job List Report" or "Excel Quote FollowUp"
                    SESCO = True
                    ExcelQuoteFU = False '04-28-15 JTC
                    txtPrimarySortSeq.Text = "SESCO Job List Report"
                    Me.pnlTypeOfRpt.Text = "Realization"
                    Me.cboSortPrimarySeq.Text = "SESCO Job List Report" '01-18-12
                    For I = 0 To 11
                        If cboSortRealization.GetItemCheckState(I) = CheckState.Checked Then
                            RealWithOneMfgCust = False : RealWithOneMfgCustCode = "" '04-29-14
                            MsgBox("No other items can be selected when the SESCO Job List Report option is selected!")
                        End If
                        cboSortRealization.SetItemCheckState(I, CheckState.Unchecked)
                    Next
                ElseIf cboSortRealization.Items(12).ToString = "Excel Quote FollowUp" Then
                    SESCO = False
                    ExcelQuoteFU = True '04-28-15 JTC
                    Me.cboSortPrimarySeq.Text = "Excel Quote FollowUp"
                    Me.pnlTypeOfRpt.Text = "Realization"
                    '04-27-15 JTC Added Below
                    Me.ChkCheckBidDates.CheckState = CheckState.Checked
                    '06-05-12 If Me.DTPicker1StartBid.Value = CDate("01/01/1900") Then '06-05-12 Dim NewDate As Date = Now '02-03-12 '06-05-12 NewDate = NewDate.AddDays(-7) 'AddYears : NewDate = NewDate.AddDays(-1)
                    Me.DTPicker1StartBid.Enabled = True
                    Me.DTPicker1EndBid.Enabled = True
                    For I = 0 To 11
                        If cboSortRealization.GetItemCheckState(I) = CheckState.Checked Then
                            RealWithOneMfgCust = False : RealWithOneMfgCustCode = "" '04-29-14
                            MsgBox("No other items can be selected when the Excel Quote FollowUp option is selected!")
                        End If
                        cboSortRealization.SetItemCheckState(I, CheckState.Unchecked)
                    Next
                End If
            End If
            If SortSeq = "" Then
                Me.txtPrimarySortSeq.Text = "Customer" '01-18-12
                SortSeq = "projectcust.NCode"
            End If
            InProcess = False '07-16-14 JTC Repeat this ?
        Catch ex As Exception
            '07-06-12 Enumerator is Bound so Changes Causes Enumerator Error
            '07-06-12 MessageBox.Show("Error in SelectedIndexChanged (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12 
        End Try
SuBExit: InProcess = False
        If RealCustomer = True And RealManufacturer = False And RealQuoteToAmtON = False And RealQuoteTOOther = False And RealSLSCustomer = False And RealArchitect = False And RealEngineer = False And RealLtgDesigner = False And RealSpecifier = False And RealContractor = False And RealOther = False And RealWithOneMfgCust = False And RealALL = False Then '03-12-14
            RealCustomerOnly = True
        Else
            RealCustomerOnly = False
        End If
    End Sub

    Private Sub ChkCheckBidDates_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkCheckBidDates.CheckedChanged
        '01-01-1900'02-03-12
        If ChkCheckBidDates.CheckState = CheckState.Checked Then
            '06-05-12 If Me.DTPicker1StartBid.Value = CDate("01/01/1900") Then
            '06-05-12 Dim NewDate As Date = Now '02-03-12
            '06-05-12 NewDate = NewDate.AddDays(-7) 'AddYears : NewDate = NewDate.AddDays(-1)
            '06-05-12 Me.DTPicker1StartBid.Value = NewDate '02-03-12 'A = Format(Now.YearaDDyear(1).Year + 1, "yyyy")
            'If Me.ChkCheckBidDates.Text = "Check Deliver Dates when Selecting Quotes" ThenStop
            chkBlankBidDates.Visible = True ' CheckState =
            Me.DTPicker1StartBid.Enabled = True '02-04-12
            Me.DTPicker1EndBid.Enabled = True '02-04-12

            'ChkIgnoreBidDates.Text = "Use Bid Dates when Selecting Quotes"
            'End If
        Else
            '06-05-12 Me.DTPicker1StartBid.Value = CDate("01/01/1900")

            chkBlankBidDates.CheckState = CheckState.Unchecked
            chkBlankBidDates.Visible = False ' CheckState =
            Me.DTPicker1StartBid.Enabled = False '02-04-12
            Me.DTPicker1EndBid.Enabled = False '02-04-12 
        End If
    End Sub

    Private Sub chkCustomerBreakdown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCustomerBreakdown.CheckedChanged
        '03-29-12stop
        If Me.chkCustomerBreakdown.CheckState = CheckState.Checked Then
            'Turn on '03-29-12 Need to Add a Check box "Only Show the Latest Quote To Each Customer
            Me.chkShowLatestCust.Visible = True '03-29-12
            Me.chkShowLatestCust.CheckState = CheckState.Checked '03-29-12
        Else
            Me.chkShowLatestCust.Visible = False
            Me.chkShowLatestCust.CheckState = CheckState.Unchecked '03-29-12
        End If
    End Sub


    Private Sub chkPrintGrayScale_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkPrintGrayScale.CheckedChanged
        If Me.chkPrintGrayScale.Checked = True Then Me.RibbonTab6.Text = "Print GrayScale" Else Me.RibbonTab6.Text = "Print Color" '01-18-13
    End Sub

    Private Sub cmdBackViewGrid_Click(sender As System.Object, e As System.EventArgs) Handles cmdBackViewGrid.Click
        '03-22-13 Go back to Select Criteria Tab From Grid cmdBackViewGrid
        System.Windows.Forms.Application.DoEvents()
        Me.tabQrt.SelectedIndex = 1
        Call tabQRT_TabActivate(1)
    End Sub

    Private Sub ChkBrandMfgRpt_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBrandReport.CheckedChanged ', ChkBrandMfgRpt22.CheckedChanged
        'Works COOP & for All N&A Brands not just PHIL '05-15-15 JTC
        '05-16-13  Public BrandReport As Boolean '05-16-13
        '05-16-13 JTC Add PHIL Brand Reports Get Brands from N&A Notes Record Call GetPHILBrands(A)
        'SELECT * FROM saw8sesco.namedefaults where NCode = '999999' and RecType = 'EDI' and Category = 'ORDERS' 
        'Comments = 'ZZ|MULTIMICRO SYST|ZZ|606449916|THMI=LAM=MORL=DAYB=MCPH=OMEG=CAPR=THMO=CHLO=GUTH'
        If chkBrandReport.CheckState = CheckState.Checked Then
            Dim PhilBrands As String = "" '05-16-13 
            Call GetPHILBrands(PhilBrands) '05-16-13
            If PhilBrands.Trim = "" Then
                BrandReport = False '05-16-13 
                MsgBox("Enter Major Brand to breakdown in Name & Address System" & vbCrLf & "999999 record, EDI type and ORDERS" & vbCrLf & "Enter (PHIL=CAPR=MORL=LOL...etc)" & vbCrLf & "or (COOP=MTLX=CORE=RSA...etc") ' "Brands Select") '05-16-13
            Else
                Me.txtQutRealCode.Text = PhilBrands
                BrandList = PhilBrands '10-20-13
                System.Windows.Forms.Application.DoEvents()
                BrandReport = True '05-16-13 
            End If
        Else
            Me.txtQutRealCode.Text = "ALL"
            System.Windows.Forms.Application.DoEvents()
            BrandReport = False '05-16-13

            Me.ChkCheckBidDates.CheckState = CheckState.Unchecked '02-24-14 
            Me.ChkCheckBidDates.Text = "Check Bid Dates when Selecting Quotes" '02-24-14
            Me.mnuBrandReport.Enabled = False '11-06-13 
            Me.mnuBrandMfgChg.Enabled = False '11-06-13
            Me.chkBlankBidDates.Visible = True '11-06-13
            Me.mnuBrandReport.Text = "Brand Reporting - Off"
            Me.chkBrandReport.Visible = False '
            Me.lblStartEntry.Enabled = True '11-06-13
            Me.DTPickerStartEntry.Enabled = True '11-06-13
            Me.DTPicker1EndEntry.Enabled = True '11-06-13
            Me.txtStatus.Text = "ALL"

        End If
    End Sub

    Private Sub chkShowCustomers_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowCustomers.CheckedChanged
        '03-24-14 Me.Label7.Text = "Customer Code" '08-02-13
        '02-25-14 Me.chkUseSpecifierCode.CheckState = CheckState.Unchecked '08-07-13
        If chkShowCustomers.Checked = True Then '03-24-14
            chkUseSpecifierCode.Checked = False
            Me.Label7.Text = "Customer Code"
        End If
    End Sub

    Private Sub chkUseSpecifierCode_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkUseSpecifierCode.CheckedChanged
        '03-24-14 Me.Label7.Text = "Specifier Code" '08-02-13
        '03-24-14 Me.chkShowCustomers.CheckState = CheckState.Unchecked '08-07-13
        If chkUseSpecifierCode.Checked = True Then '03-24-14 
            chkShowCustomers.Checked = False
            Me.Label7.Text = "Specifier Code"
        End If
    End Sub

    Private Sub _fdBranchCode_TextChanged(sender As System.Object, e As System.EventArgs) Handles _fdBranchCode.LostFocus
        '10-16-13
        If Me._fdBranchCode.Text = "" Then Me._fdBranchCode.Text = "ALL"
        Me._fdBranchCode.Text = Me._fdBranchCode.Text.ToUpper
        SecurityBrancheCodes = Me._fdBranchCode.Text
    End Sub

    Private Sub mnuBrandMfgChg_Click(sender As Object, e As System.EventArgs) Handles mnuBrandMfgChg.Click
        '03-12-13 Change Brand Name Me.mnuBrandMfgChg.Text = "Brand Mfg Code - " & BrandReportMfg '03-12-13 XXXX Brand Mfg Code - XXXX
        If BrandReportMfg.Trim = "" Then BrandReportMfg = "XXXX"
        BrandReportMfg = UCase(InputBox("Enter Major Brand to breakdown PHIL, COOP, LITH", "Enter Top Level MFG Code.", BrandReportMfg)) '03-12-13
        If BrandReportMfg.Trim = "" Then BrandReportMfg = "XXXX" '02-27-14 
        BrandReportMfg = VB.Left(BrandReportMfg, 4) '04-04-13
        Dim FileName2 As String
        '02-24-14 JTC Move "BRANDREPORT-" From UserPath to UserSysDir
        FileName2 = Dir$(UserSysDir & "BrandReport*.*") '"BrandReport-*.DAT")
        '02-06-14 If My.Computer.FileSystem.FileExists(UserPath & FileName2) = False Then '11-07-13 JTC Fix Duplicate mnuBrandMfgChg_Click BrandReportMfg = "XXXX"
        If My.Computer.FileSystem.FileExists(UserPath & "BrandReport-" & BrandReportMfg & ".DAT") = True Then '02-06-14 JTC Fix Changing MFG on "BrandReport-" & BrandReportMfg & ".DAT"
            My.Computer.FileSystem.CopyFile(UserPath & FileName2, UserSysDir & "BrandReport-" & BrandReportMfg & ".DAT")
            My.Computer.FileSystem.DeleteFile(UserPath & FileName2)
            BrandReportMfg = VB.Left(BrandReportMfg, 4) '02-27-14
        Else '02-27-14 
            If My.Computer.FileSystem.FileExists(UserSysDir & "BrandReport-" & BrandReportMfg & ".DAT") = False Then '02-27-14
                If My.Computer.FileSystem.FileExists(UserSysDir & FileName2) = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & FileName2)
                Dim FN As Integer = FreeFile() : FileClose(FN)
                FileOpen(FN, UserSysDir & "BrandReport-" & BrandReportMfg & ".DAT", OpenMode.Output)
                FileClose(FN)
                BrandReportMfg = VB.Left(BrandReportMfg, 4) '02-27-14
            End If
        End If

        '03-12-13 Change Brand Name 
        Me.mnuBrandMfgChg.Text = "Brand Mfg Code - " & BrandReportMfg '03-12-13 XXXX Brand Mfg Code - XXXX
        Me.chkBrandReport.Text = "Brand Reporting - " & BrandReportMfg '03-14-13 XXXX Brand Mfg Code - XXXX"
        '03-12-13 mnuBrandReport
        If mnuBrandReport.Text = "Brand Reporting - Off" Then
            mnuBrandReport.Text = "Brand Reporting - On"
        End If
        BrandReportMfg = BrandReportMfg.ToUpper '02-04-15 JTC Must be Upper case  BrandReportMfg = BrandReportMfg.ToUpper
    End Sub

    Private Sub mnuBrandReport_Click1(sender As Object, e As System.EventArgs) Handles mnuBrandReport.Click
        '11-06-13 mnuBrandReport
        If mnuBrandReport.Text = "Brand Reporting - Off" Then
            ''If MajSel = RptMaj.RptSpecCredit Or MajSel = RptMaj.RptRefCredit Then
            'MsgBox("Brand Reporting not allowed on this option." & vbCrLf & "Turning Brand Reporting Off.")
            'mnuBrandReport.Text = "Brand Reporting - Off"
            'Me.chkBrandReport.CheckState = CheckState.Unchecked '03-14-13
            ': Exit Sub '03-14-13
            ''End If
            mnuBrandReport.Text = "Brand Reporting - On"
            Me.chkBrandReport.CheckState = CheckState.Checked '03-14-13
            MsgBox("Brand Reporting - On") '03-13-13
        Else
            mnuBrandReport.Text = "Brand Reporting - Off"
            Me.chkBrandReport.CheckState = CheckState.Unchecked '03-14-13
            MsgBox("Brand Reporting - Off") '03-13-13
        End If
    End Sub

    Private Sub cmdSortThirdSeq_Click(sender As System.Object, e As System.EventArgs)
        On Error Resume Next
        Me.tabQrt.SelectedIndex = 1
        Call tabQRT_TabActivate(1)
    End Sub

    Private Sub cmdSortSecondarySeqCancel_Click(sender As System.Object, e As System.EventArgs)
        'gbThirdSort.Visible = False
    End Sub

    Private Sub cboSortSecondarySeq_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)
        txtSortSecondarySeq.Text = cboSortSecondarySeq.Text
    End Sub

    Private Sub cboSortPrimarySeq_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboSortPrimarySeq.SelectedIndexChanged
        txtPrimarySort.Text = cboSortPrimarySeq.Text
    End Sub

    Public Sub rbnDeleteFiles_Click(sender As System.Object, e As System.EventArgs) Handles rbnDeleteFiles.Click
        Try
            If DIST = False Then
                If File.Exists(UserSysDir & "VQrtHdrTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtHdrTGLayoutOriginal.xml")
                If File.Exists(UserSysDir & "VQrtLinesTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtLinesTGLayoutOriginal.xml")
                If File.Exists(UserSysDir & "VQrtQuoteToTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtQuoteToTGLayoutOriginal.xml")
                If File.Exists(UserSysDir & "VQrtSpecCreditOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtSpecCreditOriginal.xml")
                If File.Exists(UserDir & "VQrtHdrTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtHdrTGLayoutCurrent.xml")
                If File.Exists(UserDir & "VQrtLineItemsRepShowHide.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLineItemsRepShowHide.xml")
                If File.Exists(UserDir & "VQrtLinesTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLinesTGLayoutCurrent.xml")
                If File.Exists(UserDir & "VQrtQuoteToTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtQuoteToTGLayoutCurrent.xml")
                If File.Exists(UserDir & "VQrtRealQTOShowHideRepPrint.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtRealQTOShowHideRepPrint.xml")
                If File.Exists(UserDir & "VQrtShowHideRepPrtHdr.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtShowHideRepPrtHdr.xml")
                If File.Exists(UserDir & "VQrtSpecCreditCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtSpecCreditCurrent.xml")
                If File.Exists(UserDir & "VQrtQuoteToTGLayoutShowAllCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtQuoteToTGLayoutShowAllCurrent.xml") '07-22-14
                If File.Exists(UserDir & "VQrtRealQTOShowHideRepPrintShowAll.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtRealQTOShowHideRepPrintShowAll.xml") '07-22-14
                'VQrtQuoteToTGLayoutShowAllCurrent
                'VQrtRealQTOShowHideRepPrintShowAll.xml
            Else
                'DIST FILES
                If File.Exists(UserSysDir & "VQrtHdrDistTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtHdrDistTGLayoutOriginal.xml")
                If File.Exists(UserSysDir & "VQrtLinesDistTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtLinesDistTGLayoutOriginal.xml")
                If File.Exists(UserSysDir & "VQrtQuoteToDistTGLayoutOriginal.xml") = True Then My.Computer.FileSystem.DeleteFile(UserSysDir & "VQrtQuoteToDistTGLayoutOriginal.xml")
                If File.Exists(UserDir & "VQrtHdrDistTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtHdrDistTGLayoutCurrent.xml")
                If File.Exists(UserDir & "VQrtLineItemsDistShowHide.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLineItemsDistShowHide.xml")
                If File.Exists(UserDir & "VQrtLinesDistTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtLinesDistTGLayoutCurrent.xml")
                If File.Exists(UserDir & "VQrtQuoteToDistTGLayoutCurrent.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtQuoteToDistTGLayoutCurrent.xml")
                If File.Exists(UserDir & "VQrtRealQTOShowHideDistPrint.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtRealQTOShowHideDistPrint.xml")
                If File.Exists(UserDir & "VQrtShowHideDistPrtHdr.xml") = True Then My.Computer.FileSystem.DeleteFile(UserDir & "VQrtShowHideDistPrtHdr.xml")
            End If

            Resp = MessageBox.Show("All Done.  You will need to exit the Quote Reports and Come back in for the settings to take effect." & vbCrLf & "Do you want to close now?", "Restart the Reports", MessageBoxButtons.YesNoCancel)
            If Resp = vbYes Then Call Jump("VMENU.EXE") : Me.Close() '04-18-10

        Catch ex As Exception
            MessageBox.Show("Error in rbnDeleteFiles_Click" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12 MsgBox(exc.Message, , "Zoom")
        End Try

    End Sub

    Private Sub BrandListingLoadToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles mnuBrandListLoad.Click
        '02-21-14 JTC Added Menu Item Brand List Load and GetPHILBrands(PhilBrands)
        Dim PhilBrands As String = "" '05-16-13 
        If (BrandList.Trim <> "" Or BrandList = "ALL") And BrandReport = True Then '03-14-14 JTC Turn OFF
            BrandList = "" '            Dim TmptxtMfgLine2 As String = BrandList TmptxtMfgLine2 = Replace(TmptxtMfgLine2, "=", ",")  Me.txtMfgLine2.Text = TmptxtMfgLine2
            Me.mnuBrandListLoad.Text = "Brand Listing Load - Off" '03-14-14 
            Me.mnuBrandExclude.Text = "Exclude Brands - " & BrandReportMfg & " - Off" '03-14-14 '02-21-14
            Me.txtQutRealCode.Text = ""
            System.Windows.Forms.Application.DoEvents()
            Exit Sub
        End If
        Call GetPHILBrands(PhilBrands) '05-16-13
        If PhilBrands.Trim = "" Then
            BrandReport = False '05-16-13 
            MsgBox("Enter Major Brand to breakdown in Name & Address System" & vbCrLf & "999999 record, EDI type and ORDERS" & vbCrLf & "Enter (PHIL=CAPR=MORL=LOL...etc)" & vbCrLf & "or (COOP=MTLX=CORE=RSA...etc") ' "Brands Select") '05-16-13
        Else
            BrandList = PhilBrands
            'System.Windows.Forms.Application.DoEvents()
            BrandReport = True '05-16-13 
        End If
        If BrandList.Trim = "" Then BrandList = "ALL"
        Dim TmptxtMfgLine2 As String = BrandList
        TmptxtMfgLine2 = Replace(TmptxtMfgLine2, "=", ",")
        Me.txtQutRealCode.Text = TmptxtMfgLine2
        Me.mnuBrandListLoad.Text = "Brand Listing Load - ON" '03-14-14 
        Me.mnuBrandExclude.Text = "Exclude Brands - " & BrandReportMfg & " - Off" '03-14-14 '02-21-14
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub mnuBrandExclude_Click(sender As Object, e As System.EventArgs) Handles mnuBrandExclude.Click
        'If BrandList.Trim = "" Then Call mnuBrandListLoad_Click()
        Dim PhilBrands As String = "" '05-16-13 
        Call GetPHILBrands(PhilBrands) '05-16-13
        If PhilBrands.Trim = "" Then
            BrandReport = False '05-16-13
            MsgBox("Enter Major Brand to breakdown in Name & Address System" & vbCrLf & "999999 record, EDI type and ORDERS" & vbCrLf & "Enter (PHIL=CAPR=MORL=LOL...etc)" & vbCrLf & "or (COOP=MTLX=CORE=RSA...etc") ' "Brands Select") '05-16-13
        Else
            BrandList = PhilBrands
            'System.Windows.Forms.Application.DoEvents()
            BrandReport = True '05-16-13 
        End If
        If BrandList.Trim = "" Then BrandList = "ALL"
        If BrandList.Trim <> "ALL" Then
            Dim BL As String = BrandList '02-21-14 (-KEEN,-GLOB,-LITH) to eliminate Mfgs
            BL = Replace(BL, ",", ",-")
            If BL.StartsWith("-") Then  Else BL = "-" & BL 'add - minus in front
            Me.txtQutRealCode.Text = BL
            Me.mnuBrandExclude.Text = "Exclude Brands - " & BrandReportMfg & " - ON" '03-14-14 '02-21-14
            Me.mnuBrandListLoad.Text = "Brand Listing Load - ON" '03-14-14 
        Else
            Me.mnuBrandExclude.Text = "Exclude Brands - " & BrandReportMfg & " - Off" '03-14-14 '02-21-14
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub cboTypeCustomer_TextChanged(sender As Object, e As System.EventArgs) Handles cboTypeCustomer.TextChanged
        If cboTypeCustomer.Text = "T" Then '01-27-15 JTC "Add MFG Total Breakdown to Reports" '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero" 
            If Me.chkMfgBreakdown.Text = " Print Quote Hdr Amt when QuoteTo is Zero" Then ' "Add MFG Total Breakdown to Reports" '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero" 
                Me.chkMfgBreakdown.Visible = True
            End If
        Else
            If Me.chkMfgBreakdown.Text = " Print Quote Hdr Amt when QuoteTo is Zero" Then ' "Add MFG Total Breakdown to Reports" '01-27-15 JTC Print Quote Hdr Amt when QuoteTo is Zero" 
                Me.chkMfgBreakdown.Visible = False
            End If
        End If
    End Sub

    Private Sub rbnJoinGoToMeeting_Click(sender As System.Object, e As System.EventArgs) Handles rbnJoinGoToMeeting.Click
        On Error Resume Next '06-18-12
        System.Diagnostics.Process.Start("http://www.JoinGoToMeeting.com")
    End Sub

    Private Sub rbnHelpAbout_Click(sender As System.Object, e As System.EventArgs) Handles rbnHelpAbout.Click
        On Error Resume Next
        frmAbout.ShowDialog()
    End Sub

    Private Sub rbnHelpMaster_Click(sender As System.Object, e As System.EventArgs) Handles rbnHelpMaster.Click
        On Error Resume Next
        '05-08-15 JTC
        'Help (.chm)
        'Below doesn't work over a network 'Dim RetVal As Integer = Shell("hh.exe " & UserPath & "hlporder.chm", AppWinStyle.NormalFocus)
        Dim RetVal As Integer '  RetVal = ShellExecute(Handle.ToInt32, "OPEN", "http://www.multimicrosystems.com/helpsys/hlpOrder/Source/HTML/index.html", "", "", AppWinStyle.MaximizedFocus) '02-28-06
        Call HelpSub("HelpToolStripMenuItem") '09-10-10
    End Sub
End Class