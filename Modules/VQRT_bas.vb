Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Strings
Imports MySql
Imports MySql.Data.MySqlClient
'Imports Microsoft.VisualBasic.Strings
Imports VB = Microsoft.VisualBasic
'Imports Microsoft.VisualBasic.PowerPacks
Imports C1.C1Preview
Imports C1.C1Preview.DataBinding '09-20-08
Imports C1.Win.C1Preview '06-18-08 
'Imports C1.C1Schedule
'12-03-08 Imports C1.C1Report '09-09-08 
Imports System.Xml '09-09-08
Imports System.IO '09-09-08

Module VQRT1
	'---------------------------------------------------------
    '08-15-08 ErrorRoutine: This allows you to sto and try and fix vb errors after the error message **Get gid of Sto 'CatchStop b/4 releasing
	'MsgBox(*VB Error # =  * & Str(Err.Number) & *  ERROR at * & Str(Err.Erl()) & vbCrLf & ErrorToString(Err.Number) & vbCrLf & *In FillSentItems Routine*, 16, *FillSentItems
    'Sto 'CatchStop     Replace * with DblQuote 
	'Resume Next
	'------------------------------------------------
    'Catch ex As Exception  'Try Catch with sto so you can fix exception after the error message **Get gid of Sto 'CatchStop b/4 releasing
	'    MsgBox(ex.Message.ToString & vbCrLf & *SaveToDataBase* & vbCrLf) 
    '    Sto 'CatchStop  'Debug.WriteLine(ex.Message.ToString)
	'End Try            Replace * with DblQuote 
	'--------------------------------------------------
    Sub Main_Renamed(ByRef Cmd As String)
        Static FirstTime As Short
        Static CustSlsSelect As String
        Static Enable As Short
        Static ONECUST As String
        Static ONECODE As Short
        Static WM As String
        Static MFGCs As String
        Static MFGC As Short
        Static WN As Short
        Static Count As Short
        Static SaveStat As String
        Static savesls As String
        Static Hit As Short
        Static Sorted As Short
        Static Commpct As Decimal
        Static CommAmt As Decimal
        Static J As Short
        Static First As Short
        Static I As Short
        Static JOBSER As String
        Static SLTCommPct As Decimal
        Static SLTCommAmt As Decimal
        Static SDIS As String
        Static EK As Short
        Static BC As Short
        Static fdi As Short
        Static Temp As String
        Static DQK As String
        Static MR As String
        Static Seq As Short
        Static SEQHIGH As Short
        Static index As String
        Static CURQN As Short
        Static A As String
        Static ZLEN As Short
        Static EM As String
        Static Msg As String
        Static Resp As Short
        Static Taskid As Integer
        Static LOR As Short
        Static ZS As String
        Static ZR As String
        Static WV As Short
        Static M As Short
        Static Z As Short
        Static VB6 As Object
        ' VQRT   Quote Reports
        ' (c) COPYRIGHT 1984, 2012 MULTIMICRO SYSTEMS   ' YEAR END
        '***UNAUTHORIZED REPRODUCTION OR USE OF THIS SOFTWARE IS PROHIBITED
        'AND IS IN VIOLATION OF UNITED STATES COPYRIGHT LAWS ALL RIGHTS RESERVED.
        AbortPrtFlag = False
        On Error GoTo 500
        If Cmd = "INIT" Then GoTo INIT12
        If Cmd = "END" Then GoTo End255
        If Cmd = "ENDWIN" Then GoTo EndWin255
        If Cmd = "SHELL255" Then GoTo SHell255
        If Cmd = "VQUT.EXE" Then GoTo QUT255
        MsgBox("No Valid Command in Cmd$ of Sub Main", MsgBoxStyle.OkOnly, US)
        Exit Sub
INIT12:
        On Error GoTo 500
        US = "Quote Reports"
        Wspcs = "                                                                       "
        GoTo MSub_Exit

255:
EndWin255: FileClose() : Call frmQuoteRpt.FormSetting("Save") 'During FormClosing Event to Save Settings
        Call CloseSQL(myConnection) '06-10-09  myConnection.Close() '06-10-09
        '03-19-12 No No frmQuoteRpt.Close() '03-19-12 
        End ' This gets called if Windows Cancels our App With Ctrl Esc & End Task

End255:  'Call BtrvSub(DB_Reset)
SHell255: FileClose() : Call frmQuoteRpt.FormSetting("Save")
        Taskid = Shell("vmenu.exe " & Zarg, 1)
        'If My.Computer.FileSystem.FileExists(ProgDirName) Then
        Call CloseSQL(myConnection) '06-10-09  myConnection.Close() '06-10-09
        '03-19-12 No No frmQuoteRpt.Close() '03-19-12
        End
CLOSE255: FileClose() : Call frmQuoteRpt.FormSetting("Save")
        Call CloseSQL(myConnection) '06-10-09  myConnection.Close() '06-10-09
        '03-19-12 No No frmQuoteRpt.Close() '03-19-12
        End
        'GoTo MSub_Exit ' Close no Shell

QUT255: FileClose() : Call frmQuoteRpt.FormSetting("Save")
        Taskid = Shell("VQUT.exe " & Zarg, 1)
        Call CloseSQL(myConnection) '06-10-09   myConnection.Close() '06-10-09
        '03-19-12 No No frmQuoteRpt.Close() '03-19-12
        End
499:    Resp = MsgBox("PRINTER ERROR =" & Str(Err.Number) & "  " & Str(Erl()) & vbCrLf & "PLEASE FIX", MsgBoxStyle.OkCancel, "Printer Error")
        If Resp = MsgBoxResult.Cancel Then GoTo MSub_Exit 'Cancel
        Resume 999
500:
        If Err.Number = 7 And Erl() = 255 Then
            Msg = "Out of Memory" : FileClose()
            Resp = MsgBox(Msg, MsgBoxStyle.OkOnly, "Visual Basic Error")
            Call CloseSQL(myConnection)
            '03-19-12 No No frmQuoteRpt.Close() '03-19-12
            End
        End If
        'if Error(Err) = "Out of memory" then 04-02-01
        If Err.Number = 7 Or ErrorToString(Err.Number) = "Out of memory" Then ' Out Of Memory on Display Report '04-02-01
            MsgBox("Report is Too Large for the Display Feature - You Should:" & vbCrLf & vbCrLf & " - Print Report Instead of Display" & vbCrLf & " - Select a Totals Only Option if Available" & vbCrLf & " - Select a Smaller Range of Records")
            Resume MSub_Exit
        End If
        If Err.Number = 482 Or Err.Number = 226 Then Msg = "Print Cancelled"
        If Err.Number = 53 Then EM = "FILE NOT ON DISK"
        If Err.Number = 71 Then MsgBox("DISK IS NOT READY **CHECK FLOPPY DRIVE AND TRY AGAIN", MsgBoxStyle.OkOnly, US) : Resume 999
        If Erl() = 3720 Then MsgBox("DISKETTE ERROR ** INSERT ANOTHER DISKETTE AND TRY AGAIN", MsgBoxStyle.OkOnly, US) : Resume 999
        If Erl() = 4025 Then MsgBox("ERROR READING DISKETTE  ** INSERT ANOTHER DISKETTE AND TRY AGAIN", MsgBoxStyle.OkOnly, US) : Resume 999
        ' Following Error Section is Standard
        If Err.Number = 24 Or Err.Number = 25 Or Err.Number = 27 Then GoTo ErrBox
        If Err.Number = 53 Then Msg = "File not on disk at " & Str(Erl()) : GoTo ErrBox
        If Err.Number = 75 Then Msg = "Path/File Access Error at " & Str(Erl()) : GoTo ErrBox
        If Err.Number = 61 Then Msg = "DISK FULL " & vbCrLf & "USE BACKUP/DELETE PROG " : GoTo ErrBox
        If Err.Number = 67 Then Msg = "DIRECTORY FULL" & vbCrLf & "PERFORM BACKUP/DELETE" : GoTo ErrBox
        Msg = "VB Error # = " & Str(Err.Number) & "  ERROR AT " & Str(Erl()) & vbCrLf & "PLEASE READ BACK OF MANUAL" & vbCrLf & "ERROR MESSAGE SECTION Z-2"

ErrBox:
        Resp = MsgBox(Msg & vbCrLf & ErrorToString(Err.Number), MsgBoxStyle.OkCancel, US)
        If Resp = MsgBoxResult.Cancel Then Resume 999 ' GoTo Msub_Exit: 'Cancel
        Resume 999 'Return

998:    On Error GoTo 500 'EndDoc26: 'frmquoterpt.prtAbort.Visible = False ' TOF&CLOSE
999:    Resp = MsgBox("ALL DONE", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, US)
        Cmd = ""
        GoTo MSub_Exit

3360:   EM = "" : If Left(SI, 2) < "01" Or Left(SI, 2) > "12" Then EM = "MONTH ERROR -- PLEASE RE-ENTER" Else If Mid(SI, 3, 2) < "01" Or Mid(SI, 3, 2) > "31" Then EM = "DAY ERROR -- PLEASE RE-ENTER"

MSub_Exit:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
        System.Windows.Forms.Application.DoEvents()
        Exit Sub
    End Sub
End Module
