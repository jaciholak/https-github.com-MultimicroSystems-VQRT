Option Strict Off
Option Explicit On
Imports MySql.Data.MySqlClient
Imports C1.C1Preview
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System '11-04-13 
Imports System.IO '11-04-13 
Imports System.Collections '11-04-13 



Module VQRT3
    '---------------------------------------------------------
    '08-15-08 ErrorRoutine: This allows you to  and try and fix vb errors after the error message **Get gid of  'CatchStop b/4 releasing
    'MsgBox(*VB Error # =  * & Str(Err.Number) & *  ERROR at * & Str(Err.Erl()) & vbCrLf & ErrorToString(Err.Number) & vbCrLf & *In FillSentItems Routine*, 16, *FillSentItems
    ' 'CatchStop     Replace * with DblQuote 
    'Resume Next
    '------------------------------------------------
    'Catch ex As Exception  'Try Catch with  so you can fix exception after the error message **Get gid of  'CatchStop b/4 releasing
    '    MsgBox(ex.Message.ToString & vbCrLf & *SaveToDataBase* & vbCrLf) 
    '     'CatchStop  'Debug.WriteLine(ex.Message.ToString)
    'End Try            Replace * with DblQuote 
    '--------------------------------------------------
    Public Sub FillAdminTables()  '06-05-13
        Try
            dsadmin = New dsSAW8
            dsadmin.adminuser.Clear()
            dsadmin.admingroup.Clear()

            Dim strsql As String = "Select * From AdminUser where userid = '" & SafeSQL(UserID) & "'"
            Dim daAdminU As MySqlDataAdapter = New MySqlDataAdapter
            daAdminU.SelectCommand = New MySqlCommand(strsql, myConnection)
            Dim cbAdminU As MySqlCommandBuilder = New MySqlCommandBuilder(daAdminU)
            daAdminU.Fill(dsadmin, "AdminUser")

            If dsadmin.adminuser.Rows.Count >= 1 Then
                strsql = "Select * From AdminGroup where GroupID = '" & dsadmin.adminuser.Rows(0).Item("GroupID").ToString & "'"
                Dim daAdminG As MySqlDataAdapter = New MySqlDataAdapter
                daAdminG.SelectCommand = New MySqlCommand(strsql, myConnection)
                Dim cbAdminG As MySqlCommandBuilder = New MySqlCommandBuilder(daAdminG)
                daAdminG.Fill(dsadmin, "AdminGroup")
            End If

        Catch ex As Exception
            MessageBox.Show("Error in FillAdminTables" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VADMIN", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub SecuritySubNew(ByRef A As String)
        'From Vquote ''''''''''''''''''''''''''''''''''''''''''''''''
        '10-14-13 If A = "ReportQut" Then SecuritySubNew
        Static Date6 As String
        Static Date5 As String
        Static Date4 As String
        Static Date3 As Short
        Static MMDDYY As String
        Static Resp As Short
        Static msg As String
        Static UnLockCode As String
        Static ResultL As Integer
        Static Enteredpassword As String
        Static PswTry As Short
        Static ExpectedPassword As String
        Static UserToDayDT As String
        Static ToDayDT As String
        Static StrLenL As Integer
        Static DtaFile As String
        Static iSizeL As Integer
100:    On Error Resume Next 'A="SecurityNG" or A = "SecurityOK"
        Static SecuritySetOn As Short ' 0=First Time 1=On for this user 2=not allowed this user
        Static GroupX As String ' Users Group IE: Group2
        If My.Computer.FileSystem.FileExists(UserSysDir & "VADMINNET.INI") = False Then A = "SecurityOK" : Exit Sub ' No Security set on this system on so exit
        iSizeL = 255 : DtaFile = Space(iSizeL)

        If A = "FirstLogOn" Then Call FillAdminTables() '06-05-13 
        If Not dsadmin Is Nothing Then '06-05-13
            If dsadmin.adminuser.Rows.Count <= 0 Then
                A = "SecurityNG"
                Exit Sub
            End If
        Else
            A = "SecurityNG"
            Exit Sub
        End If

        If A = "FirstLogOn" Then GoTo 110


        If SecuritySetOn = 1 Then ' This users Security is already on so Test this function for Yes
            If A = "SecurityGroup" Then '10-15-13
                'Admin = SYSTEM, BRANCH, REGIONAL,
                '"SYSTEM" Then SecurityAdministrator = True = All Branches IE: Ignore BRANCH
                '"BRANCH" If BRANCH, GetBranchCode(UserID) 
                'REGIONAL, Then SecurityBrancheCodes = dsadmin.adminuser.Rows(0).Item("AdminBranches").ToString
                SecurityLevel = dsadmin.adminuser.Rows(0).Item("Admin").ToString
                If SecurityLevel = "SYSTEM" Then
                    SecurityBrancheCodes = "ALL"
                End If
                If SecurityLevel = "BRANCH" Then
                    SecurityBrancheCodes = GetBranchCode(UserID)
                    If SecurityBrancheCodes = "" Then
                        MessageBox.Show("You have not filled in your Branch Code in the Name & Address System." & vbCrLf & "Open the Name & Address System, click on Edit Salesman Table to Correct", "No Branch Code Set Up for " & UserID, MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        A = "SecurityHD"
                        'Else
                        '    A = "" 'JTC No  Fill from Form LoadSecurityOK" '10-28-13 Added "SecurityOK" in SecuritySubNew
                    End If
                End If
                If SecurityLevel = "REGIONAL" Then
                    SecurityBrancheCodes = dsadmin.adminuser.Rows(0).Item("AdminBranches").ToString
                End If
                Exit Sub 'If A = "SecurityGroup" Then 12-28-13 
            End If
            If A = "ReportQut" Then '10-14-13 
                If dsadmin.admingroup.Rows(0).Item("ReportQut").ToString = "N" Then
                    A = "SecurityNG"
                    Exit Sub
                Else 'Report OK & Not SYSTEM so Check BRANCH or REGIONAL
                    ''BRANCH
                    'If dsadmin.adminuser.Rows(0).Item("Admin").ToString = "BRANCH" Then
                    '    'SecurityLevel = dsadmin.adminuser.Rows(0).
                    '    SecurityBrancheCodes = dsadmin.adminuser.Rows(0).Item("AdminBranches").ToString
                    '    SecurityLevel = dsadmin.adminuser.Rows(0).Item("Admin").ToString
                    'End If
                    ''REGIONAL 
                    'If dsadmin.adminuser.Rows(0).Item("Admin").ToString = "REGIONAL" Then
                    '    SecurityBrancheCodes = dsadmin.adminuser.Rows(0).Item("AdminBranches").ToString
                    '    'SecurityLevel = dsadmin.adminuser.Rows(0).Item("Admin").ToString
                    'End If
                    A = "SecurityOK"
                    Exit Sub
                End If

            End If
            Dim FeatureCheck As String = dsadmin.admingroup.Rows(0).Item(A).ToString
            If FeatureCheck = "Y" Then
                A = "SecurityOK"
            ElseIf FeatureCheck = "N" Or FeatureCheck = "" Or FeatureCheck Is Nothing Then  '04-19-06 WNA
                A = "SecurityNG"
            Else
                A = "SecurityHD"
            End If
            Exit Sub
        End If
        If SecuritySetOn = 2 Then A = "SecurityHD" : Exit Sub ' this user not allowed
        'Has Security so determine if this user has logged in today
110:    ToDayDT = Now.ToString("yyyyMMdd")  '10-28-13   VB.Right(VB6.Format(Now, "YYYYMMDD"), 8)
        UserToDayDT = UserID.PadRight(3) & ToDayDT  '10-28-13     UserToDayDT = Left(UserID & "   ", 3) & ToDayDT
        StrLenL = GetPrivateProfileString("UserandDate", "YYYYMMDD", "none", DtaFile, iSizeL, UserDir & "Localvmenuset.ini") '07-08-09
        If "none" = Left(DtaFile, 4) Then GoTo 120 ' Not Logged in at all to \USER\XXX\
        If Left(DtaFile, 3) <> Left(UserID & "   ", 3) Then GoTo 120 ' User Not Logged in   JTCjtc0828 = User & Password
        If Mid(DtaFile, 4, 8) <> ToDayDT Then
            GoTo 120 'User Not Logged in today
        Else
            SecuritySetOn = 1 'Logged in today
            If CheckForFile(UserSysDir & "VADMINNET.INI") = True Then  '06-05-13
                SecurityAdministrator = False
                If dsadmin.adminuser.Rows(0).Item("Admin").ToString = "SYSTEM" Then SecurityAdministrator = True '06-05-13
            Else
                SecurityAdministrator = False
            End If
            A = "SecurityOK" : Exit Sub ' Logged in today IserID & Date in Windows INI so OK
        End If
        'See if this user needs a password
120:
        ExpectedPassword = dsadmin.adminuser.Rows(0).Item("Password").ToString  '06-05-13
        GroupX = dsadmin.adminuser.Rows(0).Item("GroupID").ToString  '06-05-13 
        PswTry = 0 'Not Valid UserID or Password
630:    Enteredpassword = InputBox("This Function Requires a Password." & vbCrLf & "Please Enter Password for access." & vbCrLf & vbCrLf & "See Your System Administrator for help", "Get Password", "xxxxxx") 'Need Text Boxes on Frame
635:
        GoTo GetUnLockCode
636:    If Enteredpassword = UnLockCode Or UCase(Enteredpassword) = "JTCBYPAS" Or Enteredpassword = ExpectedPassword Then
            ResultL = WritePrivateProfileString("UserandDate", "YYYYMMDD", UserToDayDT, UserDir & "Localvmenuset.ini") '07-08-09 Write \USER\XXX\ to allow logon the rest of today
            ResultL = WritePrivateProfileString("Group", "UserGroup", GroupX, UserDir & "Localvmenuset.ini") ''07-08-09 Write \USER\XXX\ to allow logon the rest of today
            iSizeL = 255 : DtaFile = Space(iSizeL)
            If CheckForFile(UserSysDir & "VADMINNET.INI") = True Then  '06-05-13
                SecurityAdministrator = False
                If dsadmin.adminuser.Rows(0).Item("Admin").ToString = "SYSTEM" Then SecurityAdministrator = True '06-05-13
            Else
                SecurityAdministrator = False
            End If
640:        A = "SecurityOK" : SecuritySetOn = 1 : Exit Sub ' Logged in today so OK
        End If
720:    PswTry = PswTry + 1
        msg = "Your UserID or Password is not valid for secured funtions." & vbCrLf & "Click Yes to try your Password again" & vbCrLf & "Click no to continue without secured functions." & vbCrLf & "Contact your Systems Administrator for Security settings"
        Resp = MsgBox(msg, MsgBoxStyle.YesNoCancel, "Security Settings")
        If Resp = MsgBoxResult.Yes And PswTry < 3 Then GoTo 630 'Try again
        A = "SecurityNG" : SecuritySetOn = 2 : Exit Sub ' Not Logged in For Security Can do everything else

800:
GetUnLockCode:  ' 'Multiply positions 2,4 and put total in 5,6  '072800 -> 072856
        MMDDYY = VB6.Format(DateString, "MMddyy") ' Todays Date
        Date3 = Val(Mid(MMDDYY, 2, 1)) * Val(Mid(MMDDYY, 4, 1))
        Date4 = Right("0" & Trim(CStr(Date3)), 2)
        Date5 = Left(MMDDYY, 4) & Date4
        'Add positions 2,4,6 and put total in 1,2   '072856 -> '212856
        Date6 = Right("0" & Trim(Str(Val(Mid(Date5, 2, 1)) + Val(Mid(Date5, 4, 1)) + Val(Mid(Date5, 6, 1)))), 2)
        Date5 = Date6 & Mid(Date5, 3, 4)
        'Add positions 2,4,6 and put total in 3,4   '212856 -> 211556
        Date6 = Right("0" & Trim(Str(Val(Mid(Date5, 2, 1)) + Val(Mid(Date5, 4, 1)) + Val(Mid(Date5, 6, 1)))), 2)
        Date5 = Left(Date5, 2) & Date6 & Mid(Date5, 5, 2)
        UnLockCode = StrReverse(Date5) '08-16-00 'Reverse position  '211556 -> 655112
        'Return
        GoTo 636
    End Sub
    Sub SecuritySub(ByRef A As String)
        'From Vquote ''''''''''''''''''''''''''''''''''''''''''''''''
        If CheckForFile(UserSysDir & "VADMINNET.INI") = True Then '08-30-13
            SecuritySubNew(A)
            Exit Sub
        End If

        Static Date6 As String
        Static Date5 As String
        Static Date4 As String
        Static Date3 As Short
        Static MMDDYY As String
        Static Resp As Short
        Static msg As String
        Static UnLockCode As String
        Static ResultL As Integer
        Static Enteredpassword As String
        Static PswTry As Short
        Static ExpectedPassword As String
        Static UserToDayDT As String
        Static ToDayDT As String
        Static StrLenL As Integer
        Static DiskDrive As String
        Static DtaFile As String
        Static iSizeL As Integer
        Static DFile As String 'UserID$ is already set VMENUSET.INI in SAW7 Dir
        '11-20-06 WNA
        'UserID = "JTC"  'For Testing
100:    On Error Resume Next 'A$="SecurityNG" or A$ = "SecurityOK"
        Static SecuritySetOn As Short ' 0=First Time 1=On for this user 2=not allowed this user
        Static GroupX As String ' Users Group IE: Group2
        '07-08-09 DFile = Dir("VMENUSET.INI") '        If Len(DFile) = 0 Then A = "Normal Lock" : Exit Sub ' No Security set on this system on so exit
        DFile = UserSysDir & "VMENUSET.INI" '07-08-09
        If My.Computer.FileSystem.FileExists(DFile) = False Then A = "SecurityOK" : Exit Sub ' No Security set on this system on so exit
        iSizeL = 255 : DtaFile = Space(iSizeL)
        '07-08-09 If Right(Trim(CurDir()), 1) <> "\" Then DiskDrive = Trim(CurDir()) & "\" Else DiskDrive = Trim(CurDir())
        If A = "FirstLogOn" Then GoTo 110
        'Only need the following in Menu for log off
        'If A$ = "LogOff" Then
        '      UserToDayDT$ = Left$(UserID$ & "   ", 3) & "20000101"
        '      ResultL& = WritePrivateProfileString("UserandDate", "YYYYMMDD", UserToDayDT$, "Localvmenuset.ini") ' Write to Windows INI an old Date
        '      frmMainM.mnuSecurityLogOff.Visible = False
        '      SecuritySetOn% = 0  'Off
        '      Exit Sub
        'End If
        If SecuritySetOn = 1 Then ' This users Security is already on so Test this function for Yes
            iSizeL = 255 : DtaFile = Space(iSizeL) '10-04-09 SLS#,SLS% = Yes No Hide
            StrLenL = GetPrivateProfileString(GroupX, A, "", DtaFile, iSizeL, UserSysDir & "VMENUSET.INI") 'Get Group
            Dim B As String = Trim(DtaFile) '10-04-09 Need to Save A = "QuoteView" or "QutLock" Chg to B
            If Len(B) <> 0 Then B = Left(B, Len(B) - 1)
            If Trim(B) = "" Then B = "Normal Lock"
            If A = "ReportQut" Then
                If B = "No" Then
                    A = "SecurityNG" : Exit Sub '10-05-09 JTC  Kicks User out of VQUT 
                Else 'Yes Option
                    A = B : Exit Sub 'Yes Option '10-05-09 JTC 
                End If
            End If
            If Left$(DtaFile$, 3) = "Yes" Then
                A$ = "SecurityOK"
            ElseIf Left$(A, 2) = "No" Then  '10-04-09 ????????? Logic 
                A$ = "SecurityNG" ' Can't 
            ElseIf Left$(A, 2) = "Hide" Then  '10-04-09 ????????? Logic   '10-04-09 SLS#,SLS% = Yes No Hide  
                A$ = "SecurityHD" ' "Hide
            End If
            Exit Sub
        End If
        '04-19-06 WNA If SecuritySetOn% = 2 Then A$ = "SecurityNG": Exit Sub ' this user not allowed
        If SecuritySetOn = 2 Then A = "SecurityHD" : Exit Sub ' this user not allowed
        'Has Security so determine if this user has logged in today
110:    ToDayDT = Now.ToString("yyyyMMdd")  '10-28-13   VB.Right(VB6.Format(Now, "YYYYMMDD"), 8)
        UserToDayDT = UserID.PadRight(3) & ToDayDT  ' VB.Left(UserID & "   ", 3) & ToDayDTToDayDT = Right(VB6.Format(Now, "YYYYMMDD"), 8)        UserToDayDT = Left(UserID & "   ", 3) & ToDayDT
        iSizeL = 255 : DtaFile = Space(iSizeL) '10-04-09
        StrLenL = GetPrivateProfileString("UserandDate", "YYYYMMDD", "none", DtaFile, iSizeL, UserDir & "Localvmenuset.ini") '07-08-09
        If "none" = Left(DtaFile, 4) Then GoTo 120 ' Not Logged in at all to \USER\XXX\
        If Left(DtaFile, 3) <> Left(UserID & "   ", 3) Then GoTo 120 ' User Not Logged in   JTCjtc0828 = User & Password
        If Mid(DtaFile, 4, 8) <> ToDayDT Then
            GoTo 120 'User Not Logged in today
        Else
            StrLenL = GetPrivateProfileString("Group", "UserGroup", "", DtaFile, iSizeL, UserDir & "Localvmenuset.ini") '07-08-09
            GroupX = Left(DtaFile, 6) : If Left(DtaFile, 5) <> "Group" Then GroupX = "Group3"
            '08-03-05 WNA frmMainM.mnuSecurityLogOff.Visible = True:
            SecuritySetOn = 1
            iSizeL = 255 : DtaFile = Space(iSizeL)
            StrLenL = GetPrivateProfileString(UserID, "Admin", "none", DtaFile, iSizeL, UserSysDir & "VMENUSET.INI")
            SecurityAdministrator = False
            'If VB.Left(DtaFile$, 4) = "none" Then
            If DtaFile.StartsWith("none") Then
                SecurityAdministrator = False 'User Not In INI File
                'ElseIf "Yes" = VB.Left(DtaFile$, 3) Then
            ElseIf DtaFile.StartsWith("Yes") Then
                SecurityAdministrator = True 'admin
                'GoTo 110 '11-03-06 See if Admin is logged in today
                '11-03-06 GoTo Exit_Done
            End If
            A = "SecurityOK" : Exit Sub ' Logged in today IserID & Date in Windows INI so OK
        End If
        'See if this user needs a password
120:    'If Right(Trim(CurDir()), 1) <> "\" Then DiskDrive = Trim(CurDir()) & "\" Else DiskDrive = Trim(CurDir())
        iSizeL = 255 : DtaFile = Space(iSizeL)
        StrLenL = GetPrivateProfileString(UserID, "Password", "none", DtaFile, iSizeL, UserSysDir & "VMENUSET.INI") '07-08-09
        If "none" = Left(DtaFile, 4) Then A = "SecurityNG Not in List" : SecuritySetOn = 2 : Exit Sub 'User Not In INI File
        ExpectedPassword = Left(DtaFile, StrLenL)
        StrLenL = GetPrivateProfileString(UserID, "Group", "", DtaFile, iSizeL, UserSysDir & "VMENUSET.INI") 'Get Group
        GroupX = Left(DtaFile, 6) : If Left(DtaFile, 5) <> "Group" Then GroupX = "Group3"

        PswTry = 0 'Not Valid UserID$ or Password
630:    Enteredpassword = InputBox("This Function Requires a Password." & vbCrLf & "Please Enter Password for access." & vbCrLf & vbCrLf & "See Your System Administrator for help", "Get Password", "xxxxxx") 'Need Text Boxes on Frame
635:    'GoSub GetUnLockCode: ' Password, UnlockCode or JTCBYPAS will work
        GoTo GetUnLockCode
636:    If Enteredpassword = UnLockCode Or UCase(Enteredpassword) = "JTCBYPAS" Or Enteredpassword = ExpectedPassword Then
            ResultL = WritePrivateProfileString("UserandDate", "YYYYMMDD", UserToDayDT, UserDir & "Localvmenuset.ini") '07-08-09 Write \USER\XXX\ to allow logon the rest of today
            ResultL = WritePrivateProfileString("Group", "UserGroup", GroupX, UserDir & "Localvmenuset.ini") ''07-08-09 Write \USER\XXX\ to allow logon the rest of today
            '08-03-05 WNA frmMainM.mnuSecurityLogOff.Visible = True
            iSizeL = 255 : DtaFile = Space(iSizeL)
            StrLenL = GetPrivateProfileString(UserID, "Admin", "none", DtaFile, iSizeL, UserSysDir & "VMENUSET.INI") '07-08-09 
            SecurityAdministrator = False
            'If VB.Left(DtaFile$, 4) = "none" Then
            If DtaFile.StartsWith("none") Then
                SecurityAdministrator = False 'User Not In INI File
                'ElseIf "Yes" = VB.Left(DtaFile$, 3) Then
            ElseIf DtaFile.StartsWith("Yes") Then
                SecurityAdministrator = True 'admin
                'GoTo 110 '11-03-06 See if Admin is logged in today
                '11-03-06 GoTo Exit_Done
            Else
                SecurityAdministrator = False
            End If
640:        A = "SecurityOK" : SecuritySetOn = 1 : Exit Sub ' Logged in today so OK
        End If
720:    PswTry = PswTry + 1
        msg = "Your UserID or Password is not valid for secured funtions." & vbCrLf & "Click Yes to try your Password again" & vbCrLf & "Click no to continue without secured functions." & vbCrLf & "Contact your Systems Administrator for Security settings"
        Resp = MsgBox(msg, MsgBoxStyle.YesNoCancel, "Security Settings")
        If Resp = MsgBoxResult.Yes And PswTry < 3 Then GoTo 630 'Try again
        A = "SecurityNG" : SecuritySetOn = 2 : Exit Sub ' Not Logged in For Security Can do everything else

800:
GetUnLockCode:  ' 'Multiply positions 2,4 and put total in 5,6  '072800 -> 072856
        MMDDYY = VB6.Format(DateString, "mmddyy") ' Todays Date
        Date3 = Val(Mid(MMDDYY, 2, 1)) * Val(Mid(MMDDYY, 4, 1))
        Date4 = Right("0" & Trim(CStr(Date3)), 2)
        Date5 = Left(MMDDYY, 4) & Date4
        'Add positions 2,4,6 and put total in 1,2   '072856 -> '212856
        Date6 = Right("0" & Trim(Str(Val(Mid(Date5, 2, 1)) + Val(Mid(Date5, 4, 1)) + Val(Mid(Date5, 6, 1)))), 2)
        Date5 = Date6 & Mid(Date5, 3, 4)
        'Add positions 2,4,6 and put total in 3,4   '212856 -> 211556
        Date6 = Right("0" & Trim(Str(Val(Mid(Date5, 2, 1)) + Val(Mid(Date5, 4, 1)) + Val(Mid(Date5, 6, 1)))), 2)
        Date5 = Left(Date5, 2) & Date6 & Mid(Date5, 5, 2)
        UnLockCode = StrReverse(Date5) '08-16-00 'Reverse position  '211556 -> 655112
        'Return
        GoTo 636
    End Sub
    Public Function GetFIRSTDayInMonth(ByVal dDate As Date) As Date
        On Error Resume Next
        '10-15-10
        'get first day of this month
        Return DateSerial(dDate.Year, dDate.Month, 1)


    End Function
    Public Function GetLastDayInMonth(ByVal dtDate As Date) As Date
        '10-15-10
        'get first day of this month
        Dim tempdate1 As DateTime = DateSerial(dtDate.Year, dtDate.Month, 1)

        'gets the first day of the next month
        Dim tempdate2 As DateTime = DateAdd("m", 1, tempdate1)

        'subtract a second to get 11:59:59 PM of the last day of the month/year passed in dDate
        Return DateAdd("s", -1, tempdate2)

        ''Function  GetLastDayInMonth("#2009-02-20#")'04-20-10 JTC 
        ''First Day Of Month = CDate(Format(dtDate, "yyyy-") & Format(dtDate, "MM-") & "01") '04-20-10**************
        'Dim NewDate As New Date ' 
        'NewDate = dtDate ' CDate(Format(dtDate, "yyyy-") & Format(dtDate, "MM-") & "01") '04-20-10 Start Of Month 
        'NewDate = NewDate.AddMonths(1) : NewDate = NewDate.AddDays(-1)
        'Return NewDate
        ''*********************************************************************
        ' ''example for #2009-02-20# we want to get the last day in the month 02,
        ' '' (ie. date for last day in Feb)
        ''Return DateAdd(DateInterval.Day, _
        ''      (Day(DateAdd(DateInterval.Month, 1, dtDate))) * -1, _
        ''       DateAdd(DateInterval.Month, 1, dtDate))

    End Function
    Public Sub FillDataSet(ByVal myform As frmQuoteRpt)
        Try
            'Dim strsql As String = ""
            'Dim dsQuote As dsSaw8 = New dsSaw8 '01-15-09
            dsQuote = New dsSaw8
            dsQuote.EnforceConstraints = False '12-03-09 
            strSql = "Select * from Quote limit 1" '02-12-09
            daQuote = New MySqlDataAdapter
            daQuote.SelectCommand = New MySqlCommand(strsql, myConnection)
            Dim cbQut As MySql.Data.MySqlClient.MySqlCommandBuilder
            cbQut = New MySqlCommandBuilder(daQuote)
            daQuote.Fill(dsQuote, "quote")

            'strsql = "Select * from project LIMIT 1 " '02-12-09"
            'daProject = New MySqlDataAdapter
            'daProject.SelectCommand = New MySqlCommand(strsql, myConnection)
            'Dim cbProj As MySql.Data.MySqlClient.MySqlCommandBuilder
            'cbProj = New MySqlCommandBuilder(daProject)
            'daProject.Fill(dsQuote, "project")

            strSql = "Select * from projectcust LIMIT 1 " '02-12-09"
            daProjCust = New MySqlDataAdapter
            daProjCust.SelectCommand = New MySqlCommand(strSql, myConnection)
            Dim cbProjCust As MySql.Data.MySqlClient.MySqlCommandBuilder
            cbProjCust = New MySqlCommandBuilder(daProjCust)
            daProjCust.Fill(dsQuote, "projectcust")

            strSql = "Select * from qutnotes LIMIT 1 " '02-12-09"
            daQuoteNotes = New MySqlDataAdapter
            daQuoteNotes.SelectCommand = New MySqlCommand(strSql, myConnection)
            Dim cbQutNote As MySql.Data.MySqlClient.MySqlCommandBuilder
            cbQutNote = New MySqlCommandBuilder(daQuoteNotes)
            daQuoteNotes.Fill(dsQuote, "qutnotes")

            strSql = "Select * from quotelines LIMIT 1 " '02-12-09"
            daQuoteLine = New MySqlDataAdapter
            daQuoteLine.SelectCommand = New MySqlCommand(strSql, myConnection)
            Dim cbQutLin As MySql.Data.MySqlClient.MySqlCommandBuilder
            cbQutLin = New MySqlCommandBuilder(daQuoteLine)
            daQuoteLine.Fill(dsQuote, "quotelines")

            '08-26-09 Out This is for Future Column header Customization
            'strsql = "Select * from qutlineprice LIMIT 1 " '02-12-09"
            'daQuoteLinePrice = New MySqlDataAdapter
            'daQuoteLinePrice.SelectCommand = New MySqlCommand(strsql, myConnection)
            'Dim cbQutLinPrc As MySql.Data.MySqlClient.MySqlCommandBuilder
            'cbQutLinPrc = New MySqlCommandBuilder(daQuoteLinePrice)
            'daQuoteLinePrice.Fill(dsQuote, "qutlineprice")

            'strsql = "Select * from qutslssplit LIMIT 1 " '02-12-09"
            'daQuoteSLS = New MySqlDataAdapter
            'daQuoteSLS.SelectCommand = New MySqlCommand(strsql, myConnection)
            'Dim cbQutSLS As MySql.Data.MySqlClient.MySqlCommandBuilder
            'cbQutSLS = New MySqlCommandBuilder(daQuoteSLS)
            'daQuoteSLS.Fill(dsQuote, "qutslssplit")

            'strsql = "Select * from projectcust LIMIT 1 " '02-12-09"
            'daQuoteTo = New MySqlDataAdapter
            'daQuoteTo.SelectCommand = New MySqlCommand(strsql, myConnection)
            'Dim cbQuoteTo As MySql.Data.MySqlClient.MySqlCommandBuilder
            'cbQuoteTo = New MySqlCommandBuilder(daQuoteTo)
            'daQuoteTo.Fill(dsQuote, "projectcust")

            'NoChangeTest = True
            '12-03-12 JH myform.QuoteBindingSource.DataSource = dsQuote
            '12-03-12 JH myform.QuoteBindingSource.DataMember = "quote"
            'NoChangeTest = False
            'myform.ProjectBindingSource.DataSource = dsQuote.project
            'myform.QuotequotelinesBindingSource.DataSource = dsQuote.quotelines
            'myform.QuotequtslssplitBindingSource.DataSource = dsQuote.qutslssplit
            'myform.QuoteNotesBindingSource.DataSource = dsQuote.qutnotes

            'dsQutLU.Tables("ProductPrice").Clear()
            'strsql = "Select * from ProductPrice" ' WHERE ProdID = " & "'" & "" & "'"
            'daProductPrice = New MySqlDataAdapter
            ''daProductPrice.SelectCommand = New MySqlCommand(strsql, myConnection)
            'cbPP = New MySqlCommandBuilder(daProductPrice)
            'daProductPrice.Fill(dsPrice, "ProductPrice")
            'FrmProduct.ProductProductPriceBindingSource.DataSource = dsPrice.Tables("productprice")

            'dsPriceLU.Tables("ProductPrice").Clear()
            'strsql = "Select * from ProductPrice" ' WHERE ProdID = " & "'" & "" & "'"
            'Dim daProductPrice2 As New MySqlDataAdapter
            'daProductPrice2.SelectCommand = New MySqlCommand(strsql, myConnection)
            'Dim cbPP2 As New MySqlCommandBuilder(daProductPrice2)
            'daProductPrice2.Fill(dsPriceLU, "ProductPrice")
        Catch ex As Exception
            MessageBox.Show("Error in FillDataSet (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox("FillDataSet " & ex.Message)
            ' If DebugOn ThenStop
        End Try

   
    End Sub
    Public Sub LockingRecord(ByVal Func As String, ByVal Answer As String) '09-11-09
        'Func = Read = Read and notify if Locked   Answer = OK, LockedOut, ReadOnly
        'Func = ReadnLock = Read then Lock         Answer = OK, Locked, ReadOnly, NoRights
        'Func = Write = Write if not locked        Answer = OK, Locked, ReadOnly, NoChangedBySomeoneElse, NoRights
        'Func = WritenLock = Write and lock        Answer = OK, Locked, ReadOnly, NoRights
        'Func = ClearLock = Admin clear any lock   Answer = OK, NoRights
        Dim Msg As String = ""
Read:
        'If Security is on, Read quote Privleges
        If Func.StartsWith("Read") Then 'On any Read see if Locked
            'If User has ReadOnly Rights then Answer = "OK" : exit sub 'They can't Save Changes
            'strsql = "Select Quote.LockedBy, Quote.LastDateTime from quote Where where quote.QuoteCode = '" & QuoteCode & "'"
            'If Quote.LockedBy.text.trim <> "" Then ' Locked By Someone
            '    If Quote.LockedBy = UserID Then Answer = "OK" : goto readnLock' 'OK I locked It
            '    'Msg "This Quote is Locked By " & Quote.LockedBy & "  At  " & Format(LastDateTime, "HH:MM:ss")
            '     Answer = "LockedOut" : Goto MsgDisplay:
            '     End if
            'Else
            'Answer = "OK" :'PublicSaveTimeStamp = Quote.LastDateTime 'Save It
            'GoTo readnLock '
            'End If

        End If
ReadnLock:
        If Func.StartsWith("ReadnLock") Then 'Read Then Lock
            ''You have already done the read steps above so Record is not Locked
            'If User has no Lock Rights then Answer = NoRights :Msg = "The User has no Lock Rights." : Goto  MsgDisplay:
            'strSql = "Update quote set quote.LockedBy = '" & UserID & "' where quote.QuoteCode = '" & QuoteCode & "'"
            'PublicSaveTimeStamp = Quote.LastDateTime
            ''If Quote.LockedBy.text.trim <> "" Then ' Locked By Someone
            ''    If Quote.LockedBy = UserID Then Answer = "OK": Exit Sub 'OK I locked It
            ''    'Msg "This Quote is Locked By " & Quote.LockedBy & "  At  " & Format(LastDateTime, "HH:MM:ss")
            ''    Answer = "LockedOut": Goto  MsgDisplay
            ''End If
        End If

        If Func.StartsWith("Write") Then 'Write if not locked
            'If User has no Write Rights then Answer = "NoRights" : Msg = "The User has no Write Rights." :Goto  MsgDisplay:
            'strsql = "Select Quote.LockedBy, Quote.LastDateTime from quote Where where quote.QuoteCode = '" & QuoteCode & "'"
            'If (Quote.LockedBy.text.trim = "" or Quote.LockedBy = UserID) then ' Not Locked or Locked by Me so KK
            'If PublicSaveTimeStamp = Quote.LastDateTime Then ' Not Changed either
            'This Quote has not been changed since I last read it so ok to write it
            'Answer = "OK" : exit sub 
            'Else
            ' Msg "Someone has updated this quote since you last accessed it." & vbcrlf "So you must reread it and then make your changes again"
            ' Msg = Msg & "It was last changed By " & Quote.LastChgBy & "  At  " & Format(LastDateTime, "HH:MM:ss")
            ' Answer = "NoChangedBySomeoneElse": Goto  MsgDisplay:
            'End If
        End If
        If Func.StartsWith("WritenLock") Then 'Write and Lock
            'If User has no Write Rights then Answer = "NoRights" :Msg = "The User has no Write Rights." :Goto  MsgDisplay:
            'If Quote.LockedBy.text.trim = "" or Quote.LockedBy = UserID Then ' Not Locked or Locked by Me
            'strSql = "Update quote set quote.LockedBy = '" & UserID & "' where quote.QuoteCode = '" & QuoteCode & "'":Answer ="OK":exit sub
            'else 
            '    'Msg "This Quote is Locked By " & Quote.LockedBy & "  At  " & Format(LastDateTime, "HH:MM:ss")
            '   Answer = "LockedOut": Goto  MsgDisplay
            'End If
        End If
        If Func.StartsWith("ClearLock") Then 'Clear Lock if Admin
            'If User is not Admin then Answer = "NoRights" :  Msg ="You do not have rights to Clear Locks." : Goto  MsgDisplay
            'Msg "This Quote is Locked By " & Quote.LockedBy & "  At  " & Format(LastDateTime, "HH:MM:ss")
            'msgBox(Are you sure you want to remove this lock?)
            'If Resp = Yes then
            '    strSql = "Update quote set quote.LockedBy = '" & "" & "' where quote.QuoteCode = '" & QuoteCode & "'":answer = "OK:exit sub
            'End if
        End If
MsgDisplay:
        MsgBox(Msg)

    End Sub
    Public Sub FillQutLUDataSet(ByVal SortSeq As String, ByVal SortDir As String, Optional ByVal SortCode As String = "")
        Try

            'Debug.Print(OrderBy)
            dsQutLU = New dsSaw8
            dsQutLU.EnforceConstraints = False
            ''SQL SECTION *********************************************************************************************
            'Dim strsql As String = ""
            ''SELECT project.ProjectName, Quote * FROM quote LEFT OUTER JOIN project ON quote.ProjectID = project.ProjectID
            ''strsql = "Select project.ProjectName, project.ProjectID, quote.QuoteID, quote.QuoteCode, quote.EntryDate, quote.EndDate, quote.BidDate, quote.Status, quote.SLSQ, quote.EnteredBy, quote.Sell from quote left join project on quote.ProjectID = project.ProjectID "
            'strsql = "Select Q.*, P.projectname, P.MarketSegment from Quote Q join Project P on P.ProjectID = Q.ProjectID " 'where Q.EntryDate > " & "'" & "2008-05-08" & "'" ' " '01-25-09 "'" & "'" & "M%" & "'"

            'If SortCode <> "" Then
            '    strsql += "where " & SortSeq & " >= '" & SortCode & "' "
            'End If

            'Dim sStartDate As String = VB6.Format(frmQuoteRpt.DTPickerStartEntry.Text, "yyyy-MM-dd") '  VB6.Format("01/01/2008", "yyyy-MM-dd") 'Testing
            'Dim sEndDate As String = VB6.Format(frmQuoteRpt.DTPicker1EndEntry.Text, "yyyy-MM-dd")
            ''If DebugOn Then 'Testing Only
            ''    sStartDate = VB6.Format("01/01/2008", "yyyy-MM-dd") 'Testing
            ''    sEndDate = VB6.Format("12/31/2009", "yyyy-MM-dd")
            ''End If
            ''11-13-09------------------------------------------------------------------
            ''If DefTypeOfJob <> "All" And DefTypeOfJob.Trim <> "" Then '11-12-09 JH
            ''    If SortCode.Trim <> "" Then strsql += " and " Else strsql += " where "
            ''    Dim JT As String = ""
            ''    If DefTypeOfJob = "Quotes" Then JT = "Q"
            ''    If DefTypeOfJob = "Planned Projects" Then JT = "P"
            ''    If DefTypeOfJob = "Spec Credit" Then JT = "S"
            ''    If DefTypeOfJob = "Submittals" Then JT = "T"
            ''    If DefTypeOfJob = "Other" Then JT = "O"
            ''    strsql += " Quote.TypeOfJob = '" & JT & "' "
            ''    hit = True
            ''End If
            'Dim JT As String = frmQuoteRpt._fdTypeofJob.Text
            'If JT = "A" Then '=ALL 11-13-09 No Select on Q.TypeOfJob JT = "*"
            '    '11-13-09 strsql += "where Q.EntryDate >= '" & sStartDate & "' and Q.BidDate <= '" & sEndDate & "' " & OrderBy '01-30-09"order by Q.QuoteCode " ' and Q.EndDate =< '20091231' " & "order by Q.QuoteCode " '01-25-09
            '    strsql += "where Q.EntryDate >= '" & sStartDate & "' and Q.Entrydate <= '" & sEndDate & "' " & OrderBy '11-13-09 01-30-09"order by Q.QuoteCode " ' and Q.EndDate =< '20091231' " & "order by Q.QuoteCode " '01-25-09
            'Else
            '    '11-13-09 strsql += "where Q.EntryDate >= '" & sStartDate & "' and Q.BidDate <= '" & sEndDate & "' " & OrderBy '01-30-09"order by Q.QuoteCode " ' and Q.EndDate =< '20091231' " & "order by Q.QuoteCode " '01-25-09
            '    strsql += "where Q.TypeOfJob = '" & JT & "' and Q.EntryDate >= '" & sStartDate & "' and Q.Entrydate <= '" & sEndDate & "' " & OrderBy '11-13-09 01-30-09"order by Q.QuoteCode " ' and Q.EndDate =< '20091231' " & "order by Q.QuoteCode " '01-25-09
            'End If
            ''Debug.Print(strsql)
            ''strsql += "order by " & SortSeq & " " & SortDir
            ''END SQL SECTION *****************************************************************************************
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Project Shortage Report" Then GoTo SkipBranch '10-16-13
            '10-15-13 Added below
            If My.Computer.FileSystem.FileExists(UserSysDir & "VADMINNET.INI") = True Or SecurityLevel = "" Then '10-16-13 JTC Added No Security
                If SecurityLevel = "BRANCH" Or SecurityLevel = "REGIONAL" Or SecurityBrancheCodes.Trim.ToUpper <> "ALL" Then '10-16-13 JTC Added No Security Or SecurityBrancheCodes.Trim.ToUpper <> "ALL"
                    Dim STR1 As String = strSql.Substring(0, strSql.IndexOf("ORDER BY"))
                    Dim STR2 As String = " " & strSql.Substring(strSql.IndexOf("ORDER BY"))
                    Dim BC As String = ""
                    If SecurityBrancheCodes.Trim.ToUpper <> "ALL" Then
                        If SecurityBrancheCodes.Contains(",") = True Then
                            BC = " and ( BranchCode = '" & SecurityBrancheCodes.Replace(",", "' or BranchCode = '") & "' or BranchCode = ''  )"
                        Else
                            BC = " and (BranchCode = '" & SecurityBrancheCodes & "'" & " or BranchCode = '')"
                        End If
                    End If
                    strSql = STR1 & BC & STR2
                End If
            End If
SkipBranch:  '10-16-13
            daQutLU = New MySqlDataAdapter
            daQutLU.SelectCommand = New MySqlCommand(strSql, myConnection)
            Dim cbQutlu As MySql.Data.MySqlClient.MySqlCommandBuilder
            cbQutlu = New MySqlCommandBuilder(daQutLU)
            daQutLU.Fill(dsQutLU, "QUTLU1")

            frmQuoteRpt.QUTLU1BindingSource.DataSource = dsQutLU.QUTLU1
            'myView.Sort = "FirmName"
            'frmQuoteRpt.QUTLU1TableAdapterBindingSource = dsSaw8.QUTLU1DataTable

            frmQuoteRpt.tgQh.Rebind(True)
            'For I = 0 To frmQuoteRpt.tgQh.Splits(0).DisplayColumns.Count - 1
            '    'If I > 41 Then Exit For '02-19-09 
            '    Dim col As C1.Win.C1TrueDBGrid.C1DisplayColumn = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I) '02-20-09 
            '    col.DataColumn.Tag = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name.ToString '07-09-09 Tag = Name col.DataColumn.Tag = I.ToString 'Add Tag to each Column
            'Next
            'frmQuoteRpt.tgQutLU.Columns(Mid(SortSeq, InStr(SortSeq, ".") + 1)).SortDirection = IIf(SortDir = "ASC", C1.Win.C1TrueDBGrid.SortDirEnum.Ascending, C1.Win.C1TrueDBGrid.SortDirEnum.Descending)
            'frmQuoteRpt.tslSortBy.Text = "Sort By: " & Mid(SortSeq, InStr(SortSeq, ".") + 1) & " " & IIf(SortDir = "ASC", "Ascending", "Decending")

        Catch ex As Exception
            MessageBox.Show("Error in FillQutLUDataSet (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox("FillQutLUDataSet " & ex.Message)
            ' If DebugOn ThenStop
        End Try
    End Sub
    Public Sub FillQutRealLUDataSet(ByRef SortSeq As String, ByVal SortDir As String, ByVal OnlyInFluenceGtZero As Boolean, Optional ByVal SortCode As String = "")
        Try
            Dim SaveStrSQL2 As String = "" 'strSql '11-27-13 Moved to top 
            Dim SaveStrSQL As String = "" 'strSql '11-27-13 Moved to top
            Dim WhereS As String = "" '06-25-13 WhereS 05-17-11 JTC Added Dates to Realization Report Added ProjectCust typeC 
            Dim MyCommand As New MySqlCommand '03-22-13
            Dim BC As String = "" '02-05-14 JTC Added
            Dim OneNCodeOnly As String = "" '07-15-14 
            Dim QuotesToOneMFG As String = "" '04-16-15 JTC
            Dim WhereSqlTypec As String = strSql '04-16-15 JTC WhereSqlTypec = Replace(strSql, WhereSqlTypec, "") '04-16-15 Jtc Save this Where Used on QuotesToOneMFG
            '04-30-15 JTC Moved Up
            Dim CTypes As String = "" 'frmQuoteRpt.cboTypeCustomer.Text '01-06-15 MOVED UP, WAS INSIDE IF
            dsQuoteRealLU = New dsSaw8
            dsQuoteRealLU.EnforceConstraints = False 'projectcust.*,  

100:        '08-20-11 If frmQuoteRpt.txtPrimarySortSeq.Text = "Architect" ThenStop '09-19-11
            If frmQuoteRpt.chkCustomerBreakdown.CheckState = CheckState.Checked Then '06-06-11 = "Show All Quote Header Fields" Then '06-06-11 "Add Cust QuoteTo Breakdown to Report"
                '06-02-11 Better But Slow ******************************************************************************
                '08-19-11 Why Limit 1 Need Latest Cust ??????????????????
                '08-28-12 JTC Add , quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, to SQL
                '08-28-12 strSql = "SELECT projectcust.*, quote.*, " '
                '07-06-17 COMMA OFF AT END strSql = "SELECT projectcust.*, quote.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, " '08-28-12 quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, 
                strSql = "SELECT projectcust.*, quote.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ " '07-06-17 COMMA OFF AT END 
                '08-20-11 strSql += " (Select projectcust.FirmName from projectcust where projectcust.Typec = 'A' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Architect, "
                If frmQuoteRpt.txtPrimarySortSeq.Text = "Architect" Then  '08-20-11
                    strSql += " (Select projectcust.FirmName from projectcust where projectcust.Typec = 'A' and Quote.QuoteID = projectcust.QuoteID) as Architect " '08-20-11
                    strSql += " FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID "
                End If

                '07-06-17 strSql += " (Select projectcust.FirmName from projectcust where projectcust.Typec = 'A' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Architect, " '07-22-14 
                '07-06-17 strSql += " (Select projectcust.FirmName from projectcust where projectcust.Typec = 'T' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Contractor, "
                '07-06-17 strSql += " (Select projectcust.FirmName from projectcust where projectcust.Typec = 'E' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Engineer, "
                '07-06-17 strSql += " (Select projectcust.FirmName from projectcust where projectcust.Typec = 'S' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Specifier, "
                '07-06-17 strSql += " (Select projectcust.FirmName from projectcust where projectcust.Typec = 'O' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Other " '11-29-11 JTC Chg Typec From X to O=Other
                strSql += " FROM Quote "
                If RealTgLookupExcel = True Then '11-27-13
                    strSql += " LEFT OUTER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID " '11-27-13 JTC LEFT OUTER
                Else
                    strSql += " INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID "
                End If
            Else
                'Not .chkCustomerBreakdown.
                '06-06-11 = "Show All Quote Header Fields"
200:            '05-31-11 Not All Fields is Much Faster 'added quote.Sell as QSell, quote.Cost as QCost, quote.Comm as QComm, to put on Specifier Value
                '07-27-12 Removed Didn't work IF(NCode = '', NCode, " Blank"), 
                '12-30-12 JTC SELECT projectcust.*, IF(projectcust.NCODE = '', Left(projectcust.FIRMNAME,6), projectcust.ncode) as NCODE, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID where Quote.TypeOfJob = 'Q'  and  Quote.EntryDate >= '2014-12-01' and Quote.EntryDate <= '2014-12-31'  and (projectcust.TypeC = 'X'  or projectcust.TypeC = 'A'  or projectcust.TypeC = 'E'  or projectcust.TypeC = 'L'  or projectcust.TypeC = 'S'  or projectcust.TypeC = 'T'  or projectcust.TypeC = 'X' )  order by projectcust.Ncode, projectcust.SLSCode, JobName
                '12-31-14 JTC Use FirmName If NCode is Blank
                'PC(names)ProjectCustID, ProjectID, QuoteCode, NCode, FirmName, ContactName, SLSCode, Got, Typec, MFGQuoteNumber, Cost, Sell, Comm, Overage, QuoteToDate, OrdDate, NotGot, Comments, SPANumber, SpecCross, LotUnit, LPCost, LPSell, LPComm, LampsIncl, Terms, FOB, QuoteID, BranchCode, LeadTime, LastChgDate, LastChgBy, Requested, FileName
                '12-31-14 strSql = "SELECT projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID " '09-14-10 INNER JOIN Projectcust ON Quote.QuoteID = projectcust.QuoteID INNER JOIN quote ON projectcust.QuoteID = quote.QuoteID "
                'ProjectCust (names)                                                                                 ???ProjectCustID, ProjectID,           QuoteCode, <<NCode,   FirmName,               ContactName,            SLSCode,             Got,              Typec,             MFGQuoteNumber,           Cost,              Sell,             Comm,            Overage,             QuoteToDate,             OrdDate,              NotGot,             Comments,             SPANumber,             SpecCross,             LotUnit,             LPCost,             LPSell,             LPComm              LampsIncl              Terms              FOB              QuoteID,             BranchCode              LeadTime,             LastChgDate,             LastChgBy,            Requested,              FileName
                '01-06-15 JH strSql = "Select IF(projectcust.NCODE = '', Left(projectcust.FIRMNAME,6), projectcust.ncode) as NCode,  projectcust.ProjectID, projectcust.QuoteCode, projectcust.FirmName, projectcust.ContactName, projectcust.SLSCode, projectcust.Got, projectcust.Typec, projectcust.MFGQuoteNumber, projectcust.Cost, projectcust.Sell, projectcust.Comm, projectcust.Overage, projectcust.QuoteToDate, projectcust.OrdDate, projectcust. NotGot, projectcust.Comments, projectcust.SPANumber, projectcust.SpecCross, projectcust.LotUnit, projectcust.LPCost, projectcust.LPSell, projectcust.LPComm, projectcust.LampsIncl, projectcust.Terms, projectcust.FOB, projectcust.QuoteID, projectcust.BranchCode, projectcust.LeadTime, projectcust.LastChgDate, projectcust.LastChgBy, projectcust.Requested, projectcust.FileName, " '12-31-14 JTC
                '01-06-15 JH strSql += " quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID "
                'strSql = "SELECT projectcust.QuoteID, IF(projectcust.NCODE = '', Left(projectcust.FIRMNAME,6), projectcust.ncode) as NCode, projectcust.Typec, projectcust.Firmname, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID "
                'strSql = "SELECT projectcust.QuoteID, projectcust.NCode as NNCODE, IF(projectcust.NCODE = '', Left(projectcust.FIRMNAME,6), projectcust.ncode) as NNCODE, projectcust.Typec, projectcust.Firmname, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID " '12-31-14 
                '06-25-15 JTC added Ucae  and Q.Status <> '" & "NOREPT" & "' 
205:            strSql = "SELECT IF(projectcust.NCODE = '', UCASE(Left(projectcust.FIRMNAME,6)), UCASE(projectcust.ncode)) as NCode,  projectcust.ProjectID, projectcust.QuoteCode, projectcust.FirmName, projectcust.ContactName, projectcust.SLSCode, projectcust.Got, projectcust.Typec, projectcust.MFGQuoteNumber, projectcust.Cost, projectcust.Sell, projectcust.Comm, projectcust.Overage, projectcust.QuoteToDate, projectcust.OrdDate, projectcust. NotGot, projectcust.Comments, projectcust.SPANumber, projectcust.SpecCross, projectcust.LotUnit, projectcust.LPCost, projectcust.LPSell, projectcust.LPComm, projectcust.LampsIncl, projectcust.Terms, projectcust.FOB, projectcust.QuoteID, projectcust.BranchCode, projectcust.LeadTime, projectcust.LastChgDate, projectcust.LastChgBy, projectcust.Requested, projectcust.FileName,  quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob, quote.SpecCross as SpecCrossH, quote.SourceQuote FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID " '11-11-15 add quote.speccross '12-31-14 JTC 01-06-14 - this is the entire two lines commented out above for clarity 05-01-17vadd sourcequote

                Dim PullBusTypeND As String = "" '01-06-15 KEEP BLANK - FILL IF THEY WANT TO PULL TYPE FROM N&A, ADD TO SQL BELOW
                Dim PullBusTypeNDJoin As String = "" '01-06-15 KEEP BLANK - FILL IF THEY WANT TO PULL TYPE FROM N&A, ADD TO SQL BELOW
                '04-30-15 Dim CTypes As String = frmQuoteRpt.cboTypeCustomer.Text '01-06-15 MOVED UP, WAS INSIDE IF
                CTypes = frmQuoteRpt.cboTypeCustomer.Text '01-06-15 MOVED UP, WAS INSIDE IF
                If RealCustomerOnly = True Then '03-11-14
                    If frmQuoteRpt.cboTypeCustomer.Text.Trim.ToUpper <> "ALL" Then
                        '01-06-15 JH strSql = "SELECT NameDetail.BUsinessType, projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID INNER JOIN namedetail ON namedetail.code  = projectcust.ncode " '03-11-14 '01-06-15 JH  new out strSql = "SELECT NameDetail.BUsinessType,  IF(projectcust.NCODE = '', Left(projectcust.FIRMNAME,6), projectcust.ncode) as NCode,  projectcust.ProjectID, projectcust.QuoteCode, projectcust.FirmName, projectcust.ContactName, projectcust.SLSCode, projectcust.Got, projectcust.Typec, projectcust.MFGQuoteNumber, projectcust.Cost, projectcust.Sell, projectcust.Comm, projectcust.Overage, projectcust.QuoteToDate, projectcust.OrdDate, projectcust. NotGot, projectcust.Comments, projectcust.SPANumber, projectcust.SpecCross, projectcust.LotUnit, projectcust.LPCost, projectcust.LPSell, projectcust.LPComm, projectcust.LampsIncl, projectcust.Terms, projectcust.FOB, projectcust.QuoteID, projectcust.BranchCode, projectcust.LeadTime, projectcust.LastChgDate, projectcust.LastChgBy, projectcust.Requested, projectcust.FileName,  quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID  INNER JOIN namedetail ON namedetail.code  = projectcust.ncode " '01-06-14 
                        '01-06-15 MOVED UP Dim CTypes As String = frmQuoteRpt.cboTypeCustomer.Text
                        PullBusTypeND = " NameDetail.BUsinessType, " '01-06-15
                        PullBusTypeNDJoin = " INNER JOIN namedetail ON namedetail.code  = projectcust.ncode " '01-06-15
                        If frmQuoteRpt.cboTypeCustomer.Text.ToUpper.Trim <> "ALL" Then
                            If CTypes.Contains(",") Then
                                CTypes = " and ( NameDetail.BUsinessType = '" & CTypes.Replace(",", "' or NameDetail.BUsinessType = '") & "' ) "
                            Else
                                CTypes = " and ( NameDetail.BUsinessType = '" & CTypes & "' ) "
                            End If
                            PullBusTypeNDJoin += CTypes '01-06-15 WAS STRSQL +=
                        End If
                    End If
                End If

                'Debug.Print(frmQuoteRpt.txtPrimarySortSeq.Text)
                'SLS SORTS NEED TO JOIN TO QUTSLSSPLIT TABLE FOR SLS 1-4 SELECT CRITERIA
                If frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman/Customer" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman" Then '01-21-14 
                    '01-06-15 JH NOTES - REALIZATION - USE SLS FROM PROJECT CUST LIKE VB6 - STILL JOIN TO QUTSLSSPLIT FOR SELECT CRITERIA SLS 1-4 BOX
                    'SELECT PROJECTCUST.Ncode, Quote.TypeOfJob, Quote.EntryDate, Quote.QuoteID FROM QUOTE INNER JOIN PROJECTCUST ON QUOTE.QUOTEID = PROJECTCUST.QUOTEID  " '01-21-14 Quote.EntryDateWhere projectcust.Ncode = 'KEEN' and PROJECTCUST.TYPEC = 'M' and QUOTE.TYPEOFJOB = 'Q'  AND  QUOTE.ENTRYDATE >= '2009-03-01' AND QUOTE.ENTRYDATE <= '2013-03-31'"
                    '02-03-14 Out QS.SLSCode as SLSCode,
                    '01-06-15 THIS IS FILLED OUT ABOVE (205:) 
                    '01-06-15 strSql = "SELECT projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID "
                    strSql += " LEFT JOIN QUTSLSSPLIT QS ON Quote.QuoteID = QS.QuoteID " '02-22-13 NEEDS TO BE IN THE SQL FOR THE SELECT CRITERIA TAB'S SLS 1-4 BOX TO LIMIT SLS CODES COMING BACK
                End If

                If frmQuoteRpt.txtSlsSplit.Text.ToUpper <> "ALL" Or frmQuoteRpt.chkSlsFromHeader.CheckState = CheckState.Checked Then '02-03-14 add QS.SLSCode as SLSCode, the following
                    '02-22-13 JTC Add QS.SLSCode as SLSCode from Quote QUTSLSSPLIT Table All Sls's from table Realization 
                    '01-06-14 JH TAKE projectcust.* OUT AND USE IF STATEMENT FROM ABOVE FOR BLANK NCODES - strSql = "SELECT QS.SLSCode as SLSCode, projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID " '02-22-13
                    strSql = "SELECT QS.SLSCode as SLSCode, IF(projectcust.NCODE = '', UCASE(Left(projectcust.FIRMNAME,6)), UCASE(projectcust.ncode)) as NCode,  projectcust.ProjectID, projectcust.QuoteCode, projectcust.FirmName, projectcust.ContactName, projectcust.SLSCode, projectcust.Got, projectcust.Typec, projectcust.MFGQuoteNumber, projectcust.Cost, projectcust.Sell, projectcust.Comm, projectcust.Overage, projectcust.QuoteToDate, projectcust.OrdDate, projectcust. NotGot, projectcust.Comments, projectcust.SPANumber, projectcust.SpecCross, projectcust.LotUnit, projectcust.LPCost, projectcust.LPSell, projectcust.LPComm, projectcust.LampsIncl, projectcust.Terms, projectcust.FOB, projectcust.QuoteID, projectcust.BranchCode, projectcust.LeadTime, projectcust.LastChgDate, projectcust.LastChgBy, projectcust.Requested, projectcust.FileName,  quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob, quote.SpecCross as SpecCrossH FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID" '11-11-15 added quote.SpecCross'01-06-15
                    strSql += " LEFT JOIN QUTSLSSPLIT QS ON Quote.QuoteID = QS.QuoteID " '02-22-13 Sls 1-4 AND QS.slsnumber = 1 "  NEEDS TO BE IN THE SQL FOR THE SELECT CRITERIA TAB'S SLS 1-4 BOX TO LIMIT SLS CODES COMING BACK
                End If

                'QUOTE TO - REALIZATION - CUSTOMER - SLS - MAJOR ON SLS = YES, SELECT CRITERIA - PULL FROM N&A BUSINESS TYPE
                If PullBusTypeND <> "" Then '01-06-14 
                    strSql = strSql.Substring(6)
                    strSql = "SELECT " & PullBusTypeND & strSql & PullBusTypeNDJoin
                End If

            End If
            'Debug.Print(frmQuoteRpt.cboSortPrimarySeq.Text)
            If SESCO = True Or ExcelQuoteFU = True Then '04-28-14 JTC Public Bool frmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-22-15 JTC 02-22-12
                '04-17-12 jh - MOVE WHERE QS.SLS TO THE JOIN
                'strSql = "SELECT QS.SLSCode as SLS1, projectcust.QuoteCode, projectcust.NCode, projectcust.SLSCode, projectcust.Typec, projectcust.Sell, projectcust.Comments, projectcust.LeadTime, projectcust.FirmName, quote.Sell as Sell, quote.JobName, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.CSR, quote.TypeOfJob, quote.TypeOfJob as Architect, quote.TypeOfJob as Engineer, quote.TypeOfJob as Distributor, quote.TypeOfJob as Contractor "
                'strSql += "FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID "
                'strSql += " LEFT JOIN QUTSLSSPLIT QS ON Quote.QuoteID = QS.QuoteID LEFT join NameContact N on N.EMPCODE = QS.SLSCODE and N.CODE = '999999' and N.Category = 'SLSMAN' "
                'strSql += "where QS.slsnumber = 1 and Quote.TypeOfJob = 'Q' and Quote.JobName <> ''  and " '02-24-12" '02-23-12
                'Backup SESCO 8888888888888888888888888888888888888888888888888888888888
                If SESCO = True Then
                    strSql = "SELECT QS.SLSCode as SLS1, projectcust.QuoteCode, projectcust.NCode, projectcust.SLSCode, projectcust.Typec, projectcust.Sell, projectcust.Comments, projectcust.LeadTime, projectcust.FirmName, quote.Sell as Sell, quote.JobName, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.CSR, quote.TypeOfJob, quote.TypeOfJob as Architect, quote.TypeOfJob as Engineer, quote.TypeOfJob as Distributor, quote.TypeOfJob as Contractor, quote.quoteid, projectcust.QUOTETODATE "
                    strSql += " FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID  AND TYPEC = 'C'" '04-17-12 ADD TYPEC <> M
                    strSql += " LEFT JOIN QUTSLSSPLIT QS ON Quote.QuoteID = QS.QuoteID AND QS.slsnumber = 1 "
                    strSql += " LEFT join NameContact N on N.EMPCODE = QS.SLSCODE and N.CODE = '999999' and N.Category = 'SLSMAN' "
                    strSql += " where Quote.TypeOfJob = 'Q' and Quote.JobName <> ''  and " '04-17-12
                End If
                '8888888888888888888888888888888888888888888888888888888888888888 '05-05-15 JTC Chg quote.Sell as SellQ
                If ExcelQuoteFU = True Then '04-28-14 JTC Public Bool frmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-22-15 JTC 02-22-12
                    '04-30-15 No SLS1 strSql = "SELECT QS.SLSCode as SLS1, projectcust.QuoteCode, projectcust.NCode, projectcust.SLSCode, projectcust.Typec, projectcust.Sell, projectcust.Comments, projectcust.LeadTime, projectcust.FirmName, quote.Sell as Sell, quote.JobName, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.CSR, quote.TypeOfJob, quote.TypeOfJob as Architect, quote.TypeOfJob as Engineer, quote.TypeOfJob as Distributor, quote.TypeOfJob as Contractor, quote.quoteid, projectcust.QUOTETODATE , NameDetail.BUsinessType  " '04-27-15  NameDetail.BUsinessType, 
                    strSql = "SELECT projectcust.QuoteCode, projectcust.NCode, projectcust.SLSCode, projectcust.Typec, projectcust.Sell, projectcust.Comments, projectcust.LeadTime, projectcust.FirmName, projectcust.SLScode, quote.Sell as SellQ, quote.JobName, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.CSR, quote.TypeOfJob, quote.TypeOfJob as Architect, quote.TypeOfJob as Engineer, quote.TypeOfJob as Distributor, quote.TypeOfJob as Contractor, quote.quoteid, projectcust.QUOTETODATE , NameDetail.BUsinessType, quote.SpecCross as SpecCrossH " '11-11-15 added quote.SpecCross '04-27-15  NameDetail.BUsinessType, 
                    strSql += " FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID  AND TYPEC = 'C'" '04-17-12 ADD TYPEC <> M
                    'No strSql += " LEFT JOIN QUTSLSSPLIT QS ON Quote.QuoteID = QS.QuoteID AND QS.slsnumber = 1 "
                    strSql += " INNER JOIN namedetail ON namedetail.code  = projectcust.ncode " '04-27-15 
                    'No Use QT strSql += " LEFT join NameContact N on N.EMPCODE = QS.SLSCODE and N.CODE = '999999' and N.Category = 'SLSMAN' "
                    strSql += " where Quote.TypeOfJob = 'Q' and Quote.JobName <> ''  and " '04-17-12
                    ' frmQuoteRpt.txtStatus.Text = "ALL" '"OPEN,LOST"
                    If frmQuoteRpt.txtStatus.Text.Trim.ToUpper <> "ALL" Then
                        CTypes = frmQuoteRpt.txtStatus.Text
                        If frmQuoteRpt.txtStatus.Text.Contains(",") Then
                            CTypes = " ( Quote.Status = '" & CTypes.Replace(",", "' or Quote.Status = '") & "' ) "
                        Else
                            CTypes = " ( Quote.Status = '" & CTypes & "' ) "
                        End If
                        strSql += CTypes & " and "
                        SaveStrSQL2 = CTypes '04-30-15 JTC Quote.Status = Status frmQuoteRpt.txtStatus.Text
                    End If
                    '05-05-15 JTC Fix Zero QuoteTo Amount on Contractors BOM print If ExcelQuoteFU = True Ask to Select on Quote Header Amout and Print it if Quote to Amt is Zero
                    Resp = MsgBox("Yes = Use Quote Header Sell Amoumt on Report Selection or " & vbCrLf & "No = Use Realization QuoteTo Sell Amount on Report", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Realization Amount") '04-17-15 JTC 
                    RealQuoteToAmtON = False '07-15-14  Public RealQuoteToAmtON As Boolean = 0 
                    If Resp = vbYes Then

                        '04-30-15 " projectcust.Sell >= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' and " '04-30-15 projectcust.Sell <= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' ) " '04-30-15 JTC
                        SaveStrSQL = strSql '04-30-15 " projectcust.Sell >= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' and " '04-30-15 projectcust.Sell <= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' ) " '04-30-15 JTC
                        If Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "0" Then
                            RealQuoteToAmtON = True '05-05-15 JTC
                            strSql += " Quote.Sell >= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' and " '04-30-15 projectcust.Sell <= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' ) " '04-30-15 JTC
                        End If
                    Else
                        SaveStrSQL = strSql '04-30-15 " projectcust.Sell >= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' and " '04-30-15 projectcust.Sell <= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' ) " '04-30-15 JTC
                        If Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "0" Then
                            strSql += " projectcust.Sell >= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' and " '04-30-15 projectcust.Sell <= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "' ) " '04-30-15 JTC
                        End If
                    End If
                End If
                    '"Project", "RetrCode", "Architect", "Engineer", "SLSCode", "SLSQ", "Distributor", "Comments", "Contractor", "LeadTime", "Sell", "BidDate"}

                GoTo GotSql '02-27-12 
                    'NG Test Method ***************** Using Stored Procedure GetSpecifiersByType'02-27-12
                    'strSql = "SELECT QS.SLSCode as SLS1, projectcust.QuoteCode, projectcust.NCode, projectcust.SLSCode, projectcust.Typec, projectcust.Sell, projectcust.Comments, projectcust.LeadTime, projectcust.FirmName, quote.Sell as Sell, quote.JobName, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.CSR, quote.TypeOfJob "
                    ''Latest Version Timed Out
                    strSql = "SELECT QS.SLSCode as SLS1, projectcust.QuoteCode, projectcust.NCode, projectcust.SLSCode, projectcust.Typec, projectcust.Sell, projectcust.Comments, projectcust.LeadTime, projectcust.FirmName, quote.Sell as Sell, quote.JobName, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.CSR, quote.TypeOfJob,GetSpecifiersByType(quote.quoteid,'A') as Architect, GetSpecifiersByType(quote.quoteid,'E') as Engineer, GetSpecifiersByType(quote.quoteid,'S') as Specifier, GetSpecifiersByType(quote.quoteid,'T') as Contractor FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID INNER JOIN QUTSLSSPLIT QS ON Quote.QuoteID = QS.QuoteID INNER join NameContact N on N.EMPCODE = QS.SLSCODE and N.CODE = '999999' and N.Category = 'SLSMAN' where QS.slsnumber = 1 and Quote.TypeOfJob = 'Q' and Quote.JobName <> '' and "
                    'strSql = "SELECT QS.SLSCode as SLS1, projectcust.QuoteCode, projectcust.NCode, projectcust.SLSCode, projectcust.Typec, projectcust.Sell, projectcust.Comments, projectcust.LeadTime, projectcust.FirmName, quote.Sell as Sell, quote.JobName, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.CSR, quote.TypeOfJob,GetSpecifiersByType(quote.quoteid,'A') as Architect, GetSpecifiersByType(quote.quoteid,'E') as Engineer, GetSpecifiersByType(quote.quoteid,'S') as Specifier, GetSpecifiersByType(quote.quoteid,'T') as Contractor FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID INNER JOIN QUTSLSSPLIT QS ON Quote.QuoteID = QS.QuoteID INNER join NameContact N on N.EMPCODE = QS.SLSCODE and N.CODE = '999999' and N.Category = 'SLSMAN' where QS.slsnumber = 1 and Quote.TypeOfJob = 'Q' and Quote.JobName <> ''  "
GotSql:
                    GoTo GetDate
                    'order by quote.JobName, projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC
                    'strSql += "SELECT projectcust.*, quote.*,  (Select projectcust.FirmName from projectcust where projectcust.Typec = 'T' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Contractor,  (Select projectcust.FirmName from projectcust where projectcust.Typec = 'E' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Engineer,  (Select projectcust.FirmName from projectcust where projectcust.Typec = 'S' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Architect " '02-22-12
                    'strSql += " FROM Quote  INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID where Quote.TypeOfJob = 'Q'  and  projectcust.TypeC = 'C'"
                End If

                '05-31-11 ok strSql = "SELECT projectcust.*, quote.* FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID " '09-14-10 INNER JOIN rojectcust ON Quote.QuoteID = projectcust.QuoteID INNER JOIN quote ON projectcust.QuoteID = quote.QuoteID "
300:            '05-31-11-2 OKstrSql = "SELECT projectcust.*, quote.* FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID " '05-31-11 
                Dim JT As String = frmQuoteRpt.cboTypeofJob.Text '12-10-09 

                If SortCode <> "" Then
                    strSql += "where " & SortSeq & " >= '" & SortCode & "' "
                    If JT <> "A" Then strSql += " and  Quote.TypeOfJob = '" & JT & "' " : WhereS = "Where" '05-17-11' =ALL 12-10-09 No Select on Q.TypeOfJob JT = " * ""
                Else
                    If JT <> "A" Then strSql += "where Quote.TypeOfJob = '" & JT & "' " : WhereS = "Where" '05-17-11' =ALL 12-10-09 No Select on Q.TypeOfJob JT = " * ""
                End If
                If WhereS = "Where" Then WhereS = " and " Else WhereS = "Where" '05-17-11'05-17-11
                'Debug.Print(strsql)VB6.Format(BidDate, "yyyy-MM-dd")
GetDate:        Dim SaveDateSql As String = "" '03-22-13
                Dim sEndDate As String = VB6.Format(frmQuoteRpt.DTPicker1EndEntry.Value, "yyyy-MM-dd") '02-03-12
                Dim sStartDate As String = VB6.Format(frmQuoteRpt.DTPickerStartEntry.Value, "yyyy-MM-dd") '02-03-12 - not /
                '02-04-14 JTC Real to MFG add projectcust.NCODE = 'DAYB' and if one Mfg from frmQuoteRpt.txtQutRealCode.Text.ToUpper in FillQutRealLUDataSet
                ' frmQuoteRpt.txtQutRealCode
                '02-05-14 OUT'If Trim(frmQuoteRpt.txtQutRealCode.Text) <> "" And Trim(frmQuoteRpt.txtQutRealCode.Text.ToUpper) <> "ALL" And frmQuoteRpt.txtQutRealCode.Text.Length < 7 Then '02-04-14
                '    WhereS = WhereS & " projectcust.NCODE =  '" & frmQuoteRpt.txtQutRealCode.Text.ToUpper & "' and "
                'End If
                BC = "" '02-05-14 JTC Added Get Many Mfg or Cust JONES,HERRY Realization FillQutRealLUDataSet
                If frmQuoteRpt.txtQutRealCode.Text.Trim <> "" Then '02-05-13
                    If frmQuoteRpt.txtQutRealCode.Text.Contains(",") = True Then
                        BC = " ( projectcust.NCODE = '" & frmQuoteRpt.txtQutRealCode.Text.Replace(",", "' or projectcust.NCODE = '") & "' ) and "
                    Else  'and ( Q.Status = 'OPEN'
                        BC = " ( projectcust.NCODE = '" & frmQuoteRpt.txtQutRealCode.Text & "'" & " ) and "
                    End If '
                End If
                WhereS = WhereS & BC '02-05-14 & " and " '02-05-14
                strSql = strSql & WhereS & " Quote.EntryDate >= '" & sStartDate & "' and Quote.EntryDate <= '" & sEndDate & "' " '05-17-11 
                SaveDateSql = WhereS & " Quote.EntryDate >= '" & sStartDate & "' and Quote.EntryDate <= '" & sEndDate & "' " '03-22-13
                If frmQuoteRpt.ChkCheckBidDates.CheckState = CheckState.Checked Then '02-03-12
                    Dim sEndBidDate As String = VB6.Format(frmQuoteRpt.DTPicker1EndBid.Value, "yyyy-MM-dd") ''02-03-12 - not /
                    Dim sStartBidDate As String = VB6.Format(frmQuoteRpt.DTPicker1StartBid.Value, "yyyy-MM-dd") '05-31-11
                    '06-05-12 strSql += " and (Quote.BidDate >= '" & sStartBidDate & "' and Quote.BidDate <= '" & sEndBidDate & "' or  quote.biddate is null) " '02-03-12
                    If frmQuoteRpt.chkBlankBidDates.CheckState = CheckState.Checked Then '06-05-12 Sesco = True 
                        strSql += " and (Quote.BidDate >= '" & sStartBidDate & "' and Quote.BidDate <= '" & sEndBidDate & "' or  quote.biddate is null) " '02-03-12 
                        SaveDateSql += " and (Quote.BidDate >= '" & sStartBidDate & "' and Quote.BidDate <= '" & sEndBidDate & "' or  quote.biddate is null) "
                    Else
                        strSql += " and (Quote.BidDate >= '" & sStartBidDate & "' and Quote.BidDate <= '" & sEndBidDate & "' ) " '06-05-12 "
                        SaveDateSql += " and (Quote.BidDate >= '" & sStartBidDate & "' and Quote.BidDate <= '" & sEndBidDate & "' ) "
                    End If
                End If
            'Debug.Print(frmQuoteRpt.txtSecondarySort.Text)
            '09-17-15 JTC Fic Q.Status to Quote.Status in FillQutRealLUDataSet
            SaveDateSql += "  and Quote.Status <> '" & "NOREPT" & "' " '06-30-15 JTC Add  and Q.Status <> '" & "NOREPT" & "'  to Realization Reports 

            If frmQuoteRpt.txtJobNameSS.Text <> "" And frmQuoteRpt.txtJobNameSS.Text.Trim.ToUpper <> "ALL" Then '01-26-16 JH
                Dim SearchString As String = " AND (Quote.JobName LIKE" & " '%" & SafeSQL(frmQuoteRpt.txtJobNameSS.Text) & "%') "
                strSql += SearchString
            End If

                'If frmQuoteRpt.cboSortSecondarySeq.Text = "Customer Type" Then '03-11-14
                '    strSql = "SELECT NameDetail.BUsinessType, projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID INNER JOIN namedetail ON namedetail.code  = projectcust.ncode " '03-11-14
                '    Dim CTypes As String = frmQuoteRpt.cboTypeCustomer.Text
                '    If frmQuoteRpt.cboTypeCustomer.Text.ToUpper.Trim <> "ALL" Then
                '        If CTypes.Contains(",") Then
                '            CTypes = " and ( NameDetail.BUsinessType = '" & CTypes.Replace(",", "' or NameDetail.BUsinessType = '") & "' ) "
                '        Else
                '            CTypes = " and ( NameDetail.BUsinessType = '" & CTypes & "' ) "
                '        End If
                '        strSql += CTypes
                '    End If
                '    Select Case frmQuoteRpt.txtSortThirdSeq.Text
                '        Case "Business Type"
                '            OrderBy = " namedetail.businesstype, projectcust.NCode " 'Name Code
                '        Case "Customer Code"
                '            OrderBy = " projectcust.NCode "
                '        Case "Quote Code"
                '            OrderBy = " projectcust.QuoteCode "
                '        Case "Job Name"
                '            OrderBy = " Quote.JobName "
                '        Case "SLSQ"
                '            OrderBy = " Quote.SLSQ "
                '        Case "Entry Date"
                '            OrderBy = " Quote.EntryDate "
                '    End Select
                'End If
                If SESCO = True Or ExcelQuoteFU = True Then GoTo 500 '04-28-14 JTC Public Bool  Or frmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then GoTo 500 '04-22-15 JTC 02-22-12 Skip this
400:            Dim PcType As String = "*" '05-17-11
                'Becarefull PCType is Filled out below also
                If RealCustomer Then PcType = "C" '02-04-12  Cust
                If RealManufacturer Then PcType = "M" '02-04-12   Mfg
                If RealQuoteTOOther Then PcType = "O" '02-04-12
                If RealArchitect Then PcType = "A" '02-04-12
                If RealEngineer Then PcType = "E" '02-04-12
                If RealLtgDesigner Then PcType = "L" '02-04-12
                If RealSpecifier Then PcType = "S" '02-04-12
                If RealContractor Then PcType = "T" '02-04-12
                If RealSLSCustomer Then PcType = "C" '01-20-12 "Salesman/Customer"
                If RealOther Then PcType = "X" '01-31-12
                If RealTgLookupExcel = True Then '11-27-13
                    SaveStrSQL2 = strSql
                End If
                If PcType <> "*" Then strSql += " and (projectcust.TypeC = '" & PcType & "' " '08-21-11 ASdded (
                '01-20-12 If PcType = "C" Or PcType = "M" Then '08-20-11 Don't Due on Cust / Mfg           '01-20-12 Else
                If PcType = "*" Then GoTo 500 '09-21-11
                If frmQuoteRpt.ChkSpecifiers.Checked = True Then RealALL = True '10-13-14 JTC Fix Check box
                WhereSqlTypec = strSql '04-16-15 JTC WhereSqlTypec = Replace(strSql, WhereSqlTypec, "") '04-16-15 Jtc Save this Where Used on QuotesToOneMFG
                If RealCustomer = True Then strSql += " or projectcust.TypeC = 'C' " 'Cust08-20-11  '01-29-13 No Or RealALL = True
                If RealManufacturer = True Then strSql += " or projectcust.TypeC = 'M' " 'MFG
                If RealQuoteTOOther = True Then strSql += " or projectcust.TypeC = 'O' " 'Other
                If RealSLSCustomer = True Then strSql += " or projectcust.TypeC = 'C' " 'SLS/Customer
                If RealArchitect = True Or RealALL = True Then strSql += " or projectcust.TypeC = 'A' " 'Arch
                If RealEngineer = True Or RealALL = True Then strSql += " or projectcust.TypeC = 'E' " 'Eng
                If RealLtgDesigner = True Or RealALL = True Then strSql += " or projectcust.TypeC = 'L' " 'LtgDesigner
                If RealSpecifier = True Or RealALL = True Then strSql += " or projectcust.TypeC = 'S' " ' Specifier
                If RealContractor = True Or RealALL = True Then strSql += " or projectcust.TypeC = 'T' " 'Contractor
                If RealOther = True Or RealALL = True Then strSql += " or projectcust.TypeC = 'X' " 'Other 01-31-12
                WhereSqlTypec = Replace(strSql, WhereSqlTypec, "") '04-16-15 Jtc Save this Where Used on QuotesToOneMFG ' WhereSqlTypec = Replace(WhereSqlTypec, " or", "") '04-16-15 Jtc Save this Where Used on QuotesToOneMFG
            If WhereSqlTypec.StartsWith(" or") Then WhereSqlTypec = Mid(WhereSqlTypec, 4) '
480:            If Right(strSql, 2) <> ") " And InStr(strSql, "(") Then strSql += ") " '09-21-11 
                'WhereSqlTypec.Replace(WhereSqlTypec, strSql, "") '04-16-15 Jtc
500:            '02-11-12 Moved dpwn strSql += " order by projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC "
                'Debug.Print(frmQuoteRpt.ChkSpecifiers.Text)                                     chkDetailTotal  Unchecked = Detail
                '02-01-14 Test frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" and frmQuoteRpt.txtPrimarySortSeq.Text = "SalesMan" and frmQuoteRpt.cboSortSecondarySeq.Text = "Descending Dollar" then 
                '02-01-14 Test frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" and frmQuoteRpt.txtPrimarySortSeq.Text = "SalesMan" and frmQuoteRpt.cboSortSecondarySeq.Text = "Descending Dollar" then 
                If frmQuoteRpt.txtPrimarySortSeq.Text = "Name Code" And VQRT2.SubSeq = VQRT2.SubSortType.SubSDescend Then '02-01-14 JTC Real
                    '02-01-14 If frmQuoteRpt.ChkSpecifiers.Text = "Sort Report by Descending Dollar" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Unchecked Then '07-27-12 Unchhecked
                    'Fix Order By SortSeq = "projectcust.Sell"''02-11-12 
                    Resp = MsgBox("Yes = Sort by Name Code / Descending Sales Dollars or " & vbCrLf & "No = Just Descending Sales Dollars", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Realization Sort") '02-11-12
                    If Resp = vbYes Then  'Sort by Name Code / Descending Sales Dollars
                        If RealCustomer = True Or RealManufacturer = True Or RealSLSCustomer = True Then '01-28-13 use PC.Sell else Quote.Sell
                            SortSeq = "projectcust.NCode, projectcust.Sell" '02-11-12
                            OrderBy = " order by projectcust.NCode, projectcust.Sell DESC, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC " '02-11-12 
                        Else
                            SortSeq = "projectcust.NCode, Quote.Sell" '01-28-13
                            OrderBy = " order by projectcust.NCode, Quote.Sell DESC, projectcust.QuoteCode, projectcust.NotGot DESC " '0
                        End If
                        frmQuoteRpt.txtPrimarySortSeq.Text = "Name Code / Descending Sales Dollars" ' Name Code" '02-11-12
                        frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" 'frmQuoteRpt.txtPrimarySortSeq.Text = "Descending Sales Dollars" ' Name Code" '02-11-12
                        frmQuoteRpt.txtSortSeq.Text = frmQuoteRpt.pnlTypeOfRpt.Text & "  Sort By = " & frmQuoteRpt.txtPrimarySortSeq.Text '02-11-12 & " / " & txtSecondarySort.Text '02-24-09
                    Else 'Just Descending Sales Dollars
                        If RealCustomer = True Or RealManufacturer = True Or RealSLSCustomer = True Then '01-28-13 use PC.Sell else Quote.Sell
                            SortSeq = "projectcust.Sell" '02-11-12
                            OrderBy = " order by projectcust.Sell DESC, Ncode " : If frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked Then frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Unchecked '09-23-15 JTC 
                        Else
                            SortSeq = "Quotecust.Sell" '02-11-12
                            OrderBy = " order by Quote.Sell DESC, Ncode " '01-28-13
                        End If
                        frmQuoteRpt.txtPrimarySortSeq.Text = "Descending Sales Dollars" ' Name Code" '02-11-12 SortSeq = "projectcust.Sell"
                        frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" 'frmQuoteRpt.txtPrimarySortSeq.Text = "Descending Sales Dollars" ' Name Code" '02-11-12
                        frmQuoteRpt.txtSortSeq.Text = frmQuoteRpt.pnlTypeOfRpt.Text & "  Sort By = " & frmQuoteRpt.txtPrimarySortSeq.Text '02-11-12 & " / " & txtSecondarySort.Text '02-24-09
                    End If
                Else 'Not Descending  If frmQuoteRpt.ChkSpecifiers.Text = "Sort Report by Descending Dollar" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Uncheck
                    If SESCO = True Then '04-28-15 JTC Or frmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-22-15 JTC 
                        'strSql += "SELECT projectcust.*, quote.*,  (Select projectcust.FirmName from projectcust where projectcust.Typec = 'T' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Contractor,  (Select projectcust.FirmName from projectcust where projectcust.Typec = 'E' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Engineer,  (Select projectcust.FirmName from projectcust where projectcust.Typec = 'S' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Architect " '02-22-12
                        'strSql += " FROM Quote  INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID where Quote.TypeOfJob = 'Q'  and  projectcust.TypeC = 'C'"
                        OrderBy = " order by Quote.JobName, projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.QUOTETODATE DESC " '02-22-12 04-17-12 TypeC
                    ElseIf ExcelQuoteFU = True Then '04-28-14 JTC Public BoolfrmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-22-15 JTC 
                        'strSql += "SELECT projectcust.*, quote.*,  (Select projectcust.FirmName from projectcust where projectcust.Typec = 'T' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Contractor,  (Select projectcust.FirmName from projectcust where projectcust.Typec = 'E' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Engineer,  (Select projectcust.FirmName from projectcust where projectcust.Typec = 'S' and Quote.QuoteID = projectcust.QuoteID Limit 1) as Architect " '02-22-12
                        'strSql += " FROM Quote  INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID where Quote.TypeOfJob = 'Q'  and  projectcust.TypeC = 'C'"
                        OrderBy = " order by Quote.JobName, projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.QUOTETODATE DESC " '02-22-12 04-17-12 TypeC
                        If frmQuoteRpt.txtSecondarySort.Text = "Bid Date" Then '04-28-15
                            OrderBy = " order by quote.BidDate DESC, Quote.JobName DESC, NameDetail.BUsinessType, projectcust.Ncode DESC  "
                            GoTo SkipOrderBy
                        End If
                    Else '07-26-12 OrderBy 'Debug.Print(frmQuoteRpt.txtPrimarySortSeq.Text)
                        'Debug.Print(frmQuoteRpt.txtSecondarySort.Text)
                        '05-16-13 JTC Added Realization When sub Sort is salesman they can change tobe Salesman Major Sequence
                        If frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman/Customer" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman" Then '05-16-13 Added frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman 
                            '10-30-12 OrderBy = " order by projectcust.SLSCode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC " '08-08-12 
                            '10-30-12 OrderBy = " projectcust.SLSCode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC " '08-08-12 
                            '02-03-14 JTC If QS.SLSCode as SLSCode, 
                            'Debug.Print(frmQuoteRpt.txtSecondarySort.Text)
                            If frmQuoteRpt.txtSecondarySort.Text = "Job Name" Then '01-21-14 JTC Realization by Salesmam by Jobname or NCode ALS
                                OrderBy = " projectcust.SLSCode, Jobname, projectcust.NotGot DESC, projectcust.LastChgDate DESC " '01-21-14
                                'If InStr(strSql, "QS.SLSCode as SLSCode") Then OrderBy = " QS.SLSCode, Jobname, projectcust.NotGot DESC, projectcust.LastChgDate DESC " '02-03-14
                            Else '     Sls/CUst
                                OrderBy = " projectcust.SLSCode, projectcust.NCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC " '10-30-12 Chg QuoteCode to Ncode
                                If frmQuoteRpt.chkSlsFromHeader.CheckState = CheckState.Checked Then '02-03-14
                                    OrderBy = " SLSQ, projectcust.Ncode, JobName, projectcust.NCode, projectcust.NotGot DESC " '02-03-14
                                ElseIf frmQuoteRpt.txtSlsSplit.Text.ToUpper <> "ALL" Then '02-03-14 add QS.SLSCode as SLSCode, the following
                                    OrderBy = " QS.SLSCode as SLSCode, projectcust.NCode, projectcust.NotGot DESC "
                                End If
                                'If InStr(strSql, "QS.SLSCode as SLSCode") Then OrderBy = " QS.SLSCode, projectcust.NCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC " '02-03-14
                            End If
                    ElseIf frmQuoteRpt.txtSecondarySort.Text = "Spread Sheet by Month" Or frmQuoteRpt.txtSecondarySort.Text = "Spread Sheet by Year" Then '06-22-15 JTC05-20-13

                        OrderBy = " projectcust.Ncode, projectcust.QuoteToDate " '05-20-13
                        'OrderBy = " Ncode, projectcust.QuoteToDate " '06-24-15 
                        'SELECT IF(projectcust.NCODE = '', left(projectcust.FIRMNAME,6), projectcust.ncode) as NCode,  proj
                        '12-31-14 JTC Chg OrderBy
                        If strSql.StartsWith("Select IF(projectcust.NCODE") Or strSql.StartsWith("SELECT IF(projectcust.NCODE") Then '06-25-15 12-31-14 JTC   OrderBy = " NCode
                            strSql += " and NCODE <> '' " '06-25-15 No Blanks
                            OrderBy = Replace(OrderBy.ToUpper, " PROJECTCUST.NCODE", " NCode") '
                        End If '12-31-14 JTC   OrderBy = " NCode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC "
                        '12-31-14  OrderBy = " NCode, projectcust.QuoteToDate " '05-20-13
                        ElseIf frmQuoteRpt.txtSecondarySort.Text = "Status" Then '01-22-14 JTC Add REAL txtSecondarySort.Text = "Status" and txtSecondarySort.Text = "Bid Date"
                            OrderBy = " projectcust.Ncode, Quote.Status "
                        ElseIf frmQuoteRpt.txtSecondarySort.Text = "Bid Date" Then '01-22-14
                            OrderBy = " projectcust.Ncode, quote.BidDate "
                        ElseIf VQRT2.RepType = VQRT2.RptMajorType.RptProj And VQRT2.SubSeq = VQRT2.SubSortType.SubSProj And RealCustomerOnly = True Then '03-11-14
                            Resp = MessageBox.Show("Do you want the Secondary Sort by Specifier SLS?", "Sort", MessageBoxButtons.YesNo)
                            If Resp = MsgBoxResult.Yes Then
                                OrderBy = " Quote.JobName , projectcust.SLScode, projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC "
                            Else
                                OrderBy = " Quote.JobName , projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC "
                            End If

                        Else
                            ''10-30-12 OrderBy = " order by projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC "
                            OrderBy = " projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC "
                        '12-31-14 JTC Chg OrderBy
                        '06-25-15 JTC 
                        If strSql.StartsWith("Select IF(projectcust.NCODE") Or strSql.StartsWith("SELECT IF(projectcust.NCODE") Then '06-25-15 12-31-14 JTC   OrderBy = " NCode
                            strSql += " and NCODE <> '' " '06-25-15 No Blanks
                            OrderBy = Replace(OrderBy.ToUpper, " PROJECTCUST.NCODE", " NCode") '
                        End If
                        '06-25-15 If strSql.StartsWith("Select IF(projectcust.NCODE") Then OrderBy = " NCode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC "
                        '10-13-14 JTC By projectcust.QuoteCode, first VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif VQRT2.RepType = VQRT2.RptMajorType.RptProj
                        '10-13-14 JTCPublic RealWithOneMfgCustSortJobName As Boolean = False '10-13-14 JTC
                        If RealWithOneMfgCustSortJobName = True Then '10-13-14 VQRT2.RepType = VQRT2.RptMajorType.RptProj And VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif Then
                            If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then '10-13-14 JTC By projectcust.QuoteCode, first
                                OrderBy = " Quote.JobName, projectcust.Ncode, projectcust.NotGot DESC, projectcust.LastChgDate DESC "
                            End If
                        End If
                        End If
                        If BranchReporting = True Then '10-30-12
                            If OrderBy <> "" Then OrderBy = " order by " & " projectcust.BranchCode, " & OrderBy '11-03-13 JTC Not UpperCase 10-30-12
                        Else
                            If OrderBy <> "" Then OrderBy = " order by " & OrderBy ''11-03-13 JTC Not UpperCase 10-13-14 JTC added space
                        End If
                        ' OrderBy = Replace(OrderBy, " order by ", "") '02-03-14
                    End If
            End If                        'chkDetailTotal  Unchecked = Detail

            If strSql.StartsWith("Select IF(projectcust.NCODE") Or strSql.StartsWith("SELECT IF(projectcust.NCODE") Then '06-25-15 Then '12-31-14 JTC  Chg OrderBy = " NCode
                OrderBy = Replace(OrderBy.ToUpper, " PROJECTCUST.NCODE", " NCode") '
                strSql += " and NCODE <> '' " '06-25-15 No Blanks
            End If '12-31-14 JTC   OrderBy = " NCode 
            '02-01-14 '02-01-14 Test 
            'Debug.Print(frmQuoteRpt.txtPrimarySortSeq.Text) ' = "SalesMan" And 
            'Debug.Print(frmQuoteRpt.cboSortSecondarySeq.Text) ' = "Descending Dollar" ThenStop
            If frmQuoteRpt.txtPrimarySortSeq.Text = "Name Code" And VQRT2.SubSeq = VQRT2.SubSortType.SubSDescend Then '02-01-14  = "SalesMan" And frmQuoteRpt.cboSortSecondarySeq.Text = "Descending Dollar" ThenStop
                '02-01-14 If frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And frmQuoteRpt.ChkSpecifiers.Text = "Sort Report by Descending Dollar" Then '07-26-12 
                frmQuoteRpt.txtPrimarySortSeq.Text = "Name Code / Descending Sales Dollars" ' 07-27-12
                frmQuoteRpt.pnlTypeOfRpt.Text = "Realization"
                frmQuoteRpt.txtSortSeq.Text = frmQuoteRpt.pnlTypeOfRpt.Text & "  Sort By = " & frmQuoteRpt.txtPrimarySortSeq.Text '02-11-12 & " / " & txtSecondarySort.Text '02-24-09
                'SELECT projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID where Quote.TypeOfJob = 'Q'  and  Quote.EntryDate >= '2008-07-01' and Quote.EntryDate <= '2012-07-31'  and (projectcust.TypeC = 'X'  or projectcust.TypeC = 'A'  or projectcust.TypeC = 'E'  or projectcust.TypeC = 'L'  or projectcust.TypeC = 'S'  or projectcust.TypeC = 'T'  or projectcust.TypeC = 'X' )  order by projectcust.Sell DESC, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC '
                '07-27-12 Maybe not ***Group BY Eliminates Blanks !!!!!! add Sql IF(FirmName = '', NCode, Firmname),
                strSql = Replace(strSql, "quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ,", "Sum(quote.Sell) as SellQ, Sum(quote.Cost) as CostQ, Sum(quote.Comm) as CommQ, IF(FirmName = '', NCode, Firmname), ") '07-27-12  
                'If OrderBy = "ORDER BY SLSBRANCH" Then OrderBy = "GROUP BY SLSBRANCH, FIRMNAME ORDER BY SLSBRANCH, ORDERSALESAMT " '07-25-12
                OrderBy = " and quote.Status <> 'NOREPT' GROUP BY FirmName order by SellQ DESC " '11-04-14 JTC Lower Case 07-27-12 and quote.Status <> 'NOREPT'  
            End If
            If VQRT2.RepType = VQRT2.RptMajorType.RptProj And VQRT2.SubSeq = VQRT2.SubSortType.SubSProj And RealCustomerOnly = True Then GoTo SkipOrderby '10-13-14 JTC
            If frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman/Customer" Or frmQuoteRpt.txtSecondarySort.Text = "Job Name" Then GoTo SkipOrderby '01-21-14 08-08-12 
            'Major NCode Sequence'03-20-13 JTC Fix Salesman Report chkSlsFromHeader SQL Order By
            If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And frmQuoteRpt.chkSlsFromHeader.CheckState = CheckState.Checked Then '03-20-13 '03-08-13 JTC Add QS.SLSCode as SLS1 from Quote QUTSLSSPLIT Table Position 1
                'OrderBy = "SLSQ" 'Salesman
                If VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Then OrderBy = " order by projectcust.Ncode, SLSQ, JobName  " '03-20-13
            Else '03-20-13 
                'OrderBy = "PC.SLSCode" 'Salesman
                If VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Then OrderBy = " order by projectcust.Ncode, projectcust.SLSCode, JobName  " '03-20-13
                If VQRT2.RepType = VQRT2.RptMajorType.RptProj And VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif Then OrderBy = " order by JobName, projectcust.Ncode" '11-04-14 JTC JobName, NCode
            End If
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Then OrderBy = " order by projectcust.Ncode, Status " '07-27-12 
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSBidDate Then OrderBy = " order by projectcust.Ncode, BidDate " '07-27-12 
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSProj Then OrderBy = " order by projectcust.Ncode, JobName " '07-27-12 
SkipOrderBy:  '10-13-14 JTC
            If BranchReporting = True Then '10-30-12
                OrderBy = Replace(OrderBy, " order by ", "")
                If OrderBy <> "" Then OrderBy = " order by " & " projectcust.BranchCode, " & OrderBy '10-30-12
            End If
            '10-13-14 JTC SkipOrderby() '08-08-12
            '10-15-13 Added Below
            'If My.Computer.FileSystem.FileExists(UserSysDir & "VADMINNET.INI") = True Then '08-30-13
            'If SecurityLevel = "BRANCH" Or SecurityLevel = "REGIONAL" Then
            If My.Computer.FileSystem.FileExists(UserSysDir & "VADMINNET.INI") = True Or SecurityLevel = "" Then '10-16-13 JTC Added No Security
                If SecurityLevel = "BRANCH" Or SecurityLevel = "REGIONAL" Or SecurityBrancheCodes.Trim.ToUpper <> "ALL" Then '10-16-13 JTC Added No Security Or SecurityBrancheCodes.Trim.ToUpper <> "ALL"
                    BC = "" '02-05-14 JTC 
                    If SecurityBrancheCodes.Trim.ToUpper <> "ALL" Then
                        If SecurityBrancheCodes.Contains(",") = True Then
                            BC = " and ( quote.BranchCode = '" & SecurityBrancheCodes.Replace(",", "' or quote.BranchCode = '") & "' or quote.BranchCode = ''  )"
                        Else
                            BC = " and (quote.BranchCode = '" & SecurityBrancheCodes & "'" & " or quote.BranchCode = '')"
                        End If
                    End If
                    strSql += BC
                End If
            End If
            strSql += OrderBy '07-26-12 
            ''**********************************************************************************************************************************************
            'One Mfg all Customers/Specifiers   One Customer/Specifier all Mfgs
            If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then  '03-22-13 
                'SELECT * FROM saw8.projectcust p order by Typec ;
                'DROP TABLE IF EXISTS TMPREPORTS ;
                'CREATE TEMPORARY TABLE TMPREPORTS AS SELECT PROJECTCUST.Ncode, PROJECTCUST.Ncode as TMPNCode, PROJECTCUST.Typec as TMPTypec, projectcust.Sell as TMPSellQ, projectcust.Cost as TMPCostQ, projectcust.Comm as TMPCommQ, Quote.TypeOfJob, Quote.EntryDate, Quote.QuoteID, Quote.JobName FROM QUOTE INNER JOIN PROJECTCUST ON QUOTE.QUOTEID = PROJECTCUST.QUOTEID   Where PROJECTCUST.Ncode = 'KEEN' and Quote.TypeOfJob = 'Q'  and  Quote.EntryDate >= '2001-04-01' and Quote.EntryDate <= '2015-04-30'   order by JobName, projectcust.Ncode ;
                'SELECT * FROM TMPREPORTS ;
                'DROP TABLE IF EXISTS TMPREPORTS2 ;
                'CREATE TEMPORARY TABLE TMPREPORTS2 AS SELECT projectcust.*, TR.TMPNCode, TR.TMPTypec, TR.TMPSellQ, TR.TMPCostQ, TR.TMPCommQ, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State,  quote.CSR,  quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID inner join TMPREPORTS TR ON TR.QuoteID = QUOTE.QUOTEID  where  projectcust.TypeC = 'A'  projectcust.TypeC = 'E'  projectcust.TypeC = 'L'  and Quote.TypeOfJob = 'Q'   and  Quote.EntryDate >= '2001-04-01' and Quote.EntryDate <= '2015-04-30'  order by JobName, projectcust.Ncode
                '11-27-13 Dim SaveStrSQL As String = strSql '11-27-13 Dim SaveStrSQL2 As String = strSql
                '12-31-14 JTC If OneMfg or Cust Don't put FirmName in NCode
                '04-16-15 Change to SELECT to Upper "SELECT IF(projectcust.NCODE = '', Left(projectcust.FIRMNAME,6), projectcust.ncode) as NCode,
                If strSql.StartsWith("SELECT IF(projectcust.NCODE = '', UCASE(Left(projectcust.FIRMNAME,6)), UCASE(projectcust.ncode)) as NCode,") Then '04-16-15 JTC Select Upper12-31-14 JTC   OrderBy = " NCode
                    strSql = Replace(strSql, "IF(projectcust.NCODE = '', UCASE(Left(projectcust.FIRMNAME,6)), UCASE(projectcust.ncode)) as NCode,", " projectcust.NCode, ") '
                End If '12-31-14 JTC   OrderBy = " NCode, projectcust.QuoteCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC "
                '04-17-15 JTC Fix Message Quote Header vs Realization QuoteTo Sell Amount on Report Can't Show Mfg amount if to one Mfg
                '04-17-15 JTC Wrong don't show MFG Resp = MsgBox("Select Yes to show total Quote Dollars Or " & vbCrLf & "No to use the dollars just for " & RealWithOneMfgCustCode.Trim, MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton1, "Realization Dollars") '07-15-14
                Resp = MsgBox("Yes = Use Quote Header Sell Amoumt on Reports or " & vbCrLf & "No = Use Realization QuoteTo Sell Amount on Report", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Realization Amount") '04-17-15 JTC 
                RealQuoteToAmtON = False '07-15-14  Public RealQuoteToAmtON As Boolean = 0 
                If Resp = vbNo Then
                    RealQuoteToAmtON = True '07-15-14  Public RealQuoteToAmtON As Boolean = 0 
                    ' "SELECT projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, 
                    strSql = Replace(strSql, "quote.Sell as", "projectcust.Sell as")
                    strSql = Replace(strSql, "quote.Cost as", "projectcust.Cost as")
                    strSql = Replace(strSql, "quote.Comm as", "projectcust.Comm as")
                End If
                SaveStrSQL = strSql '11-27-13 
                SaveStrSQL2 = strSql '11-27-13  Sample = " and (quote.BranchCode = '" & SecurityBrancheCodes & "'" & " or quote.BranchCode = '')"
                SaveStrSQL2 = Replace(SaveStrSQL2, "where Quote.TypeOfJob = 'Q'", "where projectcust.Ncode = '" & RealWithOneMfgCustCode.Trim & "'" & " and Quote.TypeOfJob = 'Q' ") '02-01-14
                '07-15-14 Delete this if one projectCust
                'and (projectcust.TypeC = 'X'  or projectcust.TypeC = 'A'  or projectcust.TypeC = 'E'  or projectcust.TypeC = 'L'  or projectcust.TypeC = 'S'  or projectcust.TypeC = 'T'  or projectcust.TypeC = 'X' )
                '07-15-14 JTC If  RealArchitect = False and RealEngineer =  False and RealLtgDesigner =  False and RealSpecifier =  False and RealContractor =  False and RealOther =  False then RealALL = True' All Specifiers
                Dim SelTypec As String = "" '07-15-14 = 'M'
                If RealManufacturer = False And RealCustomer = True And RealQuoteTOOther = False And RealSLSCustomer = False Then SelTypec = "C" '07-15-14  = 'M'Me.pnlQutRealCode.Text = "Select Cust Code"
                '07-15-14 JTC RealManufacturer Only
                If RealManufacturer = True And RealCustomer = False And RealQuoteTOOther = False And RealSLSCustomer = False Then SelTypec = "M" '07-15-14 Me.pnlQutRealCode.Text = "Select MFG Code" '"Select MFG Code"
                Dim SaveOrderBy As String = ""
                If myConnection.State <> ConnectionState.Open Then Call OpenSQL(myConnection) '05-31-12
                MyCommand.Connection = myConnection
                If RealWithOneMfgCustCode.Trim = frmQuoteRpt.txtQutRealCode.Text And (SelTypec = "M" Or SelTypec = "C") Then
                    OneNCodeOnly = RealWithOneMfgCustCode.Trim '07-15-14 Dim OneNCodeOnly As String = ""'07-15-14 'set Where ProjectCust.Ncode = 
                End If
                '04-16-15 JTC Fix Real Customers to one MFG See QuotesToOneMFG = RealWithOneMfgCustCode
                If Len(RealWithOneMfgCustCode.Trim) < 5 And OneNCodeOnly = "" Then 'Could Be Mfg
                    Resp = MessageBox.Show("If " & RealWithOneMfgCustCode & " is a MFG and you only want Quotes with this MFG answer Yes", "One MFG", MessageBoxButtons.YesNo)
                    If Resp = MsgBoxResult.Yes Then
                        QuotesToOneMFG = RealWithOneMfgCustCode
                        OneNCodeOnly = "" '04-16-15 JTC See if MFg RealWithOneMfgCustCode = "" 'Dim QuotesToOneMFG As String = "" '04-16-15 JTC
                    End If
                End If
                'strSql = "DROP TABLE IF EXISTS TMPREPORTS " '01-28-10
                MyCommand.CommandText = "DROP TABLE IF EXISTS TMPREPORTS " : MyCommand.ExecuteNonQuery()
                '07-15-14 Select add projectcust.Sell as SellQ, projectcust.Cost as CostQ, projectcust.Comm as CommQ,
                '            "SELECT PROJECTCUST.Ncode, PROJECTCUST.Ncode as TMPNCode, PROJECTCUST.Typec as TMPTypec, projectcust.Sell as TMPSellQ, projectcust.Cost as TMPCostQ, projectcust.Comm as TMPCommQ, Quote.TypeOfJob, Quote.EntryDate, Quote.QuoteID FROM QUOTE INNER JOIN PROJECTCUST ON QUOTE.QUOTEID = PROJECTCUST.QUOTEID   Where projectcust.Ncode = 'KEEN' and Quote.TypeOfJob = 'Q'  and  Quote.EntryDate >= '2009-07-01' and Quote.EntryDate <= '2014-07-31'
                '12-31-14 JTC Fix IF(projectcust.NCODE = '', Left(projectcust.FIRMNAME,6), projectcust.ncode) as NCode, 
                SaveStrSQL = "SELECT PROJECTCUST.Ncode, PROJECTCUST.Ncode as TMPNCode, PROJECTCUST.Typec as TMPTypec, projectcust.Sell as TMPSellQ, projectcust.Cost as TMPCostQ, projectcust.Comm as TMPCommQ, Quote.TypeOfJob, Quote.EntryDate, Quote.QuoteID, Quote.JobName FROM QUOTE INNER JOIN PROJECTCUST ON QUOTE.QUOTEID = PROJECTCUST.QUOTEID  " '11-12-14 JTC added quote.JobName
                '12-31-14 JTC NoNo SaveStrSQL = "SELECT PROJECTCUST.Ncode, IF(projectcust.NCODE = '', Left(projectcust.FIRMNAME,6), projectcust.ncode) as  TMPNCode, PROJECTCUST.Typec as TMPTypec, projectcust.Sell as TMPSellQ, projectcust.Cost as TMPCostQ, projectcust.Comm as TMPCommQ, Quote.TypeOfJob, Quote.EntryDate, Quote.QuoteID, Quote.JobName FROM QUOTE INNER JOIN PROJECTCUST ON QUOTE.QUOTEID = PROJECTCUST.QUOTEID  " '11-12-14 JTC added quote.JobName
                SaveStrSQL += " Where PROJECTCUST.Ncode = '" & RealWithOneMfgCustCode & "' and Quote.TypeOfJob = 'Q' " '01-21-13Ignore Type works for both Mfg and Specifierand PROJECTCUST.TYPEC = 'M' " ' "
                '06-25-15 No NO No JTC Chg <> '' to = ""
                '??????????????????
                'If QuotesToOneMFG = RealWithOneMfgCustCode Then 
                If QuotesToOneMFG.Trim <> "" Then ' RealWithOneMfgCustCode = "" '04-16-15 JTC See if MFg 
                    SaveDateSql = Replace(SaveDateSql, "and Quote.Status <> 'NOREPT' ", "") '06-30-15 '09-17-15 JTC Fic Q.Status to Quote.Status in FillQutRealLUDataSet
                End If
                SaveStrSQL += SaveDateSql '01-21-14 Added Dates
                SaveStrSQL += " " & OrderBy '10-13-14 JTC wipes out Order BY
                '   No quote,status
                SaveStrSQL = "CREATE TEMPORARY TABLE TMPREPORTS AS " & SaveStrSQL 'Select  OL.Qty, OL.LnCode, O.*, Sum((OL.Sell * OL.Qty)) as ExtSell, Sum((OL.Comm * OL.Qty)) as ExtComm, OL.MFG as LMFG, OS.SLSCode as SLSCode2 from ORDERMASTER O  left join ordslssplit OS on O.OrderID = OS.OrderID JOIN projectlines OL ON O.OrderID = OL.OrderID Where O.MFG = 'PHIL' and OS.slsnumber = 1 and O.RelHold = 'R'  and  OL.MFG <> '' and OL.Qty <> '' and OL.LnCode <> 'BTX'  and OL.LnCode <> 'NPN' and (concat(O.BuySellAB, O.BuySellSR) <> 'BS') GROUP BY PONumber, OL.MFG order by PONumber, OL.MFG "
                MyCommand.CommandText = SaveStrSQL : SubCount = MyCommand.ExecuteNonQuery() 'Get All Quotes With This Code Assigned to it
                strSql = Replace(SaveStrSQL2, "= projectcust.QuoteID", "= projectcust.QuoteID inner join TMPREPORTS TR ON TR.QuoteID = QUOTE.QUOTEID ")
                'Stop 'Replace projectcust.Ncode = 'GLOB'  it is in TMPREPORTS
                '06-25-15 JTC Chg Below from RealWithOneMfgCustCode to QuotesToOne to leave it in SQLMFG 04-16-15 JTC See if MFg QuotesToOneMFG 
                '04-16-15 JTC See if MFg RealWithOneMfgCustCode = "" 'Dim QuotesToOneMFG As String = "" '04-16-15 JTC
                If QuotesToOneMFG.Trim = "" And OneNCodeOnly = "" And RealWithOneMfgCustCode = "" Then '06-25-15    OneNCodeOnly = ""
                    strSql = Replace(strSql, " projectcust.Ncode = '" & RealWithOneMfgCustCode.Trim & "' and", "") '02-01-14 JTC Fix Real When One Code selected
                End If
                'getspecifiers(quote.quoteid), strings specifiers out to right
                'I = InStr(strSql, ")") 'If F <> 0 Then SaveStrSQL = Mid(strSql, F - 1) 'SaveStrSQL = Replace(SaveStrSQL.ToUpper, "FROM ", ", OL.Qty, OL.LnCode, Sum(OL.Sell * OL.Qty) as ExtSell, Sum(OL.
                'SaveStrSQL2= SELECT projectcust.*, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy, quote.CSR, quote.LotUnit, quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID where Quote.TypeOfJob = 'Q'  and  Quote.EntryDate >= '2009-01-01' and Quote.EntryDate <= '2014-01-31'  and (projectcust.TypeC = 'M'  or projectcust.TypeC = 'M' )  order by projectcust.Ncode, JobName 
                '07-15-14 JTCReplace projectcust.Sell as SellQ, projectcust.Cost as CostQ, projectcust.Comm as CommQ withTR.SellQ, TR.CostQ, TR.CommQ,
                'If PC="C" of "M" use PC sell else use quote.sell projectcust.TypeC = 'M' 
                '07-15-14 nostrSql = Replace(strSql, "projectcust.Sell as SellQ, projectcust.Cost as CostQ, projectcust.Comm as CommQ,", "TR.SellQ, TR.CostQ, TR.CommQ, ") '07-15-14
                '                DROP TABLE IF EXISTS TMPREPORTS;
                'CREATE TEMPORARY TABLE TMPREPORTS AS SELECT PROJECTCUST.Ncode, PROJECTCUST.Ncode as TMPNCode, PROJECTCUST.Typec as TMPTypec, projectcust.Sell as TMPSellQ, projectcust.Cost as TMPCostQ, projectcust.Comm as TMPCommQ, Quote.TypeOfJob, Quote.EntryDate, Quote.QuoteID FROM QUOTE INNER JOIN PROJECTCUST ON QUOTE.QUOTEID = PROJECTCUST.QUOTEID   Where PROJECTCUST.NCode = 'KEEN' and Quote.TypeOfJob = 'Q'  and  Quote.EntryDate >= '2009-07-01' and Quote.EntryDate <= '2014-07-31';
                'DROP TABLE IF EXISTS TMPREPORTS2;
                'CREATE TEMPORARY TABLE TMPREPORTS2 AS SELECT projectcust.*, TR.TMPNCode, TR.TMPTypec, TR.TMPSellQ, TR.TMPCostQ, TR.TMPCommQ, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State,  quote.CSR,  quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID inner join TMPREPORTS TR ON TR.QuoteID = QUOTE.QUOTEID  where Quote.TypeOfJob = 'Q'   and  Quote.EntryDate >= '2009-07-01' and Quote.EntryDate <= '2014-07-31' ORDER BY  projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC;
                'Select * from TMPREPORTS2;
                'Delete from TMPREPORTS2 where NCode <> 'KEEN' and Typec = 'M';
                'Select count(*) from TMPREPORTS2;
                strSql = Replace(strSql, "quote.lastChgBy,", "") '07-15-14 
                strSql = Replace(strSql, "quote.LotUnit, quote.StockJob,", "") '07-15-14
                MyCommand.CommandText = "DROP TABLE IF EXISTS TMPREPORTS2 " : MyCommand.ExecuteNonQuery()
                'TR.TMPNCode, TR.TMPTypec, TR.TMPSellQ, TR.TMPCostQ, TR.TMPCommQ, 
                '09-30-14 JTC Fix Dates FillQutRealLUDataSet( Relization with one MFG Bad Date  Quote.EntryDate >= '" & sStartDate & "' and Quote.EntryDate <= '" & sEndDate & "'
                '09-30-14 Bad Date strSql = "SELECT projectcust.*, TR.TMPNCode, TR.TMPTypec, TR.TMPSellQ, TR.TMPCostQ, TR.TMPCommQ, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State,  quote.CSR,  quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID inner join TMPREPORTS TR ON TR.QuoteID = QUOTE.QUOTEID  where Quote.TypeOfJob = 'Q'   and  Quote.EntryDate >= '2009-07-01' and Quote.EntryDate <= '2014-07-31' ORDER BY  projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC "
                '04-16-15 JTC Wipe out Save Where
                strSql = "SELECT projectcust.*, TR.TMPNCode, TR.TMPTypec, TR.TMPSellQ, TR.TMPCostQ, TR.TMPCommQ, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State,  quote.CSR,  quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID inner join TMPREPORTS TR ON TR.QuoteID = QUOTE.QUOTEID  where Quote.TypeOfJob = 'Q'   and  Quote.EntryDate >= '" & sStartDate & "' and Quote.EntryDate <= '" & sEndDate & "' " ''10-13-14 JTC Add BelowORDER BY  projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC " '09-30-14 Jtc Fix Date
                If QuotesToOneMFG <> "" Then '04-16-15 JTC 
                    ''09-17-15 If One MFG and have code you don't need where" & WhereSqlTypec & " = projectcust.TypeC = 'M' 
                    strSql = "SELECT projectcust.*, TR.TMPNCode, TR.TMPTypec, TR.TMPSellQ, TR.TMPCostQ, TR.TMPCommQ, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State,  quote.CSR,  quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID inner join TMPREPORTS TR ON TR.QuoteID = QUOTE.QUOTEID  where Quote.TypeOfJob = 'Q'   and  Quote.EntryDate >= '" & sStartDate & "' and Quote.EntryDate <= '" & sEndDate & "' " '09-17-15 Fix Where and SQL
                    strSql += " and PROJECTCUST.Ncode = '" & RealWithOneMfgCustCode & "' " ''06-25-15 JTC add PROJECTCUST.Ncode = '" & RealWithOneMfgCustCode & "' 
                End If
                strSql += OrderBy '10-13-14 JTC wipes out Order BY
                If (OneNCodeOnly = RealWithOneMfgCustCode.Trim And RealWithOneMfgCustCode.Trim <> "") Or (RealWithOneMfgCustCode.Trim <> "" And QuotesToOneMFG = "") Then '04-16-15  '04-16-15 JTC07-15-14
                    '04-16-15 JTC Don't do Only One Mfg
                    strSql = Replace(strSql, " where Quote.TypeOfJob = 'Q' ", " where Quote.TypeOfJob = 'Q' and projectcust.Ncode = '" & RealWithOneMfgCustCode.Trim & "' ") '07-15-14 2-03-14 JTC Fix Real When One Code selected
                    'done already strSql = Replace(strSql, "ORDER BY  projectcust.Ncode, projectcust.QuoteCode, projectcust.NotGot DESC", "") '10-13-14 JTC
                End If
                ': strSql += OrderBy '10-13-14 JTC wipes out Order BY
                MyCommand.CommandText = "CREATE TEMPORARY TABLE TMPREPORTS2 AS " & strSql : SubCount = MyCommand.ExecuteNonQuery()
                'MyCommand.CommandText = "Select * from TMPREPORTS2 " : SubCount = MyCommand.ExecuteNonQuery()
                Dim tmpds As dsSaw8 = New dsSaw8
                Dim daQ = New MySqlDataAdapter
                daQ.SelectCommand = New MySqlCommand("Select * from TMPREPORTS2", myConnection)
                daQ.Fill(tmpds, "TMPREPORTS2")
                'Debug.Print(tmpds.Tables("TMPREPORTS2").Rows.Count)

                daQuoteRealLU = New MySqlDataAdapter : strSql = "Select * from TMPREPORTS2"
                daQuoteRealLU.SelectCommand = New MySqlCommand(strSql, myConnection)
                Dim SqlDel As String = ""
                'Delete Other MFGs if any
                '04-16-15 Ship below)
                If QuotesToOneMFG = "" Then '04-16-15 JTC Not
                    MyCommand.CommandText = "Delete from TMPREPORTS2 where NCode <> '" & RealWithOneMfgCustCode.Trim & "' and Typec = '" & SelTypec & "' " : SubCount = MyCommand.ExecuteNonQuery()
                End If
                If SelTypec = "C" And OneNCodeOnly = "" Then 'If Customers Delete all MFGs
                    'Deletes 04-16-15 JTC added And OneNCodeOnly above to not deletes this MFG
                    MyCommand.CommandText = "Delete from TMPREPORTS2 where Typec = 'M' " : SubCount = MyCommand.ExecuteNonQuery()
                End If
                'strSql = "Delete from TMPREPORTS2 where NCode <> '" & RealWithOneMfgCustCode.Trim & "' and Typec = '" & SelTypec & "' "
                'daQuoteRealLU = New MySqlDataAdapter
                'daQuoteRealLU.SelectCommand = New MySqlCommand(strSql, myConnection)
                strSql = "Select * from TMPREPORTS2"
                daQ.SelectCommand = New MySqlCommand("Select * from TMPREPORTS2", myConnection)
                daQ.Fill(tmpds, "TMPREPORTS2")
                'Debug.Print(tmpds.Tables("TMPREPORTS2").Rows.Count)
            End If

            ''*******************************************************************************************************************************
            '11-29-13 As SELECT projectcust.*, Quote.QuoteCode as QuoteCodeQ, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy as LastChgByQ, quote.CSR as CSRQ, quote.LotUnit as LotUnitQ, quote.StockJob, quote.TypeOfJob FROM Quote LEFT OUTER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID where Quote.TypeOfJob = 'Q'
            'No as       SELECT projectcust.*,                                quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy,               quote.CSR,         quote.LotUnit,             quote.StockJob, quote.TypeOfJob FROM Quote INNER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID 
            'Puts Got First then LastChgDate = Most Current =s Our Old Report
            If RealTgLookupExcel = True Then '11-27-13 and Quote.EntryDate <= '2013-11-30'  and (projectcust.TypeC = 'C'  or projectcust.TypeC = 'C' )  order by projectcust.Ncode, JobName "
                If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then  '01-21-14 
                    strSql = Replace(strSql, "quote.lastChgBy, quote.CSR, quote.LotUnit", "quote.lastChgBy as LastChgByQ, quote.CSR, quote.LotUnit as LotUnitQ") '11-29-13
                    '01-21-14 SaveStrSQL2 = Replace(SaveStrSQL2, "quote.lastChgBy, quote.CSR, quote.LotUnit", "quote.lastChgBy as LastChgByQ, quote.CSR, quote.LotUnit as LotUnitQ") '11-29-13
                    'SaveStrSQL2 = All before Projectcust  strSql If Right(strSql, 2) <> ") " And IStr(strSql, "(") Then strSql += ") "
                    'Quote INNER JOIN to "Quote LEFT OUTER JOIN"
                    strSql = Replace(strSql, "Quote INNER JOIN", "Quote LEFT OUTER JOIN") '11-27-13
                    strSql = Replace(strSql, "projectcust.*, ", "projectcust.*, Quote.QuoteCode, ") '11-27-13projectcust.*, Quote.QuoteCode, 
                    '02-03-14 strSql = Replace(strSql, "projectcust.Ncode = 'KEEN' and ", "") '01-21-14 Get Specifiers projectcust.Ncode = 'KEEN'
                    strSql = Replace(strSql, " projectcust.Ncode = '" & RealWithOneMfgCustCode.Trim & "' and", "") '02-03-14 JTC Fix Real When One Code selected
                    'SaveDateSql
                    F = InStr(strSql, "order by ")
                    If F <> 0 Then SaveStrSQL = Mid(strSql, F - 1) 'SaveStrSQL = Order By 
                    '01-21-14 SaveStrSQL2 = Replace(SaveStrSQL2, "Quote INNER JOIN", "Quote LEFT OUTER JOIN") '11-27-13
                    '01-21-14 SaveStrSQL2 = Replace(SaveStrSQL2, "projectcust.*, ", "projectcust.*, Quote.QuoteCode as QuoteCodeQ, ") '11-29-13projectcust.*, Quote.QuoteCode, 
                    'SaveStrSQL = order by strSql '11-27-13 projectcust.TypeC
                    'getspecifiers(quote.quoteid), strings specifiers out to right
                    'SaveStrSQL2 = Replace(SaveStrSQL2, "projectcust.*, ", "projectcust.*, getspecifiers(quote.quoteid), ") '01-20-14 
                    strSql = Replace(strSql, "projectcust.TypeC = 'M'  or projectcust.TypeC = 'M'", " projectcust.TypeC = 'A' or projectcust.TypeC = 'E' or projectcust.TypeC = 'L' or projectcust.TypeC = 'S' ") '01-20-14
                    'strSql = SaveStrSQL2 '11-27-13 = All before Projectcust 
                    'Just + QuoteID SELECT projectcust.*, Quote.QuoteCode as QuoteCodeQ, quote.Sell as SellQ, quote.Cost as CostQ, quote.Comm as CommQ, quote.JobName, quote.MarketSegment, quote.EntryDate, quote.BidDate, quote.SLSQ, quote.Status, quote.RetrCode, quote.SelectCode, quote.City, quote.State, quote.lastChgBy as LastChgByQ, quote.CSR, quote.LotUnit as LotUnitQ, quote.StockJob, quote.TypeOfJob FROM Quote LEFT OUTER JOIN projectcust ON Quote.QuoteID = projectcust.QuoteID where Quote.TypeOfJob = 'Q'   and  Quote.EntryDate >= '2009-01-01' and Quote.EntryDate <= '2014-01-31'  and (projectcust.TypeC = 'S'  or projectcust.TypeC = 'A' or projectcust.TypeC = 'E' or projectcust.TypeC = 'L' or projectcust.TypeC = 'S') ORDER BY  projectcust.SLSCode, projectcust.NCode, projectcust.NotGot DESC, projectcust.LastChgDate DESC
                Else
                    strSql = Replace(strSql, "quote.lastChgBy, quote.CSR, quote.LotUnit", "quote.lastChgBy as LastChgByQ, quote.CSR, quote.LotUnit as LotUnitQ") '11-29-13
                    SaveStrSQL2 = Replace(SaveStrSQL2, "quote.lastChgBy, quote.CSR, quote.LotUnit", "quote.lastChgBy as LastChgByQ, quote.CSR, quote.LotUnit as LotUnitQ") '11-29-13
                    'SaveStrSQL2 = All before Projectcust  strSql If Right(strSql, 2) <> ") " And IStr(strSql, "(") Then strSql += ") "
                    'Quote INNER JOIN to "Quote LEFT OUTER JOIN"
                    strSql = Replace(strSql, "Quote INNER JOIN", "Quote LEFT OUTER JOIN") '11-27-13
                    strSql = Replace(strSql, "projectcust.*, ", "projectcust.*, Quote.QuoteCode, ") '11-27-13projectcust.*, Quote.QuoteCode, 
                    F = InStr(strSql, "order by ")
                    If F <> 0 Then SaveStrSQL = Mid(strSql, F - 1) 'SaveStrSQL = Order By 
                    SaveStrSQL2 = Replace(SaveStrSQL2, "Quote INNER JOIN", "Quote LEFT OUTER JOIN") '11-27-13
                    SaveStrSQL2 = Replace(SaveStrSQL2, "projectcust.*, ", "projectcust.*, Quote.QuoteCode as QuoteCodeQ, ") '11-29-13projectcust.*, Quote.QuoteCode, 
                    'SaveStrSQL = order by strSql '11-27-13 
                    strSql = SaveStrSQL2 '11-27-13 = All before Projectcust 
                End If
            End If
            daQuoteRealLU = New MySqlDataAdapter
            daQuoteRealLU.SelectCommand = New MySqlCommand(strSql, myConnection)
            'Dim cbQutlu As MySql.Data.MySqlClient.MySqlCommandBuilder
            'cbQutlu = New MySqlCommandBuilder(daQutLU)
            daQuoteRealLU.Fill(dsQuoteRealLU, "QuoteRealLU")
            'okDebug.Print(SortSeq)
            frmQuoteRpt.QuoteRealLUBindingSource.DataSource = dsQuoteRealLU.QuoteRealLU

            If RealCustomer = True Or RealManufacturer = True Then '02-20-17
                GoTo IncSPEC
            End If

            If RealWithOneMfgCustCode <> "" Then '10-17-16 JH 
IncSPEC:        Resp = MessageBox.Show("Do you want to include Specifiers?" & vbCrLf & "Note:  Report takes longer to generate", "Fill Specifiers", MessageBoxButtons.YesNo)
                If Resp = vbYes Then
                    Dim lblTemp As New System.Windows.Forms.Label '02-20-17
                    lblTemp.Size = New Size(290, 60)
                    lblTemp.Name = "Progress"
                    frmQuoteRpt.Controls.Add(lblTemp)
                    lblTemp.Text = "Loading Specifiers:"
                    lblTemp.TextAlign = ContentAlignment.MiddleCenter
                    lblTemp.Location = New Point(100, 100)
                    lblTemp.BackColor = Color.LightGreen
                    lblTemp.BringToFront()
                    Application.DoEvents()
                    Dim Cnt As Integer = 1
                    For Each dr As dsSaw8.QuoteRealLURow In dsQuoteRealLU.QuoteRealLU.Rows
                        Dim tmpspec As dsSaw8 = New dsSaw8 : tmpspec.EnforceConstraints = False
                        Dim tmpda As MySqlDataAdapter = New MySqlDataAdapter
                        tmpda.SelectCommand = New MySqlCommand("SELECT * FROM PROJECTCUST WHERE QUOTECODE = '" & SafeSQL(dr.QuoteCode) & "' and (projectcust.TypeC = 'X'  or projectcust.TypeC = 'A'  or projectcust.TypeC = 'E'  or projectcust.TypeC = 'L'  or projectcust.TypeC = 'S'  or projectcust.TypeC = 'T')", myConnection)
                        tmpda.Fill(tmpspec, "projectcust")
                        For Each drSpec As dsSaw8.projectcustRow In tmpspec.projectcust.Rows
                            lblTemp.Text = "Loading Specifiers: " & Cnt & "  of " & dsQuoteRealLU.QuoteRealLU.Rows.Count : Application.DoEvents() '02-21-17
                            dr.Specifiers += "," & drSpec.FirmName
                            dr.SLSCode += "," & drSpec.SLSCode
                            If drSpec.Typec = "A" Then
                                dr("Architect") = drSpec.FirmName
                            ElseIf drSpec.Typec = "E" Then
                                dr("Engineer") = drSpec.FirmName
                            ElseIf drSpec.Typec = "T" Then
                                dr("Contractor") = drSpec.FirmName
                            ElseIf drSpec.Typec = "S" Then
                                dr("Specifier") = drSpec.FirmName
                            End If
                            Cnt += 1
                        Next
                    Next
                    lblTemp.Visible = False : Application.DoEvents() '02-21-17
                End If
            End If
            'Debug.Print(SaveStrSQL2)
            'Debug.Print("and  projectcust.Sell >= '5000'")
            If SESCO = True Or ExcelQuoteFU = True Then '04-28-14 JTC Public BoolfrmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-22-15 JTC 04-17-12
                dsSESCOSpecifiers = New dsSaw8
                dsSESCOSpecifiers.EnforceConstraints = False
                Dim daSESCO As MySqlDataAdapter = New MySqlDataAdapter
                Dim tmpSQL As String = strSql
                If ExcelQuoteFU = True Then
                    'I = InStr(strSql, " projectcust.Sell >= ")
                    'If I <> 0 Then F = InStr(I, strSql, "and") ' else Exit if 
                    'If F <> 0 Then SaveStrSQL2 = Mid(strSql, F + 3) 'SaveStrSQL = Order By 
                    'tmpSQL = SaveStrSQL2 & " " & SaveStrSQL
                    Dim RString As String = " and  projectcust.Sell >= '" & frmQuoteRpt.txtStartQuoteAmt.Text & "'" '04-30-15
                    tmpSQL = strSql.Replace(RString, "")
                End If
                tmpSQL = tmpSQL.Replace("TYPEC = 'C'", " (TypeC = 'A' or TypeC = 'E' or TypeC = 'T' )") '("TYPEC = 'C'", " (TypeC = 'A' or TypeC = 'E' or TypeC = 'L' or TypeC = 'T' or TypeC = 'S')")
                daSESCO.SelectCommand = New MySqlCommand(tmpSQL, myConnection)
                daSESCO.Fill(dsSESCOSpecifiers, "QuoteRealLU")
            End If
            'myView.Sort = "FirmName"
            'frmQuoteRpt.QUTLU1TableAdapterBindingSource = dsSaw8.QUTLU1DataTable
            frmQuoteRpt.tgr.DataSource = frmQuoteRpt.QuoteRealLUBindingSource
            frmQuoteRpt.tgr.Rebind(True)
            If SESCO = True Or ExcelQuoteFU = True Then frmQuoteRpt.tgr.Rebind(False) '04-28-14 JTC Public BoolfrmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then frmQuoteRpt.tgr.Rebind(False) '04-22-15 JTC 02-24-12
            'Set col.DataColumn.Tag After ReBind in each FillQutRealLUDataSet '02-26-09
            'For I = 0 To frmQuoteRpt.tgr.Splits(0).DisplayColumns.Count - 1 '02-26-09 
            '    'If I > 41 Then Exit For '02-19-09 
            '    Dim col As C1.Win.C1TrueDBGrid.C1DisplayColumn = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I) '02-20-09 
            '    col.DataColumn.Tag = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name.ToString '07-09-09 Tag = Name  col.DataColumn.Tag = I.ToString 'Add Tag to each Column
            'Next
            If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then '03-22-13
                '01-20-14 strSql = "DROP TABLE IF EXISTS TMPREPORTS " '01-28-10
                Dim TmpstrSql As String = "DROP TABLE IF EXISTS TMPREPORTS " '01-20-14
                MyCommand.CommandText = TmpstrSql : MyCommand.ExecuteNonQuery()
                TmpstrSql = "DROP TABLE IF EXISTS TMPREPORTS2 " '07-15-14 
                MyCommand.CommandText = TmpstrSql : MyCommand.ExecuteNonQuery()
            End If
            RealWithOneMfgCustCode = QuotesToOneMFG : QuotesToOneMFG = "" : WhereSqlTypec = ""  '04-16-15 JTC
        Catch ex As Exception
            MessageBox.Show("Error in FillQutRealLUDataSet (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox("FillQutRealLUDataSet " & ex.Message)
            ' If DebugOn ThenStop
        End Try
    End Sub
    Public Function MarginOrCommCalc(ByVal Sell As Double, ByVal Cost As Double, Optional ByVal Percent As Boolean = True) As Decimal
        On Error Resume Next

        'THIS CALCULATES COMM%,  MARGIN % or MARGIN $

        'If DIST e.Value = Format(MarginOrCommCalc(Val(tgQt(e.Row, "Sell")), Val(tgQt(e.Row, "Cost"))), "###.00") 'Margind
        'If REP e.Value = Format(MarginOrCommCalc(Val(tgQt(e.Row, "Sell")), Val(tgQt(e.Row, "Comm"))), "###.00") 'Commission %
        Dim Mval As Decimal = 0
        If DIST Then '09-04-09
            If Percent = True Then '01-13-11
                If Sell <> 0 Then Mval = CDec(((Sell - Cost) / Sell) * 100) '08-09-09
                If Mval < -99 Then Mval = -99
                If Mval > 99 Then Mval = 99
            Else
                If Sell <> 0 Then Mval = CDec(Sell - Cost) '01-12-11
            End If
            MarginOrCommCalc = Decimal.Round(Mval, 2) '08-12-09
        Else  ' Cost = Comm$ for rep
            'FixMargin = ((FixSell - FixCost) / (FixSell + 0.0001)) * 100
            'If FixMargin > 900 Then FixMargin = 999 Else If FixMargin < -900 Then FixMargin = -999 '06-24-04
            Mval = CDec(((Sell - Cost) / (Sell + 0.0001)) * 100)
            If Mval < -99 Then Mval = -99 '11-08-06 JH - for margin of 100 this prints 99
            If Mval > 99 Then Mval = 99
            MarginOrCommCalc = Decimal.Round(Mval, 2) '08-12-09
        End If
    End Function
    'Public Function MarginOrCommCalc(ByVal Sell As Decimal, ByVal Cost As Decimal) As Decimal '03-09-11 Doubl to Decimal
    '    'If DIST e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), Val(tgQh(e.Row, "Cost"))), "###.00") 'Margind
    '    'If REP e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), Val(tgQh(e.Row, "Comm"))), "###.00") 'Commission %
    '    Dim Mval As Decimal = 0
    '    If Sell = 0 Then Sell = 0.01 '03-09-11 Jtc Error if Sell = -0.01 + 0.0001 to 0.01 Decimal MarginOrCommCalc(decimal)
    '    If DIST Then '09-04-09
    '        If Sell <> 0 Then Mval = CDec((Sell - Cost) / Sell * 100) '08-09-09
    '        If Mval < -99 Then Mval = -99 '11-08-06 JH - for margin of 100 this prints 99
    '        If Mval > 99 Then Mval = 99
    '        MarginOrCommCalc = Decimal.Round(Mval, 2) '08-12-09
    '    Else  ' Cost = Comm$ for rep
    '        'Rep Cost = Sell - Comm$ 
    '        '03-06-11 Mval = CDec((Cost / Sell) * 100) '03-06-11 In Order System
    '        Mval = CDec((Sell - Cost) / Sell * 100) '03-09-11 If Sell = 0 Then Sell = 0.01 '02-28-11 Jtc Error if Sell = -0.01 + 0.0001 to 0.01 Decimal MarginOrCommCalc(decimal)
    '        If Mval < -99 Then Mval = -99 '11-08-06 JH - for margin of 100 this prints 99
    '        If Mval > 99 Then Mval = 99
    '        MarginOrCommCalc = Decimal.Round(Mval, 2) '08-12-09
    '    End If
    'End Function
    Public Function BuildSQLOrderBY() As String
        'Dim OrderBy As String = ""
        '
        'already done Call SetPrimarySortValues() '01-18-09
        'sets Public OrderBy = "Q.QuoteCode"  US = "Quote Code Sequence":UH = "QUOTE CODE SEQUENCE"
        'VQRT2.RepType = VQRT2.RptMajorType.RptQutCode : UH = "QUOTE CODE SEQUENCE" : Sorted = 0 : JOBSER = ""
        'MU = VQRT2.RptMajorType.RptQutCode
        'Dim A As String = UH
        '01-30-09 
        If VQRT2.RepType = VQRT2.RptMajorType.RptBidDate = True Then OrderBy = "Q.BidDate"
        If VQRT2.RepType = VQRT2.RptMajorType.RptDescend Then OrderBy = "Q.Sell DESC" '03-08-09
        If VQRT2.RepType = VQRT2.RptMajorType.RptEntryDate Then OrderBy = "Q.EntryDate" '01-19-10
        If VQRT2.RepType = VQRT2.RptMajorType.RptLocation Then OrderBy = "Q.Location" '11-19-10 City" Location"
        If VQRT2.RepType = VQRT2.RptMajorType.RptMarketSegment Then OrderBy = "Q.MarketSegment" '10-31-12 Chg P. to Q.
        If VQRT2.RepType = VQRT2.RptMajorType.RptProj Then OrderBy = "Q.JobName" '11-23-11 quote.JobName Not project.ProjectName
        If VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Then OrderBy = "Q.QuoteCode"
        If VQRT2.RepType = VQRT2.RptMajorType.RptRetrieval Then OrderBy = "Q.RetrCode"
        If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Then OrderBy = "Q.SLSQ"
        If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Then OrderBy = "Q.FollowBy" '03-01-12
        If VQRT2.RepType = VQRT2.RptMajorType.RptEnteredBy Then OrderBy = "Q.EnteredBy" '05-14-13
        If VQRT2.RepType = VQRT2.RptMajorType.RptSpecif Then OrderBy = "PC.NCode"
        If VQRT2.RepType = VQRT2.RptMajorType.RptStatus Then OrderBy = "Q.Status"
        'VQRT2.RptMajorType.RptStatus() 'RptSpecif'RptSalesman'RptRetrieval'RptQutCode'RptProj'RptMarketSegment'RptLocation'RptEntryDate
        '.RptDescend
        If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman And frmQuoteRpt.chkSlsFromHeader.CheckState = CheckState.Checked Then
            '03-08-13 JTC Add QS.SLSCode as SLS1 from Quote QUTSLSSPLIT Table Position 1
            OrderBy = "QS.SLSCode" 'Salesman
            ' Me.chkSlsFromHeader.CheckState = CheckState.Unchecked  ' Me.chkSlsFromHeader.Text = "Use Quote SLS 1 Split for Salesman" '03-08-13 "Use Salesman From Quote Header on Report"
        End If
        'Below Only on Job Qoute Commission Shortage
        Call SetSecondarySortValues() '01-24-09 Distributors don't have Orders or Invoices
        'If VQRT2.RepType = VQRT2.RptMajorType.RptStatus Then OrderBy = "Q.Status"
        If VQRT2.SubSeq = VQRT2.SubSortType.SubSBidDate Then OrderBy += ", Q.BidDate" '03-01-11
        If VQRT2.SubSeq = VQRT2.SubSortType.SubSDescend Then OrderBy += ", Q.Sell DESC" '11-20-10
        If VQRT2.SubSeq = VQRT2.SubSortType.SubSProj Then OrderBy += ", Q.JobName" '11-23-11 quote.JobName Not project.ProjectName
        If VQRT2.SubSeq = VQRT2.SubSortType.SubSEnterDate Then OrderBy += ", Q.EntryDate" '11-20-10CurrLev2 = drQRow.EntryDate 'SubSEnterDate'SubSProjCode'11-20-10
        If VQRT2.SubSeq = VQRT2.SubSortType.SubSProjCode Then OrderBy += ", Q.QuoteCode" '11-20-10 CurrLev2 = drQRow.QuoteCode 'SubSEnterDate'SubSProjCode'11-20-10
        If VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Then OrderBy += ", Q.SLSQ"
        If VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif Then OrderBy += ", PC.NCode"
        If VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Then OrderBy += ", Q.Status"
        If VQRT2.SubSeq = VQRT2.SubSortType.SubSSelectBidDate Then OrderBy += ", Q.SelectCode, Q.BidDate" '03-06-12 SubSeq
        'VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus
        If BranchReporting = True Then '10-30-12
            If OrderBy <> "" Then OrderBy = "ORDER BY " & " Q.BranchCode, " & OrderBy '10-30-12
        Else
            If OrderBy <> "" Then OrderBy = "ORDER BY " & OrderBy
        End If
        Return OrderBy
    End Function

    Public Sub RTColSize(ByVal RT As C1.C1Preview.RenderTable, ByVal MaxCol As Int16, ByRef TgWidth() As Single)
        Dim PrtCols As Int16 'PrtCols -= 1
        Dim ColWidtmp As Single
        Try
            Dim Widthtmp As String
            For I = 0 To MaxCol '02-06-09 frmQuoteRpt.tg.Splits(0).DisplayColumns.Count - 1 ' 18
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then '02-24-09
                    ' Dim col2 As C1.Win.C1TrueDBGrid.C1DisplayColumn = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I) '02-20-09
                    'Tag = col2.DataColumn.Tag
                    If frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                    If (frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                    ColWidtmp = (frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Width / 100).ToString
                    Widthtmp = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name
                    If Widthtmp = "Margin" Or Widthtmp = "Cost" Or Widthtmp = "Sell" Or Widthtmp = "Comm-$" Or Widthtmp = "Comm-%" Or Widthtmp = "Comm" Or Widthtmp = "BKSell" Or Widthtmp = "BKComm" Or Widthtmp = "Ext Cost" Or Widthtmp = "Ext Sell" Or Widthtmp = "Ext Comm" Or Widthtmp = "Ext Marg" Or Widthtmp = "LPCost" Or Widthtmp = "LPSell" Or Widthtmp = "UOverage" Or Widthtmp = "EntryDate" Or Widthtmp.ToString.StartsWith("Price") Then '02-17809 
                        RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
                    End If
                ElseIf frmQuoteRpt.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Then '09-01-09
                    Widthtmp = frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Name
                    If frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                    If (frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                    ColWidtmp = frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Width / 100
                    If Widthtmp = "Margin" Or Widthtmp = "Cost" Or Widthtmp = "Sell" Or Widthtmp = "Comm-$" Or Widthtmp = "Comm-%" Or Widthtmp = "Comm" Or Widthtmp = "BKSell" Or Widthtmp = "BKComm" Or Widthtmp = "Ext Cost" Or Widthtmp = "Ext Sell" Or Widthtmp = "Ext Comm" Or Widthtmp = "Ext Marg" Or Widthtmp = "LPCost" Or Widthtmp = "LPSell" Or Widthtmp = "UOverage" Or Widthtmp = "EntryDate" Or Widthtmp.ToString.StartsWith("Price") Then '02-18-09 
                        RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
                    End If
                Else
                    Widthtmp = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name
                    If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                    If (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                    ColWidtmp = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100
                End If
                '@#Q ProjectName0,ProjectID1,QuoteID2,QuoteCode3,EntryDate4,RetrCode5,PRADate6,EstDelivDate7,SLSQ8,Status9,BidDate10,Cost11,Sell12,Margin13,LPCost14,LPSell15,LPMarg16,LotUnit17,StockJob18,CSR19,LastChgBy20,HeaderTab21,LinesYN22,SelectCode23,Password24,FollowBy25,OrderEntryBy26,ShipmentBy27,Remarks28,LightingGear29,Dimming30,LastDateTime31,BidBoard32,EnteredBy33,BidTime34,BranchCode35,Address36,Address237,City38,State39,Zip40,Country41,Location42,LeadTime43,"'02-22-09
                '@#R ProjectCustID0,ProjectID1,NCode2,Got3,Typec4,QuoteCode5,ProjectName6,FirmName7,ContactName8,EntryDate9,SLSCode10,Status11,Cost12,Sell13,Margin14,LPCost15,LPSell16,LPMarg17,Overage18,ChgDate19,OrdDate20,NotGot21,Comments22,SPANumber23,SpecCross24,LotUnit25,LampsIncl26,Terms27,FOB28,QuoteID29,BranchCode30,MarketSegment31,MFGQuoteNumber32,BidDate33,SLSQ34,RetrCode35,SelectCode36,LeadTime37,"
                '07-07-09 If frmShowHideGrid.tgShow(Tag, 2) = False Then Continue For
                RT.Cols(PrtCols).Width = TgWidth(I) ' ColWidtmp ' frmQuoteRpt.tg.Splits(0).DisplayColumns(I).Width / 100 ' TgWidth(PrtCols) '12-01-08
                PrtCols += 1
            Next
        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "Print Task" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop
        End Try

    End Sub
    Public Sub ReportCriteriaSU()
        '
        dsRpt = New dsSaw8
        dsRpt.EnforceConstraints = False
        strSql = ""
        Dim MyDefaultsSQL As String = ""
        Dim MyQuotesSQL As String = ""
        Dim MySLSSQL As String = ""
        Try
            'SORT ORDER CODE
            OrderBy = BuildSQLOrderBY() '01-30-09
            'sets Public OrderBy = "Q.QuoteCode"  US = "Quote Code Sequence":UH = "QUOTE CODE SEQUENCE"
            'VQRT2.RepType = VQRT2.RptMajorType.RptQutCode : UH = "QUOTE CODE SEQUENCE" : Sorted = 0 : JOBSER = ""
            'MU = VQRT2.RptMajorType.RptQutCode
            If US = "Quote Code Sequence" Then '
                strSql = ""
                '09-07-09 strSql = BuildSQLQuotes("Q.", True, strSql)
                If Trim(frmQuoteRpt.txtStartEntry.Text) <> "" And Trim(frmQuoteRpt.txtStartEntry.Text) <> "ALL" Then
                    'If QM.Date_Renamed >= ReformatDate((frmQuoteRpt.txtStartEntry.Text)) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
                End If
                'Debug.Print(frmQuoteRpt.DTPickerStartEntry.Value.ToString)
                'Debug.Print(Quote.
                'Dim A As String = 

                '"WHERE 
                'If Trim(frmQuoteRpt.txtEndEntry.Text) <> "" And Trim(frmQuoteRpt.txtEndEntry.Text) <> "ALL" Then
                '    If QM.Date_Renamed <= ReformatDate((frmQuoteRpt.txtEndEntry.Text)) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
                'End If
                '    If strsql <> "" Then strsql = "WHERE " + strsql
                Dim mystring As String = strSql
                '    strsql = "SELECT C.*, N.*, S.* FROM NAMECONTACT C LEFT OUTER JOIN NAMEDETAIL N ON N.CODE = C.CODE LEFT OUTER JOIN namslssplit S ON S.CODE = N.CODE AND S.SLSNUMBER = '1' " 'OR SLSNUMBER = 2)"
                '    strsql += " " & mystring & " " & OrderBy
                'strSql = "Select Q.*, P.projectname, S.* from Quote Q LEFT join Project P on P.ProjectID = Q.ProjectID and join QutSLSSplit S on S.QuoteID = Q.QuoteID AND (S.SLSNumber = '1' OR S.SLSNumber = '2' OR S.SLSNumber = '3' OR S.SLSNumber = '4' order by Q.QuoteCode" '01-19-09
                '11-23-11 quote.JobName Not project.ProjectName strSql = "Select Q.*, P.projectname from Quote Q LEFT join Project P on P.ProjectID = Q.ProjectID order by Q.QuoteCode" '01-19-09'07-08-10 LEFT
                strSql = "Select Q.* from Quote Q order by Q.QuoteCode"
                daQuote = New MySql.Data.MySqlClient.MySqlDataAdapter
                daQuote.SelectCommand = New MySql.Data.MySqlClient.MySqlCommand(strSql, myConnection)
                Dim cbQut As MySql.Data.MySqlClient.MySqlCommandBuilder
                cbQut = New MySqlCommandBuilder(daQuote)
                daQuote.Fill(dsRpt, "QUTLU1")
                'MsgBox(dsRpt.quote.QuoteCodeColumn.ToString)

            End If
        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "Print Task" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop
        End Try
        Exit Sub
    End Sub

    Public Sub PrtQutSpreadSheet(ByRef PriceCol As Integer)
        '09-09-09  #Top
        Dim SellCost As String = "S" '02-15-10
        Dim UnitExtd As String = "E" '02-15-10
        Dim IncludeQty As String = "Y" '02-15-10
        Dim MC As Integer = 0
        Dim OldMfg As String = ""
        Dim Mth As Integer = 0
        Dim QI2Desc As String = ""
        Dim HeaderTxt As String = ""
        On Error Resume Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim TotMonthP(12) As Decimal 'TotMonthP(12),SubTotMonthP(12),GtSubTotMonthP(12)
        Dim TotMonthQ(12) As Decimal
        Dim SubTotMonthP(12) As Decimal
        Dim SubTotMonthQ(12) As Decimal
        Dim GtSubTotMonthP(12) As Decimal
        Dim GtSubTotMonthQ(12) As Decimal
        Dim MonthToPrt(12) As Int16 '09-14-09 
        Dim SubTotYTDP As Decimal = 0
        Dim SubTotYTDQ As Decimal = 0
        Dim TotYTDP As Decimal = 0
        Dim TotYTDQ As Decimal = 0
        Dim LineQty As Decimal = 0
        'Static PriCol As Integer
        'Try '#Top  Convert This sub to QuoteLines
        Dim MaxRow As Single = 0 '09-01-09
        Dim Row As Integer = 0 '09-09-09
        Dim LnPrcCol As Short = 0
        Dim LinesToPrt As Short = 1 '06-17-10 
        Dim sStartDate As String = VB6.Format(frmQuoteRpt.DTPickerStartEntry.Text, "yyyy") '  VB6.Format("01/01/2008", "yyyy-MM-dd") 'Testing
        Dim HdrColArray() As String = {sStartDate, "-JAN-", "-FEB-", "-MAR-", "-APR-", "-MAY-", "-JUN-", "-JUL-", "-AUG-", "-SEP-", "-OCT-", "-NOV-", "-DEC-", "-YTD-"}
        Dim newsStartDate As Date = VB6.Format(frmQuoteRpt.DTPickerStartEntry.Text, "yyyy/MM") '  VB6.Format("01/01/2008", "yyyy-MM-dd") 'Testing
        Dim PrtCols As Int16 = 0
        Dim StartPrtCols As Int16 = 0 '02-15-10
        If ExportExcelProductLines = True Then '02-15-10
            StartPrtCols = 2
        Else
            StartPrtCols = 0
        End If
        For I = 1 To 12    'Sets Heading to -Jun-  -Jul- etc for 12 months
            Dim Mon As Integer = newsStartDate.Month ' (DateInterval.Month) '  .AddMonths(I) 'Mon = Mon.DatePart(DateInterval.Month) ', Mon) 'Format(Mon, "mm") 'Mon.Month.ToString
            Dim MonCalc As Integer = I + Mon - 1
            If MonCalc > 12 Then MonCalc = MonCalc - 12
            MonthToPrt(I) = MonCalc
            Dim Monname As String = MonthName(MonCalc, True)
            HdrColArray(I) = "-" & Left(Monname, 3) & "-"
        Next I
        HdrColArray(13) = "-YTD-" '02-15-10
        If frmQuoteRpt.optoptSalesorCost_Sales.Checked = True Then SellCost = "S" Else SellCost = "C" '02-15-09 
        SI = Interaction.InputBox("Do You Want to Include the Quantity Row on the Report? (Y,N)", "Quantity Row", "Y") '02-15-10
        If UCase(SI) = "N" Then IncludeQty = "N" '02-15-10
        SI = Interaction.InputBox("Do You Want One Line per Fixture or Two lines? (1,2)", "1/2 Lines", "1") '06-17-10 
        LinesToPrt = Val(SI) : If LinesToPrt < 1 Or LinesToPrt > 2 Then LinesToPrt = 1 '06-17-10 "N" Then IncludeQty = "N" '02-15-10
        If LinesToPrt = 1 Then ExportExcelProductLines = True '06-17-10 
        'SellCost = SI '02-15-09 
        If frmQuoteRpt.optUnitOrExtended_Extd.Checked = True Then '09-10-09 Extended Prices
            UnitExtd = "E" '02-15-10
        Else ' Unit 
            UnitExtd = "U" '02-15-10
        End If

6000:
        ppv.Doc.Clear() 'Clear the Doc
        Call SetupPrintPreview(FirmName) '09-18-08
        ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
        ' Because we want to show a wide table, we adjust the properties of the preview accordingly and hide all margins.           
        ppv.C1PrintPreviewControl1.PreviewPane.HideMarginsState = C1.Win.C1Preview.HideMarginsFlags.All
        ' Do not allow the user to show margins. 
        ppv.C1PrintPreviewControl1.PreviewPane.HideMargins = C1.Win.C1Preview.HideMarginsFlags.None
        ' Set padding between pages with hidden margins to 0, so that no gap is visible.          
        ppv.C1PrintPreviewControl1.PreviewPane.PagesPaddingSmall = New Size(0, 0)
        ' Set the zoom mode.
        ppv.C1PrintPreviewControl1.PreviewPane.ZoomMode = C1.Win.C1Preview.ZoomModeEnum.PageWidth

        GoTo 6310 'Skip Logo
        'UserPathImages = CurDir() & "\IMAGES\"'02-15-10
        Dim LogoName As String = UserPathImages & "LOGO.BMP" '02-15-10"C:\SAW8\LOGO.BMP" '09-18-08
        Dim ra1 As New RenderArea()
        Dim ri1 As New RenderImage(New Bitmap(LogoName))
        ' set the width of the text to 3 times the width of the page 'ri1.Width = "page*3"
        ri1.Width = "auto"
        ri1.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-15-08
        ra1.Children.Add(ri1)
        ra1.Width = "auto"
        ra1.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-15-08
        doc.Body.Children.Add(ra1)
6310:
        RT = New C1.C1Preview.RenderTable
        RT.Style.GridLines.All = LineDef.Default
        RT.CellStyle.Padding.Left = "1mm" '12-13-12
        RT.CellStyle.Padding.Right = "1mm" '12-13-12
6315:
        frmShowHideGrid.tgShow.SetDataBinding(table, "")
        'If dsGrid Is Nothing Then Exit Sub

        Dim TGNameStr As String = "" 'Documentation Set Up a String of Names
        Dim TGWidthStr As String = "" 'Set Up a String of Widths

        MaxCol = frmQuoteRpt.tgln.Splits(0).DisplayColumns.Count - 1

        'Header 
        RT.Rows.Insert(0, 1) '01-16-09 Insert Header
        RT.RowGroups(0, 1).PageHeader = True
        RT.RowGroups(0, 1).Style.BackColor = AntiqueWhite 'Color.Beige
        RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        'No RT.RowGroups(0, 1).Header = TableHe'12-04-10 & "  UserID = " & UserID 
        RT.Cells(0, 0).Text = frmQuoteRpt.txtSortSeqV.Text & "  UserID = " & UserID & " Report Date = " & VB6.Format(Now, "Short Date") & Space(4) & FirmName & Space(8) & "Page [PageNo] of [PageCount]     *" '07-02-09
        RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        RT.Cells(0, 0).Style.BackColor = AntiqueWhite 'Color.Beige
        RT.Cells(0, 0).Style.FontSize = 14
        RT.Cells(1, 0).Text = frmQuoteRpt.txtSortSeqCriteria.Text '07-02-09frmProjRpt.txtPrimarySortSeq.Text & " " & frmProjRpt.txtSecondarySort.Text '07-01-09
        RT.Cells(1, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        RT.Cells(1, 0).Style.BackColor = LemonChiffon
        RT.Cells(1, 0).Style.FontSize = 12 '06-30-09
        RT.Cells(1, 0).SpanCols = RT.Cols.Count '/ 2 '12-30-08
        RT.Cells(1, 0).SpanRows = 1 '01-16-09
        RT.Cells(0, 0).SpanCols = RT.Cols.Count '/ 2 '12-30-08
        RT.Cells(0, 0).SpanRows = 1 '01-16-09
        RT.Width = "auto"
        RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
        RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular) '05-03-10 JH
        RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
        doc.Body.Children.Add(RT) : RT = New C1.C1Preview.RenderTable : RC = 0
        RT.Style.GridLines.All = LineDef.Default
        RT.CellStyle.Padding.Left = "1mm" '12-13-12
        RT.CellStyle.Padding.Right = "1mm" '12-13-12
        Dim Headertmp As String
        PrtCols = 0   'Print Column Headers
        frmQuoteRpt.tgln.MoveFirst()
        HeaderTxt = New String(" "c, 15) & "Quote Product History" & " - "
        HeaderTxt = HeaderTxt & "Catalog Number Total Monthly Usage Report"
        'Dim PrcColArray() As String = {"", "C", "S", "P", "1"} '{C', 'S', '%', '1'}
        '09-13-09 
        'Dim HdrColArray() As String = {sStartDate, "-JAN-", "-FEB-", "-MAR-", "-APR-", "-MAY-", "-JUN-", "-JUL-", "-AUG-", "-SEP-", "-OCT-", "-NOV-", "-DEC-", "-YTD-"}
        ' A = SI ' s String 'a = Array.parse("['item1', 'item2', 'item3']")
        'ExportExcelProductLines = True '02-15-10
        I = 0 '06-17-10 
        If ExportExcelProductLines = True Then StartPrtCols = 2 '06-17-10 
        For PrtCols = StartPrtCols To StartPrtCols + 13 '02-13-10
            If (PrtCols = 0 And LinesToPrt = 1) Or (ExportExcelProductLines = True And PrtCols = 2) Then '06-17-10 
                RT.Cells(RC, 0).Text = "MFG"
                RT.Cols(0).Width = ".5in"
                RT.Cells(RC, 0).Style.TextAlignHorz = AlignHorzEnum.Center
                RT.Cells(RC, 1).Text = "Description"
                RT.Cols(1).Width = "4in"
                RT.Cells(RC, 1).Style.TextAlignHorz = AlignHorzEnum.Left
                'I = 2
            End If
            RT.Cells(RC, PrtCols + I).Text = HdrColArray((PrtCols + I) - StartPrtCols)
            RT.Cols(PrtCols + I).Width = ".7in"
            RT.Cells(RC, PrtCols + I).Style.TextAlignHorz = AlignHorzEnum.Center
            'Debug.Print(RT.Cells(RC, PrtCols + I).Text)
            'RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
        Next
        RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Left
        'Debug.Print(RT.Cols.Count)
        'Call RTColSize(RT, MaxCol, TgWidth) '02-03-09 
        'RT.Width = "auto"
        'RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
        'doc.Body.Children.Add(RT) '
        'Header End 'Done with Headings
        'RT = New C1.C1Preview.RenderTable
        RC += 1
        RT.Style.GridLines.All = LineDef.Default
        frmQuoteRpt.tgln.UpdateData()
        Dim RowCnt As Integer = 0 'Major Print Loop
        'Zero out Totals
        For I As Integer = 1 To 12 : SubTotMonthP(I) = 0 : SubTotMonthQ(I) = 0 : GtSubTotMonthP(I) = 0 : GtSubTotMonthQ(I) = 0 : Next '02-14-10
        SubTotYTDP = 0 : SubTotYTDQ = 0
        MaxRow = frmQuoteRpt.QuoteLinesBindingSource.Count - 1
        If MaxRow > -1 Then  Else MsgBox("No Line Item Records Selected. Please Try Again") : Exit Sub '09-08-09 

StartPrintLoop:
        For Row = 0 To MaxRow
            frmQuoteRpt.tgln.Row = Row
            If TotYTDQ = 9999999 Then PrevLev1 = "**GRAND TOTAL" : frmQuoteRpt.tgln(Row, "MFG").text = "TOTL" : frmQuoteRpt.tgln(Row, "Description").text = "  GRAND TOTAL " : GoTo PrintLoop2
            If Row = 0 Then PrevLev1 = frmQuoteRpt.tgln(Row, "MFG").ToString() & frmQuoteRpt.tgln(Row, "Description").ToString() '09-09-09 '09-08-09
            If PrevLev1 <> frmQuoteRpt.tgln(Row, "MFG").ToString() & frmQuoteRpt.tgln(Row, "Description").ToString() Then '09-08-09
                'Debug.Print(frmQuoteRpt.tg(Row, "MFG").ToString() & frmQuoteRpt.tg(Row, "Description").ToString())

PrintLoop2:     'Print Totals for each MFG/Desc Major Break  'THDG = "**TOTAL " & PrevLev1
                For PrtCols = 1 To 13 : RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right : Next
                RT.Cells(RC, 0).Text = frmQuoteRpt.tgln(Row, "MFG").ToString()
                RT.Cells(RC, 1).Text = frmQuoteRpt.tgln(Row, "Description").ToString() '09-09-09
                If TotYTDQ = 9999999 Then RT.Cells(RC, 0).Text = "GRAND" : RT.Cells(RC, 1).Text = "TOTAL for REPORT"
                If ExportExcelProductLines = True Then '02-15-10
                    StartPrtCols = 2
                Else
                    StartPrtCols = 0
                End If
                If TotYTDQ = 9999999 Then
                    RT.Cells(RC, 0).Text = "****" : RT.Cells(RC, 1).Text = " GRAND TOTAL FOR REPORT" ' "  GRAND TOTAL "
                    For PrtCols = 0 To 13 : RT.Cells(RC, PrtCols).Style.BackColor = AntiqueWhite : Next 'Beige
                End If
                'RT.Cols(RC).Style.BackColor = Color.
                RT.Cols(0).Width = ".7in"
                If ExportExcelProductLines = True Then '02-15-10
                    RT.Cols(1).Width = "4in"
                Else
                    RT.Cells(RC, 1).SpanCols = 7
                End If

                RT.Cells(RC, 1).Style.TextAlignHorz = AlignHorzEnum.Left
                RT.Cells(RC, 0).Style.BackColor = AntiqueWhite : RT.Cells(RC, 1).Style.BackColor = AntiqueWhite 'Beige
                'doc.Body.Children.Add(RT) '
                'RT = New C1.C1Preview.RenderTable : RC = 0 : RT.Style.GridLines.All = LineDef.Default
                'RT.Cells(RC, 1).SpanCols = 4
                If ExportExcelProductLines = True Then '02-15-10
                    StartPrtCols = 2
                Else
                    StartPrtCols = 0
                    RC += 1
                End If
                If IncludeQty = "N" Then GoTo NoQty '02-15-10 

                RT.Cells(RC, 0).Text = frmQuoteRpt.tgln(Row, "MFG").ToString()
                RT.Cells(RC, 1).Text = frmQuoteRpt.tgln(Row, "Description").ToString() '02-15-10
                RT.Cells(RC, StartPrtCols).Text = " Qty"
                For PrtCols = StartPrtCols + 1 To StartPrtCols + 12 '02-15-10
                    RT.Cells(RC, PrtCols).Text = Format(SubTotMonthQ(PrtCols - StartPrtCols), "#######0") '02-14-10
                    RT.Cols(PrtCols).Width = "1in"
                Next
                '02-15-10 
                If TotYTDQ = 9999999 Then For PrtCols = 0 To 13 : RT.Cells(RC, PrtCols).Style.BackColor = AntiqueWhite : Next 'Beige
                RT.Cells(RC, StartPrtCols + 13).Text = Format(SubTotYTDQ, "#######0") '02-14-10
                RT.Cols(0).Width = ".7in" : RT.Cols(13).Width = "1in"
                RC += 1
NoQty:
                RT.Cells(RC, 0).Text = frmQuoteRpt.tgln(Row, "MFG").ToString()
                RT.Cells(RC, 1).Text = frmQuoteRpt.tgln(Row, "Description").ToString() '02-15-10
                RT.Cells(RC, StartPrtCols).Text = " Sales"
                If SellCost = "C" Then RT.Cells(RC, StartPrtCols).Text = " Comm" '02-15-10
                If DIST And SellCost = "C" Then RT.Cells(RC, StartPrtCols).Text = " Cost" '02-15-10
                For PrtCols = StartPrtCols + 1 To StartPrtCols + 12 '02-15-10For PrtCols = 1 To 12
                    RT.Cells(RC, PrtCols).Text = Format(SubTotMonthP(PrtCols - StartPrtCols), "#######0") '02-14-10

                    RT.Cols(PrtCols).Width = ".7in"
                Next
                RT.Cells(RC, StartPrtCols + 13).Text = Format(SubTotYTDP, "#######0") '02-14-10
                RT.Cols(StartPrtCols).Width = ".7in" : RT.Cols(StartPrtCols + 13).Width = ".7in"
                If TotYTDQ = 9999999 Then For PrtCols = 0 To 13 : RT.Cells(RC, PrtCols).Style.BackColor = AntiqueWhite : Next 'Beige
                ' 'zero Totals Lev2
                PrevLev1 = frmQuoteRpt.tgln(Row, "MFG").ToString() & frmQuoteRpt.tgln(Row, "Description").ToString() '09-09-09 CurrLev1
                For I As Integer = 1 To 12 : GtSubTotMonthP(I) = GtSubTotMonthP(I) + SubTotMonthP(I) : GtSubTotMonthQ(I) = GtSubTotMonthQ(I) + SubTotMonthQ(I) : Next '02-14-10
                For I As Integer = 1 To 12 : SubTotMonthP(I) = 0 : SubTotMonthQ(I) = 0 : Next
                SubTotYTDP = 0 : SubTotYTDQ = 0
                RC += 1
                'X = "ZeroLevels" '02-09-09
                'Call(TotalsCalc(X, B, TotalLevels.TotLv1)) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            End If 'End Print Totals for each MFG/Desc Major Break  'THDG = "**TOTAL " & PrevLev1
            'Done with Subtotal Break
            Dim LineP1 As Decimal 'Each Record may go to a different month
            LineQty = CDec(Val(frmQuoteRpt.tgln(Row, "Qty").ToString()))

            Dim APrice As String = frmQuoteRpt.tgln(Row, "Sell").ToString
            Dim UnitOfM As String = frmQuoteRpt.tgln(Row, "UM").ToString
            Dim UnitMeas As Decimal = 1
            Dim UnitMeaStr As String = UnitMeaSet(APrice, UnitMeas, frmQuoteRpt.tgln.Columns("UM").Text) '' C = Hundreds M = Thousands FT =Feet '01-28-04
            If SellCost = "S" Then '02-15-10
                LineP1 = CDec(Val(frmQuoteRpt.tgln(Row, "Sell").ToString())) * LineQty / UnitMeas '02-14-10
                If UnitExtd = "U" Then LineP1 = CDec(Val(frmQuoteRpt.tgln(Row, "Sell").ToString()))
            Else
                If DIST Then '02-15-10
                    LineP1 = CDec(Val(frmQuoteRpt.tgln(Row, "Cost").ToString())) * LineQty / UnitMeas '02-14-10
                    If UnitExtd = "U" Then LineP1 = CDec(Val(frmQuoteRpt.tgln(Row, "Cost").ToString()))
                Else
                    LineP1 = CDec(Val(frmQuoteRpt.tgln(Row, "Comm-$").ToString())) * LineQty / UnitMeas '02-14-10
                    If UnitExtd = "U" Then LineP1 = CDec(Val(frmQuoteRpt.tgln(Row, "Comm-$").ToString()))
                End If
            End If
            Mth = Val(Format(frmQuoteRpt.tgln(Row, "EntryDate"), "MM"))
            'Debug.Print(frmQuoteRpt.tg(Row, "EntryDate").ToString)
            For I = 1 To 12 'Find the month 'Each Record may go to a different month
                If MonthToPrt(I) = Mth Then Mth = I : Exit For
            Next I
            SubTotMonthP(Mth) += LineP1 '02-14-10 Fill This Month For This Record
            TotMonthP(Mth) += LineP1 '=  
            SubTotYTDP += LineP1
            TotYTDP += LineP1
            'Quantity
            SubTotMonthQ(Mth) += LineQty
            TotMonthQ(Mth) += LineQty '= Grand total of Report 
            SubTotYTDQ += LineQty
            If TotYTDQ = 9999999 Then GoTo allDone '01-12-10 
        Next 'Row
        'End Rows*************************************************
        If TotYTDQ <> 9999999 Then  Else GoTo allDone '01-12-10 
        'Print Totals SubTotMonthP
        SubTotYTDP = 0 : SubTotYTDQ = 0 '02-14-10
        For I = 1 To 12 'Move Grand Total to SubTotMonthP(I) to Print 
            GtSubTotMonthP(I) = GtSubTotMonthP(I) + SubTotMonthP(I) : GtSubTotMonthQ(I) = GtSubTotMonthQ(I) + SubTotMonthQ(I) '02-15-10
            SubTotMonthP(I) = GtSubTotMonthP(I) '02-14-10
            SubTotMonthQ(I) = GtSubTotMonthQ(I) '02-14-10
            SubTotYTDP = SubTotYTDP + GtSubTotMonthP(I) '02-14-10
            SubTotYTDQ = SubTotYTDQ + GtSubTotMonthQ(I) '02-14-10
        Next I
        TotYTDQ = 9999999 'set one time switch
        GoTo StartPrintLoop 'to print Total line
allDone:
        For PrtCols = 1 To 13
            RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
        Next
        RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Left
        RT.Style.GridLines.All = LineDef.Default
        RT.Width = "auto"
        RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
        RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular) '05-03-10 JH
        RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
        doc.Body.Children.Add(RT)
        GoTo ppvshowDoc
ppvshowDoc:
        ppv.C1PrintPreviewControl1.Document = doc
        ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
        ppv.Doc.Generate()
        ppv.Show()
        Exit Sub  '#End
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
10:
    End Sub
    Public Sub PrtRealizationSpreadSheet(ByRef PriceCol As Integer) '05-20-13
        '09-09-09  #Top
        'Always use Sell column
        'Dim SellCost As String = "S" '02-15-10
        'Dim UnitExtd As String = "E" '02-15-10
        'Dim IncludeQty As String = "Y" '02-15-10
        'Dim MC As Integer = 0
        'Dim OldMfg As String = ""
        Dim Mth As Integer = 0
        'Dim QI2Desc As String = ""
        Dim HeaderTxt As String = ""
        On Error Resume Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim TotMonthP(12) As Decimal 'TotMonthP(12),SubTotMonthP(12),GtSubTotMonthP(12) 'SubTotMonthP(Mth) = Total this NCODE
        Dim TotMonthQ(12) As Decimal
        Dim SubTotMonthP(12) As Decimal 'SubTotMonthP(Mth) = Total this NCODE
        Dim SubTotMonthQ(12) As Decimal
        Dim GtSubTotMonthP(12) As Decimal
        Dim GtSubTotMonthQ(12) As Decimal
        Dim MonthToPrt(12) As Int16 '09-14-09 
        Dim SubTotYTDP As Decimal = 0
        Dim SubTotYTDQ As Decimal = 0
        Dim TotYTDP As Decimal = 0
        Dim TotYTDQ As Decimal = 0
        ' Dim LineQty As Decimal = 0
        'Static PriCol As Integer
        'Try '#Top  Convert This sub to QuoteLines
        Dim MaxRow As Single = 0 '09-01-09
        Dim Row As Integer = 0 '09-09-09
        ' Dim LnPrcCol As Short = 0
        'Dim LinesToPrt As Short = 1 '06-17-10 
        Dim sStartDate As String = VB6.Format(frmQuoteRpt.DTPickerStartEntry.Text, "yyyy") '  VB6.Format("01/01/2008", "yyyy-MM-dd") 'Testing
        Dim HdrColArray() As String = {sStartDate, "-JAN-", "-FEB-", "-MAR-", "-APR-", "-MAY-", "-JUN-", "-JUL-", "-AUG-", "-SEP-", "-OCT-", "-NOV-", "-DEC-", "-YTD-"}
        Dim newsStartDate As Date = VB6.Format(frmQuoteRpt.DTPickerStartEntry.Text, "yyyy/MM") '  VB6.Format("01/01/2008", "yyyy-MM-dd") 'Testing
        Dim PrtCols As Int16 = 0
        Dim StartPrtCols As Int16 = 2 '02-15-10
        Dim SingleMFG As String = frmQuoteRpt.txtQutRealCode.Text.Trim '08-20-14 JTC If Not "ALL"
        If frmQuoteRpt.txtSecondarySort.Text = "Spread Sheet by Year" Then '06-24-15 JTC
            GoTo SpreadSheetByYear
        End If
        Resp = MsgBox("A twelve month spreadsheet starting date = " & frmQuoteRpt.DTPickerStartEntry.Text, MsgBoxStyle.YesNoCancel) '05-20-13
        If Resp = vbYes Then  Else Exit Sub
        'If ExportExcelProductLines = True Then '02-15-10
        StartPrtCols = 2 'Else 'StartPrtCols = 0   ' End If
        For I = 1 To 12    'Sets Heading to -Jun-  -Jul- etc for 12 months
            Dim Mon As Integer = newsStartDate.Month ' (DateInterval.Month) '  .AddMonths(I) 'Mon = Mon.DatePart(DateInterval.Month) ', Mon) 'Format(Mon, "mm") 'Mon.Month.ToString
            Dim MonCalc As Integer = I + Mon - 1
            If MonCalc > 12 Then MonCalc = MonCalc - 12
            MonthToPrt(I) = MonCalc
            Dim Monname As String = MonthName(MonCalc, True)
            HdrColArray(I) = "-" & Left(Monname, 3) & "-"
        Next I
        HdrColArray(13) = "-YTD-" '02-15-10
        'If frmQuoteRpt.optoptSalesorCost_Sales.Checked = True Then SellCost = "S" Else SellCost = "C" '02-15-09 
        'SI = Interaction.InputBox("Do You Want to Include the Quantity Row on the Report? (Y,N)", "Quantity Row", "Y") '02-15-10
        'If UCase(SI) = "N" Then IncludeQty = "N" '02-15-10
        ' SI = Interaction.InputBox("Do You Want One Line per Fixture or Two lines? (1,2)", "1/2 Lines", "1") '06-17-10 
        'LinesToPrt = Val(SI) : If LinesToPrt < 1 Or LinesToPrt > 2 Then LinesToPrt = 1 '06-17-10 "N" Then IncludeQty = "N" '02-15-10
        'If LinesToPrt = 1 Then ExportExcelProductLines = True '06-17-10 
        'SellCost = SI '02-15-09 
        'If frmQuoteRpt.optUnitOrExtended_Extd.Checked = True Then '09-10-09 Extended Prices
        'UnitExtd = "E" '02-15-10
        'Else ' Unit 
        ' UnitExtd = "U" '02-15-10
        'End If

6000:
        ppv.Doc.Clear() 'Clear the Doc
        Call SetupPrintPreview(FirmName) '09-18-08
        ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
        ' Because we want to show a wide table, we adjust the properties of the preview accordingly and hide all margins.           
        '05-20-13 ppv.C1PrintPreviewControl1.PreviewPane.HideMarginsState = C1.Win.C1Preview.HideMarginsFlags.All
        '' Do not allow the user to show margins. 
        'ppv.C1PrintPreviewControl1.PreviewPane.HideMargins = C1.Win.C1Preview.HideMarginsFlags.None
        '' Set padding between pages with hidden margins to 0, so that no gap is visible.          
        'ppv.C1PrintPreviewControl1.PreviewPane.PagesPaddingSmall = New Size(0, 0)
        '' Set the zoom mode.
        'ppv.C1PrintPreviewControl1.PreviewPane.ZoomMode = C1.Win.C1Preview.ZoomModeEnum.PageWidth

        GoTo 6310 'Skip Logo
        'UserPathImages = CurDir() & "\IMAGES\"'02-15-10
        'Dim LogoName As String = UserPathImages & "LOGO.BMP" '02-15-10"C:\SAW8\LOGO.BMP" '09-18-08
        'Dim ra1 As New RenderArea()
        'Dim ri1 As New RenderImage(New Bitmap(LogoName))
        '' set the width of the text to 3 times the width of the page 'ri1.Width = "page*3"
        'ri1.Width = "auto"
        'ri1.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-15-08
        'ra1.Children.Add(ri1)
        'ra1.Width = "auto"
        'ra1.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-15-08
        'doc.Body.Children.Add(ra1)
6310:
        RT = New C1.C1Preview.RenderTable
        RT.Style.GridLines.All = LineDef.Default
        RT.CellStyle.Padding.Left = "1mm" '12-13-12
        RT.CellStyle.Padding.Right = "1mm" '12-13-12
6315:
        frmShowHideGrid.tgShow.SetDataBinding(table, "")

        'Header 
        RT.Rows.Insert(0, 1) '01-16-09 Insert Header
        RT.RowGroups(0, 1).PageHeader = True
        RT.RowGroups(0, 1).Style.BackColor = AntiqueWhite 'Color.Beige
        RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        'No RT.RowGroups(0, 1).Header = TableHe'12-04-10 & "  UserID = " & UserID 
        RT.Cells(0, 0).Text = frmQuoteRpt.txtSortSeqV.Text & "  UserID = " & UserID & " Report Date = " & VB6.Format(Now, "Short Date") & Space(4) & FirmName & Space(8) & "Page [PageNo] of [PageCount]     *" '07-02-09
        RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        RT.Cells(0, 0).Style.BackColor = AntiqueWhite 'Color.Beige
        RT.Cells(0, 0).Style.FontSize = 14
        RT.Cells(1, 0).Text = frmQuoteRpt.txtSortSeqCriteria.Text '07-02-09frmProjRpt.txtPrimarySortSeq.Text & " " & frmProjRpt.txtSecondarySort.Text '07-01-09
        RT.Cells(1, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        RT.Cells(1, 0).Style.BackColor = LemonChiffon
        RT.Cells(1, 0).Style.FontSize = 12 '06-30-09
        RT.Cells(1, 0).SpanCols = RT.Cols.Count '/ 2 '12-30-08
        RT.Cells(1, 0).SpanRows = 1 '01-16-09
        RT.Cells(0, 0).SpanCols = RT.Cols.Count '/ 2 '12-30-08
        RT.Cells(0, 0).SpanRows = 1 '01-16-09
        '05-20-13RT.Width = "auto"
        ' RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
        RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular) '05-03-10 JH
        '05-20-13 RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
        doc.Body.Children.Add(RT)

        Dim Headertmp As String
        PrtCols = 0   'Print Column Headers
        frmQuoteRpt.tgr.MoveFirst()
        HeaderTxt = New String(" "c, 15) & "Quote To Spread Sheet" & " - "
        HeaderTxt = HeaderTxt & "Code Total Monthly Dollar Report"
        'Dim PrcColArray() As String = {"", "C", "S", "P", "1"} '{C', 'S', '%', '1'}
        '09-13-09 
        'Dim HdrColArray() As String = {sStartDate, "-JAN-", "-FEB-", "-MAR-", "-APR-", "-MAY-", "-JUN-", "-JUL-", "-AUG-", "-SEP-", "-OCT-", "-NOV-", "-DEC-", "-YTD-"}
        ' A = SI ' s String 'a = Array.parse("['item1', 'item2', 'item3']")
        'ExportExcelProductLines = True '02-15-10
        I = 0 '06-17-10 
        ' If ExportExcelProductLines = True Then StartPrtCols = 2 '06-17-10 
        RT = New C1.C1Preview.RenderTable : RC = 0
        RT.Width = "auto"
        'RT.StretchColumns = StretchTableEnum.LastVectorOnPage
        RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage
        RT.Style.GridLines.All = LineDef.Default
        'RT.CellStyle.Padding.Left = "1mm" '12-13-12
        'RT.CellStyle.Padding.Right = "1mm" '12-13-12
        RT.Cells(RC, 0).Text = "Code"
        RT.Cols(0).Width = ".8in"
        RT.Cells(RC, 0).Style.TextAlignHorz = AlignHorzEnum.Center
        RT.Cells(RC, 1).Text = "Name"
        RT.Cols(1).Width = "2.8in"
        RT.Cells(RC, 1).Style.TextAlignHorz = AlignHorzEnum.Left

        For PrtCols = StartPrtCols To StartPrtCols + 13 '02-13-10
            RT.Cells(RC, PrtCols + I).Text = HdrColArray((PrtCols + I) - StartPrtCols)
            RT.Cols(PrtCols).Width = "1in"
            'Debug.Print(RT.Cells(RC, PrtCols + I).Text)
            RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
        Next
        RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Left
        'Debug.Print(RT.Cols.Count)
        RC += 1
        RT.Style.GridLines.All = LineDef.Default
        frmQuoteRpt.tgr.UpdateData()
        Dim RowCnt As Integer = 0 'Major Print Loop
        'Zero out Totals            'SubTotMonthP(Mth) = Total this NCODE
        For I As Integer = 1 To 12 : SubTotMonthP(I) = 0 : SubTotMonthQ(I) = 0 : GtSubTotMonthP(I) = 0 : GtSubTotMonthQ(I) = 0 : Next '02-14-10
        SubTotYTDP = 0 : SubTotYTDQ = 0 : PrevLev2 = ""
        MaxRow = frmQuoteRpt.QuoteRealLUBindingSource.Count - 1
        If MaxRow > -1 Then  Else MsgBox("No Quote To Records Selected. Please Try Again") : Exit Sub '09-08-09 

StartPrintLoop:
        For Row = 0 To MaxRow
            frmQuoteRpt.tgr.Row = Row
            If TotYTDQ = 9999999 Then CurrLev2 = "LastMfg" : GoTo 6397 '05-20-13 frmQuoteRpt.tgr(Row, "NCode").text = "TOTL" : frmQuoteRpt.tgr(Row, "FirmName").text = "  GRAND TOTAL " : CurrLev2 = "LastMfg" : GoTo 6397 '05-20-13PrevLev1 = "**GRAND TOTAL" :: PrevLev2 = frmQuoteRpt.tgr(Row, "FirmName").text : GoTo PrintLoop2 '05-20-13 
            If Row = 0 Then PrevLev1 = frmQuoteRpt.tgr(Row, "NCode").ToString() : PrevLev2 = frmQuoteRpt.tgr(Row, "FirmName").ToString '05-20-13
6397:       If PrevLev1 <> frmQuoteRpt.tgr(Row, "NCode").ToString() Then '  & frmQuoteRpt.tgr(Row, "Description").ToString() Then '09-08-09
                'Debug.Print(frmQuoteRpt.tgr(Row, "NCode").ToString()) ' & frmQuoteRpt.tg(Row, "Description").ToString())

PrintLoop2:     'Print Totals for each MFG/Desc Major Break  'THDG = "**TOTAL " & PrevLev1
                For PrtCols = 1 To 13 : RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right : Next
                RT.Cells(RC, 0).Text = PrevLev1 '05-20-13 frmQuoteRpt.tgr(Row, "NCode").ToString()
                RT.Cells(RC, 1).Text = PrevLev2 '06-24-15frmQuoteRpt.tgr(Row, "FirmName").ToString() '09-09-09
                If TotYTDQ = 9999999 Then
                    If CurrLev2 <> "LastMfg" Then
                        SubTotYTDP = 0
                        'RT.Cells(RC, 0).Text = "GRAND" : RT.Cells(RC, 1).Text = "TOTAL for REPORT"
                        RT.Cells(RC, 0).Text = "Total" : RT.Cells(RC, 1).Text = " GRAND TOTAL FOR REPORT" ' "  GRAND TOTAL "
                        PrevLev2 = " GRAND TOTAL FOR REPORT" '02-23-15 
                        For PrtCols = 0 To 15 : RT.Cells(RC, PrtCols).Style.BackColor = AntiqueWhite : Next 'Beige
                        'Else
                        '    CurrLev2 = "" '05-20-13 Not LastMfg"
                    End If
                End If
                RT.Cells(RC, 1).Style.TextAlignHorz = AlignHorzEnum.Left
                RT.Cells(RC, 0).Style.BackColor = AntiqueWhite : RT.Cells(RC, 1).Style.BackColor = AntiqueWhite 'Beige
NoQty:          'If TotYTDQ = 9999999 Then GoTo 6400 '05-20-13
                'RT.Cells(RC, 0).Text = PrevLev1 '05-20-13 frmQuoteRpt.tgr(Row, "NCode").ToString()
                RT.Cells(RC, 1).Text = PrevLev2 '06-24-15 frmQuoteRpt.tgr(Row, "FirmName").ToString() '02-15-10
                'RT.Cells(RC, StartPrtCols).Text = " Sales"
6400:
                'If SellCost = "C" Then RT.Cells(RC, StartPrtCols).Text = " Comm" '02-15-10
                'If DIST And SellCost = "C" Then RT.Cells(RC, StartPrtCols).Text = " Cost" '02-15-10
                For PrtCols = StartPrtCols + 1 To StartPrtCols + 12 'SubTotMonthP(Mth) = Total this NCODE '02-15-10For PrtCols = 1 To 12
                    RT.Cells(RC, PrtCols).Text = Format(SubTotMonthP(PrtCols - StartPrtCols), "#######0") '02-14-10
                    If TotYTDQ = 9999999 And CurrLev2 <> "LastMfg" And SingleMFG = "" Then '06-25-15
                        'RT.Cells(RC, PrtCols).Text = Format(GtSubTotMonthP(PrtCols - StartPrtCols), "#######0") '02-14-10
                        SubTotYTDP = SubTotYTDP + SubTotMonthP(PrtCols - StartPrtCols) '06-26-15 GtSubTotMonthP(I) '05-26-15  'SubTotYTDP ='Mth(13) 
                    End If
                Next '                                         'SubTotYTDP ='Mth(13) 
                RT.Cells(RC, StartPrtCols + 13).Text = Format(SubTotYTDP, "#######0") '02-14-10
                If TotYTDQ = 9999999 And CurrLev2 <> "LastMfg" Then
                    RT.Cells(RC, StartPrtCols + 13).Text = Format(TotYTDP, "#######0") '06-25-15
                End If
                'RT.Cols(StartPrtCols + 13).Width = "2.5in" : RT.Cols(StartPrtCols + 13).Width = "3in"
                '06-30-15 JTC                    = ""                         15
                If TotYTDQ = 9999999 And CurrLev2 = "" Then For PrtCols = 0 To 15 : RT.Cells(RC, PrtCols).Style.BackColor = AntiqueWhite : Next 'Beige
                ' 'zero Totals Lev2

                PrevLev1 = frmQuoteRpt.tgr(Row, "NCode").ToString() : PrevLev2 = frmQuoteRpt.tgr(Row, "FirmName").ToString '05-20-13
                '06-26-15 JTC Don't add to last line
                If TotYTDQ = 9999999 And CurrLev2 = "LastMfg" Then                       'SubTotMonthP(Mth) = Total this NCODE
                Else
                    For I As Integer = 1 To 12 : GtSubTotMonthP(I) = GtSubTotMonthP(I) + SubTotMonthP(I) : GtSubTotMonthQ(I) = GtSubTotMonthQ(I) + SubTotMonthQ(I) : Next '02-14-10
                End If

                If TotYTDQ = 9999999 Then '05-20-13
                    If CurrLev2 = "LastMfg" Then
                        CurrLev2 = "" '05-20-13 Not LastMfg"frmQuoteRpt.tgr(Row, "NCode").ToString = "TOTL" Then '05-20-13
                        'Print Totals SubTotMonthP
                        SubTotYTDP = 0 : SubTotYTDQ = 0 '02-14-10
                        For I = 1 To 12 'Move Grand Total to SubTotMonthP(I) to Print 'SubTotMonthP(Mth) = Total this NCODE
                            GtSubTotMonthP(I) = GtSubTotMonthP(I) + SubTotMonthP(I) : GtSubTotMonthQ(I) = GtSubTotMonthQ(I) + SubTotMonthQ(I) '02-15-10
                            SubTotMonthP(I) = GtSubTotMonthP(I) '06-26-15 
                            SubTotMonthQ(I) = GtSubTotMonthQ(I) '02-14-10
                            SubTotYTDP = SubTotYTDP + GtSubTotMonthP(I) '02-14-10 'SubTotYTDP ='Mth(13) 
                            SubTotYTDQ = SubTotYTDQ + GtSubTotMonthQ(I) '02-14-10
                        Next I
                        'If TotYTDQ = 9999999 And CurrLev2 = "" And SingleMFG <> "" Then '06-25-15 Add SingleMFG <> "")06-24-15 JTC Printed Last Total Add to YearTotalArray(
                        '    RC += 1 : GoTo PrintLoop2 '05-20-13 Print Grand total 'SubTotYTDP ='Mth(13) 
                        'ElseIf TotYTDQ = 9999999 And CurrLev2 = "" And SingleMFG = "" Then
                        '    RC += 1 : GoTo PrintLoop2 '05-20-13 Print Grand total
                        'End If
                        'For I As Integer = 1 To 12 : SubTotMonthP(I) = 0 : SubTotMonthQ(I) = 0 : Next '06-24-15 Jtc
                        RC += 1 : GoTo PrintLoop2 '05-20-13 Print Grand total
                    End If
                End If
                For I As Integer = 1 To 12 : SubTotMonthP(I) = 0 : SubTotMonthQ(I) = 0 : Next
                SubTotYTDP = 0 : SubTotYTDQ = 0 'SubTotYTDP ='Mth(13) 
                RC += 1
                If TotYTDQ = 9999999 Then GoTo allDone '05-20-13
            End If 'End Print Totals for each MFG/Desc Major Break  'THDG = "**TOTAL " & PrevLev1
            'Done with Subtotal Break
            Dim APrice As String = frmQuoteRpt.tgr(Row, "Sell").ToString
            Mth = Val(Format(frmQuoteRpt.tgr(Row, "EntryDate"), "MM"))
            'Debug.Print(frmQuoteRpt.tg(Row, "EntryDate").ToString)
            For I = 1 To 12 'Find the month 'Each Record may go to a different month
                If MonthToPrt(I) = Mth Then Mth = I : Exit For
            Next I
            TotMonthP(Mth) += APrice '= Fill This Month For This Line Record 
            SubTotMonthP(Mth) += APrice '02-14-10 Fill This Month total for this NCODE SubTotMonthP(Mth) = Total this NCODE
            SubTotYTDP += APrice 'SubTotYTDP ='Mth(13) 
            TotYTDP += APrice  ' Grand Total Report 
            If TotYTDQ = 9999999 Then GoTo allDone '01-12-10                         '06-25-15
            If (Row = MaxRow And SingleMFG <> "") Then TotYTDQ = 9999999 : CurrLev2 = "LastMfg" : GoTo PrintLoop2 ' = frmQuoteRpt.txtQutRealCode.Text '08-20-14 JTC If Not "ALL"
        Next 'Row
        'End Rows*************************************************
        If TotYTDQ <> 9999999 Then  Else GoTo allDone '01-12-10 
        'If SingleMFG <> "ALL" Then GoTo PrintLoop2 ' = frmQuoteRpt.txtQutRealCode.Text '08-20-18 JTC If Not "ALL"
        'Next I
        TotYTDQ = 9999999 'set one time switch
        GoTo StartPrintLoop 'to print Total line
allDone:
        RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Left
        RT.Style.GridLines.All = LineDef.Default
        RT.CellStyle.Padding.Left = "1mm" '05-20-13
        RT.CellStyle.Padding.Right = "1mm" '05-20-13
        RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
        RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular) '05-03-10 JH
        RT.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Right '06-24-15
        '05-20-13RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
        doc.Body.Children.Add(RT)
        GoTo ppvshowDoc
        ''****************************************************************************************************************************************
SpreadSheetByYear:  ' If frmQuoteRpt.txtSecondarySort.Text = "Spread Sheet by Year" Then '06-22-15 JTC             GoTo SpreadSheetByYear
        'Public Sub PrtSpecifierSpreadSheet(ByRef PriceCol As Integer)
        On Error Resume Next

        Dim mysqlcmd As New MySqlCommand
        '07-07-12 use A,E,L,S,T,X  If RealExtByInfluencePercent = True then Extend by % on Realization Reports by specifier
        '04-15-13 Made Public Dim RealExtByInfluencePercent As Boolean = False '07-07-12
        Resp = MsgBox("A Spreadsheet By Year Starting Date = " & VB6.Format(frmQuoteRpt.DTPickerStartEntry.Text, "yyyy/MM") & " For four Years." & vbCrLf & "Ending Year = " & (VB6.Format(frmQuoteRpt.DTPicker1EndEntry.Text, "yyyy/MM")), MsgBoxStyle.YesNoCancel) '06-26-15 05-20-13
        If Resp = vbYes Then  Else Exit Sub
        'If MajSel = RptMaj.RptSpecCredit Then 'Specifier credit If RealCustomer = True Or RealManufacturer = True Or RealOther = True Or RealArchitect = True Or RealEngineer = True Or RealLtgDesigner = True Or RealSpecifier = True Or RealContractor = True Or RealOther = True Then '02-04-12
        '    Resp = MsgBox("Yes = To multiply Specifier Sales Amounts by Influence Percent or " & vbCrLf & "No = Use actual amounts" & vbCrLf & "If Influence Percent is zero we will use 100%", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, "Realization Sort") '07-07-12
        '    If Resp = vbYes Then
        '        RealExtByInfluencePercent = True '02-05-12 RealExtByInfluencePercent 'OnlyInFluenceGtZero = True Else OnlyInFluenceGtZero = False '02-04-12'02-04-12
        '        frmQuoteRpt.TtxtSortSelV.Text += " Ext By %" '04-17-13
        '    End If
        '    'Debug.Print(frmQuoteRpt.TtxtSortSelV.Text)
        '    If Resp = vbCancel Then RealExtByInfluencePercent = False : Exit Sub '04-15-13
        'End If
        Dim YearToPrt As Int16 = 0 '06-24-15 MonthToPrt
        Dim StartYear As Date = VB6.Format(frmQuoteRpt.DTPickerStartEntry.Text, "yyyy/MM") ' .ToString("YYYY')
        'Dim HdrColArray() As String = {sStartDate, "-JAN-", "-FEB-", "-MAR-", "-APR-", "-MAY-", "-JUN-", "-JUL-", "-AUG-", "-SEP-", "-OCT-", "-NOV-", "-DEC-", "-YTD-"} '' Dim HdrColArray() As String = {"Manufacturer", "Sales/Comm") '03-20-10 , "YR", "-JAN-", "-FEB-", "-MAR-", "-APR-", "-MAY-", "-JUN-", "-JUL-", "-AUG-", "-SEP-", "-OCT-", "-NOV-", "-DEC-", "-YTD-"}
        Dim SecondYear As Date = StartYear.AddYears(1) ' VB6.Format(frmQuoteRpt.DTPicker1EndEntry.Text, "yyyy")
        'SecondYear = StartYear.AddYears(1)
        Dim ThirdYear As Date = StartYear.AddYears(2) 'VB6.Format(frmQuoteRpt.DTPicker1EndEntry.Text, "yyyy")
        Dim ForthYear As Date = StartYear.AddYears(3) 'VB6.Format(frmQuoteRpt.DTPicker1EndEntry.Text, "yyyy")
        Dim EndYear As Date = StartYear.AddYears(4)  ' frmQuoteRpt.DTPicker1EndEntry.Value 'Text, "yyyy")
        'Dim PrtCols As Int16 = 0
        'Dim StartPrtCols As Int16 = 0 '02-15-10
        Dim StrSqlTemp As String = ""
        Dim YearTotalArray(5) As Decimal '07-07-12 five years
        Dim SubTotYearFive As Decimal = 0 '06-24-16
        Dim FirmTotalArray(5) As Decimal '04-18-13 five years
        'Dim HdrColArray() As String = {"200", "201", "201", "201", "201"} '' Dim () As String = {"Manufacturer", "Sales/Comm") '03-20-10 , "YR", "-JAN-", "-FEB-", "-MAR-", "-APR-", "-MAY-", "-JUN-", "-JUL-", "-AUG-", "-SEP-", "-OCT-", "-NOV-", "-DEC-", "-YTD-"}
        Dim Temp As String = "" '04-15-13
        For I = 0 To 5    'Sets Heading to -Jun-  -Jul- etc for 12 months
            If StartYear.AddYears(I) >= EndYear Then Exit For ''" & AQuoteCode & "'
            HdrColArray(I + 2) = Format(StartYear.AddYears(I), "yyyy") & " Sales$" 'SecondYear = StartYear ' (DateInterval.Month) '  .AddMonths(I) 'Mon = Mon.DatePart(DateInterval.Month) ', Mon) 'Format(Mon, "mm") 'Mon.Month.ToString
            Temp = Format(StartYear.AddYears(I), "yyyy") & " Sales$"
            ' HdrColArray(I) = Temp
            'If RealExtByInfluencePercent = True Then '04-17-13 O.ORDERSALESAMT to ORDERSALESAMT
            '    StrSqlTemp += ", SUM(CASE WHEN EXTRACT(YEAR FROM O.ENTRYDATE) = '" & Format(StartYear.AddYears(I), "yyyy") & "' THEN  ORDERSALESAMT END) AS '" & Temp & "'" 'FIRSTYRORDERSALESAMT"
            'Else
            '    StrSqlTemp += ", SUM(CASE WHEN EXTRACT(YEAR FROM O.ENTRYDATE) = '" & Format(StartYear.AddYears(I), "yyyy") & "' THEN  O.ORDERSALESAMT END) AS '" & Temp & "'" 'FIRSTYRORDERSALESAMT"
            'End If
            'If StartYear.AddYears(I) >= EndYear Then Exit For ''" & AQuoteCode & "'
        Next I
7000:
        ppv.Doc.Clear() 'Clear the Doc
        Call SetupPrintPreview(FirmName) '09-18-08
        ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1

        Dim MaxYears As Int16 = I
        '        Dim SaveDateSql As String = strSql 'strSql = " and O.EntryDate >= '" & RptStartYearMoDa & "'  and  O.EntryDate <= '" & RptEndYearMoDa & "' "
        '        Dim strSqlLeft As String = "" '07-07-12
        'SkipExtByPercentNew:  '04-15-13
        '        '04-15-13 Test TMPREPORTS **********************************************************************************************
        '        Dim StartEntryDate As Date = frmQuoteRpt.DTPickerStartEntry.Value '11-01-10 ], 	yyyy'-'MM'-'dd HH':'mm':'ss'Z'Date = Me.DTPickerStartEntry.Value
        '        Dim RptStartYearMoDa As String '11-01-10 = Format(DTPickerStartEntry.Value, "yyyy-MM-dd")
        '        RptStartYearMoDa = StartEntryDate.ToString("yyyy-MM-dd") '11-01-10 This works
        '        Dim EndDate As Date = frmQuoteRpt.DTPicker1EndEntry.Value '11-01-10 RptStartYearMoDa, RptEndYearMoDa
        '        Dim RptEndYearMoDa As String '11-01-10 = Format(DTPicker1EndEntry.Value, "yyyy-MM-dd")
        '        RptEndYearMoDa = EndDate.ToString("yyyy-MM-dd") '11-01-10 This works
        '        If myConnection.State <> ConnectionState.Open Then Call OpenSQL(myConnection)
        '        mysqlcmd.Connection = myConnection
        '        strSql = "DROP TABLE IF EXISTS TMPREPORTS "
        '        mysqlcmd.CommandText = strSql : mysqlcmd.ExecuteNonQuery()
        '        strSql = "CREATE TEMPORARY TABLE TMPREPORTS AS SELECT  O.MFG, O.ORDERSALESAMT, O.ORDERCOMMAMT, O.ORDERNUMBER, O.CUSTCODE, O.JOBNAME, O.ENTRYDATE, O.BRANCHCODE AS SLSBRANCH, PC.NCODE AS SLSCODE, PC.SLSCODE AS SALESMAN, PC.LPSell, UPPER(PC.FirmName) as FirmName FROM ORDERMASTER O INNER JOIN PROJECTCUSTORDERMASTER PCOM ON O.ORDERID = PCOM.ORDERID INNER JOIN PROJECTCUST AS PC ON PC.PROJECTCUSTID = PCOM.PROJECTCUSTID  WHERE O.RelHold = 'R' and  O.EntryDate >= '" & RptStartYearMoDa & "'  and  O.EntryDate <= '" & RptEndYearMoDa & "' " '04-17-13 UPPER(PC.FirmName) as FirmName 04-16-13 and (concat(O.BuySellAB, O.BuySellSR) <> 'BS')  " '4-16-13 Jtc Added (concat(I.BuySellAB, I.BuySellSR)<> 'BS') 
        '        ' strSql = "CREATE TEMPORARY TABLE TMPREPORTS AS SELECT  I.InvoiceID, I.ProcessDate, I.Reference as SLSCode from invoicemaster I WHERE  I.Reference <> '' and I.ProcessDate >= '" & RptStartYearMoDa & "'  and  I.ProcessDate <= '" & RptEndYearMoDa & "'  ORDER BY SLSCode" '04-26-10 
        '        mysqlcmd.CommandText = strSql : Dim Count As Integer = mysqlcmd.ExecuteNonQuery()
        '        'frmQuoteRpt.tgLookup.UpdateData()
        '        strSql = "SELECT O.*" '04-15-13  FROM TMPREPORTS O "
        '        '04-18-13strSql += StrSqlTemp '04-15-13" SUM(CASE WHEN EXTRACT(YEAR FROM O.ENTRYDATE) = '" & Format(StartYear.AddYears(I), "yyyy") & "' THEN  O.ORDERSALESAMT END) AS '" & Temp & "'"
        '        strSql += " FROM TMPREPORTS O ORDER BY FIRMNAME " '04-18-13 ORDER BY FIRMNAME 
        '        'StopCall FillDataSetOrderLookup(frmQuoteRpt, "SortExit", "", "")
        '        'Ng daLookup.Update(ds, "ordermaster")
        '        '05-23-12 End*******************************************************************************
        'EndTemporary:
        '        'End TMPREPORTS ***********************************************************************************************
SkipExtByPercent:  '04-15-13
7300:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' Arrow
        ppv.Doc.Clear() 'Clear the Doc
        Call SetupPrintPreview(FirmName) '09-18-08
        ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
        ' Because we want to show a wide table, we adjust the properties of the preview accordingly and hide all margins.           
        '11-10-10 ppv.C1PrintPreviewControl1.PreviewPane.HideMarginsState = C1.Win.C1Preview.HideMarginsFlags.All
        ' Do not allow the user to show margins. 
        '11-10-10 ppv.C1PrintPreviewControl1.PreviewPane.HideMargins = C1.Win.C1Preview.HideMarginsFlags.None
        ' Set padding between pages with hidden margins to 0, so that no gap is visible.          
        '11-10-10 ppv.C1PrintPreviewControl1.PreviewPane.PagesPaddingSmall = New Size(0, 0)
        ' Set the zoom mode.
        ppv.C1PrintPreviewControl1.PreviewPane.ZoomMode = C1.Win.C1Preview.ZoomModeEnum.PageWidth

        'doc.Body.Children.Add(ra1)
7310:
        RT = New C1.C1Preview.RenderTable : RC = 0 '11-12-10 
        RT.Style.GridLines.All = LineDef.Default
        RT.CellStyle.Padding.Left = "1mm" '12-13-12
        RT.CellStyle.Padding.Right = "1mm" '12-13-12
7315:
        frmShowHideGrid.tgShow.SetDataBinding(table, "")
        'If dsGrid Is Nothing Then Exit Sub

        Dim TGNameStr As String = "" 'Documentation Set Up a String of Names
        Dim TGWidthStr As String = "" 'Set Up a String of Widths

        'MaxCol = frmQuoteRpt.tgLookup.Splits(0).DisplayColumns.Count - 1
        'Dim PrtCols As Int16 = 0

        'Header 
        RT.Rows.Insert(0, 1) '01-16-09 Insert Header
        RT.RowGroups(0, 1).PageHeader = True
        RT.RowGroups(0, 1).Style.BackColor = LemonChiffon '11-12-10
        RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        'No RT.RowGroups(0, 1).Header = TableHeaderEnum.All'12-04-10 & "  UserID = " & UserID 
        RT.Cells(0, 0).Text = "Quote Specifier Spreadsheet Report by Year" & "  UserID = " & UserID & "   Report Date = " & VB6.Format(Now, "Short Date") & Space(4) & FirmName & Space(8) & "Page [PageNo] of [PageCount]     *" '07-02-09
        RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        RT.Cells(0, 0).Style.BackColor = LemonChiffon '11-12-10
        RT.Cells(0, 0).Style.FontSize = 14
        'Fix (RT.Cells(1, 0).Text = frmQuoteRpt.TtxtSortSelV.Text) '07-02-09frmProjRpt.txtPrimarySortSeq.Text & " " & frmProjRpt.txtSecondarySort.Text '07-01-09
        RT.Cells(1, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        RT.Cells(1, 0).Style.BackColor = LemonChiffon '11-12-10
        RT.Cells(1, 0).Style.FontSize = 12 '06-30-09
        RT.Cells(1, 0).SpanCols = RT.Cols.Count '/ 2 '12-30-08
        RT.Cells(1, 0).SpanRows = 1 '01-16-09
        RT.Cells(0, 0).SpanCols = RT.Cols.Count '/ 2 '12-30-08
        RT.Cells(0, 0).SpanRows = 1 '01-16-09
        '11-13-10 RT.Width = "auto"
        RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
        doc.Body.Children.Add(RT) '
        RT = New C1.C1Preview.RenderTable
        RC = 0
        RT.Style.GridLines.All = LineDef.Default
        RT.CellStyle.Padding.Left = "1mm" '12-13-12
        RT.CellStyle.Padding.Right = "1mm" '12-13-12
        ' Dim Headertmp As String
        PrtCols = 0   'Print Column Headers
        'Fix frmQuoteRpt.tgLookup.MoveFirst()
        HeaderTxt = "Specifier Sales by Year" ' New String(" "c, 15) & "Quote Product History" & " - "
        'HeaderTxt = HeaderTxt & "Catalog Number Total Monthly Usage Report"
        'Dim PrcColArray() As String = {"", "C", "S", "P", "1"} '{C', 'S', '%', '1'}
        '09-13-09 
        'Dim HdrColArray() As String = {sStartDate, "-JAN-", "-FEB-", "-MAR-", "-APR-", "-MAY-", "-JUN-", "-JUL-", "-AUG-", "-SEP-", "-OCT-", "-NOV-", "-DEC-", "-YTD-"}
        ' A = SI ' s String 'a = Array.parse("['item1', 'item2', 'item3']")Dim HdrColArray() As String = {sStartDate, "-JAN-", "-FEB-", "-MAR-", "-APR-", "-MAY-", "-JUN-", "-JUL-", "-AUG-", "-SEP-", "-OCT-", "-NOV-", "-DEC-", "-YTD-"}
        'MFG, ORDERNUMBER, CUSTCODE, JOBNAME, ENTRYDATE, SLSBRANCH, FIRSTYRORDERSALESAMT, SECONDYRORDERSALESAMT, THIRDYRORDERSALESAMT, SLSCODE, SALESMAN, FIRMNAME
        'SALESMAN, SLSCODE, FIRMNAME,  
        ' PrtCols = 0 To 3 + MaxYears - 1
        RT.Cols(1).Width = "2.6in" '06-25-15
        PrtCols = 0 : RT.Cells(RC, 0).Text = "Firm Code" : RT.Cols(PrtCols).Width = ".9in" : RT.Cells(RC, PrtCols).Style.TextAlignHorz = AlignHorzEnum.Center
        PrtCols = 1 : RT.Cells(RC, 1).Text = "FirmName" : RT.Cols(PrtCols).Width = "2.6in" : RT.Cells(RC, PrtCols).Style.TextAlignHorz = AlignHorzEnum.Center
        ' PrtCols = 2 : RT.Cells(RC, 2).Text = "FirmName" : RT.Cols(PrtCols).Width = "4in" : RT.Cells(RC, PrtCols).Style.TextAlignHorz = AlignHorzEnum.Center
        For PrtCols = 2 To 2 + MaxYears '06-30-15 - 1
            RT.Cells(RC, PrtCols).Text = HdrColArray(PrtCols) : RT.Cols(PrtCols).Width = "1.3in" : RT.Cells(RC, PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
            'Debug.Print(HdrColArray(PrtCols)) ' RT.Cells(RC, PrtCols).Text)
        Next
        RT.Cells(RC, 6).Text = "4 Yr Total"
        'If PrtCols > 2 Then RT.Cells(RC, PrtCols).Text = HdrColArray(MaxYears - 8 + PrtCols) : RT.Cols(PrtCols).Width = "1.5in" : RT.Cells(RC, PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
        'Debug.Print(MaxYears - 6 + PrtCols)
        'Next
        'RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Left
        RC += 1
        RT.Style.GridLines.All = LineDef.Default
        'frmQuoteRpt.tgLookup.UpdateData()
        RowCnt = 0 'Major Print Loop
        Row = 0 'Major Print Loop
        'Zero out Totals
        'For I As Integer = 1 To 12 : SubTotMonthP(I) = 0 : SubTotMonthQ(I) = 0 : GtSubTotMonthP(I) = 0 : GtSubTotMonthQ(I) = 0 : Next '02-14-10
        'SubTotYTDP = 0 : SubTotYTDQ = 0
        '04-16-13 out Dim MaxRow As Integer = frmQuoteRpt.OrdermasterBindingSource.Count '05-21-12 No -1
        '04-16-13 out If MaxRow < 1 Then MsgBox("No Line Item Records Selected. Please Try Again") : Exit Sub '09-08-09 
        '04-18-13 Calc my own totals no Group By *****************************************************************************
        Dim Prevlev() As String '06-22-15 JTC
        Dim Year As String = ""
        Dim FirmNames As String = ""
        Dim Salesman As String = ""
        Dim SLSCode As String = ""
        MaxRow = frmQuoteRpt.QuoteRealLUBindingSource.Count - 1
        If MaxRow > -1 Then  Else MsgBox("No Quote To Records Selected. Please Try Again") : Exit Sub '09-08-09 

StartPrintLoopYear:
        For Row = 0 To MaxRow
            frmQuoteRpt.tgr.Row = Row
            If TotYTDQ = 9999999 Then CurrLev2 = "LastMfg" : GoTo 7397 '05-20-13 frmQuoteRpt.tgr(Row, "NCode").text = "TOTL" : frmQuoteRpt.tgr(Row, "FirmName").text = "  GRAND TOTAL " : CurrLev2 = "LastMfg" : GoTo 6397 '05-20-13PrevLev1 = "**GRAND TOTAL" :: PrevLev2 = frmQuoteRpt.tgr(Row, "FirmName").text : GoTo PrintLoop2 '05-20-13 
            If Row = 0 Then PrevLev1 = frmQuoteRpt.tgr(Row, "NCode").ToString() : PrevLev2 = frmQuoteRpt.tgr(Row, "FirmName").ToString '05-20-13
           
7397:       If PrevLev1 <> frmQuoteRpt.tgr(Row, "NCode").ToString() Then '  & frmQuoteRpt.tgr(Row, "Description").ToString() Then '09-08-09
                'Debug.Print(frmQuoteRpt.tgr(Row, "NCode").ToString()) ' & frmQuoteRpt.tg(Row, "Description").ToString())

PrintLoop2Year:  'Print Totals for each MFG/Desc Major Break  'THDG = "**TOTAL " & PrevLev1

StartPrintLoop2Year:
                'Print Totals for each MFG/Desc Major Break  'THDG = "**TOTAL " & PrevLev1
                For PrtCols = 1 To 6 : RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right : Next '06-29-15 JTC Fix Col 6 
                RT.Cells(RC, 0).Text = PrevLev1 '05-20-13 frmQuoteRpt.tgr(Row, "NCode").ToString()
                RT.Cells(RC, 1).Text = PrevLev2 '06-24-15 JTC frmQuoteRpt.tgr(Row, "FirmName").ToString() '09-09-09

                If TotYTDQ = 9999999 Then
                    If CurrLev2 <> "LastMfg" Then
                        'RT.Cells(RC, 0).Text = "GRAND" : RT.Cells(RC, 1).Text = "TOTAL for REPORT"
                        RT.Cells(RC, 0).Text = "Total" : RT.Cells(RC, 1).Text = " GRAND TOTAL FOR REPORT" ' "  GRAND TOTAL "
                        PrevLev2 = " GRAND TOTAL FOR REPORT" '02-23-15 
                        For PrtCols = 0 To 7 : RT.Cells(RC, PrtCols).Style.BackColor = AntiqueWhite : Next 'Beige
                        'Else
                        '    CurrLev2 = "" '05-20-13 Not LastMfg"
                    End If
                End If
                RT.Cells(RC, 1).Style.TextAlignHorz = AlignHorzEnum.Left
                RT.Cells(RC, 0).Style.BackColor = AntiqueWhite : RT.Cells(RC, 1).Style.BackColor = AntiqueWhite 'Beige
NoQtyYear:      'If TotYTDQ = 9999999 Then GoTo 6400 '05-20-13
                'RT.Cells(RC, 0).Text = PrevLev1 '05-20-13 frmQuoteRpt.tgr(Row, "NCode").ToString()
                'RT.Cells(RC, 1).Text = frmQuoteRpt.tgr(Row, "FirmName").ToString() '02-15-10
                RT.Cells(RC, 1).Text = PrevLev2 '06-24-15
                'RT.Cells(RC, StartPrtCols).Text = " Sales"
7400:
                ''Test Code Comment Out ???????????????????????????????????????????????????????????????????
                'SLSCode = frmQuoteRpt.tgr(Row, "NCode").ToString().Trim
                'If SLSCode = "COLLIN" Or SLSCode = "ECKA" Or SLSCode = "GENELY" Or SLSCode = "HIBB" Or SLSCode = "INGLE" Or SLSCode = "UNITED" Then
                '    PrevLev1 = SLSCode '= "COLLIN" Then PrevLev1 = "COLLIN"
                'Else
                '    GoTo GetNextRecord '06-26-15
                '    'Test Code Comment Out ????????????????????????????????????????????????????????????????
                '    RT.Cells(RC, 0).Text = "AAA"
                '    RT.Cells(RC, 1).Text = "AAA Architects"
                '    RT.Cells(RC, 0).Text = "ECKERT"
                '    RT.Cells(RC, 1).Text = "ECKERT Engineering"
                '    RT.Cells(RC, 0).Text = "GLENDA"
                '    RT.Cells(RC, 1).Text = "GLENDA Ltg. Design"
                '    RT.Cells(RC, 0).Text = "HENDER"
                '    RT.Cells(RC, 1).Text = "Henderson Contractors"
                '    RT.Cells(RC, 0).Text = "INGLE"
                '    RT.Cells(RC, 1).Text = "INGLE & Associates"
                '    RT.Cells(RC, 0).Text = "URBAN"
                '    RT.Cells(RC, 1).Text = "URBAN Architects"

                'End If

                ''Test Code Comment Out ????????????????????????????????????????????????????????????????
                'If SellCost = "C" Then RT.Cells(RC, StartPrtCols).Text = " Comm" '02-15-10
                'If DIST And SellCost = "C" Then RT.Cells(RC, StartPrtCols).Text = " Cost" '02-15-10
                Year = VB6.Format(frmQuoteRpt.tgr(Row, "EntryDate"), "yyyy")
                '06-26-15 JTC Print Every Yeat
                For PrtCols = 2 To 2 + MaxYears - 1
                    Year = Left(HdrColArray(PrtCols), 4) '06-25-15
                    If Year = Left(HdrColArray(PrtCols), 4) Then
                        RT.Cells(RC, PrtCols).Text = Format(FirmTotalArray(PrtCols), "#######0")
                        If TotYTDQ = 9999999 And CurrLev2 <> "LastMfg" Then
                            'FirmTotalArray(PrtCols - 3) = FirmTotalArray(PrtCols - 3) + Val(APrice) '06-24-15 JTCdrOrd.OrderSalesAmt
                            RT.Cells(RC, PrtCols).Text = Format(YearTotalArray(PrtCols), "#######0")
                            ' YearTotalArray(PrtCols - 3) = YearTotalArray(PrtCols - 3) + Val(APrice) '06-24-15 JTC drOrd.OrderSalesAmt
                        End If
                    End If
                Next
                For I = 2 To 5 : SubTotYearFive = SubTotYearFive + FirmTotalArray(I) : Next 'Add 4 years
                If TotYTDQ = 9999999 And CurrLev2 <> "LastMfg" Then
                    For I = 2 To 5 : SubTotYearFive = SubTotYearFive + YearTotalArray(I) : Next 'Add 4 years
                End If
                RT.Cells(RC, 6).Text = Format(SubTotYearFive, "#######0") : SubTotYearFive = 0
                '06-26-15 If TotYTDQ = 9999999 And CurrLev2 = "LastMfg" And SingleMFG = "" Then '06-25-15 Add SingleMFG <> "")06-24-15 JTC Printed Last Total Add to YearTotalArray(
                'For PrtCols = 2 To 2 + MaxYears - 1 : YearTotalArray(PrtCols) = YearTotalArray(PrtCols) + FirmTotalArray(PrtCols) : Next
                'End If
                For PrtCols = 2 To 2 + MaxYears - 1 : FirmTotalArray(PrtCols) = 0 : Next
                'For PrtCols = StartPrtCols + 1 To StartPrtCols + 12 '02-15-10For PrtCols = 1 To 12
                '    RT.Cells(RC, PrtCols).Text = Format(SubTotMonthP(PrtCols - StartPrtCols), "#######0") '02-14-10
                '    'RT.Cols(PrtCols).Width = "2.5in" '05-20-13
                'Next
                'RT.Cells(RC, StartPrtCols + 13).Text = Format(SubTotYTDP, "#######0") '02-14-10
                'RT.Cols(StartPrtCols + 13).Width = "2.5in" : RT.Cols(StartPrtCols + 13).Width = "3in"
                '06-30-15 JTC                    = ""                          7
                If TotYTDQ = 9999999 And CurrLev2 = "" Then For PrtCols = 0 To 7 : RT.Cells(RC, PrtCols).Style.BackColor = AntiqueWhite : Next 'Beige
                ' 'zero Totals Lev2
                PrevLev1 = frmQuoteRpt.tgr(Row, "NCode").ToString() : PrevLev2 = frmQuoteRpt.tgr(Row, "FirmName").ToString '05-20-13
                'For I As Integer = 1 To 5 : FirmTotalArray(I) = 0 :Next'06-24-15  GtSubTotMonthP(I) = GtSubTotMonthP(I) + SubTotMonthP(I) : GtSubTotMonthQ(I) = GtSubTotMonthQ(I) + SubTotMonthQ(I) : Next '02-14-10
                If TotYTDQ = 9999999 Then '05-20-13
                    If CurrLev2 = "LastMfg" Then
                        CurrLev2 = "" '05-20-13 Not LastMfg"frmQuoteRpt.tgr(Row, "NCode").ToString = "TOTL" Then '05-20-13
                        'Print Totals SubTotMonthP SubTotMonthP(Mth) = Total this NCODE
                        'SubTotYTDP = 0 : SubTotYTDQ = 0 '02-14-10
                        'For I = 1 To 12 'Move Grand Total to SubTotMonthP(I) to Print 
                        '    GtSubTotMonthP(I) = GtSubTotMonthP(I) + SubTotMonthP(I) : GtSubTotMonthQ(I) = GtSubTotMonthQ(I) + SubTotMonthQ(I) '02-15-10
                        '    SubTotMonthP(I) = GtSubTotMonthP(I) '02-14-10
                        '    SubTotMonthQ(I) = GtSubTotMonthQ(I) '02-14-10
                        '    SubTotYTDP = SubTotYTDP + GtSubTotMonthP(I) '02-14-10
                        '    SubTotYTDQ = SubTotYTDQ + GtSubTotMonthQ(I) '02-14-10
                        'Next I
                        ': For I As Integer = 1 To 12 : SubTotMonthP(I) = 0 : SubTotMonthQ(I) = 0 : Next '06-24-15 Jtc
                        'For PrtCols = 2 To 2 + MaxYears - 2 : FirmTotalArray(PrtCols) = 0 : Next '06-24-15

                        RC += 1 : GoTo PrintLoop2Year '05-20-13 Print Grand total
                    End If
                End If '                       SubTotMonthP(Mth) = Total this NCODE
                'For I As Integer = 1 To 12 : SubTotMonthP(I) = 0 : SubTotMonthQ(I) = 0 : Next
                'SubTotYTDP = 0 : SubTotYTDQ = 0
                RC += 1
                If TotYTDQ = 9999999 Then GoTo allDone '05-20-13
            End If 'End Print Totals for each MFG/Desc Major Break  'THDG = "**TOTAL " & PrevLev1
            'Equal Done with Subtotal Break
            Dim APrice As String = frmQuoteRpt.tgr(Row, "Sell").ToString
            Year = VB6.Format(frmQuoteRpt.tgr(Row, "EntryDate"), "yyyy")
            For PrtCols = 2 To 2 + MaxYears - 1
                If Year = Left(HdrColArray(PrtCols), 4) Then
                    FirmTotalArray(PrtCols) = FirmTotalArray(PrtCols) + Val(APrice) '06-24-15 JTCdrOrd.OrderSalesAmt
                    YearTotalArray(PrtCols) = YearTotalArray(PrtCols) + Val(APrice) '06-24-15 JTC drOrd.OrderSalesAmt
                End If
            Next
GetNextRecord:  '06-26-15
            If TotYTDQ = 9999999 Then GoTo allDone2 '01-12-10                '06-25-15 CurrLev2 = "LastMfg"
            If (Row = MaxRow And SingleMFG <> "") Then TotYTDQ = 9999999 : CurrLev2 = "LastMfg" : GoTo PrintLoop2Year ' = frmQuoteRpt.txtQutRealCode.Text '08-20-14 JTC If Not "ALL"
        Next 'Row
        If TotYTDQ <> 9999999 Then  Else GoTo allDone2
        TotYTDQ = 9999999 'set one time switch
        GoTo StartPrintLoopYear 'to print Last Ncode then Print Total line
allDone2:
        For PrtCols = 3 To 3 + MaxYears - 1
            RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
        Next
        RT.Style.GridLines.All = LineDef.Default
        RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
        doc.Body.Children.Add(RT)
ppvshowDoc:
        ppv.C1PrintPreviewControl1.Document = doc
        ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
        ppv.Doc.Generate()
        ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        ppv.Show() '12-06-12 JTc Moved down
        ppv.BringToFront()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
10:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow ' WaitCursor ' Arrow '04-15-11


        GoTo ppvshowDoc2
ppvshowDoc2:
        ppv.C1PrintPreviewControl1.Document = doc
        ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
        ppv.Doc.Generate()
        ppv.Show()
        Exit Sub  '#End
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
100:
    End Sub

    Public Sub DataBaseToScreen(ByVal myform As frmQuoteRpt, ByRef QutID As String, ByRef ProjID As String)
        'Get the other Tables for this quote
        'Dim strsql As String = "" '"qutSlssplit" & "projectcust" & "qutnotes"
        'Dim pos As Integer
        'Dim drqut As dsSaw8.quoteRow
        'Dim drproj As dsSaw8.projectRow
        Try

            Dim mysqlcmd As New MySqlCommand
            dsQuote = New dsSaw8 : dsQuote.EnforceConstraints = False '03-19-14
            Dim daQuoteSLS As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            strSql = "Select * from qutslssplit Where QuoteID = '" & QutID & "'"
            mysqlcmd.Connection = myConnection
            mysqlcmd.CommandText = strSql
            daQuoteSLS.SelectCommand = mysqlcmd
            dsQuote.qutslssplit.Clear()
            daQuoteSLS.Fill(dsQuote, "qutSlssplit")
            Dim cbQutSLS As MySql.Data.MySqlClient.MySqlCommandBuilder
            cbQutSLS = New MySqlCommandBuilder(daQuoteSLS)
            'myform.tgSlsSplit.Rebind(True)
            'Don'T need yet"Print Spec Credit Lines"
            If (frmQuoteRpt.pnlTypeOfRpt.Text = "Planned Projects" Or frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report") And frmQuoteRpt.chkPrtPlanLines.CheckState = CheckState.Checked Then 'Planned Projects'11-24-09
                'Debug.Print(frmQuoteRpt.chkPrtPlanLines.CheckState & frmQuoteRpt.pnlTypeOfRpt.Text)
                Dim daQuoteline As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
                strSql = "Select * from quotelines where QuoteID = '" & QutID & "' order by LnSeq"
                mysqlcmd.Connection = myConnection
                mysqlcmd.CommandText = strSql
                daQuoteline.SelectCommand = mysqlcmd
                dsQuote.quotelines.Clear()
                daQuoteline.Fill(dsQuote, "quotelines")
                Dim cbQutlin As MySql.Data.MySqlClient.MySqlCommandBuilder
                cbQutlin = New MySqlCommandBuilder(daQuoteline)
            End If
            Dim daQuoteNotes As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            strSql = "Select * from qutnotes where QuoteID = '" & QutID & "'"
            mysqlcmd.Connection = myConnection
            mysqlcmd.CommandText = strSql
            daQuoteNotes.SelectCommand = mysqlcmd
            dsQuote.qutnotes.Clear()
            daQuoteNotes.Fill(dsQuote, "qutnotes")
            Dim cbQutNote As MySql.Data.MySqlClient.MySqlCommandBuilder
            cbQutNote = New MySqlCommandBuilder(daQuoteNotes)

            Dim daQuoteTo As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            strSql = "Select * from projectcust where QuoteID = '" & QutID & "' order by projectcust.Typec, projectcust.NCode" '09-10-10
            mysqlcmd.Connection = myConnection
            mysqlcmd.CommandText = strSql
            daQuoteTo.SelectCommand = mysqlcmd
            dsQuote.projectcust.Clear()
            daQuoteTo.Fill(dsQuote, "projectcust")
            Dim cbProjCust As MySql.Data.MySqlClient.MySqlCommandBuilder '11-24-09
            cbProjCust = New MySqlCommandBuilder(daProjCust) '11-24-09 
            '"qutSlssplit", "qutnotes", "projectcust"'01-20-09
        Catch ex As Exception
            MessageBox.Show("Error in DataBaseToScreen (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox("DatabaseToScreen " & ex.Message)
            ' If DebugOn ThenStop
        End Try
    End Sub
    Private Sub MakeDoc1()
        'ppv.Doc.Clear() 'Clear the Doc
        'doc.Clear()

        'Dim rnd As New Random(DateTime.Now.Millisecond)

        'Dim n As Integer
        '' make a table
        'Dim rt1 As New RenderTable()
        '' set up some styles on the table
        ''
        '' Note: all style attributes are divided into ambient and non-ambient.
        '' Ambient attributes affect the data itself, whereas non-ambient attributes
        '' are those affecting the decorations. Examples of ambient are:
        '' Font, TextColor, text alignment. Examples of non-ambient are:
        '' Borders, Padding, Spacing.
        '' Ambient attributes are propagated via the objects containment,
        '' so that e.g. setting the Font on the table will affect text in cells.
        '' Non-ambient attributes are inherited via the styles hierarchy. In tables,
        '' to set a non-ambient on all cells, table.CellStyle should be used. Ambient
        '' attributes can be set via table.Style:
        'rt1.Style.GridLines.All = LineDef.Default
        'rt1.Style.TextAlignHorz = AlignHorzEnum.Right
        'rt1.Style.Font = New Font("Courier New", 12)
        'rt1.Style.TextColor = Color.Green
        'rt1.CellStyle.Padding.All = "1mm"

        '' fill the table with more or less random data
        'Dim nrows As Integer = rnd.Next(100, 500)
        'Dim ncols As Integer = 17 'rnd.Next(2, 4)

        'Dim row, col As Integer
        'For row = 0 To nrows
        '    For col = 0 To ncols
        '        n = rnd.Next()
        '        rt1.Cells(row, col).Text = col ' n.ToString()
        '        If rt1.Cells(nrows, col).Tag Is Nothing Then
        '            rt1.Cells(nrows, col).Tag = n
        '        Else
        '            rt1.Cells(nrows, col).Tag = n + CType(rt1.Cells(nrows, col).Tag, Long)
        '        End If
        '    Next col
        'Next row

        '' table headers and footers are implemented as row groups.

        '' The header:
        '' insert 2 rows for the header at the top:
        'rt1.Rows.Insert(0, 2)
        '' mark the top row as a table header (not repeated),
        '' set it up appropriately:
        'rt1.RowGroups(0, 1).Header = TableHeaderEnum.None
        'rt1.RowGroups(0, 1).Style.BackColor = Color.GreenYellow
        'rt1.Cells(0, 0).SpanCols = rt1.Cols.Count
        'rt1.Cells(0, 0).Text = _
        '    "This table is filled with random data. It also has a table header (this text), " + _
        '    "running headers with column captions, running footers duplicating the running headers, " + _
        '    "and a footer which prints the column totals."
        'rt1.Cells(0, 0).Style.TextAlignHorz = AlignHorzEnum.Center
        'rt1.Cells(0, 0).Style.Font = New Font("Courier New", 12, FontStyle.Bold)
        '' mark the 2nd row as a page header (i.e. repeated on each page).
        'rt1.RowGroups(1, 1).Header = TableHeaderEnum.All
        'rt1.RowGroups(1, 1).Style.TextColor = Color.Hon
        'rt1.RowGroups(1, 1).Style.BackColor = Color.DarkKhaki
        'For col = 0 To ncols
        '    rt1.Cells(1, col).Text = String.Format("Column {0}", col)
        'Next col

        '' The footer:
        '' We used the last row for totals. We push it down
        '' (to be printed at the very bottom of the table), and add
        '' a "running footer" in front, with column headers to be
        '' printed on each page:
        'n = rt1.Rows.Count - 1
        'rt1.Rows.Insert(n, 1)
        '' Orphan control:
        '' this line makes sure that at least 3 lines are printed before the
        '' footer on the same page.
        'rt1.RowGroups(n, 1).MinVectorsBefore = 3
        'rt1.RowGroups(n, 1).Footer = TableFooterEnum.All
        'rt1.RowGroups(n, 1).Style.TextColor = rt1.RowGroups(1, 1).Style.TextColor
        'rt1.RowGroups(n, 1).Style.BackColor = rt1.RowGroups(1, 1).Style.BackColor
        'For col = 0 To ncols
        '    rt1.Cells(n, col).Text = col ' rt1.Cells(1, col).Text
        'Next col

        '' the final footer with totals and a line saying "the end":
        'n = rt1.Rows.Count - 1
        'rt1.RowGroups(n, 2).Footer = TableFooterEnum.Page
        'rt1.RowGroups(n, 2).Style.BackColor = Color.SandyBrown
        'For col = 0 To ncols
        '    rt1.Cells(n, col).Text = CType(rt1.Cells(n, col).Tag, Long).ToString()
        'Next col
        'rt1.Cells(n + 1, 0).SpanCols = rt1.Cols.Count
        'rt1.Cells(n + 1, 0).Text = "The end."
        'rt1.Cells(n + 1, 0).Style.TextAlignHorz = AlignHorzEnum.Center
        ''rt1.CanSplitHorz = True
        'rt1.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage
        'doc.Body.Children.Add(rt1)
        'ppv.C1PrintPreviewControl1.Document = doc
        ''ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)
        'ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1 '12-12-08
        'ppv.Doc.Generate() '11-18-08
        'ppv.Show()
        'ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        'ppv.BringToFront()
        ''frmShowHideGrid.BringToFront() '03-10-09

    End Sub
    Private Sub MakeDoc2()
        '        On Error GoTo 9999
        '        ppv.Doc.Clear() 'Clear the Doc
        '        doc.Clear()
        '        'SetupPrintPreview
        '905:    ppv.Doc.PageLayout.PageSettings.Landscape = True
        '        'Not Used NowRTotals = New C1.C1Preview.RenderText
        '        ' define PageLayout for the first page
        '        Dim pl As New PageLayout()
        '        pl.PageSettings = New C1PageSettings()
        '        pl.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.Letter
        '        pl.PageSettings.Landscape = True
        '        pl.PageSettings.LeftMargin = ".5in" '".25cm"
        '        pl.PageSettings.RightMargin = ".5in" '".25cm"
        '        pl.PageSettings.TopMargin = ".5in"
        '        pl.PageSettings.BottomMargin = ".5in"
        '        '12-29-08doc.PageLayouts.FirstPage = pl
        '        doc.PageLayout = pl

        '        Dim rnd As New Random(DateTime.Now.Millisecond)

        '        Dim n As Integer
        '        ' make a table
        '        Dim RT As New RenderTable()
        '        ' set up some styles on the table
        '        '
        '        ' Note: all style attributes are divided into ambient and non-ambient.
        '        ' Ambient attributes affect the data itself, whereas non-ambient attributes
        '        ' are those affecting the decorations. Examples of ambient are:
        '        ' Font, TextColor, text alignment. Examples of non-ambient are:
        '        ' Borders, Padding, Spacing.
        '        ' Ambient attributes are propagated via the objects containment,
        '        ' so that e.g. setting the Font on the table will affect text in cells.
        '        ' Non-ambient attributes are inherited via the styles hierarchy. In tables,
        '        ' to set a non-ambient on all cells, table.CellStyle should be used. Ambient
        '        ' attributes can be set via table.Style:
        '        RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm" 'JTC
        '        RT.Style.GridLines.All = LineDef.Default
        '        RT.Style.TextAlignHorz = AlignHorzEnum.Right
        '        RT.Style.Font = New Font("Courier New", 12)
        '        RT.Style.TextColor = Color.Green
        '        RT.CellStyle.Padding.All = "1mm"

        '        ' fill the table with more or less random data
        '        Dim nrows As Integer = rnd.Next(100, 500)
        '        Dim ncols As Integer = 35 ' 17 'rnd.Next(2, 4)

        '        Dim row, col As Integer
        '        For row = 0 To nrows
        '            For col = 0 To ncols
        '                n = rnd.Next()
        '                RT.Cells(row, col).Text = col ' n.ToString()
        '                RT.Cols(col).Width = ".5in" '05-26-10 test
        '                If RT.Cells(nrows, col).Tag Is Nothing Then
        '                    RT.Cells(nrows, col).Tag = n
        '                Else
        '                    RT.Cells(nrows, col).Tag = n + CType(RT.Cells(nrows, col).Tag, Long)
        '                End If
        '            Next col
        '        Next row

        '        ' table headers and footers are implemented as row groups.

        '        ' The header:
        '        ' insert 2 rows for the header at the top:
        '        RT.Rows.Insert(0, 2)
        '        ' mark the top row as a table header (not repeated),
        '        ' set it up appropriately:
        '        'RT.RowGroups(0, 1).Header = TableHeaderEnum.None
        '        ' RT.RowGroups(0, 1).Style.BackColor = Color.GreenYellow
        '        RT.Cells(0, 0).SpanCols = RT.Cols.Count
        '        RT.Cells(0, 0).Text = _
        '            "This table is filled with random data. It also has a table header (this text), " + _
        '            "running headers with column captions, running footers duplicating the running headers, " + _
        '            "and a footer which prints the column totals."
        '        RT.Cells(0, 0).Style.TextAlignHorz = AlignHorzEnum.Center
        '        RT.Cells(0, 0).Style.Font = New Font("Courier New", 12, FontStyle.Bold)
        '        ' mark the 2nd row as a page header (i.e. repeated on each page).
        '        RT.RowGroups(0, 2).Header = TableHeaderEnum.All
        '        RT.RowGroups(0, 2).Style.TextColor = Color.Hon
        '        RT.RowGroups(0, 2).Style.BackColor = Color.DarkKhaki
        '        For col = 0 To ncols
        '            RT.Cells(1, col).Text = String.Format("Column {0}", col)
        '        Next col

        '        ' The footer:
        '        ' We used the last row for totals. We push it down
        '        ' (to be printed at the very bottom of the table), and add
        '        ' a "running footer" in front, with column headers to be
        '        ' printed on each page:
        '        n = RT.Rows.Count - 1
        '        RT.Rows.Insert(n, 1)
        '        ' Orphan control:
        '        ' this line makes sure that at least 3 lines are printed before the
        '        ' footer on the same page.
        '        RT.RowGroups(n, 1).MinVectorsBefore = 3
        '        RT.RowGroups(n, 1).Footer = TableFooterEnum.All
        '        RT.RowGroups(n, 1).Style.TextColor = RT.RowGroups(1, 1).Style.TextColor
        '        RT.RowGroups(n, 1).Style.BackColor = RT.RowGroups(1, 1).Style.BackColor
        '        For col = 0 To ncols
        '            RT.Cells(n, col).Text = col ' RT.Cells(1, col).Text
        '            RT.Cols(col).Width = ".5in" '05-26-10 test
        '        Next col

        '        ' the final footer with totals and a line saying "the end":
        '        n = RT.Rows.Count - 1
        '        RT.RowGroups(n, 2).Footer = TableFooterEnum.None
        '        RT.RowGroups(n, 2).Style.BackColor = Color.SandyBrown
        '        For col = 0 To ncols
        '            RT.Cells(n, col).Text = CType(RT.Cells(n, col).Tag, Long).ToString()
        '            RT.Cols(col).Width = ".5in" '05-26-10 test
        '        Next col
        '        RT.Cells(n + 1, 0).SpanCols = RT.Cols.Count
        '        RT.Cells(n + 1, 0).Text = "The end."
        '        RT.Cells(n + 1, 0).Style.TextAlignHorz = AlignHorzEnum.Center
        '        'RT.CanSplitHorz = True
        '        RT.Width = "auto"
        '        RT.RowGroups(0, 2).Header = C1.C1Preview.TableHeaderEnum.Page



        '        RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage
        '        RT.StretchColumns = StretchTableEnum.LastVectorOnPage '05-26-10
        '        RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage
        '        'RT.SplitHorzBehavior = SplitBehaviorEnum.SplitIfNeeded
        '        doc.Body.Children.Add(RT)
        '        ppv.C1PrintPreviewControl1.Document = doc
        '        'ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)
        '        ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1 '12-12-08
        '        ppv.Doc.Generate() '11-18-08
        '        ppv.Show()
        '        ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        '        ppv.BringToFront()
        '        'frmShowHideGrid.BringToFront() '03-10-09
        '        Exit Sub
        '9999:   Resume Next



    End Sub


    Public Function GetProjectCust(ByVal QuoteID As String) As dsSaw8.projectcustDataTable '05-03-10 JH
        Dim strsql As String = ""
        Dim myds As dsSaw8 : myds = New dsSaw8 : myds.EnforceConstraints = False
        Dim dt As New dsSaw8.projectcustDataTable

        Try
            '09-10-10 strsql = "Select * from ProjectCust where QuoteID = " & "'" & QuoteID & "'" & "' order by ProjectCust.Typec, ProjectCust.NCode" '09-10-10
            'Do For Everyone If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Then '03-06-12 
            '03-06-12 strsql = "Select * from projectcust where QuoteID = '" & QuoteID & "' order by projectcust.Typec, projectcust.NCode " 
            strsql = "Select * from projectcust where QuoteID = '" & QuoteID & "' order by projectcust.Typec, projectcust.NCode, projectcust.QuoteToDate DESC" '03-06=12
            Dim mydataadapterProject As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            mydataadapterProject.SelectCommand = New MySqlCommand(strsql, myConnection)
            Dim cbP As MySql.Data.MySqlClient.MySqlCommandBuilder = New MySqlCommandBuilder(mydataadapterProject)
            mydataadapterProject.Fill(myds, "ProjectCust")
            Return myds.projectcust
        Catch ex As Exception
            MessageBox.Show("Error in GetProjectCust (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox(ex.Message & " in GetProjectCust(VQRT)")
            Return Nothing
        End Try


    End Function

    Public Function GetNotes(ByVal QuoteID As String) As dsSaw8.qutnotesDataTable '05-03-10 JH
        Dim strsql As String = ""
        Dim myds As dsSaw8 : myds = New dsSaw8 : myds.EnforceConstraints = False
        Dim dt As New dsSaw8.qutnotesDataTable
        Try
            strsql = "Select * from QutNotes where QuoteID = " & "'" & QuoteID & "'"
            Dim mydataadapterProject As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            mydataadapterProject.SelectCommand = New MySqlCommand(strsql, myConnection)
            Dim cbP As MySql.Data.MySqlClient.MySqlCommandBuilder = New MySqlCommandBuilder(mydataadapterProject)
            mydataadapterProject.Fill(myds, "QutNotes")
            Return myds.qutnotes
        Catch ex As Exception
            MessageBox.Show("Error in GetNotes (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox(ex.Message & " in GetNotes(VQRT)")
            Return Nothing
        End Try


    End Function
    Public Sub PrintReportQuotes()
        Try '#Top
            'If frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked ThenStop 'chkDetailTotal  Unchecked = Detail
            'Public RC As Integer = 0 'Row Count '06-11-10 
            CommAmtA(0) = 0 : CommAmtA(1) = 0 : CommAmtA(2) = 0 : CommAmtA(3) = 0 ' 0=TotalPaid  1=PaidThisQuote 2=TotalUnPaid 3=UnpaidPaidthisquote'09-20-12
            Dim TGNameStr As String = "" 'Documentation Set Up a String of Names
            Dim TGWidthStr As String = "" 'Set Up a String of Widths
            Dim PrtCols As Int16 = 0
            'Dim PrtCostCommCol As Boolean = False  '01-12-13 JTC If DIST = True and No Cost Column then don't print Cost on QuoteTO PrtCostCommCol
            If frmQuoteRpt.ChkQuoteNoSpecifiers.CheckState = CheckState.Checked Then frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked '08-14-12 
6000:
            ppv.Doc.Clear() 'Clear the Doc
            Call SetupPrintPreview(FirmName) '09-18-08
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
            ' Because we want to show a wide table, we adjust the properties of the preview accordingly and hide all margins.           
            '04-30-10 JH ppv.C1PrintPreviewControl1.PreviewPane.HideMarginsState = C1.Win.C1Preview.HideMarginsFlags.All
            '04-30-10 JH ppv.C1PrintPreviewControl1.PreviewPane.HideMargins = C1.Win.C1Preview.HideMarginsFlags.None
            '04-30-10 JH ppv.C1PrintPreviewControl1.PreviewPane.PagesPaddingSmall = New Size(0, 0)' Set padding between pages with hidden margins to 0, so that no gap is visible.     
            ' Set the zoom mode.
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomMode = C1.Win.C1Preview.ZoomModeEnum.PageWidth
6310:
            ' Because we want to show a wide table, we adjust the properties of the preview accordingly: 
            ' Hide all margins.
            ' ppv.C1PrintPreviewControl1.PreviewPane.HideMarginsState = C1.Win.C1Preview.HideMarginsFlags.All
            ' Do not allow the user to show margins.
            'ppv.C1PrintPreviewControl1.PreviewPane.HideMargins = C1.Win.C1Preview.HideMarginsFlags.None
            ' Set padding between pages with hidden margins to 0, so that no gap is visible:
            'ppv.C1PrintPreviewControl1.PreviewPane.PagesPaddingSmall = New Size(0, 0)
            ' Set zoom mode:
            'ppv.C1PrintPreviewControl1.PreviewPane.ZoomMode = C1.Win.C1Preview.ZoomModeEnum.PageWidth

6315:
            frmShowHideGrid.tgShow.SetDataBinding(table, "")
            MaxCol = frmQuoteRpt.tgQh.Splits(0).DisplayColumns.Count - 1

            'Header 
            Dim RArea As C1.C1Preview.RenderArea = New C1.C1Preview.RenderArea

            'Type of Report - & Agency Name
            RT = New C1.C1Preview.RenderTable
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10 RT.SplitHorzBehavior = True '05-26-10 Test
            'RT.SplitVertBehavior = True '05-26-10 Test
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            RT.CellStyle.Padding.Left = "1mm" '12-13-12
            RT.CellStyle.Padding.Right = "1mm" '12-13-12
            RT.Style.GridLines.All = LineDef.Empty '  LineDef.Default  '12-04-10 & "  UserID = " & UserID 
            RT.Cells(0, 0).Text = "Report: " & frmQuoteRpt.pnlTypeOfRpt.Text.Trim & "  UserID = " & UserID & "    Report Date = " & Format(Now, "MM/dd/yyyy") '10-17-10

            RT.Cells(0, 1).Text = AGnam : RT.Cells(0, 1).Style.TextAlignHorz = AlignHorzEnum.Right : RT.Cells(0, 1).Style.FontBold = True '05-09-14 JTC Bold Company name
            Dim fs As Integer = frmQuoteRpt.FontSizeComboBox.Text
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Bold)
            '05-27-10 RT.RowGroups(0, 2).Header = C1.C1Preview.TableHeaderEnum.Page '05-26-10 
            'doc.Body.Children.Add(RT) '12-29-06
            'RArea.Children.Add(RT)
            'Sort Sequence & Page #
            'Select Criteria
            'RT = New C1.C1Preview.RenderTable
            '06-12-10 RT.RowGroups(0, 3).PageHeader = True '06-12-10 
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10RT.SplitHorzBehavior = True '05-26-10 Test
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '11-02-10 
            'RT.SplitVertBehavior = True '05-26-10 Test
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            RT.Style.GridLines.All = LineDef.Empty '  LineDef.Default
            'Debug.Print(frmQuoteRpt.txtSortSeq.Text)

            'Dim sort As String = "Primary Sort: " & frmQuoteRpt.txtPrimarySortSeq.Text
            'If frmQuoteRpt.txtSecondarySort.Text.Trim <> "" Then sort = sort + "Secondary Sort: " & frmQuoteRpt.txtSecondarySort.Text
            RT.Rows(1).Style.TextAlignHorz = AlignHorzEnum.Left
            RT.Rows(1).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs)
            RT.Cells(1, 0).Text = frmQuoteRpt.txtSortSeq.Text '11-19-10 
            RT.Cells(1, 1).Text = "Page [PageNo] of [PageCount]"
            RT.Cells(1, 1).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Right
            RT.Cells(2, 0).Text = "Select Criteria: " & SelectionText 'frmQuoteRpt.TtxtSortSelV.Text '07-02-09frmProjRpt.txtPrimarySortSeq.Text & " " & frmProjRpt.txtSecondarySort.Text '07-01-09
            RT.Cells(2, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
            RT.Cells(2, 0).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, 9)  '04-30-10 jh - FONT COMBO
            '06-12-10 RT.RowGroups(0, 3).Header = C1.C1Preview.TableHeaderEnum.Page '05-26-10  TableHeaderEnum.Page '05-26-10

            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            '03-19-14 RT.Style.BackColor = LemonChiffon
            '02-04-12 Header on First Page Col hdg on All Pages
            '02-04-12RArea.Children.Add(RT)
            '02-04-12doc.Body.Document.PageLayout.PageHeader = RArea
            doc.Body.Children.Add(RT) ''02-04-12 Header on First Page Col hdg on All Pages


            'END PAGE HEADER - DIFFERENT THAN THE TABLE HEADERS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            RT = New C1.C1Preview.RenderTable : RC = 0 '06-19-10
            RT.Width = "auto" '07-15-10'RT.Width = "" '07-14-10
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            RT.CellStyle.Padding.Left = "1mm" '12-13-12
            RT.CellStyle.Padding.Right = "1mm" '12-13-12
            RT.Style.GridLines.All = LineDef.Default : RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            '11-03-10 Moved to end06-19-10 Moved Up Column Headers
            PrtCols = 0   'Set Column Widths ***********************************************************************
            ''Debug.Print(RC.ToString)
            '01-12-13 JTC If DIST = True and No Cost Column then don't print Cost on QuoteTO Dim NoCostCommCol As Int16 = 0  
            '01-25-13 UnCommented Below Section Need to set Col Width01-24-13 Header prints below Commented out this section
            Dim ColName As String = "" '01-25-13
            For I = 0 To frmQuoteRpt.tgQh.Splits(0).DisplayColumns.Count - 1 ' Set Column Widths 
                TgWidth(I) = (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) '02-25-09
                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                If (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
                PrtCols += 1
                If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then
                    ColName = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name
                    If ColName = "Comm" Or ColName = "Margin" Or ColName = "LPComm" Or ColName = "LPMarg" Or ColName = "Cost" Or ColName = "Comm-$" Or ColName = "Comm-%" Then
                        'If DIST And ColName = "Cost" Then 'See Header Print at End 05-09-14 JTC Keep Cost on DIST
                        'Else
                        RT.Cells(0, PrtCols).Text = ""  'Skip
                        'End If
                    End If
                End If
                '01-12-13 JTC If DIST = True and No Cost Column then don't print Cost on QuoteTO Dim PrtCostCommCol PrtCostCommCol = True '01-23-13 Has Cost so print
                'If DIST = True And frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name = "Cost" Then PrtCostCommCol = True '01-23-13 Has Cost so print
                'If DIST = False And frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name = "Comm" Then PrtCostCommCol = True '01-23-13 Has Comm so print
                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name = "Sell" And PrtCols < 6 Then '11-17-11
                    Resp = MsgBox("The Sell column needs to be in the sixth column" & vbCrLf & "or greater to Print totals correctly." & vbCrLf & "Click Yes to stop and move the Sell to the right (or add columns on the left)." & vbCrLf & "Click No to run the report as is.", MsgBoxStyle.YesNoCancel, "Sell column totals print") '11-17-11
                    If Resp = vbYes Then Exit Sub
                End If
            Next
            PrtCols -= 1 '06-12-10  RC = 0 '01-24-13 Commented out above

            'REPORT BODY '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim RCS As Int32 = 0 '06-13-11 JTC Added = 0  Sub Table 
            '01-25-13 Don't Dim till you use it Sets RTS.Rows.Count to 1 ''''
            Dim RTS As RenderTable = New RenderTable '11-03-10 Sub Table 
            'Debug.Print(RTS.Rows.Count)

            RT = New C1.C1Preview.RenderTable ' Main Table
            RC = 0 '06-12-10
            RT.Width = "auto" '07-15-10
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm" '07-14-10
            RT.CellStyle.Padding.Left = "1mm" '12-13-12
            RT.CellStyle.Padding.Right = "1mm" '12-13-12
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
            RT.Style.GridLines.All = LineDef.Default : RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage : RT.Style.GridLines.All = LineDef.Default
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)

            Dim C As Integer = 0 ' Lev.TotGt
            Dim X As String = "ZeroLevels" ' "AddAllLevels" ''01-26-09
            Call TotalsCalc(X, B, C) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            CurrLev1 = "" : PrevLev1 = "" : CurrLev2 = "" : PrevLev2 = "" : Cmd = "" 'Cmd = "EOF"
            Dim A As String = "PrintLine"
            frmQuoteRpt.tgQh.UpdateData()
            Dim PrimarySortSeq As String = frmQuoteRpt.txtPrimarySortSeq.Text
            Dim SeconarySortSeq As String = frmQuoteRpt.cboSortSecondarySeq.Text
            Dim RowCnt As Integer = 0 'Major Print Loop
            Dim FirstLoop As Int16 = 0 '07-09-09
            drQRow = dsQutLU.QUTLU1.Rows(0)
            For Each drQRow In dsQutLU.QUTLU1.Rows 'dsQutLU
                If drQRow.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                'Debug.Print(frmQuoteRpt.tgQh.RowCount)
                If FirstLoop = 0 Then 'FirstLoop = 1
                    frmQuoteRpt.tgQh.MoveFirst() '07-07-09
                End If  'If drQRow.QuoteCode = "A11-0309" ThenStop
                '06-13-11 If It deleted in dsQutLU.QUTLU1.Rows it is not in TrueGrid
                If RowCnt > frmQuoteRpt.tgQh.RowCount - 1 Then GoTo next170 '06-13-11 165 '06-19-10Continue For 'Filter caused fewer records '07-07-0
                If drQRow.RowState = DataRowState.Deleted Then GoTo next170 '06-13-11 165 '06-19-10 Continue For ' GoTo 235 '06-19-08
                'Test Dim ColText2 As String = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(3).DataColumn.Text  '06-13-11 
                'test If ColText2 <> drQRow.QuoteCode ThenStop '06-13-11
                'Debug.Print(frmQuoteRpt.tgQh.Splits(0).DisplayColumns("Description").DataColumn.Text)
                'If FirstLoop = 0 Then '06-13-11 JTC 
                '06-13-11 moved to end of loop
                'If RCS <> 0 Then '11-04-10 RTS Sub RT
                'RT.Cells(RC, 0).RenderObject = RTS : RT.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '11-04-10
                'RTS = New C1.C1Preview.RenderTable : RCS = 0 '06-12-10
                'RTS.Style.Padding.All = "0mm" : RTS.Style.Padding.Top = "0mm" : RTS.Style.Padding.Bottom = "0mm" '07-14-10
                'RTS.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                'RTS.Style.GridLines.All = LineDef.Default : RTS.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                'RTS.StretchColumns = StretchTableEnum.LastVectorOnPage : RTS.Style.GridLines.All = LineDef.Default
                'RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
                'RC += 1
                'End If
                Dim Hit As Short ' = 1 '06-19-10 
                'Debug.Print(drQRow.QuoteCode)
                Call SelectHit9500(Hit, multsrtrvs) '01-25-09
                'Hit = 1 ' Test
                FirstLoop = 1 '11-11-10 
                If Hit = 0 Then GoTo 165 'Get Next '07-07-09
                Dim QutID As String, ProjID As String
                QutID = drQRow.QuoteID 'Me.tgQutLU.Columns("QuoteID").CellText(Me.tgQutLU.SelectedRows.Item(0))
                ProjID = drQRow.ProjectID ' Me.tgQutLU.Columns("ProjectID").CellText(Me.tgQutLU.SelectedRows.Item(0))

                If frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" Or (frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report") Then ' "Realization" Then '"Quote Summary"
                Else 'Not Needed on Summary
                    Call DataBaseToScreen(frmQuoteRpt, QutID, ProjID) 'to get other Tables for this quote
                End If
                '03-03-12 Moved Up Dim Hit As Short ' = 1 '06-19-10 
                Dim ColText As String = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(3).DataColumn.Text  '06-13-11Debug.Print(drQRow.QuoteCode)
                RowCnt += 1 '02-08-09
                'Debug.Print(RC.ToString)
                Call SubTotChk9360(RT, doc) '02-08-09
                '11-11-10 Moved Down to Fix Blank Lines and no hit
                If frmQuoteRpt.chkBlankLine.CheckState = CheckState.Checked Or frmQuoteRpt.chkMfgBreakdown.CheckState = CheckState.Checked Or frmQuoteRpt.chkCustomerBreakdown.CheckState = CheckState.Checked Or frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked Then
                    If frmQuoteRpt.chkBlankLine.CheckState = CheckState.Checked Then
                        RT.Cells(RC, 0).Text = "  " : RT.Cells(RC, 0).SpanCols = RT.Cols.Count  '11-04-10  'RT.Cells(RC, 0).Text = "  " : RT.Cells(RC, 0).SpanCols = 10 : RT.Cells(RC, 1).Text = "  "
                        RT.Rows(RC).Style.BackColor = Color.White '11-04-10
                        RC += 1
                    End If
164:            End If
                FirstLoop = 1
                Call PrintQuoteLineRpt946(A, RT) '01-30-09 '946 Format routine and Calc Routine
                'If frmQuoteRpt.ChkTotalsOnly.CheckState = CheckState.Checked ThenStop '01-29-13 
                If frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Unchecked Then ''chkDetailTotal  Unchecked = Detail
                    If frmQuoteRpt.ChkQuoteNoSpecifiers.CheckState = CheckState.Checked Then '08-14-12
                    Else
                        RC += 1
                    End If
                End If
                'Debug.Print(RC.ToString)
                Call TotalsCalc("AddAllLevels", B, C) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
                'Print Other stuff

155:
                'Debug.Print(frmQuoteRpt.ChkSpecifiers.Text)
                '02-01-14 JTC Don't Want to skip Specifiers on Descending
                '02-01-14If frmQuoteRpt.ChkSpecifiers.Text = "Sort Report by Descending Dollar" Then GoTo ChkSpecifiersSkip '09-06-12 JTC Fix SLS/Desc Dollar frmQuoteRpt.ChkSpecifiers.Text used for Descending Dollars also ChkSpecifiersSkip:  '09-06-12 
                '02-01-14 Test frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" and frmQuoteRpt.txtPrimarySortSeq.Text = "SalesMan" and frmQuoteRpt.cboSortSecondarySeq.Text = "Descending Dollar" then 

                If frmQuoteRpt.chkMfgBreakdown.CheckState = CheckState.Checked Or frmQuoteRpt.chkCustomerBreakdown.CheckState = CheckState.Checked Or frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked Then
                    Dim dtPC As dsSaw8.projectcustDataTable = GetProjectCust(drQRow.QuoteID)
                    Dim FirstTimePC As Int16 = 0 '07-15-10
                    Dim LastCustQT As String = "" '03-29-12 
                    Hit = 0 ''Dim Hit As Int16 = 0
                    'ProjectCust
                    For Each drQPCRow As dsSaw8.projectcustRow In dtPC.Rows
                        If drQPCRow.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                        Hit = 0
                        If frmQuoteRpt.ChkQuoteNoSpecifiers.CheckState = CheckState.Checked Then '08-14-12 frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked 
                            If (drQPCRow.Typec <> "M" And drQPCRow.Typec <> "C" And drQPCRow.Typec <> "O") Then
                                Hit = 1 '08-14-12 Has Specifier
                                Exit For '08-14-12
                            Else
                                Continue For
                            End If
                        End If
                        If frmQuoteRpt.chkMfgBreakdown.CheckState = CheckState.Checked And drQPCRow.Typec = "M" Then Hit = 1 'Then  Else Continue For
                        If frmQuoteRpt.chkCustomerBreakdown.CheckState = CheckState.Checked And drQPCRow.Typec = "C" Then Hit = 1 'Else Continue For
                        'QuoteTo Records are M,C,O  Specifiers are A,E,S,T,O
                        If frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked And (drQPCRow.Typec <> "M" And drQPCRow.Typec <> "C" And drQPCRow.Typec <> "O") Then Hit = 1 '09-09-10Else Continue For
                        If frmQuoteRpt.pnlTypeOfRpt.Text = "Planned Projects" And frmQuoteRpt.chkPrtPlanLines.CheckState = CheckState.Checked Then Hit = 1 'Planned Projects'11-24-09
                        If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.chkPrtPlanLines.CheckState = CheckState.Checked Then Hit = 1 '11-24-09 fix chkPrtPlanLines 
                        'Debug.Print(frmQuoteRpt.chkPrtPlanLines.CheckState.ToString)
                        If Hit = 1 Then  Else Continue For 'No Hit
                        If FirstTimePC = 0 Then '07-15-10
                            RTS = New C1.C1Preview.RenderTable : RCS = 0 '06-12-10
                            RTS.Style.Padding.All = "0mm" : RTS.Style.Padding.Top = "0mm" : RTS.Style.Padding.Bottom = "0mm" '07-14-10
                            RTS.CellStyle.Padding.Left = "1mm" '12-13-12
                            RTS.CellStyle.Padding.Right = "1mm" '12-13-12
                            RTS.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                            RTS.Style.GridLines.All = LineDef.Default : RTS.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                            RTS.StretchColumns = StretchTableEnum.LastVectorOnPage : RTS.Style.GridLines.All = LineDef.Default
                            RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
                        End If 'FirstTimePC As As Int16 = 0'07-15-10
                        FirstTimePC += 1 '07-15-10 
                        Dim I As Int16 = 0
                        If frmQuoteRpt.chkCustomerBreakdown.CheckState = CheckState.Checked And drQPCRow.Typec = "C" Then '03-29-12
                            '03-29-12 Need to Add a Check box "Only Show the Latest Quote To Each Customer
                            'Debug.Print(drQPCRow.NCode & drQPCRow.LastChgDate.ToString & drQPCRow.QuoteToDate.ToString)
                            If frmQuoteRpt.chkShowLatestCust.CheckState = CheckState.Checked Then '03-29-12
                                'Dim LastCustQT As String = "" '03-29-12 
                                If drQPCRow.NCode.Trim = "" And drQPCRow.FirmName.Trim = "" Then GoTo NoHitNext '03-29-12 Don't Print 
                                If LastCustQT = drQPCRow.NCode.Trim Then GoTo NoHitNext '03-29-12 JTC Only Print First Cust If More Than One
                                If LastCustQT = "" Then LastCustQT = drQPCRow.NCode.Trim '03-29-12
                                If LastCustQT <> drQPCRow.NCode.Trim Then LastCustQT = drQPCRow.NCode.Trim '03-29-12
                            End If
                        End If
                        RTS.Cells(RCS, I).Text = drQPCRow.Typec : RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Center : I += 1
                        RTS.Cells(RCS, I).Text = drQPCRow.NCode : RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Center : I += 1
                        RTS.Cells(RCS, I).Text = drQPCRow.FirmName : I += 1
                        RTS.Cells(RCS, I).Text = drQPCRow.ContactName : I += 1 '09-09-10 
                        RTS.Cells(RCS, I).Text = drQPCRow.SLSCode : I += 1 '03-06-12 Always print SLSCode on QuoteTO
                        'QuoteTo Records are M,C,O  Specifiers are A,E,S,T,O '09-09-10 Don't Print Dollars on Specifiers
                        If frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked And (drQPCRow.Typec <> "M" And drQPCRow.Typec <> "C" And drQPCRow.Typec <> "C") Then
                            'Specifiers are A,E,S,T,O     '03-06-12 RTS.Cells(RCS, I).Text = drQPCRow.SLSCode : I += 1 '09-09-10 
                            If drQPCRow.NCode.Trim = "" And drQPCRow.FirmName.Trim = "" Then GoTo NoHitNext '03-29-12 Don't Print 
                            GoTo HitEqualsHit '09-10-10  
                        End If
                        'QuoteTo Records are M,C,O 
                        If DIST Then                                 '  $  "%" '03-14-12
                            RTS.Cells(RCS, I).Text = Format(drQPCRow.Cost, DecFormat) : RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right '01-23-13 I += 1 '02-04-09
                            'If PrtCostCommCol = False Then RTS.Cells(RCS, I).Text = "  " '01-12-13 JTC If DIST = True and No Cost Column then don't print Cost on QuoteTO Dim PrtCostCommCol PrtCostCommCol = True '01-23-13 Has Cost so print
                            If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" '01-23-13 10-17-10 GoTo 850 'Skip
                            I += 1 '01-23-13
                            RTS.Cells(RCS, I).Text = Format(drQPCRow.Sell, DecFormat)
                            If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" '05-09-14 JTC Eliminate Dollars on QuoteTo for dist Report Quote summary if chkIncludeCommDolPer.Checked = CheckState.Unchecked
                            RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right : I += 1 '02-04-09
                        Else  'rep 10-16-10
                            RTS.Cells(RCS, I).Text = Format(drQPCRow.Sell, DecFormat) : RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right : I += 1 '02-04-09
                            '02-07-12 RTS.Cells(RCS, I).Text = Format(drQPCRow.Sell - drQPCRow.Cost, "########0.00")
                            RTS.Cells(RCS, I).Text = Format(drQPCRow.Comm, DecFormat) '02-07-12 
                            If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" '10-17-10 GoTo 850 'Skip
                            RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right : I += 1 '10-17-10 COMM $
                        End If
                        'Margin below
                        If DIST Then
                            FixSell = drQPCRow.Sell : FixProfit = FixSell - drQPCRow.Cost
                            If FixSell <> 0 Then FixProfitPer = FixProfit / (FixSell + 0.00001) Else FixProfitPer = 0 '08-22-02 WNA
                            RTS.Cells(RCS, I).Text = Format(FixProfitPer, "##0.00") & "%" '03-14-12
                            If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" '05-09-14 JTC Eliminate Dollars on QuoteTo for dist Report Quote summary if chkIncludeCommDolPer.Checked = CheckState.Unchecked
                            RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right
                            'If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" '10-17-10 GoTo 850 'Skip
                            'LampSell = drQRow.LPSell : LampCost = drQRow.LPCost
                        Else
                            Amt = drQPCRow.Sell : CommAmt = drQPCRow.Comm
                            If Amt <> 0 Then Commpct = (CommAmt / (Amt + 0.0001)) * 100 '
                            If Commpct > 900 Then Commpct = 999 Else If Commpct < -900 Then Commpct = -999 '06-24-04
                            RTS.Cells(RCS, I).Text = Format(Commpct, "##0.00") & "%" '03-14-12
                            RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right
                            If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" '10-17-10 GoTo 850 'Skip
                            'LampSell = drQRow.LPSell : LampCost = drQRow.LPCost
                        End If
HitEqualsHit:           '09-10-10 
                        If Hit <> 0 Then 'Print Proj Cust if any
                            '11-03-10 RTS.Width = "auto"
                            RTS.Width = "auto" '05-09-14 JH Added 
                            RTS.Cols(0).Width = ".5in" '07-15-10 RTS.Cols(0).Width = ".5in"
                            RTS.Cols(1).Width = "1in" '05-09-14 JTC  ".75in"
                            RTS.Cols(2).Width = "2.2in" '11-24-09 
                            RTS.Cols(3).Width = "2.in"
                            RTS.Cols(4).Width = ".6In" '05-09-14 JTC 5 to 603-14-12 "1in"
                            RTS.Cols(5).Width = "1in"
                            RTS.Cols(6).Width = "1in"
                            RTS.Cols(7).Width = "1in" '05-09-15 
                            RTS.Rows(RCS).Style.BackColor = LemonChiffon '01-19-13Color.LightGoldenrodYellow
                            RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular) '05-03-10 JH
                            RCS += 1 'Next Row 06-12-10 RT = New C1.C1Preview.RenderTable '11-24-09
                        End If
NoHitNext:              '03-29-12
                    Next 'End For Each drQPCRow As dsSaw8.projectcustRow In dtPC.Rows
                    'Debug.Print(frmQuoteRpt.ChkQuoteNoSpecifiers.Text)
                    If frmQuoteRpt.ChkQuoteNoSpecifiers.CheckState = CheckState.Checked Then '08-14-12
                        If frmQuoteRpt.ChkQuoteNoSpecifiers.CheckState = CheckState.Checked And Hit = 0 Then
                            'Hit = 0 '= no 08-14-12  Specifier
                            RC += 1 '08-14-12 so print
                            GoTo Next170Move
                        Else   'Hit = 1 = Has a specifier
                            ' RC -= 1 '08-14-12 Reduce RC should erase the line
                            GoTo Next170Move
                        End If
                    End If
                End If
ChkSpecifiersSkip:  'frmQuoteRpt.ChkSpecifiers.Text 09-06-12 ChkSpecifiersSkip: 
                If frmQuoteRpt.ChkQuoteNoSpecifiers.CheckState = CheckState.Checked Then GoTo Next170 '08-14-12
                'PRINT NOTES
                If frmQuoteRpt.chkNotes.CheckState = CheckState.Checked Then
                    Dim dTPC As dsSaw8.qutnotesDataTable = GetNotes(drQRow.QuoteID) '05-15-10 JH 
                    For Each drQNRow As dsSaw8.qutnotesRow In dTPC.Rows
                        If drQNRow.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                        RT.Cells(RC, 0).Text = "     " & drQNRow.Notes ''02-11-12 JTC Add Notes to Realization
                        RT.Cells(RC, 0).Style.BackColor = LightGray
                        RT.Cells(RC, 0).SpanCols = RT.Cols.Count '05-09-14  - 1 '11-04-10  9 '07-15-10 
                        RC += 1 '06-12-10 RT = New C1.C1Preview.RenderTable
                    Next
                End If

                If (frmQuoteRpt.pnlTypeOfRpt.Text = "Planned Projects" Or frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report") And frmQuoteRpt.chkPrtPlanLines.CheckState = CheckState.Checked Then 'Planned Projects'11-24-09
                    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                    RTS = New C1.C1Preview.RenderTable : RCS = 0 '11-03-10
                    RTS.Style.Padding.All = "0mm" : RTS.Style.Padding.Top = "0mm" : RTS.Style.Padding.Bottom = "0mm" '07-14-10
                    RTS.CellStyle.Padding.Left = "1mm" '12-13-12
                    RTS.CellStyle.Padding.Right = "1mm" '12-13-12
                    RTS.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                    RTS.Style.GridLines.All = LineDef.Default : RTS.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                    RTS.StretchColumns = StretchTableEnum.LastVectorOnPage : RTS.Style.GridLines.All = LineDef.Default
                    RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
                    '****************************************************

                    If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Then
                        If frmQuoteRpt.cboLinesInclude.Text = "Include All Lines on Job" Then
                        Else
                            'Debug.Print(frmQuoteRpt.cboLinesInclude.Text)
                            ' If dsQutLU. = True And frmQuoteRpt.cboLinesInclude.Text <> "Include Only Paid Items on the Job" Then GoTo GetNextLookup '06-19-10  Continue For '06-19-10
                            ' If drQRow.Paid = False And frmQuoteRpt.cboLinesInclude.Text <> "Include Only UnPaid Items on the Job" Then GoTo GetNextLookup '06-19-10  Continue For '06-19-10
                        End If


GetNextLookup:
                        GoTo Next170Move 'GetNextQuote Lookup
                    End If
                    '***********************************************************************
                    ' 0=TotalPaid  1=PaidThisQuote 2=TotalUnPaid 3=UnpaidPaidthisquote'09-20-12 CommAmtA(1) = 0 : CommAmtA(3) = 0
                    CommAmtA(0) += CommAmtA(1) : CommAmtA(2) += CommAmtA(3)
                    CommAmtA(1) = 0 : CommAmtA(3) = 0 'Zero out Low Level'= Paid
                    Dim drQutLn As dsSaw8.quotelinesRow '11-24-09
                    Dim LnQty As Double = 0 '09-21-12
                    For Each drQutLn In dsQuote.quotelines
                        If drQutLn.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                        '12-01-09 Exclude Some Lines
                        'Include All Lines on Job
                        'Include Only Paid Items on the Job
                        'Include Only UnPaid Items on the Job
                        'Debug.Print(frmQuoteRpt.cboLinesInclude.Text & frmQuoteRpt.txtJobNameSS.Text)
                        ' If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report")
                        If frmQuoteRpt.cboLinesInclude.Text = "Include All Lines on Job" Then
                        Else
                            'Debug.Print(frmQuoteRpt.cboLinesInclude.Text)
                            If drQutLn.Paid = True And frmQuoteRpt.cboLinesInclude.Text <> "Include Only Paid Items on the Job" Then GoTo GetNextQuoteLine '06-19-10  Continue For '06-19-10
                            If drQutLn.Paid = False And frmQuoteRpt.cboLinesInclude.Text <> "Include Only UnPaid Items on the Job" Then GoTo GetNextQuoteLine '06-19-10  Continue For '06-19-10
                        End If
MoreLineTests:          If frmQuoteRpt.txtJobNameSS.Text <> "ALL" Then
                            If Trim(drQutLn.MFG) = "" Then Continue For '12-01-09
                            If InStr("," & Trim(frmQuoteRpt.txtJobNameSS.Text) & ",", "," & Trim(drQutLn.MFG) & ",") Then '
                                'Sto ' MfgHit = 1
                            Else
                                Continue For '12-01-09
                            End If
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        LnQty = Val(drQutLn.Qty) '09-21-12
                        RTS.Cells(RCS, 0).Text = drQutLn.LnCode
                        RTS.Cells(RCS, 1).Text = drQutLn.Qty
                        RTS.Cells(RCS, 2).Text = drQutLn.Type
                        RTS.Cells(RCS, 3).Text = drQutLn.MFG
                        RTS.Cells(RCS, 4).Text = drQutLn.Description
                        If DIST Then
                            RTS.Cells(RCS, 5).Text = Val(drQutLn.Cost).ToString
                            RTS.Cells(RCS, 6).Text = Val(drQutLn.Sell).ToString
                        Else
                            If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '12-01-09
                                '12-01-09 RTS.Cells(RCS, 5).Text = drQutLn.Paid '12-01-09 Format(drQutLn.Paid, "##")
                                'Debug.Print(drQutLn.Paid)
                                RTS.Cells(RCS, 5).Text = Val(drQutLn.Comm).ToString '09-10-12
                                If drQutLn.Paid = True Then
                                    RTS.Cells(RCS, 6).Text = "Paid"
                                    ' 0=TotalPaid  1=PaidThisQuote 2=TotalUnPaid 3=UnpaidPaidthisquote'09-20-12
                                    CommAmtA(1) += Val(drQutLn.Comm) * LnQty '09-21-12  0 : CommAmtA(3) = 0
                                Else
                                    : RTS.Cells(RCS, 6).Text = "UnPaid" '09-10-12 
                                    CommAmtA(3) += Val(drQutLn.Comm) * LnQty '09-21-12 
                                End If
                                RTS.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Center
                            Else
                                RTS.Cells(RCS, 5).Text = Val(drQutLn.Sell).ToString
                                RTS.Cells(RCS, 6).Text = Val(drQutLn.Comm).ToString
                            End If
                        End If
                        RCS += 1 '09-10-12
GetNextQuoteLine:       '06-19-10           'Margin Calc Only Do Once Per Row *********************************************

                    Next 'Row in Quote Lines
                    RCS += 1
                    For I = 0 To 5 : RTS.Cells(RCS, I).Text = "  " : Next 'vbNewLine '06-19-10 vbCrLf
                    RTS.Cells(RCS, 1).Text = " Paid = " & Format(CommAmtA(1), "$#####0.00") & "     Unpaid = " & Format(CommAmtA(3), "$#####0.00")
                    RTS.Cells(RCS, 1).SpanCols = RTS.Cols.Count - 1 ' 4
                    RTS.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Left
                    '.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '04-04-12 Zero fix

                    RCS += 1
                    '06-12-10 RT.Width = "auto"
                    ' 0=TotalPaid  1=PaidThisQuote 2=TotalUnPaid 3=UnpaidPaidthisquote'09-20-12 CommAmtA(1) = 0 : CommAmtA(3) = 0
                    CommAmtA(0) += CommAmtA(1) : CommAmtA(2) += CommAmtA(3)
                    'Set up Quote Line Width
                    RTS.Cols(0).Width = ".75in" : RTS.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Center 'LnCode
                    RTS.Cols(1).Width = ".75in" : RTS.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Right  'Qty
                    RTS.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Left
                    RTS.Cols(2).Width = ".75in" : RTS.Cols(2).Style.TextAlignHorz = AlignHorzEnum.Center  'Type
                    RTS.Cols(3).Width = ".75in" : RTS.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Center  'Mfg
                    RTS.Cols(4).Width = "4.5in" : RTS.Cols(4).Style.TextAlignHorz = AlignHorzEnum.Left 'desc
                    RTS.Cols(5).Width = ".75in" : RTS.Cols(5).Style.TextAlignHorz = AlignHorzEnum.Right  'Cost
                    RTS.Cols(6).Width = ".75in" : RTS.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Right 'Sell
                    For I = 0 To 5 : RTS.Cells(RCS, I).Text = "  " : Next 'vbNewLine '06-19-10 vbCrLf
                    RCS += 1
                    RTS.Rows.Insert(0, 1)
                    RTS.Cells(0, 0).Text = "LnCode"
                    RTS.Cells(0, 1).Text = "QTY" ''Qty"
                    RTS.Cells(0, 2).Text = "TYPE" '  'Type
                    RTS.Cells(0, 3).Text = "MFG"   'Mfg
                    RTS.Cells(0, 4).Text = "DESCRIPTION" ' 'desc
                    RTS.Cells(0, 5).Text = "COMM"   'COMM
                    RTS.Cells(0, 6).Text = "PAID"   'Cost
                    'RTS.Rows(0).Style.BackColor = LemonChiffon '06-19-10 
                    RTS.Cols(0).Visible = False : RTS.Cols(0).Width = Unit.Empty 'Don't print LnCode
                    'End Line Items^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                End If
165:            '06-13-11 moved to end of loop
                'Debug.Print(RC & "RCS" & RCS & "rows= " & RTS.Rows.Count)
                'If RTS Is Nothing ThenStop
                'Debug.Print(RTS.Rows.Count)

                If RTS.Rows.Count <> 0 Then '03-13-12 If RCS <> 0 Then '11-04-10 RTS Sub RT
                    'Debug.Print(RT.Cols.Count)
                    RT.Cells(RC, 0).RenderObject = RTS
                    If RT.Cols.Count - 1 <> 0 Then RT.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '04-04-12 Zero fix
                    RT.Cells(RC, 0).SpanCols = RT.Cols.Count '05-09-14 JH Chg SpanCols No - 1 
                    RTS = New C1.C1Preview.RenderTable : RCS = 0 '06-12-10
                    RTS.Style.Padding.All = "0mm" : RTS.Style.Padding.Top = "0mm" : RTS.Style.Padding.Bottom = "0mm" '07-14-10
                    RTS.CellStyle.Padding.Left = "1mm" '12-13-12
                    RTS.CellStyle.Padding.Right = "1mm" '12-13-12
                    RTS.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                    RTS.Style.GridLines.All = LineDef.Default : RTS.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                    RTS.StretchColumns = StretchTableEnum.LastVectorOnPage : RTS.Style.GridLines.All = LineDef.Default
                    RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
                    If frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Unchecked Then '07-26-12 'chkDetailTotal  Unchecked = Detail
                        RC += 1  'chkDetailTotal  Unchecked = Detail
                    End If
                End If
Next170Move:    '08-14-12 
                frmQuoteRpt.tgQh.MoveNext() '07-09-09
Next170:        '06-13-11
            Next ' Quote For Each drQRow In dsQutLU.QUTLU1.Rows 'dsQutLU
            'Debug.Print(RT.Rows.Count)
            Cmd = "EOF" 'PUBLIC
            If frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked Then GoTo 175 '01-29-13 Skip when Totals Causing Blank Line Below'chkDetailTotal  Unchecked = Detail
            If RTS.Rows.Count <> 0 Then '03-13-12 If RCS <> 0 Then '11-04-10 RTS Sub RT
                RT.Cells(RC, 0).RenderObject = RTS
                If RT.Cols.Count - 1 <> 0 Then RT.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '05-21-13 Fix Zero
                RC += 1
            End If
175:
            If RT.Rows.Count < 1 Then '07-24-10
                RT = New C1.C1Preview.RenderTable : RC = 0 '06-12-10
                RT.Width = "auto" '07-15-10
                RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm" '07-14-10
                RT.CellStyle.Padding.Left = "1mm" '12-13-12
                RT.CellStyle.Padding.Right = "1mm" '12-13-12
                RT.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                RT.Style.GridLines.All = LineDef.Default : RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                RT.StretchColumns = StretchTableEnum.LastVectorOnPage : RT.Style.GridLines.All = LineDef.Default
                RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            End If
            Call RTColSize(RT, MaxCol, TgWidth) '06-19-10
            Call SubTotChk9360(RT, doc) '02-08-09'SellFixtureA(I),CostFixtureA(I),SellFixtureAExt(I)
            THDG = "**GRAND TOTAL  " & "Record Count = " & QuantityA(0).ToString '02-07-10 QuantityA(I)
            TRCT = "Record Count = " & QuantityA(0).ToString
            'FixSell, FixCost, FixProfit, LampSell, LampCost, ProfitLamp, CommAmt, Commpct
            '01-25-09 TotalLevels.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            Call TotPrt9250(THDG, TotalLevels.TotGt, RT, doc)
            RT.Rows(RC).PageBreakBehavior = BreakEnum.None '11-19-10 
            RT.BreakAfter = BreakEnum.None
            Call TotalsCalc("ZeroLevels", B, TotalLevels.TotGt) '02-18-10
            LnQuantityA = 0 '02-18-10
            Cmd = "" 'Off "EOF" 'PUBLIC
6010:       ''Start INSERT COLUMN HEADERs &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            RT.Rows.Insert(0, 1) '06-14-10
            RT.RowGroups(0, 1).PageHeader = True '06-09-10
            Dim Headertmp As String = ""
            PrtCols = 0   'Print Column Headers ***********************************************************************
            'Dim ColName As String = "" '01-24-13
            For I = 0 To frmQuoteRpt.tgQh.Splits(0).DisplayColumns.Count - 1
                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                If (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
                RT.Cells(0, PrtCols).Text = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name
                '03-19-14 
                RT.Cells(0, PrtCols).Style.BackColor = LemonChiffon '06-14-10 
                RT.Cells(0, PrtCols).Style.FontBold = True '3-19-14
                RT.Cols(PrtCols).Width = TgWidth(I) '02-22-09
                If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then '01-24-13
                    ColName = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name
                    If ColName = "Comm" Or ColName = "Margin" Or ColName = "LPComm" Or ColName = "LPMarg" Or ColName = "Cost" Or ColName = "Comm-$" Or ColName = "Comm-%" Then '01-24-13
                        'If DIST And ColName = "Cost" Then '05-09-14 JTC Keep Cost on DIST
                        'Else
                        RT.Cells(0, PrtCols).Text = "" '01-24-13 'Skip
                        'End If
                    End If
                End If
                Headertmp = Headertmp & frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name & TgWidth(I).ToString
                PrtCols += 1
            Next
            PrtCols -= 1 '06-12-10  RC = 0
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            '03-19-14  RT.Style.BackColor = LemonChiffon
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            'Split across pages
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage : RT.StretchColumns = StretchTableEnum.LastVectorOnPage
            'Grid Lines
            RT.Style.GridLines.All = LineDef.Default '06-12-10 Else RT.Style.GridLines.All = LineDef.Empty
            doc.Body.Children.Add(RT) '06-12-10 
            'Footer
            Dim RF As New C1.C1Preview.RenderTable
            RF.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
            RF.Style.Padding.All = "0mm" : RF.Style.Padding.Top = "0mm" : RF.Style.Padding.Bottom = "0mm"
            RF.CellStyle.Padding.Left = "1mm" '12-13-12
            RF.CellStyle.Padding.Right = "1mm" '12-13-12
            RF.Style.GridLines.All = LineDef.Empty
            RF.Cells(0, 0).Text = Now.ToShortDateString
            RF.Cells(0, 1).Text = "Page [PageNo] of [PageCount]"
            RF.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs - 2, FontStyle.Bold)
            '03-19-14 RF.Style.BackColor = LemonChiffon '06-14-10 
            doc.Body.Document.PageLayout.PageFooter = RF
            'END HEADER &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

6090:
            ppv.C1PrintPreviewControl1.Document = doc
            'ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1 '12-12-08
            ppv.Doc.Generate() '11-18-08
            ppv.Show()
            ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
            ppv.BringToFront()
            frmShowHideGrid.BringToFront() '03-10-09



ExitReportLoop:
        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "PrintReportQuotes" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop
        End Try
Exit_Done:  '#End

    End Sub
    Public Sub PrintReportQuotesFollowBySesco()
        Try '#Top
            'Public RC As Integer = 0 'Row Count '06-11-10 
            Dim TGNameStr As String = "" 'Documentation Set Up a String of Names
            Dim TGWidthStr As String = "" 'Set Up a String of Widths
            Dim PrtCols As Int16 = 0
6000:
            ppv.Doc.Clear() 'Clear the Doc
            Call SetupPrintPreview(FirmName) '09-18-08
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
            ' Because we want to show a wide table, we adjust the properties of the preview accordingly and hide all margins.           
            '04-30-10 JH ppv.C1PrintPreviewControl1.PreviewPane.HideMarginsState = C1.Win.C1Preview.HideMarginsFlags.All
            '04-30-10 JH ppv.C1PrintPreviewControl1.PreviewPane.HideMargins = C1.Win.C1Preview.HideMarginsFlags.None
            '04-30-10 JH ppv.C1PrintPreviewControl1.PreviewPane.PagesPaddingSmall = New Size(0, 0)' Set padding between pages with hidden margins to 0, so that no gap is visible.     
            ' Set the zoom mode.
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomMode = C1.Win.C1Preview.ZoomModeEnum.PageWidth
6310:
            ' Because we want to show a wide table, we adjust the properties of the preview accordingly: 
            ' Hide all margins.
            ' ppv.C1PrintPreviewControl1.PreviewPane.HideMarginsState = C1.Win.C1Preview.HideMarginsFlags.All
            ' Do not allow the user to show margins.
            'ppv.C1PrintPreviewControl1.PreviewPane.HideMargins = C1.Win.C1Preview.HideMarginsFlags.None
            ' Set padding between pages with hidden margins to 0, so that no gap is visible:
            'ppv.C1PrintPreviewControl1.PreviewPane.PagesPaddingSmall = New Size(0, 0)
            ' Set zoom mode:
            'ppv.C1PrintPreviewControl1.PreviewPane.ZoomMode = C1.Win.C1Preview.ZoomModeEnum.PageWidth

6315:
            frmShowHideGrid.tgShow.SetDataBinding(table, "")
            MaxCol = frmQuoteRpt.tgQh.Splits(0).DisplayColumns.Count - 1

            'Header 
            Dim RArea As C1.C1Preview.RenderArea = New C1.C1Preview.RenderArea

            'Type of Report - & Agency Name
            RT = New C1.C1Preview.RenderTable
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10 RT.SplitHorzBehavior = True '05-26-10 Test
            'RT.SplitVertBehavior = True '05-26-10 Test
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            RT.CellStyle.Padding.Left = "1mm" '12-13-12
            RT.CellStyle.Padding.Right = "1mm" '12-13-12
            RT.Style.GridLines.All = LineDef.Empty '  LineDef.Default  '12-04-10 & "  UserID = " & UserID 
            RT.Cells(0, 0).Text = "Report: " & frmQuoteRpt.pnlTypeOfRpt.Text.Trim & "  UserID = " & UserID & "    Report Date = " & Format(Now, "MM/dd/yyyy") '10-17-10

            RT.Cells(0, 1).Text = AGnam : RT.Cells(0, 1).Style.TextAlignHorz = AlignHorzEnum.Right
            Dim fs As Integer = frmQuoteRpt.FontSizeComboBox.Text
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Bold)
            '05-27-10 RT.RowGroups(0, 2).Header = C1.C1Preview.TableHeaderEnum.Page '05-26-10 
            'doc.Body.Children.Add(RT) '12-29-06
            'RArea.Children.Add(RT)
            'Sort Sequence & Page #
            'Select Criteria
            'RT = New C1.C1Preview.RenderTable
            '06-12-10 RT.RowGroups(0, 3).PageHeader = True '06-12-10 
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10RT.SplitHorzBehavior = True '05-26-10 Test
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '11-02-10 
            'RT.SplitVertBehavior = True '05-26-10 Test
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            RT.Style.GridLines.All = LineDef.Empty '  LineDef.Default
            'Debug.Print(frmQuoteRpt.txtSortSeq.Text)

            'Dim sort As String = "Primary Sort: " & frmQuoteRpt.txtPrimarySortSeq.Text
            'If frmQuoteRpt.txtSecondarySort.Text.Trim <> "" Then sort = sort + "Secondary Sort: " & frmQuoteRpt.txtSecondarySort.Text
            RT.Rows(1).Style.TextAlignHorz = AlignHorzEnum.Left
            RT.Rows(1).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs)
            RT.Cells(1, 0).Text = frmQuoteRpt.txtSortSeq.Text '11-19-10 
            RT.Cells(1, 1).Text = "Page [PageNo] of [PageCount]"
            RT.Cells(1, 1).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Right
            RT.Cells(2, 0).Text = "Select Criteria: " & SelectionText 'frmQuoteRpt.TtxtSortSelV.Text '07-02-09frmProjRpt.txtPrimarySortSeq.Text & " " & frmProjRpt.txtSecondarySort.Text '07-01-09
            RT.Cells(2, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
            RT.Cells(2, 0).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, 9)  '04-30-10 jh - FONT COMBO
            '06-12-10 RT.RowGroups(0, 3).Header = C1.C1Preview.TableHeaderEnum.Page '05-26-10  TableHeaderEnum.Page '05-26-10

            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            RT.Style.BackColor = LemonChiffon
            '02-04-12 Header on First Page Col hdg on All Pages
            '02-04-12RArea.Children.Add(RT)
            '02-04-12doc.Body.Document.PageLayout.PageHeader = RArea
            doc.Body.Children.Add(RT) ''02-04-12 Header on First Page Col hdg on All Pages
            'END PAGE HEADER

            'Start TABLE HEADERS Below' DIFFERENT THAN THE PAGE Header''''''''''''''''''''''''''
            RT = New C1.C1Preview.RenderTable : RC = 0 'TABLE HEADERS
            RT.Width = "auto" '07-15-10'RT.Width = "" '07-14-10
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            RT.CellStyle.Padding.Left = "1mm" '12-13-12
            RT.CellStyle.Padding.Right = "1mm" '12-13-12
            RT.Style.GridLines.All = LineDef.Default : RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            '11-03-10 Moved to end06-19-10 Moved Up Column Headers
            PrtCols = 0   'Print Column Headers ***********************************************************************
            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then PrtCols = 0 : GoTo ReportBody '03-09-12

            ''Debug.Print(RC.ToString)
            For I = 0 To frmQuoteRpt.tgQh.Splits(0).DisplayColumns.Count - 1
                TgWidth(I) = (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) '02-25-09
                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                If (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
                PrtCols += 1
                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name = "Sell" And PrtCols < 6 Then '11-17-11
                    Resp = MsgBox("The Sell column needs to be in the sixth column" & vbCrLf & "or greater to Print totals correctly." & vbCrLf & "Click Yes to stop and move the Sell to the right (or add columns on the left)." & vbCrLf & "Click No to run the report as is.", MsgBoxStyle.YesNoCancel, "Sell column totals print") '11-17-11
                    If Resp = vbYes Then Exit Sub
                End If
            Next
            PrtCols -= 1
            '03-11-12 doc.Body.Children.Add(RT) '03-11-12 
ReportBody:  '03-08-10 
            'REPORT BODY '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim RCS As Int32 = 0 '06-13-11 JTC Added = 0  Sub Table 
            Dim RTS As RenderTable = New RenderTable '11-03-10 Sub Table 

            RT = New C1.C1Preview.RenderTable ' Main Table
            RC = 0 '06-12-10
            RT.Width = "auto" '07-15-10
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm" '07-14-10
            RT.CellStyle.Padding.Left = "1mm" '12-13-12
            RT.CellStyle.Padding.Right = "1mm" '12-13-12
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
            RT.Style.GridLines.All = LineDef.Default : RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage : RT.Style.GridLines.All = LineDef.Default
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)

            Dim C As Integer = 0 ' Lev.TotGt
            Dim X As String = "ZeroLevels" ' "AddAllLevels" ''01-26-09
            Call TotalsCalc(X, B, C) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            CurrLev1 = "" : PrevLev1 = "" : CurrLev2 = "" : PrevLev2 = "" : Cmd = "" 'Cmd = "EOF"
            Dim A As String = "PrintLine"
            frmQuoteRpt.tgQh.UpdateData()
            Dim PrimarySortSeq As String = frmQuoteRpt.txtPrimarySortSeq.Text
            Dim SeconarySortSeq As String = frmQuoteRpt.cboSortSecondarySeq.Text
            Dim RowCnt As Integer = 0 'Major Print Loop
            Dim FirstLoop As Int16 = 0 '07-09-09
            Dim LastQuoteCode As String = "" '03-10-12 
            drQRow = dsQutLU.QUTLU1.Rows(0)
            For Each drQRow In dsQutLU.QUTLU1.Rows 'dsQutLU
                If drQRow.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                If FirstLoop = 0 Then 'FirstLoop = 1
                    frmQuoteRpt.tgQh.MoveFirst() '07-07-09
                End If  'If drQRow.QuoteCode = "A11-0309" ThenStop
                '06-13-11 If It deleted in dsQutLU.QUTLU1.Rows it is not in TrueGrid
                If RowCnt > frmQuoteRpt.tgQh.RowCount - 1 Then GoTo next170 '06-13-11 165 '06-19-10Continue For 'Filter caused fewer records '07-07-0
                If drQRow.RowState = DataRowState.Deleted Then GoTo next170 '06-13-11 165 '06-19-10 Continue For ' GoTo 235 '06-19-08
                If FirstLoop = 0 Then LastQuoteCode = drQRow.QuoteCode '03-10-12 
                Dim Hit As Short ' = 1 '06-19-10 
                'Debug.Print(drQRow.QuoteCode)
                Call SelectHit9500(Hit, multsrtrvs) '01-25-09
                'Hit = 1 ' Test
                FirstLoop = 1 '11-11-10 
                If Hit = 0 Then GoTo 165 'Get Next '07-07-09
                Dim QutID As String, ProjID As String
                QutID = drQRow.QuoteID 'Me.tgQutLU.Columns("QuoteID").CellText(Me.tgQutLU.SelectedRows.Item(0))
                ProjID = drQRow.ProjectID ' Me.tgQutLU.Columns("ProjectID").CellText(Me.tgQutLU.SelectedRows.Item(0))
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" Then ' "Realization" Then '"Quote Summary"
                Else 'Not Needed on Summary
                    Call DataBaseToScreen(frmQuoteRpt, QutID, ProjID) 'to get other Tables for this quote
                End If
                Dim ColText As String = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(3).DataColumn.Text  '06-13-11Debug.Print(drQRow.QuoteCode)
                RowCnt += 1 '02-08-09
                'Debug.Print(RC.ToString)
                Call SubTotChk9360(RT, doc) '02-08-09
                '11-11-10 Moved Down to Fix Blank Lines and no hit
                If frmQuoteRpt.chkBlankLine.CheckState = CheckState.Checked Or frmQuoteRpt.chkMfgBreakdown.CheckState = CheckState.Checked Or frmQuoteRpt.chkCustomerBreakdown.CheckState = CheckState.Checked Or frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked Then
                    If frmQuoteRpt.chkBlankLine.CheckState = CheckState.Checked Then
                        If frmQuoteRpt.chkSalesmanPerPage.CheckState = CheckState.Checked Then GoTo 164 ' No Blank Line if One per page
                        If LastQuoteCode <> drQRow.QuoteCode Then LastQuoteCode = drQRow.QuoteCode : GoTo 164 '03-10-12 '03-10-12 
                        RT.Cells(RC, 0).Text = "  "
                        If RT.Cols.Count > 1 Then RT.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '03-09-12    = RT.Cols.Count  '11-04-10  'RT.Cells(RC, 0).Text = "  " :  = 10 : RT.Cells(RC, 1).Text = "  "
                        RT.Rows(RC).Style.BackColor = Color.White '11-04-10
                        RC += 1
                    End If
164:            End If
                'Debug.Print(RC)
                FirstLoop = 1
                Call PrintQuoteLineRpt946(A, RT) '01-30-09 '946 Format routine and Calc Routine
                If frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Unchecked Then '02-15-09 'chkDetailTotal  Unchecked = Detail
                    If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-13-12
                        'No
                    Else
                        RC += 1 '06-12-10 RT = New C1.C1Preview.RenderTable
                    End If

                End If

                Call TotalsCalc("AddAllLevels", B, C) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
155:            'Print QuoteTo Rows Project Cust
                Dim PrevNCode As String = "" '03-06-12
                If frmQuoteRpt.chkMfgBreakdown.CheckState = CheckState.Checked Or frmQuoteRpt.chkCustomerBreakdown.CheckState = CheckState.Checked Or frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked Then
                    Dim dtPC As dsSaw8.projectcustDataTable = GetProjectCust(drQRow.QuoteID)
                    PrevNCode = "" '03-06-12 
                    Dim FirstTimePC As Int16 = 0 '07-15-10
                    Hit = 0 ''Dim Hit As Int16 = 0
                    For Each drQPCRow As dsSaw8.projectcustRow In dtPC.Rows
                        If drQPCRow.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line

                        Hit = 0
                        If frmQuoteRpt.chkMfgBreakdown.CheckState = CheckState.Checked And drQPCRow.Typec = "M" Then Hit = 1 'Then  Else Continue For
                        '03-05-12
                        If frmQuoteRpt.chkCustomerBreakdown.CheckState = CheckState.Checked And drQPCRow.Typec = "C" Then
                            Hit = 1
                            'If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Then '03-06-12 
                            '    If PrevNCode = drQPCRow.NCode Then Continue For Else PrevNCode = drQPCRow.NCode '03-06-12
                            'End If
                        End If
                        'QuoteTo Records are M,C,O  Specifiers are A,E,S,T,O
                        If frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked And (drQPCRow.Typec <> "M" And drQPCRow.Typec <> "C" And drQPCRow.Typec <> "O") Then Hit = 1 '09-09-10Else Continue For
                        If frmQuoteRpt.pnlTypeOfRpt.Text = "Planned Projects" And frmQuoteRpt.chkPrtPlanLines.CheckState = CheckState.Checked Then Hit = 1 'Planned Projects'11-24-09
                        If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.chkPrtPlanLines.CheckState = CheckState.Checked Then Hit = 1 '11-24-09 fix chkPrtPlanLines 
                        'Debug.Print(frmQuoteRpt.chkPrtPlanLines.CheckState.ToString)
                        If Hit = 1 Then  Else Continue For 'Skip if no Hit
                        'Sto '7-15-10 Add RT & New RT
                        '11-03-10 Dim RTS As RenderTable = New RenderTable ' Dim RCS As Int32 '11-03-10 Sub Table 
                        If FirstTimePC = 0 Then '07-15-10
                            RTS = New C1.C1Preview.RenderTable : RCS = 0 '06-12-10
                            RTS.Style.Padding.All = "0mm" : RTS.Style.Padding.Top = "0mm" : RTS.Style.Padding.Bottom = "0mm" '07-14-10
                            RTS.CellStyle.Padding.Left = "1mm" '12-13-12
                            RTS.CellStyle.Padding.Right = "1mm" '12-13-12
                            RTS.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                            RTS.Style.GridLines.All = LineDef.Default : RTS.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                            RTS.StretchColumns = StretchTableEnum.LastVectorOnPage : RTS.Style.GridLines.All = LineDef.Default
                            RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
                        End If 'FirstTimePC As As Int16 = 0'07-15-10
                        FirstTimePC += 1 '07-15-10 
                        Dim I As Int16 = 0
                        RTS.Cells(RCS, I).Text = drQPCRow.Typec : RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Center : I += 1
                        RTS.Cells(RCS, I).Text = drQPCRow.NCode : RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Center : I += 1
                        RTS.Cells(RCS, I).Text = drQPCRow.FirmName : I += 1
                        RTS.Cells(RCS, I).Text = drQPCRow.ContactName : I += 1 '09-09-10 
                        RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Center '03-22-12
                        RTS.Cells(RCS, I).Text = drQPCRow.SLSCode : I += 1 '03-06-12 Always print SLSCode on QuoteTO
                        'QuoteTo Records are M,C,O  Specifiers are A,E,S,T,O '09-09-10 Don't Print Dollars on Specifiers
                        'If frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked And (drQPCRow.Typec <> "M" And drQPCRow.Typec <> "C" And drQPCRow.Typec <> "C") Then QuoteTo :Hit = 1 'Else Specifiers Continue For
                        'frmQuoteRpt.chkCustomerBreakdown.CheckState
                        'Debug.Print(drQRow.QuoteCode)
                        If frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked And (drQPCRow.Typec <> "M" And drQPCRow.Typec <> "C" And drQPCRow.Typec <> "O") Then
                            'Specifiers are A,E,S,T,O 
                            GoTo HitEqualsHit '09-10-10 No Dollar Fields 
                        End If
                        'QuoteTo Records are M,C,O 
                        If DIST Then                                  '03-14-12 Added $
                            RTS.Cells(RCS, I).Text = Format(drQPCRow.Cost, DecFormat) : RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right : I += 1 '02-04-09
                            RTS.Cells(RCS, I).Text = Format(drQPCRow.Sell, DecFormat) : RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right : I += 1 '02-04-09
                            RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right / 3 - 14 - 12
                        Else  'rep 10-16-10
                            RTS.Cells(RCS, I).Text = Format(drQPCRow.Sell, DecFormat) : RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right : I += 1 '02-04-09
                            '02-07-12 RTS.Cells(RCS, I).Text = Format(drQPCRow.Sell - drQPCRow.Cost, "########0.00")
                            RTS.Cells(RCS, I).Text = Format(drQPCRow.Comm, DecFormat) '02-07-12 
                            If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" '10-17-10 GoTo 850 'Skip
                            RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right : I += 1 '10-17-10 COMM $
                        End If
                        'Margin below
                        If DIST Then
                            FixSell = drQPCRow.Sell : FixProfit = FixSell - drQPCRow.Cost
                            If FixSell <> 0 Then FixProfitPer = FixProfit / (FixSell + 0.00001) Else FixProfitPer = 0 '08-22-02 WNA
                            RTS.Cells(RCS, I).Text = Format(FixProfitPer, "##0.00") & "%" '03-14-12
                            RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right '03-14-12
                            'If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" '10-17-10 GoTo 850 'Skip
                            'LampSell = drQRow.LPSell : LampCost = drQRow.LPCost
                        Else
                            Amt = drQPCRow.Sell : CommAmt = drQPCRow.Comm
                            If Amt <> 0 Then Commpct = (CommAmt / (Amt + 0.0001)) * 100 '
                            If Commpct > 900 Then Commpct = 999 Else If Commpct < -900 Then Commpct = -999 '06-24-04
                            RTS.Cells(RCS, I).Text = Format(Commpct, "##0.00") & "%" '03-14-12
                            RTS.Cols(I).Style.TextAlignHorz = AlignHorzEnum.Right '03-14-12
                            If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RTS.Cells(RCS, I).Text = "" '10-17-10 GoTo 850 'Skip
                            'LampSell = drQRow.LPSell : LampCost = drQRow.LPCost
                        End If
HitEqualsHit:           '09-10-10 
                        If Hit <> 0 Then 'Print Proj Cust if any
                            '11-03-10 RTS.Width = "auto"
                            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-0912" '07-15-10 RTS.Cols(0).Width = ".5in"
                                RTS.Cols(0).Width = ".5in"
                                RTS.Cols(1).Width = ".75in"
                                RTS.Cols(2).Width = "2.in" '11-24-09 
                                RTS.Cols(3).Width = "2.in"
                                RTS.Cols(4).Width = ".5in" '03-14-12 "1in"
                                RTS.Cols(5).Width = "1in"
                                RTS.Cols(6).Width = "1in"
                                RTS.Cols(7).Width = "1in" '03-06-12 
                                '06-12-10 RTS.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
                                '06-12-10  RTS.Style.GridLines.All = LineDef.Default
                                RTS.Rows(RCS).Style.BackColor = LemonChiffon '01-19-13.LightGoldenrodYellow
                                If drQPCRow.Typec = "C" Then RTS.Rows(RCS).Style.BackColor = AntiqueWhite '01-19-13 Color.Honeydew ' Color.Khaki '03-06-12 
                                RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular) '05-03-10 JH
                                '06-12-10 doc.Body.Children.Add(RT) '12-29-06
                                RCS += 1 '06-12-10 RT = New C1.C1Preview.RenderTable '11-24-09
                            End If
                        End If
                        RCS += 1 '03-11-12
                    Next 'End For Each drQPCRow As dsSaw8.projectcustRow In dtPC.Rows
                    '03-09-12 Moved down PRINT NOTES Start
                    If frmQuoteRpt.chkNotes.CheckState = CheckState.Checked Then
                        '03-10-12 Dim dTPC As dsSaw8.qutnotesDataTable = GetNotes(drQRow.QuoteID) '05-15-10 JH 
                        Dim dTPN As dsSaw8.qutnotesDataTable = GetNotes(drQRow.QuoteID) '05-15-10 JH 
                        For Each drQNRow As dsSaw8.qutnotesRow In dTPN.Rows
                            If drQNRow.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-09-12
                                'Call SetRTSWidth(RTS) 'ByRef R As RenderTable) '03-09-12 call
                                RTS.Cells(RCS, 0).Text = "     " & drQNRow.Notes ''02-11-12 JTC Add Notes to Realization
                                RTS.Cells(RCS, 0).Style.BackColor = LightGray
                                'RTS.Cols(0).Width = "9.5in"
                                RTS.Cells(RCS, 0).SpanCols = RTS.Cols.Count '01-24-13 02-12-12
                                RCS += 1
                            Else
                                'RT.Cells(RC, 0).Text = drQNRow.Notes '06-12-10 
                                RT.Cells(RC, 0).Text = "     " & drQNRow.Notes ''02-11-12 JTC Add Notes to Realization
                                RT.Cells(RC, 0).Style.BackColor = LightGray
                                If RT.Cols.Count > 1 Then RT.Cells(RC, 0).SpanCols = RT.Cols.Count Else RT.Cols(0).Width = "7.25" '03-12-12  
                                RC += 1 '06-12-10 RT = New C1.C1Preview.RenderTable    'Debug.Print(RC.ToString)
                            End If
                        Next
                    End If
                    'Blank Line
                    If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-09-12
                        'If RC = 0 ThenStop
                        If frmQuoteRpt.chkBlankLine.CheckState = CheckState.Checked Then
                            'Debug.Print(RTS.Cols.Count)
                            RTS.Cells(RCS, 0).Text = "  "
                            RTS.Cells(RCS, 0).Style.BackColor = LightGray  'RTS.Cols(0).Width = "9.5in"
                            RTS.Cells(RCS, 0).SpanCols = RTS.Cols.Count     ' RTS.Cells(RC, 0).SpanCols = RTS.Cols.Count - 1 '03-10-12
                            RCS += 1
                        End If

                    End If
                    'End Print Notes
                    'If FirstTimePC > 0 Then '07-15-10
                    '    'FirstTimePC = 0
                    'End If 'FirstTimePC As As Int16 = 0'07-15-10
                End If '

                If (frmQuoteRpt.pnlTypeOfRpt.Text = "Planned Projects" Or frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report") And frmQuoteRpt.chkPrtPlanLines.CheckState = CheckState.Checked Then 'Planned Projects'11-24-09
                    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                    RTS = New C1.C1Preview.RenderTable : RCS = 0 '11-03-10
                    '11-03-10 RT.Width = "auto" '07-15-10
                    RTS.Style.Padding.All = "0mm" : RTS.Style.Padding.Top = "0mm" : RTS.Style.Padding.Bottom = "0mm" '07-14-10
                    RTS.CellStyle.Padding.Left = "1mm" '12-13-12
                    RTS.CellStyle.Padding.Right = "1mm" '12-13-12
                    RTS.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                    RTS.Style.GridLines.All = LineDef.Default : RTS.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                    RTS.StretchColumns = StretchTableEnum.LastVectorOnPage : RTS.Style.GridLines.All = LineDef.Default
                    RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
                    Dim drQutLn As dsSaw8.quotelinesRow '11-24-09
                    For Each drQutLn In dsQuote.quotelines
                        If drQutLn.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                        '12-01-09 Exclude Some Lines
                        'Include All Lines on Job
                        'Include Only Paid Items on the Job
                        'Include Only UnPaid Items on the Job
                        'Debug.Print(frmQuoteRpt.cboLinesInclude.Text & frmQuoteRpt.txtJobNameSS.Text)
                        If frmQuoteRpt.cboLinesInclude.Text = "Include All Lines on Job" Then
                        Else
                            'Debug.Print(frmQuoteRpt.cboLinesInclude.Text)
                            If drQutLn.Paid = True And frmQuoteRpt.cboLinesInclude.Text <> "Include Only Paid Items on the Job" Then GoTo GetNextQuoteLine '06-19-10  Continue For '06-19-10
                            If drQutLn.Paid = False And frmQuoteRpt.cboLinesInclude.Text <> "Include Only UnPaid Items on the Job" Then GoTo GetNextQuoteLine '06-19-10  Continue For '06-19-10
                        End If
MoreLineTests:          If frmQuoteRpt.txtJobNameSS.Text <> "ALL" Then
                            If Trim(drQutLn.MFG) = "" Then Continue For '12-01-09
                            If InStr("," & Trim(frmQuoteRpt.txtJobNameSS.Text) & ",", "," & Trim(drQutLn.MFG) & ",") Then '
                                'Sto ' MfgHit = 1
                            Else
                                Continue For '12-01-09
                            End If
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        RTS.Cells(RCS, 0).Text = drQutLn.LnCode
                        RTS.Cells(RCS, 1).Text = drQutLn.Qty
                        RTS.Cells(RCS, 2).Text = drQutLn.Type
                        RTS.Cells(RCS, 3).Text = drQutLn.MFG
                        RTS.Cells(RCS, 4).Text = drQutLn.Description
                        If DIST Then
                            RTS.Cells(RCS, 5).Text = Val(drQutLn.Cost).ToString
                            RTS.Cells(RCS, 6).Text = Val(drQutLn.Sell).ToString
                        Else
                            If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '12-01-09
                                '12-01-09 RTS.Cells(RCS, 5).Text = drQutLn.Paid '12-01-09 Format(drQutLn.Paid, "##")
                                If drQutLn.Paid = True Then RTS.Cells(RCS, 5).Text = "Paid" Else RTS.Cells(RCS, 5).Text = "UnPaid" '12-01-09 
                                RTS.Cols(5).Style.TextAlignHorz = AlignHorzEnum.Center
                            Else
                                RTS.Cells(RCS, 5).Text = Val(drQutLn.Sell).ToString
                                RTS.Cells(RCS, 6).Text = Val(drQutLn.Comm).ToString
                            End If
                        End If
                        '11-04-10 RCS += 1
GetNextQuoteLine:       '06-19-10           'Margin Calc Only Do Once Per Row *********************************************

                    Next 'Row in Quote Lines

                    '06-12-10 RT.Width = "auto"
                    'Set up Quote Line Width
                    RTS.Cols(0).Width = ".75in" : RTS.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Center 'LnCode
                    RTS.Cols(1).Width = ".75in" : RTS.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Right  'Qty
                    RTS.Cols(2).Width = ".75in" : RTS.Cols(2).Style.TextAlignHorz = AlignHorzEnum.Center  'Type
                    RTS.Cols(3).Width = ".75in" : RTS.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Center  'Mfg
                    RTS.Cols(4).Width = "4.5in" : RTS.Cols(4).Style.TextAlignHorz = AlignHorzEnum.Left 'desc
                    RTS.Cols(5).Width = ".75in" : RTS.Cols(5).Style.TextAlignHorz = AlignHorzEnum.Right  'Cost
                    RTS.Cols(6).Width = ".75in" : RTS.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Right 'Sell
                    '11-02-10 RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
                    '11-02-10 RT.Style.GridLines.All = LineDef.Default
                    '11-02-10 RT.Rows(RC).Style.BackColor = Color.Li '11-24-09
                    '11-02-10 RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular) '05-03-10 JH
                    '11-02-10  RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
                    'Sto 'Blank Line
                    'RT.Rows(RC).PageBreakBehavior = RepeatGridLinesHorz = True
                    '11-04-10 RCS += 1
                    For I = 0 To 5 : RTS.Cells(RCS, I).Text = "  " : Next 'vbNewLine '06-19-10 vbCrLf
                    RCS += 1
                    RTS.Rows.Insert(0, 1)
                    RTS.Cells(0, 0).Text = "LnCode"
                    RTS.Cells(0, 1).Text = "QTY" ''Qty"
                    RTS.Cells(0, 2).Text = "TYPE" '  'Type
                    RTS.Cells(0, 3).Text = "MFG"   'Mfg
                    RTS.Cells(0, 4).Text = "DESCRIPTION" ' 'desc
                    RTS.Cells(0, 5).Text = "PAID??"   'Cost
                    RTS.Rows(0).Style.BackColor = LemonChiffon '06-19-10 
                    RTS.Cols(0).Visible = False : RTS.Cols(0).Width = Unit.Empty 'Don't print LnCode
                    'If RCS <> 0 Then '11-03-10 
                    '    RT.Cells(RC, 0).RenderObject = RTS : RT.Cells(RC, 0).SpanCols = RT.Cols.Count
                    '    RC += 1
                    'End If
                End If
165:
                'Debug.Print(RCS)
                If RTS IsNot Nothing Then '03-10-12  iIf RCS <> 0 Then '11-04-10 RTS Sub RT
                    'Debug.Print(RCS, RC)
                    If RC <> 0 Then ' If RT IsNot Nothing Then '03-10-12
                        doc.Body.Children.Add(RT)
                        RT = New C1.C1Preview.RenderTable : RC = 0 '06-12-10
                        RT.Width = "auto" '07-15-10
                        RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm" '07-14-10
                        RT.CellStyle.Padding.Left = "1mm" '12-13-12
                        RT.CellStyle.Padding.Right = "1mm" '12-13-12
                        RT.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                        RT.Style.GridLines.All = LineDef.Default : RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                        RT.StretchColumns = StretchTableEnum.LastVectorOnPage : RT.Style.GridLines.All = LineDef.Default
                        RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
                    End If
                    If RCS <> 0 Then doc.Body.Children.Add(RTS)
                    'RT.Cells(RC, 0).RenderObject = RTS : RC += 1 '03-10-12 
                    '03-09-12 If RT.Cols.Count > 1 Then RT.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '03-09-12  '03-09-12 Error Less than 1 RT.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '11-04-10
                    RTS = New C1.C1Preview.RenderTable : RCS = 0 '06-12-10
                    RTS.Style.Padding.All = "0mm" : RTS.Style.Padding.Top = "0mm" : RTS.Style.Padding.Bottom = "0mm" '07-14-10
                    RTS.CellStyle.Padding.Left = "1mm" '12-13-12
                    RTS.CellStyle.Padding.Right = "1mm" '12-13-12
                    RTS.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                    RTS.Style.GridLines.All = LineDef.Default : RTS.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                    RTS.StretchColumns = StretchTableEnum.LastVectorOnPage : RTS.Style.GridLines.All = LineDef.Default
                    RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
                    '03-09-12 RC += 1
                End If
                '06-13-11 Move Done once too Often
                frmQuoteRpt.tgQh.MoveNext() '07-09-09
Next170:        '06-13-11
            Next ' Quote For Each drQRow In dsQutLU.QUTLU1.Rows 'dsQutLU
            'Debug.Print(RT.Rows.Count)
            Cmd = "EOF" 'PUBLIC
            If RTS IsNot Nothing Then '03-10-12  If RCS <> 0 Then '11-04-10 RTS Sub RT
                RT.Cells(RC, 0).RenderObject = RTS
                If RT.Cols.Count > 1 Then RT.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '03-09-12  '03-09-12 Less Than 1 RT.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '11-04-10 
                'RTS.Style.Padding.All = "0mm" : RTS.Style.Padding.Top = "0mm" : RTS.Style.Padding.Bottom = "0mm" '07-14-10
                'RTS.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                'RTS.Style.GridLines.All = LineDef.Default : RTS.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                'RTS.StretchColumns = StretchTableEnum.LastVectorOnPage : RTS.Style.GridLines.All = LineDef.Default
                'RTS.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
                RC += 1
            End If
            If RT.Rows.Count < 1 Then '07-24-10
                RT = New C1.C1Preview.RenderTable : RC = 0 '06-12-10
                RT.Width = "auto" '07-15-10
                RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm" '07-14-10
                RT.CellStyle.Padding.Left = "1mm" '12-13-12
                RT.CellStyle.Padding.Right = "1mm" '12-13-12
                RT.StretchColumns = StretchTableEnum.LastVectorOnPage '07-14-10
                RT.Style.GridLines.All = LineDef.Default : RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
                RT.StretchColumns = StretchTableEnum.LastVectorOnPage : RT.Style.GridLines.All = LineDef.Default
                RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            End If
            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-08-12
            Else
                Call RTColSize(RT, MaxCol, TgWidth)
            End If
            Call SubTotChk9360(RT, doc) '02-08-09'SellFixtureA(I),CostFixtureA(I),SellFixtureAExt(I)
            THDG = "**GRAND TOTAL  " & "Record Count = " & QuantityA(0).ToString '02-07-10 QuantityA(I)
            TRCT = "Record Count = " & QuantityA(0).ToString
            'FixSell, FixCost, FixProfit, LampSell, LampCost, ProfitLamp, CommAmt, Commpct
            '01-25-09 TotalLevels.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            Call TotPrt9250(THDG, TotalLevels.TotGt, RT, doc)
            '03-09-12 RT.Rows(RC).PageBreakBehavior = BreakEnum.None '11-19-10 
            'RT.BreakAfter = BreakEnum.None
            Call TotalsCalc("ZeroLevels", B, TotalLevels.TotGt) '02-18-10
            LnQuantityA = 0 '02-18-10
            Cmd = "" 'Off "EOF" 'PUBLIC
            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then PrtCols = 0 : GoTo NoHeader '03-09-12
6010:       ''Start INSERT COLUMN HEADERs &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            RT.Rows.Insert(0, 1) '06-14-10
            RT.RowGroups(0, 1).PageHeader = True '06-09-10
            Dim Headertmp As String = ""
            PrtCols = 0   'Print Column Headers ***********************************************************************
            'Debug.Print(RC.ToString)
            For I = 0 To frmQuoteRpt.tgQh.Splits(0).DisplayColumns.Count - 1
                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                If (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
                RT.Cells(0, PrtCols).Text = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name
                RT.Cells(0, PrtCols).Style.BackColor = LemonChiffon '06-14-10 
                RT.Cols(PrtCols).Width = TgWidth(I) '02-22-09
                Headertmp = Headertmp & frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name & TgWidth(I).ToString
                PrtCols += 1
            Next
            PrtCols -= 1
NoHeader:   '03-09-12
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            RT.Style.BackColor = LemonChiffon
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            'Split across pages
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage : RT.StretchColumns = StretchTableEnum.LastVectorOnPage
            'Grid Lines
            RT.Style.GridLines.All = LineDef.Default '06-12-10 Else RT.Style.GridLines.All = LineDef.Empty
            doc.Body.Children.Add(RT) '06-12-10 

            'Footer
            Dim RF As New C1.C1Preview.RenderTable
            RF.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
            RF.Style.Padding.All = "0mm" : RF.Style.Padding.Top = "0mm" : RF.Style.Padding.Bottom = "0mm"
            RF.CellStyle.Padding.Left = "1mm" '12-13-12
            RF.CellStyle.Padding.Right = "1mm" '12-13-12
            RF.Style.GridLines.All = LineDef.Empty
            RF.Cells(0, 0).Text = Now.ToShortDateString
            RF.Cells(0, 1).Text = "Page [PageNo] of [PageCount]"
            RF.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs - 2, FontStyle.Bold)
            RF.Style.BackColor = LemonChiffon '06-14-10 
            doc.Body.Document.PageLayout.PageFooter = RF
            'END HEADER &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

6090:
            ppv.C1PrintPreviewControl1.Document = doc
            'ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1 '12-12-08
            ppv.Doc.Generate() '11-18-08
            ppv.Show()
            ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
            ppv.BringToFront()
            frmShowHideGrid.BringToFront() '03-10-09



ExitReportLoop:
        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "PrintReportQuotes" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop
        End Try
Exit_Done:  '#End

    End Sub
    Public Sub PrintReportQuoteLines() '09-01-09 
        Try '#Top  Convert This sub to QuoteLines
            'SESCO = False 'Stop:SESCO = True'test 7-24-14
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor '10-04-13
            Dim MaxRow As Single = 0 '09-01-09
            Dim Row As Integer = 0 '09-09-09
            Dim PrevJob As String = "" '09-21-12
            Dim CurrJob As String = "" '09-21-12
            Dim HeaderInfoPrinted As Boolean = False '03-19-14

            Dim mysqlcmd As New MySqlCommand '02-28-14
            CommAmtA(0) = 0 : CommAmtA(1) = 0 : CommAmtA(2) = 0 : CommAmtA(3) = 0 ' 0=TotalPaid  1=PaidThisQuote 2=TotalUnPaid 3=UnpaidPaidthisquote'09-20-12
6000:
            ppv.Doc.Clear() 'Clear the Doc
            Call SetupPrintPreview(FirmName) '09-18-08
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomMode = C1.Win.C1Preview.ZoomModeEnum.PageWidth
            '10-20-13 Moved Dim up
            Dim PC As Int16 = 0 '09-01-09
            Dim ColText As String = ""
            Dim ColName As String = ""
            'GoTo 6310 'Skip Logo
            '12-28-13 Moved Up
            Dim RArea As C1.C1Preview.RenderArea = New C1.C1Preview.RenderArea '10-20-13
            Dim fs As Integer = frmQuoteRpt.FontSizeComboBox.Text '10-20-13
            Dim C As Int16 = 0 '10-28-13 Moved up Dim C As Int16 = 0
            '12-28-13 Moved Up 
            Dim sEndBidDate As String = VB6.Format(frmQuoteRpt.DTPicker1EndBid.Value, "yyyy-MM-dd") ''02-03-12 - not /
            Dim sStartBidDate As String = VB6.Format(frmQuoteRpt.DTPicker1StartBid.Value, "yyyy-MM-dd") '10-23-13 Q.biddate is null or Q.biddate <>
            Dim SaveStartDate As String = sStartBidDate '05-28-14
            If ForecastAllMfg Then
                sEndBidDate = VB6.Format(frmQuoteRpt.DTPicker1EndEntry, "yyyy-MM-dd") ''02-03-12 - not /
                sStartBidDate = VB6.Format(frmQuoteRpt.DTPickerStartEntry.Value, "yyyy-MM-dd") '10-23-13 Q.biddate is <> null or Q.biddate <>
                SaveStartDate = sStartBidDate '05-28-14
            End If
6310:
            'Start Forecast & Brand Report %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            '12-28-13 moved If Here 
            If (frmQuoteRpt.chkBrandReport.CheckState = CheckState.Checked And frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And frmQuoteRpt.txtPrimarySortSeq.Text = "Forecasting") Or ForecastAllMfg = True Then '10-20-13 'And RptCatSel = RptCat.RptOrd Then '05-29-13 
                'Insert Header Forecast ******************************************************
                '10-23-13 JTC PhilipsForecast Report Quotes A=total each quote B=total each brand/Desc on each quote
                '1 - Done Select Just Status of Submit & Got or what they enter
                '2 - Done Select Exstimated  Delivery Date from Bid Date 
                '3 - Done  Report date to Last EstDelData IE: 10-30-13
                '4 - Done  Rep#
                '5 - Ignore EntryDates
                'Dim StartTime As String = Format(Now, "hh:mm tt") '02-27-14 'Dim EndTime As String = Format(Now, "hh:mm tt") '
                '03-24-14 Dim StartTimeM As DateTime = Format(Now, "hh:mm tt") 'Now.Minute '02-27-14
                '03-24-14 Dim EndTimeM As DateTime = Format(Now, "hh:mm tt") ' StartTimeM.AddMinutes(10)
                '03-24-14 Dim IntervalTypeM As DateInterval = DateInterval.Minute  ' Specifies Minute as interval.
                '03-24-14 Dim Step1 As String
                '03-24-14 Dim Step2 As String
                '03-24-14 Dim Step3 As String
                'If StartTimeM <= EndTimeM Then
                '    EndTimeM = DateAdd(IntervalTypeM, 15, StartTimeM) 'Add 15 minutes 
                'End If
                '03-24-14 StartTimeM = Format(Now, "hh:mm tt") ' EndTimeM = Format(Now, "hh:mm tt")
                'Step1 = DateDiff(DateInterval.Minute, StartTimeM, EndTimeM)
                'StartTimeM = EndTimeM : EndTimeM = DateAdd(IntervalTypeM, 15, StartTimeM) 'Add 15 minutes 
                'Step2 = DateDiff(DateInterval.Minute, StartTimeM, EndTimeM)
                'StartTimeM = EndTimeM : EndTimeM = DateAdd(IntervalTypeM, 15, StartTimeM) 'Add 15 minutes 
                'Step3 = DateDiff(DateInterval.Minute, StartTimeM, EndTimeM)
                'MsgBox("Step1= " & Step1 & "  Step2= " & Step2 & "  Step3= " & Step3 & " Minutes by Step" & vbCrLf & "PLease view the report file. ") '02-27-14 
                'MsgBox(msg)
                'DateTime End ))))))))))))))))))))))))))))))))))))))))))))))))))))))))))
                '02-28-14 Dim mysqlcmd As New MySqlCommand '10-20-13 
                Dim TmpStrSql As String = "" '10-21-13
                Dim LevBStrSQl As String = ""
                Dim SaveStrSQL As String = ""
                Dim LevASummary As String = ""
                Dim GtExtQlSell As Decimal = 0 '10-23-13 Grand Total
                Dim objExcel As Object = Nothing '10-27-13 Add Excel to Forecast Report
                Dim objBooks As Object = Nothing
                Dim objSheets As Object = Nothing
                Dim objSheet As Object = Nothing
                Dim objBook As Object = Nothing '10-27-13 
                objExcel = CreateObject("Excel.Application")
                '05-15-15 JTC Put at End objExcel.visible = True
                objExcel.DisplayAlerts = False
                objBook = objExcel.Workbooks.Add()
                objBooks = objExcel.Workbooks
                objBook = objBooks(1)
                objSheet = objBook.ActiveSheet()
                '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
                'If ForecastAllMfg = True Then '05-14-15 JTC Forecasting for MFGs Except Philips and SESCO

                'End If
SetWSCount:     '01-27-14 JTC Need three to find out thecurrent number of worksheets and then call Workbook.Worksheets.Add() 
                Dim Cnt As Int32 = objBook.Worksheets.Count '01-27-14 JTC Set Excel Forecast WookSheets.count to 3 Need three to find out thecurrent number of worksheets and then call Workbook.Worksheets.Add() 
                If Cnt < 3 Then
                    objBook.WorkSheets.Add()
                    GoTo SetWSCount '01-27-14
                End If
                objSheets = objBook.Worksheets
                objSheets(1).name = "Forecast By Quote Total"
                objSheets(2).name = "Quotes By Brand" '11-15-13
                objSheets(3).name = "Hold Orders By Brand" '11-15-13
                RC = 0
                '06-13-14 JTC added s(1) need Sheet numberto first Sheet to fix Heading not showing on Forecast report
                objSheets(1).Cells(RC + 1, 1) = "A" ' Rep Forecasting Report" "Forecast By Quote Total"
                objSheets(1).Cells(RC + 1, 2) = "Q-O" ''Qty"
                objSheets(1).Cells(RC + 1, 3) = "Rep#" '  'Type
                objSheets(1).Cells(RC + 1, 4) = "Rpt-Date" '  
                objSheets(1).Cells(RC + 1, 5) = "Deliv-Date" '  
                objSheets(1).Cells(RC + 1, 6) = "Qut-Num" '
                objSheets(1).Cells(RC + 1, 7) = "Status" '  '11-14-13 Add Status Column 
                objSheets(1).Cells(RC + 1, 8) = "MFG-Q#" ' SourceQuote #
                objSheets(1).Cells(RC + 1, 9) = "Ext-Sell" '  
                objSheets(1).Cells(RC + 1, 10) = "Description"
                objSheets(1).Cells(RC + 1, 11) = "U-Qty"
                objSheets(1).Cells(RC + 1, 12) = "U-Sell"
                'objSheets(I).Columns(6).NumberFormat = "@" '05-20-15 JTCHorizontalAlignment = 2 'Columns(1).NumberFormat = "@"  Columns(1).NumberFormat = "@"
                'Set for Text RT.Cols(Excel0).Width = ".75in" : RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Center 'LnCode
                ' objSheets(1).Column()
                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                '06-13-14 JTC added s(1) need Sheet numberto first Sheet to fix Heading not showing on Forecast report
                '07-22-14 JTC Fix Header Quotes By Brand Below Commented out in Error
                objSheets(2).Cells(RC + 1, 1) = "B" ' Rep Forecasting Report" "Quotes By Brand" 
                objSheets(2).Cells(RC + 1, 2) = "Q-B" ''Qty"
                objSheets(2).Cells(RC + 1, 3) = "Rep#" '  'Type
                objSheets(2).Cells(RC + 1, 4) = "Rpt-Date" '  
                objSheets(2).Cells(RC + 1, 5) = "Deliv-Date" '  
                objSheets(2).Cells(RC + 1, 6) = "Qut-Num" ' RT.Cols(0).Width = ".75in" : RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Center 'LnCode  
                objSheets(2).Cells(RC + 1, 7) = "Status" '  '11-14-13 Add Status Column 
                objSheets(2).Cells(RC + 1, 8) = "MFG" ' or Job-Name"
                objSheets(2).Cells(RC + 1, 9) = "Ext-Sell" '  
                objSheets(2).Cells(RC + 1, 10) = "Description"
                objSheets(2).Cells(RC + 1, 11) = "U-Qty"
                objSheets(2).Cells(RC + 1, 12) = "U-Sell"
                RC += 1 '07-22-14 JTC Added 1 to Fix Header not Showing Rept B Brand
                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                '12-28-13 moved Up If frmQuoteRpt.ChkBrandMfgRpt.CheckState = CheckState.Checked And frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And frmQuoteRpt.txtPrimarySortSeq.Text = "Forecasting" Then '10-20-13 'And RptCatSel = RptCat.RptOrd Then '05-29-13 
                If ForecastAllMfg = False Then '05-14-15 JTC Forecasting for MFGs Except Philips and SESCO
                    If BrandReportMfg.Trim = "" Then '=PHIL
                        BrandReportMfg = UCase(InputBox("Major Brand not set to breakdown PHIL, COOP, LITH, MFG, Etc." & vbCrLf & "Please check your Brand setup.", "Please check your Brand setup.", "MFG")) '03-12-13
                    End If
                End If
                If myConnection.State <> ConnectionState.Open Then Call OpenSQL(myConnection)
                mysqlcmd.Connection = myConnection
                TmpStrSql = "DROP TABLE IF EXISTS TMPREPORTS1 " '01-28-10
                mysqlcmd.CommandText = TmpStrSql : mysqlcmd.ExecuteNonQuery()
                TmpStrSql = "" '05-20-15 JTC
                'Fix where ( Q.Status = 'SUBMIT' or Q.Status = 'GOT' ) If UserDocDir.EndsWith(",") = False
                'Fix Select Q.* from Quote Q  where Q.TypeOfJob = 'Q'  and (Q.EstDelivDate >= '2013-04-01' and Q.EstDelivDate <= '2013-04-30' or  Q.biddate is null or Q.biddate = '1900-01-01') "
                Dim BC As String = "" '10-22-13
                Dim BCstatus As String = "" '10-22-13
                If frmQuoteRpt.txtStatus.Text.Contains(",") = True Then  '10-22-13 JTC Added
                    BCstatus = " and ( Q.Status = '" & frmQuoteRpt.txtStatus.Text.Replace(",", "' or Q.Status = '") & "' ) " '12-12-13 JTC Forecast Report Fix Status GOT or SUBMIT No Mfg need Status10-20-13 No Blanks or QL.MFG = ''  )"
                Else  'and ( Q.Status = 'OPEN'
                    BCstatus = " and ( Q.Status = '" & frmQuoteRpt.txtStatus.Text & "'" & " ) " '10-20-13 No Blanks or QL.MFG = '')"
                End If '"                '10-21-13 strSql = "Select  QL.*, Q.QuoteCode, Q.JobName, Sum(QL.Sell * QL.Qty) as ExtSell from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  where Q.EntryDate >= '2009-10-01' and Q.EntryDate <= '2013-10-31' and Q.TypeOfJob = 'Q'  and QL.Active = '1' and QL.LnCode <> 'NTE' and QL.LnCode <> 'NPE' and QL.LnCode <> 'SUB'  and QL.LnCode <> 'BTX' and QL.LnCode <> 'TXL'  and QL.LnCode <> 'TXS'  and QL.LnCode <> 'TXF'  and QL.LnCode <> 'TAX'  and QL.Description <> ''  and QL.MFG <> '' order by Q.QuoteCode, QL.MFG, QL.Description, QL.EntryDate DESC "
                If ForecastAllMfg = True And frmQuoteRpt.txtStatus.Text.ToUpper = "ALL" Then BCstatus = "" '05-14-15 JTC Forecasting for MFGs Except Philips and SESCO
                If ForecastAllMfg = True And frmQuoteRpt.txtStatus.Text.ToUpper = "ALL" Then BCstatus = "" '05-20-15 JTC Fix where Q.Status = 'ALL" Forecasting for MFGs Except Philips and SESCO
                'If Trim(frmQuoteRpt.txtStatus.Text.ToUpper) = "NOREPT" Then BC = " Q.Status <> "NOREPT" '05-12-15 JTC
                If ForecastAllMfg = True And BCstatus = "" Then BCstatus += " and Q.Status <> '" & "NOREPT" & "' " Else BCstatus += " And Q.Status <> '" & "NOREPT" & "' " '05-12-15 JTC"
                'Q.EstDelivDate Forecastin
                If frmQuoteRpt.ChkCheckBidDates.CheckState = CheckState.Checked Then '10-22-13
                    '12-28-13 Dim sEndBidDate As String = VB6.Format(frmQuoteRpt.DTPicker1EndBid.Value, "yyyy-MM-dd") ''02-03-12 - not /
                    'Dim sStartBidDate As String = VB6.Format(frmQuoteRpt.DTPicker1StartBid.Value, "yyyy-MM-dd") '10-23-13 Q.biddate is<> null or Q.biddate <>
                    SaveStrSQL = " and Q.BidDate >= '" & sStartBidDate & "' and Q.BidDate <= '" & sEndBidDate & "' " '10-23-13 or  Q.biddate is<> null or Q.biddate <> '" & "1900-01-01" & "'  " '04-04-12 added or Q.biddate = '" & "1900-01-01" & "'
                    'NO  No NullsSaveStrSQL = " and Q.BidDate >= '" & sStartBidDate & "' and Q.BidDate <= '" & sEndBidDate & "'  and  (Q.BidDate <> null or Q.biddate <> '" & "1900-01-01" & "' ) " '04-04-12 added or Q.biddate = '" & "1900-01-01" & "'

                    SaveStrSQL = Replace(SaveStrSQL, "Q.BidDate", "Q.EstDelivDate") '10-22-13
                    If ForecastAllMfg = False And BCstatus <> "" Then SaveStrSQL += BCstatus '05-28-15 Put status if First SQL
                Else
                    SaveStrSQL = " and Q.EntryDate >= '" & sStartBidDate & "' and Q.EntryDate <= '" & sEndBidDate & "' " '05-20-15 JTC10-23-13 or  Q.biddate is<> null or Q.biddate <> '" & "1900-01-01" & "'  " '04-04-12 added or Q.biddate = '" & "1900-01-01" & "'
                    SaveStrSQL = Replace(SaveStrSQL, "Q.BidDate", "Q.EstDelivDate") '10-22-13
                    If BCstatus <> "" Then SaveStrSQL += BCstatus '05-28-15 Put Q.EstDelivDate or Q.EntryDate if First Temporary SQL in PrintReportQuote lines
                End If
                'Q.biddate is<> null or Q.biddate <>                        '02-28-15 Added Q.SourceQuote
                '10-23-13 strSql = "Select  QL.MFG, QL.Description, QL.Sell, QL.UM, Q.QuoteCode, Q.JobName, Q.Status, Sum(QL.Qty) as Qty, Sum(QL.Sell * QL.Qty) as ExtSell from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  where (Q.Status = 'SUBMIT' or Q.Status = 'GOT') and Q.EntryDate >= '2009-10-01' and Q.EntryDate <= '2013-10-31' and Q.TypeOfJob = 'Q' and TypeOfJob = 'Q'   and QL.Active = '1' and QL.LnCode <> 'NTE' and QL.LnCode <> 'NPE' and QL.LnCode <> 'SUB'  and QL.LnCode <> 'BTX' and QL.LnCode <> 'TXL'  and QL.LnCode <> 'TXS'  and QL.LnCode <> 'TXF'  and QL.LnCode <> 'TAX'  and QL.Description <> ''  and QL.MFG <> '' order by Q.QuoteCode, QL.MFG, QL.Description, QL.EntryDate DESC "
                strSql = "Select  QL.MFG, QL.Description, QL.Sell, QL.UM, Q.QuoteCode, Q.SourceQuote, Q.JobName, Q.Status, Sum(QL.Qty) as Qty, Sum(QL.Sell * QL.Qty) as ExtSell, Q.EstDelivDate from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  where Q.TypeOfJob = 'Q' and TypeOfJob = 'Q'   and QL.Active = '1' and QL.LnCode <> 'NTE' and QL.LnCode <> 'NPE' and QL.LnCode <> 'SUB'  and QL.LnCode <> 'BTX' and QL.LnCode <> 'TXL'  and QL.LnCode <> 'TXS'  and QL.LnCode <> 'TXF'  and QL.LnCode <> 'TAX'  and QL.Description <> ''  and QL.MFG <> '' order by Q.QuoteCode, QL.MFG, QL.Description, QL.EntryDate DESC " '05-28-14 Q.EstDelivDate  
                If ForecastAllMfg = False Then 'Skip 05-14-15 JTC Forecasting for MFGs Except Philips and SESCO
                    strSql = Replace(strSql, "Q.BidDate", "Q.EstDelivDate") '10-23-13
                End If
                ' strSql = " and TypeOfJob = 'Q'   and QL.Active = '1' and QL.LnCode <> 'NTE' and QL.LnCode <> 'NPE' and QL.LnCode <> 'SUB'  and QL.LnCode <> 'BTX' and QL.LnCode <> 'TXL'  and QL.LnCode <> 'TXS'  and QL.LnCode <> 'TXF'  and QL.LnCode <> 'TAX'  and QL.Description <> ''  and QL.MFG <> '' order by Q.QuoteCode, QL.MFG, QL.Description, QL.EntryDate DESC "
                Dim STR1 As String = strSql.Substring(0, strSql.IndexOf("where")) '"Select  QL.MFG, QL.Description, QL.Sell, QL.UM, Q.QuoteCode, Q.JobName, Q.Status, Sum(QL.Qty) as Qty, Sum(QL.Sell * QL.Qty) as ExtSell from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  "

                '10-20-13 .txtQutRealCode.Text = PhilBrands
                If BrandList <> "" Then '10-20-13 BrandList=DAYB,CAPR,MCPH,OMEG,LAM,CHLO,MORL,GUTH,ARDE,CRES,GARD,EMCO,FORE,HADC,EXCE,TRAN,LOL,ALKC,LEDA,LUME,CRES,ALLS,HANO,THMO,THMI"
                    'STR1 = strSql.Substring(0, strSql.IndexOf("order by"))
                    Dim STR2 As String = " " & strSql.Substring(strSql.IndexOf("order by")) 'STR2 = " order by Q.QuoteCode, QL.MFG, QL.Description, QL.EntryDate DESC "
                    ' Dim BC As String = ""
                    ' TmpStrSql = " and ( QL.MFG = "
                    If ForecastAllMfg = True And BrandList = "" Then GoTo SkipMfgBrands '05-15-15 JTC Forecasting for MFGs Except Philips and SESCO
                    If BrandList.Contains(",") = True Then  '10-19-13 JTC Added QL.QL.MFG to Line Items
                        BC = " and ( QL.MFG = '" & BrandList.Replace(",", "' or QL.MFG = '") & "' ) " '10-20-13 No Blanks or QL.MFG = ''  )"
                    Else
                        BC = " ( QL.MFG = '" & BrandList & "'" & " ) " '10-20-13 No Blanks or QL.MFG = '')"
                    End If '" and ( QL.MFG = 'DAYB' or QL.MFG = 'CAPR' or QL.MFG = 'MCPH' or QL.MFG = 'OMEG' or QL.MFG = 'LAM' or QL.MFG = 'CHLO' or QL.MFG = 'MORL' or QL.MFG = 'GUTH' or QL.MFG = 'ARDE' or QL.MFG = 'CRES' or QL.MFG = 'GARD' or QL.MFG = 'EMCO' or QL.MFG = 'FORE' or QL.MFG = 'HADC' or QL.MFG = 'EXCE' or QL.MFG = 'TRAN' or QL.MFG = 'LOL' or QL.MFG = 'ALKC' or QL.MFG = 'LEDA' or QL.MFG = 'LUME' or QL.MFG = 'CRES' or QL.MFG = 'ALLS' or QL.MFG = 'HANO' or QL.MFG = 'THMO' or QL.MFG = 'THMI' or QL.MFG = ''  )"
                    '
                    STR1 += " where " ' & STR1 '05-15-15 Add where 
                    strSql = STR1 & TmpStrSql & SaveStrSQL & " and TypeOfJob = 'Q'   and QL.Active = '1' and QL.LnCode <> 'NTE' and QL.LnCode <> 'NPE' and QL.LnCode <> 'SUB'  and QL.LnCode <> 'BTX' and QL.LnCode <> 'TXL'  and QL.LnCode <> 'TXS'  and QL.LnCode <> 'TXF'  and QL.LnCode <> 'TAX'  and QL.Description <> ''  and QL.MFG <> '' "
                    If SaveStrSQL = "" Then '05-15-15 JTC No Bid Dates so no and needed
                        strSql = STR1 & TmpStrSql & SaveStrSQL & " TypeOfJob = 'Q'   and QL.Active = '1' and QL.LnCode <> 'NTE' and QL.LnCode <> 'NPE' and  QL.LnCode <> 'SUB'  and QL.LnCode <> 'BTX' and QL.LnCode <> 'TXL'  and QL.LnCode <> 'TXS'  and QL.LnCode <> 'TXF'  and QL.LnCode <> 'TAX'  and QL.Description <> ''  and QL.MFG <> '' "
                    End If
                    'Add Brand Codes and Status Codes
                    strSql += BC 'order by Q.QuoteCode, QL.MFG, QL.Description, QL.EntryDate DESC "
                    strSql = Replace(strSql, "where and", "where") '05-20-15 1 spc
                    strSql = Replace(strSql, "where  and", "where") '05-20-15 2 spc
                    If ForecastAllMfg = True Then strSql += BCstatus '05-20-15 JTC 
                    'First
                    'SaveStrSQL = STR1 & BC & " GROUP BY Q.QuoteCode order by Q.QuoteCode "
                    'If SaveStrSQL.EndsWith("GROUP BY Q.QuoteCode order by Q.QuoteCode ") Then LevASummary = "A" '10-21-13
                    '05-14-15 Moved Down LevASummary = "A" '10-22-13
                    If ForecastAllMfg = False Then '05-15-15 JTC  
                        strSql = Replace(strSql, "Q.BidDate", "Q.EstDelivDate") '10-22-13
                    End If
SkipMfgBrands:
                    LevASummary = "A" '10-22-13
                    If ForecastAllMfg = False Then 'Skip 05-14-15 JTC Forecasting for MFGs Except Philips and SESCO
                        LevBStrSQl = strSql & " GROUP BY Q.QuoteCode, QL.MFG, QL.Description order by Q.QuoteCode, QL.MFG, QL.Description "
                        strSql = strSql & " GROUP BY Q.QuoteCode order by Q.QuoteCode " '10-23-13 " GROUP BY Q.QuoteCode, QL.MFG, QL.Description order by Q.QuoteCode, QL.MFG, QL.Description "
                    Else 'Forecasting
                        '05-15-15 Already Set STR2 = strSql.Substring(0, strSql.IndexOf("order by")) '
                        LevBStrSQl = strSql & " GROUP BY Q.QuoteCode, QL.MFG, QL.Description order by Q.QuoteCode, QL.MFG, QL.Description "
                        strSql = strSql & " GROUP BY Q.QuoteCode order by Q.QuoteCode " '95-15-15  " GROUP BY Q.QuoteCode, QL.MFG, QL.Description order by Q.QuoteCode, QL.MFG, QL.Description "
                    End If
                End If
                SaveStrSQL = "CREATE TEMPORARY TABLE TMPREPORTS1 AS " & strSql '10-22-13
                mysqlcmd.CommandText = SaveStrSQL : SubCount = mysqlcmd.ExecuteNonQuery() 'Get All PHIL or COOP
                strSql = "Select * from TMPREPORTS1 QL"
                MsgBox("Please be patient!" & vbCrLf & "Quote Line item data can be slow when extracting to Excel.") '05-15-15 JTC
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor '05-15-15 JTC
                GoTo FillDataSetHere
                'TmpStrSql = "CREATE TEMPORARY TABLE TMPREPORTS1 AS " & strSql 'Select  OL.Qty, OL.LnCode, O.*, Sum((OL.Sell * OL.Qty)) as ExtSell, Sum((OL.Comm * OL.Qty)) as ExtComm, OL.MFG as LMFG, OS.SLSCode as SLSCode2 from ORDERMASTER O  left join ordslssplit OS on O.OrderID = OS.OrderID JOIN projectlines OL ON O.OrderID = OL.OrderID Where O.MFG = 'PHIL' and OS.slsnumber = 1 and O.RelHold = 'R'  and  OL.MFG <> '' and OL.Qty <> '' and OL.LnCode <> 'BTX'  and OL.LnCode <> 'NPN' and (concat(O.BuySellAB, O.BuySellSR) <> 'BS') GROUP BY PONumber, OL.MFG order by PONumber, OL.MFG "
                ''Get All PHIL or COOP Only based on User SQL
                '12-28-13 Movrd Down End If '10-21-13 ******************************************************************************************************************************************
FillDataSetHere:  'Forecast & Brand Report
                dsQuote = New dsSaw8 : dsQuote.EnforceConstraints = False
                daQuoteLine = New MySqlDataAdapter
                daQuoteLine.SelectCommand = New MySqlCommand(strSql, myConnection)
                Dim cbQutLin As MySql.Data.MySqlClient.MySqlCommandBuilder
                cbQutLin = New MySqlCommandBuilder(daQuoteLine)
                daQuoteLine.Fill(dsQuote, "quotelines")
                Dim ExtQlSell As Decimal = 0 '10-20-13 
                '10-28-13 Moved up Dim C As Int16 = 0
                Dim SaveMfg As String = ""
                Dim SaveMfgName As String = ""
                'RC += 1 '08-15-11
                '08-15-11
                'Debug.Print(ds.Tables("Invoicemaster").Rows.Count)
                '04-07-12 If RptCatSel = RptCat.RptInvHistory Then GoTo InvoiceSection '08-15-11
                '12-28-13 Moved UpDim RArea As C1.C1Preview.RenderArea = New C1.C1Preview.RenderArea '10-20-13
                '12-28-13 Dim fs As Integer = frmQuoteRpt.FontSizeComboBox.Text '10-20-13
                If LevASummary = "HoldOrder" Then  Else RC = 0 '11-03-13 Don't Zero If  HoldOrder
                PC = 0
                If LevASummary = "A" Then '11-03-13
                    'Type of Report - & Agency Name
                    RT = New C1.C1Preview.RenderTable
                    RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10 RT.SplitHorzBehavior = True '05-26-10 Test
                    'RT.SplitVertBehavior = True '05-26-10 Test
                    RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
                    RT.Rows(0).Height = "4.6mm" '02-07-13 Lower Heading Height
                    RT.CellStyle.Padding.Left = "1mm" '12-13-12
                    RT.CellStyle.Padding.Right = "1mm" '12-13-12
                    RT.Style.GridLines.All = LineDef.Empty '11-10-10 
                    RT.Rows(0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left '10-11-10 AlignHorzEnum.Left '11-10-10 

                    RT.Rows(0).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs)
                    '11-10-10 RT.Cells(0, 0).Style.Te
                    'Debug.Print(ds.Tables("quotelines").Rows.Count - 1)
                End If 'End Level a'11-03-13 
SecondLoop:
                Dim drQutLn As dsSaw8.quotelinesRow '11-24-09
                For Each drQutLn In dsQuote.quotelines
                    If drQutLn.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                    'If frmQuoteRpt.txtMfgLine2.Text.ToUpper <> "ALL" Then
                    '    If InStr("," & Trim(frmQuoteRpt.txtMfgLine2.Text.ToUpper) & ",", "," & drQutLn("MFG") & ",") = 0 Then drQutLn.Delete() : Continue For 'Don't Use
                    'End If
                    'Debug.Print(drQutLn("OrderNumber") & "  " & drQutLn("SlsBranch"))
                    '02-28-14 If IsDBNull(drQutLn("Status")) Then drQutLn("Status") = "" '11-14-13
                    If IsDBNull(drQutLn("MFG")) Then drQutLn("MFG") = "" '02-28-14 
                    If SaveMfg = "" Then SaveMfg = drQutLn("MFG") 'SaveMfgName = drQutLn("Firmname")
                    If SaveMfg = "" Then Continue For '11-19-10
                    If LevASummary = "A" Then
                        If IsDBNull(drQutLn("SourceQuote")) Then drQutLn("SourceQuote") = "" '02-28-14
                    End If
                    If LevASummary = "HoldOrder" Then '11-01-13
                        If IsDBNull(drQutLn("OrderNumber")) Then drQutLn("OrderNumber") = ""
                        If IsDBNull(drQutLn("EntryDate")) Then drQutLn("EntryDate") = ""
                    Else
                        If IsDBNull(drQutLn("QuoteCode")) Then drQutLn("QuoteCode") = "" '04-02-12 07-08-11 JTC
                    End If
                    If IsDBNull(drQutLn("Qty")) Then drQutLn("Qty") = "" '07-08-11 JTC
                    If IsDBNull(drQutLn("Description")) Then drQutLn("Description") = "" '07-08-11 JTC
                    'TildaCheck = : drQutLn("Description") = "~"
                    Dim ZC As Short '07-08-14
                    ZC = InStr(drQutLn("Description"), "~")
                    If ZC = 1 Or ZC = 2 Then Continue For '07-08-14 JTC Skip Tilda on Forecasting Quote Lines Delete All text after "~" Tilda
                    If ZC <> 0 Then
                        drQutLn("Description") = Mid(drQutLn("Description"), ZC + 1) '07-08-14
                    End If
                    'If frmQuoteRpt.txtPrimarySortSeq.Text = "Forecasting" Then
                    '    If InStr(drQutLn("Description"), "~") Then Continue For '07-08-14 JTC Skip Tilda on Forecasting Quote Lines
                    'End If
                    If IsDBNull(drQutLn("Sell")) Then drQutLn("Sell") = "0" '07-08-11 JTC
                    '02-28-14  Sum(OL.Qty) as Qty, Sum(OL.Sell * OL.Qty) as ExtSell
                    'Dim ExtSell As Double = drQutLn("Qty") * drQutLn("Sell") '02-28-14
                    If IsDBNull(drQutLn("ExtSell")) Then drQutLn("ExtSell") = "0" '07-08-11 JTC
                    '02-28-14 If IsDBNull(drQutLn("Status")) Then drQutLn("Status") = "" '11-14-13
                    Dim LineQty As Decimal = 0
                    LineQty = CDec(Val(frmQuoteRpt.tgln(Row, "Qty").ToString()))

                    Dim APrice As String = drQutLn("Sell") 'frmQuoteRpt.tgln(Row, "Sell").ToString
                    Dim UnitOfM As String = drQutLn("UM") 'frmQuoteRpt.tgln(Row, "UM").ToString
                    Dim UnitMeas As Decimal = 1
                    Dim UnitMeaStr As String = UnitMeaSet(APrice, UnitMeas, UnitOfM) '' C = Hundreds M = Thousands FT =Feet '01-28-04
                    'If IsDBNull(drQutLn("Ext Sell")) Then drQutLn("ExtSell") = 0 '07-08-11 JTC
                    '12-18-14 jh LOT PRICE QUOTE, THIS KEEPS IT OFF THE REPORT - If Trim(APrice) = "" Then Continue For
                    '12-18-14 jh LOT PRICE QUOTE, THIS KEEPS IT OFF THE REPORT - If Val(drQutLn("ExtSell")) = 0 Then Continue For '12-23-13 
                    ExtQlSell = Val(drQutLn("ExtSell"))
                    GtExtQlSell += ExtQlSell '10-23-13 Grand Total
                    ''07-24-14 JTC Forecast
                    If SESCO = True Then GtExtQlSell = 0 : ExtQlSell = 0 : APrice = 0
                    'ExtQlSell = Val(drQutLn("ExtSell"))'Else  ExtQlSell = Val(drQutLn("Sell")) * (Val(drQutLn("Qty") * UnitMeas))
                    'If SaveMfg <> drQutLn("MFG") Then '08-15-11
                    PC = 0 'Levels **********************************************
                    RT.Cells(RC, PC).Text = "B" ') SaveMfg & " Sales$" '12-15-11
                    If LevASummary = "HoldOrder" Then RT.Cells(RC, PC).Text = "C" '11-03-13
                    'If LevASummary = "A" Then 'SaveStrSQL.EndsWith("GROUP BY Q.QuoteCode order by Q.QuoteCode") then LevASummary = "A":'10-21-13
                    If LevASummary = "A" Then RT.Cells(RC, PC).Text = "A" '10-21-13
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center
                    If LevASummary = "A" Or LevASummary = "HoldOrder" Then RT.Rows(RC).Style.BackColor = AntiqueWhite Else RT.Rows(PC).Style.BackColor = Color.White '10-21-13
                    PC += 1  'Type
                    RT.Cells(RC, PC).Text = "QT"
                    If LevASummary = "HoldOrder" Then RT.Cells(RC, PC).Text = "HO" '11-01-13
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center
                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 
                    PC += 1  'Rep#
                    RT.Cells(RC, PC).Text = RepCustNumber '10-28-13  "12345678" ' RepNum RepCustNumber
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center
                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 

                    PC += 1  'ReportDate
                    RT.Cells(RC, PC).Text = Format(Now, "yyyy-MM-dd") '"2012-10-10"

                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center
                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 


                    PC += 1  'EstDeliveryDate
                    Try
                        If LevASummary = "A" Or LevASummary = "DoneWithA" Then '12-18-14 jh LevASummary = "DoneWithA" - IT'S IN OUR DS AND ON THE EXCEL SHEET, USE IT
                            RT.Cells(RC, PC).Text = drQutLn("EstDelivDate") '05-28-14  sEndBidDate '10-28-13"2012-15-10"
                            'If ForecastAllMfg Then RT.Cells(RC, PC).Text = drQutLn.EntryDate 'QL.EntryDate
                            RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center
                        End If
                        If LevASummary = "HoldOrder" Then RT.Cells(RC, PC).Text = Format(drQutLn("EntryDate"), "MM/dd/yyyy") & " " '11-15-13 
                    Catch ex As Exception
                    End Try

                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 

                    PC += 1  'QuoteCode
                    If LevASummary = "HoldOrder" Then '11-01-13
                        RT.Cells(RC, PC).Text = drQutLn("OrderNumber")
                    Else
                        RT.Cells(RC, PC).Text = drQutLn("QuoteCode")
                    End If
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Left
                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 
                    '11-14-13 Add Status Column 
                    PC += 1  'Status
                    RT.Cells(RC, PC).Text = "Status" '11-14-13 drQutLn("Mfg")
                    If LevASummary = "A" Or LevASummary = "DoneWithA" Then RT.Cells(RC, PC).Text = drQutLn("Status") '"PHIL" '11-14-13
                    If LevASummary = "HoldOrder" Then RT.Cells(RC, PC).Text = "HOLD" '11-14-13
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center

                    PC += 1  'Brand
                    RT.Cells(RC, PC).Text = drQutLn("Mfg")
                    If LevASummary = "A" Then RT.Cells(RC, PC).Text = drQutLn("SourceQuote") '02-28-14
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center
                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 

                    PC += 1  'ExtSell$
                    RT.Cells(RC, PC).Text = Format(ExtQlSell, "0")
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 

                    PC += 1  'drQutLn("Description")
                    RT.Cells(RC, PC).Text = drQutLn("Description")
                    If LevASummary = "A" Then RT.Cells(RC, PC).Text = drQutLn("JobName") '10-21-13 
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Left
                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 

                    PC += 1  'Qty
                    RT.Cells(RC, PC).Text = Format(Val(drQutLn("Qty")), "0")
                    If LevASummary = "A" Then RT.Cells(RC, PC).Text = "" '10-21-13
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 

                    PC += 1  'Sell$
                    RT.Cells(RC, PC).Text = Format(Val(drQutLn("Sell")), "0")
                    If LevASummary = "A" Then RT.Cells(RC, PC).Text = "" '10-21-13
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                    ''07-24-14 JTC 
                    If SESCO = True Then
                        GtExtQlSell = 0 : ExtQlSell = 0 : APrice = 0
                        RT.Cells(RC, 8).Text = "0" : RT.Cells(RC, 11).Text = "0"
                    End If
                    '02-28-14 Add JobName to HoldOrder report
                    If LevASummary = "HoldOrder" Then
                        PC += 1
                        RT.Cells(RC, PC).Text = drQutLn("JobName")
                    End If
                    'RT.Rows(RC).Style.BackColor = AntiqueWhite '11-19-10 
                    'Debug.Print(PC, LastBranch)
                    '11-14-13 Add Status Column 
                    For I = 0 To 11 ' 'Excel Starts with 1 not Zero !@#$%^&*
                        If LevASummary = "A" Then
                            objSheets(1).Cells(RC + 2, I + 1) = RT.Cells(RC, I).Text
                        ElseIf LevASummary = "HoldOrder" Then
                            objSheets(3).Cells(RC + 2, I + 1) = RT.Cells(RC, I).Text '11-15-13
                        Else
                            objSheets(2).Cells(RC + 2, I + 1) = RT.Cells(RC, I).Text '11-15-13
                        End If
                    Next
                    If LevASummary = "HoldOrder" Then '02-28-13
                        'I += 1
                        objSheets(3).Cells(RC + 2, I + 1) = RT.Cells(RC, I).Text '11-15-13
                    End If
                    'objSheet.Cells(RC, 2) = "Test" ' RetrCode : objSheet.Columns(2).ColumnWidth = 15
                    'objSheet.Cells(RC, 3) = ARCH : objSheet.Columns(3).ColumnWidth = 30
                    'objSheet.Cells(RC, 4) = ENG : objSheet.Columns(4).ColumnWidth = 30
                    'objSheet.Cells(RC, 5) = SLS1 : objSheet.Columns(5).ColumnWidth = 5
                    'objSheet.Cells(RC, 6) = SLSQ : objSheet.Columns(6).ColumnWidth = 5
                    'objSheet.Cells(RC, 7) = Distributor : objSheet.Columns(7).ColumnWidth = 12
                    'objSheet.Cells(RC, 8) = Contact : objSheet.Columns(8).ColumnWidth = 30
                    'objSheet.Cells(RC, 9) = CONTR : objSheet.Columns(9).ColumnWidth = 30
                    'objSheet.Cells(RC, 10) = LeadTime : objSheet.Columns(10).ColumnWidth = 30
                    'objSheet.Cells(RC, 11) = Sell : objSheet.Columns(11).ColumnWidth = 10
                    'objSheet.Cells(RC, 12) = StBidDate : objSheet.Columns(12).ColumnWidth = 10

                    RC += 1 : PC = 0  ''Print Line & Zero Line & Mfg
                Next 'End Of Section
                'HoldOrderSection ***HoldOrder*****HoldOrder******HoldOrder*****HoldOrder*****HoldOrder************************************************************
                TmpStrSql = "DROP TABLE IF EXISTS TMPREPORTS1 " '01-28-10
                mysqlcmd.CommandText = TmpStrSql : mysqlcmd.ExecuteNonQuery()
                If LevASummary = "DoneWithA" Then
                    If ForecastAllMfg = True Then GoTo EndHoldOrder '05-14-15 JTC '05-14-15 JTC Forecasting for MFGs Except Philips and SESCO
                    '03-24-14 EndTimeM = Format(Now, "hh:mm tt") '02-27-14
                    '03-24-14 Step2 = DateDiff(DateInterval.Minute, StartTimeM, EndTimeM)
                    '03-24-14 StartTimeM = Format(Now, "hh:mm tt")
                    LevASummary = "HoldOrder" 'objSheets(3)
                    GtExtQlSell = 0 : ExtQlSell = 0 : RC += 1 : PC = 0 '11-03-13 Zero Grand Total
                    RC = 0 '11-15-13 For Excel Sheet 2
                    '11-14-13 Add Status Column  objSheets(2).Cells(
                    For I = 0 To 12 : RT.Cells(RC, I).Text = "" : Next '02-28-14
                    'For I = 0 To 11 : objSheets(3).Cells(RC + 2, I + 1) = RT.Cells(RC, I).Text : Next : RC += 1 '11-03-13
                    'For I = 0 To 11 : RT.Cells(RC, I).Text = "HOLD" : Next '11-03-13
                    'RT.Cells(RC, 0).Text = "H" : RT.Cells(RC, 1).Text = "HO" : RT.Cells(RC, 9).Text = " ** Hold Orders **"
                    'For I = 0 To 11 : objSheets(3).Cells(RC + 2, I + 1) = RT.Cells(RC, I).Text : Next:    RC += 1 '11-03-13
                    '11-15-13 Hold Header((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((
                    objSheets(3).Cells(RC + 1, 1) = "H" ' Rep Forecasting Report" "Hold Orders By Brand"
                    objSheets(3).Cells(RC + 1, 2) = "H-O" ''Qty"
                    objSheets(3).Cells(RC + 1, 3) = "Rep#" '  'Type
                    objSheets(3).Cells(RC + 1, 4) = "Rpt-Date" '  
                    objSheets(3).Cells(RC + 1, 5) = "EnterDate" '                     'objSheets(3).Cells(RC + 1, 5).
                    objSheets(3).Cells(RC + 1, 6) = "PO Number" '  
                    objSheets(3).Cells(RC + 1, 7) = "HOLD" '  '11-14-13 Add Status Column 
                    objSheets(3).Cells(RC + 1, 8) = "MFG" ' or Job-Name"
                    objSheets(3).Cells(RC + 1, 9) = "Ext-Sell" '  
                    objSheets(3).Cells(RC + 1, 10) = "Description"
                    objSheets(3).Cells(RC + 1, 11) = "U-Qty"
                    objSheets(3).Cells(RC + 1, 12) = "U-Sell"
                    objSheets(3).Cells(RC + 1, 13) = "Job Name" '02-28-14

                    '05-29-14 - somehow headers are not on sesco's report - reset them
                    '06-13-14 JTC added s(1) need Sheet numberto first Sheet to fix Heading not showing on Forecast report
                    'objSheet.Cells(1, 1) = "ABC" ' Rep Forecasting Report"
                    'objSheet.Cells(1, 2) = "Q-O" ''Qty"
                    'objSheet.Cells(1, 3) = "Rep#" '  'Type
                    'objSheet.Cells(1, 4) = "Rpt-Date" '  
                    'objSheet.Cells(1, 5) = "Deliv-Date" '  
                    'objSheet.Cells(1, 6) = "Qut-Num" '  RT.Cols(0).Width = ".75in" : RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Center 'LnCode 
                    'objSheet.Cells(1, 7) = "Status" '  '11-14-13 Add Status Column 
                    'objSheet.Cells(1, 8) = "MFG-Q#" ' SourceQuote #
                    'objSheet.Cells(1, 9) = "Ext-Sell" '  
                    'objSheet.Cells(1, 10) = "Description"
                    'objSheet.Cells(1, 11) = "U-Qty"
                    'objSheet.Cells(1, 12) = "U-Sell"

                    RC += 1 '11-03-13
                    '(((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((
                    'strSql = "SELECT O.*, CONCAT(O.RELHOLD,O.OPENCLOSED,' ',O.STOCKJOB,O.SELECTCODE) AS RC_SS, OS.SLSCODE, OS.SLSSPLIT FROM ORDERMASTER O  LEFT JOIN ORDSLSSPLIT OS ON O.ORDERID = OS.ORDERID WHERE (OS.SLSNUMBER = 1) AND (CONCAT(O.BUYSELLAB, O.BUYSELLSR)<> 'BS') AND O.ENTRYDATE >= '2009-10-01' AND  O.ENTRYDATE <= '2013-10-31' and O.RelHold = 'R'  ORDER BY O.MFG"
                    'Select  OL.*, O.MFG as TmpMFG, O.CustCode as CustCode, O.CustName as FirmName, O.OpenClosed, O.JobName, O.RelHold,  O.StockJob, (OL.Sell * OL.Qty) as ExtSell, (OL.Comm * OL.Qty) as ExtComm, OS.SLSCode as SLSCode2 from ORDERMASTER O  left join ordslssplit OS on O.OrderID = OS.OrderID JOIN projectlines OL ON O.OrderID = OL.OrderID where OS.slsnumber = 1  and OL.OEDate >= '2009-10-01' and OL.OEDate <= '2013-10-31'  and OL.Active = '1' and OL.LnCode <> 'NPN' and OL.LnCode <> 'NTE'  and OL.Active = '1' and OL.LnCode <> 'NTE'  and OL.LnCode <> 'SUB'  and OL.LnCode <> 'BTX' and OL.LnCode <> 'TXL'  and OL.LnCode <> 'TXS'  and OL.LnCode <> 'TXF'  and OL.LnCode <> 'TAX'  and OL.Description <> ''  and OL.MFG <> ''  order by OL.MFG, OL.Description, OL.OEDate "
                    Dim NewDate As Date = Now '11-03-13
                    NewDate = NewDate.AddYears(-3) '=SubtractYears AddYears : NewDate = NewDate.AddDays(-1) Me.DTPicker1StartBid.Value = NewDate '02-03-12 'A = Format(Now.YearaDDyear(1).Year + 1, "yyyy")
                    sStartBidDate = VB6.Format(NewDate, "yyyy-MM-dd") ''11-03-13 JTC Include HoldOrder within the last 3 years
                    ' SaveStrSQL = " and O.EntryDate >= '" & sStartBidDate & "' "'11-03-13 Added to end of Sql
                    '11-14-13 Add Status Column After OrderNumber  O.MFGSubGroup as status '11-15-13 add O.EntryDate,
                    '02-28-14 JTC Put Hold Orders in Tmp (((((((((((((((((((((((((((((((((((((((((((((((((((((
                    strSql = "Select  O.EntryDate, O.OrderNumber, O.JobName, O.ORDERID from OrderMaster O where O.RelHold = 'H' and O.EntryDate >= '" & sStartBidDate & "' " '
                    TmpStrSql = "DROP TABLE IF EXISTS TMPREPORTS1 " '01-28-10
                    mysqlcmd.CommandText = TmpStrSql : mysqlcmd.ExecuteNonQuery()
                    SaveStrSQL = "CREATE TEMPORARY TABLE TMPREPORTS1 AS " & strSql '10-22-13
                    mysqlcmd.CommandText = SaveStrSQL : SubCount = mysqlcmd.ExecuteNonQuery() 'Get All PHIL or COOP
                    'strSql = "Select * from TMPREPORTS1 QL"
                    'strSql = Replace(SaveStrSQL2, "= projectcust.QuoteID", "= projectcust.QuoteID inner join TMPREPORTS TR ON TR.QuoteID = QUOTE.QUOTEID ")
                    'strSql = "Select  TR.*, OL.MFG, OL.Description, OL.Sell, OL.UM from TMPREPORTS1 TR left join ProjectLines OL  ON TR.OrderID = OL.OrderID " '02-28-14 , Sum(OL.Qty) as Qty, Sum(OL.Sell * OL.Qty) as ExtSell 
                    strSql = "Select  TR.*, OL.MFG, OL.Description, OL.Active, OL.LnCode, OL.Qty, OL.Sell, OL.UM , Sum(OL.Qty), Sum(OL.Sell * OL.Qty) as ExtSell from TMPREPORTS1 TR inner join ProjectLines OL  ON TR.OrderID = OL.OrderID " '02-28-13 GROUP BY TR.OrderNumber, OL.MFG, OL.Description order by TR.OrderNumber, OL.MFG, OL.Description
                    'strSql = "Select  OL.MFG, OL.Description, OL.Sell, OL.UM, O.EntryDate, O.OrderNumber, O.MFGSubGroup as Status, O.JobName, O.RelHold, Sum(OL.Qty) as Qty, Sum(OL.Sell * OL.Qty) as ExtSell from OrderMaster O LEFT JOIN ProjectLines OL ON O.ORDERID = OL.OrderID inner join TMPREPORTS1 TR ON TR.QuoteID = QUOTE.QUOTEID "
                    'TmpStrSql = "DROP TABLE IF EXISTS TMPREPORTS1 " '01-28-10
                    ' mysqlcmd.CommandText = TmpStrSql : mysqlcmd.ExecuteNonQuery()
                    '))))))))))))))))))))))))))))))))))))))))))))))))))))
                    'B/4 TEMPREPORTS1 strSql = "Select  OL.MFG, OL.Description, OL.Sell, OL.UM, O.EntryDate, O.OrderNumber, O.MFGSubGroup as Status, O.JobName, O.RelHold, Sum(OL.Qty) as Qty, Sum(OL.Sell * OL.Qty) as ExtSell from OrderMaster O LEFT JOIN ProjectLines OL ON O.ORDERID = OL.OrderID  where O.RelHold = 'H' and OL.Active = '1' and OL.LnCode <> 'NPN' and OL.LnCode <> 'NTE'  and OL.Active = '1' and OL.LnCode <> 'NTE'  and OL.LnCode <> 'SUB'  and OL.LnCode <> 'BTX' and OL.LnCode <> 'TXL'  and OL.LnCode <> 'TXS'  and OL.LnCode <> 'TXF'  and OL.LnCode <> 'TAX'  and OL.Description <> ''  and OL.MFG <> '' and O.EntryDate >= '" & sStartBidDate & "' " '
                    'STR1 = strSql.Substring(0, strSql.IndexOf("where")) '"Select  QL.MFG, QL.Description, QL.Sell, QL.UM, Q.QuoteCode, Q.JobName, Q.Status, Sum(QL.Qty) as Qty, Sum(QL.Sell * QL.Qty) as ExtSell from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  "
                    '10-20-13 .txtQutRealCode.Text = PhilBrands
                    If BrandList <> "" Then '10-20-13 BrandList=DAYB,CAPR,MCPH,OMEG,LAM,CHLO,MORL,GUTH,ARDE,CRES,GARD,EMCO,FORE,HADC,EXCE,TRAN,LOL,ALKC,LEDA,LUME,CRES,ALLS,HANO,THMO,THMI"
                        'Dim STR2 As String = " " & strSql.Substring(strSql.IndexOf("order by"))
                        If BrandList.Contains(",") = True Then  '10-19-13 JTC Added QL.QL.MFG to Line Items
                            BC = " and ( OL.MFG = '" & BrandList.Replace(",", "' or OL.MFG = '") & "' ) " '10-20-13 No Blanks or QL.MFG = ''  )"
                        Else
                            BC = " and ( OL.MFG = '" & BrandList & "'" & " ) " '10-20-13 No Blanks or QL.MFG = '')"
                        End If '" and ( QL.MFG = 'DAYB' or QL.MFG = 'CAPR' or QL.MFG = 'MCPH' or QL.MFG = 'OMEG' or QL.MFG = 'LAM' or QL.MFG = 'CHLO' or QL.MFG = 'MORL' or QL.MFG = 'GUTH' or QL.MFG = 'ARDE' or QL.MFG = 'CRES' or QL.MFG = 'GARD' or QL.MFG = 'EMCO' or QL.MFG = 'FORE' or QL.MFG = 'HADC' or QL.MFG = 'EXCE' or QL.MFG = 'TRAN' or QL.MFG = 'LOL' or QL.MFG = 'ALKC' or QL.MFG = 'LEDA' or QL.MFG = 'LUME' or QL.MFG = 'CRES' or QL.MFG = 'ALLS' or QL.MFG = 'HANO' or QL.MFG = 'THMO' or QL.MFG = 'THMI' or QL.MFG = ''  )"
                        strSql += BC 'Add Brands of Line Items OL.MFG <> '' and OL.Qty <> ''
                        strSql += " and OL.LnCode <> 'NPN' and OL.LnCode <> 'NTE'  and OL.Active = '1' and OL.LnCode <> 'SUB'  and OL.LnCode <> 'BTX' and OL.LnCode <> 'TXL'  and OL.LnCode <> 'TXS'  and OL.LnCode <> 'TXF'  and OL.LnCode <> 'TAX'  and OL.Description <> ''  and OL.MFG <> '' "
                        LevASummary = "HoldOrder" '10-22-13 order by OL.MFG, OL.Description, order by OL.MFG, OL.Description,
                        strSql += " GROUP BY TR.OrderNumber, OL.MFG, OL.Description order by TR.OrderNumber, OL.MFG, OL.Description"
                    End If
                    'Test Below Select  OL.MFG, OL.Description, OL.Sell, OL.UM, O.OrderNumber, O.JobName, O.entrydate, O.RelHold, Sum(OL.Qty) as Qty, Sum(OL.Sell * OL.Qty) as ExtSell from OrderMaster O LEFT JOIN ProjectLines OL ON O.ORDERID = OL.OrderID  where o.entryDate < '2013-08-01' and O.RelHold = 'R' and OL.Active = '1' and OL.LnCode <> 'NPN' and OL.LnCode <> 'NTE'  and OL.Active = '1' and OL.LnCode <> 'NTE'  and OL.LnCode <> 'SUB'  and OL.LnCode <> 'BTX' and OL.LnCode <> 'TXL'  and OL.LnCode <> 'TXS'  and OL.LnCode <> 'TXF'  and OL.LnCode <> 'TAX'  and OL.Description <> ''  and OL.MFG <> ''  and ( OL.MFG = 'DAYB' or OL.MFG = 'CAPR' or OL.MFG = 'OMEG' or OL.MFG = 'MCPH' or OL.MFG = 'MORL' or OL.MFG = 'LAML' or OL.MFG = 'CHLO' or OL.MFG = 'GUTH' or OL.MFG = 'ARDE' or OL.MFG = 'CRST' or OL.MFG = 'EXCE' or OL.MFG = 'HADC' or OL.MFG = 'LOL' or OL.MFG = 'TRAN' or OL.MFG = 'GARD' or OL.MFG = 'ALKC' or OL.MFG = 'ALLS' or OL.MFG = 'BRON' or OL.MFG = 'CMT' or OL.MFG = 'FORE' or OL.MFG = 'HANO' or OL.MFG = 'LEDA' or OL.MFG = 'LIGH' or OL.MFG = 'LUME' or OL.MFG = 'THOM' or OL.MFG = 'WIDE' or OL.MFG = 'PHIL' ) GROUP BY O.OrderNumber, OL.MFG, OL.Description order by O.OrderNumber, OL.MFG, OL.Description
                    'Test strSql = "Select  OL.MFG, OL.Description, OL.Sell, OL.UM, O.OrderNumber, O.JobName, O.RelHold, Sum(OL.Qty) as Qty, Sum(OL.Sell * OL.Qty) as ExtSell from OrderMaster O LEFT JOIN ProjectLines OL ON O.ORDERID = OL.OrderID  where O.RelHold = 'H' and OL.Active = '1' and OL.LnCode <> 'NPN' and OL.LnCode <> 'NTE'  and OL.Active = '1' and OL.LnCode <> 'NTE'  and OL.LnCode <> 'SUB'  and OL.LnCode <> 'BTX' and OL.LnCode <> 'TXL'  and OL.LnCode <> 'TXS'  and OL.LnCode <> 'TXF'  and OL.LnCode <> 'TAX'  and OL.Description <> ''  and OL.MFG <> ''  and ( OL.MFG = 'DAYB' or OL.MFG = 'CAPR' or OL.MFG = 'OMEG' or OL.MFG = 'MCPH' or OL.MFG = 'MORL' or OL.MFG = 'LAML' or OL.MFG = 'CHLO' or OL.MFG = 'GUTH' or OL.MFG = 'ARDE' or OL.MFG = 'CRST' or OL.MFG = 'EXCE' or OL.MFG = 'HADC' or OL.MFG = 'LOL' or OL.MFG = 'TRAN' or OL.MFG = 'GARD' or OL.MFG = 'ALKC' or OL.MFG = 'ALLS' or OL.MFG = 'BRON' or OL.MFG = 'CMT' or OL.MFG = 'FORE' or OL.MFG = 'HANO' or OL.MFG = 'LEDA' or OL.MFG = 'LIGH' or OL.MFG = 'LUME' or OL.MFG = 'THOM' or OL.MFG = 'WIDE' or OL.MFG = 'PHIL' )  GROUP BY O.OrderNumber, OL.MFG, OL.Description order by O.OrderNumber, OL.MFG, OL.Description "

                    'SaveStrSQL = "CREATE TEMPORARY TABLE TMPREPORTS1 AS " & strSql '10-22-13

                    'mysqlcmd.CommandText = SaveStrSQL : SubCount = mysqlcmd.ExecuteNonQuery() 'Get All PHIL or COOP
                    'strSql = "Select * from TMPREPORTS1 OL "
                    mysqlcmd.CommandText = strSql : SubCount = mysqlcmd.ExecuteNonQuery() 'Get All PHIL or COOP
                    GoTo FillDataSetHere 'Goto Hold Data
                End If
EndHoldOrder:   '05-14-15 JTC
                'End HoldOrderSection *****************************************************************************************
                'Start B Quotes by Brand (((((((((((((((((((((((((((((((((((((((((((((((((((((((((((
                If LevASummary = "A" Then 'A = Total by Quote Code B = Total by 
                    '03-24-14 EndTimeM = Format(Now, "hh:mm tt") '02-27-14
                    '03-24-14 Step1 = DateDiff(DateInterval.Minute, StartTimeM, EndTimeM)
                    '03-24-14 StartTimeM = Format(Now, "hh:mm tt")
                    LevASummary = "DoneWithA"
                    'Print Grand Total here for If LevASummary = "A"
                    ''07-24-14 JTC 
                    If SESCO = True Then GtExtQlSell = 0 : ExtQlSell = 0
                    For PC = 0 To 9
                        RT.Cells(RC, PC).Text = "" '10-23-13 Format(Val(drQutLn("Sell")), "0")
                    Next
                    RT.Cells(RC, 8).Text = Format(Val(GtExtQlSell), "0")
                    RT.Cells(RC, 9).Text = "Grand Total Quotes"
                    RT.Rows(RC).Style.BackColor = AntiqueWhite '  Else RT.Rows(PC).Style.BackColor = Color.White '10-21-13
                    For I = 0 To 10 ' 'Excel Starts with 1 not Zero !@#$%^&*
                        objSheets(1).Cells(RC + 2, I + 1) = RT.Cells(RC, I).Text '11-03-13  
                    Next
                    'Done With objSheets(1)
                    RC += 1 : PC = 0
                    RC = 0 '11-15-13 For Excel Sheet 2
                    GtExtQlSell = 0
                    'Doing B    BBBBBBBBB
                    'LevBStrSQl = STR1 & BC & " Q.QuoteCode, QL.MFG, QL.Description order by Q.QuoteCode, QL.MFG, QL.Description "
                    TmpStrSql = "DROP TABLE IF EXISTS TMPREPORTS1 " '01-28-10
                    mysqlcmd.CommandText = TmpStrSql : mysqlcmd.ExecuteNonQuery()
                    SaveStrSQL = "CREATE TEMPORARY TABLE TMPREPORTS1 AS " & LevBStrSQl  '10-21-13
                    'strSql = LevBStrSQl ' 10-21-13
                    mysqlcmd.CommandText = SaveStrSQL : SubCount = mysqlcmd.ExecuteNonQuery() 'Get All PHIL or COOP
                    strSql = "Select * from TMPREPORTS1 QL"
                    dsQuote = New dsSaw8 : dsQuote.EnforceConstraints = False
                    daQuoteLine = New MySqlDataAdapter
                    daQuoteLine.SelectCommand = New MySqlCommand(strSql, myConnection)
                    'Dim cbQutLin as MySql.Data.MySqlClient.MySqlCommandBuilder
                    cbQutLin = New MySqlCommandBuilder(daQuoteLine)
                    daQuoteLine.Fill(dsQuote, "quotelines")
                    RT = New RenderTable '12-18-14 - EMPTY THIS, THE SECOND TAB HAS THE DELIVER DATE FROM A DIFFERENT ROW IN IT 
                    GoTo SecondLoop ' Up FillDataSetHere
                End If
                For I = 1 To 3
                    objSheets(I).Columns(1).Columnwidth = 6
                    objSheets(I).Columns(1).Columnwidth = 6 '".4"
                    objSheets(I).Columns(2).ColumnWidth = 6 '".4"
                    objSheets(I).Columns(3).ColumnWidth = 12 '".9" 'Rep
                    objSheets(I).Columns(4).ColumnWidth = 12 '".9" 'Date
                    objSheets(I).Columns(5).ColumnWidth = 12 ' ".9" 'date
                    objSheets(I).Columns(6).ColumnWidth = 15 ' "1" 'Quote
                    '11-14-13 Add Status Column 
                    objSheets(I).Columns(7).ColumnWidth = 8 'Status
                    objSheets(I).Columns(8).ColumnWidth = 8 ' ".6" 'brand
                    objSheets(I).Columns(9).ColumnWidth = 10 ' ".8" 'ExtSell
                    objSheets(I).Columns(10).ColumnWidth = 25 ' "2" 'Desc
                    objSheets(I).Columns(11).ColumnWidth = 10 ' = ".8" 'Qty
                    objSheets(I).Columns(12).ColumnWidth = 10 ' ".8" 'Sell
                    objSheets(I).Columns(6).HorizontalAlignment = 2 'Left 11-17-13  XlVAlign.xlVAlignCenter -4108 ' xlWs.cells(r, 3).HorizontalAlignment = xlCenter'.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                Next
                '''''''''''''''''''''''''''''''''''''''''''''''''
                'Print Grand Total here 
                For PC = 0 To 10
                    RT.Cells(RC, PC).Text = "" '10-23-13 Format(Val(drQutLn("Sell")), "0")
                Next
                ''07-24-14 JTC If SESCO = True then delete Sell ExtSell & Unit Sell GtExtQlSell = 0 : ExtQlSell = 0
                If SESCO = True Then GtExtQlSell = 0 : ExtQlSell = 0
                RT.Cells(RC, 8).Text = Format(Val(GtExtQlSell), "0")
                RT.Cells(RC, 9).Text = "Grand Total Hold Orders"
                RT.Rows(RC).Style.BackColor = AntiqueWhite '  Else RT.Rows(PC).Style.BackColor = Color.White '10-21-13
                '11-14-13 Add Status Column 
                For I = 0 To 12 '02-28-14  'Excel Starts with 1 not Zero !@#$%^&*
                    objSheets(3).Cells(RC + 2, I + 1) = RT.Cells(RC, I).Text '11-03-13  
                Next
                'ActiveCell.row = 1
                'Dim range1 As objExcel.range ''// set range location eg row 5  "5:5"  'range1.Insert(objExcel.xlShiftDown, objExcel.xlFormatFromRightOrBelow)
                ''objRange = objSheet.Cells(1, 1).EntireRow  ''objSheet.Cells(1, 1).EntireRow()  'objRange.Insert(-4121) ' xlShiftDown)
                RC += 1 : PC = 0
                GtExtQlSell = 0
                'Forecast End
                '10-04-13 'Insert Header Forecast ******************************************************
                '11-04-13 Move to Front of Reoprt
                RT.Rows.Insert(0, 2) '12-15-11 2 to 301-16-09 Insert Header
                RT.RowGroups(0, 1).Style.BackColor = LemonChiffon '11-12-10
                RT.Cells(0, 1).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
                RT.Cells(0, 0).SpanCols = 11 ' RT.Cols.Count '/ 2 '12-30-08
                RT.RowGroups(1, 1).Style.BackColor = LemonChiffon '11-12-10
                RT.Cells(0, 0).Text = "Philips Rep Forecast Report    Run Date = " & VB6.Format(Now, "Short Date") & Space(4) & "" & Space(8) & "Page [PageNo] of [PageCount]    " '07-02-09
                RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
                RT.Cells(0, 0).Style.BackColor = LemonChiffon '11-12-10
                RT.Cells(0, 0).Style.FontSize = 14
                RT.Cells(1, 0).Text = "ABC"
                RT.Cells(1, 1).Text = "Q-O" 'Qut-Order
                RT.Cells(1, 2).Text = "Rep#" '  
                RT.Cells(1, 3).Text = "Rpt-Date" '  
                RT.Cells(1, 4).Text = "Deliv-Date" '  
                RT.Cells(1, 5).Text = "Qut-Num" '  
                '11-14-13 Add Status Column 
                RT.Cells(1, 6).Text = "Status" ' or Job-Name"
                RT.Cells(1, 7).Text = "MFG" ' or Job-Name"
                RT.Cells(1, 8).Text = "Ext-Sell" '  
                RT.Cells(1, 9).Text = "Description"
                RT.Cells(1, 10).Text = "U-Qty"
                RT.Cells(1, 11).Text = "U-Sell"

                PC = 0 : RT.Cols(PC).Width = ".4"
                PC = 1 : RT.Cols(PC).Width = ".4"
                PC = 2 : RT.Cols(PC).Width = ".9" 'Rep
                PC = 3 : RT.Cols(PC).Width = ".9" 'Date
                PC = 4 : RT.Cols(PC).Width = ".9" 'date
                PC = 5 : RT.Cols(PC).Width = "1" 'Quote
                '11-14-13 Add Status Column 
                PC = 6 : RT.Cols(PC).Width = ".6" 'Status
                PC = 7 : RT.Cols(PC).Width = ".6" 'brand
                PC = 8 : RT.Cols(PC).Width = ".8" 'ExtSell
                PC = 9 : RT.Cols(PC).Width = "2" 'Desc
                PC = 10 : RT.Cols(PC).Width = ".8" 'Qty
                PC = 11 : RT.Cols(PC).Width = ".8" 'Sell
                For I = 0 To PC : RT.Cells(RC, I).Style.BackColor = LemonChiffon : Next
                RT.Rows(RC).Style.BackColor = LemonChiffon '11-19-10 
                RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
                'Const xlShiftDown = -4121
                'Dim objRange = objSheets.range("A1:A9").EntireRow
                'objRange.Insert(xlShiftDown)
                If ForecastAllMfg = True Then '05-14-15 JTC '05-14-15 JTC Forecasting for MFGs Except Philips and SESCO
                    '05-14-15 JTC objBook.WorkSheets.Delete()
                    For Each objSheets In objExcel.Worksheets
                        If objSheets.Name = "Hold Orders By Brand" Then
                            objSheets.delete()
                        End If
                    Next '
                End If
                If frmQuoteRpt.chkBrandReport.CheckState = CheckState.Checked And frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And frmQuoteRpt.txtPrimarySortSeq.Text = "Forecasting" Then
                    '11-15-13 Don't Display the report See Excel
                    GoTo NoDisplayForecastRpT
                End If
                If ForecastAllMfg = True Then GoTo NoDisplayForecastRpT '05-14-15 JTC '05-14-15 JTC Forecasting for MFGs Except Philips and SESCO

                RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
                RT.Style.GridLines.All = LineDef.Default
                doc.Body.Children.Add(RT) '12-29-06
609022:         ppv.C1PrintPreviewControl1.Document = doc
                'ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)
                ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1 '12-12-08
                ppv.Doc.Generate() '11-18-08
                ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
                ppv.Show() '12-06-12 JTc Moved down
                ppv.BringToFront()
                frmShowHideGrid.BringToFront() '03-10-09
                '10-22-13 Save Excel File name  update Quote q set q.EstDelivDate = '2013-04-14' where q.BidDate > '2013-04-01' and q.BidDate < '2013-04-30';
NoDisplayForecastRpT:  'SELECT q.QuoteCode, q.EntryDate, q.BidDate, q.EstDelivDate FROM saw8prud.`quote` q where q.BidDate > '2013-04-01' and q.BidDate < '2013-04-30'
                'FORECAST-RepCustNumber-StartQ.EstDelivDate-To:End-EstDelivDate.xls
                'FileName += ".csv" '06-17-10 
                'FileName = SaveDialog(FileName, "Export to CSV", "csv Files (*.csv)|*.csv")
                'If FileNaTrim = "" Then GoTo NoFileName '06-16-10 
                'Done Above FileName = UserPath & "DATA\OrderReport001.Xls"
                '06-30-13 Start File In Use *************************************************************************
                Dim FileNum As Short = 0 '05-15-15 JTC 3
                '05-20-15 JTC Alkways Save a Report even if No RepCustNumber If RepCustNumber Is Nothing Then GoTo NoFileName '05-14-15 JTC
                If RepCustNumber Is Nothing Then RepCustNumber = VB6.Format("12345", "00000") & "  " '05-14-15 JTC 
                If RepCustNumber.Trim = "" Then RepCustNumber = VB6.Format("12345", "00000") & "  " '&02-28-14 Format(Now, "yyyy-MM-dd") '10-04-13 
                Dim FileNumStr As String = VB6.Format(RepCustNumber, "00000") & "-" '01-16-14 & Format(Now, "yyyy-MM-dd") '10-04-13 

                'FileNumStr = VB6.Format(RepCustNumber, "000") & "-" & Format(Now, "yyyy-MM-dd") '10-04-13 
                Dim Range As String = VB6.Format(SaveStartDate, "MM-dd-yy") & " Thru " & VB6.Format(sEndBidDate, "MM-dd-yy") '05-28-14 SaveStartDate

                'Dim sStartBidDate As String = VB6.Format(frmQuoteRpt.DTPicker1StartBid.Value, "yyyy-MM-dd") '10-23-13 Q.biddate is<> null or Q.biddate <>
                'If RepCustNumber.Trim = "" Then FileNumStr = VB6.Format("12345678", "000") & "  " & Format(Now, "yyyy-MM-dd") '10-04-13 

                '05-28-14Dim FileName As String = UserPath & "DATA\RepForecastRpt-" & FileNumStr & Range & "Run" & Format(Now, "yyyy-MM-dd") & ".Xls" '01-16-14 08-14-12 JTC Added UserID to file Name
                Dim FileName As String = UserPath & "DATA\RepForecastRpt-" & FileNumStr & Range & " Run " & sEndBidDate & ".Xls" '01-16-14 08-14-12 JTC Added UserID to file Name
StartFileN:
                If My.Computer.FileSystem.FileExists(FileName) Then
                    FileNumStr = Left(FileNumStr, Len(FileNumStr) - 1) + VB6.Format(FileNum + 1) : FileNum = FileNum + 1 '05-15-15 
                    FileName = UserPath & "DATA\RepForecastRpt-" & FileNumStr & Range & " Run " & sEndBidDate & "-2.Xls" '06-12-14 JTC Try Version-2 08-14-12 JTC Added UserID to file Name
                    If My.Computer.FileSystem.FileExists(FileName) Then
                        GoTo StartFileN '  My.Computer.FileSystem.DeleteFile(FileName)
                    End If
                End If
                'Start Delete old Report files
                If My.Computer.FileSystem.DirectoryExists(UserPath & "DATA\") = True Then
                    For Each fName As String In Directory.GetFiles(UserPath & "DATA\", "RepForecast*") '10-04-13  Only get files that begin with RepForecast
                        Dim fdate As Date = File.GetLastWriteTime(fName)
                        If Date.Compare(fdate, Now.AddDays(-45)) < 0 Then File.Delete(fName) 'Delete Older Than 45 days
                    Next
                End If 'End Delete old report files 
                Try '08-14-12 
                    If My.Computer.FileSystem.FileExists(FileName) Then My.Computer.FileSystem.DeleteFile(FileName)
                Catch myException As Exception '08-14-12 
                    Resp = MsgBox("Export File " & FileName & vbCrLf & "is in use. You must close that Excel Application fist and try again.", MsgBoxStyle.Critical + MsgBoxStyle.YesNo + RealWithOneMfgCustCode.Trim, "VQRT Export to Excel") '10-04-13 : GoTo Exit_Done '10-04-13 Exit Sub 'in Use
                    If Resp = vbYes Then GoTo StartFileN
                    GoTo Exit_Done '10-04-13 
                End Try
                objBook.SaveAs(FileName, 39, , , False, False, , False, False, , , ) '39 MessageBox.Show("All Done   Created " & FileName)
                objExcel.VISIBLE = True 'Show User the Excel Report
                'Clean Up
                objSheet = Nothing
                objExcel = Nothing
                '03-24-14 EndTimeM = Format(Now, "hh:mm tt")
                '03-24-14 Step3 = DateDiff(DateInterval.Minute, StartTimeM, EndTimeM)
                '03-24-14 MsgBox("Step1= " & Step1 & "  Step2= " & Step2 & "  Step3= " & Step3 & " Minutes by Step" & vbCrLf & "PLease view the RepForecast Excel report file to send to Philips. ") '02-27-14 
NoFileName:
                If frmQuoteRpt.txtPrimarySortSeq.Text = "Forecasting" And frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" Then
                    'If My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") Then '07-24-14 JTC SESCO Forecasting No dollars
                    SESCO = False '  End If
                End If
                MsgBox("All Done. PLease view the RepForecast Excel report file to send to your Factory. ") '05-15-15 JTC
                GoTo ExitReportLoop
                '12-28-13 Moved Down
            End If '10-21-13 ******************************************************************************************************************************************
            'End Forecast %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
6315:

            frmShowHideGrid.tgShow.SetDataBinding(table, "")
            'If dsGrid Is Nothing Then Exit Sub

            Dim TGNameStr As String = "" 'Documentation Set Up a String of Names
            Dim TGWidthStr As String = "" 'Set Up a String of Widths

            MaxCol = frmQuoteRpt.tgln.Splits(0).DisplayColumns.Count - 1
            Dim PrtCols As Int16 = 0

            'Header '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '10-20-13Dim RArea As C1.C1Preview.RenderArea = New C1.C1Preview.RenderArea
            'Type of Report - & Agency Name
            RT = New C1.C1Preview.RenderTable
            RT.Rows.Insert(0, 4) '01-16-09 Insert Header
            RT.RowGroups(0, 4).Header = True
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10 RT.SplitHorzBehavior = True '05-26-10 Test
            'RT.SplitVertBehavior = True '05-26-10 Test
            RT.Width = "auto"
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            RT.Style.GridLines.All = LineDef.Empty '  LineDef.Default   '12-04-10 & "  UserID = " & UserID 
            ' Dim AAA As String = "Report: Quote " & frmQuoteRpt.pnlTypeOfRpt.Text.Trim & "  UserID = " & UserID & "    Report Date = " & Format(Now, "MM/dd/yyyy")
            RT.Cells(0, 0).Text = "Report: Quote " & frmQuoteRpt.pnlTypeOfRpt.Text.Trim & "  UserID = " & UserID & "    Report Date = " & Format(Now, "MM/dd/yyyy") '11-20-10 
            'RT.Cells(0, 1).Text = AGnam : RT.Cells(0, 1).Style.TextAlignHorz = AlignHorzEnum.Right
            RT.Cells(0, 0).SpanCols = 7 '09-20-12
            RT.Rows(0).Style.GridLines.All = LineDef.Empty
            RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
            '10-20-13 Dim fs As Integer = frmQuoteRpt.FontSizeComboBox.Text
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Bold)
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10RT.SplitHorzBehavior = True '05-26-10 Test
            'RT.SplitVertBehavior = True '05-26-10 Test
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            'RT.Style.GridLines.All = LineDef.Empty '  LineDef.Default
            Dim sort As String = "Primary Sort: " & frmQuoteRpt.txtPrimarySortSeq.Text
            If frmQuoteRpt.txtSecondarySort.Text.Trim <> "" Then sort = sort + " Secondary Sort: " & frmQuoteRpt.txtSecondarySort.Text
            RT.Rows(1).Style.TextAlignHorz = AlignHorzEnum.Left
            RT.Rows(1).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs)
            RT.Cells(1, 0).Text = sort & "           Page [PageNo] of [PageCount]"
            'RT.Cells(1, 0).Text = "Page [PageNo] of [PageCount]"
            ' RT.Cells(1, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Right
            RT.Cells(1, 0).SpanCols = 7 '09-20-12
            RT.Rows(1).Style.GridLines.All = LineDef.Empty
            RT.Cells(2, 0).Text = "Select Criteria: " & SelectionText 'frmQuoteRpt.TtxtSortSelV.Text '07-02-09frmProjRpt.txtPrimarySortSeq.Text & " " & frmProjRpt.txtSecondarySort.Text '07-01-09
            RT.Cells(2, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
            RT.Cells(2, 0).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, 9)  '04-30-10 jh - FONT COMBO
            RT.Cells(2, 0).SpanCols = 7 '09-20-12
            RT.Rows(2).Style.TextAlignHorz = AlignHorzEnum.Left
            RT.Rows(2).Style.GridLines.All = LineDef.Empty

            'RT.RowGroups(0, 1).Header = True
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            '03-19-14 RT.Style.BackColor = LemonChiffon
            RC = 0
            'If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then '09-19-12  GoTo QutLineHistoryRpt
            '    RT.Cells(3, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
            '    RT.Cells(3, 0).Text = "LnCode"
            '    RT.Cells(3, 1).Text = "QTY" ''Qty"
            '    RT.Cells(3, 2).Text = "TYPE" '  'Type
            '    RT.Cells(3, 3).Text = "MFG"   'Mfg
            '    RT.Cells(3, 4).Text = "DESCRIPTION" ' 'desc
            '    RT.Cells(3, 5).Text = "COMM-$"   'COMM
            '    RT.Cells(3, 6).Text = "PAID Y/N"   'Cost

            '    RT.Cols(1).Width = ".75in" : RT.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Right  'Qty
            '    RT.Cols(2).Width = ".75in" : RT.Cols(2).Style.TextAlignHorz = AlignHorzEnum.Center  'Type
            '    RT.Cols(3).Width = ".75in" : RT.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Center  'Mfg
            '    RT.Cols(4).Width = "5in" : RT.Cols(4).Style.TextAlignHorz = AlignHorzEnum.Left 'desc
            '    RT.Cols(5).Width = ".75in" : RT.Cols(5).Style.TextAlignHorz = AlignHorzEnum.Right  'Cost
            '    RT.Cols(6).Width = "2.4in"
            '    RT.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Right 'Sell
            '    RT.Cols(0).Visible = False : RT.Cols(0).Width = Unit.Empty 'Don't print LnCode
            '    RT.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Right
            '    RT.Rows(3).Style.GridLines.All = LineDef.Default
            'End If '09-23-12

            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            RArea.Children.Add(RT)
            doc.Body.Document.PageLayout.PageHeader = RArea
            '09-23-12End If
            ''Dim RTSave As RenderTable = New RenderTable '09-20-12 
            'RTSave = New RenderTable : RTSave = RT.Clone '09-20-12


            'doc.Body.Children.Add(RT) ''02-04-12 Header on First Page Col hdg on All Pages
            'Dim RThdr As RenderTable = RT
            'END PAGE HEADER - DIFFERENT THAN THE TABLE HEADERS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'If DebugOn ThenDebug.Print(frmQuoteRpt.cboSortPrimarySeq.Text.ToString) 'MFG Sub-Totals in Catalog # Sequence
            If frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("Catalog # Totals By Month") Then
                frmQuoteRpt.fraSalesorCost.Visible = True '02-15-10
                frmQuoteRpt.fraUnitorExtended.Visible = True '02-15-10
                If Not DIST Then frmQuoteRpt.optSalesorCost_Cost.Text = "Cost Dollars" '02-15-10"
                Call PrtQutSpreadSheet(1) '09-09-09 ByRef PriceCol As Integer)
                GoTo Exit_Done '10-04-13 
            End If

            'If frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("Catalog # Detail Report") Then  '"Catalog # Detail Report - MFG/Cat # Sequence"'frmMenu.optRptByDate.Value Then
            'Print Column Headers *******************************************************************
            '06-14-10 RT = New C1.C1Preview.RenderTable
            '06-14-10 RT.Style.GridLines.All = LineDef.Default
            'Dim Headertmp As String = ""
            ' PrtCols = 0   'Print Column Headers
            'frmQuoteRpt.tgln.MoveFirst()
            'For I = 0 To frmQuoteRpt.tgln.Splits(0).DisplayColumns.Count - 1
            '    'Dim col2 As C1.Win.C1TrueDBGrid.C1DisplayColumn = frmQuoteRpt.tg.Splits(0).DisplayColumns(I) '02-20-09
            '    'Dim Tag As String = col2.DataColumn.Tag 'Debug.Print(tgShow(col.DataColumn.Tag, 0))'col.DataColumn.Tag = I.ToString
            '    TgWidth(I) = (frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Width / 100) '02-25-09
            '    If frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
            '    If (frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
            '    RT.Cells(0, PrtCols).Text = frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Name
            '    RT.Cols(PrtCols).Width = TgWidth(I) '02-22-09
            '    Headertmp = Headertmp & frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Name & TgWidth(I).ToString
            '    PrtCols += 1
            'Next
            'PrtCols -= 1
            '06-14-10 MaxCol = frmQuoteRpt.tgln.Splits(0).DisplayColumns.Count - 1
            '06-14-10 Call RTColSize(RT, MaxCol, TgWidth) '02-03-09 
            'RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
            'RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            'RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            'doc.Body.Children.Add(RT)
            'Header End 'Done with Headings *********************************************************************

            'REPORT BODY '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            RT = New C1.C1Preview.RenderTable
            RC = 0
            RT.Style.GridLines.All = LineDef.Default
            RT.CellStyle.Padding.Left = "1mm" '12-13-12
            RT.CellStyle.Padding.Right = "1mm" '12-13-12
            '10-20-13Dim C As Integer
            Dim X As String = "ZeroLevels" ' "AddAllLevels" ''01-26-09
            Call TotalsCalc(X, B, C) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            CurrLev1 = "" : PrevLev1 = "" : CurrLev2 = "" : PrevLev2 = "" : Cmd = "" 'Cmd = "EOF"
            Dim A As String = "PrintLine"
            frmQuoteRpt.tgln.UpdateData()
            Dim PrimarySortSeq As String = frmQuoteRpt.txtPrimarySortSeq.Text
            Dim SeconarySortSeq As String = frmQuoteRpt.cboSortSecondarySeq.Text
            Dim FirstLoop As Int16 = 0 '07-09-09
            Dim MFGSubTotals As Boolean = False '09-10-09 PrevMfg As String = "" '09-08-09
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then '09-23-12 
                MFGSubTotals = True '09-19-12  GoTo
PaidUnpaidBothAgain:
                Dim PaidUnpaidBoth As String = UCase(InputBox("Please enter P = Paid,  U = Unpaid or B = Both.", "P, U, B", "B")) ''09-23-12 
                If PaidUnpaidBoth = "P" Or PaidUnpaidBoth = "U" Or PaidUnpaidBoth = "B" Then  Else GoTo PaidUnpaidBothAgain
                If PaidUnpaidBoth = "B" Then frmQuoteRpt.cboLinesInclude.Text = "Include All Lines on Job"
                If PaidUnpaidBoth = "P" Then frmQuoteRpt.cboLinesInclude.Text = "Include Only Paid Items on the Job"
                If PaidUnpaidBoth = "U" Then frmQuoteRpt.cboLinesInclude.Text = "Include Only UnPaid Items on the Job"
            End If
            'Debug.Print(frmQuoteRpt.tgln(Row, "Paid").ToString)
            If frmQuoteRpt.ChkTotalsOnly.Checked = True Then MFGSubTotals = True '02-10-10
            If frmQuoteRpt.cboSortPrimarySeq.Text.ToString.StartsWith("MFG Sub-Totals in Catalog # Sequence") Then MFGSubTotals = True '09-10-09
            'Line Detail &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            '06-14-10 MaxRow = frmQuoteRpt.QuoteLinesBindingSource.Count - 1
            MaxRow = frmQuoteRpt.tgln.Splits(0).Rows.Count - 1 '06-18-10 
            If MaxRow > -1 Then  Else MsgBox("No Line Item Records Selected. Please Try Again") : GoTo Exit_Done '10-04-13  
            'CurrLev1 = frmQuoteRpt.tgln(Row, "MFG").ToString()
            For Row = 0 To MaxRow
                'If MFGSubTotals And CurrLev1 <> "JKH" ThenStop
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then CurrJob = frmQuoteRpt.tgln(Row, "JobName") '09-21-12
                CurrLev1 = frmQuoteRpt.tgln(Row, "MFG").ToString()
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Then '09-19-12  GoTo QutLineHistoryRpt
                    CurrLev1 = frmQuoteRpt.tgln(Row, "SLSQ").ToString()
                    'ElseIf frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.txtPrimarySortSeq.Text = "Quote Summary" Then '09-21-12  GoTo QutLineHistoryRpt
                    '    CurrLev1 = frmQuoteRpt.tgln(Row, "ProjectName").ToString()
                End If
                'if frmQuoteRpt.tgln.Splits(0).Rows.'If drQToRow.RowState = DataRowState.Deleted Then Continue For '12-10-09 GoTo 165 '12-09-09 frmQuoteRpt.tgr.MoveNext()
                If MFGSubTotals And Row = 0 Then PrevLev1 = CurrLev1 ' frmQuoteRpt.tgln(Row, "MFG").ToString() '09-08-09
                If MFGSubTotals And PrevLev1 <> CurrLev1 Then '09-19-12frmQuoteRpt.tgln(Row, "MFG").ToString() Then '09-08-09
                    'Print Totals 
                    If Trim$(PrevLev1) <> CurrLev1 Then '09-19-12 Trim(frmQuoteRpt.tgln(Row, "MFG").ToString()) Then
                        'Print sub totals for each Major Break
                        THDG = "**TOTAL " & PrevLev1 & "  Qty Cnt = " & Format(QuantityA(1), "######0") '02-10-10
                        Call TotPrt9250(THDG, TotalLevels.TotLv1, RT, doc)
                        ' 0=TotalPaid  1=PaidThisQuote 2=TotalUnPaid 3=UnpaidPaidthisquote'09-20-12 CommAmtA(1) = 0 : CommAmtA(3) = 0
                        If frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Then '03-19-14
                            ' doc.Body.Children.Add(RT)
                            'doc.Body.Children.Add(RT)
                            'ppv.C1PrintPreviewControl1.Document = doc
                            'ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
                            'ppv.Doc.Generate()
                            'ppv.Show()
                            'Exit Sub
                        End If
                        CommAmtA(0) += CommAmtA(1) : CommAmtA(2) += CommAmtA(3)
                        CommAmtA(1) = 0 : CommAmtA(3) = 0 'Zero out Low Level'= Paid

                        If frmQuoteRpt.chkSalesmanPerPage.CheckState = CheckState.Checked Then RT.Rows(RC).PageBreakBehavior = BreakEnum.Page '09-19-12
                        GoTo 333
                        '09-20-12******************************************************************************************
333:                    PrevLev1 = CurrLev1 '09-20-12 frmQuoteRpt.tgln(Row, "MFG").ToString() '09-09-09 CurrLev1

                        X = "ZeroLevels" '02-09-09
                        Call TotalsCalc(X, B, TotalLevels.TotLv1) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
                    End If
                End If
946:
                Dim APrice As String = frmQuoteRpt.tgln(Row, "Sell").ToString '02-15-10 
                Dim UnitOfM As String = frmQuoteRpt.tgln(Row, "UM").ToString
                Dim UnitMeas As Decimal = 1
                Dim UnitMeaStr As String = UnitMeaSet(APrice, UnitMeas, frmQuoteRpt.tgln.Columns("UM").Text) '' C = Hundreds M = Thousands FT =Feet '01-28-04

                LnQuantityA = CDec(Val(frmQuoteRpt.tgln(Row, "Qty"))) '09-08-09 
                'Aprice? FixSell?
                FixSell = Val(frmQuoteRpt.tgln(Row, "Sell").ToString)
                If frmQuoteRpt.pnlTypeOfRpt.Text <> "Terr Spec Credit Report" Then '03-19-14
                    If DIST Then '06-25-13
                        FixCost = Val(frmQuoteRpt.tgln(Row, "Cost").ToString)
                    Else
                        FixCost = Val(frmQuoteRpt.tgln(Row, "Book").ToString) '06-25-13 JTC Quote Lines Rep Book > FixCost
                    End If
                End If

                If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then CommAmt = 0 : GoTo TerrSpecSkip1 '09-21-12
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then '09-19-12  GoTo QutLineHistoryRpt
                    CommAmt = Val(frmQuoteRpt.tgln(Row, "Comm-$").ToString) '02-11-10 
                    '09-23-12 CommAmt = Val(frmQuoteRpt.tgln(Row, "Comm").ToString) '09-23-12 
                End If
                FixSellExt = Val(frmQuoteRpt.tgln(Row, "Ext Sell").ToString) '09-11-09 
                If DIST Then
                    FixCostExt = Val(frmQuoteRpt.tgln(Row, "Ext Cost").ToString) 'Ext Marg,Ext Cost,Ext Comm,LPSell,LPCost
                Else   'Rep Ext Comm to FixCostExt '02-11-10 
                    FixCostExt = Val(frmQuoteRpt.tgln(Row, "Ext Comm-$").ToString)
                End If
                LampSell = Val(frmQuoteRpt.tgln(Row, "LPSell").ToString)
                LampCost = Val(frmQuoteRpt.tgln(Row, "LPCost").ToString)

                FixProfit = FixSell - FixCost
                If FixSell <> 0 Then FixProfitPer = FixProfit / (FixSell + 0.00001) Else FixProfitPer = 0 '08-22-02 WNA
                LampProfit = LampSell - LampCost
                If LampSell <> 0 Then LampProfitPer = LampProfit / (LampSell + 0.00001) Else LampProfitPer = 0 '08-22-02 WNA
                '02-11-10If MFGSubTotals Then 
TerrSpecSkip1:
                Call TotalsCalc("AddAllLevels", B, C) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
                If DIST Then 'All Decimals FixSell!,FixProfit!, FixProfitPer!,LampSell!,LampCost!,Amt!,CommAmt!,Commpct!
                    'FixProfit = FixSell - FixCost
                    FixMargin = (FixSell - FixCost) / (FixSell + 0.0001) * 100 '07-08-09
                    'LampProfit = LampSell - LampCost
                    LpMargin = (LampSell - LampCost) / (LampSell + 0.0001) * 100 '07-08-09
                    'Commpct = (CommAmt / (Amt + 0.0001)) * 100 '
                    If FixMargin > 900 Then FixMargin = 999 Else If FixMargin < -900 Then FixMargin = -999
                    If LpMargin > 900 Then LpMargin = 999 Else If LpMargin < -900 Then LpMargin = -999
                Else 'Rep()
                    If DAYB Then
                        'CommAmt = drQRow.Sell '11-26-01 Reverse Cost & Sell
                        'Amt! = drQRow.Comm:
                    End If
                    If MFG Then '               'Comm
                        FixMargin = ((FixSell - FixCost) / (FixSell + 0.0001)) * 100
                    Else              'Comm
                        FixMargin = (FixCost / (FixSell + 0.0001)) * 100 '
                    End If
                    If FixMargin > 900 Then FixMargin = 999 Else If FixMargin < -900 Then FixMargin = -999 '06-24-04
                    FixProfitPer = FixMargin '02-05-09
                    LnQuantityA = CDec(Val(frmQuoteRpt.tgln(Row, "Qty"))) '09-08-09 
                End If 'End Margin )))))))))))))))))))))))))))))))))))))))))))))))))))))))))
                'Change values to show negatives
                If frmQuoteRpt.tgln(Row, "Qty").ToString.StartsWith("-") Then '02-12-10
                    'If frmQuoteRpt.tg(Row, "Sell").ToString.StartsWith("-") Then  Else frmQuoteRpt.tg(Row, "Sell") = "-" & frmQuoteRpt.tg(Row, "Sell").ToString
                    'Debug.Print(Len(frmQuoteRpt.tg(Row, "Type").ToString))
                    ' If frmQuoteRpt.tg(Row, "Cost").ToString.StartsWith("-") Then  Else frmQuoteRpt.tg(Row, "Cost") = "-" & frmQuoteRpt.tg(Row, "Cost").ToString
                    'If frmQuoteRpt.tg(Row, "Comm-$").ToString.StartsWith("-") Then  Else frmQuoteRpt.tg(Row, "Comm-$") = "-" & frmQuoteRpt.tg(Row, "Comm-$").ToString
                End If
                '06-17-10 RowCnt += 1
                '05-07-10 RC += 1 '02-10-10
                'Debug.Print(RC.ToString)
                frmQuoteRpt.tgln.Row = Row '07-27-09 Wrong Grid tgt.Row = Row

                If frmQuoteRpt.ChkTotalsOnly.Checked = True Then GoTo GetNextRow '05-07-10
                '06-14-10 RC += 1 '05-07-10 
                PC = 0 'Do Each Column
                '09-19-12 *************************************************************************************************************************
                'Dim RCS As Int32 = 0 '06-13-11 JTC Added = 0  Sub Table 
                'Dim RTS As RenderTable = New RenderTable '11-03-10 Sub Table If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.txtPrimarySortSeq.Text = "Quote Summary" Then '09-21-12  GoTo QutLineHistoryRpt
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Or frmQuoteRpt.txtPrimarySortSeq.Text = "Quote Summary" Then '09-19-12  GoTo QutLineHistoryRpt

                    If frmQuoteRpt.cboLinesInclude.Text = "Include All Lines on Job" Then
                    Else
                        'Debug.Print(frmQuoteRpt.tgln(Row, "Paid").ToString)
                        If frmQuoteRpt.cboLinesInclude.Text = "Include Only Paid Items on the Job" Then
                            If frmQuoteRpt.tgln(Row, "Paid") <> True Then GoTo GetNextQuoteLine '06-23-10  Continue For
                            GoTo KeepLine
                        End If
                        If frmQuoteRpt.cboLinesInclude.Text = "Include Only UnPaid Items on the Job" And frmQuoteRpt.tgln(Row, "Paid") <> False Then GoTo GetNextQuoteLine '06-19-10  Continue For '06-19
                        '09-23-12If frmQuoteRpt.tgln(Row, "Paid") = True And frmQuoteRpt.cboLinesInclude.Text <> "Include Only Paid Items on the Job" Then GoTo GetNextQuoteLine '06-19-10  Continue For '06-19-10
                        '09-23-12 If frmQuoteRpt.tgln(Row, "Paid") = False And frmQuoteRpt.cboLinesInclude.Text <> "Include Only UnPaid Items on the Job" Then GoTo GetNextQuoteLine '06-19-10  Continue For '06-19-10
                    End If
KeepLine:
                    'MoreLineTests:      If frmQuoteRpt.txtJobNameSS.Text <> "ALL" Then
                    '                        If Trim(MFG) = "" Then Continue For '12-01-09
                    '                        If InStr("," & Trim(frmQuoteRpt.txtJobNameSS.Text) & ",", "," & Trim(drQutLn.MFG) & ",") Then '
                    '                            'Sto ' MfgHit = 1
                    '                        Else
                    '                            Continue For '12-01-09
                    '                        End If
                    '                    End If
                    '09-21-12 Show Job Name''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Dim RCO As Short = 0
                    If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then
                        If CurrJob <> PrevJob Then
                            Dim EDate As Date = frmQuoteRpt.tgln(Row, "EntryDate").ToString()
                            If Row <> 0 Then
                                Dim ld As New C1.C1Preview.LineDef("1mm", Color.Black)
                                RT = New C1.C1Preview.RenderTable : RC = 0 : RCO = 0
                                RT.Cells(0, 0).Text = " "
                                RT = New C1.C1Preview.RenderTable
                                RT.Style.GridLines.Top = ld ' LineDef.Default
                                RT.Style.Padding.Top = ".20in"
                                RT.Style.Padding.Bottom = ".20in"
                                doc.Body.Children.Add(RT)
                            End If
                            RT = New C1.C1Preview.RenderTable : RC = 0 : RCO = 0
                            RT.Cells(RC, RCO).Text = "Job Name: " : RCO += 1
                            RT.Cells(RC, RCO).Text = "Project Code: " : RCO += 1
                            RT.Cells(RC, RCO).Text = "Entry Date: " : RCO += 1
                            RT.Cells(RC, RCO).Text = "SLSQ" : RCO += 1
                            RT.Cells(RC, RCO).Text = "Status: " : RCO += 1
                            RT.Cells(RC, RCO).Text = "Project Location: " : RCO += 1
                            RT.Rows(RC).Style.FontBold = True
                            RC += 1 : RCO = 0
                            RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "JobName").ToString() : RCO += 1
                            RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "QuoteCode").ToString() : RCO += 1
                            RT.Cells(RC, RCO).Text = EDate.ToShortDateString : RCO += 1
                            RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "SLSQ").ToString() : RCO += 1
                            RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "Status").ToString() : RCO += 1
                            RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "City").ToString() & ", " & frmQuoteRpt.tgln(Row, "State").ToString() : RCO += 1
                            RC += 1 : RCO = 1
                            RT.Cells(RC, RCO).Text = " "
                            doc.Body.Children.Add(RT)

                            If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then '03-19-14
                                Dim tmpds As dsSaw8
                                If frmQuoteRpt.chkIncludeSpecifiers.Checked = True And CurrJob <> PrevJob Then

                                    tmpds = New dsSaw8 : tmpds.EnforceConstraints = False
                                    tmpds = GetSpecifiers(frmQuoteRpt.tgln(Row, "QuoteID").ToString())
                                    If tmpds.projectcust.Rows.Count <> 0 Then
                                        RT = New C1.C1Preview.RenderTable : RCO = 0 : RC = 1
                                        RT.Cells(RC, RCO).Text = "Specifier Information" : RT.Cells(RC, RCO).Style.FontBold = True : RC += 1
                                        RT.Cells(RC, RCO).Text = "Specifier Type" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                        RT.Cells(RC, RCO).Text = "Firm Name" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                        RT.Cells(RC, RCO).Text = "SLS" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                        RT.Cells(RC, RCO).Text = "City, State" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                        RT.Cells(1, 0).SpanCols = RCO - 1 : RT.Cells(1, 0).Style.TextAlignHorz = AlignHorzEnum.Center
                                        RC += 1
                                        RCO = 0
                                        For Each dr As dsSaw8.projectcustRow In tmpds.projectcust.Rows
                                            Dim SPEC As String = ""
                                            Select Case dr.Typec
                                                Case "A" : SPEC = "Architect"
                                                Case "E" : SPEC = "Engineer"
                                                Case "S" : SPEC = "Specifier"
                                                Case "T" : SPEC = "Contractor"
                                                Case "X" : SPEC = "Other"
                                            End Select
                                            RT.Cells(RC, RCO).Text = SPEC : RT.Cells(RC, RCO).Style.FontBold = False : RCO += 1
                                            RT.Cells(RC, RCO).Text = dr.FirmName : RT.Cells(RC, RCO).Style.FontBold = False : RCO += 1
                                            RT.Cells(RC, RCO).Text = dr.SLSCode : RT.Cells(RC, RCO).Style.FontBold = False : RCO += 1
                                            Dim dsCityState As dsSaw8 = New dsSaw8 : dsCityState.EnforceConstraints = False
                                            dsCityState = GetSLSNameDetail(dr.NCode)
                                            If dsCityState.namedetail.Rows.Count <> 0 Then
                                                RT.Cells(RC, RCO).Text = dsCityState.namedetail.Rows(0)("City") & ", " & dsCityState.namedetail.Rows(0)("State") : RT.Cells(RC, RCO).Style.FontBold = False : RCO += 1
                                            Else
                                                RT.Cells(RC, RCO).Text = "  " : RCO += 1
                                            End If
                                            RT.Cells(RC, RCO).Text = "                               "
                                            RC += 1 : RCO = 0
                                        Next
                                        RT.Cells(RC, RCO).Text = " "
                                        doc.Body.Children.Add(RT)
                                    End If
                                End If

                                If frmQuoteRpt.chkIncludeSLSSPlit.Checked = True And CurrJob <> PrevJob Then
                                    tmpds = New dsSaw8 : tmpds.EnforceConstraints = False
                                    tmpds = GetSLSQuote(frmQuoteRpt.tgln(Row, "QuoteID").ToString())
                                    If tmpds.qutslssplit.Rows.Count <> 0 Then

                                        RT = New C1.C1Preview.RenderTable : RCO = 0 : RC = 0
                                        Dim SLSString As String = ""
                                        RT.Cells(0, 0).Text = "Project Salesmen:" : RT.Cells(0, 0).Style.FontBold = True : RT.Cols(0).Width = "1.5in"
                                        For Each dr As dsSaw8.qutslssplitRow In tmpds.qutslssplit.Rows
                                            SLSString += dr.SLSCODE & " " & dr.SLSSplit & " %" & ", "
                                        Next
                                        RT.Cells(0, 1).Text = SLSString
                                        RT.Cells(1, 0).Text = " "
                                        doc.Body.Children.Add(RT)
                                    End If
                                End If

                                If frmQuoteRpt.chkIncludeNotesLineItems.Checked = True And CurrJob <> PrevJob Then
                                    tmpds = New dsSaw8 : tmpds.EnforceConstraints = False
                                    tmpds = GetQuoteNotes(frmQuoteRpt.tgln(Row, "QuoteID").ToString())
                                    If tmpds.qutnotes.Rows.Count <> 0 Then
                                        RT = New C1.C1Preview.RenderTable : RCO = 0 : RC = 0
                                        RT.Cells(0, 0).Text = "Project Notes:" : RT.Cells(0, 0).Style.FontBold = True : RT.Cols(0).Width = "1.5in"
                                        If tmpds.qutnotes.Rows.Count <> 0 Then RT.Cells(0, 1).Text = tmpds.qutnotes.Rows(0)("notes").ToString.Trim
                                        RT.Cells(1, 0).Text = " "
                                        doc.Body.Children.Add(RT)
                                    End If
                                End If

                            End If

                            If CurrJob <> PrevJob Then
                                RT = New C1.C1Preview.RenderTable : RCO = 0 : RC = 0
                                RT.Cells(RC, RCO).Text = "Qty" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                RT.Cells(RC, RCO).Text = "Type" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                RT.Cells(RC, RCO).Text = "MFG" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                RT.Cells(RC, RCO).Text = "Description" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                RT.Cells(RC, RCO).Text = "Comm" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                RT.Cells(RC, RCO).Text = "Paid/Unpaid" : RT.Cells(RC, RCO).Style.FontBold = True : RCO += 1
                                RT.Cols(3).Width = "4.5in" : RT.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Left 'desc
                                doc.Body.Children.Add(RT)

                            End If

                            'If frmQuoteRpt.chkIncludeSpecifiers.Checked = True Or frmQuoteRpt.chkIncludeSLSSPlit.Checked = True Or frmQuoteRpt.chkIncludeNotesLineItems.Checked = True Then
                            'Else
                            PrevJob = CurrJob
                            'End If

                            'HeaderInfoPrinted = True
                        Else
                            'HeaderInfoPrinted = False
                        End If
                    End If
                    'If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Quote Summary") Then '09-21-12
                    'Else
                    '    RT.Cells(RC, 0).Text = frmQuoteRpt.tgln(Row, "LnCode") 'drQutLn.LnCode
                    'End If



                    HeaderInfoPrinted = True

                    RT = New C1.C1Preview.RenderTable : RCO = 0 : RC = 0
                    RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "Qty") : RT.Cells(RC, RCO).Style.FontBold = False : RCO += 1
                    RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "Type") : RT.Cells(RC, RCO).Style.FontBold = False : RCO += 1
                    RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "MFG") : RT.Cells(RC, RCO).Style.FontBold = False : RCO += 1
                    RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "Description") : RT.Cells(RC, RCO).Style.FontBold = False : RCO += 1
                    RT.Cells(RC, RCO).Text = frmQuoteRpt.tgln(Row, "Comm").ToString.Trim : RT.Cells(RC, RCO).Style.FontBold = False : RCO += 1

                    'If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then RT.Cols(0).Visible = False : RT.Cols(0).Width = Unit.Empty 'Don't print LnCode
                    ' '' 0=TotalPaid  1=PaidThisQuote 2=TotalUnPaid 3=UnpaidPaidthisquote'09-20-12 CommAmtA(1) = 0 : CommAmtA(3) = 0
                    'CommAmtA(0) += CommAmtA(1) : CommAmtA(2) += CommAmtA(3)
                    If frmQuoteRpt.tgln(Row, "Paid") = True Then
                        RT.Cells(RC, RCO).Text = "Paid"
                        CommAmtA(1) += Val(frmQuoteRpt.tgln(Row, "Comm")) * LnQuantityA '09-21-12 
                    Else
                        RT.Cells(RC, RCO).Text = "UnPaid" '09-10-12 
                        CommAmtA(3) += Val(frmQuoteRpt.tgln(Row, "Comm")) * LnQuantityA '09-21-12
                    End If
                    'Set up Quote Line Width

                    If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then
                    Else
                        RT.Cols(0).Width = ".75in" : RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Center 'LnCode
                    End If
                    'RT.Cols(0).Width = "1in" : RT.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Right  'Qty
                    'RT.Cols(1).Width = "1in" : RT.Cols(2).Style.TextAlignHorz = AlignHorzEnum.Center  'Type
                    'RT.Cols(2).Width = "1in" : RT.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Center  'Mfg
                    RT.Cols(3).Width = "4.5in" : RT.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Left 'desc
                    'RT.Cols(4).Width = ".75in" : RT.Cols(5).Style.TextAlignHorz = AlignHorzEnum.Right  'Cost
                    'RT.Cols(5).Width = ".75in" : RT.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Right 'Sell
                    'rT.Cols(6).Visible = False : RT.Cols(0).Width = Unit.Empty 'Don't print LnCode
                    'RT.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Center
                    'RC += 1 '09-10-12
                    doc.Body.Children.Add(RT)
                    'ppv.C1PrintPreviewControl1.Document = doc
                    'ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
                    'ppv.Doc.Generate()
                    'ppv.Show()
                    'Exit Sub

GetNextQuoteLine:   GoTo GetNextRow
                    '06-19-10           'Margin Calc Only Do Once Per Row *********************************************


                    '09-19-12 ************************************************************************************************************************
                End If
                For I = 0 To frmQuoteRpt.tgln.Splits(0).DisplayColumns.Count - 1  ' 02-03-09 frmFoll.tg.Splits(0).DisplayColumns.Count - 1
                    'This goes to QuoteLines '08-31-09 &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                    'Debug.Print(I.ToString & "Row=" & Row.ToString)
                    If frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If tglines.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                    If (frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                    ColText = frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).DataColumn.Text.ToString  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                    ColName = frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Name
                    If ColText = "True" Then ColText = "Y"
                    If ColText = "False" Then ColText = "N"
DoneMarginCalc:     If DIST Then
                        If ColName = "Margin-%" Then ColText = Format(FixMargin, "###.00") '
                        If ColName = "Ext Marg" Then ColText = Format(FixMargin, "###.00") '12-08-09Format(FixSellExt - FixCostExt, "########0") '09-11-09 
                    End If
                    '*************************************************
                    'Debug.Print(frmQuoteRpt.ChkTotalsOnly.Checked.ToString)
                    '06-17-10 If Not frmQuoteRpt.ChkTotalsOnly.Checked = True Then
                    RT.Cells(RC, PC).Text = ColText
                    If ColName = "EntryDate" Then RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-10-10
                    PC += 1
                    'End If
                Next 'Col
                'Debug.Print(RT.Cells(RC, 0).Text & RT.Cells(RC, 1).Text & RT.Cells(RC, 2).Text & RT.Cells(RC, 4).Text)
                If Not frmQuoteRpt.ChkTotalsOnly.Checked = True Then RC += 1 '06-17-10 Row Count
GetNextRow:
            Next 'Row

            'Debug.Print(RC.ToString)
            If MFGSubTotals Then '09-12-09
                THDG = "**TOTAL " & PrevLev1 & "  Qty Cnt = " & Format(QuantityA(1), "######0") ' Mfg Total '02-10-10
                Call TotPrt9250(THDG, TotalLevels.TotLv1, RT, doc)
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '09-23-12
                    ' 0=TotalPaid  1=PaidThisQuote 2=TotalUnPaid 3=UnpaidPaidthisquote'09-20-12 CommAmtA(1) = 0 : CommAmtA(3) = 0
                    CommAmtA(0) += CommAmtA(1) : CommAmtA(2) += CommAmtA(3)
                    '09-23-12 CommAmtA(1) = 0 : CommAmtA(3) = 0 'Zero out Low Level'= Paid
                    'set 1 & 3 to grand totals
                    CommAmtA(1) = CommAmtA(0) : CommAmtA(3) = CommAmtA(2)
                End If
            End If
            THDG = "**GRAND TOTAL " & "  Qty Cnt = " & Format(QuantityA(0), "######0") ' Mfg Total '02-10-10 ' Grabd Mfg Total
            Call TotPrt9250(THDG, TotalLevels.TotGt, RT, doc)
            Call TotalsCalc("ZeroLevels", B, TotalLevels.TotGt) '02-18-10
            LnQuantityA = 0 '02-18-10
            PrtCols -= 1

            If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Quote Summary") Then '09-19-12  GoTo QutLineHistoryRpt

                'Set up Quote Line Width
                RT.Cols(0).Width = ".75in" : RT.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Center 'LnCode
                RT.Cols(1).Width = ".75in" : RT.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Right  'Qty
                RT.Cols(2).Width = ".75in" : RT.Cols(2).Style.TextAlignHorz = AlignHorzEnum.Center  'Type
                RT.Cols(3).Width = ".75in" : RT.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Center  'Mfg
                RT.Cols(4).Width = "4.5in" : RT.Cols(4).Style.TextAlignHorz = AlignHorzEnum.Left 'desc
                RT.Cols(5).Width = ".75in" : RT.Cols(5).Style.TextAlignHorz = AlignHorzEnum.Right  'Cost
                RT.Cols(6).Width = ".75in" : RT.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Right 'Sell
                RT.Cols(0).Visible = False : RT.Cols(0).Width = Unit.Empty 'Don't print LnCode
                'For I = 0 To 5 : RT.Cells(RC, I).Text = "  " : Next 'vbNewLine '06-19-10 vbCrLf
                RC += 1
                'Header 
                ' RT.Rows.Insert(0, 4) '01-16-09 Insert Header
                ' RT.RowGroups(0, 4).PageHeader = True
                'RT.RowGroups(0, 1).Style.BackColor = C
                'RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
                '''''''''''''''''''''''''''''''''''''''''''
                'RT.Cells(0, 0).Text = "LnCode"
                'RT.Cells(0, 1).Text = "QTY" ''Qty"
                'RT.Cells(0, 2).Text = "TYPE" '  'Type
                'RT.Cells(0, 3).Text = "MFG"   'Mfg
                'RT.Cells(0, 4).Text = "DESCRIPTION" ' 'desc
                'RT.Cells(0, 5).Text = "COMM"   'COMM
                'RT.Cells(0, 6).Text = "PAID??"   'Cost
                'RT.Rows(0).Style.BackColor = LemonChiffon '06-19-10 
                'RT.Cols(0).Visible = False : RT.Cols(0).Width = Unit.Empty 'Don't print LnCode
                'End Line Items^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                'doc.Body.Children.Add(RT)
                GoTo 6090
            End If
            '06-14-10 RC = 0
            'Column Headers
            'INSERT COLUMN HEADERS 
            Dim Headertmp As String = ""
            PrtCols = 0   'Print Column Headers
            RT.Rows.Insert(0, 1) '06-14-10
            RT.RowGroups(0, 1).PageHeader = True '06-09-10
            frmQuoteRpt.tgln.MoveFirst()
            For I = 0 To frmQuoteRpt.tgln.Splits(0).DisplayColumns.Count - 1
                'Dim col2 As C1.Win.C1TrueDBGrid.C1DisplayColumn = frmQuoteRpt.tg.Splits(0).DisplayColumns(I) '02-20-09
                'Dim Tag As String = col2.DataColumn.Tag 'Debug.Print(tgShow(col.DataColumn.Tag, 0))'col.DataColumn.Tag = I.ToString
                TgWidth(I) = (frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Width / 100) '02-25-09
                If frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                If (frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
                RT.Cells(0, PrtCols).Text = frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Name
                If RT.Cells(0, PrtCols).Text = "Margin-$" Or RT.Cells(0, PrtCols).Text = "Margin-%" Then
                    RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
                End If
                RT.Cells(0, PrtCols).Style.BackColor = LemonChiffon '06-14-10 
                RT.Cols(PrtCols).Width = TgWidth(I) '02-22-09
                'Headertmp = Headertmp & frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Name & TgWidth(I).ToString
                PrtCols += 1
            Next
            PrtCols -= 1

6090:       RT.Width = "auto" : RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            'Start at Column 1 on Mfg
            '09-20-12 If PrimarySortSeq.ToUpper.StartsWith("MFG") Then RT.Cols(0).Visible = False Else RT.Cols(0).Visible = True '06-17-10 no Mfg Col
            '09-20-12 If PrimarySortSeq.ToUpper.StartsWith("MFG") Then RT.Cols(0).Width = Unit.Empty
            RT.Width = "auto" '
            'Split across pages
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage : RT.StretchColumns = StretchTableEnum.LastVectorOnPage
            'Grid Lines
            'If frmFoll.chkPrintGridLines.CheckState = CheckState.Checked Then 
            RT.Style.GridLines.All = LineDef.Default '06-12-10 Else RT.Style.GridLines.All = LineDef.Empty
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage
            Try
                doc.Body.Children.Add(RT) '06-12-10 
                'doc.Body.Children.Add(RThdr) '06-12-10 
            Catch ex As Exception
                ' If DebugOn ThenStop 'CatchStop
            End Try
ppvshowDoc:
            ppv.C1PrintPreviewControl1.Document = doc
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
            ppv.Doc.Generate()
            ppv.Show()
            '09-13-09 JTC *******************************************************************
            ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
            ppv.BringToFront()
            frmShowHideGrid.BringToFront() '03-10-09
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Quote Summary") Then '09-19-12  GoTo QutLineHistoryRpt
                frmQuoteRpt.chkSalesmanPerPage.CheckState = CheckState.Unchecked '09-23-12
            End If
ExitReportLoop:
            'objExcel.visible = True '05-15-15 JTC Put at End objExcel.visible = True


            If myConnection.State <> ConnectionState.Open Then Call OpenSQL(myConnection) '03-19-14
            mysqlcmd.Connection = myConnection '03-19-14
            strSql = "DROP TABLE IF EXISTS TMPREPORTS1 " '02-28-14
            mysqlcmd.CommandText = strSql : mysqlcmd.ExecuteNonQuery()

            GoTo Exit_Done '10-04-13 
        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "PrintReportQuotesLines" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop
        End Try
Exit_Done:  '#End

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow '10-04-13
    End Sub

    Public Function GetSLSNameDetail(ByVal Code As String) As dsSaw8 '03-19-14
        Dim FatalExceptionCnt As Boolean = False
Start:  Try
            Dim tmpds As dsSaw8 = New dsSaw8 : tmpds.EnforceConstraints = False
            Dim da As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            da.SelectCommand = New MySqlCommand("Select * from namslssplit WHERE Code = " & "'" & SafeSQL(Code) & "'", myConnection)
            da.Fill(tmpds, "namslssplit")
            da.SelectCommand = New MySqlCommand("Select * from namedetail WHERE Code = " & "'" & SafeSQL(Code) & "'", myConnection)
            da.Fill(tmpds, "namedetail")
            Return (tmpds)

        Catch ex As Exception
            If ex.Message.Contains("Fatal error encountered during command execution") = True Then
                'Leaving order system open for a long time and then coming back later on and trying to LU
                If FatalExceptionCnt = False Then Call OpenSQL(myConnection) : FatalExceptionCnt = True : GoTo Start Else GoTo DFLTMSG '05-31-12
            Else
DFLTMSG:        MessageBox.Show("Error in GetSLS" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return Nothing
        End Try

    End Function

    Public Function GetSLSQuote(ByVal QuoteID As String) As dsSaw8 '03-19-14
        Dim FatalExceptionCnt As Boolean = False
Start:  Try
            Dim tmpds As dsSaw8 = New dsSaw8 : tmpds.EnforceConstraints = False
            Dim mdaND As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            mdaND.SelectCommand = New MySqlCommand("Select * from qutslssplit Where QuoteID = '" & QuoteID & "'", myConnection)
            mdaND.Fill(tmpds, "qutslssplit")
            Return (tmpds)
        Catch ex As Exception
            If ex.Message.Contains("Fatal error encountered during command execution") = True Then
                If FatalExceptionCnt = False Then Call OpenSQL(myConnection) : FatalExceptionCnt = True : GoTo Start Else GoTo DFLTMSG
            Else
DFLTMSG:        MessageBox.Show("Error in GetSLS" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return Nothing
        End Try

    End Function

    Public Function GetQuoteNotes(ByVal QuoteID As String) As dsSaw8 '03-19-14
        Dim FatalExceptionCnt As Boolean = False
Start:  Try
            Dim tmpds As dsSaw8 = New dsSaw8 : tmpds.EnforceConstraints = False
            Dim mdaND As MySql.Data.MySqlClient.MySqlDataAdapter = New MySqlDataAdapter
            mdaND.SelectCommand = New MySqlCommand("Select * from qutnotes where QuoteID = '" & QuoteID & "'", myConnection)
            mdaND.Fill(tmpds, "qutnotes")
            Return (tmpds)
        Catch ex As Exception
            If ex.Message.Contains("Fatal error encountered during command execution") = True Then
                If FatalExceptionCnt = False Then Call OpenSQL(myConnection) : FatalExceptionCnt = True : GoTo Start Else GoTo DFLTMSG
            Else
DFLTMSG:        MessageBox.Show("Error in GetQuoteNotes" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return Nothing
        End Try

    End Function

    Public Function GetSpecifiers(ByVal QuoteID As String) As dsSaw8 '03-19-14
        Dim FatalExceptionCnt As Boolean = False
Start:  Try
            Dim tmpds As dsSaw8 = New dsSaw8 : tmpds.EnforceConstraints = False
            Dim da As MySqlDataAdapter = New MySqlDataAdapter
            da = New MySqlDataAdapter
            strSql = "Select * from projectcust where QuoteID = '" & QuoteID & "' and (TypeC = 'X' or Typec = 'A' or Typec = 'E' or Typec = 'S' or Typec = 'T' or Typec = 'L') order by NCode" '09-20-10 X = OTHER
            da.SelectCommand = New MySqlCommand(strSql, myConnection) '09-14-09 JH 
            da.Fill(tmpds, "projectcust")
            Return tmpds
        Catch ex As Exception
            If ex.Message.Contains("Fatal error encountered during command execution") = True Then
                If FatalExceptionCnt = False Then Call OpenSQL(myConnection) : FatalExceptionCnt = True : GoTo Start Else GoTo DFLTMSG
            Else
DFLTMSG:        MessageBox.Show("Error in GetSLS" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Return Nothing
        End Try

    End Function

    Public Sub PrintRealizationReportQutTO()
        '09-23-15 JTC If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ = Fixed Length font
        Dim FontSaveReal As String = frmQuoteRpt.RibbonFontComboBox2.Text
        Try '#Top
6000:       '07-06-12 If Specifiers Checked then ask Question
            'RealExtByInfluencePercent Dim ExtendByInfluencePercent As Boolean = False '07-06-12 in LpSell
            '07-06-12 If Specifiers Checked then ask Question
            'Do you want to multiply Sales Amounts by the Probability? Yes/No
            If (RealArchitect = True) Or (RealEngineer = True) Or (RealLtgDesigner = True) Or (RealSpecifier = True) Or (RealContractor = True) Or (RealOther = True) Then '12-07-16
                Resp = MsgBox("Yes = To multiply Specifier Sales Amounts by Influence Percent or " & vbCrLf & "No = Use actual amounts" & vbCrLf & "If Influence Percent is zero we will use 100%", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton2, "Realization Sort") '07-06-12
                If Resp = vbYes Then RealExtByInfluencePercent = True
            End If
           
            Dim LastCustQT As String = "" '03-29-12 
            Dim TGNameStr As String = "" 'Documentation Set Up a String of Names
            Dim TGWidthStr As String = "" 'Set Up a String of Widths
            ppv.Doc.Clear() 'Clear the Doc
            Call SetupPrintPreview(FirmName) '09-18-08
            '09-23-15 JTC If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)) Then
                frmQuoteRpt.RibbonFontComboBox2.Text = "Consolas"
            End If
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1
            '05-03-10 JH ppv.C1PrintPreviewControl1.PreviewPane.HideMarginsState = C1.Win.C1Preview.HideMarginsFlags.All
            '05-03-10 JH ppv.C1PrintPreviewControl1.PreviewPane.HideMargins = C1.Win.C1Preview.HideMarginsFlags.None
            '05-03-10 JH ppv.C1PrintPreviewControl1.PreviewPane.PagesPaddingSmall = New Size(0, 0)
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomMode = C1.Win.C1Preview.ZoomModeEnum.PageWidth

            'Header 
            Dim RArea As C1.C1Preview.RenderArea = New C1.C1Preview.RenderArea
            'Type of Report - & Agency Name
            RT = New C1.C1Preview.RenderTable
            RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            RT.CellStyle.Padding.Left = "1mm" '12-13-12
            RT.CellStyle.Padding.Right = "1mm" '12-13-12
            RT.Style.GridLines.All = LineDef.Empty '  LineDef.Default    '12-04-10 & "  UserID = " & UserID 
            RT.Cells(0, 0).Text = "Report: " & frmQuoteRpt.pnlTypeOfRpt.Text.Trim & "  UserID = " & UserID & "    Report Date = " & Format(Now, "MM/dd/yyyy") '10-17-10
            RT.Cells(0, 1).Text = AGnam : RT.Cells(0, 1).Style.TextAlignHorz = AlignHorzEnum.Right
            Dim fs As Integer = frmQuoteRpt.FontSizeComboBox.Text
            ''09-23-15 JTC If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
            'Dim FontSaveReal As String = frmQuoteRpt.RibbonFontComboBox2.Text
            'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)) Then
            '    frmQuoteRpt.RibbonFontComboBox2.Text = "Consolas"
            'End If
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs + 2, FontStyle.Bold)
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            '06-14-10 RArea.Children.Add(RT)
            'Select Criteria
            '06-14-10 RT = New C1.C1Preview.RenderTable
            '06-14-10 RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            '06-14-10 RT.Style.GridLines.All = LineDef.Empty '  LineDef.Default
            Dim sort As String = "Primary Sort: " & frmQuoteRpt.txtPrimarySortSeq.Text
            If frmQuoteRpt.txtSecondarySort.Text.Trim <> "" Then sort = sort + " Secondary Sort: " & frmQuoteRpt.txtSecondarySort.Text
            'If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" ThenStop '01-21-14 
            If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then sort = sort + " All Specifiers for Code = " & RealWithOneMfgCustCode '01-21-14 
            RT.Rows(1).Style.TextAlignHorz = AlignHorzEnum.Left
            RT.Rows(1).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs)
            ''09-18-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
            'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True) Then
            '    'fix RT.Rows(1).Style.Font = New Font("Consolas", FontStyle.Bold)
            'End If 'Consolas or Courier

            RT.Cells(1, 0).Text = sort
            RT.Cells(1, 1).Text = "Page [PageNo] of [PageCount]"
            RT.Cells(1, 1).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Right

            RT.Cells(2, 0).Text = "Select Criteria: " & SelectionText 'frmQuoteRpt.TtxtSortSelV.Text '07-02-09frmProjRpt.txtPrimarySortSeq.Text & " " & frmProjRpt.txtSecondarySort.Text '07-01-09
            RT.Cells(2, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
            RT.Cells(2, 0).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs - 1)  '04-30-10 jh - FONT COMBO
            ''09-18-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
            'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True) Then
            '    'fix  RT.Cells(2, 0).Style.Font = New Font("Consolas", FontStyle.Bold)
            'End If 'Consolas or Courier
            RT.Cells(2, 0).SpanCols = RT.Cols.Count

            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            '03-19-14 RT.Style.BackColor = LemonChiffon
            '02-04-12 RArea.Children.Add(RT)
            '02-04-12 doc.Body.Document.PageLayout.PageHeader = RArea
            doc.Body.Children.Add(RT) ''02-04-12 Header on First Page Col hdg on All Pages
            'END PAGE HEADER - DIFFERENT THAN THE TABLE HEADERS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'REPORT BODY '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            RT = New C1.C1Preview.RenderTable
            RC = 0 '06-14-10 
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            ''09-18-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
            'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True) Then
            '    'fix RT.Style.Font = New Font("Consolas", FontStyle.Bold)
            'End If 'Consolas or Courier
            RT.Style.GridLines.All = LineDef.Default
            RT.CellStyle.Padding.Left = "1mm" '12-13-12
            RT.CellStyle.Padding.Right = "1mm" '12-13-12
6315:
            frmShowHideGrid.tgShow.SetDataBinding(table, "")
            frmQuoteRpt.tgr.Refresh()

            Dim C As Integer
            Dim X As String = "ZeroLevels" ' "AddAllLevels" ''01-26-09
            Call TotalsCalc(X, B, C) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            'Update the data to the dataset:
            CurrLev1 = "" : PrevLev1 = "" : CurrLev2 = "" : PrevLev2 = "" : Cmd = "" 'Cmd = "EOF"
            Dim A As String = "PrintLine"
            frmQuoteRpt.tgr.UpdateData()
            frmQuoteRpt.tgr.MoveFirst() '02-14-09 
            Dim PrimarySortSeq As String = frmQuoteRpt.txtPrimarySortSeq.Text
            Dim SeconarySortSeq As String = frmQuoteRpt.cboSortSecondarySeq.Text
            'If RealCustomer = True And RealManufacturer = False And frmQuoteRpt.cboTypeCustomer.Text.Trim.ToUpper <> "ALL" And SeconarySortSeq <> "Cusstomer Code" Then '03-11-14
            '    PrimarySortSeq = frmQuoteRpt.cboSortSecondarySeq.Text
            '    SeconarySortSeq = ""
            'End If
            Dim RowCnt As Integer = 0 'Major Print Loop
            Dim FirstReal As Int16 = 0 '-7-09-09 daQuoteRealLU.Fill(dsQuoteRealLU, "QuoteRealLU")
            drQToRow = dsQuoteRealLU.QuoteRealLU.Rows(0) '02-11-09 
            For Each drQToRow In dsQuoteRealLU.QuoteRealLU.Rows 'dsQutLU
                If drQToRow.RowState = DataRowState.Deleted Then Continue For '12-10-09 GoTo 165 '12-09-09 frmQuoteRpt.tgr.MoveNext()
                If FirstReal = 0 Then FirstReal = +1 : frmQuoteRpt.tgr.MoveFirst() '02-10-09
                'If DebugOn ThenDebug.Print(frmQuoteRpt.tgr.Splits(0).DisplayColumns("SLSQ").DataColumn.Text & frmQuoteRpt.tgr.Splits(0).DisplayColumns("Ncode").DataColumn.Text & drQToRow.SLSQ & drQToRow.NCode)
                'Debug.Print(drQToRow.SLSQ & drQToRow.NCode)
                If drQToRow.RowState = DataRowState.Deleted Then GoTo 165 '12-09-09 frmQuoteRpt.tgr.MoveNext() : Continue For ' 12-09-09 GoTo 235 '06-19-08
                'Debug.Print(drQToRow.NCode & drQToRow.QuoteCode) '  & drQToRow.QuoteToDate.ToString)
                If frmQuoteRpt.chkShowLatestCust.CheckState = CheckState.Checked Then '03-29-12
                    'Dim LastCustQT As String = "" '03-29-12 
                    If drQToRow.NCode.Trim = "" And drQToRow.FirmName.Trim = "" Then GoTo 165 '03-29-12 Don't Print 
                    If LastCustQT = drQToRow.NCode.Trim & drQToRow.QuoteCode.Trim Then GoTo 165 '03-29-12 JTC Only Print First Cust If More Than One
                    If LastCustQT = "" Then LastCustQT = drQToRow.NCode.Trim & drQToRow.QuoteCode.Trim '03-29-12
                    If LastCustQT <> drQToRow.NCode.Trim & drQToRow.QuoteCode.Trim Then LastCustQT = drQToRow.NCode.Trim & drQToRow.QuoteCode.Trim '03-29-12
                End If
                Dim Hit As Short = 0
                Call SelectHit9500(Hit, multsrtrvs) '01-25-09
                If Hit = 0 Then GoTo 165 '02-14-09 Continue For
                RowCnt += 1 '02-08-09
                '07-06-12 use A,E,L,S,T,X**********************************************************************************
                If RealExtByInfluencePercent = True Then '07-06-12 
                    If (RealArchitect And drQToRow.Typec = "A") Or (RealEngineer And drQToRow.Typec = "E") Or (RealLtgDesigner And drQToRow.Typec = "L") Or (RealSpecifier And drQToRow.Typec = "S") Or (RealContractor And drQToRow.Typec = "T") Or (RealOther And drQToRow.Typec = "X") Then '07-06-12 
                        'If RealEngineer And drQToRow.Typec "E" '02-04-12
                        'If RealLtgDesigner And drQToRow.Typec = "L" '02-04-12
                        'If RealSpecifier And drQToRow.Typec = "S" '02-04-12
                        'If RealContractor And drQToRow.Typec = "T" '02-04-12
                        'If RealSLSCustomer Then PcType = "C" '01-20-12 "Salesman/Customer"
                        'If  '01-31-12
                        'If PcType <> "*" Then strSql += " and (projectcust.TypeC = '" & PcType & "' " '08-21-11 ASdded (
                        If drQToRow.LPSell <> 0 Then ' drQToRow.LPSell = 1 ' If Influence % = 0 make 1
                            drQToRow.Sell = drQToRow.Sell * drQToRow.LPSell
                            drQToRow.Cost = drQToRow.Cost * drQToRow.LPSell
                            drQToRow.Comm = drQToRow.Comm * drQToRow.LPSell
                        End If
                    End If
                End If

                '******************************************************************************************

                Call SubTotChk9360(RT, doc) '02-08-09
                Call PrintQuoteRealLineRpt946(A, RT) '02-25-09 '946 Format routine and Calc Routine
                LnQuantityA = 1 '11-04-10   'C
                If frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Unchecked Then '02-15-09 chkDetailTotal  Unchecked = Detail
                    'If Not TotalsOnly Then
                    RC += 1 '06-14-10
                End If
                'toget other Tables for this quote
                Dim QutID As String, ProjID As String
                QutID = drQToRow.QuoteID 'Me.tgQutLU.Columns("QuoteID").CellText(Me.tgQutLU.SelectedRows.Item(0))
                ProjID = drQToRow.ProjectID ' Me.tgQutLU.Columns("ProjectID").CellText(Me.tgQutLU.SelectedRows.Item(0))
                Call TotalsCalc("AddAllLevels", B, C) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
                Hit = 0
                'Print Other stuff
                'PRINT NOTES '02-11-12 JTC Add Notes to Realization
                If frmQuoteRpt.chkNotes.CheckState = CheckState.Checked Then '02-11-12 JTC Add Notes to Realization
                    Dim dTPC As dsSaw8.qutnotesDataTable = GetNotes(drQToRow.QuoteID) '05-15-10 JH 
                    For Each drQNRow As dsSaw8.qutnotesRow In dTPC.Rows
                        If drQNRow.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                        RT.Cells(RC, 0).Text = "     " & drQNRow.Notes ''02-11-12 JTC Add Notes to Realization
                        RT.Cells(RC, 0).Style.BackColor = LightGray
                        RT.Cells(RC, 0).SpanCols = RT.Cols.Count - 1 '11-04-10  9 '07-15-10 
                        RC += 1
                        'Debug.Print(RC.ToString)
                    Next
                End If

165:            frmQuoteRpt.tgr.MoveNext() '02-14-09
            Next
            Cmd = "EOF" 'PUBLIC
            '
            '05-04-10 If drQToRow.RowState = DataRowState.Deleted ThenStop 'GoTo Hell
            Call SubTotChk9360(RT, doc) '02-08-09
            THDG = "**GRAND TOTAL  " & "Record Count = " & QuantityA(0).ToString '02-07-10 QuantityA(I)
            ''09-23-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)) Then
                THDG = Left("**GRAND TOTAL RECORDS = " & Space(25), 41) & "Count = " & QuantityA(0).ToString '09-18-15 
            End If
            TRCT = Left(("Record Count = " & Str(RecordCt)) & Wspcs, 20)

            'FixSell, FixCost, FixProfit, LampSell, LampCost, ProfitLamp, CommAmt, Commpct
            '01-25-09 TotalLevels.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            Call TotPrt9250(THDG, TotalLevels.TotGt, RT, doc)
            If THDG.StartsWith("**GRAND") Then '11-20-10
                RT.Rows(RC).PageBreakBehavior = BreakEnum.None '11-19-10 
                RT.BreakAfter = BreakEnum.None
            End If
            Call TotalsCalc("ZeroLevels", B, TotalLevels.TotGt) '02-18-10
            LnQuantityA = 0 '02-18-10
            Cmd = "" 'Off "EOF" 'PUBLIC
6010:
            'Column Headers  'INSERT COLUMN HEADERS 
            RT.Rows.Insert(0, 1) '06-14-10
            RT.RowGroups(0, 1).PageHeader = True '06-09-10
            MaxCol = frmQuoteRpt.tgr.Splits(0).DisplayColumns.Count - 1
            Dim PrtCols As Int16 = 0 'Print Column Headers
            For I = 0 To MaxCol '02-03-09frmQuoteRpt.tgr.Splits(0).DisplayColumns.Count - 1 
                'Dim col2 As C1.Win.C1TrueDBGrid.C1DisplayColumn = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I) '02-20-09
                'Dim Tag As String = col2.DataColumn.Tag
                TgWidth(I) = (frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Width / 100) '02-25-09
                If frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                'If DebugOn ThenDebug.Print(I.ToString & frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Name & TgWidth(I).ToString)
                If (frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For
                RT.Cells(0, PrtCols).Text = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Name.ToString '07-09-09 frmShowHideGrid.tgShow(Tag, 1) '02-25 User HeadingfrmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Name
                If RT.Cells(0, PrtCols).Text = "Margin-$" Or RT.Cells(0, PrtCols).Text = "Margin-%" Then
                    RT.Cols(PrtCols).Style.TextAlignHorz = AlignHorzEnum.Right
                End If
                RT.Cols(PrtCols).Width = TgWidth(I) '02-22-09
                'Debug.Print(TgWidth(I).ToString)
                RT.Cells(0, PrtCols).Style.BackColor = LemonChiffon '06-14-10
                TgName(PrtCols) = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Name
                Dim FF As String = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Name & frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).DataColumn.Caption
                If DIST And TgName(PrtCols) = "Comm" Then TgName(PrtCols) = "Margin" 'Can't do  frmQuoteRpt.tgr.Splits(0).DisplayColumns(col).Name = "Margin" '02-15-09 
                If DIST And TgName(PrtCols) = "LPComm" Then TgName(PrtCols) = "LPMarg" '02-15-09TgName(PrtCols) = frmQuoteRpt.tgr.Splits(0).DisplayColumns(col).Name
                If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then
                    If TgName(PrtCols) = "Comm" Or TgName(PrtCols) = "Margin" Or TgName(PrtCols) = "LPComm" Or TgName(PrtCols) = "LPMarg" Or TgName(PrtCols) = "Cost" Or TgName(PrtCols) = "Comm-$" Then '01-24-13
                        RT.Cells(0, PrtCols).Text = "" '01-24-13 'Skip
                    End If
                End If
                TgCol(PrtCols) = I.ToString 'col
                TGNameStr = TGNameStr & TgName(PrtCols) & " " & TgWidth(I).ToString & ":"
                'Done Above RT.Cols(PrtCols).Width = TgWidth(I) '12-01-08
                ''09-23-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ

                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)) Then
                    ''Debug.Print(TgName(PrtCols))
                    If RealWithOneMfgCustCode <> "" And RealManufacturer = True Then GoTo SkipWidthChange '09-23-15 JTC RealWithOneMfgCustCodeQuotesToOneMFG
                    If TgName(PrtCols) = "EntryDate" Or TgName(PrtCols) = "SLSCode" Or TgName(PrtCols) = "Status" Or TgName(PrtCols) = "SLSQ" Then
                        RT.Cols(PrtCols).Width = 0 ' .5
                        'frmQuoteRpt.tgr.Splits(0).DisplayColumns(PrtCols).Visible = False
                        ' frmQuoteRpt.tgr.Splits(0).DisplayColumns(PrtCols).Width = 0
                    End If
                    If TgName(PrtCols) = "JobName" Then
                        RT.Cols(PrtCols).Width = 5
                        'frmQuoteRpt.tgr.Splits(0).DisplayColumns(PrtCols).Visible = False
                        'frmQuoteRpt.tgr.Splits(0).DisplayColumns(PrtCols).Width = 2
                    End If
                    ''09-22-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
                    If TgName(PrtCols) = "NCode" Then
                        '09-23-15 JTCRT.Cols(PrtCols).Style.Font = New Font("Consolas", FontStyle.Bold)
                        'RT.Cells(RC, PC).Style.Font = New Font("Consolas", FontStyle.Bold)
                    End If 'Consolas or Courier
                End If ' End '09-18-15 
SkipWidthChange:  '09-23-15 JTC
                PrtCols += 1
                'CNVT .Width = VB6.TwipsToPixels(1440)' Twips/inch=1440  Pixels/inch=96   twips/15 = Pixels  on .Width = and .Height =
145:        Next
            'Debug.Print(RT.Cols(0).Width.ToString)
            PrtCols -= 1
            '@#R ProjectCustID0,ProjectID1,NCode2,Got3,Typec4,QuoteCode5,ProjectName6,FirmName7,ContactName8,EntryDate9,SLSCode10,Status11,Cost12,Sell13,Margin14,LPCost15,LPSell16,LPMarg17,Overage18,ChgDate19,OrdDate20,NotGot21,Comments22,SPANumber23,SpecCross24,LotUnit25,LampsIncl26,Terms27,FOB28,QuoteID29,BranchCode30,MarketSegment31,MFGQuoteNumber32,BidDate33,SLSQ34,RetrCode35,SelectCode36,LeadTime37,"
            'Debug.Print(TGNameStr)
            '06-14-10 RT.Width = "auto"RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            '06-14-10RT.Style.BackColor = LemonChiffon
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
            '06-14-10 RT.Rows(RC).Style.BackColor = LemonChiffon
            RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular)
            ''09-18-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True) Then
                RT.Style.Font = New Font("Consolas", FontStyle.Bold) '09-23-15 JTC
            End If 'Consolas or Courier
            '05-27-10 Doesn't work RArea.Children.Add(RT)
            '05-27-10 doc.Body.Document.PageLayout.PageHeader = RArea
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            '06-14-10 doc.Body.Children.Add(RT) '
            '06-14-10 RT = New C1.C1Preview.RenderTable
            '06-14-10  RT.Style.GridLines.All = LineDef.Default
            'ProjectCustID 1:ProjectID 1:QuoteCode 1:NCode 1:FirmName 1:ContactName 1:SLSCode 1:Got 1:Typec 1:MFGQuoteNumber 1:Cost 1:Sell 1:Comm 1:Overage 1:ChgDate 1:OrdDate 1:NotGot 1:Comments 1:SPANumber 1:SpecCross 1:LotUnit 1:LPCost 1:LPSell 1:LPComm 1:LampsIncl 1:Terms 1:FOB 1:QuoteID 1:BranchCode 1:ProjectName 1:MarketSegment 1:EntryDate 1:BidDate 1:SLSQ 1:Status 1:RetrCode 1:SelectCode 1:LastChgBy 1:City 1:State 1:CSR 1:StockJob 1:LotUnit1 1:

            'no MaxCol = 43 'PrtCols
            '06-14-10 RC = 0

            'Footer
            '06-14-10 RC = 0
            '10-17-10 Deleted Footer '''''''''''''''''''''''
            'Dim RF As New C1.C1Preview.RenderTable
            'RF.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10
            'RF.Style.Padding.All = "0mm" : RF.Style.Padding.Top = "0mm" : RF.Style.Padding.Bottom = "0mm"
            'RF.Style.GridLines.All = LineDef.Empty
            'RF.Cells(0, 0).Text = Now.ToShortDateString
            'RF.Cells(0, 1).Text = "Page [PageNo] of [PageCount]"
            'RF.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs - 2, FontStyle.Bold)
            'doc.Body.Document.PageLayout.PageFooter = RF
            '''''''''''''''''''''''
            'RT = New C1.C1Preview.RenderTable '02-04-09
            'RT.Rows.Insert(RC, 1)
            '' Orphan control: this line makes sure that at least 3 lines are printed before the footer on the same page.
            'RT.RowGroups(RC, 1).MinVectorsBefore = 3
            'RT.RowGroups(RC, 1).Footer = TableFooterEnum.All
            'RT.RowGroups(RC, 1).Style.TextColor = RT.RowGroups(RC, 1).Style.TextColor
            'RT.RowGroups(RC, 1).Style.BackColor = RT.RowGroups(RC, 1).Style.BackColor
            'RT.Cells(RC, 0).Style.BackColor = Color
            'RT.Cells(RC, 0).Style.FontSize = 14
            'RT.Cells(RC, 0).Text = "Record count = " & RowCnt.ToString & "    Page [PageNo] of [PageCount]     *" '01-22-10
            'RT.Cells(RC, 0).SpanCols = RT.Cols.Count '07-02-09 - 1 ' 22 '12-30-08RT.Cols.Count
            'RT.Width = "auto"
            'RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
            'RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Regular) '05-03-10 JH
            'RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            'doc.Body.Children.Add(RT) '12-29-06
            ''Header 
            'RT.Rows.Insert(0, 0) '01-16-09 Insert Header
            ''RT.RowGroups(0, 2).Style.BackColor = Color.
            'RT.RowGroups(0, 1).PageHeader = True
            'RT.RowGroups(0, 1).Style.BackColor = Color.
            'RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
            ''No RT.RowGroups(0, 1).Header = TableHeaderEnum.All
            'RT.Cells(0, 0).Text = " Quote Report    Report Date = " & VB6.Format(Now, "Short Date") & Space(4) & FirmName & Space(8) & "Page [PageNo] of [PageCount]     *" '12-30-08
            'RT.Cells(0, 0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
            'RT.Cells(0, 0).Style.BackColor = Color.
            'RT.Cells(0, 0).Style.FontSize = 14
            ''RT.Cells(0, 1).Text = "Page [PageNo] of [PageCount]     *" '  & "     " '06-18-08 Spacing
            ''RT.Cells(0, 1).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
            ''RT.Cells(0, 1).Style.BackColor = Color.B
            ''RT.Cells(0, 1).Style.FontSize = 14
            'RT.Cells(0, 0).SpanCols = RT.Cols.Count '/ 2 '12-30-08
            'RT.Cells(0, 0).SpanRows = 1 '01-16-09
            RT.Width = "auto" : RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
            'Start at Column 1
            '09-14-10 No No RT.Cols(0).Visible = False : RT.Cols(0).Width = Unit.Empty : RT.Width = "auto" '
            'Split across pages
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage : RT.StretchColumns = StretchTableEnum.LastVectorOnPage
            'Grid Lines
            'If frmFoll.chkPrintGridLines.CheckState = CheckState.Checked Then 
            RT.Style.GridLines.All = LineDef.Default '06-12-10 Else RT.Style.GridLines.All = LineDef.Empty
            RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage
            RT.StretchColumns = StretchTableEnum.LastVectorOnPage
            doc.Body.Children.Add(RT) '06-12-10 
6090:       ppv.C1PrintPreviewControl1.Document = doc
            'ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor.Equals(100)
            ppv.C1PrintPreviewControl1.PreviewPane.ZoomFactor = 1 '12-12-08
            ppv.Doc.Generate() '11-18-08
            ppv.Show()
            ppv.MaximumSize = New System.Drawing.Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
            ppv.BringToFront()
            frmShowHideGrid.BringToFront() '03-10-09
            If BeginCode = "CSV Tab" Or BeginCode = "CSV Comma" Then '01-06-09
                Dim CsvCode As String = Microsoft.VisualBasic.Constants.vbTab
                Dim CsvFileName As String = UserDocDir & "QuoteReportTmp.txt" 'BeginCode = "CSV Tab" '01-06-09 Public
                If BeginCode = "CSV Tab" Then CsvFileName = UserDocDir & "QuoteReportTmp.txt" 'tab  '01-06-09 Public
                If BeginCode = "CSV Comma" Then CsvFileName = UserDocDir & "QuoteReportTmp.csv" ' CsvCode = "," '01-06-09 Public
                fileExists = CheckForFile(CsvFileName, False)
                If fileExists = True Then
                    Kill(CsvFileName)
                End If '01-06-09
                FileClose(3) : FileOpen(3, CsvFileName, OpenMode.Output)
                Dim Rr As Integer = RT.Rows.Count - 1
                Dim CC As Integer = RT.Cols.Count - 1
                Dim LineData As String
                For Rr = 1 To RT.Rows.Count - 1
                    LineData = ""
                    For CC = 0 To RT.Rows.Count - 1
                        If BeginCode = "CSV Tab" And RT.Cols(CC).Visible = True Then LineData = LineData & RT.Cells(Rr, CC).Text & Microsoft.VisualBasic.Constants.vbTab '01-06-09
                        If BeginCode = "CSV Comma" And RT.Cols(CC).Visible = True Then
                            If InStr(RT.Cells(Rr, CC).Text, ",") Then RT.Cells(Rr, CC).Text = Replace(RT.Cells(Rr, CC).Text, ",", " ") '09-13-09 Space in Text with comma
                            LineData = LineData & RT.Cells(Rr, CC).Text & "," '01-06-09
                        End If
                    Next CC
                    PrintLine(3, LineData) 'eol
                Next Rr 'Row
                FileClose(3)
            End If ' Not CSV

ExitReportLoop:
        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "PrintRealizationReportQutTo" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop
        End Try
Exit_Done:  '#End
        RealWithOneMfgCustSortJobName = False '10-13-14 JTCPublic RealWithOneMfgCustSortJobName As Boolean = False '10-13-14 JTC
        '09-23-15 JTC set Font BackDim FontSaveReal As String = frmQuoteRpt.RibbonFontComboBox2.Text
        If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)) Then
            frmQuoteRpt.RibbonFontComboBox2.Text = FontSaveReal '09-23-15 "Consolas"
            System.Windows.Forms.Application.DoEvents()
        End If

    End Sub


    Public Sub ExcelQuoteFollowUp()
        Try '#Top '04-24-15 JTC Create ExcelQuoteFollowUp Public Sub from PrintSESCOJobListRealReportQutTO(
            Dim B As String = ""
            Dim Saw8MaxCols As Int16
            Static EndingLine As Short
            Static StartingLine As Short
            Dim iCol As Int32 '        'For iRow = 0 To 1 'ProdDataArray
            Dim objExcel As Object = Nothing
            Dim objBooks As Object = Nothing
            Dim objSheets As Object = Nothing
            Dim objSheet As Object = Nothing
            Dim objBook As Object = Nothing
            Dim A As String = ""
            Dim FileName As String = ""
            Dim StBidDate As String = ""
            Dim BranchCode As String = ""
            Dim ProjectName As String = ""
            Dim QuoteCode As String = "" '04-24-15 JTC
            Dim Status As String = "" '04-24-15 JTC
            Dim TypeC As String = "" '04-27-15 JTC
            Dim SLSQ As String = ""
            Dim RetrCode As String = ""
            Dim SLSQT As String = "" 'SLS1 to SLSQT '04-30-15 JTC
            Dim Distributor As String = ""
            Dim Sell As String = ""
            Dim NameCode66 As String = ""
            Dim StockJob As String = ""
            Dim ARCH As String = ""
            Dim ENG As String = ""
            Dim CONTR As String = ""
            Dim Comments As String = ""
            Dim SelectCode As String = ""
            Dim CSR As String = ""
            Dim RecType As String = "*"
            Dim NameCode As String = ""
            Dim LeadTime As String = ""
            Dim Contact As String = ""
            Dim SKIPDUPLICATECNT As Integer = 0
            Try
                objExcel = CreateObject("Excel.Application")
                objExcel.DisplayAlerts = False
                objBook = objExcel.Workbooks.Add()
                objBooks = objExcel.Workbooks
                objBook = objBooks(1)
                objSheet = objBook.ActiveSheet()
                objSheet.Name = "Excel Quote FollowUp" ' "SESCO Job List"
                '04-24-15 JTC If frmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then objSheet.Name = "Excel Quote FollowUp" '04-22-15 JTC
                objSheets = objBook.Worksheets
                StartingLine = 0 : EndingLine = 0 '10-19-11
5038:
                Dim ExcelHdgArrayProd() As String
                Dim ExcelDataArrayProd() As String
                'If SESCO = True Then
                '04-24-15 oldExcelHdgArrayProd = New String() {"", "PROJECT", "SPECCODE", "ARCHITECT", "ENGINEER", "SLS1", "SLSQ", "DISTRIBUTOR", "QT-CONTACT", "CONTRACTOR", "MISC", "QT-AMT", "BIDDATE"}
                ''04-24-15 JTC add QuoteCode, StatuscboSortPrimarySeq.Text = "Excel Quote FollowUp" 
                'ExcelHdgArrayProd = New String() {"", "Project Name", "QuoteCode", "RetrCode", "Status", "ARCHITECT", "ENGINEER", "SLS1", "SLSQ", "Quote To:", "QT-CONTACT", "CONTRACTOR", "MISC", "QT-AMT", "BIDDATE"}
                'ExcelDataArrayProd = New String() {"", "JobName", "QuoteCode", "RetrCode", "Status", "Architect", "Engineer", "SLSCode", "SLSQ", "Distributor", "Comments", "Contractor", "LeadTime", "Sell", "BidDate"}
                ExcelHdgArrayProd = New String() {"", "Project Name", "QuoteCode", "RetrCode", "Status", "BidDate", "SLSQT", "SLSQ", "Sell", "Quote To:", "TypeC", "QT-CONTACT", "CONTRACTOR", "ARCHITECT", "ENGINEER", "LeadTime"} '04-24-15 JTC
                ExcelDataArrayProd = New String() {"", "JobName", "QuoteCode", "RetrCode", "Status", "BidDate", "SLSCode", "SLSQ", "Sell", "Distributor", "TypeC", "Comments", "Contractor", "Architect", "Engineer", "LeadTime"} '04-24-15 JTC
                'ProjectCustID, ProjectID, QuoteCode, NCode, FirmName, ContactName, SLSCode, Got, Typec, MFGQuoteNumber, Cost, Sell, Comm, Overage, QuoteToDate, OrdDate, NotGot, Comments, SPANumber, SpecCross, LotUnit, LPCost, LPSell, LPComm, LampsIncl, Terms, FOB, QuoteID, BranchCode, LeadTime, LastChgDate, LastChgBy, Requested, FileName, QuoteID, ProjectID, QuoteCode, EntryDate, BidDate, BidTime, SLSQ, LotUnit, Status, StockJob, EnteredBy, CSR, LastChgBy, Cost, Sell, Comm, HeaderTab, RetrCode, LinesYN, SelectCode, Password, LPCost, LPSell, PRADate, EstDelivDate, FollowBy, OrderEntryBy, ShipmentBy, Remarks, LightingGear, Dimming, LastDateTime, BidBoard, BranchCode, Address, Address2, City, State, Zip, Country, Location, LeadTime, LockedBy, SourceQuote, SpecCross, Probability, MarketSegment, TypeOfJob, SpecCredit, SubmCover, SubmSinglePDF, JobName, LastSaveDate, LockOut, DISTRIBUTOR, Engineer, Architect
                'ExcelDataArrayProd = New String() {"", "JobName", "RetrCode", "Architect",
                '06-03-11 Retriev not   SPEC CODE            ARC              ENG                 SLS1                  SLSQ               DIST               Comment                 CONT             MISC  LeadTime=23                   AMT                              BID DATE
                'A$ = QM.Proj & vbTab & QM.Retriev & vbTab & FF$(15) & vbTab & FF$(16) & vbTab & QM.SlsMan1 & vbTab & QM.QuoteSls & vbTab & QC.Cust & vbTab & QC.Comment & vbTab & FF$(39) & vbTab & FF$(23) & vbTab & Right$(Wspcs$ & Format$(QC.Sell, "######"), 8) & vbTab & ExcelDate(QM.BidDate) & vbTab '06-02-11
ST:
                Saw8MaxCols = ExcelHdgArrayProd.GetUpperBound(0)
                frmQuoteRpt.tgr.UpdateData()
                frmQuoteRpt.tgr.MoveFirst() '02-14-09 

                Dim LastJobCode As String = ""
                Dim LastNCode As String = ""
                Dim LastQuoteCode As String = ""
                Dim RowCnt As Integer = 1
                Dim FirstReal As Int16 = 0
                drQToRow = dsQuoteRealLU.QuoteRealLU.Rows(0)
                For Each drQToRow In dsQuoteRealLU.QuoteRealLU.Rows
                    '                  04-24-15
                    ProjectName = "" : QuoteCode = "" : RetrCode = "" : Status = "" : ARCH = "" : ENG = "" : SLSQT = "" : Distributor = "" : TypeC = "" : Comments = "" : CONTR = "" : LeadTime = "" : Sell = "" : StBidDate = ""

                    If RowCnt = 1 Then
                        For iCol = 1 To Saw8MaxCols : objSheet.Cells(RowCnt, iCol).Value = ExcelHdgArrayProd(iCol) : Next
                        RowCnt += 1
                    End If

                    'CHECK BOX TO ONLY PRINT THE MOST RECENT DIST - SEE BOTTOM OF LOOP
                    If frmQuoteRpt.chkShowLatestCust.Checked = True Then
                        If SKIPDUPLICATECNT > 0 Then SKIPDUPLICATECNT = SKIPDUPLICATECNT - 1 : GoTo NextRecord
                    End If

                    If drQToRow.RowState = DataRowState.Deleted Then Continue For '12-10-09 GoTo 165 '12-09-09 frmQuoteRpt.tgr.MoveNext()
                    If drQToRow.RowState = DataRowState.Deleted Then GoTo NextRecord '02-24-12 frmQuoteRpt.tgr.MoveNext() : Continue For ' 12-09-09 GoTo 235 '06-19-08
                    If FirstReal = 0 Then FirstReal = +1 : frmQuoteRpt.tgr.MoveFirst() '02-10-09

                    SLSQ = drQToRow.SLSQ
                    RetrCode = drQToRow.RetrCode
                    Distributor = drQToRow.NCode
                    LeadTime = drQToRow.LeadTime

                    If IsDBNull(drQToRow("StockJob")) Then drQToRow.StockJob = "" '05-05-15 JTC fix Null 
                    StockJob = drQToRow.StockJob '05-05-15 JTC
                    '04-24-15 JTCProjectName = drQToRow.JobName
                    ProjectName = drQToRow.JobName
                    QuoteCode = drQToRow.QuoteCode '04-24-15
                    Status = drQToRow.Status '04-24-15
                    TypeC = drQToRow.BusinessType 'TypeC = drQToRow.Typec '04-27-15
                    SelectCode = drQToRow.SelectCode
                    NameCode = drQToRow.NCode '05-04-10 
                    Sell = drQToRow.Sell
                    If RealQuoteToAmtON = True And Sell = 0 Then Sell = drQToRow.SellQ '05-05-15 JtC
                    '05-05-15 JTC Sell = VB6.Format(drQToRow.Sell, "####0") '04-24-15 JTC
                    Sell = VB6.Format(Sell, "####0") '05-05-15 JTC
                    If IsDBNull(drQToRow("SLSCode")) Then drQToRow.SLSCode = "" '04-22-15 JTC fix Null 
                    SLSQT = drQToRow.SLSCode '04-30-15 
                    Try
                        '04-22-15 JTC Error Below
                        If Sell = "0" Then Sell = Format(frmQuoteRpt.tgr(frmQuoteRpt.tgr.Row, "Sell1"), "####0")
                    Catch ex As Exception
                        Sell = 0
                    End Try
                    '04-22-15 JTC OUT below
                    Try
                        SLSQT = frmQuoteRpt.tgr(frmQuoteRpt.tgr.Row, "SLSCode")
                    Catch ex As Exception
                        SLSQT = ""
                    End Try
                    Try
                        StBidDate = drQToRow.BidDate.ToShortDateString
                    Catch ex As Exception
                        StBidDate = ""
                    End Try
                    Try
                        Contact = drQToRow.Comments
                    Catch ex As Exception
                        Contact = ""
                    End Try

                    'ARCHITECTS
                    Dim drArchitects() As DataRow = dsSESCOSpecifiers.QuoteRealLU.Select("TypeC = 'A' and QuoteID = '" & drQToRow.QuoteID & "'")
                    For Each dr As dsSaw8.QuoteRealLURow In drArchitects
                        If ARCH = "" Then ARCH += dr.FirmName Else ARCH += "," & dr.FirmName
                    Next

                    'ENGINEERS
                    Dim drEngineers() As DataRow = dsSESCOSpecifiers.QuoteRealLU.Select("TypeC = 'E' and QuoteID = '" & drQToRow.QuoteID & "'")
                    For Each dr As dsSaw8.QuoteRealLURow In drEngineers
                        If ENG = "" Then ENG += dr.FirmName Else ENG += "," & dr.FirmName
                    Next

                    'CONTRACTORS
                    Dim drContractors() As DataRow = dsSESCOSpecifiers.QuoteRealLU.Select("TypeC = 'T' and QuoteID = '" & drQToRow.QuoteID & "'")
                    For Each dr As dsSaw8.QuoteRealLURow In drContractors
                        If CONTR = "" Then CONTR += dr.FirmName Else CONTR += "," & dr.FirmName
                    Next

                    objSheet.Cells(RowCnt, 1) = ProjectName : objSheet.Columns(1).ColumnWidth = 30
                    objSheet.Cells(RowCnt, 2) = QuoteCode : objSheet.Columns(2).ColumnWidth = 15 '04-24-15 
                    objSheet.Cells(RowCnt, 3) = RetrCode : objSheet.Columns(3).ColumnWidth = 15
                    objSheet.Cells(RowCnt, 4) = Status : objSheet.Columns(4).ColumnWidth = 8
                    objSheet.Cells(RowCnt, 5) = StBidDate : objSheet.Columns(5).ColumnWidth = 10
                    objSheet.Cells(RowCnt, 6) = SLSQT : objSheet.Columns(6).ColumnWidth = 5
                    objSheet.Cells(RowCnt, 7) = SLSQ : objSheet.Columns(7).ColumnWidth = 5
                    objSheet.Cells(RowCnt, 8) = Sell : objSheet.Columns(8).ColumnWidth = 10
                    objSheet.Cells(RowCnt, 9) = Distributor : objSheet.Columns(9).ColumnWidth = 12 'Quote To:
                    objSheet.Cells(RowCnt, 10) = TypeC : objSheet.Columns(10).ColumnWidth = 7
                    objSheet.Cells(RowCnt, 11) = Contact : objSheet.Columns(11).ColumnWidth = 10
                    objSheet.Cells(RowCnt, 12) = CONTR : objSheet.Columns(12).ColumnWidth = 20
                    objSheet.Cells(RowCnt, 13) = ARCH : objSheet.Columns(13).ColumnWidth = 20
                    objSheet.Cells(RowCnt, 14) = ENG : objSheet.Columns(14).ColumnWidth = 20
                    objSheet.Cells(RowCnt, 15) = LeadTime : objSheet.Columns(15).ColumnWidth = 15


                    LastNCode = drQToRow.NCode

                    RowCnt += 1

                    If frmQuoteRpt.chkShowLatestCust.Checked = True Then
                        'WE SORT BY QUOTE TO DATE SO THE MOST RECENT IS THE FIRST ONE WE HIT - SKIP THE REST OF THEM 
                        Dim drDIST() As DataRow = dsQuoteRealLU.QuoteRealLU.Select("NCODE = '" & drQToRow.NCode & "' and QuoteID = '" & drQToRow.QuoteID & "'")
                        If drDIST.Length > 1 Then
                            SKIPDUPLICATECNT = drDIST.Length - 1
                        End If
                    End If
NextRecord:
165:                frmQuoteRpt.tgr.MoveNext() '02-14-09
                Next 'drQToRow
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                objSheet.rows(1).insert() '04-28-15 JTC Add Header
                objSheet.Cells(1, 1) = "Excel Quote FollowUp---Sequence = " & frmQuoteRpt.txtSortSecondarySeq.Text & " Select Criteria: " & SelectionText & "  UserID = " & UserID & "  Report Date = " & Format(Now, "MM/dd/yyyy") '04-28-15 
                'Debug.Print(frmQuoteRpt.txtSortSecondarySeq.text) frmQuoteRpt.txtSortSeqV.Text
                'If frmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-22-14 JTC
                '    objSheet.Cells(1, 1) = "Project Name" '04-22-15 JTC 
                '    objSheet.Cells(1, 7) = "Quote To:"
                'End If
9299:           'SAVE THE EXCEL DOCUMENT
                Try
                    '11-30-15 LET THEM SAVE IT ANYWHERE
                    '                    Dim FileNum As Short = 3 '04-22-15 JTC SESCO uses 003
                    '                    Dim FileNumStr As String = "003"
                    'GetNextNum:
                    '                    FileName = UserPath & "DATA\QUTRP" & FileNumStr & ".Xls" '04-22-15 JTC
                    '                    If My.Computer.FileSystem.FileExists(FileName) Then
                    '                        'If FileNum > 13 Then MsgBox("Too many -DATA\QUTRP003.Xls - Files " & vbCrLf & "Delete Some QUTRP* files from: " & UserPath & "DATA\") : GoTo 6000 '04-24-15 JTC
                    '                        If FileNum > 13 Then MessageBox.Show("Too many -DATA\QUTRP003.Xls - Files " & vbCrLf & "Delete Some QUTRP* files from: " & UserPath & "DATA\", "ExselQuoteFollowUp", MessageBoxButtons.OK, MessageBoxIcon.Error) : GoTo 6000 '04-24-15 JTC
                    '                        FileNum = FileNum + 1 : FileNumStr = Format(FileNum, "000") : GoTo GetNextNum
                    '                    End If
                    '                    objBook.SaveAs(FileName, 39, , , False, False, , False, False, , , )
                    '                    '04-22-15 JTC Show Sesco or Me.cboSortPrimarySeq.Text = "Excel Quote FollowUp"
                    '                    US = "Excel Quote FollowUp"
                    '11-30-15 LET THEM SAVE IT ANYWHERE

                    MessageBox.Show("All Done   Created " & FileName & vbCrLf & US & " will display Shortly.")

                    objExcel.VISIBLE = True
                    GoTo 6000
                Catch ex As Exception  '11-01-12 
                    MessageBox.Show("Error Saving Excel Quote FollowUp (VQRT)" & vbCrLf & ex.Message & vbCrLf & FileName & vbCrLf & "Make Sure Excel does not have the file open!" & vbCrLf & "If the problem persists call Multimicro for support", "ExselQuoteFollowUp", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    objExcel.VISIBLE = False
                End Try

            Catch ex As Exception
                MessageBox.Show("Error in Excel Quote FollowUP (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "ExcelQuoteFollowup", MessageBoxButtons.OK, MessageBoxIcon.Error)
                objExcel.VISIBLE = True
            End Try

6000:

ExitReportLoop:
        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "ExcelQuoteFollowUp Report" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop
        End Try
Exit_Done:  '#End

    End Sub
    Public Sub PrintSESCOJobListRealReportQutTO()
        Try '#Top '02-22-12 JTC PrintSESCOJobListRealReportQutTO(
            Dim B As String = ""
            Dim Saw8MaxCols As Int16
            Static EndingLine As Short
            Static StartingLine As Short
            Dim iCol As Int32 '        'For iRow = 0 To 1 'ProdDataArray
            Dim objExcel As Object = Nothing
            Dim objBooks As Object = Nothing
            Dim objSheets As Object = Nothing
            Dim objSheet As Object = Nothing
            Dim objBook As Object = Nothing
            Dim A As String = ""
            Dim FileName As String = ""
            Dim StBidDate As String = ""
            Dim BranchCode As String = ""
            Dim ProjectName As String = ""
            Dim SLSQ As String = ""
            Dim RetrCode As String = ""
            Dim SLS1 As String = ""
            Dim Distributor As String = ""
            Dim Sell As String = ""
            Dim NameCode66 As String = ""
            Dim StockJob As String = ""
            Dim ARCH As String = ""
            Dim ENG As String = ""
            Dim CONTR As String = ""
            Dim Comments As String = ""
            Dim SelectCode As String = ""
            Dim CSR As String = ""
            Dim RecType As String = "*"
            Dim NameCode As String = ""
            Dim LeadTime As String = ""
            Dim Contact As String = ""
            Dim SKIPDUPLICATECNT As Integer = 0
            Try
                objExcel = CreateObject("Excel.Application")
                objExcel.DisplayAlerts = False
                objBook = objExcel.Workbooks.Add()
                objBooks = objExcel.Workbooks
                objBook = objBooks(1)
                objSheet = objBook.ActiveSheet()
                objSheet.Name = "SESCO Job List"
                'If frmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then objSheet.Name = "Excel Quote FollowUp" '04-22-15 JTC
                objSheets = objBook.Worksheets
                StartingLine = 0 : EndingLine = 0 '10-19-11
5038:
                Dim ExcelHdgArrayProd() As String
                Dim ExcelDataArrayProd() As String
                'If SESCO = True Then
                ExcelHdgArrayProd = New String() {"", "PROJECT", "SPECCODE", "ARCHITECT", "ENGINEER", "SLS1", "SLSQ", "DISTRIBUTOR", "QT-CONTACT", "CONTRACTOR", "MISC", "QT-AMT", "BIDDATE"}
                ''04-22-15 JTCcboSortPrimarySeq.Text = "Excel Quote FollowUp" 
                'ExcelHdgArrayProd = New String() {"", "Project Name", "SPECCODE", "ARCHITECT", "ENGINEER", "SLS1", "SLSQ", "Quote To:", "QT-CONTACT", "CONTRACTOR", "MISC", "QT-AMT", "BIDDATE"}
                'ProjectCustID, ProjectID, QuoteCode, NCode, FirmName, ContactName, SLSCode, Got, Typec, MFGQuoteNumber, Cost, Sell, Comm, Overage, QuoteToDate, OrdDate, NotGot, Comments, SPANumber, SpecCross, LotUnit, LPCost, LPSell, LPComm, LampsIncl, Terms, FOB, QuoteID, BranchCode, LeadTime, LastChgDate, LastChgBy, Requested, FileName, QuoteID, ProjectID, QuoteCode, EntryDate, BidDate, BidTime, SLSQ, LotUnit, Status, StockJob, EnteredBy, CSR, LastChgBy, Cost, Sell, Comm, HeaderTab, RetrCode, LinesYN, SelectCode, Password, LPCost, LPSell, PRADate, EstDelivDate, FollowBy, OrderEntryBy, ShipmentBy, Remarks, LightingGear, Dimming, LastDateTime, BidBoard, BranchCode, Address, Address2, City, State, Zip, Country, Location, LeadTime, LockedBy, SourceQuote, SpecCross, Probability, MarketSegment, TypeOfJob, SpecCredit, SubmCover, SubmSinglePDF, JobName, LastSaveDate, LockOut, DISTRIBUTOR, Engineer, Architect
                ExcelDataArrayProd = New String() {"", "JobName", "RetrCode", "Architect", "Engineer", "SLSCode", "SLSQ", "Distributor", "Comments", "Contractor", "LeadTime", "Sell", "BidDate"}
                '06-03-11 Retriev not   SPEC CODE            ARC              ENG                 SLS1                  SLSQ               DIST               Comment                 CONT             MISC  LeadTime=23                   AMT                              BID DATE
                'A$ = QM.Proj & vbTab & QM.Retriev & vbTab & FF$(15) & vbTab & FF$(16) & vbTab & QM.SlsMan1 & vbTab & QM.QuoteSls & vbTab & QC.Cust & vbTab & QC.Comment & vbTab & FF$(39) & vbTab & FF$(23) & vbTab & Right$(Wspcs$ & Format$(QC.Sell, "######"), 8) & vbTab & ExcelDate(QM.BidDate) & vbTab '06-02-11
ST:
                Saw8MaxCols = ExcelHdgArrayProd.GetUpperBound(0)
                frmQuoteRpt.tgr.UpdateData()
                frmQuoteRpt.tgr.MoveFirst() '02-14-09 

                Dim LastJobCode As String = ""
                Dim LastNCode As String = ""
                Dim LastQuoteCode As String = ""
                Dim RowCnt As Integer = 1
                Dim FirstReal As Int16 = 0
                drQToRow = dsQuoteRealLU.QuoteRealLU.Rows(0)
                For Each drQToRow In dsQuoteRealLU.QuoteRealLU.Rows

                    ProjectName = "" : RetrCode = "" : ARCH = "" : ENG = "" : SLS1 = "" : Distributor = "" : Comments = "" : CONTR = "" : LeadTime = "" : Sell = "" : StBidDate = ""

                    If RowCnt = 1 Then
                        For iCol = 1 To Saw8MaxCols : objSheet.Cells(RowCnt, iCol).Value = ExcelHdgArrayProd(iCol) : Next
                        RowCnt += 1
                    End If

                    'CHECK BOX TO ONLY PRINT THE MOST RECENT DIST - SEE BOTTOM OF LOOP
                    If frmQuoteRpt.chkShowLatestCust.Checked = True Then
                        If SKIPDUPLICATECNT > 0 Then SKIPDUPLICATECNT = SKIPDUPLICATECNT - 1 : GoTo NextRecord
                    End If

                    If drQToRow.RowState = DataRowState.Deleted Then Continue For '12-10-09 GoTo 165 '12-09-09 frmQuoteRpt.tgr.MoveNext()
                    If drQToRow.RowState = DataRowState.Deleted Then GoTo NextRecord '02-24-12 frmQuoteRpt.tgr.MoveNext() : Continue For ' 12-09-09 GoTo 235 '06-19-08
                    If FirstReal = 0 Then FirstReal = +1 : frmQuoteRpt.tgr.MoveFirst() '02-10-09

                    SLSQ = drQToRow.SLSQ
                    RetrCode = drQToRow.RetrCode
                    Distributor = drQToRow.NCode
                    LeadTime = drQToRow.LeadTime
                    StockJob = drQToRow.Sell.ToString
                    ProjectName = drQToRow.JobName
                    SelectCode = drQToRow.SelectCode
                    NameCode = drQToRow.NCode '05-04-10 
                    Sell = drQToRow.Sell
                    Sell = VB6.Format(drQToRow.Sell, "####0") '04-24-15 JTC
                    If IsDBNull(drQToRow("SLSCode")) Then drQToRow.SLSCode = "" '04-22-15 JTC fix Null 
                    SLS1 = drQToRow.SLSCode
                    Try
                        '04-22-15 JTC Error Below
                        If Sell = "0" Then Sell = Format(frmQuoteRpt.tgr(frmQuoteRpt.tgr.Row, "Sell1"), "####0.#0")
                    Catch ex As Exception
                        Sell = 0
                    End Try
                    '04-22-15 JTC OUT below
                    Try
                        SLS1 = frmQuoteRpt.tgr(frmQuoteRpt.tgr.Row, "SLS1")
                    Catch ex As Exception
                        SLS1 = ""
                    End Try
                    Try
                        StBidDate = drQToRow.BidDate.ToShortDateString
                    Catch ex As Exception
                        StBidDate = ""
                    End Try
                    Try
                        Contact = drQToRow.Comments
                    Catch ex As Exception
                        Contact = ""
                    End Try

                    'ARCHITECTS
                    Dim drArchitects() As DataRow = dsSESCOSpecifiers.QuoteRealLU.Select("TypeC = 'A' and QuoteID = '" & drQToRow.QuoteID & "'")
                    For Each dr As dsSaw8.QuoteRealLURow In drArchitects
                        If ARCH = "" Then ARCH += dr.FirmName Else ARCH += "," & dr.FirmName
                    Next

                    'ENGINEERS
                    Dim drEngineers() As DataRow = dsSESCOSpecifiers.QuoteRealLU.Select("TypeC = 'E' and QuoteID = '" & drQToRow.QuoteID & "'")
                    For Each dr As dsSaw8.QuoteRealLURow In drEngineers
                        If ENG = "" Then ENG += dr.FirmName Else ENG += "," & dr.FirmName
                    Next

                    'CONTRACTORS
                    Dim drContractors() As DataRow = dsSESCOSpecifiers.QuoteRealLU.Select("TypeC = 'T' and QuoteID = '" & drQToRow.QuoteID & "'")
                    For Each dr As dsSaw8.QuoteRealLURow In drContractors
                        If CONTR = "" Then CONTR += dr.FirmName Else CONTR += "," & dr.FirmName
                    Next

                    objSheet.Cells(RowCnt, 1) = ProjectName : objSheet.Columns(1).ColumnWidth = 30
                    objSheet.Cells(RowCnt, 2) = RetrCode : objSheet.Columns(2).ColumnWidth = 15
                    objSheet.Cells(RowCnt, 3) = ARCH : objSheet.Columns(3).ColumnWidth = 30
                    objSheet.Cells(RowCnt, 4) = ENG : objSheet.Columns(4).ColumnWidth = 30
                    objSheet.Cells(RowCnt, 5) = SLS1 : objSheet.Columns(5).ColumnWidth = 5
                    objSheet.Cells(RowCnt, 6) = SLSQ : objSheet.Columns(6).ColumnWidth = 5
                    objSheet.Cells(RowCnt, 7) = Distributor : objSheet.Columns(7).ColumnWidth = 12
                    objSheet.Cells(RowCnt, 8) = Contact : objSheet.Columns(8).ColumnWidth = 30
                    objSheet.Cells(RowCnt, 9) = CONTR : objSheet.Columns(9).ColumnWidth = 30
                    objSheet.Cells(RowCnt, 10) = LeadTime : objSheet.Columns(10).ColumnWidth = 30
                    objSheet.Cells(RowCnt, 11) = Sell : objSheet.Columns(11).ColumnWidth = 10
                    objSheet.Cells(RowCnt, 12) = StBidDate : objSheet.Columns(12).ColumnWidth = 10
                    LastNCode = drQToRow.NCode

                    RowCnt += 1

                    If frmQuoteRpt.chkShowLatestCust.Checked = True Then
                        'WE SORT BY QUOTE TO DATE SO THE MOST RECENT IS THE FIRST ONE WE HIT - SKIP THE REST OF THEM 
                        Dim drDIST() As DataRow = dsQuoteRealLU.QuoteRealLU.Select("NCODE = '" & drQToRow.NCode & "' and QuoteID = '" & drQToRow.QuoteID & "'")
                        If drDIST.Length > 1 Then
                            SKIPDUPLICATECNT = drDIST.Length - 1
                        End If
                    End If
NextRecord:
165:                frmQuoteRpt.tgr.MoveNext() '02-14-09
                Next 'drQToRow
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'If frmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then '04-22-14 JTC
                '    objSheet.Cells(1, 1) = "Project Name" '04-22-15 JTC 
                '    objSheet.Cells(1, 7) = "Quote To:"
                'End If
9299:           'SAVE THE EXCEL DOCUMENT
                Try
                    Dim FileNum As Short = 3 '04-22-15 JTC SESCO uses 003
                    Dim FileNumStr As String = "003"
GetNextNum:
                    FileName = UserPath & "DATA\QUTRP" & FileNumStr & ".Xls" '04-22-15 JTC
                    If My.Computer.FileSystem.FileExists(FileName) Then
                        'If FileNum > 10 Then MsgBox("Too many -DATA\QUTRP003.Xls - Files  Delete Some") : GoTo 6000
                        If FileNum > 10 Then MsgBox("Too many -DATA\QUTRP003.Xls - Files " & vbCrLf & "Delete Some QUTRP* files from: " & UserPath & "DATA\") : GoTo 6000 '04-24-15 JTC
                        FileNum = FileNum + 1 : FileNumStr = Format(FileNum, "000") : GoTo GetNextNum
                    End If
                    objBook.SaveAs(FileName, 39, , , False, False, , False, False, , , )
                    '04-22-15 JTC Show Sesco or Me.cboSortPrimarySeq.Text = "Excel Quote FollowUp"
                    If SESCO = True Then
                        US = "SESCO Excel Report"
                        'ElseIf frmQuoteRpt.cboSortPrimarySeq.Text = "Excel Quote FollowUp" Then
                        '    US = "Excel Quote FollowUp"
                    End If
                    MessageBox.Show("All Done   Created " & FileName & vbCrLf & US)
                    objExcel.VISIBLE = True
                    GoTo 6000
                Catch ex As Exception  '11-01-12 
                    MessageBox.Show("Error Saving SESCOJobList (VQRT)" & vbCrLf & ex.Message & vbCrLf & FileName & vbCrLf & "Make Sure Excel does not have the file open!" & vbCrLf & "If the problem persists call Multimicro for support", "SESCOJobListRealReport", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    objExcel.VISIBLE = False
                End Try

            Catch ex As Exception
                MessageBox.Show("Error in SESCOJobList (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "SESCOJobListRealReport", MessageBoxButtons.OK, MessageBoxIcon.Error)
                objExcel.VISIBLE = True
            End Try

6000:

ExitReportLoop:
        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "SESCOJobList Report" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop
        End Try
Exit_Done:  '#End

    End Sub
    Public Sub AutoSizeTableRow(ByVal A As String, ByVal row As Integer, ByVal rt As RenderTable) 'A = "ThisRowOnly" "SetGridWidths"
        'A = "ThisRowOnly" "SetGridWidths"
        If rt.Document Is Nothing Then
            Throw New Exception("The table must be already added to the document")
        End If
        Dim col As Integer
        If A = "ThisRowOnly" Then
            '        For row = 0 To rt.Rows.Count - 1
            For col = 0 To rt.Cols.Count - 1
                If rt.Cells(row, col).RenderObject IsNot Nothing Then
                    Dim s As SizeD = rt.Cells(row, col).RenderObject.CalcSize(Unit.Auto, Unit.Auto)
                    widths(col) = Math.Max(widths(col), s.Width)
                End If
            Next col
            ' Next row
            Exit Sub
        End If
        ' 1. grid line widths are added to the columns' widths, so we must take them into consideration.
        ' 2. for calculations in the document, the maximum width is used, so we do that too.
        ' 3. first and last columns include an extra half-width of a line.
        A = "SetGridWidths"
        Dim wVert As Double = 0
        If rt.Style.GridLines.Vert IsNot Nothing Then
            wVert = rt.Style.GridLines.Vert.Width.ConvertUnit(rt.Document.ResolvedUnit)
        End If
        Dim wLeft As Double = 0
        If rt.Style.GridLines.Left IsNot Nothing Then
            wLeft = rt.Style.GridLines.Left.Width.ConvertUnit(rt.Document.ResolvedUnit)
        End If
        Dim wRight As Double = 0
        If rt.Style.GridLines.Right IsNot Nothing Then
            wRight = rt.Style.GridLines.Right.Width.ConvertUnit(rt.Document.ResolvedUnit)
        End If
        Dim lineWidths As Double = Math.Max(wVert, Math.Max(wLeft, wRight))
        'CNVT .Width = VB6.TwipsToPixels(1440)' Twips/inch=1440  Pixels/inch=96   twips/15 = Pixels  on .Width = and .Height =
        For col = 0 To rt.Cols.Count - 1
            If col = 0 OrElse col = rt.Cols.Count - 1 Then
                rt.Cols(col).Width = New Unit(widths(col) + lineWidths * 1.5, rt.Document.ResolvedUnit)
            Else
                rt.Cols(col).Width = New Unit(widths(col) + lineWidths, rt.Document.ResolvedUnit)
            End If
        Next
        ' the default for a table is 100% of the parent's width, so we must set the width
        ' to auto (which means the sum of all columns' widths).
        rt.Width = Unit.Auto
    End Sub

    Public Sub SetupPrintPreview(ByVal FirmName As String)
        Try
            ' make the document:-+
900:        doc.Clear()
            '04-30-10 JH RT = New C1.C1Preview.RenderTable ' RT  Table is Public
            '04-30-10 JH RT.CellStyle.Padding.All = ".5mm"
            '04-30-10 JH RT.Style.Padding.All = "2mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm" '04-30-10 JH 
            '04-30-10 JH RT.Style.GridLines.All = LineDef.Default
            'SetupPrintPreview
905:        ppv.Doc.PageLayout.PageSettings.Landscape = True
            'Not Used NowRTotals = New C1.C1Preview.RenderText
            ' define PageLayout for the first page
            Dim pl As New PageLayout()
            pl.PageSettings = New C1PageSettings()
            pl.PageSettings.PaperKind = System.Drawing.Printing.PaperKind.Letter
            pl.PageSettings.Landscape = True
            pl.PageSettings.LeftMargin = ".5in" '".25cm"
            pl.PageSettings.RightMargin = ".5in" '".25cm"
            pl.PageSettings.TopMargin = ".5in"
            pl.PageSettings.BottomMargin = ".5in"
            'pl.PageSettings.Width = "8in"
            '12-29-08doc.PageLayouts.FirstPage = pl
            doc.PageLayout = pl

910:

        Catch myException As Exception
            MsgBox(myException.Message & vbCrLf & "SetupPrintPreview" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop
        End Try
    End Sub
    Public Function CheckEmptyDecimal(ByVal StringToCheck As String) As String '09-11-09 
        'Trys to parse a string to decimal if it can returns original string or returns 0
        Return IIf(Decimal.TryParse(StringToCheck, Nothing), StringToCheck, "0")
    End Function
    'Public Function MyControls() As ArrayList
    '    Dim res As New ArrayList '01-24-09
    '    'Public MyControlNamesList As String = ""
    '    Dim MyCtlNames As New ArrayList()
    '    Call MyControlsGetAll(frmQuoteRpt, res, MyCtlNames)
    '    MyCtlNames.Sort()
    '    MyControlNamesList = ""
    '    Dim I As Integer
    '    For I = 0 To MyCtlNames.Count - 1
    '        MyControlNamesList = MyControlNamesList & MyCtlNames(I) & ", "
    '    Next
    '    'Debug.Print(MyControlNamesList)
    '    Return res
    'End Function
    'Public Function MyControlsGetAll(ByVal c As Control, ByVal res As ArrayList, ByRef MyCtlNames As ArrayList)
    '    Dim curControl As Control '01-24-09
    '    For Each curControl In c.Controls
    '        'MyControlNamesList = MyControlNamesList & curControl.Name & ", "
    '        MyCtlNames.Add(curControl.Name)
    '        res.Add(curControl)
    '        'Return MyControlsGetAll(curControl, res, MyCtlNames)
    '    Next
    '    Return MyControlsGetAll(curControl, res, MyCtlNames)
    'End Function
    Public Function CheckForFile(ByVal FileName As String, Optional ByVal DisplayMesssage As Boolean = False) As Boolean
        'Function that Checks for the existance of a file and returns True or False
        'FileName:        Full Path and File Name
        'DisplayMesssage: Boolean to dispaly a message if the file doesn't exist

        Dim fileExists As Boolean
        fileExists = My.Computer.FileSystem.FileExists(FileName)
        If DisplayMesssage = True Then
            If fileExists = False Then MsgBox("Can't Find File. " & FileName)
        End If
        Return fileExists

    End Function
    Public Function UnitMeaSet(ByVal Aprice As String, ByRef UnitMeas As Decimal, ByRef UnitOfMeasure As String) As String '05-15-04
        '05-04-05 JH UnitMeaStr = UnitMeaSet(A, UnitMeas, (tgItem.Columns("UM").Text)) '' C = Hundreds M = Thousands FT =Feet '01-28-04
        On Error Resume Next
        UnitMeas = 1 : UnitMeaSet = "" ' C = Hundreds  M = Thousands  EA = Each
        If Right(Aprice, 2) = "FT" Or Right(Aprice, 2) = "EA" Then If Len(Aprice) > 2 Then If IsNumeric(Left(Aprice, Len(Aprice) - 2)) = 0 Then Exit Function Else GoTo 238 '05-15-04  IsNumeric() = 0 Then "Entry is not Numeric Exit
        If Len(Aprice) > 1 Then If IsNumeric(Left(Aprice, Len(Aprice) - 1)) = 0 Then Exit Function '01-28-04 IsNumeric() = 0 Then "Entry is not Numeric Exit
        If Right(Aprice, 1) = "C" Or Trim(UnitOfMeasure) = "C" Then UnitMeaSet = "C" : UnitMeas = 100 : UnitOfMeasure = "C" '05-15-04    ' C = Hundreds
        If Right(Aprice, 1) = "M" Or Trim(UnitOfMeasure) = "M" Then UnitMeaSet = "M" : UnitMeas = 1000 : UnitOfMeasure = "M" '05-15-04 ' M = Thousands
        If Right(Aprice, 1) = "E" Then UnitMeaSet = "EA" : UnitMeas = 1 : UnitOfMeasure = "E" '05-15-04 ' E = Each Only If it is on Input
238:    If Right(Aprice, 2) = "FT" Or Trim(UnitOfMeasure) = "FT" Then UnitMeaSet = "FT" : UnitMeas = 1 : UnitOfMeasure = "FT" '05-15-04  ' FT = Feet
        If Right(Aprice, 2) = "EA" Then UnitMeaSet = "EA" : UnitMeas = 1 : UnitOfMeasure = "EA" '05-15-04 Only Put EA if it is on Input Screen  ' EA = Each  '08-15-02 WNA

    End Function
    Public Function FormatDate(ByVal DateString As String) As String
        'Converts 20080101 to 010108
        Try
            If Len((DateString)) >= 8 Then
                DateString = Mid(DateString, 5, 2) & Mid(DateString, 7, 2) & Mid(DateString, 3, 2)
            End If

            Return DateString

        Catch ex As Exception
            Return DateString
        End Try

    End Function

    Public Function FormatDate() As String '09-11-07 JH
        'Converts the Current Date to "MM/DD/YYYY"
        'DateString is a Member of: Microsoft.VisualBasic.DateAndTime (represents the current date)
        FormatDate = Left(DateString, 2) & "/" & Mid(DateString, 4, 2) & "/" & Right(DateString, 2)
    End Function


    Function CompSlashes(ByRef FirmName As String) As String
        Static ZC As Short ' Compress and Eliminate Slashes & Backslash
        On Error Resume Next
        'A$ = A$ & "." & "/" & "\" & "&" & "<" & ">" & ":" & "|" & "*" & "?" & Chr$(34) & Chr$(30) 'test
        Dim strA As String = FirmName
        strA = Replace(strA, ".", "") '01-28-09
        strA = Replace(strA, "/", "") '01-28-09
        strA = Replace(strA, "\", "") '01-28-09
        strA = Replace(strA, "&", "") '01-28-09
        strA = Replace(strA, "<", "") '01-28-09
        strA = Replace(strA, ">", "") '01-28-09
        strA = Replace(strA, ":", "") '01-28-09
        strA = Replace(strA, "|", "") '01-28-09
        strA = Replace(strA, "?", "") '01-28-09
        strA = Replace(strA, "*", "") '01-28-09
        strA = Replace(strA, Chr(34), "") '01-28-09
        strA = Replace(strA, ",", "") '01-28-09
        strA = Replace(strA, " ", "") '01-28-09
        GoTo 190
185:
190:
        For ZC = 1 To Len(strA) ' Eliminate if asc < 32
            If Asc(Mid(strA, ZC, 1)) < 32 Then
                strA = Left(strA, ZC - 1) & Mid(strA, ZC + 1, 20) ' 02-20-02
            End If
        Next

        CompSlashes = strA

    End Function

    Public Function GetFirmNameCompSlashes(ByRef B As String, ByRef FirmName As String) As String
        Dim ZC As Short '02-07-07 JH

        ZC = InStr(FirmName, B)
        If ZC <> 0 Then
            GetFirmNameCompSlashes = Left(FirmName, ZC - 1) & Mid(FirmName, ZC + 1, 20)
            Call GetFirmNameCompSlashes(B, FirmName)
        Else
            GetFirmNameCompSlashes = FirmName
        End If

    End Function
    Public Function FormatString(ByRef Line As String) As String
        Dim F As Short

        '07-29-05 JH
        'This function allows the User to use the following symbols: " -quotation marks  ' - apostrophe   , - comma
        'when importing lines from a text file w/o altering the data
        Const ascQuotes As Short = 34
        On Error Resume Next
        If Line = vbNullString Then FormatString = "" : Exit Function
        'The , symbol puts quotes around the String
        If InStr(Line, ",") Then
            If Asc(Left(Line, 1)) = 34 Then Line = Mid(Line, 2, Len(Line) - 2)
        End If
        'The ' symbol puts quotes around the String
        If InStr(Line, "'") Then
            If Asc(Left(Line, 1)) = 34 Then Line = Mid(Line, 2, Len(Line) - 2)
        End If
        'if There Are Quotes within the string (used for 4" )
        F = InStr(Line, """""")
        While F > 1
            If Asc(Mid(Line, F + 1, 1)) = 34 Then
                Line = Left(Line, F) & Mid(Line, F + 2, Len(Line))
            End If
            F = InStr(Line, """""")
        End While
        If Asc(Left(Line, 1)) = 34 Then Line = Mid(Line, 2, Len(Line) - 2)
        If Asc(Right(Line, 1)) = 34 Then Line = Mid(Line, 1, Len(Line) - 1)

        FormatString = Line

    End Function

    Public Sub OpenSQL(ByRef myCon As MySqlConnection) '03-19-07 JH

        Try '02-19-10 UserPath & & FormSetting("Load") must be set before OpenSQL Call OpenSQL(myConnection)**************************************** ****************************************
            'server=localhost;user id=root;Password=saw987;database=saw8;persist security info=True
            Dim FileName As String = UserPath & "DBConnectString.Dat" '02-19-10  *****************

            If ServerPath <> "" Then '01-03-17 
                FileName = ServerPath & "DBConnectString.Dat"
            Else
                FileName = UserPath & "DBConnectString.Dat"
            End If

            If My.Computer.FileSystem.FileExists(FileName) Then ' fill My.Settings.saw8ConnectionString
                '05-04-12 FileClose(3) : FileOpen(3, FileName, OpenMode.Input)
                'If Not EOF(3) Then myConnectionString = LineInput(3) 'Read Only My.Settings.saw8ConnectionString = myConnectionString 'Read Only
                'FileClose(3)
                FileClose(3) : FileOpen(3, FileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared) : If Not EOF(3) Then myConnectionString = LineInput(3)
                FileClose(3) '05-04-12 OpenShare.Shared)
                'MySqlCommand cmd = new MySqlCommand():cmd.CommandTimeout = 60; 600=10 min 900=15 min
                '02-01-14 server=JTC7-PC;database=SAW8;user id=root;password=saw987;port=3306;persist security info=True;allow user variables=true;use procedure bodies=false;default command timeout=120
                I = InStr(myConnectionString.ToUpper, "TIMEOUT") '02-4-14 OpenSQL set myConnectionString timeout=1800" '30 min for Reports
                If I > 0 Then
                    myConnectionString = Left(myConnectionString, I - 1) & "timeout=1800" '1800 =30Min 900=15 min'
                End If
                myCon = New MySqlConnection(myConnectionString)

            Else
                myCon = New MySqlConnection(My.Settings.saw8ConnectionString1)
            End If '02-19-10 ******************************************************************
            myCon.Open()
        Catch myException As Exception
            MsgBox(myException.Message)

        End Try

    End Sub

    Public Sub CloseSQL(ByRef myCon As MySqlConnection) '11-21-07 JH

        Try
            myCon.Close()
        Catch myException As Exception
            MsgBox(myException.Message)

        End Try

    End Sub

    Function CheckCreateFile(ByRef FName As String, ByRef Create As Short) As Boolean
        Static FN As Short
        On Error GoTo Error_Routine '05-01-06 JH
5609:
        Static FileExists As Boolean
        'This function Takes in a String Filename and a Integer Create Flag.
        'This function Returns an Bool True if the File Exists (or created) and a False if it doens't exist
        'Takes:
        '   FName     Filename and Path to Check
        '   Create    0-Dont Create, 1-Create File
        'Returns:
        '             Boolean for If the File Existed

        If Len(Dir(FName)) = 0 Then FileExists = False Else  : FileExists = True

        If FileExists = False And Create = 1 Then
            FN = FreeFile() : FileClose(FN)
            FileOpen(FN, FName, OpenMode.Output) : FileClose(FN)
        End If

        CheckCreateFile = FileExists
        GoTo Exit_Done

Error_Routine: Resume Exit_Done

Exit_Done:

    End Function



    Public Sub SetPrimarySortValues() 'VQRT2.RepType = VQRT2.RptMajorType.RptQutCode
        Dim JOBSER As String
        Dim Sorted As Short
        On Error Resume Next
        SortNeeded = "" '06-18-02 WNA
        'Debug.Print(frmQuoteRpt.txtPrimarySortSeq.Text) 'MFG Follow-Up Report
        Select Case frmQuoteRpt.txtPrimarySortSeq.Text
            Case "Quote Code", "Project Code" '11-24-09
                OrderBy = "Q.QuoteCode" '  "OrderBy = "N.FirmName"
                US = "Quote Code Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptQutCode : UH = "QUOTE CODE SEQUENCE" : Sorted = 0 : JOBSER = ""
                If frmQuoteRpt.pnlPrimarySortSeq.Text = "Planned Project Code" Then UH = "PLANNED PROJECT CODE SEQUENCE"
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then UH = "SPEC CREDIT PROJECT CODE SEQUENCE"
                'Debug.Print(frmQuoteRpt.pnlPrimarySortSeq.Text)
                B = "Print Menu"
                MU = VQRT2.RptMajorType.RptQutCode

            Case "Job Name"
                US = "Project Name Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptProj : UH = "PROJECT NAME SEQUENCE" : QN = 1 : Sorted = 0 : JOBSER = ""
                B = "Print Menu"
                MU = VQRT2.RptMajorType.RptProj

            Case "Bid Date"
                US = "Bid Date Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptBidDate : UH = "BID DATE SEQUENCE" : QN = 3 : Sorted = 0 : JOBSER = ""
                B = "Print Menu"
                MU = VQRT2.RptMajorType.RptBidDate

            Case "Entry Date"
                US = "Entry Date Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptEntryDate : UH = "ENTRY DATE SEQUENCE" : QN = 2 : Sorted = 0 : JOBSER = ""
                B = "Print Menu"
                MU = VQRT2.RptMajorType.RptEntryDate

            Case "Followed By" '03-01-12
                US = "Followed By"
                VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy : UH = "FOLLOWED BY SEQUENCE" : QN = 1 : Sorted = 0 : JOBSER = ""
                B = "Print Menu"
                MU = VQRT2.RptMajorType.RptFollowBy
                OrderBy = "Q.FollowBy" ' RptFollowBy = 16 'FollowedBy 03-01-12  SubSProj = 1 'JobName

            Case "Entered By" '05-14-13
                US = "Entered By"
                VQRT2.RepType = VQRT2.RptMajorType.RptEnteredBy : UH = "ENTERED BY SEQUENCE" : QN = 1 : Sorted = 0 : JOBSER = ""
                B = "Print Menu"
                MU = VQRT2.RptMajorType.RptEnteredBy
                OrderBy = "Q.EnteredBy" '05-14-13  RptFollowBy = 16 'FollowedBy 03-01-12  SubSProj = 1 'JobName

            Case "Salesman"
                '      If MARK Then ''MARK = Swap Status & Salesman
                '         US$ = "Status Sequence" 'MARK
                '         RepType = RptSalesman: UH$ = "STATUS"
                '      Else
                US = "Salesman Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptSalesman : UH = "SALESMAN"
                MU = VQRT2.RptMajorType.RptSalesman
                SortNeeded = "YES"
                B = "SEC"
                '   If MFG Then
                '     If DAYB Then
                '        Case "Status"
                '     Else
            Case "Rep Number"
            Case "Status"
                'End If
                '      If MARK Then 'MARK = Swap Status & Salesman
                '         US$ = "Salesman Sequence"
                '         RepType = RptStatus: UH$ = "SALESMAN"
                '      Else
                US = "Status Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptStatus : UH = "STATUS"
                MU = VQRT2.RptMajorType.RptStatus
                SortNeeded = "YES"
                B = "SEC"
            Case "Descending Dollar"
                US = "Descending $ Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptDescend : UH = "DESCENDING DOLLAR SEQ" : B = "Print Menu"
                MU = VQRT2.RptMajorType.RptDescend
                SortNeeded = "YES"
            Case "Specifier Credit"
                frmQuoteRpt.pnlQutRealCode.Visible = True '05-05-10 
                frmQuoteRpt.txtQutRealCode.Visible = True '05-05-10 
                frmQuoteRpt.pnlSpecifierCode.Visible = True '05-05-10 
                frmQuoteRpt.txtSpecifierCode.Visible = True '05-05-10 
                frmQuoteRpt.pnlQuoteToSls.Visible = True '05-07-10
                frmQuoteRpt.txtQuoteToSls.Visible = True '05-07-10
                US = "Specifier Credit Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptSpecif : UH = "SPECIFIER CREDIT SEQ" : B = "SEC"
                MU = VQRT2.RptMajorType.RptSpecif
                'ZE$ = "12"
                SortNeeded = "YES"
            Case "Location"
                US = "Location Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptLocation : UH = "LOCATION SEQUENCE" : QN = 4 : Sorted = 0 : JOBSER = ""
                B = "Print Menu"
                MU = VQRT2.RptMajorType.RptLocation
            Case "Retrieval Code"
                US = "Retrieval Code Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptRetrieval : UH = "RETRIEVAL CODE SEQUENCE" : B = "SEC"
                SortNeeded = "YES"
                MU = VQRT2.RptMajorType.RptRetrieval

            Case "Market Segment" '07-30-04 JH
                US = "Market Segment Sequence" '10-30-12 Fix .RptMarketSegment
                VQRT2.RepType = VQRT2.RptMajorType.RptMarketSegment : UH = "MARKET SEGMENT SEQUENCE" : B = "SEC"
                SortNeeded = "YES"
                MU = VQRT2.RptMajorType.RptMarketSegment

        End Select
    End Sub
    Public Sub SetSecondarySortValues()
        On Error Resume Next 'Only on Job Qoute Commission Shortage
        'Debug.Print(frmQuoteRpt.ChkSpecifiers.Text)
        RealWithOneMfgCustSortJobName = False '10-13-14 JTCPublic RealWithOneMfgCustSortJobName As Boolean = False ' JTC
        SubSeq = 0 '10-31-12 ""
        'Debug.Print(frmQuoteRpt.txtSecondarySort.Text)
        '08-13-12 JTC Can't do this frmQuoteRpt.ChkSpecifiers.Text = "" ' 07-27-12 Sort Report by Descending Dollar" '07-27-12
        'If frmQuoteRpt.ChkSpecifiers.Text = "Sort Report by Descending Dollar" Then '08-13-12
        '    frmQuoteRpt.ChkSpecifiers.Text = "Add Specifiers (Arch, Eng, Etc) to Reports" '
        '    frmQuoteRpt.ChkSpecifiers.Checked = False  '08-13-12 
        'End If
        If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.txtPrimarySortSeq.Text = "Name Code" And RealCustomer = True And RealManufacturer = False Then '03-11-14
            VQRT2.SubSeq = 0 '03-11-14
            VQRT2.RepType = 0 '03-11-14
        End If


        Select Case frmQuoteRpt.txtSecondarySort.Text
            Case "Enter Date" '11-20-10
                US = US & " - Enter Date"
                UH = UH & "/ Enter Date"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSEnterDate 'SubSEnterDate'SubSProjCode'11-20-10 
                '10-31-12 SubSeq = "Q.EntryDate"
            Case "Project Code", "Quote Code" '11-20-10
                US = US & " - Quote Code"
                UH = UH & "/ Quote Code"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSProjCode 'SubSEnterDate'SubSProjCode
                '10-31-12 SubSeq = "Q.QuoteCode"

            Case "Select-Priority / BidDate" '03-06-12 "Job Name/BidDate"'Select-Priority= Q.SelectCode
                US = US & " - Quote Name / BidDate"
                UH = UH & "/QUOTE FOLLOW UP"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSSelectBidDate '03-03-12
                '10-31-12 SubSeq = "Q.SelectCode, Q.BidDate" '03-06-12
            Case "Job Name"
                US = US & " - Quote Name"
                UH = UH & "/QUOTE FOLLOW UP"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSProj
                '10-31-12 SubSeq = "Q.JobName" '11-23-11 quote.JobName Not project.ProjectName 
                If (RealCustomerOnly = True Or RealWithOneMfgCust = True) And frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then 'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.txtPrimarySortSeq.Text = "Name Code" And RealCustomer = True And RealManufacturer = False Then '03-11-14
                    Resp = MsgBox("Do You Want the Major Sort on Job Name?" & vbCrLf & "No = Name Code / Job Name.", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Do You Want the Major Sort on Job Name?") '11-04-14 JTc
                    If Resp = vbYes Then
                        US = US & " - JobName" '01-20-14 JTC for ALS
                        '11-04-14 JTC VQRT2.SubSeq = VQRT2.SubSortType.SubSProj
                        VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif '11-04-14 JTC
                        VQRT2.RepType = VQRT2.RptMajorType.RptProj
                        frmQuoteRpt.txtPrimarySortSeq.Text = "Job Name" : frmQuoteRpt.txtSecondarySort.Text = "Name Code" '10-13-14 01-20-14
                        RealWithOneMfgCustSortJobName = True '10-13-14 JTCPublic RealWithOneMfgCustSortJobName As Boolean = False '10-13-14 JTC
                        GoTo SKipJobQuestion '11-04-14 JTC Duplicate Major Sort on Job Name? Question
                    End If
                End If
                '10-13-14 JTC Mfg in Job Name Seq OneNCodeOnly
                If RealManufacturer = True And frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then 'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.txtPrimarySortSeq.Text = "Name Code" And RealCustomer = True And RealManufacturer = False Then '03-11-14
                    If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then '10-13-14 JTC Realization By JobName Sequence first RealWithOneMfgCustSortJobName = True
                        'Resp = MsgBox("Do You Want the Major Sort on Job Name?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Do Want the Major Sort on Job Name?") '03-11-14
                        Resp = MsgBox("Do You Want the Major Sort on Job Name?" & vbCrLf & "No = Name Code / Job Name.", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Do You Want the Major Sort on Job Name?") '11-04-14 JTc
                        If Resp = vbYes Then
                            US = US & " - JobName" '01-20-14 JTC for ALS
                            VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif
                            VQRT2.RepType = VQRT2.RptMajorType.RptProj
                            frmQuoteRpt.txtPrimarySortSeq.Text = "Job Name" : frmQuoteRpt.txtSecondarySort.Text = "Name Code" '01-20-14
                            RealWithOneMfgCustSortJobName = True '10-13-14 JTCPublic RealWithOneMfgCustSortJobName As Boolean = False '10-13-14 JTC
                        End If
                    End If
                End If
SKipJobQuestion:  '11-04-14 JTC Duplicate Major Sort on Job Name? Question
            Case "Salesman"
                '      If MARK Then   'MARK = Swap Status & Salesman
                '          US$ = US$ + " - Status" 'MARK
                '          UH$ = UH$ + "/STATUS FOLLOW UP"
                '      Else
                US = US & " - Salesman"
                '10-31-12 SubSeq = "Q.SLSQ"
                UH = UH & "/SALESMAN FOLLOW UP"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSSls
                '05-16-13 JTC Added Realization When sub Sort is salesman they can change tobe Salesman Major Sequence
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.txtPrimarySortSeq.Text = "Name Code" Then '05-16-13
                    Resp = MsgBox("Do Want the Major Sort on Salesman?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Do Want the Major Sort on Salesman?") '05-16-13 
                    If Resp = vbYes Then
                        Resp = MsgBox("Yes = Salesman by Name Code." & vbCrLf & "No = Salesman by Job Name.", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Salesman Major Sort on Salesman?") '05-16-13 
                        If Resp = vbYes Then
                            US = US & " - Salesman/NameCode"
                            '10-31-12 SubSeq = "Q.SLSQ"
                            UH = UH & "/SALESMAN FOLLOW UP"
                            VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif
                            VQRT2.RepType = VQRT2.RptMajorType.RptSalesman
                            frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman" : frmQuoteRpt.txtSecondarySort.Text = "Name Code"
                        Else
                            'No                             If Resp = vbNo 
                            US = US & " - Salesman/JobName" '01-20-14 JTC for ALS
                            '10-31-12 SubSeq = "Q.SLSQ"
                            UH = UH & "/SALESMAN FOLLOW UP"
                            VQRT2.SubSeq = VQRT2.SubSortType.SubSProj '01-21-14
                            VQRT2.RepType = VQRT2.RptMajorType.RptSalesman
                            frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman" : frmQuoteRpt.txtSecondarySort.Text = "Job Name" '01-20-14

                        End If
                    End If
                End If

            Case "Status"
                '      If MARK Then 'MARK = Swap Status & Salesman
                '          US$ = US$ + " - Salesman"
                '          UH$ = UH$ + "/SALESMAN FOLLOW UP"
                '      Else
                US = US & " - Status"
                UH = UH & "/STATUS FOLLOW UP"
                '10-31-12 SubSeq = "Q.Status"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus

            Case "Bid Date"
                US = US & " - Bid Date"
                UH = UH & "/BID DATE FOLLOW UP"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSBidDate
                '10-31-12 SubSeq = "Q.BidDate"
            Case "Descending Dollar"
                US = US & " - Descending $"
                UH = UH & "/DESCEND $ FOLLOW UP"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSDescend
                '10-31-12 SubSeq = "Q.Sell"
                '02-01-14 JTC NO
                'frmQuoteRpt.ChkSpecifiers.Text = "Sort Report by Descending Dollar" '07-27-12
                'frmQuoteRpt.ChkSpecifiers.Checked = True '08-13-12 
            Case "Specifiers"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif
                UH = "SPECIFIER CREDIT SEQ By SALESMAN FOLLOW UP"
                US = "Salesman / Specifier Credit Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptSpecif
                '10-31-12 SubSeq = "PC.Ncode"
                B = "SEC"
                MU = VQRT2.RptMajorType.RptSpecif
                SortNeeded = "YES"

            Case "Retrieval Code"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif
                UH = "RETRIEVAL CODE SEQ By SALESMAN FOLLOW UP"
                US = "Salesman / Retrieval Code Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptRetrieval
                '10-31-12 SubSeq = "Q.RetrCode"
                B = "SEC"
                SortNeeded = "YES"
                MU = VQRT2.RptMajorType.RptRetrieval

            Case "Market Segment"
                VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif
                UH = "Market Segment SEQ By SALESMAN FOLLOW UP"
                US = "Salesman / Retrieval Code Sequence"
                VQRT2.RepType = VQRT2.RptMajorType.RptMarketSegment
                '10-31-12 SubSeq = "P.MarketSegment"
                B = "SEC"
                SortNeeded = "YES"
                MU = VQRT2.RptMajorType.RptMarketSegment

            Case "Spread Sheet by Month" '05-20-13 frmQuoteRpt.txtSecondarySort.Text = "Spread Sheet by Month" '05-20-13
                US = US & " - Quote Code"
                UH = UH & "/ Quote Code"
                'VQRT2.SubSeq = VQRT2.SubSortType.SubS'SubSEnterDate'SubSProjCode
                '10-31-12 SubSeq = "Q.QuoteCode"
                '06-22-15 JTC Add Realization Report Spread Sheet by Year
            Case "Spread Sheet by Year" '06-22-15 frmQuoteRpt.txtSecondarySort.Text = "Spread Sheet by Year
                US = US & " - Quote Code"
                UH = UH & "/ Quote Code"

        End Select
    End Sub
    Public Sub SetupSelectCriteria()
        On Error Resume Next
        If DIST Then
            'frmQuoteRpt.pnlCSR.Visible = False
            'frmQuoteRpt.pnlSlsSplits.Visible = False
            'frmQuoteRpt.pnlLotUnit.Visible = False
            'frmQuoteRpt.pnlSpecCross.Visible = False
            'frmQuoteRpt.pnlStkJob.Visible = False
            'frmQuoteRpt.pnlSltCode.Visible = False
            'frmQuoteRpt.txtCSR.Visible = False
            'frmQuoteRpt.txtSlsSplit.Visible = False
            'frmQuoteRpt.cboLotUnit.Visible = False
            'frmQuoteRpt.cboSpecCross.Visible = False
            'frmQuoteRpt.cboStockJob.Visible = False
            'frmQuoteRpt.txtSelectCode.Visible = False
            frmQuoteRpt.pnlCSRdist.Visible = True
            frmQuoteRpt.txtCSRofCust.Visible = True
        End If
    End Sub
    Public Function ShortDate(ByRef OrigDate As String) As String
        On Error Resume Next
        ShortDate = Right(OrigDate, 4) & Mid(OrigDate, 3, 2)
    End Function
    Public Function ExcelDate(ByRef OrigDate As String) As String '09-13-05
        On Error Resume Next
        If Asc(Left(OrigDate, 1)) = 0 Then ExcelDate = "" : Exit Function '09-13-05 0 = Nulls
        If Trim(OrigDate) <> "" Then
            'YYYYMMDD to MM/DD/YYYY
            ExcelDate = Mid(OrigDate, 5, 2) & "/" & Right(OrigDate, 2) & "/" & Left(OrigDate, 4) '09-13-05
        Else
            ExcelDate = Trim(OrigDate)
        End If
    End Function


    Public Sub FillPrimarySortCombo()
        On Error Resume Next
        frmQuoteRpt.cboSortPrimarySeq.Text = ""
        frmQuoteRpt.cboSortPrimarySeq.Items.Clear()
        'If DebugOn ThenDebug.Print(frmQuoteRpt.pnlTypeOfRpt.Text.ToString)
        'Debug.Print(frmQuoteRpt.pnlTypeOfRpt.Text)
        Select Case frmQuoteRpt.pnlTypeOfRpt.Text
            '"Planned Projects","Quote Summary", "Project Shortage Report","Product Sales History - Line Items""Terr Spec Credit Report"

            'Me.C1SuperTooltip1.SetToolTip(
            Case "Realization"
                frmQuoteRpt.C1SuperTooltip1.SetToolTip(frmQuoteRpt.cboSortPrimarySeq, "From the Quote To Tab in the Quotation System. (C = Customer, M = Mfg/Agent) From Specifier Tab (A = Architect, E = Engineer, S = Specifier, T = Contractor, O = Other)") '11-29-11 JTC 05-03-05 JH
                '11-29-11 'QuoteTo Records are M,C,O  Specifiers are A,E,S,T,O
                'Not Used *******##### use cbo
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Customer")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Manufacturer")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Salesman/Customer")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Architect")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Engineer")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Specifier")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Contractor")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Other")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("All")
                If My.Computer.FileSystem.FileExists(UserPath & "VQRTSESCOJOBLIST.DAT") Then '02-25-12
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("SESCO Job List Report") '02-25-12
                Else
                    SESCO = False
                    ExcelQuoteFU = True '04-28-14 JTC Public Bool
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("Excel Quote FollowUp") '04-21-15 JTC Chg "SESCO Job List Report" to "Excel Quote FollowUp" Realization
                End If
                'Me.pnlTypeOfRpt.Text = "Project Shortage Report" '03-15-10
            Case "Quote Summary", "Project Shortage Report" '03-15-10
                frmQuoteRpt.C1SuperTooltip1.SetToolTip(frmQuoteRpt.cboSortPrimarySeq, "") '05-03-05 JH
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Quote Code")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Job Name")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Bid Date")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Entry Date")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Salesman")
                '     If MFG Then
                '        If DAYB Then '08-28-01  Leave As Is for DAYB
                '        Else
                '           frmQuoteRpt.cboSortPrimarySeq.AddItem "Rep Number"
                '        End If
                '     Else
                If frmQuoteRpt.pnlTypeOfRpt.Text <> "Project Shortage Report" Then '05-16-10
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("Status")
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("Descending Dollar")
                    '06-29-12 Specifiers are on RealizationfrmQuoteRpt.cboSortPrimarySeq.Items.Add("Specifier Credit")
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("Location")
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("Retrieval Code")
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("Market Segment")
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("Followed By") '03-01-12
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("Entered By") '05-14-13
                    '05-14-15 JTC If BrandReportMfg = "COOP" Then GoTo NoForecasting '10-23-13
                    If DIST = True Then GoTo NoForecasting '07-31-14
                    '11-01-13 On/Off  If BrandReportMfg = "COOP" Then GoTo NoForecasting 
                    '05-14-15 
                    If BrandReportMfg = "PHIL" Or BrandReportMfg = "DAYB" Or BrandReportMfg = "DAY" Or SESCO = True Then
                        ForecastAllMfg = False '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
                    Else 'Not PHIL Not SESCO
                        ForecastAllMfg = True '05-14-15 JTC Public ForecastAllMfg = True Forecasting for MFGs Except Philips and SESCO
                    End If
                    'Debug.Print(SESCO)
                    '07-31-14 If DIST Then  Else frmQuoteRpt.cboSortPrimarySeq.Items.Add("Forecasting") '10-28-13 No Forecasting on Dist
                    '07-31-14 Only Philips Reps have forecasting and must have BrandReportMfg = something
                    frmQuoteRpt.cboSortPrimarySeq.Items.Add("Forecasting")
                End If
NoForecasting:  frmQuoteRpt.chkBrandReport.Visible = True '10-15-13


            Case "Product Sales History - Line Items" '08-18-09 JH 
                frmQuoteRpt.C1SuperTooltip1.SetToolTip(frmQuoteRpt.cboSortPrimarySeq, "")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Catalog # Detail Report - MFG/Cat # Sequence")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("MFG Sub-Totals in Catalog # Sequence")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Catalog # Totals By Month - Spreadsheet")

            Case "Terr Spec Credit Report" '06-18-10 "Out of Terr Spec Credit" 
                frmQuoteRpt.C1SuperTooltip1.SetToolTip(frmQuoteRpt.cboSortPrimarySeq, "")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Quote Summary") '06-18-18 "Spec Project Report")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("MFG Follow-Up Report")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Salesman Follow-Up Report")
                '06-18-10 Done in Quote frmQuoteRpt.cboSortPrimarySeq.Items.Add("Factory Spec Registration Letter")
                '06-18-10 Done in Quote frmQuoteRpt.cboSortPrimarySeq.Items.Add("Rep Spec Registration Letter")

            Case "Other Quote Types" '06-18-10
                frmQuoteRpt.C1SuperTooltip1.SetToolTip(frmQuoteRpt.cboSortPrimarySeq, "")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Planned Projects") '06-18-10 "Submittals" = "T" Other= "O"
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Submittals")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Other Quote Types")
                '    If DefTypeOfJob = "Quotes" Then JT = "Q"
                '    If DefTypeOfJob = "Planned Projects" Then JT = "P"
                '    If DefTypeOfJob = "Spec Credit" Then JT = "S"
                '    If DefTypeOfJob = "Submittals" Then JT = "T"
                '    If DefTypeOfJob = "Other" Then JT = "O"

            Case "Planned Projects" '11-23-09
                frmQuoteRpt.C1SuperTooltip1.SetToolTip(frmQuoteRpt.cboSortPrimarySeq, "")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Planned Project Code")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Job Name")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Bid Date")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Entry Date")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Salesman")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Status")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Descending Dollar")
                '11-24-09 frmQuoteRpt.cboSortPrimarySeq.Items.Add("Specifier Credit")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Location")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Retrieval Code")
                frmQuoteRpt.cboSortPrimarySeq.Items.Add("Market Segment")


        End Select
        frmQuoteRpt.cboSortPrimarySeq.Text = VB6.GetItemString(frmQuoteRpt.cboSortPrimarySeq, 0)
    End Sub
    Public Sub FormatCommamtPercentTabToA_4440(ByRef A As String)
        Dim Commpct As Decimal
        Dim C As String
        Dim CommAmt As Decimal 'Call FormatCommamtPercentTabToA_4440(A$)
        'FormatCommamtPercentTabToA_4440:  '07-22-03
4440:   C = VB6.Format(CommAmt, "#####,###") : If C = "" Then C = "0" '02-26-01 WNA
        A = A & Right(Wspcs & C, 10) & vbTab '11-27-01 9 to 10 Spaces on Second column
        C = VB6.Format(Commpct, "##0.00") : If C = "" Then C = "0"
        A = A & Right(Wspcs & C, 9) & vbTab
    End Sub
    Public Function SafeSQL(ByVal strMySQL As String) As String
        'To avoid SQL Injection
        'Strip out ',\," and other characters to avoid them messing with the Sql Select Statement.
        'Backslash
        If InStr(strMySQL, "\") > 0 Then
            If InStr(strMySQL, "\\") = 0 Then strMySQL = Replace(strMySQL, "\", "\\")
        End If
        'Single Quote - Apostrophe
        If InStr(strMySQL, "'") > 0 Then
            If InStr(strMySQL, "\'") = 0 Then strMySQL = Replace(strMySQL, "'", "\'")
        End If
        'Double Quote
        If InStr(strMySQL, "\" & """") = 0 Then
            If InStr(strMySQL, "\" & """") <> 0 Then strMySQL = Replace(strMySQL, """", "\" & """")
        End If
        Return strMySQL
    End Function
    Public Sub TotPrt9250(ByRef A As String, ByVal Lev As TotalLevels, ByRef RT As C1.C1Preview.RenderTable, ByVal doc As C1PrintDocument) '
        Try  '#Top
TotPrt9250:
9250:       'FixSell, FixCost, FixProfit, LampSell, LampCost, ProfitLamp, CommAmt, Commpct
            '01-25-09Lev=Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            Dim PC As Int16 = 0 'PC = Print Column
            FixProfit = SellFixtureA(Lev) - CostFixtureA(Lev)
            If SellFixtureA(Lev) <> 0 Then FixProfitPer = FixProfit / (SellFixtureA(Lev) + 0.00001) Else FixProfitPer = 0 '02-15-09
            LampProfit = LampSellA(Lev) - LampCostA(Lev)
            If LampSellA(Lev) <> 0 Then LampProfitPer = LampProfit / (LampSellA(Lev) + 0.00001) Else LampProfitPer = 0
            Dim FixMargin As Decimal '07-08-09
            Dim LPMargin As Decimal '07-08-09
            FixMargin = (SellFixtureA(Lev) - CostFixtureA(Lev)) / (SellFixtureA(Lev) + 0.0001) * 100 '07-08-09
            LPMargin = (LampSellA(Lev) - LampCostA(Lev)) / (LampSellA(Lev) + 0.0001) * 100 '07-08-09
            If FixMargin > 900 Then FixMargin = 999 Else If FixMargin < -900 Then FixMargin = -999
            If LPMargin > 900 Then LPMargin = 999 Else If LPMargin < -900 Then LPMargin = -999
            'If DIST e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), Val(tgQh(e.Row, "Cost"))), "###.00") 'Margind
            'If REP e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), Val(tgQh(e.Row, "Comm"))), "###.00") 'Commission %
            '06-12-10 RC = 0
            Dim ColText As String
            Dim ColName As String
            Dim I As Int16
            'Dim Col As Int16
            'Dim Tag As String
            '09-08-08
            '02-11-10 If frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("MFG Sub-Totals") Then
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report") Then '03-19-14
                RT = New RenderTable
                RT.Cells(0, 0).Text = Replace(A, "*", "") & "   Paid = " & Format(CommAmtA(1), "$#####0.00") & "     Unpaid = " & Format(CommAmtA(3), "$#####0.00")
                'RT.Cells(RC, 5).Text = Format(CommAmtA(1), "####0.00") '09-23-12 
                'RT.Cells(0, 0).SpanCols = RT.Cols.Count - 1 '01-24-13 5
                RT.Cells(0, 0).Style.FontBold = True
                RT.Style.Padding.Top = ".10in" : RT.Style.Padding.Bottom = ".10in"
                doc.Body.Children.Add(RT)
                RT = New RenderTable
                RT.Cells(0, 0).Text = " "
                If frmQuoteRpt.chkSalesmanPerPage.Checked = True Then RT.Rows(RC).PageBreakBehavior = BreakEnum.Page
                doc.Body.Children.Add(RT)
                GoTo 9255 'ExitPrtCel
            End If

            If frmQuoteRpt.pnlTypeOfRpt.Text.StartsWith("Product Sales History") Then '02-11-10
                RC = RT.Rows.Count
            Else
                '06-14-10 Done belowRC += 1 '06-12-10 RT = New C1.C1Preview.RenderTable

                '06-12-10 RC = 0 '02-10-10
            End If
            'Debug.Print(RC.ToString)
            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-09-12
                PC = 0                'RT.Cells(RC).Clear() '03-09-12  'Call SetRTSWidth(RT) 'ByRef R As RenderTable) '03-09-12 
                RT.Cells(RC, PC).Text = A '02-04-09 frmQuoteRpt.tg.Splits(0).DisplayColumns(Col).DataColumn.Text  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Left '02-24-09
                RT.Cols(PC).Width = "6.74in" : PC += 1
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                RT.Cells(RC, PC).Text = Format(SellFixtureA(Lev), DecFormat) '02-11-10
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Right
                RT.Cols(PC).Width = "1.06in" : PC += 1
                If DIST Then RT.Cells(RC, PC).Text = Format(CostFixtureA(Lev), DecFormat) Else RT.Cells(RC, PC).Text = Format(FixProfit, DecFormat) '01-18-12
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Right
                RT.Cols(PC).Width = "1.06in" : PC += 1
                RT.Cells(RC, PC).Text = " " : RT.Cols(PC).Width = ".5in"
                doc.Body.Children.Add(RT) : RT = New C1.C1Preview.RenderTable : RC = 0 '03-10-12
                RT.CellStyle.Padding.Left = "1mm" '12-13-12
                RT.CellStyle.Padding.Right = "1mm" '12-13-12
                GoTo AddRCExit '03-09-12
            End If
            'Debug.Print(frmQuoteRpt.cboSortPrimarySeq.Text) ' pnlTypeOfRpt.Text) ' frmQuoteRpt.cboSortPrimarySeq.Text)
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Or frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Then '06-28-14 Moved up Test for Realization first 03-19-14 frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report"
                MaxCol = frmQuoteRpt.tgr.Splits(0).DisplayColumns.Count - 1 '09-10-09 
            ElseIf frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("Catalog # Detail Report") Or frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("MFG Sub-Totals in Catalog # Sequence") Then '09-11-09 
                MaxCol = frmQuoteRpt.tgln.Splits(0).DisplayColumns.Count - 1 '09-10-09 
                '06-28-14 Moved Up ElseIf frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Or frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Then '03-19-14 frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report"
                '    MaxCol = frmQuoteRpt.tgr.Splits(0).DisplayColumns.Count - 1 '09-10-09 
            Else
                MaxCol = frmQuoteRpt.tgQh.Splits(0).DisplayColumns.Count - 1 '09-10-09 
            End If
            '"Catalog # Detail Report - MFG/Cat # Sequence"'frmMenu.optRptByDate.Value Then
            For I = 0 To MaxCol ' 02-03-09 frmFoll.tg.Splits(0).DisplayColumns.Count - 1
                'Debug.Print(frmQuoteRpt.tg.Splits(0).DisplayColumns(Col).Name)
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
                    'Dim col2 As C1.Win.C1TrueDBGrid.C1DisplayColumn = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I) '02-20-09
                    'Tag = col2.DataColumn.Tag
                    ''07-09-09 Col = Tag 'Tag '= original Grid Sequence LnNum% in old program
                    ColText = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).DataColumn.Text  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                    ColName = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Name 'TgName(I) '02-15-09 frmQuoteRpt.tgr.Splits(0).DisplayColumns(Col).Name
                ElseIf frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("Catalog # Detail Report") Or frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("MFG Sub-Totals in Catalog # Sequence") Or frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Then '09-11-09 Then 'If frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("MFG Sub-Totals") Then 03-19-14 frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report"
                    If frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If tglines.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                    If (frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                    ColText = frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).DataColumn.Text.ToString  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                    ColName = frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Name '09-09-09

                Else 'Quote Summary Regular
                    If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If tglines.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                    If (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                    ColText = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).DataColumn.Text  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                    ColName = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name 'TgName(I) '02-15-09 frmQuoteRpt.tg.Splits(0).DisplayColumns(Col).Name
                    'If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-09-12
                    '    PC = 0
                    '    'RT.Cells(RC).Clear() '03-09-12
                    '    RT.Cells(RC, PC).Text = A '02-04-09 frmQuoteRpt.tg.Splits(0).DisplayColumns(Col).DataColumn.Text  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                    '    RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Left '02-24-09
                    '    RT.Cols(PC).Width = "7.25in" : PC += 1
                    '    RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    '    RT.Cells(RC, PC).Text = Format(SellFixtureA(Lev), "########0.00") '02-11-10
                    '    RT.Cols(PC).Width = "1in" : PC += 1
                    '    If DIST Then RT.Cells(RC, PC).Text = Format(CostFixtureA(Lev), "########0.00") Else RT.Cells(RC, PC).Text = Format(FixProfit, "########0.00") '01-18-12
                    '    RT.Cols(PC).Width = "1in"
                    '    doc.Body.Children.Add(RT) : RT = New C1.C1Preview.RenderTable : RC = 0 '03-09-12
                    '    GoTo AddRCExit '03-09-12
                    'End If
                End If
                'Debug.Print(frmQuoteRpt.cboSortPrimarySeq.Text)
                '@#ProjectName0,ProjectID1,QuoteID2,QuoteCode3,EntryDate4,RetrCode5,PRADate6,EstDelivDate7,SLSQ8,Status9,BidDate10,Cost11,Sell12,Margin13,LPCost14,LPSell15,LPMarg16,LotUnit17,StockJob18,CSR19,LastChgBy20,HeaderTab21,LinesYN22,SelectCode23,Password24,FollowBy25,OrderEntryBy26,ShipmentBy27,Remarks28,LightingGear29,Dimming30,LastDateTime31,BidBoard32,EnteredBy33,BidTime34,BranchCode35,Address36,Address237,City38,State39,Zip40,Country41,Location42,LeadTime43,"'02-22-09
                '@#ProjectCustID0,ProjectID1,NCode2,Got3,Typec4,QuoteCode5,ProjectName6,FirmName7,ContactName8,EntryDate9,SLSCode10,Status11,Cost12,Sell13,Margin14,LPCost15,LPSell16,LPMarg17,Overage18,ChgDate19,OrdDate20,NotGot21,Comments22,SPANumber23,SpecCross24,LotUnit25,LampsIncl26,Terms27,FOB28,QuoteID29,BranchCode30,MarketSegment31,MFGQuoteNumber32,BidDate33,SLSQ34,RetrCode35,SelectCode36,LeadTime37,"
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then '07-08-09
                    If frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                    If (frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                ElseIf frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("Catalog # Detail Report") Or frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("MFG Sub-Totals in Catalog # Sequence") Then '09-11-09 Then 'If Not frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("MFG Sub-Totals") Then
                    If frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                    If (frmQuoteRpt.tgln.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                    'Me.pnlTypeOfRpt.Text = "Quote Summary"
                ElseIf frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" Then
                    If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                    If (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                    RT.Cells(RC, PC).Text = ""
                    'Debug.Print(RC.ToString)
                End If
                If PC = 0 Then   'First Column
                    Dim ColSpan As Int16 = 0 '01-29-13
                    If MaxCol > 7 Then
                        For F = 0 To MaxCol - 1
                            'Debug.Print(frmQuoteRpt.pnlTypeOfRpt.Text) 'Product Sales History - Line Items

                            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
                                If frmQuoteRpt.tgr.Splits(0).DisplayColumns(F).Visible = False Then Continue For '01-29-13
                                If (frmQuoteRpt.tgr.Splits(0).DisplayColumns(F).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                                If frmQuoteRpt.tgr.Splits(0).DisplayColumns(F).Name.ToString = "Sell" Then Exit For
                                If frmQuoteRpt.tgr.Splits(0).DisplayColumns(F).Name.ToString = "Cost" Then Exit For '06-22-11
                                ColSpan += 1 ' ColSpan += 1
                            ElseIf frmQuoteRpt.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then '06-26-13
                                If frmQuoteRpt.tgln.Splits(0).DisplayColumns(F).Visible = False Then Continue For '01-29-13
                                If (frmQuoteRpt.tgln.Splits(0).DisplayColumns(F).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                                If frmQuoteRpt.tgln.Splits(0).DisplayColumns(F).DataColumn.Caption = "Book" Then Exit For '06-26-13
                                If frmQuoteRpt.tgln.Splits(0).DisplayColumns(F).Name.ToString = "Sell" Then Exit For
                                If frmQuoteRpt.tgln.Splits(0).DisplayColumns(F).Name.ToString = "Cost" Then Exit For '06-22-11
                                ColSpan += 1 ' ColSpan += 1
                            Else
                                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(F).Visible = False Then Continue For '01-29-13
                                If (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(F).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(F).DataColumn.Caption = "Book" Then Exit For '06-26-13
                                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(F).Name.ToString = "Sell" Then Exit For
                                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(F).Name.ToString = "Cost" Then Exit For '06-22-11
                                ColSpan += 1 ' ColSpan += 1
                            End If
                        Next
                    End If
                    If ColSpan < 5 Then ColSpan = 5
                    'If ColSpan > 9 Then ColSpan = 9 '06-22-11 Inv Split SLS14 fix
                    'RT.Cells(RC, PC).SpanCols = ColSpan - 1 
                    RT.Style.GridLines.All = LineDef.Default '06-14-10 
                    Dim RealizRatio As Decimal = RealizSellAExt(Lev) / (SellFixtureAExt(Lev) + 0.00001)
                    If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
                        '' Public RealizSellExt As Decimal '12-10-09    Public RealizSellAExt(5) As Decimal
                        ' If A <> A Or A <> A & " Realized= " & Format(RealizSellAExt(Lev), DecFormat) & " %= " & Format(RealizRatio, "##0.00") Then
                        '09-21-15 JTC Don't Duplicate Header Info
                        If InStr(A, Trim("Realized=")) = 0 Then
                            A = A & " Realized= " & Format(RealizSellAExt(Lev), DecFormat) & " %= " & Format(RealizRatio, "##0.00")
                        End If
                    End If
                    'Debug.Print(RC)
                    RT.Cells(RC, PC).Text = A '02-04-09 frmQuoteRpt.tg.Splits(0).DisplayColumns(Col).DataColumn.Text  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                    '09-23-15 JTC 
                    'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True) Then
                    '    RT.Style.Font = New Font("Consolas", FontStyle.Bold) ' 
                    '    RT.Cells(RC, PC).Style.Font = New Font("Consolas", FontStyle.Bold)
                    'End If

                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Left '02-24-09
                    RT.Cells(RC, PC).SpanCols = ColSpan - 1  '01-29-13RT.Cells(RC, PC).SpanCols = 5 '01-29-13 RT.Cols.Count - 1 '01-24-13 was=5 '12-10-09 '07-02-09 2 to 4 '02-24-09
                    RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Left
                    '***********************************************
SPECRPTS:           If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report" Then '09-19-12
                        RT.Cells(RC, PC).Text = A & "   Paid = " & Format(CommAmtA(1), "$#####0.00") & "     Unpaid = " & Format(CommAmtA(3), "$#####0.00")
                        RT.Cells(RC, 5).Text = Format(CommAmtA(1), "####0.00") '09-23-12 
                        RT.Rows(RC).Style.BackColor = LemonChiffon
                        RT.Cells(RC, PC).SpanCols = RT.Cols.Count - 1 '01-24-13 5
                        '10-17-10 Onlymake the Row RC Bold
                        'Dim fs As Integer = frmQuoteRpt.FontSizeComboBox.Text '10-17-10 
                        'RT.Cells(RC, PC).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Bold)
                        RC += 1 '09-19-12 
                        GoTo 9255 'ExitPrtCel
MFGFOLLOWUP:        ElseIf frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Then '03-19-14
                        RT = New RenderTable
                        RT.Cells(0, 0).Text = Replace(A, "*", "") & "   Paid = " & Format(CommAmtA(1), "$#####0.00") & "     Unpaid = " & Format(CommAmtA(3), "$#####0.00")
                        'RT.Cells(RC, 5).Text = Format(CommAmtA(1), "####0.00") '09-23-12 
                        'RT.Cells(0, 0).SpanCols = RT.Cols.Count - 1 '01-24-13 5
                        RT.Cells(0, 0).Style.FontBold = True
                        RT.Style.Padding.Top = ".10in" : RT.Style.Padding.Bottom = ".10in"
                        doc.Body.Children.Add(RT)
                        RC += 1 '09-19-12 
                        GoTo 9255 'ExitPrtCel

                    End If
                End If  '12-08-09
                If ColName = "EntryDate" Or ColName = "Qty" Or ColName = "LPCode" Or ColName = "LPCost" Or ColName = "LPSell" Or ColName = "LPQty" Or ColName = "ChgDate" Or ColName = "OrdDate" Then '12-08-09
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '12-08-09
                End If
                If ColName = "Type" Or ColName = "LnCode" Then '02-12-10 
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center
                End If
                'Dim Hdg As String
                ''Cost11,Sell12,Margin13,LPCost14,LPSell15,LPMarg16,LotUnit17,StockJob18,CSR19,LastChgBy20,HeaderTab21,LinesYN22,SelectCode23,Password24,FollowBy25,OrderEntryBy26,ShipmentBy27,Remarks28,LightingGear29,Dimming30,LastDateTime31,BidBoard32,EnteredBy33,BidTime34,BranchCode35,Address36,Address237,City38,State39,Zip40,Country41,Location42,LeadTime43,"'02-22-09
                '@#Q=ProjectName0,ProjectID1,QuoteID2,QuoteCode3,EntryDate4,RetrCode5,PRADate6,EstDelivDate7,SLSQ8,Status9,BidDate10,Cost11,Sell12,Margin13,LPCost14,LPSell15,LPMarg16,LotUnit17,StockJob18,CSR19,LastChgBy20,HeaderTab21,LinesYN22,SelectCode23,Password24,FollowBy25,OrderEntryBy26,ShipmentBy27,Remarks28,LightingGear29,Dimming30,LastDateTime31,BidBoard32,EnteredBy33,BidTime34,BranchCode35,Address36,Address237,City38,State39,Zip40,Country41,Location42,LeadTime43,"'02-22-09
                '@#R=ProjectCustID0,ProjectID1,NCode2,Got3,Typec4,QuoteCode5,ProjectName6,FirmName7,ContactName8,EntryDate9,SLSCode10,Status11,Cost12,Sell13,Margin14,LPCost15,LPSell16,LPMarg17,Overage18,ChgDate19,OrdDate20,NotGot21,Comments22,SPANumber23,SpecCross24,LotUnit25,LampsIncl26,Terms27,FOB28,QuoteID29,BranchCode30,MarketSegment31,MFGQuoteNumber32,BidDate33,SLSQ34,RetrCode35,SelectCode36,LeadTime37,"
                'If ColName = "Qty" Then '07-07-09 Tag = Val(Hdg) Then
                '    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                '    RT.Cells(RC, PC).Text = QuantityA(Lev).ToString '09-12-09 
                'End If
                'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then Hdg = "12" Else Hdg = "11"
                'If DIST Then Hdg = "Cost" Else Hdg = "Comm" '02-05-09
                If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then GoTo SkipComm625 '10-17-10 'Skip
                'Debug.print(frmQuoteRpt.ChkTotalsOnly.Checked)
                'If frmQuoteRpt.chkDetailTotal.Checked = True ThenStop 'If frmQuoteRpt.ChkTotalsOnly.Checked = True ThenStop
                If frmQuoteRpt.ChkTotalsOnly.Checked = True And (ColName = "Comm-%" Or ColName = "Comm-$") Then GoTo PrtCell '11-20-10

                If ColName = "Cost" Or ColName = "Comm-$" Or ColName = "Comm" Then '01-18-12 07-07-09 Tag = Val(Hdg) Then
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    If DIST Then RT.Cells(RC, PC).Text = Format(CostFixtureA(Lev), DecFormat) Else RT.Cells(RC, PC).Text = Format(FixProfit, DecFormat) '01-18-12
                End If
                If ColName = "Ext Cost" Or ColName = "Ext Comm" Or ColName = "Ext Cost" Or ColName = "Ext Comm-$" Or ColName = "Ext Book" Then '09-11-09 
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    'If DIST Then
                    RT.Cells(RC, PC).Text = Format(CostFixtureAExt(Lev), DecFormat)
                End If
                If ColName = "Book" Then '06-26-13 JTC Added Show Book Total TotPrt9250
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    RT.Cells(RC, PC).Text = Format(CostFixtureAExt(Lev), DecFormat)
                End If
SkipComm625:    '10-17-10 
                'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then Hdg = "13" Else Hdg = "12"
                If ColName = "Sell" Then '07-07-09If Tag = Val(Hdg) Then ' ColName = "Sell" Then
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    RT.Cells(RC, PC).Text = Format(SellFixtureA(Lev), DecFormat) '02-11-10
                    '02-12-02 Don't Print Format(SellFixtureA(Lev), "########0.00") '02-11-10
                End If
                If ColName = "Ext Sell" Then '09-11-09 
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    RT.Cells(RC, PC).Text = Format(SellFixtureAExt(Lev), DecFormat)
                End If
                'FixSell, FixCost, FixProfit, LampSell, LampCost, ProfitLamp, CommAmt, Commpct
                If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then GoTo 650 'Skip
                'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then Hdg = "14" Else Hdg = "13"
                'If DIST Then Hdg = "Margin" Else Hdg = "Comm" '02-04-09
                If ColName = "Margin" Or ColName = "Comm-%" Or ColName = "Ext Marg" Or ColName = "Comm-$" Then '12-09 -09 9 If Tag = Val(Hdg) Then ' ColName = Hdg Or ColName = "Comm" Then '02-15-09
                    'If DIST e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), Val(tgQh(e.Row, "Cost"))), "###.00") 'Margind
                    'If REP e.Value = Format(MarginOrCommCalc(Val(tgQh(e.Row, "Sell")), Val(tgQh(e.Row, "Comm"))), "###.00") 'Commission %
                    If DIST Then
                        RT.Cells(RC, PC).Text = Format(FixMargin, "##0.00")
                        If ColName = "Ext Marg" Then RT.Cells(RC, PC).Text = Format(MarginOrCommCalc(SellFixtureAExt(Lev), CostFixtureAExt(Lev)), "###.00") '12-08-09Format(SellFixtureAExt(Lev) - CostFixtureAExt(Lev), "########0") '09-11-09 
                        If ColName = "Margin" Then RT.Cells(RC, PC).Text = Format(MarginOrCommCalc(SellFixtureAExt(Lev), CostFixtureAExt(Lev)), "###.00") 'Margind
                    Else 'Rep
                        'CommAmtA()
                        RT.Cells(RC, PC).Text = Format(FixProfitPer, "##0.00") 'CostFixtureAExt(Lev)
                        If ColName = "Comm-$" Then RT.Cells(RC, PC).Text = Format(CommAmtA(Lev), DecFormat) '01-06-13 02-18-10 02-12-02 Don't Print Format(CommAmtA(Lev), "####0.00")
                        If ColName = "Comm-%" Then RT.Cells(RC, PC).Text = Format(MarginOrCommCalc(SellFixtureAExt(Lev), SellFixtureAExt(Lev) - CommAmtA(Lev)), "###.00") 'Margin'10-17-10 
                        '10-17-10 If ColName = "Comm-%" Then RT.Cells(RC, PC).Text = Format(MarginOrCommCalc(SellFixtureAExt(Lev), CommAmtA(Lev)), "###.00") 'Margind 10-17-10 
                        'Debug.Print(RT.Cells(RC, PC).Text)
                    End If
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    If frmQuoteRpt.chkDetailTotal.Checked = True And (ColName = "Comm-%" Or ColName = "Margin") Then GoTo PrtCell '01-29-13
                End If

650:            If frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("MFG Sub-Totals") Then GoTo PrtCell '09-09-09 
                'LampSell, LampCost, ProfitLamp, CommAmt, Commpct
                'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then Hdg = "15" Else Hdg = "14"
                If ColName = "LPCost" Then '07-07-09  If Tag = Val(Hdg) Then ' If ColName = "LPCost" Then
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    RT.Cells(RC, PC).Text = Format(LampCostA(Lev), DecFormat)
                End If
                'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then Hdg = "16" Else Hdg = "15"
                If ColName = "LPSell" Then '07-07-09  If Tag = Val(Hdg) Then ' If ColName = "LPSell" Then
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    RT.Cells(RC, PC).Text = Format(LampSellA(Lev), DecFormat)
                End If
                If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then GoTo PrtCell '10-17-10 'Skip
                'If DIST Then Hdg = "LPMarg" Else Hdg = "LPComm" '02-04-09
                'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then Hdg = "17" Else Hdg = "16"
                If ColName = "LPComm" Or ColName = "LPMarg" Then '07-07-09  If Tag = Val(Hdg) Then ' If ColName = Hdg Or ColName = "LPComm" Then '02-15-09
                    If DIST Then
                        RT.Cells(RC, PC).Text = Format(LPMargin, "##0.00") '07-08-09 
                    Else
                        RT.Cells(RC, PC).Text = Format(LampProfitPer, "##0.00")
                    End If
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                End If
PrtCell:
                RT.Cells(RC, PC).Style.BackColor = LemonChiffon
                '10-17-10 Only make the Row RC Bold
                Dim fs As Integer = frmQuoteRpt.FontSizeComboBox.Text '10-17-10 
                RT.Cells(RC, PC).Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Bold)
                ''09-23-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
                'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)
                '    RT.Cells(RC, 0).Style.Font = New Font("Courier New", FontStyle.Bold)
                'End If 'Consolas or Courier
                PC += 1 '02-20-09 Dim PC As Int16 'PC = Print Column
            Next
            '10-17-10 Move DownStop :Debug.Print(TgWidth(0).ToString)

            'RT.Cells(RC, 0).SpanCols = 5 '10-17-10 
            '06-14-10RT.Rows(RC).Style.BackColor = Color.L'06-14-10 
            If frmQuoteRpt.pnlTypeOfRpt.Text.StartsWith("Product Sales History") Then '02-11-10
                RC = RT.Rows.Count
                GoTo 9255
            End If
            If frmQuoteRpt.cboSortPrimarySeq.Text.StartsWith("MFG Sub-Totals") Then GoTo 9255 'ExitPrtCell '09-09-09 
            '
            'Debug.Print(RT.Cells(0, 0).Text & RT.Cells(0, 5).Text)
            '06-12-10 RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '12-29-08
            '06-12-10 RT.Style.GridLines.All = LineDef.Default
            '06-17-10 If frmQuoteRpt.chkSalesmanPerPage.CheckState = CheckState.Checked Then
            '06-17-10 RT.Rows(RC).PageBreakBehavior = BreakEnum.Page '06-14-10 
            '06-17-10 End If
            '10-17-10 RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Bold) '05-03-10 JH
            'RT.Style.
            '06-12-10 RT.StretchColumns = StretchTableEnum.LastVectorOnPage '06-01-10
            '06-12-10 doc.Body.Children.Add(RT) '12-29-06
            '06-12-10 Call RTColSize(RT, MaxCol, TgWidth)
            'Call AutoSizeTableRow("ThisRowOnly", RC, RT) '02-03-09 A = "ThisRowOnly" "SetGridWidths"
            'Debug.Print(RC.ToString) 'RT.Rows.Count
            RC += 1
            '06-12-10 RC = 0 '02-10-10
AddRCExit:

            RT.Style.GridLines.All = LineDef.Default
            GoTo 9255 'Exit 02-15-09

        Catch ex As Exception  'Try Catch with  so you can fix exception after the error message **Get gid of  'CatchStop b/4 releasing
            MessageBox.Show("Error in TotPrt9250 (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox(ex.Message.ToString & vbCrLf & "TotPrt9250(VQRT)" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop  'Debug.WriteLine(ex.Message.ToString)
        End Try '
9255:   '#End
    End Sub
    Public Sub SelectHit9500(ByRef Hit As Short, ByRef multsrtrvs() As String) ' ByVal drQRow As dsSaw8.QUTLU1Row
        '#Top
        Try
            Dim K As Short = 0
            Dim SearchStat As String = ""
            Dim SaveHit As Short = 0
            Dim J As Short = 0
            Dim splitret As Short = 0
            Dim I As Short = 0
            Dim retrievalhold As String = ""
            Dim CompAmt As Decimal = 0       'Dim CheckState As Object 'Call SelectHit9500(Hit, multsrtrvs())
            '***********************   SELECT CRITERIA   ********************************
SelectHit9500:
9500:       Hit = 1 'Because 'EntryStart and End is in The Sql
            ' 'Dim drQRow As dsSaw8.QUTLU1Row
            Dim BidDate As Date
            Dim BranchCode As String = ""
            Dim ProjectName As String = ""
            Dim SLSQ As String = ""
            Dim RetrCode As String = ""
            Dim Status As String = ""
            Dim City As String = ""
            Dim State As String = ""
            Dim LastChgBy As String = ""
            Dim StockJob As String = ""
            Dim LotUnit As String = ""
            Dim SelectCode As String = ""
            Dim CSR As String = ""
            Dim RecType As String = "*"
            Dim NameCode As String = ""
            Dim SpecCross As String = "" '02-13-13 drQRow.SpecCross
            Dim MarketSegment As String = "" '02-13-13 drQRow.MarketSegment
            'If frmQuoteRpt.pnlTypeOfRpt.Text = "Product Sales History - Line Items" Then
            '    'Debug.Print(frmQuoteRpt.pnlTypeOfRpt.Text)
            '    SLSQ = drQline.SpecCross
            'End If

            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
                If IsDBNull(drQToRow("BranchCode")) Then drQToRow.BranchCode = "" '04-22-15 JTC fix Null BranchCode =
                BranchCode = drQToRow.BranchCode
                Try
                    If IsDBNull(drQToRow("BidDate")) Then '07-25-13
                        drQToRow.BidDate = "#1/1/1900#"
                    End If
                    If IsDBNull(drQToRow("BidDate")) = False Then '07-25-13 IsDBNull(drQToRow.BidDate) = False Then  '10-13-11
                        BidDate = drQToRow.BidDate
                    End If
                    If IsDBNull(drQToRow("SpecCross")) Then drQToRow.SpecCross = "" '07-25-13
                Catch
                End Try
                ProjectName = drQToRow.JobName ' JobName '09-14-10 ProjectName 01-26-16 THIS WAS COMMENTED OUT.  REALIZATION WILL BE IN THE SQL BUT LEAVE IT HERE UNTIL WE GET THIS ORGANIZED
                'Debug.Print(drQToRow.QuoteCode)
                SLSQ = drQToRow.SLSQ
                RetrCode = drQToRow.RetrCode
                Status = drQToRow.Status
                If IsDBNull(drQToRow("City")) Then drQToRow.City = "" '04-22-15 JTC fix Null 
                City = CType(drQToRow.City, String)
                If IsDBNull(drQToRow("State")) Then drQToRow.State = "" '04-22-15 JTC fix Null 
                State = CType(drQToRow.State, String) 'drQToRow.State
                If IsDBNull(drQToRow("LastChgBy")) Then drQToRow.LastChgBy = "" '04-30-15 JTC fix Null
                LastChgBy = drQToRow.LastChgBy
                StockJob = drQToRow.StockJob
                LotUnit = drQToRow.LotUnit
                SelectCode = drQToRow.SelectCode
                CSR = drQToRow.CSR
                NameCode = drQToRow.NCode '05-04-10
                MarketSegment = drQToRow.MarketSegment '02-13-13
                SpecCross = drQToRow.SpecCross '02-13-13
                'If SpecCross = "C" ThenStop
                '08-20-11 Deleted This because it is in StrSQL
                'Select Case frmQuoteRpt.txtPrimarySortSeq.Text
                '    Case "Customer"
                '        RecType = "C"
            Else
                ' SLSQ = drQRow.SLSCode
                BranchCode = drQRow.BranchCode '06-15-10
                Try
                    If IsDBNull(drQRow("BidDate")) Then '07-25-13
                        drQRow.BidDate = "#1/1/1900#"
                    End If
                    '04-04-12 Could use If 
                    If drQRow.IsBidDateNull = True Then '05-04-12 
                        drQRow.BidDate = "#1/1/1900#"
                    End If
                    If IsDBNull(drQRow.BidDate) = False Then  '10-13-11
                        BidDate = drQRow.BidDate
                    End If
                Catch
                End Try
                'If BidDate = "#12:00:00 AM#" Then BidDate = Nothing 'Null ' "Null" '05-04-12IsDBNull()
                ProjectName = drQRow.JobName '09-14-10 
                SLSQ = drQRow.SLSQ
                RetrCode = drQRow.RetrCode
                Status = drQRow.Status
                City = drQRow.City
                State = drQRow.State
                LastChgBy = drQRow.LastChgBy
                StockJob = drQRow.StockJob
                LotUnit = drQRow.LotUnit
                SelectCode = drQRow.SelectCode
                CSR = drQRow.CSR
                MarketSegment = drQRow.MarketSegment '02-13-13
                SpecCross = drQRow.SpecCross '02-13-13

                If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Then '03-03-12 Chg SpecifierCode to FollowedBY
                    NameCode = drQRow.FollowBy
                    'NameCode = drQRow.SLSQ ' 
                End If
            End If
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
                Dim st As String = drQToRow.Typec
                Hit = 1 : GoTo EndRecType9510 '08-20-11 Deleted This because it is in StrSQL
                '08-20-11 Deleted This because it is in StrSQL
                'If st = "M" Then 
                'If RecType = "*" Then Hit = 1 : GoTo EndRecType9510 '08-20-11 Deleted This because it is in StrSQL
                'If RecType = "C" Then
                '    If RecType = st Then Hit = 1 : GoTo EndRecType9510 Else Hit = 0 : GoTo SelExit9530 '02-14-09 Customer
                'End If
                'If RecType = "M" Then
                '    If RecType = st Then Hit = 1 : GoTo EndRecType9510 Else Hit = 0 : GoTo SelExit9530 '02-14-09 Mfg
                'End If
                'If RecType = "T" Then
                '    If RecType = st Then Hit = 1 : GoTo EndRecType9510 Else Hit = 0 : GoTo SelExit9530 '02-14-09 Contractor
                'End If
                ''All Except D & M
                'If st = "C" Or st = "M" Then Hit = 0 : GoTo SelExit9530
                ''Must be Specifier so give them all But C & M
                'Hit = 1 : GoTo EndRecType9510 '02-14-09 
                ' If drQToRow.Typec <> RecType Then Hit = 0 : GoTo SelExit9530 '02-14-09 
EndRecType9510:
            End If
            'If frmQuoteRpt.DTPicker1StartBid.Text = "1/1/1900" Then GoTo 9510 'OK10-18-10 JTC If StartBidDate = "#1/1/1900#" then ignore BidDate
            'Checked means Include Blank BidDate '04-04-12 Fix Select EntryDate&Biddate&InvludeBlankBid 
            '05-21-12 Fix Object errorDebug.Print(drQRow.QuoteCode) SelectHit9500
            'Debug.Print(frmQuoteRpt.chkBlankBidDates.Visible)
            If frmQuoteRpt.txtPrimarySortSeq.Text = "Forecasting" Then GoTo 9510 '12-12-13 JTC Bypass Bid Date test in SelectHit9500
            If frmQuoteRpt.chkBlankBidDates.CheckState = CheckState.Checked Then '05-21-13 deleted frmQuoteRpt.chkBlankBidDates.Visible = True And 
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then '04-26-12 Fix object Reference Error
                    'No If drQToRow.IsBidDateNull = True Then GoTo 9510 'OK on BidDate 04-04-12 and chkBlankBidDates.CheckState = CheckState.Checked
                Else
                    If drQRow.IsBidDateNull = True Then GoTo 9510 'BidDate OK
                End If
                If Trim(BidDate) = "" Or BidDate = "#1/1/1900#" Or BidDate = "#12:00:00 AM#" Then GoTo 9510 '05-21-12 OK on BidDate
            End If
            If frmQuoteRpt.ChkCheckBidDates.CheckState = CheckState.Checked Then '05-21-13 Check Regular Bid Datea
                'Debug.Print(frmQuoteRpt.DTPicker1StartBid.Text & frmQuoteRpt.DTPicker1EndBid.Text)
                '05-21-13If Trim(BidDate) <> "" And BidDate <> "#1/1/1900#" And BidDate <> "#12:00:00 AM#" Then '11-03-09 "1-1-1900" to "1/1/1900"
                '05-21-13 If Trim(BidDate) = "" Or BidDate = "#1/1/1900#" Or BidDate = "#12:00:00 AM#" Then '11-03-09 "1-1-1900" to "1/1/1900"=
                If Trim(frmQuoteRpt.DTPicker1StartBid.Text) <> "" And Trim(frmQuoteRpt.DTPicker1StartBid.Text) <> "ALL" And Trim(frmQuoteRpt.DTPicker1EndBid.Text) <> "" And Trim(frmQuoteRpt.DTPicker1EndBid.Text) <> "ALL" Then      ''BidDate OK
                    '05-21-12 If frmQuoteRpt.chkBlankBidDates.CheckState = CheckState.Checked Then  Else Hit = 0 : GoTo SelExit9530 '05-04-12
                    If VB6.Format(BidDate, "yyyy-MM-dd") >= VB6.Format(frmQuoteRpt.DTPicker1StartBid.Text, "yyyy-MM-dd") And VB6.Format(BidDate, "yyyy-MM-dd") <= VB6.Format(frmQuoteRpt.DTPicker1EndBid.Text, "yyyy-MM-dd") Then Hit = 1 : GoTo 9510 Else Hit = 0 : GoTo SelExit9530
                    Hit = 0 : GoTo SelExit9530 '05-04-12
                End If
            End If

9510:       ' BidDate OK
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
                Try
                    If IsDBNull(drQToRow("SellQ")) Then '07-25-13
                        drQToRow.SellQ = 0 '
                    End If
                    If IsDBNull(drQToRow.SellQ) = False Then

                    Else
                        CompAmt = 0 '08-28-12 drQToRow.SellQ '08-27-12 JTC Select Realization on Total Quote Amount not sell of each row drQToRow.Sell
                    End If
                Catch
                    drQToRow.SellQ = 0 '08-28-12 JTC Fix Null on Error Realization"
                End Try
                '01-06-14 JTC RealQuoteToAmtON = True  Use Realization QuoteTo Sell Amount on Report
                CompAmt = drQToRow.SellQ '08-27-12 JTC Select Realization on Total Quote Amount not sell of each row drQToRow.Sell
                If RealQuoteToAmtON = True Then
                    If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then  '12-07-16 JH - IF TMPSellQ isn't in the Quote To - MFG sql. 
                        If IsDBNull(drQToRow("TMPSellQ")) Then drQToRow("TMPSellQ") = "" '04-22-15 JTC
                        CompAmt = drQToRow("TMPSellQ") '04-22-15 JTC Fix RealQuoteToAmtON = True  one MFG Use Realization QuoteTo Sell Amount on Report CompAmt = drQToRow("TMPSellQ")
                    Else
                        CompAmt = drQToRow.Sell '12-07-16 
                    End If

                End If
            Else
                CompAmt = drQRow.Sell
            End If
            If BranchReporting Then '06-15-10 
                If Trim(BranchCodeRpt) <> "" And Trim(BranchCodeRpt) <> "ALL" Then '10-30-12 JTC Fix Branch Selection
                    If InStr("," & Trim(BranchCodeRpt) & ",", "," & Trim(BranchCode) & ",") = 0 Then Hit = 0 : GoTo SelExit9530 '06-15-10 
                End If
            End If
            'If MFG Then
            '    If DAYB Then
            '        ' CompAmt! = drQRow.Comm
            '    End If
            'End If
            'Debug.Print(frmQuoteRpt.txtStartQuote.Text)
            'Check Dollar Range Amount  If RealQuoteToAmtON = True then CompAmt = drQToRow.Sell
            If Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "0" And Trim(frmQuoteRpt.txtEndQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtEndQuoteAmt.Text) <> "999999999" Then '"999999999" '03-24-08 JTC Added 9 "999,999,999"
                If CompAmt >= Val(frmQuoteRpt.txtStartQuoteAmt.Text) And CompAmt <= Val(frmQuoteRpt.txtEndQuoteAmt.Text) Then Hit = 1 : GoTo 9515 Else Hit = 0 : GoTo SelExit9530 '09-23-02 WNA
            End If
            If Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "0" Then
                If CompAmt >= Val(frmQuoteRpt.txtStartQuoteAmt.Text) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If
            If Trim(frmQuoteRpt.txtEndQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtEndQuoteAmt.Text) <> "999999999" Then '"999999999" '03-24-08 JTC Added 9 "999,999,999"
                If CompAmt <= Val(frmQuoteRpt.txtEndQuoteAmt.Text) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If
9515:
9520:       If frmQuoteRpt.pnlTypeOfRpt.Text <> "Terr Spec Credit Report" Then '12-01-09 
                If Trim(frmQuoteRpt.txtJobNameSS.Text) <> "" And frmQuoteRpt.txtJobNameSS.Text.Trim <> "ALL" Then '12-09-09 
                    If InStr(ProjectName, Trim(frmQuoteRpt.txtJobNameSS.Text)) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
                End If
            Else
            End If
9525:       'Debug.Print(frmQuoteRpt.txtQuoteToSls.Text)
            If Trim(frmQuoteRpt.txtSalesman.Text) <> "" And Trim(frmQuoteRpt.txtSalesman.Text) <> "ALL" Then
                If InStr("," & Trim(frmQuoteRpt.txtSalesman.Text) & ",", "," & Trim(SLSQ) & ",") Then Hit = 1 Else Hit = 0 : GoTo SelExit9530 '05-05-10 
            End If
            '10-29-12
            If Trim(frmQuoteRpt.txtRetrieval.Text) <> "" And Trim(frmQuoteRpt.txtRetrieval.Text) <> "ALL" Then '10-29-12 JTC Fix Retrieval Code Select Error
                If InStr("," & Trim(frmQuoteRpt.txtRetrieval.Text) & ",", "," & Trim(RetrCode) & ",") Then Hit = 1 Else Hit = 0 : GoTo SelExit9530 '05-05-10 
            End If

            'Debug.Print(NameCode & "    " & frmQuoteRpt.txtSpecifierCode.Text)
            'If NameCode.Trim <> "" ThenStop
            If Trim(frmQuoteRpt.txtSpecifierCode.Text) <> "" And Trim(frmQuoteRpt.txtSpecifierCode.Text) <> "ALL" Then
                If InStr(Trim(frmQuoteRpt.txtSpecifierCode.Text) & ",", NameCode & ",") = False Or Trim(NameCode = "") Then Hit = 0 : GoTo SelExit9530 '03-03-12 or Trim(frmQuoteRpt.txtSpecifierCode.Text)
            End If
            If frmQuoteRpt.txtPrimarySortSeq.Text = "Forecasting" Then GoTo 9526 '01-16-14 JTC Bypass txtQutRealCode.Text =  test in SelectHit9500
            If Trim(frmQuoteRpt.txtQutRealCode.Text) <> "" And Trim(frmQuoteRpt.txtQutRealCode.Text) <> "ALL" Then
                'If DebugOn ThenStop'05-05-10 Added
                If NameCode.Trim = "" Then Hit = 0 : GoTo SelExit9530 '05-07-10 Blank Code so skip 
                If InStr(Trim(frmQuoteRpt.txtQutRealCode.Text) & ",", NameCode & ",") = False Or Trim(NameCode = "") Then Hit = 0 : GoTo SelExit9530 '03-03-12 Or Trim(NameCode = "") 
            End If
9526:
            '''''''''''''''''''''''''''''''''''''''''''''''''
            If Trim(frmQuoteRpt.txtStatus.Text) <> "" And Trim(frmQuoteRpt.txtStatus.Text) <> "ALL" Then
                If Trim(Status) <> "" Then ' allow mult. codes
                    If InStr("," & Trim(frmQuoteRpt.txtStatus.Text) & ",", ",#") Or InStr("," & Trim(frmQuoteRpt.txtStatus.Text) & ",", "*,") Then
                        If InStr(Trim(frmQuoteRpt.txtStatus.Text), Trim(Status)) Then
                            Hit = 1
                        Else
                            SaveHit = 0
                            SI = "," & Trim(frmQuoteRpt.txtStatus.Text) & ","
                            I = InStr(SI, ",#") '01-30-01 WNA
                            If I > 0 Then
                                SearchStat = Mid(SI, I + 2, (InStr(I + 1, SI, ",")) - (I + 2))
                                If InStr(Status, SearchStat) Then SaveHit = 1
                            End If
                            I = InStr(SI, "*,") '01-30-01 WNA
                            If I > 2 Then
                                K = 1
                                Do While K > 0
                                    J = InStr(K + 1, SI, ",")
                                    If J = I + 1 Then
                                        SearchStat = Mid(SI, K + 1, I - K - 1)
                                        Exit Do
                                    Else
                                        K = J
                                    End If
                                Loop
                                If Left(Status, Len(SearchStat)) = SearchStat Then SaveHit = 1
                            End If
                            If SaveHit = 1 Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
                        End If
                    Else 'no wildcards.
                        If InStr(Trim(frmQuoteRpt.txtStatus.Text), Trim(Status)) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530 '09-16-02 WNA
                    End If
                Else
                    Hit = 0 : GoTo SelExit9530
                End If
9528:       End If
            If Trim(frmQuoteRpt.txtState.Text) <> "" And Trim(frmQuoteRpt.txtState.Text) <> "ALL" Then
                If frmQuoteRpt.txtState.Text = State Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If

            If Trim(frmQuoteRpt.txtCity.Text) <> "" And Trim(frmQuoteRpt.txtCity.Text) <> "ALL" Then
                If frmQuoteRpt.txtCity.Text = City Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If

            If Trim(frmQuoteRpt.txtMktSegment.Text) <> "" And Trim(frmQuoteRpt.txtMktSegment.Text) <> "ALL" Then
                If InStr(UCase(MarketSegment), UCase(frmQuoteRpt.txtMktSegment.Text)) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If

            If Trim(frmQuoteRpt.txtLastChgBy.Text) <> "" And Trim(frmQuoteRpt.txtLastChgBy.Text) <> "ALL" Then
                If InStr("," & Trim(frmQuoteRpt.txtLastChgBy.Text) & ",", "," & Trim(LastChgBy) & ",") Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If

            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And Trim(frmQuoteRpt.txtQuoteToSls.Text) <> "" And Trim(frmQuoteRpt.txtQuoteToSls.Text) <> "ALL" Then '07-16-02 WNA
                If InStr("," & Trim(frmQuoteRpt.txtQuoteToSls.Text) & ",", "," & Trim(SLSQ) & ",") Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If

            If Trim(frmQuoteRpt.cboLotUnit.Text) <> "ALL" Then
                If Left$(frmQuoteRpt.cboLotUnit.Text, 1) = Trim$(LotUnit) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If
            'frmQuoteRpt.cboSpecCross.Text = "C"
            If Trim(frmQuoteRpt.cbospeccross.Text) <> "ALL" Then
                If Left$(frmQuoteRpt.cbospeccross.Text, 1) = Trim$(SpecCross) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If
            If Trim(frmQuoteRpt.cboStockJob.Text) <> "ALL" Then
                If Left$(frmQuoteRpt.cboStockJob.Text, 1) = Trim$(StockJob) Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If

            If Trim(frmQuoteRpt.txtSelectCode.Text) <> "" And Trim(frmQuoteRpt.txtSelectCode.Text) <> "ALL" Then
                If frmQuoteRpt.txtSelectCode.Text = SelectCode Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If

            If Trim(frmQuoteRpt.txtCSR.Text) <> "" And Trim(frmQuoteRpt.txtCSR.Text) <> "ALL" Then
                If InStr("," & Trim(frmQuoteRpt.txtCSR.Text) & ",", "," & Trim(CSR) & ",") Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If

            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And Trim(frmQuoteRpt.txtSlsSplit.Text) <> "" And Trim(frmQuoteRpt.txtSlsSplit.Text) <> "ALL" Then '02-22-13
                If IsDBNull(drQToRow("SLSCode")) Then drQToRow("SLSCode") = "" '02-22-13
                Dim tmpSLS As String = drQToRow.SLSCode.Trim
                'Uses all Sls In qutslssplit
                If InStr("," & Trim$(frmQuoteRpt.txtSlsSplit.Text) & ",", "," & Trim$(tmpSLS) & ",") Then Hit = 1 Else Hit = 0 : GoTo SelExit9530 '02-22-13
                '           ElseIf InStr("," & Trim$(frmQuoteRpt.txtSlsSplit.Text) & ",", "," & Trim$(drQRow.SlsMan2) & ",") Then
            End If
SkipSLSCode:
            'Else
            If Trim(frmQuoteRpt.txtCSRofCust.Text) <> "" And Trim(frmQuoteRpt.txtCSRofCust.Text) <> "ALL" Then
                If InStr("," & Trim(frmQuoteRpt.txtCSRofCust.Text) & ",", "," & Trim(CSR) & ",") Then Hit = 1 Else Hit = 0 : GoTo SelExit9530
            End If
            'Quote Lines Selection'''''''''''''''''''''''''''''''''''''''
            Dim PaidHit As Int16 = 0 '06-19-10
            Dim UnPaidHit As Int16 = 0 '06-19-10
            Dim MfgHit As Int16 = 0
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And Hit = 1 Then '12-01-09 
                '12-01-09 Select Mfg's here
                Dim drQutLn As dsSaw8.quotelinesRow '11-24-09
                For Each drQutLn In dsQuote.quotelines
                    If drQutLn.RowState = DataRowState.Deleted Then Continue For '03-01-12 Added Line
                    '12-01-09 Exclude Some Lines
                    'Include All Lines on Job
                    'Include Only Paid Items on the Job
                    'Include Only UnPaid Items on the Job
                    'Debug.Print(frmQuoteRpt.cboLinesInclude.Text)
                    'If DebugOn ThenDebug.Print(frmQuoteRpt.cboLinesInclude.Text & drQutLn.QuoteID.ToString)
                    'Dim PaidHit As Int16 = 0 '06-19-10  Dim UnPaidHit As Int16 = 0 '06-19-10
                    If frmQuoteRpt.cboLinesInclude.Text = "Include All Lines on Job" Then
                    Else
                        If frmQuoteRpt.cboLinesInclude.Text = "Include Only Paid Items on the Job" Then
                            If drQutLn.Paid = True Then PaidHit = 1 '06-19-10Hit = 0 : Exit For '12-01-09
                        End If
                        If frmQuoteRpt.cboLinesInclude.Text = "Include Only UnPaid Items on the Job" Then
                            If drQutLn.Paid = False Then UnPaidHit = 1 '06-19-10  Hit = 0 : Exit For '12-01-09
                        End If
                    End If

MoreLineTests:      If frmQuoteRpt.txtJobNameSS.Text <> "ALL" Then
                        If InStr("," & Trim(frmQuoteRpt.txtJobNameSS.Text) & ",", "," & Trim(drQutLn.MFG) & ",") Then '
                            MfgHit = 1
                        End If
                    End If
                Next 'Row
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then '06-20-10
                    If frmQuoteRpt.cboLinesInclude.Text = "Include Only Paid Items on the Job" Then
                        If PaidHit = 1 Then  Else Hit = 0 '06-19-10 : Exit For '12-01-09
                    End If
                    If frmQuoteRpt.cboLinesInclude.Text = "Include Only UnPaid Items on the Job" Then
                        If UnPaidHit = 1 Then  Else Hit = 0 '06-19-10: Exit For '06-19-20
                    End If
                End If '  If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report
                If frmQuoteRpt.txtJobNameSS.Text <> "ALL" Then
                    If MfgHit = 1 Then Hit = 1 Else Hit = 0 '12-01-09 
                End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End If
            '********************* END SELECT CRITERIA *********************

SelExit9530:  '02-21-02
9530:
            'Keep Duplicates out of reports
            If frmQuoteRpt.chkExcludeDuplicates.CheckState = CheckState.Checked Then
                If Trim(Status) = "NOREPT" Then Hit = 0 '08-11-98
            End If
        Catch ex As Exception  'Try Catch with  so you can fix exception after the error message **Get gid of  'CatchStop b/4 releasing
            MessageBox.Show("Error in SelectHit9500 (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox(ex.Message.ToString & vbCrLf & "SelectHit9500(VQRT)" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop  'Debug.WriteLine(ex.Message.ToString)
        End Try '
        '#End
    End Sub
    Public Sub SetupSpec(ByVal LoadSave As String, ByVal forGrid As Boolean, Optional ByVal tg As C1.Win.C1TrueDBGrid.C1TrueDBGrid = Nothing) '11-20-08
        '12-17-09 JH
        Dim I As Integer

600:    Try
            Dim SelectCodeArray(50, 2) As String '07-29-09 jh
            Dim FileName As String = UserSysDir & "VSPEC.DAT"

            If LoadSave = "SAVE" Then GoTo SaveDefaults

            'Find out if the File Exist - if not, Create some Default Selet Codes
            If My.Computer.FileSystem.FileExists(FileName) Then
            Else
                FileClose(3) : FileOpen(3, FileName, OpenMode.Binary, OpenShare.Shared) '06-17-15 JTC add , OpenShare.Shared
                SelectCodeArray(0, 0) = "S" : SelectCodeArray(0, 1) = "Spec"
                SelectCodeArray(1, 0) = "C" : SelectCodeArray(1, 1) = "Cross"
                SelectCodeArray(2, 0) = "E" : SelectCodeArray(2, 1) = "Equal" '09-10-10 
                For I = 0 To 50
                    If Trim(SelectCodeArray(I, 0)) = "" Then Exit For
                    SCRecord.Header = SelectCodeArray(I, 0)
                    SCRecord.TypeName = SelectCodeArray(I, 1)
                    FilePut(3, SCRecord)
                Next
                FileClose(3)
            End If

            'Create the Dataset to Store Select Codes in
            Dim myView As DataView
            dsSCross = New DataSet : Dim dt As New DataTable
            dt.Columns.Add("Code", GetType(String)) : dt.Columns.Add("Description", GetType(String))
            dsSCross.Tables.Add(dt)
650:
GetRead:
            'Read the File Contents into the Dataset
            FileClose(3) : FileOpen(3, FileName, OpenMode.Binary, OpenShare.Shared) '06-17-15 JTC add , OpenShare.Shared
            For I = 0 To 50
                If EOF(3) Then Exit For
                FileGet(3, SCRecord)
                SelectCodeArray(I, 0) = SCRecord.Header : SelectCodeArray(I, 1) = SCRecord.TypeName
                dsSCross.Tables(0).Rows.Add(New Object() {Trim(SCRecord.Header), Trim(SCRecord.TypeName)})
            Next
            FileClose(3)
            'Create the Add New... option.
            If forGrid = False And SecurityAdministrator = True Then dsSCross.Tables(0).Rows.Add(New Object() {"", "Add New..."})
675:
GetExit:
            'Set up either the Form's Combobox or the Dialog's TrueGrid
            If forGrid = False Then
                'frmQuote._fdSpecCross.DataBindings.Clear()
                'frmQuote._fdSpecCross.DataSource = dsSCross.Tables(0)
                'frmQuote._fdSpecCross.ColumnHeaders = False
                'frmQuote._fdSpecCross.Splits(0).DisplayColumns(0).Width = 30
                'frmQuote._fdSpecCross.Splits(0).DisplayColumns(1).Width = 75
                'frmQuote._fdSpecCross.ValueMember = "Code" : frmQuote._fdSpecCross.DisplayMember = "Code"
            Else
                tg.DataSource = dsSCross.Tables(0)
                tg.Columns(0).DataWidth = 1
                tg.Splits(0).DisplayColumns(0).Width = 40
                tg.Splits(0).DisplayColumns(1).Width = 230
            End If
            GoTo Exit_Done


SaveDefaults:
            'Not a global dataset, so read the Grid Contents to put it back in a dataset to sort before saving
            dsSCross = New DataSet : dt = New DataTable
            dt.Columns.Add("Code", GetType(String)) : dt.Columns.Add("Description", GetType(String)) : dsSCross.Tables.Add(dt)
            For I = 0 To tg.RowCount - 1 : dsSCross.Tables(0).Rows.Add(New Object() {tg(I, 0).ToString.ToUpper, tg(I, 1).ToString}) : Next  '12-11-12
            myView = dsSCross.Tables(0).DefaultView : myView.Sort = "Code"

            'Set the grid to the Sorted dataset
            tg.DataSource = dsSCross.Tables(0)

            'Resave the file
            Kill(FileName) : FileClose(3) : FileOpen(3, FileName, OpenMode.Binary, OpenShare.Shared) '06-17-15 JTC add , OpenShare.Shared)
            For I = 0 To tg.RowCount - 1
                SCRecord.Header = tg(I).Item("Code").ToString.ToUpper   '12-11-12
                SCRecord.TypeName = tg(I).Item("Description").ToString  '12-11-12
                FilePut(3, SCRecord)
            Next

        Catch ex As Exception
            MessageBox.Show("Error in SetUpSpec" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            FileClose(3)
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
        End Try

Exit_Done:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
    End Sub
    'Vname Code *****************************************************
    Public Function BuildSQLDetail(ByVal CC As String, ByVal FillMainGrid As Boolean, Optional ByVal strsql As String = "") As String
        'Jaci Sql Select  Q.QuoteCode, QL.* from QUOTE Q LEFT  JOIN QUOTELINES QL ON Q.QUOTEID = QL.QUOTEID  
        'This goes to QuoteLines '08-31-09 &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        '08-31-09
        'If Trim(frmQuoteRpt.TxtSingleCatNum.Text) <> "" Then
        '    If strsql <> "" Then strsql += " AND "
        '    strsql += CC & "Description LIKE " & "'" & SafeSQL(frmQuoteRpt.TxtSingleCatNum.Text.Trim) & "'"
        'End If
        'quoteLines
        'If Trim(frmQuoteRpt.txtMfgLine.Text) <> "" Then
        '    If strsql <> "" Then strsql += " AND "
        '    strsql += CC & "MFG = " & "'" & SafeSQL(Trim(frmQuoteRpt.txtMfgLine.Text)) & "'"
        'End If

        'If Trim(frmQuoteRpt.TxtSingleCatNum.Text) <> "" Then
        '    If ColText = frmQuoteRpt.tg.Splits(0).DisplayColumns(I).DataColumn.Text.ToString Then  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
        '        'If strSql <> "" Then strSql += " AND "
        '        'strSql += CC & "Description LIKE " & "'" & SafeSQL(frmQuoteRpt.TxtSingleCatNum.Text.Trim) & "'"
        '    End If


        'This goes to QuoteLines '08-31-09 &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Dim sStartDate As String = VB6.Format(frmQuoteRpt.DTPickerStartEntry.Text, "yyyy-MM-dd") '  VB6.Format("01/01/2008", "yyyy-MM-dd") 'Testing
        Dim sEndDate As String = VB6.Format(frmQuoteRpt.DTPicker1EndEntry.Text, "yyyy-MM-dd")
        'Dim sEndDate As String = VB6.Format("12/31/2009", "yyyy-MM-dd")
        If frmQuoteRpt.pnlTypeOfRpt.Text.StartsWith("Product Sales History - Line Items") Then
            strsql += "where Q.EntryDate >= '" & sStartDate & "' and Q.EntryDate <= '" & sEndDate & "' and Q.TypeOfJob = 'Q' " '05-08-10
        Else
            strsql += "where Q.EntryDate >= '" & sStartDate & "' and Q.EntryDate <= '" & sEndDate & "' " '09-03-09
        End If
        SelectionText += " Start = " & sStartDate & " End = " & sEndDate '11-20-14 JTC add SelectionText to Product Lines Report
        If frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" Then strsql += " and Q.TypeOfJob = 'S' " '09-21-12  If DefTypeOfJob = "Spec Credit" Then JT = "S"
        'Old Public Sub PrtQutLinFilterRecords(ByRef RecordMatch As Integer, ByRef AbortPrint As Integer)
        'Debug.Print(frmQuoteRpt.DTPicker1StartBid.Text & frmQuoteRpt.DTPicker1EndBid.Text)
        'If Trim(BidDate) <> "" Or BidDate <> "1-1-1900" Then
        '    If Trim(frmQuoteRpt.DTPicker1StartBid.Text) <> "" And Trim(frmQuoteRpt.DTPicker1StartBid.Text) <> "ALL" And Trim(frmQuoteRpt.DTPicker1EndBid.Text) <> "" And Trim(frmQuoteRpt.DTPicker1EndBid.Text) <> "ALL" Then
        '        If VB6.Format(BidDate, "yyyy-MM-dd") >= VB6.Format(frmQuoteRpt.DTPicker1StartBid.Text, "yyyy-MM-dd") And VB6.Format(BidDate, "yyyy-MM-dd") <= VB6.Format(frmQuoteRpt.DTPicker1EndBid.Text, "yyyy-MM-dd") Then Hit = 1 : GoTo 9510 Else Hit = 0 : GoTo SelExit9530
        '    End If
        'End If
        If frmQuoteRpt.chkUseSpecifierCode.Checked = True Or frmQuoteRpt.chkShowCustomers.Checked = True Then '08-02-13 
            strsql += " and NCode <> '' " '02-24-134 Chg (PC.NCode <> to Ncode 07-15-13 No specifier"
        End If
        'Add txtPrcCode.Text.Trim
        'Eliminate if No MFG Code ????
        Dim rcode As String = ""
        rcode = frmQuoteRpt.txtStat.Text.Trim '09-02-09 Status Code
        '08-31-09 Quote Selection
        If rcode <> "" And rcode <> "ALL" Then '09-02-09
            'SelectionText += " Status = " & rcode '11-20-14 JTC add SelectionText to Product Lines Report
            'If strsql <> "" Then strsql += " AND "
            'strsql += CC & "Status = " & "'" & SafeSQL(rcode) & "'"
            '05-28-15 JTC Fix Multiple Status Codes("OPEN,GOT")  in Select Line Items Not Forecasting Replace above If rcode.Contains(",") = True Then 
            Dim BC As String = ""
            If rcode.Contains(",") = True Then  '10-19-13 JTC Added QL.QL.MFG to Line Items
                BC = " and ( Q.Status = '" & rcode.Replace(",", "' or Q.Status = '") & "' ) " '10-20-13 No Blanks or QL.MFG = ''  )"
            Else
                BC = " and ( Q.Status = '" & rcode & "'" & " ) " '10-20-13 No Blanks or QL.MFG = '')"
            End If '" and ( QL.MFG = 'DAYB' or QL.MFG = 'CAPR' or QL.MFG = 'MCPH' or QL.MFG = 'OMEG' or QL.MFG = 'LAM' or QL.MFG = 'CHLO' or QL.MFG = 'MORL' or QL.MFG = 'GUTH' or QL.MFG = 'ARDE' or QL.MFG = 'CRES' or QL.MFG = 'GARD' or QL.MFG = 'EMCO' or QL.MFG = 'FORE' or QL.MFG = 'HADC' or QL.MFG = 'EXCE' or QL.MFG = 'TRAN' or QL.MFG = 'LOL' or QL.MFG = 'ALKC' or QL.MFG = 'LEDA' or QL.MFG = 'LUME' or QL.MFG = 'CRES' or QL.MFG = 'ALLS' or QL.MFG = 'HANO' or QL.MFG = 'THMO' or QL.MFG = 'THMI' or QL.MFG = ''  )"
            'If strsql <> "" Then strsql += " AND "
            strsql += BC
        End If
        rcode = frmQuoteRpt.txtRetr.Text.Trim 'Retrieval Code
        If rcode <> "" And rcode <> "ALL" Then '09-02-09
            SelectionText += " RetrCode = " & rcode '11-20-14 JTC add SelectionText to Product Lines Report
            If strsql <> "" Then strsql += " AND "
            strsql += CC & "RetrCode = " & "'" & SafeSQL(rcode) & "'"
        End If
        rcode = Trim(frmQuoteRpt.txtSpecCross.Text) 'SpecCross field
        If rcode <> "" And rcode <> "ALL" Then '09-02-09
            SelectionText += " SpecCross = " & rcode '11-20-14 JTC add SelectionText to Product Lines Report
            If strsql <> "" Then strsql += " AND "
            strsql += CC & "SpecCross = " & "'" & SafeSQL(rcode) & "'"
        End If
        '11-20-14 JTC Add SLSQ Quoter to Line Item Report
        rcode = frmQuoteRpt.txtSlsTerr.Text.Trim
        If rcode <> "" And rcode <> "ALL" Then '11-20-14 JTC Add SLSQ Quoter to Line Item Report
            SelectionText += " SLSQ = " & rcode '11-20-14 JTC add SelectionText to Product Lines Report
            If strsql <> "" Then strsql += " AND "
            strsql += CC & "SLSQ = " & "'" & SafeSQL(rcode) & "'"
        End If
        rcode = frmQuoteRpt.txtSalesman.Text.Trim
        If rcode <> "" And rcode <> "ALL" Then '09-02-09
            SelectionText += " SLSQ = " & rcode '11-20-14 JTC add SelectionText to Product Lines Report
            If strsql <> "" Then strsql += " AND "
            strsql += CC & "SLSQ = " & "'" & SafeSQL(rcode) & "'"
        End If
        '09-03-09 Price Code  Select QL. Line Items *************************
        rcode = frmQuoteRpt.txtPrcCode.Text.Trim
        If rcode <> "" And rcode <> "ALL" Then '09-02-09
            If strsql <> "" Then strsql += " AND "
            strsql += "QL." & "PriceCode = " & "'" & SafeSQL(rcode) & "'"
        End If
        rcode = frmQuoteRpt.txtMfgLine.Text.Trim
        'If rcode <> "" And rcode <> "ALL" Then '09-03-09
        '    SelectionText += " MFG = " & rcode '11-20-14 JTC add SelectionText to Product Lines Report
        '    If strsql <> "" Then strsql += " AND "
        '    strsql += "QL." & "MFG = " & "'" & SafeSQL(rcode) & "'"
        'End If
        '05-28-15 JTC Fix Multiple MFgs in Select Line Items Not Forecasting Replace above BuildSQLDetail
        If rcode <> "" And rcode <> "ALL" Then '09-03-09
            'SelectionText += " MFG = " & rcode '11-20-14 JTC add SelectionText to Product Lines Report
            'If strsql <> "" Then strsql += " AND "
            'strsql += "QL." & "MFG = " & "'" & SafeSQL(rcode) & "'"
            Dim BC As String = ""
            If rcode.Contains(",") = True Then  '10-19-13 JTC Added QL.QL.MFG to Line Items
                BC = " and ( QL.MFG = '" & rcode.Replace(",", "' or QL.MFG = '") & "' ) " '10-20-13 No Blanks or QL.MFG = ''  )"
            Else
                BC = " and ( QL.MFG = '" & rcode & "'" & " ) " '10-20-13 No Blanks or QL.MFG = '')"
            End If '" and ( QL.MFG = 'DAYB' or QL.MFG = 'CAPR' or QL.MFG = 'MCPH' or QL.MFG = 'OMEG' or QL.MFG = 'LAM' or QL.MFG = 'CHLO' or QL.MFG = 'MORL' or QL.MFG = 'GUTH' or QL.MFG = 'ARDE' or QL.MFG = 'CRES' or QL.MFG = 'GARD' or QL.MFG = 'EMCO' or QL.MFG = 'FORE' or QL.MFG = 'HADC' or QL.MFG = 'EXCE' or QL.MFG = 'TRAN' or QL.MFG = 'LOL' or QL.MFG = 'ALKC' or QL.MFG = 'LEDA' or QL.MFG = 'LUME' or QL.MFG = 'CRES' or QL.MFG = 'ALLS' or QL.MFG = 'HANO' or QL.MFG = 'THMO' or QL.MFG = 'THMI' or QL.MFG = ''  )"
            'If strsql <> "" Then strsql += " AND "
            strsql += BC
        End If         '05-28-15 JTC Multiple MFgs in Select Line Items Not Forecasting
        'Debug.Print(frmQuoteRpt.TxtSingleCatNum.Text)
        rcode = frmQuoteRpt.txtCustomerCodeLine.Text.Trim
        '05-28-15 JTC Customer Option Must Check Specifier or Customer In Line Items Reports Multiple Customer Codes("GES/AT,GES/BR")  in Select Line Items Not Forecasting BuildSQLDetail If rcode.Contains(",") = True Then 
        If frmQuoteRpt.chkUseSpecifierCode.Checked = True Or frmQuoteRpt.chkShowCustomers.Checked = True Then '08-02-13 
            strsql += " and NCode <> '' " '02-24-134 Chg (PC.NCode <> to Ncode 07-15-13 No specifier"
            If rcode <> "" And rcode <> "ALL" Then '08-02-13
                'SelectionText += " NCode = " & rcode '11-20-14 JTC add SelectionText to Product Lines Report
                'If strsql <> "" Then strsql += " AND "
                'strsql += "NCode = " & "'" & SafeSQL(rcode) & "'" '02-25-14 Chg "PC." & "NCode = " to Ncode
                '05-28-15 JTC Fix Multiple Customer Codes("GES/AT,GES/BR") If frmQuoteRpt.chkUseSpecifierCode.Checked = True Or frmQuoteRpt.chkShowCustomers.Checked = True  in Select Line Items Not Forecasting Replace above If rcode.Contains(",") = True Then 
                Dim BC As String = ""
                If rcode.Contains(",") = True Then  '10-19-13 JTC Added QL.QL.MFG to Line Items
                    BC = " and ( PC.NCode = '" & rcode.Replace(",", "' or PC.NCode = '") & "' ) " '10-20-13 No Blanks or QL.MFG = ''  )"
                Else
                    BC = " and ( PC.NCode = '" & rcode & "'" & " ) " '10-20-13 No Blanks or QL.MFG = '')"
                End If '" and ( QL.MFG = 'DAYB' or QL.MFG = 'CAPR' or QL.MFG = 'MCPH' or QL.MFG = 'OMEG' or QL.MFG = 'LAM' or QL.MFG = 'CHLO' or QL.MFG = 'MORL' or QL.MFG = 'GUTH' or QL.MFG = 'ARDE' or QL.MFG = 'CRES' or QL.MFG = 'GARD' or QL.MFG = 'EMCO' or QL.MFG = 'FORE' or QL.MFG = 'HADC' or QL.MFG = 'EXCE' or QL.MFG = 'TRAN' or QL.MFG = 'LOL' or QL.MFG = 'ALKC' or QL.MFG = 'LEDA' or QL.MFG = 'LUME' or QL.MFG = 'CRES' or QL.MFG = 'ALLS' or QL.MFG = 'HANO' or QL.MFG = 'THMO' or QL.MFG = 'THMI' or QL.MFG = ''  )"
                'If strsql <> "" Then strsql += " AND "
                strsql += BC
            End If
        End If
        If (frmQuoteRpt.pnlTypeOfRpt.Text = "Terr Spec Credit Report" And (frmQuoteRpt.txtPrimarySortSeq.Text = "MFG Follow-Up Report" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman Follow-Up Report")) Then '03-19-14 was above
            rcode = frmQuoteRpt.txtJobNameSS.Text.Trim '03-19-14
            If rcode <> "" And rcode <> "ALL" Then
                If strsql <> "" Then strsql += " AND "
                strsql += "QL." & "MFG = " & "'" & SafeSQL(rcode) & "'"
            End If
        End If
        If Trim(frmQuoteRpt.TxtSingleCatNum.Text) <> "" And frmQuoteRpt.TxtSingleCatNum.Text.ToUpper <> "ALL" Then '03-05-13
            If strsql <> "" Then strsql += " AND "
            strsql += "QL." & "Description LIKE " & "'" & SafeSQL(frmQuoteRpt.TxtSingleCatNum.Text.Trim) & "'"
        End If

        If frmQuoteRpt.TxtSearchString.Text.Trim <> "" Then '03-19-14
            If strsql <> "" Then strsql += " AND "
            strsql += "QL." & "Description LIKE " & "'%" & SafeSQL(frmQuoteRpt.TxtSearchString.Text.Trim) & "%'" '06-23-15 JTC fic TxtSearchString
        End If

        Return strsql
        Exit Function
        'End Vname Code ***************************************************
        'OldCode'Header Item Check''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '        If QI2.HeadItem.Value = "H" Then GoTo 6900
        '        ''Non-Catalog Number Line Check''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '        If Conversion.Val(QI2.Qty.Value) = 9996 Or Conversion.Val(QI2.Qty.Value) = 9997 Or Conversion.Val(QI2.Qty.Value) = 9998 Or Conversion.Val(QI2.Qty.Value) = 9999 Then
        '            GoTo 6900
        '        End If
        '        If DIST Then
        '            If QN = 1 Then
        '                If QI2.MFG.Value.Trim() = "" Then GoTo 6900
        '            Else
        '                If QI2.LpMfg.Value.Trim() = "" Then GoTo 6900
        '            End If
        '        Else
        '            If Trim$(QI2.MFG) = "" Then GoTo 6900
        '        End If
        '        '''Date Check''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '        If RptBy <> "MONT" Then
        '            TEDT = QI2.Date_Renamed.Value.Substring(0, Math.Min(QI2.Date_Renamed.Value.Length, 8))
        '            If TEDT > EndD Or TEDT < StartD Or TEDT > EndD Then GoTo 6900
        '        Else
        '            If StartD.Substring(0, Math.Min(StartD.Length, 4)) <> QI2.Date_Renamed.Value.Substring(0, Math.Min(QI2.Date_Renamed.Value.Length, 4)) Then GoTo 6900
        '        End If
        '        '''Bid Date Check''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '        If BidEndD <> "A" And BidStartD <> "A" Then
        '            QK = QI2.Code.Value : SaveFQ = FQ : SaveQN = QN
        '            FQ = 3 : QN = 0 : CheckMaster = 1
        '            DbaseFunction = "Db_GetEq" : DBExecute(DbaseFunction, FQ, QN, QK)
        '            FQ = SaveFQ : QN = SaveQN
        '            TEDT = QM.BidDate.Value
        '            If TEDT > BidEndD Or TEDT < BidStartD Or TEDT > BidEndD Then GoTo 6900
        '        End If
        '        '''Mfg Check'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '        If SelMfg <> "A" Then ' And SelMfg$ <> "M" Then
        '            If DIST Then
        '                If QN = 1 Then
        '                    If SelMfg <> QI2.MFG.Value Then
        '                        If QN = 1 Then
        '                            AbortPrint = 1 : GoTo 6900
        '                        Else
        '                            GoTo 6900
        '                        End If
        '                    End If
        '                Else
        '                    If SelMfg <> QI2.LpMfg.Value Then
        '                        If QN = 1 Then
        '                            AbortPrint = 1 : GoTo 6900
        '                        Else
        '                            GoTo 6900
        '                        End If
        '                    End If
        '                End If
        '            Else
        '                If Trim$(SelMfg$) <> Trim$(QI2.MFG) Then
        '                    If QN% = 1 Then AbortPrint% = 1 : GoTo 6900 Else GoTo 6900
        '                End If
        '            End If
        '        End If
        '        '''Salesman Check''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '        If SelSlsMan <> "A" Then
        '            If CheckMaster = 0 Then
        '                If DIST Then
        '                    QK = QI2.Code.Value
        '                Else
        '                    QK$ = QI2.Code & Chr$(0) & Chr$(0)
        '                End If
        '                SaveFQ = FQ : SaveQN = QN
        '                FQ = 3 : QN = 0
        '                DbaseFunction = "Db_GetEq" : DBExecute(DbaseFunction, FQ, QN, QK)
        '                FQ = SaveFQ : QN = SaveQN : CheckMaster = 1
        '            End If
        '            If DIST Then
        '                If QM.Sls.Value <> SelSlsMan.ToUpper() Then
        '                Else
        '                    If QM.QuoteSls <> UCase$(SelSlsMan$) Then
        '                    End If
        '                    GoTo 6900
        '                End If
        '            End If
        '            '''Retrieval Check'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '            If SelRet <> "A" Then
        '                If CheckMaster = 0 Then
        '                    If DIST Then
        '                        QK = QI2.Code.Value
        '                    Else
        '                        QK$ = QI2.Code & Chr$(0) & Chr$(0)
        '                    End If
        '                    SaveFQ = FQ : SaveQN = QN
        '                    FQ = 3 : QN = 0
        '                    DbaseFunction = "Db_GetEq" : DBExecute(DbaseFunction, FQ, QN, QK)
        '                    FQ = SaveFQ : QN = SaveQN : CheckMaster = 1
        '                End If
        '                If QM.Retriev.Value <> SelRet.ToUpper() Then GoTo 6900
        '            End If
        '            '''Status Check''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '            If SelStatus <> "A" Then
        '                If CheckMaster = 0 Then
        '                    If DIST Then
        '                        QK = QI2.Code.Value
        '                    Else
        '                        QK$ = QI2.Code & Chr$(0) & Chr$(0)
        '                    End If
        '                    SaveFQ = FQ : SaveQN = QN
        '                    FQ = 3 : QN = 0
        '                    DbaseFunction = "Db_GetEq" : DBExecute(DbaseFunction, FQ, QN, QK)
        '                    FQ = SaveFQ : QN = SaveQN : CheckMaster = 1
        '                End If
        '                If QM.Status.Value <> SelStatus.ToUpper() Then GoTo 6900
        '            End If
        '            '''Select individual catalog number
        '            If SelCatNum <> "A" Then
        '                If DIST Then
        '                    If QN = 1 Then
        '                        If SelCatNum.Trim() <> QI2.Desc.Value.Trim() Then GoTo 6900
        '                    Else
        '                        If SelCatNum.Trim() <> QI2.LpDesc.Value.Trim() Then GoTo 6900
        '                    End If
        '                Else
        '                    If Trim$(SelCatNum$) <> Trim$(QI2.Desc) Then GoTo 6900
        '                End If
        '            End If
        '            '''search for string (ie, esb)
        '            If SelString <> "" Then
        '                If DIST Then
        '                    If QN = 1 Then
        '                        StrMatch = 0 : StrMatch = (QI2.Desc.Value.ToUpper().IndexOf(SelString) + 1) '08-08-00 WNA
        '                        If StrMatch = 0 Then GoTo 6900
        '                    Else
        '                        StrMatch = 0 : StrMatch = (QI2.LpDesc.Value.ToUpper().IndexOf(SelString) + 1) '08-08-00 WNA
        '                        If StrMatch = 0 Then GoTo 6900
        '                    End If
        '                Else
        '                    StrMatch% = 0 : StrMatch% = InStr(UCase$(QI2.Desc), SelString$)  '08-08-00 WNA
        '                    If StrMatch% = 0 Then GoTo 6900
        '                End If
        '            End If
        '            If frmMenu.txtPrcCode.Text.Trim() <> "ALL" And frmMenu.txtPrcCode.Text.Trim() <> "" Then
        '                If QI2.PrcCode.Value.Trim() <> frmMenu.txtPrcCode.Text.Trim() Then GoTo 6900
        '            End If
        '            If DIST <> 1 Then
        '                If Trim$(frmMenu.txtLastChgBy.Text) <> "ALL" And Trim$(frmMenu.txtLastChgBy.Text) <> "" Then
        '                    If Trim$(QI2.LastChgBy) <> Trim$(frmMenu.txtLastChgBy.Text) Then GoTo 6900
        '                End If
        '                If Trim$(frmMenu.txtSpecCross.Text) <> "ALL" And Trim$(frmMenu.txtSpecCross.Text) <> "" Then
        '                    If Trim$(QI2.SpecCrossed) <> Trim$(frmMenu.txtSpecCross.Text) Then GoTo 6900
        '                End If
        '            End If
        '            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '            CheckForInfo = 1 'flag to see if any records in the search match (no rec meet criteria)
        '            RecordMatch = 1 'flag to see if this record meets criteria, 1 = yes, 0 = no(skip to next record)
        '            Exit Sub
        '            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '6899:   Resume 6900 ''on error skip this record
        '6900:   RecordMatch = 0 ''this means to skip this record and get next, it has not met all of
        '        '            ''the above criteria
    End Function


    Public Sub PrtQutLinParseDesc()
        Dim TildaCheck As Integer
        Dim DescString As String = ""

        On Error Resume Next
        ''Logic to parse string when finding '~', which is an indication the catalog number
        ''has ended, and a comment is to follow
        'DescString = QI2.Desc.Value
        'TildaCheck = (DescString.IndexOf("~"c) + 1)
        'If TildaCheck <> 0 Then
        '    QI2.Desc.Value = (QI2.Desc.Value.Substring(0, Math.Min(QI2.Desc.Value.Length, TildaCheck - 1)) & Wspcs).Substring(0, Math.Min((QI2.Desc.Value.Substring(0, Math.Min(QI2.Desc.Value.Length, TildaCheck - 1)) & Wspcs).Length, 45))
        'End If


    End Sub




    Public Sub PrtMajTot9390(ByRef A As String, ByVal Lev As TotalLevels, ByRef RT As C1.C1Preview.RenderTable, ByVal doc As C1PrintDocument)
        On Error GoTo 99997
        'Dim A As String = PrimarySortSeq
        Dim TempFds As String
        Dim J As Short
        Dim SLTCommPct As Decimal
        Dim SLTCommAmt As Decimal 'Call PrtMajTot9390(A$)
PrtMajTot9390:  '02-21-02
9390:
        THDG = Left("**TOTAL " & A & Wspcs, 20) : Call TotPrt9250(THDG, Lev, RT, doc)
        GoTo 998 'Exit
        '      If MARK Then 'MARK
        '         If Left$(UH$, 6) = "STATUS" Then
        '            TOTD# = SLT# + TOTD#: SLT# = TOTD#:
        '            THDG$ = Left$("**TOTAL " + PREVSLS$ + Wspcs$, 20): Call TotPrt9250(A$):
        '            If DIST Then
        '               PREVSLS$ = QM.Sls:
        '            Else
        '               PREVSLS$ = QM.QuoteSls:
        '            End If
        '            SLT# = 0: TOTD# = 0
        '         End If
        '      Else    ' Not MARK
        If Left(UH, 8) = "SALESMAN" Or Left(UH, 6) = "STATUS" Or Left(UH, 6) = "FOLLOWEDBY" Or Left(UH, 6) = "ENTEREDBY" Then '' RptFollowBy = 16 'FollowedBy 03-01-12  SubSProj = 1 'JobName
            TOTD = SLT + TOTD : SLT = TOTD
            THDG = Left("**TOTAL " & A & Wspcs, 20) : Call TotPrt9250(THDG, Lev, RT, doc)
            SLT = 0 : TOTD = 0
            TOTDCommAmt = 0 : TOTDCommPct = 0 '02-26-01 WNA
            SLTCommAmt = 0 : SLTCommPct = 0 '05-07-02 WNA
        End If
        If Left(UH, 20) = "SPECIFIER CREDIT SEQ" Then '02-27-01 WNA
            If (VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Or VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Or VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif) And Trim(PrevStat) <> "" Then '2 or 3  '02-22-02
                THDG = Left("**TOTAL " & PrevStat & Wspcs, 20) : Call TotPrt9250(A, TotalLevels.TotGt, RT, doc)
                TOTD = SLT + TOTD : SLT = TOTD
                TOTD = 0
            End If
            If PREVSLS <> Left(Fds(J - 1), 6) Then
                If VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Or VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Or VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif Then
                    TOTD = SLT ' 2 or 3  '02-22-02
                    TOTDCommAmt = SLTCommAmt '05-07-02 WNA
                    TOTDCommPct = SLTCommPct '05-07-02 WNA
                End If
                THDG = Left("****TOTAL " & PREVSLS & Wspcs, 20) : Call TotPrt9250(A, TotalLevels.TotGt, RT, doc) '02-22-02 Added **
                frmQuoteRpt.Text = vbCrLf
                Call AbortCheck52(A) : If A = "Cancel" Then GoTo 998 '01-13-09
                SLT = 0 : TOTD = 0
                TOTDCommAmt = 0 : TOTDCommPct = 0 '02-26-01 WNA
                SLTCommAmt = 0 : SLTCommPct = 0 '05-07-02 WNA
            End If
        End If
        If Left(UH, 18) = "RETRIEVAL CODE SEQ" Or Left(UH, 18) = "MARKET SEGMENT SEQ" Then '07-30-04 JH
            If (VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Or VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Or VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif) And Trim(PrevStat) <> "" Then '2 or 3  '02-22-02
                THDG = Left("**TOTAL " & PrevStat & Wspcs, 20) : Call TotPrt9250(A, TotalLevels.TotGt, RT, doc)
                TOTD = SLT + TOTD : SLT = TOTD
                TOTD = 0
            End If
            If PREVSLS <> TempFds And Trim(PREVSLS) <> "" Then
                If VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Or VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Or VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif Then
                    TOTD = SLT ' 2 or 3  '02-22-02
                    TOTDCommPct = SLTCommPct
                    TOTDCommAmt = SLTCommAmt
                End If
                THDG = Left("****TOTAL " & PREVSLS & Wspcs, 20) : Call TotPrt9250(THDG, TotalLevels.TotGt, RT, doc) '02-22-02 Added **
                frmQuoteRpt.Text = vbCrLf
                Call AbortCheck52(A) : If A = "Cancel" Then GoTo 998 '01-13-09
                SLT = 0 : TOTD = 0
                TOTDCommAmt = 0 : TOTDCommPct = 0 '02-26-01 WNA
                SLTCommPct = 0 : SLTCommAmt = 0 '05-07-02 WNA
            End If
        End If
        GoTo 998
99997:  'On Error GoTo 99997:
ErrBox: Dim Msg As String
        Msg = "VB Error # = " & Str(Err.Number) & "  ERROR AT " & Str(Erl()) & vbCrLf & "PLEASE READ BACK OF MANUAL" & vbCrLf & "ERROR MESSAGE SECTION Z-2"
        Dim resp As Integer
        resp = MsgBox(Msg & vbCrLf & ErrorToString(Err.Number), MsgBoxStyle.OkCancel, US)
        If resp = MsgBoxResult.Cancel Then Resume 99998 ' GoTo Msub_Exit: 'Cancel
        Resume 99998 'Return
99998:
9400:
998:
    End Sub

    Public Sub SubTotChk9360(ByRef RT As C1.C1Preview.RenderTable, ByVal doc As C1PrintDocument)
        Try  '#top
            Dim MARK As Boolean = 0
            Dim First As Short
            PREVSLS = PrevLev1
            PrevStat = PrevLev2
            Dim TmpPREVBRK As String = "" '07-02-13 
SubTotChk9360:
9300:       ' If Cmd <> "EOF" Then '10-30-12 
            '     If drOrd.RowState = DataRowState.Deleted = False Then '03-17-11
            '           If IsDBNull(drOrd("SLSBranch")) Then drOrd("SLSBranch") = "" '10-30-12
            '      End If
            ' End If

            If Cmd = "EOF" Then '03-14-12
                CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ"
                If Left(UH, 8) = "SALESMAN" Or Left(UH, 6) = "STATUS" Or Left(UH, 11) = "FOLLOWED BY" Or Left(UH, 10) = "ENTERED BY" Then '05-14-13
                    GoTo SlsStatusFoll '' RptFollowBy = 16 'FollowedBy 03-01-12  SubSProj = 1 'JobName SlsStatusFoll:'03-14-12
                Else
                    GoTo Lev2_Brk_9400 '05-07-10 Need here for Deleted Row Problem
                End If
            End If
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
                If drQToRow.RowState = DataRowState.Deleted Then GoTo Lev2_Brk_9400 '05-04-10 
                If drQToRow.RowState = DataRowState.Deleted = False Then '03-17-11
                    If IsDBNull(drQToRow("BranchCode")) Then drQToRow("BranchCode") = "" '10-30-12
                    If drQToRow("BranchCode") = "" Then drQToRow("BranchCode") = "000" '10-30-12
                End If
                If IsDBNull(drQToRow("SLSCode")) Then drQToRow("SLSCode") = "" '01-30-14 JTC Fix Realization Null on drQToRow.SLSCode SubTotChk9360
                If drQToRow.SLSCode.Trim = "" Then drQToRow.SLSCode = "000" '10-30-12 
                If PrevLev1 = "" Then
                    '05-16-13 JTC Added Realization When sub Sort is salesman they can change tobe Salesman Major Sequence
                    If frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman/Customer" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman" Then '05-16-13 Added frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman
                        PrevLev1 = drQToRow.SLSCode : PrevLev2 = drQToRow.NCode
                        CurrLev1 = drQToRow.SLSCode : CurrLev2 = drQToRow.NCode
                    Else  'No Lev2
                        PrevLev1 = drQToRow.NCode : PrevLev2 = "" 'drQTRow.NCode
                        CurrLev1 = drQToRow.NCode : CurrLev2 = "" 'drQToRow.NCode
                        If RealWithOneMfgCustSortJobName = True Then '10-13-14 JTC Job Name
                            PrevLev1 = drQToRow.JobName : PrevLev2 = ""
                            CurrLev1 = drQToRow.JobName : CurrLev2 = ""
                        End If
                    End If
                    '09-23-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
                    If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)) Then
                        SortCode = drQToRow.FirmName
                    End If ' End '09-18-15 
                    '09-23-15 JTC RealSLSCustomer only
                    If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And (RealManufacturer = False Or RealCustomerOnly = False Or RealSLSCustomer = True Or RealALL = False) Then
                        SortCode = drQToRow.FirmName
                    End If
                    If BranchReporting = True Then '10-30-12
                        CurrLev2 = CurrLev1 : PrevLev2 = PrevLev1 'Move Down a Level
                        CurrLev1 = drQToRow.BranchCode : PrevLev1 = drQToRow.BranchCode '10-30-12 JTC SubTotChk9360
                    End If
                Else  'CurrLev
                    '05-16-13 JTC Added Realization When sub Sort is salesman they can change tobe Salesman Major Sequence
                    If frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman/Customer" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman" Then '05-16-13 Added frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman 
                        If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ" : GoTo Lev2_Brk_9400 '02-14-09
                        CurrLev1 = drQToRow.SLSCode : CurrLev2 = drQToRow.NCode
                    Else    'No Lev2
                        If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ" : GoTo Lev1_Brk_9500 '02-14-09
                        CurrLev1 = drQToRow.NCode : CurrLev2 = "" : PrevLev2 = "" ' 'drQTRow.NCode
                        If RealWithOneMfgCustSortJobName = True Then '10-13-14 JTC Job Name
                            'PrevLev1 = drQToRow.JobName : PrevLev2 = ""
                            CurrLev1 = drQToRow.JobName : CurrLev2 = "" : PrevLev2 = ""
                        End If
                    End If
                    If BranchReporting = True Then '10-30-12
                        CurrLev2 = CurrLev1 : CurrLev1 = drQToRow.BranchCode '10-30-12 JTC SubTotChk9360
                    End If
                End If

                If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ"
                'Debug.Print(SortSeq) ' = "projectcust.Sell" '02-11
                '02-11-12
                'Debug.Print(frmQuoteRpt.ChkSpecifiers.Text)
                '02-01-14 frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" and VQRT2.SubSeq = VQRT2.SubSortType.SubSDescend then 
                '02-01-14 JTC out If frmQuoteRpt.ChkSpecifiers.Text = "Sort Report by Descending Dollar" And frmQuoteRpt.ChkSpecifiers.CheckState = CheckState.Checked And (SortSeq = "projectcust.Sell" Or SortSeq = "Quotecust.Sell") Then GoTo 9375 '01-28-13 "Quotecust.Sell" 02-11-12'02-11-12 " "Add Specifiers (Arch, Eng, Etc) to Reports" '02-11-12 
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And RealCustomer = True And RealManufacturer = False And VQRT2.RepType = VQRT2.RptMajorType.RptProj And VQRT2.SubSeq = VQRT2.SubSortType.SubSProj Then GoTo 9375 '03-11-14
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And frmQuoteRpt.txtPrimarySortSeq.Text = "SalesMan" And frmQuoteRpt.cboSortSecondarySeq.Text = "Descending Dollar" And (SortSeq = "projectcust.Sell" Or SortSeq = "Quotecust.Sell") Then GoTo 9375 '02-01-14
                'If SortSeq = "projectcust.Sell" Then GoTo 9375 '02-11-12
                GoTo Lev2_Brk_9400
            End If 'End Realization &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

            If frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" Then
                'If BranchReporting = True Then '10-30-12
                '    If Cmd = "EOF" Then CurrLev1 = "ZZZ"
                '    CurrLev2 = "" : CurrLev1 = drQRow.BranchCode '10-30-12 JTC SubTotChk9360
                'End If
                If drQRow.RowState = DataRowState.Deleted Then GoTo Lev2_Brk_9400 '05-04-10 
            End If
            'Stk/Job ??
            'If MajSel = RptMaj.RptCust And Cmd <> "EOF" Then PREVBRK = drOrd("CustName") '07-03-13 
            'If MajSel = RptMaj.RptSls14 And SSeqSel = SecSort.SSCust Then PREVBRK = drOrd("CustName") '07-03-13
            'Not Realization
            If drQRow.RowState = DataRowState.Deleted = False Then '10-30-12
                If IsDBNull(drQRow("BranchCode")) Then drQRow("BranchCode") = "" '10-30-12
                If drQRow("BranchCode") = "" Then drQRow("BranchCode") = "000" '10-30-12
            End If
            If VQRT2.RepType = VQRT2.RptMajorType.RptBidDate Then GoTo Exit9375 '12-23-09 02-10-09 OrderBy = "Q.BidDate"
            If VQRT2.RepType = VQRT2.RptMajorType.RptDescend Then GoTo Exit9375 ' : OrderBy = "Q.Sell DESC" '03-08-09
            If VQRT2.RepType = VQRT2.RptMajorType.RptEntryDate Then GoTo Exit9375 '02-10-09 OrderBy = "Q.EnterDate"
            'If VQRT2.RepType = VQRT2.RptMajorType.RptLocation Then OrderBy = "Q.Location"
            'If VQRT2.RepType = VQRT2.RptMajorType.RptMarketSegment Then OrderBy = "Q.MarketSegment"
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) And frmQuoteRpt.cboSortSecondarySeq.Text = "Salesman 1-4 Splits" Then
                CurrLev1 = Left(drQRow.FollowBy, 3)
                If PrevLev1 = "" Then PrevLev1 = CurrLev1
                GoTo SlsStatusFoll '06-02-15 JTC 
            End If
            If VQRT2.RepType = VQRT2.RptMajorType.RptProj Then GoTo Exit9375 '02-10-09 OrderBy = "P.ProjectName"
            If VQRT2.RepType = VQRT2.RptMajorType.RptQutCode And BranchReporting = False Then GoTo Exit9375 '10-30-12 JTC BranchReporting = True
            'If VQRT2.RepType = VQRT2.RptMajorType.RptRetrieval Then OrderBy = "Q.RetrCode"
            'If VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Then OrderBy = "Q.SLSQ"
            'If VQRT2.RepType = VQRT2.RptMajorType.RptSpecif Then OrderBy = "Reference"
            'If VQRT2.RepType = VQRT2.RptMajorType.RptStatus Then OrderBy = "Q.Status"
            ''VQRT2.RptMajorType.RptStatus() 'RptSpecif'RptSalesman'RptRetrieval'RptQutCode'RptProj'RptMarketSegment'RptLocation'RptEntryDate
SlsStatusFoll:  '03-14-12
            If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ" : GoTo 9365 '04-13-12 JTC Fix Deleted Row Problem

            If Left(UH, 8) = "SALESMAN" Or Left(UH, 6) = "STATUS" Or Left(UH, 11) = "FOLLOWED BY" Or Left(UH, 10) = "ENTERED BY" Then '05-14-13' RptFollowBy = 16 'FollowedBy 03-01-12  SubSProj = 1 'JobName SlsStatusFoll:'03-14-12
                If SortSeq = "quote.SLSQ" Or SortSeq = "QS.SLSCode" Then ' Primary Sort 03-08-13
                    CurrLev1 = drQRow.SLSQ
                    CurrLev2 = drQRow.Status
                    'ElseIf SortSeq = "QS.SLSCode" Then ' Primary Sort 03-08-13
                    '    CurrLev1 = drQRow.SLSQ
                    '    CurrLev2 = drQRow.Status
                ElseIf SortSeq = "quote.FollowedBy" Then '03-03-12 
                    CurrLev1 = drQRow.FollowBy
                    'CurrLev1 = drQRow.SLSQ '
                ElseIf SortSeq = "quote.EnteredBy" Then '05-14-123
                    CurrLev1 = drQRow.EnteredBy
                    CurrLev2 = drQRow.Status
                Else
                    CurrLev1 = drQRow.Status
                    CurrLev2 = drQRow.SLSQ
                End If
                If VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Then
                    CurrLev2 = drQRow.Status
                Else
                    If VQRT2.SubSortType.SubSSelectBidDate Then '03-03-12 Don't Change
                    Else
                        CurrLev2 = drQRow.SLSQ
                    End If
                End If
                If BranchReporting = True Then '10-30-12
                    CurrLev2 = CurrLev1 : CurrLev1 = drQRow.BranchCode '10-30-12 JTC SubTotChk9360
                End If
                If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ"

                GoTo 9365

            ElseIf Left(UH, 18) = "RETRIEVAL CODE SEQ" Or Left(UH, 18) = "MARKET SEGMENT SEQ" Or Left(UH, 20) = "SPECIFIER CREDIT SEQ" Then  'Then  '07-30-04 JH
                'PrevLev1 = ""
                If PrevLev1 = "" Then
                    If VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif Then PrevLev2 = CurrLev2 'Else PREVSLS = CurrLev2 ' Left(Fds(J - 1), 10)
                    If VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Then PrevLev2 = drQRow.Status ' PrevLev2 'Mid(Fds(J - 1), 11, 6) '02-22-02 Else PrevStat$ = Mid$(Fds(J% - 1), 7, 3) '3
                    If VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Then PrevLev2 = drQRow.SLSQ ' PrevLev2 Else PrevStat = PrevLev2 'Mid(Fds(J - 1), 11, 3) '3
                    If VQRT2.SubSeq = VQRT2.SubSortType.SubSProj Then PrevLev2 = drQRow.JobName '09-14-10
                    If VQRT2.SubSeq = VQRT2.SubSortType.SubSEnterDate Then PrevLev2 = drQRow.EntryDate 'SubSEnterDate'SubSProjCode'11-20-10
                    If VQRT2.SubSeq = VQRT2.SubSortType.SubSProjCode Then PrevLev2 = drQRow.QuoteCode 'SubSEnterDate'SubSProjCode'11-20-10
                    If VQRT2.SubSeq = VQRT2.SubSortType.SubSDescend Then PrevLev2 = drQRow.Sell
                    If VQRT2.SubSeq = VQRT2.SubSortType.SubSBidDate Then PrevLev2 = drQRow.BidDate.ToString '10-13-11
                End If
                If Left(UH, 18) = "RETRIEVAL CODE SEQ" Then 'If First Then ' Left$(UH$, 8) = "SALESMAN" Or Left$(UH$, 6) = "STATUS"
                    ' Primary Sort
                    If PrevLev1 = "" Then
                        PrevLev1 = drQRow.RetrCode
                        CurrLev1 = drQRow.RetrCode
                    Else
                        CurrLev1 = drQRow.RetrCode
                        If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ"
                    End If
                    If BranchReporting = True Then '10-30-12
                        CurrLev2 = CurrLev1 : CurrLev1 = drQRow.BranchCode '10-30-12 JTC SubTotChk9360
                    End If
                ElseIf Left(UH, 20) = "SPECIFIER CREDIT SEQ" Then  '02-27-01 WNA
                    If PrevLev1 = "" Then
                        PrevLev1 = drQRow.RetrCode
                        CurrLev1 = drQRow.RetrCode
                        If VQRT2.RepType = VQRT2.RptMajorType.RptSpecif Then 'OrderBy = "Reference"
                            PrevLev1 = drQRow("Reference") '02-06-10
                            CurrLev1 = drQRow("Reference") '02-06-10
                        End If
                    Else
                        CurrLev1 = drQRow.RetrCode
                        If VQRT2.RepType = VQRT2.RptMajorType.RptSpecif Then 'OrderBy = "Reference"
                            CurrLev1 = drQRow("Reference") '02-06-10
                        End If
                        If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ"
                    End If
                    If BranchReporting = True Then '10-30-12
                        CurrLev2 = CurrLev1 : CurrLev1 = drQRow.BranchCode '10-30-12 JTC SubTotChk9360
                    End If
                ElseIf Left(UH, 18) = "MARKET SEGMENT SEQ" Then  '"MARKET SEGMENT SEQ"
                    'If DebugOn ThenStop' 'Punch MarketSegment
                    If PrevLev1 = "" Then
                        PrevLev1 = drQRow.RetrCode
                        CurrLev1 = drQRow.RetrCode
                    Else
                        CurrLev1 = drQRow.RetrCode
                        If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ"
                    End If
                    'CurrLev2 = drQRow.Status : PrevLev2 = drQRow.Status
                End If
                If BranchReporting = True Then '10-30-12
                    CurrLev2 = CurrLev1 : CurrLev1 = drQRow.BranchCode '10-30-12 JTC SubTotChk9360
                End If
            End If
            If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ" '10-30-12 
            '05-17-11 None If VQRT2.SubSeq = VQRT2.SubSortType.SubSSpecif ThenStop : If DebugOn ThenStop 'Punch CurrLev2 = CurrLev2 'Else PREVSLS = CurrLev2
            '03-03-12 SubSSelectBidDate do nothingIf VQRT2.SubSeq = VQRT2.SubSortType.SubSSelectBidDate Then CurrLev2 = drQRow.Status ' PrevLev2 
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Then CurrLev2 = drQRow.Status ' PrevLev2 
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Then CurrLev2 = drQRow.SLSQ ' PrevLev2 Else
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSProj Then CurrLev2 = drQRow.JobName '09-14-10ProjectName
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSEnterDate Then CurrLev2 = drQRow.EntryDate 'SubSEnterDate'SubSProjCode'11-20-10
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSProjCode Then CurrLev2 = drQRow.QuoteCode 'SubSEnterDate'SubSProjCode'11-20-10
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSDescend Then CurrLev2 = drQRow.Sell
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSBidDate Then CurrLev2 = drQRow.BidDate.ToString '10-13-11
            If BranchReporting = True Then '10-30-12
                CurrLev2 = CurrLev1 : CurrLev1 = drQRow.BranchCode '10-30-12 JTC SubTotChk9360
            End If
            If Cmd = "EOF" Then CurrLev1 = "ZZZ" : CurrLev2 = "ZZZ"
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSProj Or VQRT2.SubSortType.SubSEnterDate Or VQRT2.SubSortType.SubSEnterDate Or VQRT2.RptMajorType.RptFollowBy Then GoTo Lev1_Brk_9500 '11-20-10 No Job Sub TotalsCurrLev2 = drQRow.ProjectName
Lev2_Brk_9400:  '12-09-09 If Lev1 break force Level 2 break
            If PrevLev2 <> CurrLev2 Or Trim$(PrevLev1) <> Trim(CurrLev1) Then
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And RealCustomer = True And RealManufacturer = False And VQRT2.RepType = VQRT2.RptMajorType.RptProj And VQRT2.SubSeq = VQRT2.SubSortType.SubSProj Then   '03-11-14
                    GoTo 9375 '12-10-09
                ElseIf frmQuoteRpt.txtPrimarySortSeq.Text <> "Salesman/Customer" Then
                    GoTo Lev1_Brk_9500 '12-10-09
                End If
                '03-11-14 If frmQuoteRpt.txtPrimarySortSeq.Text <> "Salesman/Customer" Then GoTo Lev1_Brk_9500 '12-10-09
                'Print sub totals for each Level 2 = Lower level Break
                THDG = "**TOTAL " & PrevLev1 & " / " & PrevLev2
                THDG += "  Count = " & QuantityA(2).ToString '01-29-13 Add Count 
                '09-23-15 JTC 
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And (RealManufacturer = False Or RealCustomerOnly = False Or RealSLSCustomer = True Or RealALL = False) Then '09-23-14 JTC Add Or RealALL = True)) ThenRealSLSCustomer = True Then '09-23-14 JTC Add Or RealALL = True)) Then
                    'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)) Then
                    THDG = "**TOTAL " & Left(PrevLev1 & "           ", 10) & Left(SortCode & Space(25), 25) '09-23-15
                    If Cmd = "EOF" Then  Else SortCode = drQToRow.FirmName '09-22-15 JTC 
                End If
                Call TotPrt9250(THDG, TotalLevels.TotLv2, RT, doc)
                PrevLev2 = CurrLev2
                ' 'zero Totals Lev2
                Dim X As String = "ZeroLevels" '02-09-09
                Call TotalsCalc(X, B, TotalLevels.TotLv2) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
            End If
Lev1_Brk_9500:
            If Trim$(PrevLev1) <> Trim(CurrLev1) Then
                'Print sub totals for each Major Break
                THDG = "**TOTAL " & PrevLev1 & "  " & PrevLev1 '12-09-09 
                '09-23-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)) Then
                    THDG = "**TOTAL " & Left(PrevLev1 & "           ", 10) & Left(SortCode & Space(25), 25) '09-23-15
                    If Cmd = "EOF" Then  Else SortCode = drQToRow.FirmName '09-22-15 JTC 
                    GoTo SkipNextIf '09-21-15
                    'If TgName(PrtCols) = "EntryDate" Or TgName(PrtCols) = "SLSCode" Or TgName(PrtCols) = "Status" Or TgName(PrtCols) = "SLSQ" Then
                    '    frmQuoteRpt.tgr.Splits(0).DisplayColumns(PrtCols).Visible = False
                    '    frmQuoteRpt.tgr.Splits(0).DisplayColumns(PrtCols).Width = 0
                    'End If
                End If ' End '09-18-15 
                If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Or VQRT2.RepType = VQRT2.RptMajorType.RptSalesman Or VQRT2.RepType = VQRT2.RptMajorType.RptStatus Or VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Or BranchReporting = True Then  '01-28-13 frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" 01-24-13 RetQutCode10-30-12 03-10-12 Chg SpecifierCode to FollowedBY
                    THDG += "  Count = " & QuantityA(1).ToString '03-03-12 "
                End If
SkipNextIf:     '09-21-15
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.txtPrimarySortSeq.Text = "Descending Sales Dollars" And SortSeq = "projectcust.Sell" Then GoTo SkipSubTotal '06-28-14 JTC Skip SubTotal if "Descending Sales Dollars" only Sequence
                Call TotPrt9250(THDG, TotalLevels.TotLv1, RT, doc)
SkipSubTotal:
                ' 'zero Totals Lev2
                PrevLev1 = CurrLev1
                Dim X As String = "ZeroLevels" '02-09-09
                Call TotalsCalc(X, B, TotalLevels.TotLv1) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
                If frmQuoteRpt.chkSalesmanPerPage.CheckState = CheckState.Checked Then
                    If frmQuoteRpt.ChkTotalsOnly.CheckState = CheckState.Unchecked Then
                        If THDG.StartsWith("**GRAND") Then '11-19-10
                            RT.Rows(RC).PageBreakBehavior = BreakEnum.None '11-19-10 
                            RT.BreakAfter = BreakEnum.None
                        Else
                            RT.Rows(RC).PageBreakBehavior = BreakEnum.Page '11-19-10 
                            RT.BreakAfter = BreakEnum.Page '11-19-10 
                        End If
                    End If
                ElseIf frmQuoteRpt.chkBlankLine.CheckState = CheckState.Checked Then
                    If frmQuoteRpt.chkSalesmanPerPage.CheckState = CheckState.Checked Then GoTo 9364 '
                    If Trim$(PrevLev1) = "ZZZ" And Trim(CurrLev1) = "ZZZ" Then GoTo 9364 '10-25-10 No blank Line on ZZZ
                    RC += 1 '10-25-10 
                    'Dim tmp As Integer = RT.Rows.Count
                    'Debug.Print(RC.ToString) RT.Cells(tmp, 0).Text = "  " : RT.Cells(tmp, 0).SpanCols = RT.Cols.Count '10-25-10
                    RT.Cells(RC, 0).Text = "  " : RT.Cells(RC, 0).SpanCols = RT.Cols.Count '10-25-10
                    RC += 1 '10-25-10  'RT.Cells(RC, 0).Text = "  " : RT.Cells(RC, 0).SpanCols = 10 : RT.Cells(RC, 1).Text = "  "
9364:           End If '10-25-10 
            End If
            GoTo Exit9375

9365:       'SLS or Status FollowedBy
            If PrevLev1 = "" Then 'If First Then ' Left$(UH$, 8) = "SALESMAN" Or Left$(UH$, 6) = "STATUS"
                If SortSeq = "quote.SLSQ" Or SortSeq = "QS.SLSCode" Then ' Primary Sort '03-08-13  or SortSeq = "QS.SLSCode" 
                    CurrLev1 = drQRow.SLSQ : PrevLev1 = drQRow.SLSQ
                    CurrLev2 = drQRow.Status : PrevLev2 = drQRow.Status
                ElseIf SortSeq = "quote.FollowedBy" Then '03-03-12 Primary Sort
                    CurrLev1 = drQRow.FollowBy : PrevLev1 = drQRow.FollowBy
                    'CurrLev1 = drQRow.SLSQ ' PrevLev1 = drQRow.SLSQ ' '03-03-12
                    'No CurrLev2 = drQRow.Status : PrevLev2 = drQRow.Status
                ElseIf SortSeq = "quote.EnteredBy" Then '05-14-13
                    CurrLev1 = drQRow.EnteredBy : PrevLev1 = drQRow.EnteredBy
                    If Trim(PrevLev2) = "" Then PrevLev2 = Trim(CurrLev2) '05-14-13
                Else
                    CurrLev1 = drQRow.Status : PrevLev1 = drQRow.Status
                    CurrLev2 = drQRow.SLSQ : PrevLev2 = drQRow.SLSQ
                End If
                First = 0
                If BranchReporting = True Then '10-30-12
                    CurrLev2 = CurrLev1 : CurrLev1 = drQRow.BranchCode '10-30-12 JTC SubTotChk9360
                    PrevLev2 = PrevLev1 : PrevLev1 = drQRow.BranchCode '10-30-12 JTC SubTotChk9360
                End If
            End If

            'If DebugOn ThenDebug.Print(UH)
            'Do Sub Totals first
            If VQRT2.SubSeq = VQRT2.SubSortType.SubSStatus Or VQRT2.SubSeq = VQRT2.SubSortType.SubSSls Or (BranchReporting = True And (Left(UH, 8) = "SALESMAN" Or Left(UH, 6) = "STATUS" Or Left(UH, 11) = "FOLLOWED BY" Or Left(UH, 11) = "ENTERED BY")) Then '10-30-12
                '    CurrLev2 = CurrLev1 : CurrLev1 = drQRow.BranchCode '10-30-12 JTC SubTotChk9360
                'End IfThen
                '12-09-09 Force a lev2 break on Lev1 break
                If Trim$(PrevLev2) <> Trim(CurrLev2) Or Trim$(PrevLev1) <> Trim(CurrLev1) Then
                    'Call PrtMajTot9390(PrevLev1 &  PrevLev2, TotalLevels.TotLv2, RT, doc) 'Print sub totals for each specifier
                    THDG = "**TOTAL " & PrevLev1 & "  " & PrevLev2 & "  Count = " & QuantityA(2).ToString '10-30-12
                    Call TotPrt9250(THDG, TotalLevels.TotLv2, RT, doc)
                    PrevLev2 = CurrLev2
                    ' 'zero Totals Lev2
                    Dim X As String = "ZeroLevels" '02-09-09
                    Call TotalsCalc(X, B, TotalLevels.TotLv2) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
                End If
            End If
            If Trim$(PrevLev1) <> Trim(CurrLev1) Then
                'Call PrtMajTot9390(PrevLev1 & " " & PrevLev1, TotalLevels.TotLv1, RT, doc) 'Print sub totals for each specifier
                THDG = "**TOTAL " & PrevLev1 & " " & PrevLev1 & "  Count = " & QuantityA(1).ToString '03-03-12 
                'Debug.Print(RC.ToString)
                Call TotPrt9250(THDG, TotalLevels.TotLv1, RT, doc)
                'Debug.Print(RC.ToString)

                PrevLev1 = CurrLev1
                Dim X As String = "ZeroLevels" '02-09-09
                Call TotalsCalc(X, B, TotalLevels.TotLv1) 'Lev.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
                If frmQuoteRpt.chkSalesmanPerPage.CheckState = CheckState.Checked Then
                    If frmQuoteRpt.ChkTotalsOnly.CheckState = CheckState.Unchecked Then 'Not On Totals Only
                        If THDG.StartsWith("**GRAND") Then '11-19-10
                            RT.Rows(RC).PageBreakBehavior = BreakEnum.None '11-19-10 
                            RT.BreakAfter = BreakEnum.None
                        Else
                            RT.Rows(RC).PageBreakBehavior = BreakEnum.Page '11-19-10 
                            RT.BreakAfter = BreakEnum.Page '11-19-10 
                            'RT.Rows(RC).PageBreakBehavior = PageBreakBehaviorEnum.MustBreak '03-11-12 
                        End If
                    End If
                    If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-08-12
                        doc.Body.Children.Add(RT) : RT = New C1.C1Preview.RenderTable : RC = 0 '03-09-12
                        RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10 RT.SplitHorzBehavior = True '05-26-10 Test
                        RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
                        RT.CellStyle.Padding.Left = "1mm" '12-13-12
                        RT.CellStyle.Padding.Right = "1mm" '12-13-12
                        RT.Style.GridLines.All = LineDef.Empty '  
                    End If
                ElseIf frmQuoteRpt.chkBlankLine.CheckState = CheckState.Checked Then
                    If Trim$(PrevLev1) = "ZZZ" And Trim(CurrLev1) = "ZZZ" Then GoTo SkipBlank
                    If THDG.StartsWith("**GRAND") Then GoTo SkipBlank '03-10-12 
                    If RT.BreakAfter = BreakEnum.Page Then GoTo SkipBlank '03-10-12 
                    '03-13-12 RC += 1 '10-25-10 
                    'Debug.Print(RC.ToString)
                    RT.Cells(RC, 0).Text = "  " : RT.Cells(RC, 0).SpanCols = RT.Cols.Count '10-25-10
                    RC += 1 '10-25-10 
                    'RT.Cells(RC, 0).Text = "  " : RT.Cells(RC, 0).SpanCols = 10 : RT.Cells(RC, 1).Text = "  "
                End If '10-25-10 
SkipBlank:
            End If
            'Debug.Print(RC.ToString)
        Catch ex As Exception  'Try Catch with  so you can fix exception after the error message **Get gid of  'CatchStop b/4 releasing
            MessageBox.Show("Error in SubTotChk9360 (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox(ex.Message.ToString & vbCrLf & "SubTotChk9360(VQRT)" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop  'Debug.WriteLine(ex.Message.ToString)
        End Try '

9375:
Exit9375:  ' #End
    End Sub
    Public Sub PrintQuoteLineRpt946(ByRef A As String, ByRef RT As C1.C1Preview.RenderTable) '01-30-09
        Try
            'Debug.Print(RC.ToString)
            'Call PrintQuoteLineRpt946(A$,FixSell!,FixProfit!, FixProfitPer!,LampSell!,LampCost!,Amt!,CommAmt!,Commpct!)'01-13-09
            '01-25-09 TotalLevels.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
946:        Dim FixMargin As Decimal '07-08-09
            Dim LpMargin As Decimal '07-08-09
            QuantityA(0) += 1 : QuantityA(1) += 1 : QuantityA(2) += 1 '02-07-10 Count RowCnt += 1
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
                FixSell = drQToRow.Sell : FixCost = drQToRow.Cost
                LampSell = drQToRow.LPSell : LampCost = drQToRow.LPCost
            Else
                Dim Probability As Decimal = 1 '10-11-09 'Extend By quote probability 
                If frmQuoteRpt.ChkExtendByProb.CheckState = CheckState.Checked Then Probability = Val(drQRow.Probability) '10-11-09
                FixSell = drQRow.Sell * Probability : FixCost = drQRow.Cost * Probability '10-11-09
                LampSell = drQRow.LPSell * Probability : LampCost = drQRow.LPCost * Probability '10-11-09 
            End If
            FixProfit = FixSell - FixCost
            If FixSell <> 0 Then FixProfitPer = FixProfit / (FixSell + 0.00001) Else FixProfitPer = 0 '08-22-02 WNA
            LampProfit = LampSell - LampCost
            If LampSell <> 0 Then LampProfitPer = LampProfit / (LampSell + 0.00001) Else LampProfitPer = 0 '08-22-02 WNA
            '12-09-09 Totals needed ''''''''''''''''''''
            CommAmt = drQRow.Comm '12-09-09LampSell = drQRow.LPSell:FixSell = drQRow.Sell :LampCost = drQRow.LPCost:FixCost = drQRow.Cost
            FixSellExt = FixSell
            FixCostExt = FixCost
            'None On Header LnQuantityA = CDec(Val(frmQuoteRpt.tg(Row, "Qty"))) '09-08-09 

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If DIST Then 'All Decimals FixSell!,FixProfit!, FixProfitPer!,LampSell!,LampCost!,Amt!,CommAmt!,Commpct!
                'FixProfit = FixSell - FixCost
                FixMargin = (FixSell - FixCost) / (FixSell + 0.0001) * 100 '07-08-09
                'LampProfit = LampSell - LampCost
                LpMargin = (LampSell - LampCost) / (LampSell + 0.0001) * 100 '07-08-09
                'Commpct = (CommAmt / (Amt + 0.0001)) * 100 '
                If FixMargin > 900 Then FixMargin = 999 Else If FixMargin < -900 Then FixMargin = -999
                If LpMargin > 900 Then LpMargin = 999 Else If LpMargin < -900 Then LpMargin = -999
            Else
                Amt = drQRow.Sell '02-26-01 WNA'''''''''''''''''''''''''
                'CommAmt! = drQRow.Comm
                If DAYB Then
                    CommAmt = drQRow.Sell '11-26-01 Reverse Cost & Sell
                    'Amt! = drQRow.Comm:
                End If
                If MFG Then '05-09-01 WNA
                    Commpct = ((Amt - CommAmt) / (Amt + 0.0001)) * 100
                Else
                    Commpct = (CommAmt / (Amt + 0.0001)) * 100 '
                End If
                If Commpct > 900 Then Commpct = 999 Else If Commpct < -900 Then Commpct = -999 '06-24-04
                FixProfitPer = Commpct '02-05-09
            End If
            Dim ColText As String
            Dim ColName As String
            Dim ColCaption As String
            Dim I As Int16
            Dim PC As Int16 = 0 'PC = Print Column
            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy And SESCO = True Then '03-09-12
                'Debug.Print(RC.ToString)         '03-09-12RT.Cols.Clear()
                ColText = frmQuoteRpt.tgQh.Splits(0).DisplayColumns("FollowBy").DataColumn.Text '
                RT.Cells(RC, PC).Text = ColText '03-12-12 vbCrLf &
                RT.Cols(PC).Width = ".7in" 'FollowBy .7'04-25-12 Increase size of Quote# field on RptMajorType.RptFollowBy And SESCO = True
                RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center
                RT.Cells(RC, PC).Style.TextAlignVert = AlignVertEnum.Center '03-13-12
                'Debug.Print(frmQuoteRpt.tgQh.Splits(0).DisplayColumns(PC).Height.ToString)
                RT.Rows(RC).Height = ".4in" : PC += 1 '.5
                RT.Cells(RC, PC).Text = "Select" & vbCrLf & frmQuoteRpt.tgQh.Splits(0).DisplayColumns("SelectCode").DataColumn.Text '03-12-12 CD Code
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Center
                RT.Cols(PC).Width = ".5in" : PC += 1 '.6
                ColText = frmQuoteRpt.tgQh.Splits(0).DisplayColumns("Sell").DataColumn.Text ' ), "########0.00")
                'ColText = Replace(ColText, "$", "")
                'ColText = Replace(ColText, ",", "")
                RT.Cells(RC, PC).Text = "   " & frmQuoteRpt.tgQh.Splits(0).DisplayColumns("JobName").DataColumn.Text & vbCrLf & Space(20) & "Total= " & ColText '03-09-12
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Left
                RT.Cells(RC, PC).Style.FontBold = True '03-12-12
                RT.Cols(PC).Width = "3.15in" : PC += 1 '03-13-12 "2.8in"
                RT.Cells(RC, PC).Text = "Bid Date:" & vbCrLf & frmQuoteRpt.tgQh.Splits(0).DisplayColumns("BidDate").DataColumn.Text
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Center
                RT.Cols(PC).Width = ".7in" : PC += 1 ' 1 to .7 '03-22-12
                RT.Cells(RC, PC).Text = "Quote#" & vbCrLf & frmQuoteRpt.tgQh.Splits(0).DisplayColumns("QuoteCode").DataColumn.Text
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Center
                RT.Cols(PC).Width = "1.2in" : PC += 1 ' 1 to .7 '04-25-12 .7
                RT.Cells(RC, PC).Text = "SLSQ:" & vbCrLf & frmQuoteRpt.tgQh.Splits(0).DisplayColumns("SLSQ").DataColumn.Text
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Center
                RT.Cols(PC).Width = ".5in" : PC += 1 '04-25-12 
                '03-22-12 Added Status Column
                RT.Cells(RC, PC).Text = "Status:" & vbCrLf & frmQuoteRpt.tgQh.Splits(0).DisplayColumns("Status").DataColumn.Text
                RT.Cells(RC, PC).Style.TextAlignHorz = AlignHorzEnum.Center
                RT.Cols(PC).Width = ".8in" : PC += 1
                RT.Cells(RC, PC).Text = "Location:" & vbCrLf & frmQuoteRpt.tgQh.Splits(0).DisplayColumns("City").DataColumn.Text & "," & frmQuoteRpt.tgQh.Splits(0).DisplayColumns("State").DataColumn.Text
                RT.Cols(PC).Width = "1.4" ' "1.85" ' "1.5in" '03-13-12 PC += 1
                RT.Rows(RC).Style.BackColor = LightSkyBlue '01-19-13 .LightBlue
                RT.Rows(RC).Style.FontSize = 11 '03-12-12
                'Debug.Print(RT.Rows(RC).Style.FontName.ToString & "    " & RT.Rows(RC).Style.FontSize.ToString) '03-12-12
                RT.Cells(RC, PC).Style.TextAlignVert = AlignVertEnum.Center '03-13-12
                RT.Rows(RC).Style.TextAlignVert = AlignVertEnum.Center '03-13-12
                RT.Style.GridLines.All = LineDef.Empty '03-12-12
                'RT.Width = "9.5" '"8.9in" ' auto" '03-13-12 'RT.Width = ""
                doc.Body.Children.Add(RT) : RT = New C1.C1Preview.RenderTable : RC = 0 '03-9-12
                ''03-12-12 *************************************************
                RT.SplitHorzBehavior = SplitBehaviorEnum.SplitNewPage '05-26-10 RT.SplitHorzBehavior = True '05-26-10 Test
                RT.Style.Padding.All = "0mm" : RT.Style.Padding.Top = "0mm" : RT.Style.Padding.Bottom = "0mm"
                RT.CellStyle.Padding.Left = "1mm" '12-13-12
                RT.CellStyle.Padding.Right = "1mm" '12-13-12
                RT.Style.GridLines.All = LineDef.Empty '  LineDef.Default  '12-04-10 & "  UserID = " & UserID 
                'Dim fs As Integer = frmQuoteRpt.FontSizeComboBox.Text
                'RT.Style.Font = New Font(frmQuoteRpt.RibbonFontComboBox2.Text, fs, FontStyle.Bold)
                ''03-12-12 *************************************************
                Exit Sub
            End If
            'ProjectCustID 1:ProjectID 1:QuoteCode 1:NCode 1:FirmName 1:ContactName 1:SLSCode 1:Got 1:Typec 1:MFGQuoteNumber 1:Cost 1:Sell 1:Comm 1:Overage 1:ChgDate 1:OrdDate 1:NotGot 1:Comments 1:SPANumber 1:SpecCross 1:LotUnit 1:LPCost 1:LPSell 1:LPComm 1:LampsIncl 1:Terms 1:FOB 1:QuoteID 1:BranchCode 1:ProjectName 1:MarketSegment 1:EntryDate 1:BidDate 1:SLSQ 1:Status 1:RetrCode 1:SelectCode 1:LastChgBy 1:City 1:State 1:CSR 1:StockJob 1:LotUnit1 1:
            'Debug.Print(frmQuoteRpt.tgQh.Splits(0).DisplayColumns("QuoteCode").DataColumn.Text)
            For I = 0 To MaxCol ' 02-03-09 frmFoll.tg.Splits(0).DisplayColumns.Count - 1
                'Debug.Print(RC.ToString & " / " & RT.Cols.Count)
                ' Dim col2 As C1.Win.C1TrueDBGrid.C1DisplayColumn = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I) '02-20-09
                If frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                If (frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                ColText = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).DataColumn.Text  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                ColName = frmQuoteRpt.tgQh.Splits(0).DisplayColumns(I).Name '
                '@#ProjectName0,ProjectID1,QuoteID2,QuoteCode3,EntryDate4,RetrCode5,PRADate6,EstDelivDate7,SLSQ8,Status9,BidDate10,Cost11,Sell12,Margin13,LPCost14,LPSell15,LPMarg16,LotUnit17,StockJob18,CSR19,LastChgBy20,HeaderTab21,LinesYN22,SelectCode23,Password24,FollowBy25,OrderEntryBy26,ShipmentBy27,Remarks28,LightingGear29,Dimming30,LastDateTime31,BidBoard32,EnteredBy33,BidTime34,BranchCode35,Address36,Address237,City38,State39,Zip40,Country41,Location42,LeadTime43,"'02-22-09
                RT.Cells(RC, PC).Text = ColText '02-04-09 frmQuoteRpt.tg.Splits(0).DisplayColumns(Col).DataColumn.Text  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                RT.Cols(PC).Width = TgWidth(I) '07-14-10 
                '01-06-13
                If ColName = "CustName" Or ColName = "JobName" Or ColName = "FirmName" Then '01-06-13 FirmName Left
                    RT.Cols(PC).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left '11-10-10 AlignHorzEnum.Left '11-10-10 
                    If ColName = "CustName" Then If ColText.Length > MaxNameLength Then RT.Cells(RC, PC).Text = ColText.Substring(0, MaxNameLength) '01-03-12 
                    If ColName = "JobName" Then If ColText.Length > MaxJobLength Then RT.Cells(RC, PC).Text = ColText.Substring(0, MaxJobLength) '01-03-12  
                End If
                If ColName = "Type" Then
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center '12-03-09
                End If
                If ColName = "BidBoard" Then
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center '12-03-09
                End If
                If ColName = "BidTime" And ColText = "00:00:00" Then RT.Cells(RC, PC).Text = " " '09-10-10 Blank Bid Date
                If ColName = "EntryDate" Or ColName = "BidDate" Then
                    Dim MyDate As Date
                    If IsDate(RT.Cells(RC, PC).Text) Then
                        MyDate = CDate(RT.Cells(RC, PC).Text)
                        RT.Cells(RC, PC).Text = VB6.Format(MyDate, "MM-dd-yy")
                        If ColName = "BidDate" And ColText = "01/01/00" Then RT.Cells(RC, PC).Text = " " '09-10-10 Blank Bid Date
                        RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    End If
                End If
                Dim Hdg As String
                'Cost11,Sell12,Margin13,LPCost14,LPSell15,LPMarg16,LotUnit17,StockJob18,CSR19,LastChgBy20,HeaderTab21,LinesYN22,SelectCode23,Password24,FollowBy25,OrderEntryBy26,ShipmentBy27,Remarks28,LightingGear29,Dimming30,LastDateTime31,BidBoard32,EnteredBy33,BidTime34,BranchCode35,Address36,Address237,City38,State39,Zip40,Country41,Location42,LeadTime43,"'02-22-09
                If ColName = "Comm-$" Then '12-05-09
                    ColText = Replace(ColText, "$", "") '12-05-09
                    ColText = Replace(ColText, ",", "") '12-05-09
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                    RT.Cells(RC, PC).Text = Format(Val(ColText), DecFormat) '01-06-12 "########0.00"
                    If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RT.Cells(RC, PC).Text = "" '10-17-10 GoTo 652 'Skip
                End If
652:
                If ColName = "Cost" Then
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                    RT.Cells(RC, PC).Text = Format(FixCost, DecFormat) '10-12-09 
                End If
                If ColName = "Sell" Then '10-12-09
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                    RT.Cells(RC, PC).Text = Format(FixSell, DecFormat) '10-12-09 
                End If
                If ColName = "LPCost" Then '10-12-09
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                    RT.Cells(RC, PC).Text = Format(LampCost, DecFormat) '10-12-09 
                End If
                If ColName = "LPSell" Then '10-12-09
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                    RT.Cells(RC, PC).Text = Format(LampSell, DecFormat) '10-12-09 
                End If
                If ColName = "Margin-$" Or ColName = "Margin-%" Then '03-21-14
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                End If
                If DIST Then Hdg = "Cost" Else Hdg = "Comm" '02-05-09
                ColText = Replace(ColText, ",", "") '07-07-09
                If DIST Then Hdg = "Margin" Else Hdg = "Comm" '02-04-09
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And ColName = "Comm" Then ColCaption = "Margin"
                If ColName = Hdg Or ColName = "Comm" Or ColName = "Margin" Then '07-08-09 Tag = 13 Then '
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '12-03-09
                    If DIST Then
                        RT.Cells(RC, PC).Text = Format(FixMargin, "##0.00")
                    Else 'Rep Comm %
                        RT.Cells(RC, PC).Text = Format(FixProfitPer, "##0.00")
                        GoTo 750
                    End If
650:                If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RT.Cells(RC, PC).Text = "" '10-17-10 GoTo 650 'Skip
                End If
                If DIST Then Hdg = "LPMarg" Else Hdg = "LPComm" '02-04-09
                If ColName = "LPComm" Or ColName = "LPMarg" Then ' Tag = 16 Then 'ColName = Hdg Or ColName = "LPComm" Then
                    If DIST Then
                        RT.Cells(RC, PC).Text = Format(LpMargin, "##0.00")
                    Else
                        'LampProfit
                        RT.Cells(RC, PC).Text = Format(LampProfit, "#####0.00") '12-09-09 Was LampProfitPer
                        If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RT.Cells(RC, PC).Text = "" '10-17-10 GoTo 850 'Skip
                    End If
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
750:            End If
                If ColName = "Comm-%" Then '12-05-09
                    RT.Cells(RC, PC).Text = Format(Val(ColText), "##0.00")
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right
                    If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RT.Cells(RC, PC).Text = "" '10-17-10 GoTo 850 'Skip
                End If
850:            RT.Cells(RC, PC).Style.BackColor = Color.White '06-14-10 
                PC += 1 '02-20-09 Dim PC As Int16 'PC = Print Column

            Next
            '#End
            Exit Sub
        Catch ex As Exception  'Try Catch with  so you can fix exception after the error message **Get gid of  'CatchStop b/4 releasing
            MessageBox.Show("Error in PrintQuoteLineRpt946 (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox(ex.Message.ToString & vbCrLf & "PrintQuoteLineRpt946(VQRT)" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop  'Debug.WriteLine(ex.Message.ToString)
        End Try '           Replace * with DblQuote 
    End Sub
    Public Sub PrintQuoteRealLineRpt946(ByRef A As String, ByRef RT As C1.C1Preview.RenderTable) '01-30-09
        Try  '#Top
            'Call PrintQuoteRealLineRpt946(A$,FixSell!,FixProfit!, FixProfitPer!,LampSell!,LampCost!,Amt!,CommAmt!,Commpct!)'01-13-09
            '01-25-09 TotalLevels.TotGt TotLv1 TotLv2 TotLv3 TotLv4 TotPrt 
946:        'Debug.Print(drQToRow.NCode & drQToRow.QuoteCode)
            '02-05-12If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
            A = drQToRow.Sell : A = Replace(A, "$", "") : A = Replace(A, ",", "") : FixSell = Val(A) '02-05-12
            A = drQToRow.Cost : A = Replace(A, "$", "") : A = Replace(A, ",", "") : FixCost = Val(A) '02-05-12
            A = drQToRow.Comm : A = Replace(A, "$", "") : A = Replace(A, ",", "") : CommAmt = Val(A) '02-05-12
            '02-05-12 FixSell = drQToRow.Sell : FixCost = drQToRow.Cost
            LampSell = drQToRow.LPSell : LampCost = drQToRow.LPCost
            '02-05-12 CommAmt = drQToRow.Comm '12-09-09LampSell = drQRow.LPSell:FixSell = drQRow.Sell :LampCost = drQRow.LPCost:FixCost = drQRow.Cost
            '07-06-12 If RealExtByInfluencePercent = True Then '02-04-12 
            'Lpsell has Influence% 'Do You only want Specifiers with an Influence % Greater Than Zero?"
            'FixSell = FixSell * Val(LampSell) '02-04-12
            'FixCost = FixCost * Val(LampSell) '02-04-12
            'CommAmt = CommAmt * Val(LampSell) '02-04-12
            'End If
            'Else
            'FixSell = drQRow.Sell : FixCost = drQRow.Cost
            'LampSell = drQRow.LPSell : LampCost = drQRow.LPCost
            'End If
            ' Public RealizSellExt As Decimal '12-10-09    Public RealizSellAExt(5) As Decimal
            RealizSellExt = 0 '01-18-12
            If drQToRow.Got = True Then RealizSellExt = drQToRow.Sell '12-10-09
            FixSellExt = FixSell '12-09-09
            FixCostExt = FixCost
            FixProfit = FixSell - FixCost
            If FixSell <> 0 Then FixProfitPer = FixProfit / (FixSell + 0.00001) Else FixProfitPer = 0 '08-22-02 WNA
            LampProfit = LampSell - LampCost
            If LampSell <> 0 Then LampProfitPer = LampProfit / (LampSell + 0.00001) Else LampProfitPer = 0 '08-22-02 WNA
            Dim FixMargin As Decimal '-07-09-09
            Dim LpMargin As Decimal '-07-09-09
            'Debug.Print(frmQuoteRpt.tgQh.Splits(0).DisplayColumns("LPSell").DataColumn.Text) ' & TgWidth(Col).ToString & "  " & CStr(I))
            If DIST Then 'All Decimals FixSell!,FixProfit!, FixProfitPer!,LampSell!,LampCost!,Amt!,CommAmt!,Commpct!
                FixMargin = (FixSell - FixCost) / (FixSell + 0.0001) * 100 '07-08-09
                'LampProfit = LampSell - LampCost
                LpMargin = (LampSell - LampCost) / (LampSell + 0.0001) * 100 '07-08-09
                If FixMargin > 900 Then FixMargin = 999 Else If FixMargin < -900 Then FixMargin = -999
                If LpMargin > 900 Then LpMargin = 999 Else If LpMargin < -900 Then LpMargin = -999
            Else
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" Then
                    Amt = drQToRow.Sell
                    If DAYB Then
                        CommAmt = drQToRow.Sell '11-26-01 Reverse Cost & Sell
                        'Amt! = drQRow.Comm:
                    End If
                End If
            End If
            If MFG Then '05-09-01 WNA
                Commpct = ((Amt - CommAmt) / (Amt + 0.0001)) * 100
            Else
                Commpct = (CommAmt / (Amt + 0.0001)) * 100 '
            End If
            If Commpct > 900 Then Commpct = 999 Else If Commpct < -900 Then Commpct = -999 '06-24-04
            FixProfitPer = Commpct '02-05-09

            Dim ColText As String = ""
            Dim ColName As String = ""
            Dim ColCaption As String = ""
            Dim I As Int16 = 0
            Dim PC As Int16 = 0 'PC = Print Column
            'Dim Col As Int16
            '  Dim Tag As String '2-24-09
            MaxCol = frmQuoteRpt.tgr.Splits(0).DisplayColumns.Count - 1 '06-14-10 
            For I = 0 To MaxCol ' 02-03-09 frmFoll.tgr.Splits(0).DisplayColumns.Count - 1
                'ColName = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Name 'Test
                If frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Visible = False Then Continue For ' GoTo 145 '02-03-09If frmQuoteRpt.tg.Splits(0).DisplayColumns(col).Visible = False Then GoTo 145 '02-03-09
                If (frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Width / 100) < 0.1 Then Continue For ' Too Small so don't print
                ColText = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).DataColumn.Text  'dis  'Columns(col).CellText(row) '.ToString 'frmFoll.tg.Splits(0).DisplayColumns(Cat).DataColumn.Text 'Trim(drFRow.Category)
                ColName = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Name
                '@#R ProjectCustID0,ProjectID1,NCode2,Got3,Typec4,QuoteCode5,ProjectName6,FirmName7,ContactName8,EntryDate9,SLSCode10,Status11,Cost12,Sell13,Margin14,LPCost15,LPSell16,LPMarg17,Overage18,ChgDate19,OrdDate20,NotGot21,Comments22,SPANumber23,SpecCross24,LotUnit25,LampsIncl26,Terms27,FOB28,QuoteID29,BranchCode30,MarketSegment31,MFGQuoteNumber32,BidDate33,SLSQ34,RetrCode35,SelectCode36,LeadTime37,"
                RT.Cells(RC, PC).Text = ColText '02-04-09 )
                RT.Cols(PC).Width = frmQuoteRpt.tgr.Splits(0).DisplayColumns(I).Width / 100 '09-15-10 
                'Debug.Print(RT.Cols(PC).Width)
                If ColName = "CustName" Or ColName = "JobName" Or ColName = "FirmName" Then '12-23-12 FirmName Left
                    RT.Cols(PC).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left '11-10-10 AlignHorzEnum.Left '11-10-10 
                    If ColName = "CustName" Then If ColText.Length > MaxNameLength Then RT.Cells(RC, PC).Text = ColText.Substring(0, MaxNameLength) '01-03-12 
                    If ColName = "JobName" Then If ColText.Length > MaxJobLength Then RT.Cells(RC, PC).Text = ColText.Substring(0, MaxJobLength) '01-03-12  
                End If
                If ColName = "Got" Then '07-09-09 Tag = 3 Then 'Got
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center '02-04-09
                    If ColText = False Then RT.Cells(RC, PC).Text = "N" Else RT.Cells(RC, PC).Text = "Y" '02-25-09
                End If
                If ColName = "Type" Then '07-09-09 Tag = 4 Then 'Type 
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Center '02-04-09
                    RT.Cells(RC, PC).Text = ColText
                End If
                If ColName = "EntryDate" Or ColName = "BidDate" Then '07-09-09 Tag = 9 Or Tag = 33 Then 'If ColName = "EntryDate" Or ColName = "BidDate" Then
                    Dim MyDate As Date
                    If IsDate(RT.Cells(RC, PC).Text) Then
                        MyDate = CDate(RT.Cells(RC, PC).Text)
                        RT.Cells(RC, PC).Text = VB6.Format(MyDate, "MM-dd-yy")
                        If ColName = "BidDate" And ColText = "01/01/00" Then RT.Cells(RC, PC).Text = " " '09-10-10 Blank Bid Date
                        RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    End If
                End If
                'Cost12,Sell13,Margin14,LPCost15,LPSell16,LPMarg17
                If ColName = "Sell" Then '07-09-09 If Tag = 13 Then 'If ColName = "Sell" Then
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    '02-05-12 ColText = Replace(ColText, "$", "") '07-07-09
                    ColText = Replace(ColText, ",", "") '07-07-09
                    RT.Cells(RC, PC).Text = Format(Val(FixSell), DecFormat) '02-05-12
                End If

                'If DIST Then Hdg = "Cost" Else Hdg = "Comm" '02-05-09
                If ColName = "Cost" Or ColName = "Comm-$" Or ColName = "Comm" Then '01-18-11 12-10-09 If Tag = 12 Then ' If ColName = Hdg Then
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    '02-05-12 ColText = Replace(ColText, "$", "") '07-07-09
                    '02-05-12 ColText = Replace(ColText, ",", "") '07-07-09
                    If ColName = "Cost" Then RT.Cells(RC, PC).Text = Format(Val(FixCost), DecFormat) '02-05-12
                    If ColName = "Comm-$" Or ColName = "Comm" Then RT.Cells(RC, PC).Text = Format(Val(CommAmt), DecFormat) '02-05-12
                    If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RT.Cells(RC, PC).Text = "" '01-24-13 'Skip
                End If
                'If DIST Then Hdg = "Margin" Else Hdg = "Comm" '02-04-09
                If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And ColName = "Comm" Then ColCaption = "Margin"
                If ColName = "Comm-%" Or ColName = "Margin" Then '12-10-09 Tag = 14 Then 'If ColName = Hdg Or ColName = "Comm" Then '02-15-09
                    If DIST Then
                        RT.Cells(RC, PC).Text = Format(FixMargin, "##0.00")
                    Else 'Rep Comm %
                        RT.Cells(RC, PC).Text = Format(FixProfitPer, "##0.00")
                    End If
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RT.Cells(RC, PC).Text = "" '10-17-10 GoTo 850 '10-17-10  'Skip
                End If
SkipComm956:    '01-24-13
650:            'LPCost15,LPSell16,LPMarg17 Realization
                If ColName = "LPSell" Or ColName = "LPCost" Then '07-07-09 If Tag = 16 Then 'If ColName = "LPSell" The
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
                    ColText = Replace(ColText, "$", "") '07-07-09
                    ColText = Replace(ColText, ",", "") '07-07-09
                    RT.Cells(RC, PC).Text = Format(Val(ColText), DecFormat)
                    If ColName = "LPCost" And frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RT.Cells(RC, PC).Text = "" '01-24-13 'Skip
                End If
                If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then GoTo 750 '01-24-13 'Skip
                If ColName = "LPComm" Or ColName = "LPMarg" Then '07-07-09 If Tag = 17 Then 'If ColName = Hdg Or ColName = "LPComm" Then
                    If DIST Then
                        RT.Cells(RC, PC).Text = Format(LpMargin, "##0.00")
                    Else
                        RT.Cells(RC, PC).Text = Format(LampProfitPer, "##0.00")
                    End If
                    If frmQuoteRpt.chkIncludeCommDolPer.Checked = CheckState.Unchecked Then RT.Cells(RC, PC).Text = "" '10-17-10 GoTo 750 'Skip
                    RT.Cols(PC).Style.TextAlignHorz = AlignHorzEnum.Right '02-04-09
750:            End If
                '09-23-15 If TotalsOnly on SLS-Cust or Cust  or Mfg Delete EntryDate,SLSCode,STATUS,SLSQ
                'If frmQuoteRpt.pnlTypeOfRpt.Text = "Realization" And frmQuoteRpt.chkDetailTotal.CheckState = CheckState.Checked And (RealManufacturer = True Or RealCustomerOnly = True Or RealSLSCustomer = True Or RealALL = True) Then '09-23-14 JTC Add Or RealALL = True)) Then
                '    If ColName = "EntryDate" Or ColName = "SLSCode" Or ColName = "Status" Or ColName = "SLSQ" Then
                '        '09-21-15 RT.Cells(RC, PC).Text = "" 'frmQuoteRpt.tgr.Splits(0).DisplayColumns(PrtCols).Visible = False 'frmQuoteRpt.tgr.Splits(0).DisplayColumns(PrtCols).Width = 0
                '    End If
                'End If ' End '09-18-15 


                PC += 1 '02-20-09 Dim PC As Int16 'PC = Print Column
            Next
            'RT.Cols(RC).Style.BackColor = Color.White ' RT.Cells(RC, PC).Style.BackColor = Color.White '06-14-10 
            '
            Exit Sub

        Catch ex As Exception  'Try Catch with  so you can fix exception after the error message **Get gid of  'CatchStop b/4 releasing
            MessageBox.Show("Error in PrintQuoteRealLineRpt946 (VQRT)" & vbCrLf & ex.Message & vbCrLf & "If the problem persists call Multimicro for support", "VQRT)", MessageBoxButtons.OK, MessageBoxIcon.Error) '07-06-12MsgBox(ex.Message.ToString & vbCrLf & "PrintQuoteRealLineRpt946(VQRT)" & vbCrLf)
            ' If DebugOn ThenStop 'CatchStop  'Debug.WriteLine(ex.Message.ToString)
        End Try '           Replace * with DblQuote 
        '#End
    End Sub
    Public Sub AbortCheck52(ByRef A As String)
        System.Windows.Forms.Application.DoEvents()
        If AbortPrtFlag = True Then
            'frmQuoteRpt.vsPrinter1.KillDoc()
            A = "Cancel" 'GoTo 998
        End If
    End Sub

    Public Sub FillSecondarySortCombo()
        On Error Resume Next
        'Debug.Print(frmQuoteRpt.txtPrimarySortSeq.Text)
        Dim IFollowBy As Int16 = 0
        frmQuoteRpt.cboSortSecondarySeq.Items.Clear() '11-19-10
        If frmQuoteRpt.txtPrimarySortSeq.Text = "Print Spec Credit Lines" Then '11-20-10
            '11-19-10 frmQuoteRpt.cboSortSecondarySeq.Text = ""
            frmQuoteRpt.cboSortSecondarySeq.Items.Clear()
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Quote Code")
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Job Name")
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Salesman")
        ElseIf frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And (VQRT2.RepType = VQRT2.RptMajorType.RptQutCode Or VQRT2.RepType = VQRT2.RptMajorType.RptProj) Then '11-04-14 JTC
            frmQuoteRpt.cboSortSecondarySeq.Items.Clear() '11-04-14 JTC Add Quote Summary Rept (Job Name or Quote Code) Seconary Sort Option Salesman 1-4 Split Show Dollars at split percent FollowBy = Sls * percent
            frmQuoteRpt.cboSortSecondarySeq.Text = "None"
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("None")
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Salesman 1-4 Splits")
        ElseIf ExcelQuoteFU = True Then '04-29-15 JTC
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Job Name")
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Bid Date")
        Else

            frmQuoteRpt.cboSortSecondarySeq.Text = ""
            frmQuoteRpt.cboSortSecondarySeq.Items.Clear()
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Job Name")
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman" Then '11-20-10
            Else
                frmQuoteRpt.cboSortSecondarySeq.Items.Add("Salesman")
            End If
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Status")
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Bid Date")
            frmQuoteRpt.cboSortSecondarySeq.Items.Add("Descending Dollar")
            If VQRT2.RepType = VQRT2.RptMajorType.RptFollowBy Then '03-03-12
                '03-06-12 frmQuoteRpt.cboSortSecondarySeq.Items.Add("Job Name/BidDate") '03-03-12
                frmQuoteRpt.cboSortSecondarySeq.Items.Add("Select-Priority / BidDate") '03-03-12 'Select-Priority= Q.SelectCode
                'Debug.Print(frmQuoteRpt.cboSortSecondarySeq.Items.Count - 1)
                IFollowBy = frmQuoteRpt.cboSortSecondarySeq.Items.Count - 1
                frmQuoteRpt.cboSortSecondarySeq.SelectedIndex = IFollowBy '03-06-12 frmQuoteRpt.cboSortSecondarySeq.Items.Count - 1
                'IFollowBy = frmQuoteRpt.cboSortSecondarySeq.Items.Count - 1

            End If
            If frmQuoteRpt.pnlTypeOfRpt.Text = "Quote Summary" And (frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman" Or frmQuoteRpt.txtPrimarySortSeq.Text = "Status") Then '11-22-10
                frmQuoteRpt.cboSortSecondarySeq.Items.Add("Quote Code")
                frmQuoteRpt.cboSortSecondarySeq.Items.Add("Enter Date")
            Else
                If frmQuoteRpt.txtPrimarySortSeq.Text = "Salesman" Then
                    frmQuoteRpt.cboSortSecondarySeq.Items.Add("Specifiers")
                    frmQuoteRpt.cboSortSecondarySeq.Items.Add("Retrieval Code")
                    frmQuoteRpt.cboSortSecondarySeq.Items.Add("Market Segment") '07-30-04 JH
                End If
            End If
            'PrimarySortSeq = Name Code
            If SESCO = False And frmQuoteRpt.txtPrimarySortSeq.Text.StartsWith("Salesman") = False And frmQuoteRpt.pnlTypeOfRpt.Text <> "Quote Summary" Then '05-21-13
                'If RealManufacturer = True And SESCO = False And RealCustomer = False And RealALL = False And RealOther = False And RealArchitect = False And RealEngineer = False And RealLtgDesigner = False And RealSpecifier = False And RealContractor = False And RealOther = False Then '03-24-13
                frmQuoteRpt.cboSortSecondarySeq.Items.Add("Spread Sheet by Month") ''05-20-13
                frmQuoteRpt.cboSortSecondarySeq.Items.Add("Spread Sheet by Year") '06-22-15 JTC Add Realization Report Spread Sheet by Year

            End If


        End If

        frmQuoteRpt.cboSortSecondarySeq.Text = VB6.GetItemString(frmQuoteRpt.cboSortSecondarySeq, IFollowBy) '03-03-12
    End Sub
    Public Function ReformatDate(ByRef OrigDate As String) As String
        Dim Century As String
        On Error Resume Next
        If Right(OrigDate, 2) < "90" Then Century = "20" Else Century = "19"
        ReformatDate = Century & Right(OrigDate, 2) & Left(OrigDate, 4)
    End Function
    Public Function ReturnZarg(ByVal ZargSearch As String) As String
        Dim B As Integer, E As Integer '    Call function  Agnam = ReturnZarg("/Nam=")
        B = Zarg.IndexOf(ZargSearch) : If B >= 0 Then E = Zarg.IndexOf("|", B) : If E >= 0 Then Return Zarg.Substring(B + ZargSearch.Length, E - B - ZargSearch.Length) : Exit Function
        Return ""

    End Function
    Public Sub SetRTSWidth(ByRef R As RenderTable) '03-09-12 call
        R.Cols(0).Width = ".5in"
        R.Cols(1).Width = ".75in"
        R.Cols(2).Width = "2.in" '11-24-09 
        R.Cols(3).Width = "2.in"
        R.Cols(4).Width = "1in"
        R.Cols(5).Width = "1in"
        R.Cols(6).Width = "1in"
        R.Cols(7).Width = "1in" '03-06-12 

    End Sub
    Public Sub SetSelectionHeader()
        'Dim CheckState As Object
        On Error Resume Next
        SelectionText = ""
        frmQuoteRpt.txtStartEntry.Text = VB6.Format(frmQuoteRpt.DTPickerStartEntry, "MMddyy") '01-26-09
        frmQuoteRpt.txtEndEntry.Text = VB6.Format(frmQuoteRpt.DTPicker1EndEntry, "MMddyy") '01-26-09
        'Entry Date
        If Trim(frmQuoteRpt.txtStartEntry.Text) <> "" And Trim(frmQuoteRpt.txtStartEntry.Text) <> "ALL" And Trim(frmQuoteRpt.txtEndEntry.Text) <> "" And Trim(frmQuoteRpt.txtEndEntry.Text) <> "ALL" Then
            SelectionText = "Entry- " & frmQuoteRpt.txtStartEntry.Text & " - " & frmQuoteRpt.txtEndEntry.Text
        End If
        'Bid  Date
        If frmQuoteRpt.ChkCheckBidDates.CheckState = CheckState.Checked Then '02-04-12 No Bid Info
            frmQuoteRpt.txtStartBid.Text = VB6.Format(frmQuoteRpt.DTPicker1StartBid, "MMddyy") '01-26-09
            frmQuoteRpt.txtEndBid.Text = VB6.Format(frmQuoteRpt.DTPicker1EndBid, "MMddyy") '01-26-09
            If Trim(frmQuoteRpt.txtStartBid.Text) <> "" And Trim(frmQuoteRpt.txtStartBid.Text) <> "ALL" And Trim(frmQuoteRpt.txtEndBid.Text) <> "" And Trim(frmQuoteRpt.txtEndBid.Text) <> "ALL" Then
                SelectionText = SelectionText & " Bid- " & Trim(frmQuoteRpt.txtStartBid.Text) & " - " & frmQuoteRpt.txtEndBid.Text
            End If
            'Print Blank Bids
            If frmQuoteRpt.chkBlankBidDates.CheckState = CheckState.Checked Then
                SelectionText = SelectionText & "Include Blank Bid,"
            End If
        End If
        'Quote Amount
        If Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "0" And Trim(frmQuoteRpt.txtEndQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtEndQuoteAmt.Text) <> "999999999" Then '"999999999" '03-24-08 JTC Added 9 "999,999,999"
            SelectionText = SelectionText & " QAmt- " & VB6.Format(frmQuoteRpt.txtStartQuoteAmt.Text, "$##,###,###") & "-" & VB6.Format(frmQuoteRpt.txtEndQuoteAmt.Text, "$##,###,###")
        Else
            If Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtStartQuoteAmt.Text) <> "0" And (Trim(frmQuoteRpt.txtEndQuoteAmt.Text) = "" Or Trim(frmQuoteRpt.txtEndQuoteAmt.Text) = "999999999") Then '"999999999" '03-24-08 JTC Added 9 "999,999,999"
                SelectionText = SelectionText & " QAmt- " & VB6.Format(frmQuoteRpt.txtStartQuoteAmt.Text, "$##,###,###") & "-$9,999,999"
            End If
            If (Trim(frmQuoteRpt.txtStartQuoteAmt.Text) = "" Or Trim(frmQuoteRpt.txtStartQuoteAmt.Text) = "0") And Trim(frmQuoteRpt.txtEndQuoteAmt.Text) <> "" And Trim(frmQuoteRpt.txtEndQuoteAmt.Text) <> "999999999" Then '"999999999" '03-24-08 JTC Added 9 "999,999,999"
                SelectionText = SelectionText & " QAmt- $0-" & VB6.Format(frmQuoteRpt.txtEndQuoteAmt.Text, "$##,###,###")
            End If
            '02-04-12 Don't Print QAmt Hdr if Default
            If (Trim(frmQuoteRpt.txtStartQuoteAmt.Text) = "" Or Trim(frmQuoteRpt.txtStartQuoteAmt.Text) = "0") And (Trim(frmQuoteRpt.txtEndQuoteAmt.Text) = "" Or Trim(frmQuoteRpt.txtEndQuoteAmt.Text) = "999999999") Then '"999999999" '03-24-08 JTC Added 9 "999,999,999"
                '02-04-12 SelectionText = SelectionText & " QAmt- $0-$9,999,999"
            End If
        End If
        'SLS
        If Trim(frmQuoteRpt.txtSalesman.Text) <> "" And Trim(frmQuoteRpt.txtSalesman.Text) <> "ALL" Then
            SelectionText = SelectionText & " Sls- " & Trim(frmQuoteRpt.txtSalesman.Text)
        Else
            '02-04-12 Don't Print on Hdr if Default
            '02-04-12 SelectionText = SelectionText & " Sls- ALL"
        End If
        'Status
        If Trim(frmQuoteRpt.txtStatus.Text) = "" Or Trim(frmQuoteRpt.txtStatus.Text) = "ALL" Then
            '02-04-12 Don't Print on Hdr if Default SelectionText = SelectionText & " Status- ALL"
        Else
            SelectionText = SelectionText & " Status- " & Left(frmQuoteRpt.txtStatus.Text, 20)
        End If
        If Trim(frmQuoteRpt.txtSpecifierCode.Text) <> "" And Trim(frmQuoteRpt.cbospeccross.Text) <> "ALL" Then '04-25-05 JH
            SelectionText = SelectionText & " Spec- " & Trim(frmQuoteRpt.cbospeccross.Text) '04-25-05 JH
        Else
            '02-04-12 Don't Print on Hdr if DefaultSelectionText = SelectionText & " Spec- ALL"
        End If
        If Trim(frmQuoteRpt.txtRetrieval.Text) <> "" And Trim(frmQuoteRpt.txtRetrieval.Text) <> "ALL" Then
            SelectionText = SelectionText & " Retrieval- " & Trim(frmQuoteRpt.txtRetrieval.Text)
        Else
            '02-04-12 Don't Print on Hdr if DefaultSelectionText = SelectionText & " Retrieval- ALL"
        End If
        If Trim(frmQuoteRpt.txtLastChgBy.Text) <> "" And Trim(frmQuoteRpt.txtLastChgBy.Text) <> "ALL" Then '09-16-02 WNA
            SelectionText = SelectionText & " LastChgBy- " & Trim(frmQuoteRpt.txtLastChgBy.Text)
        Else
            '02-04-12 Don't Print on Hdr if Default SelectionText = SelectionText & " LastChgBy- ALL"
        End If
        If RealWithOneMfgCust = True And RealWithOneMfgCustCode.Trim <> "" Then  '03-22-13 
            SelectionText = SelectionText & " With One Code = " & RealWithOneMfgCustCode.Trim '03-22-13 
        End If
    End Sub
End Module