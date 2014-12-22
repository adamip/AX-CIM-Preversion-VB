Imports System
Imports System.IO
Imports System.Xml
Imports devBiz.Net.Mail
Imports System.Windows.Forms
Imports System.Data.SqlClient

Imports CIM_PreVersion.Constants
Imports CIM_PreVersion.EventLog

Public Class IntegrationParser
    'Private objlogger As EventLog = New EventLog
    Private DManager As DataManager
    Private mQuoteKeyString As String = ""
    Private mDatabaseServer As String = ""
    Private mDatabaseName As String = ""
    Private mDatabasePassword As String = ""
    Private mDatabaseUsername As String = ""
    Private mQuoteKey As String = ""
    Private mInfoKey As String = ""
    Private mCustKey As String = ""
    Private NumberOfSites As Integer = 0
    Private CurrentSite As Integer = 0
    Private BaseQuoteKey As String = ""
    Private SolutionType As String = ""
    Private strFormattedXML As String = ""
    Private RecordEffected As Long
    Private sFolder As String = ""
    Private Subject As String = ""
    Private Company As String = ""
    Private Owner As String = ""

    Dim TicketQueue As String = "", DataAreaID As String = "", Partner3_ As String = "", Caller As String = "", CallerName As String = ""
    Dim CallerPhone As String = "", CallerEmail As String = "", PartnerTicketSource As String = ""
    Dim SeverityLevel As String = "", SeverityLevelDescription As String = ""
    Dim Customer As String = "", CustomerName As String = ""
    Dim PartnerTicketNo As String = "", SiteName As String = ""
    Dim Product As String = "", SN As String = "", ContactName As String = "", ContactPhone As String = ""
    Dim ContactEmail As String = "", ProblemDescription As String = ""
    Dim UpdatePartner As Integer = 0, UpdateCaller As Integer = 0, UpdateContact As Integer = 0

    Dim RequestStatus As String = "", SQLTicketChangesString As String = "", SQLTicketChangesEventString As String = ""

    Public Shared sEventEntry As String = ""           ' Adam Ip 2011-04-07
    Public Shared Debug_Mode As Boolean = False
    Public Shared Inform_Mode As Boolean = True

    Public Sub New()
        If Not System.IO.Directory.Exists(Application.StartupPath & "\Attachments\") Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Attachments\")
        End If
        sFolder = Application.StartupPath & "\Attachments\"

        DatabaseServer = "db0.adtech.net"
        DatabaseName = "DynamicsAxProd"
        DatabaseUserName = "Integration"
        DatabasePassword = "1ntegr@te"
    End Sub

    ''' <summary>
    ''' This Function is the Entry Point called by Windows Service             Adam Ip
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ProcessEmailInbox() As Boolean
        Const MAX_AREA As Integer = 2
        Dim i As Integer = 0, InboxCount As Integer = 0
        Dim DataAreaIDcode(MAX_AREA - 1) As String
        Dim TakeABreak(MAX_AREA - 1) As Boolean   ' Adam Ip 2011-08-11
        Dim PSM As Boolean = False, UMI As Boolean = False, UMIIP As Long

        Dim msg As MailMessage
        Dim Inbox(MAX_AREA - 1) As POP3

        DataAreaIDcode(0) = "UK"
        Inbox(0) = New POP3()
        Inbox(0).Username = "emeasupport"
        Inbox(0).Password = "zZ3k.D&tk9bew"
        Inbox(0).Host = "INPOP.it.adtech.net"
        TakeABreak(0) = False
        DataAreaIDcode(1) = "US"
        Inbox(1) = New POP3()
        Inbox(1).Username = "support"
        Inbox(1).Password = "zZ3k.D&tk9bew"
        Inbox(1).Host = "INPOP.it.adtech.net"
        TakeABreak(1) = False

        ProcessEmailInbox = False
        InboxCount = 0

        If Debug_Mode Then
            sEventEntry = vbTab & "Service Entry Point - UK US" & vbCrLf & "CIM Service: #01 Function ProcessEmailInbox"
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        End If
        For i = 0 To MAX_AREA - 1
            DataAreaID = DataAreaIDcode(i)

            If Inbox(i).Connect() = False Then
                sEventEntry = "CIM Service: [" & DataAreaID & "] #02 Function ProcessEmailInbox connecting to mailbox failed"
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                TakeABreak(i) = True
            Else
                REM Connecting inbox succcessfully
                sEventEntry = "CIM Service: [" & DataAreaID & "]  #03 Function ProcessEmailInbox started." & vbCrLf _
                            & vbTab & Convert.ToString(Inbox(i).Count) & " [" & DataAreaID & "] e-mails found."
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                If OpenDB(TakeABreak(i)) = False Then
                    Inbox(i).Disconnect()
                    TakeABreak(i) = True
                    sEventEntry = "CIM Service: [" & DataAreaID & "] #04 Function ProcessEmailInbox connecting to database failed, returns " _
                        & Convert.ToString(ProcessEmailInbox) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak(i))
                    WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                ElseIf TakeABreak(i) = False Then
                    REM Connecting database successfully
                    For Each msg In Inbox(i)
                        If TakeABreak(i) = False Then
                            Try
                                ' make sure this email is from the past week at least!
                                ' Dim TodaysDate As DateTime = DateTime.Now()

                                Application.DoEvents()
                                REM No Message ID found, then this is a new message, so process.  
                                REM If Message ID found, then previously read e-mail, so no process.
                                If FindMessageID(msg.MessageID, TakeABreak(i)) = False Then
                                    If TakeABreak(i) = False Then
                                        If Debug_Mode Then
                                            sEventEntry = "CIM Service: [" & DataAreaID & "] #05 Function ProcessEmailInbox" & vbCrLf _
                                                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak(i)) & vbCrLf _
                                                & "Data Area ID: " & DataAreaID & vbCrLf _
                                                & "Message ID: " & msg.MessageID.ToString() & vbCrLf _
                                                & "Message Date: " & msg.Date().ToString() & vbCrLf _
                                                & "Message Body: " & vbCrLf & vbTab & "---- Beginning of lines ----" & vbCrLf _
                                                & msg.PlainMessage.Body.ToString() & vbTab & vbCrLf _
                                                & "---- End of lines ----"
                                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                        End If

                                        'if e-mail is NOT from Verint nor aip@adtech.net then ValidateEmailSource(msg.From.ToString()) returns false
                                        '   then simply only update the MessageIDs table through UpdateMessageID()
                                        'if ValidateEmailSource(msg.From.ToString()) returns true, then email is either from  
                                        '    from Verint nor aip@adtech.net, then proceed with ProcessSupportMessage()
                                        'Don't combine these 2 following IF's into one IF
                                        '   If a = false Or B Then ...    the logic is different
                                        If ValidateEmailSource(msg.MessageID, msg.From.ToString(), TakeABreak(i)) = False Then
                                            UMI = UpdateMessageID(msg.MessageID, TakeABreak(i))
                                            If Debug_Mode Then WriteToEventLog("CIM Service: [" & DataAreaID & "] #06 Function ProcessEmailInbox, UpdateMessageID returns " _
                                                & Convert.ToString(UMI), EventLogEntryType.Information)
                                        Else
                                            UMIIP = UpdateMessageIDinProgress(msg.MessageID, "A", TakeABreak(i))
                                            If Debug_Mode Then WriteToEventLog("CIM Service: [" & DataAreaID & "] #07 Function ProcessEmailInbox" & vbCrLf _
                                                    & "UpdateMessageIDinProgress returns " & Convert.ToString(UMIIP), EventLogEntryType.Information)
                                            If UMIIP = 0 And TakeABreak(i) = False Then
                                                REM most processing mechanism happens in ProcessSupportMessage( )
                                                PSM = ProcessSupportMessage(msg, DataAreaID, TakeABreak(i))
                                                If Debug_Mode Then WriteToEventLog("CIM Service: [" & DataAreaID & "] #08 Function ProcessEmailInbox" & vbCrLf _
                                                    & "ProcessSupportMessage returns " & Convert.ToString(PSM), EventLogEntryType.Information)
                                                If PSM = True Then
                                                    If TakeABreak(i) = False Then UMI = UpdateMessageID(msg.MessageID, TakeABreak(i))
                                                    If TakeABreak(i) = False Then UpdateMessageIDinProgress(msg.MessageID, "D", TakeABreak(i))
                                                    If Debug_Mode Then WriteToEventLog("CIM Service: [" & DataAreaID & "] #09 Function ProcessEmailInbox" & vbCrLf _
                                                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak(i)) & vbCrLf _
                                                        & "UMI: " & Convert.ToString(UMI) & vbCrLf _
                                                        & "UMIIP: " & Convert.ToString(UMIIP), EventLogEntryType.Information)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                                REM TakeABreak(i) = True
                                sEventEntry = "CIM Service Exception: [" & DataAreaID & "] #10 Function ProcessEmailInbox returns " _
                                    & Convert.ToString(ProcessEmailInbox) & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak(i)) & vbCrLf _
                                    & "Message ID: " & msg.MessageID.ToString() & vbCrLf _
                                    & "Message Date: " & msg.Date().ToString() & vbCrLf _
                                    & "Message Body: " & vbCrLf & vbTab & "-- Beginning of lines --" & vbCrLf _
                                    & msg.PlainMessage.Body.ToString() & vbCrLf & vbTab _
                                    & "-- End of lines --"
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                REM here should not Exit Function, because Exit Function in Prorgram Entry Point is actually Exit Program
                                'Exit Function
                            End Try
                        End If  REM TakeABreak(i) = False
                    Next msg
                    InboxCount = InboxCount + 1
                    If Debug_Mode Then
                        sEventEntry = "CIM Service: [" & DataAreaID & "] #11 Function ProcessEmailInbox" & vbCrLf _
                            & "TakeABreak(i) = " & Convert.ToString(TakeABreak(i)) & vbCrLf _
                            & "i = " & Convert.ToString(i) & vbCrLf _
                            & "PSM = " & Convert.ToString(PSM) & vbCrLf _
                            & "UMI = " & Convert.ToString(UMI) & vbCrLf _
                            & "UMIIP = " & Convert.ToString(UMIIP) & vbCrLf _
                            & "InboxCount = " & Convert.ToString(InboxCount)
                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    End If
                End If
                Inbox(i).Disconnect()
            End If
        Next i
        If TakeABreak(0) = False And TakeABreak(1) = False And InboxCount = MAX_AREA Then ProcessEmailInbox = True
        REM Else
        REM     ProcessEmailInbox = False
        REM End If
        If Debug_Mode Then
            sEventEntry = "CIM Service: Function ProcessEmailInbox #12 returns " & Convert.ToString(ProcessEmailInbox) & vbCrLf _
                & "[" & DataAreaIDcode(0) & "] TakeABreak(0) = " & Convert.ToString(TakeABreak(0)) & vbCrLf _
                & "[" & DataAreaIDcode(1) & "] TakeABreak(1) = " & Convert.ToString(TakeABreak(1)) & vbCrLf _
                & "i = " & Convert.ToString(i) & vbCrLf _
                & "PSM = " & Convert.ToString(PSM) & vbCrLf _
                & "UMI = " & Convert.ToString(UMI) & vbCrLf _
                & "UMIIP = " & Convert.ToString(UMIIP) & vbCrLf _
                & "InboxCount = " & Convert.ToString(InboxCount)
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        End If
        CloseDB()
        ClearMemory()
    End Function

    ' ProcessSupportMessage returns whether the message has been successfully processed
    Private Function ProcessSupportMessage(ByVal msg As MailMessage, ByVal DataAreaID As String, ByRef TakeABreak As Boolean) As Boolean
        Dim SQLString As String = ""

        REM Local variables
        Dim sTicketEmailBody As String = msg.PlainMessage.Body.ToString()
        Dim sTicketID As String = ""
        Dim sIntegrationRequestType As String = "NEW"
        Dim NextTicketRecID As Long
        Dim NextTicketEventRecID As Long

        REM return value initialized
        ProcessSupportMessage = False

        REM Global variables
        Partner3_ = ""
        PartnerTicketNo = ""
        TicketQueue = ""
        CallerName = ""
        CallerPhone = ""
        CallerEmail = ""
        SeverityLevel = ""
        SeverityLevelDescription = ""
        UpdatePartner = 0
        UpdateCaller = 0
        UpdateContact = 0
        SQLTicketChangesEventString = ""

        If Debug_Mode Then
            sEventEntry = "CIM Service: Function ProcessSupportMessage #01" & vbCrLf _
                & "Data Area ID: " & DataAreaID & vbCrLf _
                & "From: " & msg.From.ToString() & vbCrLf _
                & "To: " & msg.To.ToString() & vbCrLf _
                & "Subject: " & msg.Subject.ToString() & vbCrLf _
                & "Body: " & vbTab & "---- Beginning of lines ----" & vbCrLf _
                & sTicketEmailBody & vbCrLf & vbTab & "---- End of lines ----" & vbCrLf
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        End If

        If (msg.To.ToString().ToLower().Contains("cellstacksupport") Or msg.Cc.ToString().ToLower().Contains("cellstacksupport")) Then
            Try
                'DataAreaID = "US"
                Partner3_ = "OTH"
                TicketQueue = "CellStack"
                'Caller = "CON012643"
                'CallerName = msg.To.ToString()
                'CallerPhone = "1-800-494-8637"
                'CallerEmail = "feg-amer@witness.com"
                RequestStatus = ""
                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #02, Cellstack", EventLogEntryType.Information)
            Catch ex As Exception
                TakeABreak = True
                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #03 returns " & Convert.ToString(ProcessEmailInbox) & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                    & "Exception Target: " & ex.TargetSite.ToString()
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                Exit Function
            End Try

        ElseIf (InStr(UCase(msg.Subject), "CUSTOMER INCIDENT") > 0) And ((InStr(UCase(msg.Subject), "WITNESS") > 0) Or (InStr(UCase(msg.Subject), "VERINT") > 0)) Then
            Try
                If (InStr(UCase(msg.Subject), "UPDATED") > 0) Then
                    sIntegrationRequestType = "UPDATE"
                End If
                'DataAreaID = "US"
                Partner3_ = "VerintDir"
                TicketQueue = "Support Ct"
                Caller = "CON012643"
                CallerName = msg.To.ToString()
                CallerPhone = "1-800-494-8637"
                CallerEmail = "feg-amer@witness.com"
                RequestStatus = ""
                If (InStr(UCase(msg.Subject), "VERINT") > 0) Then ' spoof where it came from for testing...Oracle emails come with Verint in subject, Onyx with Witness
                    PartnerTicketSource = "Oracle"
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #04, Oracle", EventLogEntryType.Information)
                    ParseVerintOracleEmail(sTicketEmailBody, msg.MessageID, sIntegrationRequestType, TakeABreak)
                Else
                    PartnerTicketSource = "Onyx"
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #05, Onyx", EventLogEntryType.Information)
                    ParseVerintOnyxEmail(sTicketEmailBody, msg.MessageID, sIntegrationRequestType, TakeABreak)
                End If
            Catch ex As Exception
                TakeABreak = True
                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #06 returns " _
                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                    & "Exception Target: " & ex.TargetSite.ToString()
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                Exit Function
            End Try
        End If

        Try
            If TakeABreak = True Then Exit Function

            ' last level of defense!  Can't update w/o these values
            If Len(Partner3_) = 0 Or Len(PartnerTicketNo) = 0 Then UpdatePartner = 0
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function ProcessSupportMessage #07" & vbCrLf _
                    & "Partner3_: " & Partner3_ & vbCrLf _
                    & "PartnerTicketNo: " & PartnerTicketNo & vbCrLf _
                    & "UpdatePartner: " & Convert.ToString(UpdatePartner)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            If Len(ContactEmail) = 0 Then UpdateContact = 0
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function ProcessSupportMessage #08" & vbCrLf _
                    & "ContactEmail: " & ContactEmail & vbCrLf _
                    & "UpdateContact: " & Convert.ToString(UpdateContact)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            If Len(CallerEmail) = 0 Then UpdateCaller = 0
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function ProcessSupportMessage #09" & vbCrLf _
                    & "CallerEmail: " & CallerEmail & vbCrLf _
                    & "UpdateCaller: " & Convert.ToString(UpdateCaller)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #10 returns " _
                & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Function
        End Try

        REM Dim eventViewerString As String = ""
        If Len(Partner3_) > 0 Then ' partner integration needs to have a partner to process!
            If Len(PartnerTicketNo) > 0 Then ' Verint integration needs to have a ticket number specified to process!
                REM by this point, you should have parsed all applicable partner integration email messages...new or update.  
                REM The insert/update commands to follow are partner indepedent

                REM only insert if this is a new ticket
                If String.Compare(sIntegrationRequestType, "NEW", True) = 0 Then
                    Try
                        ' Get the next ticket id.  
                        sTicketID = GetNextTicketId(DataAreaID, TakeABreak)
                        'If CInt(Replace(TicketID, "AGST", "")) + 1 > GetTicketNumberSequence() Then ' Check if the number sequence table needs to be synched up too...
                        '    With DManager
                        '        SQLString = "UPDATE NumberSequenceTable SET NextRec = " & CInt(Replace(TicketID, "AGST", "")) + 2 & " Where DataAreaID = '" & DataAreaID & "' and numbersequence = 'CC_Tickets'"
                        '        .ExecuteSQL(SQLString)
                        '        'Write("CIM Service - Updated numbersequenceTable", "SQL: " & SQLString, 0)
                        '    End With
                        'End If
                        If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #11" & vbCrLf & "sTicketID:" & sTicketID, EventLogEntryType.Information)
                    Catch ex As Exception
                        TakeABreak = True
                        sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #12 returns " _
                            & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                            & "sTicketID:" & sTicketID & vbCrLf _
                            & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                            & "Exception Target: " & ex.TargetSite.ToString()
                        WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                        Exit Function
                    End Try

                    If TakeABreak Then Exit Function

                    'C_Tickets
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #13" & vbCrLf & "sTicketID: " & sTicketID, EventLogEntryType.Information)

                    With DManager
                        Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #14", EventLogEntryType.Information)
                            'SQLString = "INSERT INTO C_Tickets (" & _
                            '    "   CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, TicketID_Display, Source, Status, TicketQueue, Tier, Partner3_, PartnerTicketSource, UpdatePartner, CallerPhone, " & _
                            '    "   CallerEmail, RecID, RecVersion, DataAreaID, UpdateContact, UpdateCaller, Caller, CallerName, severityLevel, CustomerName, " & _
                            '    "   PartnerTicketNo, SiteName, Product, SN, ContactName, ContactPhone, ContactEmail, ProblemDescription) " & _
                            '    "" & _
                            '    "VALUES (" & _
                            '    "   'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & TicketID & "', 'Email', 'Open', '" & TicketQueue & "', 'Tier 1', '" & Partner3_ & "', '" & PartnerTicketSource & "', 1, '" & Replace(CallerPhone, "'", "''") & "', " & _
                            '    "   '" & Replace(CallerEmail, "'", "''") & "', '" & GetNextTicketRecId() & "', 1, '" & DataAreaID & "', 1, 1, '" & Replace(Caller, "'", "''") & "', '" & Replace(Mid(CallerName, 1, 60), "'", "''") & "', '" & SeverityLevel & "', '" & Replace(Mid(CustomerName, 1, 60), "'", "''") & "', " & _
                            '    "   '" & PartnerTicketNo & "', '" & Replace(Mid(SiteName, 1, 60), "'", "''") & "', '" & Replace(Product, "'", "''") & "', '" & Replace(SN, "'", "''") & "', '" & Replace(Mid(ContactName, 1, 60), "'", "''") & "', '" & Replace(ContactPhone, "'", "''") & "', '" & Replace(ContactEmail, "'", "''") & "', '" & Replace(ProblemDescription, "'", "''") & "')"
                            NextTicketRecID = GetNextTicketRecID(TakeABreak)
                            If TakeABreak Then Exit Function
                            If Debug_Mode Then
                                sEventEntry = "CIM Service: Function ProcessSupportMessage #15" & vbCrLf & "sTicketID: " & sTicketID & vbCrLf _
                                    & "NextTicketRecID: " & Convert.ToString(NextTicketRecID)
                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            End If

                            SQLString = "INSERT INTO C_Tickets ( " _
                                        & " CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, TicketID_Display, Source, Status, TicketQueue, Tier," _
                                        & " Partner3_, PartnerTicketSource, UpdatePartner, CallerPhone, " & " CallerEmail," _
                                        & " RecID, RecVersion, DataAreaID, UpdateContact, UpdateCaller," _
                                        & " Caller, CallerName," _
                                        & " SeverityLevel, CustomerName," _
                                        & " PartnerTicketNo, SiteName," _
                                        & " Product, SN," _
                                        & " ContactName, ContactPhone," _
                                        & " ContactEmail," _
                                        & " ProblemDescription )" _
                                        & "VALUES ( " _
                                        & " 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & sTicketID & "', 'Email', 'Open', '" & TicketQueue & "', 'Tier 1', '" _
                                        & Partner3_ & "', '" & PartnerTicketSource & "', 1, '" & Replace(CallerPhone, "'", "''") & "', " & "   '" & Replace(CallerEmail, "'", "''") & "', '" _
                                        & NextTicketRecID & "', 1, '" & DataAreaID & "', 1, 1, '" _
                                        & Replace(Caller, "'", "''") & "', '" & Replace(Mid(CallerName, 1, 60), "'", "''") & "', '" _
                                        & SeverityLevel & "', '" & Replace(Mid(CustomerName, 1, 60), "'", "''") & "', '" _
                                        & PartnerTicketNo & "', '" & Replace(Mid(SiteName, 1, 60), "'", "''") & "', '" _
                                        & Replace(Product, "'", "''") & "', '" & Replace(SN, "'", "''") & "', '" _
                                        & Replace(Mid(ContactName, 1, 60), "'", "''") & "', '" & Replace(ContactPhone, "'", "''") & "', '" _
                                        & Replace(ContactEmail, "'", "''") & "', '" _
                                        & Replace(ProblemDescription, "'", "''") & "' )"

                            If Debug_Mode Or Inform_Mode Then
                                sEventEntry = "CIM Service: Function ProcessSupportMessage #16" & vbCrLf _
                                    & "sTicketID: " & sTicketID & vbCrLf _
                                    & "NextTicketRecID: " & Convert.ToString(NextTicketRecID) & vbCrLf & vbCrLf _
                                    & "SN: " & ContactName & vbCrLf _
                                    & "Contact Name: " & SN & vbCrLf _
                                    & "Contact Phone: " & ContactPhone & vbCrLf _
                                    & "Contact Email: " & ContactEmail & vbCrLf _
                                    & "Problem Desc: " & ProblemDescription & vbCrLf & vbCrLf _
                                    & "SQL to execute: " & SQLString
                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            End If
                            .ExecuteSQL(SQLString)
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #17", EventLogEntryType.Information)
                        Catch ex As Exception
                            TakeABreak = True
                            sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #18 returns " _
                                & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                & "Exception Target: " & ex.TargetSite.ToString()
                            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                            Exit Function
                        End Try
                        If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #19", EventLogEntryType.Information)
                        IncrementTicketNumberSequence(TakeABreak)
                    End With

                    If TakeABreak Then Exit Function
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #20", EventLogEntryType.Information)

                    'C_TicketEvents
                    NextTicketEventRecID = GetNextTicketEventRecID(TakeABreak)
                    If TakeABreak Then Exit Function

                    If Debug_Mode Then
                        sEventEntry = "CIM Service: Function ProcessSupportMessage #21" & vbCrLf _
                            & "NextTicketEventRecID: " & Convert.ToString(NextTicketEventRecID)
                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    End If

                    Try
                        With DManager
                            SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, " _
                                        & "CreatedDateTime, ModifiedBy, ModifiedDateTime, SeverityDescription, SeverityLevel, " _
                                        & "DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('New Ticket', 0, '" & sTicketID & "', 'New " & Partner3_ & " Integration Email received: " _
                                        & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(sTicketEmailBody, "'", "''") _
                                        & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription _
                                        & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & NextTicketEventRecID & "', 1, 1, 0)"

                            'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                            'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                            If Debug_Mode Or Inform_Mode Then
                                sEventEntry = "CIM Service: Function ProcessSupportMessage #22" & vbCrLf _
                                    & "Inserting Ticket Event" & vbCrLf _
                                    & "Ticket ID: " & sTicketID & vbCrLf _
                                    & "Ticket Event ID : " & Convert.ToString(NextTicketEventRecID) & vbCrLf _
                                    & "SQL to execute: " & SQLString
                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            End If
                            .ExecuteSQL(SQLString)
                        End With
                    Catch ex As Exception
                        TakeABreak = True
                        sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #23 returns " _
                            & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                            & "SQL: " & SQLString & vbCrLf & vbCrLf _
                            & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                            & "Exception Target: " & ex.TargetSite.ToString()
                        WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                        Exit Function
                    End Try
                Else ' update ticket
                    ' get AGS ticket from partner's ticket number
                    Dim ds As DataSet = New DataSet
                    Dim dt As DataTable = New DataTable
                    Dim Status As String = ""
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #24", EventLogEntryType.Information)
                    Try

                        SQLString = "SELECT TOP 1 TicketID_Display, Status FROM C_Tickets WHERE PartnerTicketNo = '" & PartnerTicketNo _
                            & "' ORDER BY TicketID_Display DESC"

                        If Debug_Mode Then
                            sEventEntry = "CIM Service: Function ProcessSupportMessage #25" & vbCrLf _
                                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                & "Partner Ticket ID: " & PartnerTicketNo & vbCrLf _
                                & "sTicketID: " & sTicketID & vbCrLf _
                                & "Status: " & Status & vbCrLf & vbCrLf _
                                & "SQL to execute: " & SQLString
                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                        End If
                        With DManager
                            ds = .GetDataSet(SQLString)
                            If Not ds Is Nothing Then
                                dt = ds.Tables(0)
                                If dt.Rows.Count > 0 Then
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #26", EventLogEntryType.Information)
                                    sTicketID = dt.Rows(0).Item("TicketID_Display")
                                    Status = dt.Rows(0).Item("Status")
                                End If
                            End If
                            If Debug_Mode Then
                                sEventEntry = "CIM Service: Function ProcessSupportMessage #27" & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                    & "sTicketID: " & sTicketID & vbCrLf _
                                    & "Status: " & Status & vbCrLf & vbCrLf _
                                    & "SQL executed: " & SQLString
                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            End If
                        End With
                    Catch ex As Exception
                        TakeABreak = True
                        sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #28 returns " _
                            & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                            & "Partner Ticket ID: " & PartnerTicketNo & vbCrLf _
                            & "sTicketID: " & sTicketID & vbCrLf _
                            & "Status: " & Status & vbCrLf & vbCrLf _
                            & "SQL: " & SQLString & vbCrLf & vbCrLf _
                            & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                            & "Exception Target: " & ex.TargetSite.ToString()
                        WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                        Exit Function
                    End Try
                    If TakeABreak = True Then Exit Function
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #29", EventLogEntryType.Information)
                    If Len(sTicketID) = 0 Then ' send email to Arnett Kelly and Brian Brazil that an update request was recieved but no ticket was found internally
                        SendUpdateTicketNotFoundEmail(msg, Partner3_, PartnerTicketNo, DataAreaID)
                    Else ' process accordingly
                        NextTicketEventRecID = GetNextTicketEventRecID(TakeABreak)
                        If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #30", EventLogEntryType.Information)
                        If TakeABreak = True Or NextTicketEventRecID <= 0 Then Exit Function
                        If Status = "Closed" And RequestStatus <> "Close" Then ' Re-Open...change value on header, update appropriate fields and add Event entry.
                            Try
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), Status = 'Open', IsRead = 0 WHERE TicketID_Display = '" & sTicketID & "'"
                                    ' " & SQLTicketChangesString & "
                                    'eventViewerString &= vbCrLf & vbCrLf & "Re-Opened Ticket ID '" & TicketID & "' SQL: " & SQLString
                                    'Write("CIM Service - Ticket ID: " & TicketID & " Re-Opened", "SQL: " & SQLString, 0)
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #31" & vbCrLf _
                                            & "Ticket ID: " & sTicketID & " Re-opened" & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #32", EventLogEntryType.Information)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #33 returns " _
                                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #34", EventLogEntryType.Information)
                            Try
                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Ticket Re-Opened', 0, '" & sTicketID & "', 'Ticket Re-Opened due to an update from the partner.  Please review and process accordingly." _
                                        & vbCrLf & vbCrLf & Replace(SQLTicketChangesEventString, "'", "''") & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf _
                                        & Replace(sTicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" _
                                        & SeverityLevel & "', '" & DataAreaID & "', '" & NextTicketEventRecID & "', 1, 1, 0)"
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #35" & vbCrLf _
                                            & "Ticket Event ID: " & sTicketID & vbCrLf _
                                            & "Next Ticket Event Rec ID: " & Convert.ToString(NextTicketEventRecID) & " Re-opened" & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #36 returns " _
                                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                    & "Ticket Event ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf & vbCrLf _
                                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                        ElseIf Status <> "Closed" And RequestStatus = "Close" Then ' add Event entry.  Notify ticket owner that partner has closed the ticket
                            Try
                                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #37", EventLogEntryType.Information)
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), IsRead = 0 WHERE TicketID_Display = '" & sTicketID & "'"
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #38 Ticket ID: " & sTicketID _
                                            & " Receive Closed Ticket per partner's request" & vbCrLf _
                                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #39", EventLogEntryType.Information)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #40 returns " _
                                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                    & "Ticket ID: " & sTicketID & " Receive Closed Ticket per partner's request" & vbCrLf & vbCrLf _
                                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #41", EventLogEntryType.Information)
                            Try
                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Comment', 1, '" & sTicketID & "', 'Please note that it appears that this ticket may have been closed by the partner.  Please review and process accordingly." _
                                        & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(sTicketEmailBody, "'", "''") _
                                        & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel _
                                        & "', '" & DataAreaID & "', '" & NextTicketEventRecID & "', 1, 1, 0)"
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #42" & vbCrLf _
                                            & "Inserting Ticket Event" & vbCrLf _
                                            & "Ticket ID: " & sTicketID & vbCrLf _
                                            & "Next Ticket Event Rec ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf _
                                            & " Receive Closed Ticket per partner's request" & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #43 returns " _
                                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                    & "Ticket Event ID: " & Convert.ToString(NextTicketEventRecID) & " Receive Closed Ticket per partner's request" & vbCrLf & vbCrLf _
                                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #44", EventLogEntryType.Information)
                            REM Close request when our ticket was already closed...probably just a receipt confirmation from our partner.  No notification of ticket owner
                        ElseIf Status = "Closed" And RequestStatus = "Close" Then
                            Try
                                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #45", EventLogEntryType.Information)
                                ' no update needed to ticket header (C_Tickets)
                                'Write("CIM Service - Ticket ID: " & TicketID & " Closed per partners request", "Ticket was already closed internally", 0)
                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Comment', 1, '" & sTicketID & "', 'Please note that this ticket was closed by the partner (internal ticket already closed).  Review not necessary." _
                                        & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(sTicketEmailBody, "'", "''") _
                                        & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel _
                                        & "', '" & DataAreaID _
                                        & "', '" & NextTicketEventRecID & "', 1, 1, 0)"
                                    REM eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event (close request on a closed ticket) SQL: " & SQLString
                                    REM Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #46" & vbCrLf _
                                            & "Closed per partner's request" & vbCrLf _
                                            & "Ticket ID: " & sTicketID & vbCrLf _
                                            & "Next Ticket Event Rec ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf _
                                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                            & "Ticket was already closed internally" & vbCrLf & vbCrLf & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #47", EventLogEntryType.Information)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #48 returns " _
                                        & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                        & "Ticket Event ID: " & Convert.ToString(NextTicketEventRecID) & " Closed per partner's request" & vbCrLf _
                                        & "Ticket was already closed internally" & vbCrLf & vbCrLf _
                                        & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                        & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                        & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #49", EventLogEntryType.Information)
                        Else ' regular update request...insert and update accordingly
                        Try
                                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #50", EventLogEntryType.Information)
                            With DManager
                                    SQLString = "UPDATE C_Tickets SET " _
                                            & "SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), IsRead = 0 " _
                                            & "WHERE TicketID_Display = '" & sTicketID & "'"

                                'eventViewerString &= vbCrLf & vbCrLf & "Received Update Ticket '" & TicketID & "' request, Ticket updated SQL: " & SQLString
                                'Write("CIM Service - Updated Ticket ID: " & TicketID, SQLString, 0)
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #51" & vbCrLf _
                                            & "Receive updated Ticket ID: " & sTicketID & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                End With
                        Catch ex As Exception
                            TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #52 returns " _
                                & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                & "Updated Ticket ID: " & sTicketID & vbCrLf & vbCrLf _
                                & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                & "Exception Target: " & ex.TargetSite.ToString()
                            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                            Exit Function
                        End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #53", EventLogEntryType.Information)
                        Try
                            With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Comment', 1, '" & sTicketID & "', 'Partner Integration Update Email Received.  Please review and process accordingly." _
                                        & vbCrLf & vbCrLf & Replace(SQLTicketChangesEventString, "'", "''") & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf _
                                        & Replace(sTicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription _
                                        & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & NextTicketEventRecID & "', 1, 1, 0)"
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #54 Inserted Ticket Event ID: " _
                                            & Convert.ToString(NextTicketEventRecID) & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #55", EventLogEntryType.Information)
                                End With
                        Catch ex As Exception
                            TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #56 returns " _
                                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                    & "Inserted Ticket Event ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf & vbCrLf _
                                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString()
                            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                            Exit Function
                        End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #57", EventLogEntryType.Information)
                        End If
                    End If
                End If
                ProcessSupportMessage = True    'returns

            Else ' process other partner ticket emails (non integration)
                Try
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #58", EventLogEntryType.Information)
                    ' Get the next ticket id.  Check if the number sequence table needs to be synched up too...
                    sTicketID = GetNextTicketId(DataAreaID, TakeABreak)
                    'If CInt(Replace(TicketID, "AGST", "")) + 1 > GetTicketNumberSequence() Then
                    '    With DManager
                    '        SQLString = "UPDATE NumberSequenceTable SET NextRec = " & CInt(Replace(TicketID, "AGST", "")) + 2 & " Where DataAreaID = '" & DataAreaID & "' and numbersequence = 'CC_Tickets'"
                    '        .ExecuteSQL(SQLString)
                    '    End With
                    'End If
                    If TakeABreak = True Or Len(sTicketID) = 0 Then Exit Function
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #59", EventLogEntryType.Information)
                    If msg.Priority = MessagePriority.High Then
                        SeverityLevel = "5637144578"
                        SeverityLevelDescription = "High"
                    ElseIf msg.Priority = MessagePriority.Low Then
                        SeverityLevel = "5637144580"
                        SeverityLevelDescription = "Informational"
                    Else
                        SeverityLevel = "5637144579"
                        SeverityLevelDescription = "Normal"
                    End If
                    If Debug_Mode Then
                        sEventEntry = "CIM Service: Function ProcessSupportMessage #60" & vbCrLf _
                            & "Ticket ID:" & sTicketID & vbCrLf _
                            & "Severity Level: " & SeverityLevel & vbCrLf _
                            & "Severity Level Description: " & SeverityLevelDescription
                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    End If

                    'DataAreaID = "US"
                    'Partner3_ = "OTH"
                    'TicketQueue = "CellStack"

                    'C_Tickets
                    NextTicketRecID = GetNextTicketRecID(TakeABreak)
                    If Debug_Mode Then
                        sEventEntry = "CIM Service: Function ProcessSupportMessage #61" & vbCrLf _
                            & "Next Ticket Rec ID:" & Convert.ToString(NextTicketRecID)
                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    End If
                    If TakeABreak = True Or NextTicketRecID <= 0 Then Exit Function
                Catch ex As Exception
                    TakeABreak = True
                    sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #62 returns " _
                        & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                        & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                        & "Exception Target: " & ex.TargetSite.ToString()
                    WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                    Exit Function
                End Try
                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #63", EventLogEntryType.Information)

                Try
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #64", EventLogEntryType.Information)
                    With DManager
                        SQLString = "INSERT INTO C_Tickets (" & _
                            " CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, TicketID_Display, Source, Status, TicketQueue, Tier, Partner3_," & _
                            " PartnerTicketSource, UpdatePartner, CallerPhone," & _
                            " CallerEmail, RecID, RecVersion, DataAreaID, UpdateContact, UpdateCaller, Caller, CallerName, severityLevel, CustomerName," & _
                            " PartnerTicketNo, SiteName, Product, SN, ContactName, ContactPhone, ContactEmail, ProblemDescription" & _
                            ")" & _
                            "VALUES (" & _
                            "   'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & sTicketID & "', 'Email', 'Open', '" & _
                            TicketQueue & "', 'Tier 1', '" & Partner3_ & "', 'email', 1, '', " & _
                            "   '" & Replace(msg.From.EMail.ToString(), "'", "''") & "', '" & NextTicketRecID & "', 1, '" & DataAreaID & "', 1, 1, '" & _
                            Replace(msg.From.Name.ToString(), "'", "''") & "', '" & Replace(Mid(msg.From.Name.ToString(), 1, 60), "'", "''") & _
                            "', '" & SeverityLevel & "', '', " & _
                            "   '', '', '', '', '" & Replace(msg.From.Name.ToString(), "'", "''") & "', '', '" & Replace(msg.From.EMail.ToString(), "'", "''") & _
                            "', '" & Replace(msg.PlainMessage.Body.ToString(), "'", "''") & "')"

                        'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' SQL: " & SQLString
                        'Write("CIM Service - Inserted Ticket ID: " & TicketID, "SQL: " & SQLString, 0)
                        If Debug_Mode Or Inform_Mode Then
                            sEventEntry = "CIM Service: Function ProcessSupportMessage #65" & vbCrLf _
                                & "Inserting Ticket ID: " & sTicketID & vbCrLf & vbCrLf & "SQL to execute: " & SQLString
                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                        End If
                        .ExecuteSQL(SQLString)
                    End With
                Catch ex As Exception
                    TakeABreak = True
                    sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #66 returns " _
                        & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                        & "Ticket ID: " & sTicketID & vbCrLf & vbCrLf _
                        & "SQL: " & SQLString & vbCrLf & vbCrLf _
                        & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                        & "Exception Target: " & ex.TargetSite.ToString()
                    WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                    Exit Function
                End Try

                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #67", EventLogEntryType.Information)
                IncrementTicketNumberSequence(TakeABreak)
                If TakeABreak = True Then Exit Function
                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #68", EventLogEntryType.Information)

                'C_TicketEvents
                REM Cast a long number to an integer 
                NextTicketEventRecID = GetNextTicketEventRecID(TakeABreak)
                If TakeABreak = True Or NextTicketEventRecID <= 0 Then Exit Function
                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #69", EventLogEntryType.Information)

                Try
                    With DManager
                        SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " & _
                            "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " & _
                            "VALUES ('New Ticket', 0, '" & sTicketID & "', 'New " & Partner3_ & " Integration Email received: " & vbCrLf & vbCrLf & _
                            "---- Original message ----" & vbCrLf & vbCrLf & Replace(sTicketEmailBody, "'", "''") & _
                            "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & _
                            SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & NextTicketEventRecID & "', 1, 1, 0)"
                        'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                        'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                        If Debug_Mode Or Inform_Mode Then
                            sEventEntry = "CIM Service: Function ProcessSupportMessage #70" & vbCrLf _
                                & "Ticket ID: " & sTicketID & vbCrLf _
                                & "Next Ticket Events Rec ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf _
                                & vbCrLf & vbCrLf & "SQL to execute: " & SQLString
                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                        End If
                        .ExecuteSQL(SQLString)
                        If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #71", EventLogEntryType.Information)
                    End With
                Catch ex As Exception
                    TakeABreak = True
                    sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #72 returns " _
                        & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                        & "Ticket Events ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf & vbCrLf _
                        & "SQL: " & SQLString & vbCrLf & vbCrLf _
                        & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                        & "Exception Target: " & ex.TargetSite.ToString()
                    WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                    Exit Function
                End Try

                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #73", EventLogEntryType.Information)
                'Write("CIM Service - received ticket integration email", eventViewerString, 0)
                'WriteToEventLog("CIM Service - Received ticket integration email" & vbCrLf & "Event Viewer String: " & eventViewerString, EventLogEntryType.Information)
                ProcessSupportMessage = True    'returns
            End If
        Else
            ProcessSupportMessage = True    'returns
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function ProcessSupportMessage #74, returns" _
                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                    & "Email received that was not an integration request" & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                    & "From: " & msg.From.ToString() & vbCrLf _
                    & "To: " & msg.To.ToString() & vbCrLf _
                    & "Subject: " & msg.Subject.ToString() & vbCrLf _
                    & "Body: " & vbCrLf & "-- beginning of line --" & vbCrLf _
                    & sTicketEmailBody & vbCrLf _
                    & "-- end of line --"
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        End If
    End Function

    Sub ParseVerintOnyxEmail(ByVal TicketEmailBody As String, ByVal MsgID As String, ByVal IntegrationRequestType As String, ByRef TakeABreak As Boolean)
        Dim Priority As String = "", PriorityStart As Integer = 0, PriorityEnd As Integer = 0, PriorityDescription As String = ""
        Dim CompanyName As String = "", CompanyNameStart As Integer = 0, CompanyNameEnd As Integer = 0, CompanyNameEndA As Integer = 0, CompanyNameEndB As Integer = 0
        Dim SiteNameStart As Integer = 0, SiteNameEnd As Integer = 0
        Dim OnyxCustomerNumber As String = "", OnyxCustomerNumberStart As Integer = 0, OnyxCustomerNumberEnd As Integer = 0
        Dim OnyxIncidentNumber As String = "", OnyxIncidentNumberStart As Integer = 0, OnyxIncidentNumberEnd As Integer = 0
        Dim CustomerPrimaryContact As String = "", CustomerPrimaryContactStart As Integer = 0, CustomerPrimaryContactEnd As Integer = 0
        Dim Location As String = "", LocationStart As Integer = 0, LocationEnd As Integer = 0
        Dim PostalCode As String = "", PostalCodeStart As Integer = 0, PostalCodeEnd As Integer = 0
        Dim IncidentContact As String = "", IncidentContactStart As Integer = 0, IncidentContactEnd As Integer = 0
        Dim BusinessPhone As String = "", BusinessPhoneStart As Integer = 0, BusinessPhoneEnd As Integer = 0
        Dim CellPhone As String = "", CellPhoneStart As Integer = 0, CellPhoneEnd As Integer = 0
        Dim Fax As String = "", FaxStart As Integer = 0, FaxEnd As Integer = 0
        Dim Pager As String = "", PagerStart As Integer = 0, PagerEnd As Integer = 0
        Dim EmailAddress As String = "", EmailAddressStart As Integer = 0, EmailAddressEnd As Integer = 0
        Dim SupportLevel As String = "", SupportLevelStart As Integer = 0, SupportLevelEnd As Integer = 0
        Dim IncidentType As String = "", IncidentTypeStart As Integer = 0, IncidentTypeEnd As Integer = 0
        Dim ProductStart As Integer = 0, ProductEnd As Integer = 0 ' Product As String = "", 
        Dim Version As String = "", VersionStart As Integer = 0, VersionEnd As Integer = 0
        Dim SerialNumber As String = "", SerialNumberStart As Integer = 0, SerialNumberEnd As Integer = 0
        Dim SubmittedBy As String = "", SubmittedByStart As Integer = 0, SubmittedByEnd As Integer = 0
        Dim WorkNotes As String = "", WorkNotesStart As Integer = 0, WorkNotesEnd As Integer = 0

        If Debug_Mode Then
            sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #01" & vbCrLf _
                & "Msg ID: " & MsgID & vbCrLf _
                & "Integration Request Type: " & IntegrationRequestType & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "Ticket Email Body:" & vbCrLf & vbTab & "---- Beginnning of lines ----" & vbCrLf _
                & TicketEmailBody & vbCrLf & vbTab & "---- End of lines ----"
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        End If

        'Parse message body...Verint's emails are formatted the same regardless of them being a new or an update request

        'Priority:  4 - Low
        Try
            PriorityStart = InStr(TicketEmailBody, "Priority: ") + 9
            If (PriorityStart > 9) Then
                PriorityEnd = FindMin(InStr(PriorityStart, TicketEmailBody, vbCr), _
                    InStr(PriorityStart, TicketEmailBody, vbLf), _
                    InStr(PriorityStart, TicketEmailBody, vbCrLf), _
                    InStr(PriorityStart, TicketEmailBody, vbNewLine))
                If (PriorityEnd > 0) Then
                    Priority = Trim(Mid(TicketEmailBody, PriorityStart, PriorityEnd - PriorityStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #02" & vbCrLf & "Priority: " & Priority, EventLogEntryType.Information)

            'Company Name:  Hartford, The (Server #1) - Charlotte, NC
            CompanyNameStart = InStr(TicketEmailBody, "Company Name: ") + 13
            If (CompanyNameStart > 13) Then
                ' if they don't submit the customer name and site name together (no " - " then), we have to handle it...
                CompanyNameEndA = InStr(CompanyNameStart, TicketEmailBody, " - ")
                'CompanyNameEndB = InStr(CompanyNameStart, TicketEmailBody, vbCrLf)
                CompanyNameEndB = FindMin(InStr(CompanyNameStart, TicketEmailBody, vbCr), _
                    InStr(CompanyNameStart, TicketEmailBody, vbLf), _
                    InStr(CompanyNameStart, TicketEmailBody, vbCrLf), _
                    InStr(CompanyNameStart, TicketEmailBody, vbNewLine))
                If (CompanyNameEndA < CompanyNameEndB) Then ' use the lesser one...ie if there was a return before the split, we only want to go to the return
                    CompanyNameEnd = CompanyNameEndA
                Else
                    CompanyNameEnd = CompanyNameEndB
                End If
                If (CompanyNameEnd > 0) Then
                    CompanyName = Trim(Mid(TicketEmailBody, CompanyNameStart, CompanyNameEnd - CompanyNameStart))

                    If (CompanyNameEndA < CompanyNameEndB) Then ' site name will only be there if the split " - " existed
                        SiteNameStart = CompanyNameEnd + 3
                        If (SiteNameStart > 16) Then
                            'SiteNameEnd = InStr(SiteNameStart, TicketEmailBody, vbCrLf)
                            SiteNameEnd = FindMin(InStr(SiteNameStart, TicketEmailBody, vbCr), _
                                InStr(SiteNameStart, TicketEmailBody, vbLf), _
                                InStr(SiteNameStart, TicketEmailBody, vbCrLf), _
                                InStr(SiteNameStart, TicketEmailBody, vbNewLine))
                            If (SiteNameEnd > 0) Then
                                SiteName = Trim(Mid(TicketEmailBody, SiteNameStart, SiteNameEnd - SiteNameStart))
                            End If
                        End If
                    End If
                End If
            End If
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #03" & vbCrLf _
                    & "CompanyName: " & CompanyName & vbCrLf _
                    & "SiteName: " & SiteName
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Onyx Customer Number:  159048
            OnyxCustomerNumberStart = InStr(TicketEmailBody, "Onyx Customer Number: ") + 21
            If (OnyxCustomerNumberStart > 21) Then
                'OnyxCustomerNumberEnd = InStr(OnyxCustomerNumberStart, TicketEmailBody, vbCrLf)
                OnyxCustomerNumberEnd = FindMin(InStr(OnyxCustomerNumberStart, TicketEmailBody, vbCr), _
                    InStr(OnyxCustomerNumberStart, TicketEmailBody, vbLf), _
                    InStr(OnyxCustomerNumberStart, TicketEmailBody, vbCrLf), _
                    InStr(OnyxCustomerNumberStart, TicketEmailBody, vbNewLine))
                If (OnyxCustomerNumberEnd > 0) Then
                    OnyxCustomerNumber = Trim(Mid(TicketEmailBody, OnyxCustomerNumberStart, OnyxCustomerNumberEnd - OnyxCustomerNumberStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #04" & vbCrLf & "OnyxCustomerNumber: " & OnyxCustomerNumber, EventLogEntryType.Information)

            'Onyx Incident Number:  3810615
            OnyxIncidentNumberStart = InStr(TicketEmailBody, "Onyx Incident Number: ") + 21
            If (OnyxIncidentNumberStart > 21) Then
                'OnyxIncidentNumberEnd = InStr(OnyxIncidentNumberStart, TicketEmailBody, vbCrLf)
                OnyxIncidentNumberEnd = FindMin(InStr(OnyxIncidentNumberStart, TicketEmailBody, vbCr), _
                    InStr(OnyxIncidentNumberStart, TicketEmailBody, vbLf), _
                    InStr(OnyxIncidentNumberStart, TicketEmailBody, vbCrLf), _
                    InStr(OnyxIncidentNumberStart, TicketEmailBody, vbNewLine))
                If (OnyxIncidentNumberEnd > 0) Then
                    OnyxIncidentNumber = Trim(Mid(TicketEmailBody, OnyxIncidentNumberStart, OnyxIncidentNumberEnd - OnyxIncidentNumberStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #05" & vbCrLf & "OnyxIncidentNumber: " & OnyxIncidentNumber, EventLogEntryType.Information)

            'Customer Primary Contact:  Frank Beatty
            CustomerPrimaryContactStart = InStr(TicketEmailBody, "Customer Primary Contact: ") + 25
            If (CustomerPrimaryContactStart > 25) Then
                'CustomerPrimaryContactEnd = InStr(CustomerPrimaryContactStart, TicketEmailBody, vbCrLf)
                CustomerPrimaryContactEnd = FindMin(InStr(CustomerPrimaryContactStart, TicketEmailBody, vbCr), _
                    InStr(CustomerPrimaryContactStart, TicketEmailBody, vbLf), _
                    InStr(CustomerPrimaryContactStart, TicketEmailBody, vbCrLf), _
                    InStr(CustomerPrimaryContactStart, TicketEmailBody, vbNewLine))
                If (CustomerPrimaryContactEnd > 0) Then
                    CustomerPrimaryContact = Trim(Mid(TicketEmailBody, CustomerPrimaryContactStart, CustomerPrimaryContactEnd - CustomerPrimaryContactStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #06" & vbCrLf & "CustomerPrimaryContact: " & CustomerPrimaryContact, EventLogEntryType.Information)

            'Location:  Charlotte Service Ctr, 8711 University East Dr, Charlotte, North Carolina, United States
            LocationStart = InStr(TicketEmailBody, "Location: ") + 9
            If (LocationStart > 9) Then
                'LocationEnd = InStr(LocationStart, TicketEmailBody, vbCrLf)
                LocationEnd = FindMin(InStr(LocationStart, TicketEmailBody, vbCr), _
                    InStr(LocationStart, TicketEmailBody, vbLf), _
                    InStr(LocationStart, TicketEmailBody, vbCrLf), _
                    InStr(LocationStart, TicketEmailBody, vbNewLine))
                If (LocationEnd > 0) Then
                    Location = Trim(Mid(TicketEmailBody, LocationStart, LocationEnd - LocationStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #07" & vbCrLf & "Location: " & Location, EventLogEntryType.Information)

            'Postal Code:  28213
            PostalCodeStart = InStr(TicketEmailBody, "Postal Code: ") + 12
            If (PostalCodeStart > 12) Then
                'PostalCodeEnd = InStr(PostalCodeStart, TicketEmailBody, vbCrLf)
                PostalCodeEnd = FindMin(InStr(PostalCodeStart, TicketEmailBody, vbCr), _
                    InStr(PostalCodeStart, TicketEmailBody, vbLf), _
                    InStr(PostalCodeStart, TicketEmailBody, vbCrLf), _
                    InStr(PostalCodeStart, TicketEmailBody, vbNewLine))
                If (PostalCodeEnd > 0) Then
                    PostalCode = Trim(Mid(TicketEmailBody, PostalCodeStart, PostalCodeEnd - PostalCodeStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #08" & vbCrLf & "PostalCode: " & PostalCode, EventLogEntryType.Information)

            'Incident Contact:  Steven Hudak
            IncidentContactStart = InStr(TicketEmailBody, "Incident Contact: ") + 17
            If (IncidentContactStart > 17) Then
                'IncidentContactEnd = InStr(IncidentContactStart, TicketEmailBody, vbCrLf)
                IncidentContactEnd = FindMin(InStr(IncidentContactStart, TicketEmailBody, vbCr), _
                    InStr(IncidentContactStart, TicketEmailBody, vbLf), _
                    InStr(IncidentContactStart, TicketEmailBody, vbCrLf), _
                    InStr(IncidentContactStart, TicketEmailBody, vbNewLine))
                If (IncidentContactEnd > 0) Then
                    IncidentContact = Trim(Mid(TicketEmailBody, IncidentContactStart, IncidentContactEnd - IncidentContactStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #09" & vbCrLf & "IncidentContact: " & IncidentContact, EventLogEntryType.Information)

            'Business Phone:  8609197122
            BusinessPhoneStart = InStr(TicketEmailBody, "Business Phone: ") + 15
            If (BusinessPhoneStart > 15) Then
                'BusinessPhoneEnd = InStr(BusinessPhoneStart, TicketEmailBody, vbCrLf)
                BusinessPhoneEnd = FindMin(InStr(BusinessPhoneStart, TicketEmailBody, vbCr), _
                    InStr(BusinessPhoneStart, TicketEmailBody, vbLf), _
                    InStr(BusinessPhoneStart, TicketEmailBody, vbCrLf), _
                    InStr(BusinessPhoneStart, TicketEmailBody, vbNewLine))
                If (BusinessPhoneEnd > 0) Then
                    BusinessPhone = Trim(Mid(TicketEmailBody, BusinessPhoneStart, BusinessPhoneEnd - BusinessPhoneStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #10" & vbCrLf & "BusinessPhone: " & BusinessPhone, EventLogEntryType.Information)

            'Cell Phone:  8609197122
            CellPhoneStart = InStr(TicketEmailBody, "Cell Phone: ") + 11
            If (CellPhoneStart > 11) Then
                'CellPhoneEnd = InStr(CellPhoneStart, TicketEmailBody, vbCrLf)
                CellPhoneEnd = FindMin(InStr(CellPhoneStart, TicketEmailBody, vbCr), _
                    InStr(CellPhoneStart, TicketEmailBody, vbLf), _
                    InStr(CellPhoneStart, TicketEmailBody, vbCrLf), _
                    InStr(CellPhoneStart, TicketEmailBody, vbNewLine))
                If (CellPhoneEnd > 0) Then
                    CellPhone = Trim(Mid(TicketEmailBody, CellPhoneStart, CellPhoneEnd - CellPhoneStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #11" & vbCrLf & "CellPhone: " & CellPhone, EventLogEntryType.Information)

            'Fax: 8609197122
            FaxStart = InStr(TicketEmailBody, "Fax: ") + 4
            If (FaxStart > 4) Then
                'FaxEnd = InStr(FaxStart, TicketEmailBody, vbCrLf)
                FaxEnd = FindMin(InStr(FaxStart, TicketEmailBody, vbCr), _
                    InStr(FaxStart, TicketEmailBody, vbLf), _
                    InStr(FaxStart, TicketEmailBody, vbCrLf), _
                    InStr(FaxStart, TicketEmailBody, vbNewLine))
                If (FaxEnd > 0) Then
                    Fax = Trim(Mid(TicketEmailBody, FaxStart, FaxEnd - FaxStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #12" & vbCrLf & "Fax: " & Fax, EventLogEntryType.Information)

            'Pager: 8609197122
            PagerStart = InStr(TicketEmailBody, "Pager: ") + 6
            If (PagerStart > 6) Then
                'PagerEnd = InStr(PagerStart, TicketEmailBody, vbCrLf)
                PagerEnd = FindMin(InStr(PagerStart, TicketEmailBody, vbCr), _
                    InStr(PagerStart, TicketEmailBody, vbLf), _
                    InStr(PagerStart, TicketEmailBody, vbCrLf), _
                    InStr(PagerStart, TicketEmailBody, vbNewLine))
                If (PagerEnd > 0) Then
                    Pager = Trim(Mid(TicketEmailBody, PagerStart, PagerEnd - PagerStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #13" & vbCrLf & "Pager: " & Pager, EventLogEntryType.Information)

            'Email Address:  steven.hudak@thehartford.com
            EmailAddressStart = InStr(TicketEmailBody, "Email Address: ") + 14
            If (EmailAddressStart > 14) Then
                'EmailAddressEnd = InStr(EmailAddressStart, TicketEmailBody, vbCrLf)
                EmailAddressEnd = FindMin(InStr(EmailAddressStart, TicketEmailBody, vbCr), _
                    InStr(EmailAddressStart, TicketEmailBody, vbLf), _
                    InStr(EmailAddressStart, TicketEmailBody, vbCrLf), _
                    InStr(EmailAddressStart, TicketEmailBody, vbNewLine))
                If (EmailAddressEnd > 0) Then
                    EmailAddress = Trim(Mid(TicketEmailBody, EmailAddressStart, EmailAddressEnd - EmailAddressStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #14" & vbCrLf & "EmailAddress: " & EmailAddress, EventLogEntryType.Information)

            'Support Level:  Advantage
            SupportLevelStart = InStr(TicketEmailBody, "Support Level: ") + 14
            If (SupportLevelStart > 14) Then
                'SupportLevelEnd = InStr(SupportLevelStart, TicketEmailBody, vbCrLf)
                SupportLevelEnd = FindMin(InStr(SupportLevelStart, TicketEmailBody, vbCr), _
                    InStr(SupportLevelStart, TicketEmailBody, vbLf), _
                    InStr(SupportLevelStart, TicketEmailBody, vbCrLf), _
                    InStr(SupportLevelStart, TicketEmailBody, vbNewLine))
                If (SupportLevelEnd > 0) Then
                    SupportLevel = Trim(Mid(TicketEmailBody, SupportLevelStart, SupportLevelEnd - SupportLevelStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #15" & vbCrLf & "SupportLevel: " & SupportLevel, EventLogEntryType.Information)

            'Incident Type:  Web - Unknown
            IncidentTypeStart = InStr(TicketEmailBody, "Incident Type: ") + 14
            If (IncidentTypeStart > 14) Then
                'IncidentTypeEnd = InStr(IncidentTypeStart, TicketEmailBody, vbCrLf)
                IncidentTypeEnd = FindMin(InStr(IncidentTypeStart, TicketEmailBody, vbCr), _
                    InStr(IncidentTypeStart, TicketEmailBody, vbLf), _
                    InStr(IncidentTypeStart, TicketEmailBody, vbCrLf), _
                    InStr(IncidentTypeStart, TicketEmailBody, vbNewLine))
                If (IncidentTypeEnd > 0) Then
                    IncidentType = Trim(Mid(TicketEmailBody, IncidentTypeStart, IncidentTypeEnd - IncidentTypeStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #16" & vbCrLf & "IncidentType: " & IncidentType, EventLogEntryType.Information)

            'Product:  eQContactStore
            ProductStart = InStr(TicketEmailBody, "Product: ") + 8
            If (ProductStart > 8) Then
                'ProductEnd = InStr(ProductStart, TicketEmailBody, vbCrLf)
                ProductEnd = FindMin(InStr(ProductStart, TicketEmailBody, vbCr), _
                    InStr(ProductStart, TicketEmailBody, vbLf), _
                    InStr(ProductStart, TicketEmailBody, vbCrLf), _
                    InStr(ProductStart, TicketEmailBody, vbNewLine))
                If (ProductEnd > 0) Then
                    Product = Trim(Mid(TicketEmailBody, ProductStart, ProductEnd - ProductStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #17" & vbCrLf & "Product: " & Product, EventLogEntryType.Information)

            'Version:  CS - 7.2
            VersionStart = InStr(TicketEmailBody, "Version: ") + 8
            If (VersionStart > 8) Then
                'VersionEnd = InStr(VersionStart, TicketEmailBody, vbCrLf)
                VersionEnd = FindMin(InStr(VersionStart, TicketEmailBody, vbCr), _
                    InStr(VersionStart, TicketEmailBody, vbLf), _
                    InStr(VersionStart, TicketEmailBody, vbCrLf), _
                    InStr(VersionStart, TicketEmailBody, vbNewLine))
                If (VersionEnd > 0) Then
                    Version = Trim(Mid(TicketEmailBody, VersionStart, VersionEnd - VersionStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #18" & vbCrLf & "Version: " & Version, EventLogEntryType.Information)

            'Serial Number: RN-XXXXXX
            SerialNumberStart = InStr(TicketEmailBody, "Serial Number: ") + 14
            If (SerialNumberStart > 14) Then
                'SerialNumberEnd = InStr(SerialNumberStart, TicketEmailBody, vbCrLf)
                SerialNumberEnd = FindMin(InStr(SerialNumberStart, TicketEmailBody, vbCr), _
                    InStr(SerialNumberStart, TicketEmailBody, vbLf), _
                    InStr(SerialNumberStart, TicketEmailBody, vbCrLf), _
                    InStr(SerialNumberStart, TicketEmailBody, vbNewLine))
                If (SerialNumberEnd > 0) Then
                    SerialNumber = Trim(Mid(TicketEmailBody, SerialNumberStart, SerialNumberEnd - SerialNumberStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #19" & vbCrLf & "SerialNumber: " & SerialNumber, EventLogEntryType.Information)

            'Submitted By:  Janice Wells
            SubmittedByStart = InStr(TicketEmailBody, "Submitted By: ") + 13
            If (SubmittedByStart > 13) Then
                'SubmittedByEnd = InStr(SubmittedByStart, TicketEmailBody, vbCrLf)
                SubmittedByEnd = FindMin(InStr(SubmittedByStart, TicketEmailBody, vbCr), _
                    InStr(SubmittedByStart, TicketEmailBody, vbLf), _
                    InStr(SubmittedByStart, TicketEmailBody, vbCrLf), _
                    InStr(SubmittedByStart, TicketEmailBody, vbNewLine))
                If (SubmittedByEnd > 0) Then
                    SubmittedBy = Trim(Mid(TicketEmailBody, SubmittedByStart, SubmittedByEnd - SubmittedByStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #20" & vbCrLf & "SubmittedBy: " & SubmittedBy, EventLogEntryType.Information)

            'WorkNotes: 
            WorkNotesStart = InStr(TicketEmailBody, "WorkNotes: ") + 10
            If (WorkNotesStart > 10) Then
                WorkNotesEnd = Len(TicketEmailBody)
                If (WorkNotesEnd > 0) Then
                    'Trim occurs
                    WorkNotes = Trim(Mid(TicketEmailBody, WorkNotesStart, WorkNotesEnd - WorkNotesStart))
                    ' remove up to the first 3 carriage returns so that the problem description will show!
                    'If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    '    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                    'End If
                    'If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    '    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                    'End If
                    'If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    '    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                    'End If
                    Dim HasLeadingWhite As Boolean = True
                    Do While HasLeadingWhite
                        If Left(WorkNotes, Len(vbCr)) = vbCr Then
                            WorkNotes = Mid(WorkNotes, Len(vbCr) + 1)
                        ElseIf Left(WorkNotes, Len(vbLf)) = vbLf Then
                            WorkNotes = Mid(WorkNotes, Len(vbLf) + 1)
                        ElseIf Left(WorkNotes, Len(vbCrLf)) = vbCrLf Then
                            WorkNotes = Mid(WorkNotes, Len(vbCrLf) + 1)
                        ElseIf Left(WorkNotes, Len(vbNewLine)) = vbNewLine Then
                            WorkNotes = Mid(WorkNotes, Len(vbNewLine) + 1)
                        Else
                            HasLeadingWhite = False
                        End If
                    Loop
                End If
            End If
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #21" & vbCrLf & "WorkNotes:" _
                    & vbCrLf & vbTab & "---- Beginning of lines ----" & vbCrLf _
                    & WorkNotes & vbCrLf & vbTab & "---- End of lines ----"
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            If (InStr(UCase(Priority), "CRITICAL") > 0) Then
                SeverityLevel = "5637144577"
                SeverityLevelDescription = "Critical"
            ElseIf (InStr(UCase(Priority), "HIGH") > 0) Then
                SeverityLevel = "5637144578"
                SeverityLevelDescription = "High"
            ElseIf (InStr(UCase(Priority), "MEDIUM") > 0) Then
                SeverityLevel = "5637144579"
                SeverityLevelDescription = "Normal"
            ElseIf (InStr(UCase(Priority), "LOW") > 0) Then
                SeverityLevel = "5637144580"
                SeverityLevelDescription = "Informational"
            Else
                SeverityLevel = "5637144576"
                SeverityLevelDescription = "Compliance"
            End If

            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #22" _
                    & vbCrLf & "SeverityLevel:" & SeverityLevel _
                    & vbCrLf & "SeverityLevelDescription:" & SeverityLevelDescription
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
            CustomerName = CompanyName
            PartnerTicketNo = OnyxIncidentNumber
            SN = SerialNumber
            ContactName = IncidentContact
            ContactPhone = BusinessPhone
            ContactEmail = EmailAddress
            ProblemDescription = WorkNotes
            If InStr(UCase(WorkNotes), "CLOSING") > 0 Or InStr(UCase(WorkNotes), "CLOSED") > 0 Then
                RequestStatus = "Close"
            End If

            If Debug_Mode Or Inform_Mode Then
                'Write("CIM Service - Parsed an Onyx integration email Message (ID: " & MsgID & ")", "Priority: '" & Priority & "'" & vbCrLf & "Company Name: '" & CompanyName & "'" & vbCrLf & "Site Name: '" & SiteName & "'" & vbCrLf & "Onyx Customer Number: '" & OnyxCustomerNumber & "'" & vbCrLf & "Onyx Incident Number: '" & OnyxIncidentNumber & "'" & vbCrLf & "Customer Primary Contact: '" & CustomerPrimaryContact & "'" & vbCrLf & "Location: '" & Location & "'" & vbCrLf & "Postal Code: '" & PostalCode & "'" & vbCrLf & "Incident Contact: '" & IncidentContact & "'" & vbCrLf & "Business Phone: '" & BusinessPhone & "'" & vbCrLf & "Cell Phone: '" & CellPhone & "'" & vbCrLf & "Fax: '" & Fax & "'" & vbCrLf & "Pager: '" & Pager & "'" & vbCrLf & "Email Address: '" & EmailAddress & "'" & vbCrLf & "Support Level: '" & SupportLevel & "'" & vbCrLf & "Incident Type: '" & IncidentType & "'" & vbCrLf & "Product: '" & Product & "'" & vbCrLf & "Version: '" & Version & "'" & vbCrLf & "Serial Number: '" & SerialNumber & "'" & vbCrLf & "Submitted By: '" & SubmittedBy & "'" & vbCrLf & "WorkNotes: '" & WorkNotes & "'", 0)
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail *Summary* #23" & vbCrLf _
                    & "Message ID: (" & MsgID & ")" & vbCrLf _
                    & "Priority: '" & Priority & "'" & vbCrLf _
                    & "Company Name: '" & CompanyName & "'" & vbCrLf _
                    & "Site Name: '" & SiteName & "'" & vbCrLf _
                    & "Onyx Customer Number: " & OnyxCustomerNumber & vbCrLf _
                    & "Onyx Incident Number: " & OnyxIncidentNumber & vbCrLf _
                    & "Customer Primary Contact: " & CustomerPrimaryContact & vbCrLf _
                    & "Location: " & Location & vbCrLf _
                    & "Postal Code: " & PostalCode & vbCrLf _
                    & "Incident Contact: " & IncidentContact & vbCrLf _
                    & "Business Phone: " & BusinessPhone & vbCrLf _
                    & "Cell Phone: " & CellPhone & vbCrLf _
                    & "Fax: " & Fax & vbCrLf _
                    & "Pager: " & Pager & vbCrLf _
                    & "Email Address: " & EmailAddress & vbCrLf _
                    & "Support Level: " & SupportLevel & vbCrLf _
                    & "Incident Type: " & IncidentType & vbCrLf _
                    & "Product: " & Product & vbCrLf _
                    & "Version: " & Version & vbCrLf _
                    & "Serial Number: " & SerialNumber & vbCrLf _
                    & "Submitted By: '" & SubmittedBy & vbCrLf _
                    & "Severity Level: " & SeverityLevel & vbCrLf _
                    & "Severity Level Description: " & SeverityLevelDescription & vbCrLf _
                    & "Request Status: " & RequestStatus & vbCrLf _
                    & "WorkNotes: " & vbCrLf & vbTab & "---- Beginning of lines ----" & vbCrLf _
                    & WorkNotes & vbCrLf & vbTab & "---- End of lines ----"
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Sub ParseVerintOnyxEmail #24" & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Sub
        End Try

        Dim SQLString As String = ""

        If IntegrationRequestType <> "NEW" Then

            ' Get old stuff to compare...need to compare against what was originally submitted by the Partner
            Dim ds As DataSet = New DataSet, dt As DataTable = New DataTable
            Dim OldCustomerName As String = "", OldStatus As String = "", OldSiteName As String = "", OldSN As String = "", Old As String = "", OldSeverity_Priority As String = "", OldProduct As String = "", OldCallerName As String = ""
            Dim OldCallerPhone As String = "", OldCallerEmail As String = "", OldContactName As String = "", OldContactPhone As String = "", OldContactEmail As String = ""

            SQLString = "SELECT TOP 1 * FROM PartnerIntegrationEmailData WHERE PartnerTicketNo = '" & PartnerTicketNo & "' ORDER BY LastChange DESC"
            Try
                If Debug_Mode Then
                    sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #25" & vbCrLf & vbCrLf & "SQL to execute: " & SQLString
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
                With DManager
                    ds = .GetDataSet(SQLString)
                    If Not ds Is Nothing Then
                        dt = ds.Tables(0)
                        If dt.Rows.Count > 0 Then
                            OldCustomerName = dt.Rows(0).Item("CustomerName")
                            OldSiteName = dt.Rows(0).Item("SiteName")
                            OldSN = dt.Rows(0).Item("SN")
                            OldSeverity_Priority = dt.Rows(0).Item("Severity_Priority")
                            OldProduct = dt.Rows(0).Item("Product")
                            OldCallerName = dt.Rows(0).Item("CallerName")
                            OldCallerPhone = dt.Rows(0).Item("CallerPhone")
                            OldCallerEmail = dt.Rows(0).Item("CallerEmail")
                            OldContactName = dt.Rows(0).Item("ContactName")
                            OldContactPhone = dt.Rows(0).Item("ContactPhone")
                            OldContactEmail = dt.Rows(0).Item("ContactEmail")
                        End If
                    End If
                End With
            Catch ex As Exception
                TakeABreak = True
                sEventEntry = "CIM Service Exception: Sub ParseVerintOnyxEmail #26" & vbCrLf & vbCrLf _
                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                    & "Exception Target: " & ex.TargetSite.ToString()
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                Exit Sub
            End Try

            Try
                If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOnyxEmail #27" & vbCrLf & "IF..THEN block begins", EventLogEntryType.Information)
                If (OldCustomerName <> CustomerName) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Customer Name" & vbCrLf & "OLD: " & OldCustomerName & vbCrLf & "NEW: " & CustomerName
                    SQLTicketChangesString = SQLTicketChangesString & ", CustomerName = '" & CustomerName & "'"
                    SQLTicketChangesString = SQLTicketChangesString & ", Customer = ''"
                End If
                If (OldSiteName <> SiteName) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Site Name" & vbCrLf & "OLD: " & OldSiteName & vbCrLf & "NEW: " & SiteName
                    SQLTicketChangesString = SQLTicketChangesString & ", SiteName = '" & SiteName & "'"
                    SQLTicketChangesString = SQLTicketChangesString & ", Sites = ''"
                End If
                If (OldSN <> SN) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Serial Number" & vbCrLf & "OLD: " & OldSN & vbCrLf & "NEW: " & SN
                    SQLTicketChangesString = SQLTicketChangesString & ", SN = '" & SN & "'"
                End If
                If (OldSeverity_Priority <> Priority) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Priority" & vbCrLf & "OLD: " & OldSeverity_Priority & vbCrLf & "NEW: " & Priority
                    SQLTicketChangesString = SQLTicketChangesString & ", SeverityLevel = '" & SeverityLevel & "'"
                    SQLTicketChangesString = SQLTicketChangesString & ", SeverityLevelDescription = '" & SeverityLevelDescription & "'"
                End If
                If (OldProduct <> Product) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Product" & vbCrLf & "OLD: " & OldProduct & vbCrLf & "NEW: " & Product
                    SQLTicketChangesString = SQLTicketChangesString & ", Product = '" & Product & "'"
                End If
                If (OldCallerName <> CallerName) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Caller Name" & vbCrLf & "OLD: " & OldCallerName & vbCrLf & "NEW: " & CallerName
                    SQLTicketChangesString = SQLTicketChangesString & ", CallerName = '" & CallerName & "'"
                    SQLTicketChangesString = SQLTicketChangesString & ", Caller = ''"
                End If
                If (OldCallerPhone <> CallerPhone) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Caller Phone" & vbCrLf & "OLD: " & OldCallerPhone & vbCrLf & "NEW: " & CallerPhone
                    SQLTicketChangesString = SQLTicketChangesString & ", CallerPhone = '" & CallerPhone & "'"
                End If
                If (OldCallerEmail <> CallerEmail) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Caller Email" & vbCrLf & "OLD: " & OldCallerEmail & vbCrLf & "NEW: " & CallerEmail
                    SQLTicketChangesString = SQLTicketChangesString & ", CallerEmail = '" & CallerEmail & "'"
                End If
                If (OldContactName <> ContactName) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Contact Name" & vbCrLf & "OLD: " & OldContactName & vbCrLf & "NEW: " & ContactName
                    SQLTicketChangesString = SQLTicketChangesString & ", ContactName = '" & ContactName & "'"
                    SQLTicketChangesString = SQLTicketChangesString & ", Contact = ''"
                End If
                If (OldContactPhone <> ContactPhone) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Contact Phone" & vbCrLf & "OLD: " & OldContactPhone & vbCrLf & "NEW: " & ContactPhone
                    SQLTicketChangesString = SQLTicketChangesString & ", ContactPhone = '" & ContactPhone & "'"
                End If
                If (OldContactEmail <> ContactEmail) Then
                    SQLTicketChangesEventString = SQLTicketChangesEventString & vbCrLf & vbCrLf & "Contact Email" & vbCrLf & "OLD: " & OldContactEmail & vbCrLf & "NEW: " & ContactEmail
                    SQLTicketChangesString = SQLTicketChangesString & ", ContactEmail = '" & ContactEmail & "'"
                End If

                If SQLTicketChangesEventString <> "" Then
                    SQLTicketChangesEventString = "The following fields were changed from their originally submitted values: " & vbCrLf & vbCrLf & SQLTicketChangesEventString
                End If
            Catch ex As Exception
                TakeABreak = True
                sEventEntry = "CIM Service Exception: Sub ParseVerintOnyxEmail #28" & vbCrLf _
                    & "SQLTicketChangesString: " & SQLTicketChangesString & vbCrLf _
                    & "SQLTicketChangesEventString: " & SQLTicketChangesEventString & vbCrLf _
                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                    & "Exception Target: " & ex.TargetSite.ToString()
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                Exit Sub
            End Try
        End If

        Try
            With DManager
                SQLString = "INSERT INTO PartnerIntegrationEmailData (" & _
                            " PartnerTicketNo, CustomerName, Status, SiteName, SN, Severity_Priority," & _
                            " Product, CallerName, CallerPhone, CallerEmail, ContactName, ContactPhone, ContactEmail, LastChange" & _
                            ") VALUES (" & _
                            "   '" & Replace(PartnerTicketNo, "'", "''") & "', '" & Replace(CustomerName, "'", "''") & "', '', '" & _
                            Replace(SiteName, "'", "''") & "', '" & Replace(SN, "'", "''") & "', '" & Priority & "', '" & _
                            Replace(Product, "'", "''") & "', '" & Replace(CallerName, "'", "''") & "'," & _
                            "   '" & Replace(CallerPhone, "'", "''") & "', '" & Replace(CallerEmail, "'", "''") & "', '" & _
                            Replace(ContactName, "'", "''") & "', '" & Replace(ContactPhone, "'", "''") & "', '" & _
                            Replace(ContactEmail, "'", "''") & "', GetDate()" & _
                            ")"
                If Debug_Mode Or Inform_Mode Then
                    sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #29 Partner Ticket Event Number: " & PartnerTicketNo _
                        & vbCrLf & vbCrLf & "SQL to execute: " & SQLString
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
                .ExecuteSQL(SQLString)
            End With
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Sub ParseVerintOnyxEmail #30" & vbCrLf & vbCrLf _
                & "SQL: " & SQLString & vbCrLf & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Sub
        End Try
    End Sub

    Sub ParseVerintOracleEmail(ByVal TicketEmailBody As String, ByVal MsgID As String, ByVal IntegrationRequestType As String, ByRef TakeABreak As Boolean)
        Dim SRSeverity As String = "", SRSeverityStart As Integer = 0, SRSeverityEnd As Integer = 0, SRSeverityDescription As String = ""
        Dim CustomerNameStart As Integer = 0, CustomerNameEnd As Integer = 0
        Dim SiteNameStart As Integer = 0, SiteNameEnd As Integer = 0
        Dim CustomerNumber As String = "", CustomerNumberStart As Integer = 0, CustomerNumberEnd As Integer = 0
        Dim IncidentNumber As String = "", IncidentNumberStart As Integer = 0, IncidentNumberEnd As Integer = 0
        Dim Location As String = "", LocationStart As Integer = 0, LocationEnd As Integer = 0
        Dim PostalCode As String = "", PostalCodeStart As Integer = 0, PostalCodeEnd As Integer = 0
        Dim IncidentContact As String = "", IncidentContactStart As Integer = 0, IncidentContactEnd As Integer = 0
        Dim IncidentPrimaryContactPhone As String = "", IncidentPrimaryContactPhoneStart As Integer = 0, IncidentPrimaryContactPhoneEnd As Integer = 0
        Dim IncidentPrimaryContactEmail As String = "", IncidentPrimaryContactEmailStart As Integer = 0, IncidentPrimaryContactEmailEnd As Integer = 0
        Dim SupportLevel As String = "", SupportLevelStart As Integer = 0, SupportLevelEnd As Integer = 0
        Dim SRType As String = "", SRTypeStart As Integer = 0, SRTypeEnd As Integer = 0
        Dim ServiceProductVertical As String = "", ServiceProductVerticalStart As Integer = 0, ServiceProductVerticalEnd As Integer = 0
        Dim SerialNumber As String = "", SerialNumberStart As Integer = 0, SerialNumberEnd As Integer = 0
        Dim TaskStatus As String = "", TaskStatusStart As Integer = 0, TaskStatusEnd As Integer = 0
        Dim LoggedBy As String = "", LoggedByStart As Integer = 0, LoggedByEnd As Integer = 0
        Dim WorkNotes As String = "", WorkNotesStart As Integer = 0, WorkNotesEnd As Integer = 0

        If Debug_Mode Then
            sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #01" & vbCrLf _
                & "Msg ID: " & MsgID & vbCrLf _
                & "Integration Request Type: " & IntegrationRequestType & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "Ticket Email Body" & vbCrLf & vbTab & "---- Beginnning of lines ----" & vbCrLf _
                & TicketEmailBody & vbCrLf & vbTab & "---- End of lines ----"
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        End If

        Try
            'Parse message body...Verint's emails are formatted the same regardless of them being a new or an update request
            'SR Severity         : VRNT P3
            SRSeverityStart = InStr(TicketEmailBody, "SR Severity        : ") + 21
            'sEventEntry = "CIM Service : Sub ParseVerintOracleEmail - SRSeverity #1: " & SRSeverity & vbCrLf & "SRSeverityStart: " & Str(SRSeverityStart) & vbCrLf & "SRSeverityEnd: " & Str(SRSeverityEnd)
            'WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            If (SRSeverityStart > 21) Then
                'SRSeverityEnd = InStr(SRSeverityStart, TicketEmailBody, vbNewLine)
                'sEventEntry = "CIM Service : Sub ParseVerintOracleEmail - SRSeverityStart: " & Str(SRSeverityStart) & vbCrLf & "SRSeverityEnd via vbNewLine: " & Str(SRSeverityEnd)
                'WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                'SRSeverityEnd = InStr(SRSeverityStart, TicketEmailBody, vbCrLf)
                'sEventEntry = "CIM Service : Sub ParseVerintOracleEmail - SRSeverityStart: " & Str(SRSeverityStart) & vbCrLf & "SRSeverityEnd via vbCrLf: " & Str(SRSeverityEnd)
                'WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                'SRSeverityEnd = InStr(SRSeverityStart, TicketEmailBody, vbCr)
                'sEventEntry = "CIM Service : Sub ParseVerintOracleEmail - SRSeverityStart: " & Str(SRSeverityStart) & vbCrLf & "SRSeverityEnd via vbCr: " & Str(SRSeverityEnd)
                'WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                'SRSeverityEnd = InStr(SRSeverityStart, TicketEmailBody, vbLf)
                'sEventEntry = "CIM Service : Sub ParseVerintOracleEmail - SRSeverityStart: " & Str(SRSeverityStart) & vbCrLf & "SRSeverityEnd via vbLf: " & Str(SRSeverityEnd)
                'WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                SRSeverityEnd = FindMin(InStr(SRSeverityStart, TicketEmailBody, vbCr), _
                    InStr(SRSeverityStart, TicketEmailBody, vbLf), _
                    InStr(SRSeverityStart, TicketEmailBody, vbCrLf), _
                    InStr(SRSeverityStart, TicketEmailBody, vbNewLine))
                If (SRSeverityEnd > 0) Then
                    SRSeverity = Trim(Mid(TicketEmailBody, SRSeverityStart, SRSeverityEnd - SRSeverityStart))
                    'sEventEntry = "CIM Service : Sub ParseVerintOracleEmail - SRSeverity #3: " & SRSeverity & vbCrLf & "SRSeverityStart: " & Str(SRSeverityStart) & vbCrLf & "SRSeverityEnd: " & Str(SRSeverityEnd)
                    'WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #02" & vbCrLf & "SRSeverity: " & SRSeverity, EventLogEntryType.Information)

            'Customer Name       : VERINT GMBH
            CustomerNameStart = InStr(TicketEmailBody, "Customer Name      : ") + 21
            If (CustomerNameStart > 21) Then
                'CustomerNameEnd = InStr(CustomerNameStart, TicketEmailBody, vbCrLf)
                CustomerNameEnd = FindMin(InStr(CustomerNameStart, TicketEmailBody, vbCr), _
                    InStr(CustomerNameStart, TicketEmailBody, vbLf), _
                    InStr(CustomerNameStart, TicketEmailBody, vbCrLf), _
                    InStr(CustomerNameStart, TicketEmailBody, vbNewLine))
                If (CustomerNameEnd > 0) Then
                    CustomerName = Trim(Mid(TicketEmailBody, CustomerNameStart, CustomerNameEnd - CustomerNameStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #03" & vbCrLf & "CustomerName: " & CustomerName, EventLogEntryType.Information)

            'Site Name           : VERINT GMBH KARLSRUHE,  D-76139 DE_Site
            SiteNameStart = InStr(TicketEmailBody, "Site Name          : ") + 21
            If (SiteNameStart > 21) Then
                'SiteNameEnd = InStr(SiteNameStart, TicketEmailBody, vbCrLf)
                SiteNameEnd = FindMin(InStr(SiteNameStart, TicketEmailBody, vbCr), _
                 InStr(SiteNameStart, TicketEmailBody, vbLf), _
                 InStr(SiteNameStart, TicketEmailBody, vbCrLf), _
                 InStr(SiteNameStart, TicketEmailBody, vbNewLine))
                If (SiteNameEnd > 0) Then
                    SiteName = Trim(Mid(TicketEmailBody, SiteNameStart, SiteNameEnd - SiteNameStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #04" & vbCrLf & "SiteName: " & SiteName, EventLogEntryType.Information)

            'Customer Number     : 295464
            CustomerNumberStart = InStr(TicketEmailBody, "Customer Number    : ") + 21
            If (CustomerNumberStart > 21) Then
                'CustomerNumberEnd = InStr(CustomerNumberStart, TicketEmailBody, vbCrLf)
                CustomerNumberEnd = FindMin(InStr(CustomerNumberStart, TicketEmailBody, vbCr), _
                    InStr(CustomerNumberStart, TicketEmailBody, vbLf), _
                    InStr(CustomerNumberStart, TicketEmailBody, vbCrLf), _
                    InStr(CustomerNumberStart, TicketEmailBody, vbNewLine))

                If (CustomerNumberEnd > 0) Then
                    CustomerNumber = Trim(Mid(TicketEmailBody, CustomerNumberStart, CustomerNumberEnd - CustomerNumberStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #05" & vbCrLf & "CustomerNumber: " & CustomerNumber, EventLogEntryType.Information)

            'Incident Number     : 1445980
            IncidentNumberStart = InStr(TicketEmailBody, "Incident Number    : ") + 21
            If (IncidentNumberStart > 21) Then
                'IncidentNumberEnd = InStr(IncidentNumberStart, TicketEmailBody, vbCrLf)
                IncidentNumberEnd = FindMin(InStr(IncidentNumberStart, TicketEmailBody, vbCr), _
                    InStr(IncidentNumberStart, TicketEmailBody, vbLf), _
                    InStr(IncidentNumberStart, TicketEmailBody, vbCrLf), _
                    InStr(IncidentNumberStart, TicketEmailBody, vbNewLine))
                If (IncidentNumberEnd > 0) Then
                    IncidentNumber = Trim(Mid(TicketEmailBody, IncidentNumberStart, IncidentNumberEnd - IncidentNumberStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #06" & vbCrLf & "IncidentNumber: " & IncidentNumber, EventLogEntryType.Information)

            'Location            : AM STORRENACKER 2, , KARLSRUHE, D-76139  DE
            LocationStart = InStr(TicketEmailBody, "Location           : ") + 21
            If (LocationStart > 21) Then
                'LocationEnd = InStr(LocationStart, TicketEmailBody, vbCrLf)
                LocationEnd = FindMin(InStr(LocationStart, TicketEmailBody, vbCr), _
                    InStr(LocationStart, TicketEmailBody, vbLf), _
                    InStr(LocationStart, TicketEmailBody, vbCrLf), _
                    InStr(LocationStart, TicketEmailBody, vbNewLine))

                If (LocationEnd > 0) Then
                    Location = Trim(Mid(TicketEmailBody, LocationStart, LocationEnd - LocationStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #07" & vbCrLf & "Location: " & Location, EventLogEntryType.Information)

            'Postal Code         : D-76139
            PostalCodeStart = InStr(TicketEmailBody, "Postal Code        : ") + 21
            If (PostalCodeStart > 21) Then
                'PostalCodeEnd = InStr(PostalCodeStart, TicketEmailBody, vbCrLf)
                PostalCodeEnd = FindMin(InStr(PostalCodeStart, TicketEmailBody, vbCr), _
                    InStr(PostalCodeStart, TicketEmailBody, vbLf), _
                    InStr(PostalCodeStart, TicketEmailBody, vbCrLf), _
                    InStr(PostalCodeStart, TicketEmailBody, vbNewLine))
                If (PostalCodeEnd > 0) Then
                    PostalCode = Trim(Mid(TicketEmailBody, PostalCodeStart, PostalCodeEnd - PostalCodeStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #08" & vbCrLf & "PostalCode: " & PostalCode, EventLogEntryType.Information)

            'Incident Primary Contact First name    : 
            IncidentContactStart = InStr(TicketEmailBody, "Incident Primary Contact First name   : ") + 40
            If (IncidentContactStart > 40) Then
                'IncidentContactEnd = InStr(IncidentContactStart, TicketEmailBody, vbCrLf)
                IncidentContactEnd = FindMin(InStr(IncidentContactStart, TicketEmailBody, vbCr), _
                    InStr(IncidentContactStart, TicketEmailBody, vbLf), _
                    InStr(IncidentContactStart, TicketEmailBody, vbCrLf), _
                    InStr(IncidentContactStart, TicketEmailBody, vbNewLine))
                If (IncidentContactEnd > 0) Then
                    IncidentContact = Trim(Mid(TicketEmailBody, IncidentContactStart, IncidentContactEnd - IncidentContactStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #09" & vbCrLf & "IncidentContact First Name: " & IncidentContact, EventLogEntryType.Information)

            'Incident Primary Contact Last name     : 
            IncidentContactStart = InStr(TicketEmailBody, "Incident Primary Contact Last name    : ") + 40
            If (IncidentContactStart > 40) Then
                'IncidentContactEnd = InStr(IncidentContactStart, TicketEmailBody, vbCrLf)
                IncidentContactEnd = FindMin(InStr(IncidentContactStart, TicketEmailBody, vbCr), _
                    InStr(IncidentContactStart, TicketEmailBody, vbLf), _
                    InStr(IncidentContactStart, TicketEmailBody, vbCrLf), _
                    InStr(IncidentContactStart, TicketEmailBody, vbNewLine))
                If (IncidentContactEnd > 0) Then
                    REM if first name exists, then adding a space between the first name and the last name
                    If Len(IncidentContact) > 0 Then IncidentContact = IncidentContact & " "
                    IncidentContact = IncidentContact & Trim(Mid(TicketEmailBody, IncidentContactStart, IncidentContactEnd - IncidentContactStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #10" & vbCrLf & "IncidentContact First Name <space> Last Name: " & IncidentContact, EventLogEntryType.Information)

            'Incident Primary Contact Phone         : 
            IncidentPrimaryContactPhoneStart = InStr(TicketEmailBody, "Incident Primary Contact Phone        : ") + 40
            If (IncidentPrimaryContactPhoneStart > 40) Then
                'IncidentPrimaryContactPhoneEnd = InStr(IncidentPrimaryContactPhoneStart, TicketEmailBody, vbCrLf)
                IncidentPrimaryContactPhoneEnd = FindMin(InStr(IncidentPrimaryContactPhoneStart, TicketEmailBody, vbCr), _
                    InStr(IncidentPrimaryContactPhoneStart, TicketEmailBody, vbLf), _
                    InStr(IncidentPrimaryContactPhoneStart, TicketEmailBody, vbCrLf), _
                    InStr(IncidentPrimaryContactPhoneStart, TicketEmailBody, vbNewLine))
                If (IncidentPrimaryContactPhoneEnd > 0) Then
                    IncidentPrimaryContactPhone = Trim(Mid(TicketEmailBody, IncidentPrimaryContactPhoneStart, IncidentPrimaryContactPhoneEnd - IncidentPrimaryContactPhoneStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #11" & vbCrLf & "IncidentPrimaryContactPhone: " & IncidentPrimaryContactPhone, EventLogEntryType.Information)

            'Incident Primary Contact Email         : 
            IncidentPrimaryContactEmailStart = InStr(TicketEmailBody, "Incident Primary Contact Email        : ") + 40
            If (IncidentPrimaryContactEmailStart > 40) Then
                'IncidentPrimaryContactEmailEnd = InStr(IncidentPrimaryContactEmailStart, TicketEmailBody, vbCrLf)
                IncidentPrimaryContactEmailEnd = FindMin(InStr(IncidentPrimaryContactEmailStart, TicketEmailBody, vbCr), _
                    InStr(IncidentPrimaryContactEmailStart, TicketEmailBody, vbLf), _
                    InStr(IncidentPrimaryContactEmailStart, TicketEmailBody, vbCrLf), _
                    InStr(IncidentPrimaryContactEmailStart, TicketEmailBody, vbNewLine))
                If (IncidentPrimaryContactEmailEnd > 0) Then
                    IncidentPrimaryContactEmail = Trim(Mid(TicketEmailBody, IncidentPrimaryContactEmailStart, IncidentPrimaryContactEmailEnd - IncidentPrimaryContactEmailStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #12" & vbCrLf & "IncidentPrimaryContactEmail: " & IncidentPrimaryContactEmail, EventLogEntryType.Information)

            'Support Level       : 
            SupportLevelStart = InStr(TicketEmailBody, "Support Level      : ") + 21
            If (SupportLevelStart > 21) Then
                'SupportLevelEnd = InStr(SupportLevelStart, TicketEmailBody, vbCrLf)
                SupportLevelEnd = FindMin(InStr(SupportLevelStart, TicketEmailBody, vbCr), _
                    InStr(SupportLevelStart, TicketEmailBody, vbLf), _
                    InStr(SupportLevelStart, TicketEmailBody, vbCrLf), _
                    InStr(SupportLevelStart, TicketEmailBody, vbNewLine))
                If (SupportLevelEnd > 0) Then
                    SupportLevel = Trim(Mid(TicketEmailBody, SupportLevelStart, SupportLevelEnd - SupportLevelStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #13" & vbCrLf & "SupportLevel: " & SupportLevel, EventLogEntryType.Information)

            'SR Type             : VRNT System Malfunction
            SRTypeStart = InStr(TicketEmailBody, "SR Type            : ") + 21
            If (SRTypeStart > 21) Then
                'SRTypeEnd = InStr(SRTypeStart, TicketEmailBody, vbCrLf)
                SRTypeEnd = FindMin(InStr(SRTypeStart, TicketEmailBody, vbCr), _
                    InStr(SRTypeStart, TicketEmailBody, vbLf), _
                    InStr(SRTypeStart, TicketEmailBody, vbCrLf), _
                    InStr(SRTypeStart, TicketEmailBody, vbNewLine))
                If (SRTypeEnd > 0) Then
                    SRType = Trim(Mid(TicketEmailBody, SRTypeStart, SRTypeEnd - SRTypeStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #14" & vbCrLf & "SRType: " & SRType, EventLogEntryType.Information)

            'Service Product Vertical : Next Generation (NG) WAS ServiceProductVertical Version : VI360
            ServiceProductVerticalStart = InStr(TicketEmailBody, "Service Product Vertical: ") + 26
            If (ServiceProductVerticalStart > 26) Then
                ServiceProductVerticalEnd = FindMin(InStr(ServiceProductVerticalStart, TicketEmailBody, vbCr), _
                    InStr(ServiceProductVerticalStart, TicketEmailBody, vbLf), _
                    InStr(ServiceProductVerticalStart, TicketEmailBody, vbCrLf), _
                    InStr(ServiceProductVerticalStart, TicketEmailBody, vbNewLine))
                If (ServiceProductVerticalEnd > 0) Then
                    ServiceProductVertical = Trim(Mid(TicketEmailBody, ServiceProductVerticalStart, ServiceProductVerticalEnd - ServiceProductVerticalStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #15" & vbCrLf & "ServiceProductVertical: " & ServiceProductVertical, EventLogEntryType.Information)

            'Serial Number       : 
            SerialNumberStart = InStr(TicketEmailBody, "Serial Number      : ") + 21
            If (SerialNumberStart > 21) Then
                SerialNumberEnd = FindMin(InStr(SerialNumberStart, TicketEmailBody, vbCr), _
                    InStr(SerialNumberStart, TicketEmailBody, vbLf), _
                    InStr(SerialNumberStart, TicketEmailBody, vbCrLf), _
                    InStr(SerialNumberStart, TicketEmailBody, vbNewLine))
                If (SerialNumberEnd > 0) Then
                    SerialNumber = Trim(Mid(TicketEmailBody, SerialNumberStart, SerialNumberEnd - SerialNumberStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #16" & vbCrLf & "SerialNumber: " & SerialNumber, EventLogEntryType.Information)

            'Task Status         : VRNT Closed
            TaskStatusStart = InStr(TicketEmailBody, "Task Status        : ") + 21
            If (TaskStatusStart > 21) Then
                TaskStatusEnd = FindMin(InStr(TaskStatusStart, TicketEmailBody, vbCr), _
                    InStr(TaskStatusStart, TicketEmailBody, vbLf), _
                    InStr(TaskStatusStart, TicketEmailBody, vbCrLf), _
                    InStr(TaskStatusStart, TicketEmailBody, vbNewLine))
                If (TaskStatusEnd > 0) Then
                    TaskStatus = Trim(Mid(TicketEmailBody, TaskStatusStart, TaskStatusEnd - TaskStatusStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #17" & vbCrLf & "TaskStatus: " & TaskStatus, EventLogEntryType.Information)

            'Logged By           : NSINGH
            LoggedByStart = InStr(TicketEmailBody, "Logged By          : ") + 21
            If (LoggedByStart > 21) Then
                LoggedByEnd = FindMin(InStr(LoggedByStart, TicketEmailBody, vbCr), _
                    InStr(LoggedByStart, TicketEmailBody, vbLf), _
                    InStr(LoggedByStart, TicketEmailBody, vbCrLf), _
                    InStr(LoggedByStart, TicketEmailBody, vbNewLine))
                If (LoggedByEnd > 0) Then
                    LoggedBy = Trim(Mid(TicketEmailBody, LoggedByStart, LoggedByEnd - LoggedByStart))
                End If
            End If
            If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #18" & vbCrLf & "LoggedBy: " & LoggedBy, EventLogEntryType.Information)

            'Work Notes          : 
            WorkNotesStart = InStr(TicketEmailBody, "Work Notes         : ") + 21
            If (WorkNotesStart > 21) Then
                WorkNotesEnd = Len(TicketEmailBody)
                'WorkNotesEnd cannot be calculated as following because its content is more than one line
                'WorkNotesEnd = FindMin(InStr(WorkNotesStart, TicketEmailBody, vbCr), _
                '    InStr(WorkNotesStart, TicketEmailBody, vbLf), _
                '    InStr(WorkNotesStart, TicketEmailBody, vbCrLf), _
                '    InStr(WorkNotesStart, TicketEmailBody, vbNewLine))

                If (WorkNotesEnd > 0) Then
                    ' Trim occurs
                    WorkNotes = Trim(Mid(TicketEmailBody, WorkNotesStart, WorkNotesEnd - WorkNotesStart))
                    ' remove up to the first 3 carriage returns so that the problem description will show!
                    'If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    '    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                    'End If
                    'If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    '    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                    'End If
                    'If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    '    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                    'End If
                    Dim HasLeadingWhite As Boolean = True
                    Do While HasLeadingWhite
                        If Mid(WorkNotes, 1, Len(vbCr)) = vbCr Then
                            WorkNotes = Mid(WorkNotes, Len(vbCr) + 1)
                        ElseIf Mid(WorkNotes, 1, Len(vbLf)) = vbLf Then
                            WorkNotes = Mid(WorkNotes, Len(vbLf) + 1)
                        ElseIf Mid(WorkNotes, 1, Len(vbCrLf)) = vbCrLf Then
                            WorkNotes = Mid(WorkNotes, Len(vbCrLf) + 1)
                        ElseIf Mid(WorkNotes, 1, Len(vbNewLine)) = vbNewLine Then
                            WorkNotes = Mid(WorkNotes, Len(vbNewLine) + 1)
                        Else
                            HasLeadingWhite = False
                        End If
                    Loop
                End If
            End If
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #19" & vbCrLf & "WorkNotes: " & vbCrLf _
                    & vbTab & "---- Beginning of lines ----" & vbCrLf & WorkNotes & vbCrLf _
                    & vbTab & "---- End of lines ----"
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Sub ParseVerintOracleEmail #20: " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Sub
        End Try

        Try
            If (InStr(UCase(SRSeverity), "P5") > 0) Then
                SeverityLevel = "5637144577"
                SeverityLevelDescription = "Critical"
            ElseIf (InStr(UCase(SRSeverity), "P4") > 0) Then
                SeverityLevel = "5637144578"
                SeverityLevelDescription = "High"
            ElseIf (InStr(UCase(SRSeverity), "P3") > 0) Then
                SeverityLevel = "5637144579"
                SeverityLevelDescription = "Normal"
            ElseIf (InStr(UCase(SRSeverity), "P2") > 0) Then
                SeverityLevel = "5637144580"
                SeverityLevelDescription = "Informational"
            Else
                SeverityLevel = "5637144576"
                SeverityLevelDescription = "Compliance"
            End If

            PartnerTicketNo = IncidentNumber
            SN = SerialNumber
            ContactName = IncidentContact
            ContactPhone = IncidentPrimaryContactPhone
            ContactEmail = IncidentPrimaryContactEmail
            Product = ServiceProductVertical
            ProblemDescription = WorkNotes
            If (TaskStatus = "VRNT Closed" Or TaskStatus = "VRNT Cancelled") Then
                RequestStatus = "Close"
            End If
            If Debug_Mode Or Inform_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail *Summary* #21" & vbCrLf _
                    & "Parsed an Oracle integration email Message ID: " & MsgID & vbCrLf _
                    & "SR Severity: " & SRSeverity & vbCrLf _
                    & "Customer Name: " & CustomerName & vbCrLf _
                    & "Site Name: " & SiteName & vbCrLf _
                    & "Customer Number: " & CustomerNumber & vbCrLf _
                    & "Partner Ticket Number: " & PartnerTicketNo & vbCrLf _
                    & "Location: " & Location & vbCrLf _
                    & "Postal Code: " & PostalCode & vbCrLf _
                    & "Contact Name: " & ContactName & vbCrLf _
                    & "Contact Phone: " & ContactPhone & vbCrLf _
                    & "Contact Email: " & ContactEmail & vbCrLf _
                    & "Support Level: " & SupportLevel & vbCrLf _
                    & "SR Type: " & SRType & vbCrLf _
                    & "Product: " & Product & vbCrLf _
                    & "Serial Number: " & SN & vbCrLf _
                    & "Task Status: " & TaskStatus & vbCrLf _
                    & "Logged By: " & LoggedBy & vbCrLf _
                    & "Severity Level: " & SeverityLevel & vbCrLf _
                    & "Severity Level Description: " & SeverityLevelDescription & vbCrLf _
                    & "Request Status: " & RequestStatus & vbCrLf _
                    & "Problem Description: " & vbCrLf & vbTab & "---- Beginning of lines ----" & vbCrLf & ProblemDescription & vbCrLf & vbTab & "---- End of lines ----"
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Sub ParseVerintOracleEmail #22 " & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Sub
        End Try

        Dim SQLString As String = ""
        Dim ds As DataSet = New DataSet, dt As DataTable = New DataTable
        Dim OldCustomerName As String = "", OldStatus As String = "", OldSiteName As String = "", OldSN As String = "", Old As String = "", OldSeverity_Priority As String = "", OldProduct As String = "", OldCallerName As String = ""
        Dim OldCallerPhone As String = "", OldCallerEmail As String = "", OldContactName As String = "", OldContactPhone As String = "", OldContactEmail As String = ""

        If IntegrationRequestType <> "NEW" Then
            Try
                ' Get old stuff to compare...need to compare against what was originally submitted by the Partner

                SQLString = "SELECT TOP 1 * FROM PartnerIntegrationEmailData WHERE PartnerTicketNo = '" & PartnerTicketNo & "' ORDER BY LastChange DESC"
                If Debug_Mode Then
                    sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #23" & vbCrLf & vbCrLf & "SQL to execute: " & SQLString
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If

                With DManager
                    ds = .GetDataSet(SQLString)
                    If Not ds Is Nothing Then
                        dt = ds.Tables(0)
                        If dt.Rows.Count > 0 Then
                            OldCustomerName = dt.Rows(0).Item("CustomerName")
                            'OldStatus = dt.Rows(0).Item("Status")
                            OldSiteName = dt.Rows(0).Item("SiteName")
                            OldSN = dt.Rows(0).Item("SN")
                            OldSeverity_Priority = dt.Rows(0).Item("Severity_Priority")
                            OldProduct = dt.Rows(0).Item("Product")
                            OldCallerName = dt.Rows(0).Item("CallerName")
                            OldCallerPhone = dt.Rows(0).Item("CallerPhone")
                            OldCallerEmail = dt.Rows(0).Item("CallerEmail")
                            OldContactName = dt.Rows(0).Item("ContactName")
                            OldContactPhone = dt.Rows(0).Item("ContactPhone")
                            OldContactEmail = dt.Rows(0).Item("ContactEmail")
                        End If
                    End If
                End With
            Catch ex As Exception
                TakeABreak = True
                sEventEntry = "CIM Service Exception: Sub ParseVerintOracleEmail #24:" & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                    & "Exception Target: " & ex.TargetSite.ToString()
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                Exit Sub
            End Try

            Try
                If Debug_Mode Then WriteToEventLog("CIM Service: Sub ParseVerintOracleEmail #25" & vbCrLf & "IF..THEN block begins", EventLogEntryType.Information)
                If (OldCustomerName <> CustomerName) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Customer Name" & vbCrLf & "OLD: " & OldCustomerName & vbCrLf & "NEW: " & CustomerName
                    SQLTicketChangesString = ", CustomerName = '" & CustomerName & "'"
                End If
                If (OldSiteName <> SiteName) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Site Name" & vbCrLf & "OLD: " & OldSiteName & vbCrLf & "NEW: " & SiteName
                    SQLTicketChangesString = ", SiteName = '" & SiteName & "'"
                End If
                If (OldSN <> SN) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Serial Number" & vbCrLf & "OLD: " & OldSN & vbCrLf & "NEW: " & SN
                    SQLTicketChangesString = ", SN = '" & SN & "'"
                End If
                If (OldSeverity_Priority <> SRSeverity) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Severity" & vbCrLf & "OLD: " & OldSeverity_Priority & vbCrLf & "NEW: " & SRSeverity
                    SQLTicketChangesString = ", SeverityLevel = '" & SeverityLevel & "'"
                    SQLTicketChangesString = ", SeverityLevelDescription = '" & SeverityLevelDescription & "'"
                End If
                If (OldProduct <> Product) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Product" & vbCrLf & "OLD: " & OldProduct & vbCrLf & "NEW: " & Product
                    SQLTicketChangesString = ", Product = '" & Product & "'"
                End If
                If (OldCallerName <> CallerName) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Caller Name" & vbCrLf & "OLD: " & OldCallerName & vbCrLf & "NEW: " & CallerName
                    SQLTicketChangesString = ", CallerName = '" & CallerName & "'"
                End If
                If (OldCallerPhone <> CallerPhone) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Caller Phone" & vbCrLf & "OLD: " & OldCallerPhone & vbCrLf & "NEW: " & CallerPhone
                    SQLTicketChangesString = ", CallerPhone = '" & CallerPhone & "'"
                End If
                If (OldCallerEmail <> CallerEmail) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Caller Email" & vbCrLf & "OLD: " & OldCallerEmail & vbCrLf & "NEW: " & CallerEmail
                    SQLTicketChangesString = ", CallerEmail = '" & CallerEmail & "'"
                End If
                If (OldContactName <> ContactName) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Contact Name" & vbCrLf & "OLD: " & OldContactName & vbCrLf & "NEW: " & ContactName
                    SQLTicketChangesString = ", ContactName = '" & ContactName & "'"
                End If
                If (OldContactPhone <> ContactPhone) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Contact Phone" & vbCrLf & "OLD: " & OldContactPhone & vbCrLf & "NEW: " & ContactPhone
                    SQLTicketChangesString = ", ContactPhone = '" & ContactPhone & "'"
                End If
                If (OldContactEmail <> ContactEmail) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Contact Email" & vbCrLf & "OLD: " & OldContactEmail & vbCrLf & "NEW: " & ContactEmail
                    SQLTicketChangesString = ", ContactEmail = '" & ContactEmail & "'"
                End If
            Catch ex As Exception
                TakeABreak = True
                sEventEntry = "CIM Service Exception: Sub ParseVerintOracleEmail #26: " & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                    & "Exception Target: " & ex.TargetSite.ToString()
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                Exit Sub
            End Try
        End If  'Paired with "If IntegrationRequestType <> "NEW" Then"

        Try
            With DManager
                SQLString = "INSERT INTO PartnerIntegrationEmailData (" _
                            & " PartnerTicketNo, CustomerName, Status, SiteName, SN, Severity_Priority, Product, CallerName, CallerPhone, CallerEmail," _
                            & " ContactName, ContactPhone, ContactEmail, LastChange" _
                            & ") VALUES (" _
                            & "   '" & Replace(PartnerTicketNo, "'", "''") & "', '" & Replace(CustomerName, "'", "''") & "', '" & Replace(TaskStatus, "'", "''") & "', '" _
                            & Replace(SiteName, "'", "''") & "', '" & Replace(SN, "'", "''") & "', '" & SRSeverity & "', '" & Replace(Product, "'", "''") & "', '" _
                            & Replace(CallerName, "'", "''") & "'," & "   '" & Replace(CallerPhone, "'", "''") & "', '" & Replace(CallerEmail, "'", "''") _
                            & "', '" & Replace(ContactName, "'", "''") & "', '" & Replace(ContactPhone, "'", "''") & "', '" & Replace(ContactEmail, "'", "''") & "', GetDate() " _
                            & ")"
                If Debug_Mode Or Inform_Mode Then
                    'Write("CIM Service - Inserted PartnerIntegrationEmailData Record: " & PartnerTicketNo & " Event", SQLString, 0)
                    sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #27" & vbCrLf _
                        & "Partner Ticket Event Number: " & PartnerTicketNo & vbCrLf _
                        & "SQL to execute: " & SQLString
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
                .ExecuteSQL(SQLString)
            End With
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Sub ParseVerintOracleEmail #28: " & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                & "SQL: " & SQLString & vbCrLf & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Sub
        End Try
    End Sub

    Public Function GetNextTicketRecID(ByRef TakeABreak As Boolean) As Long
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        Dim sql As String = ""
        GetNextTicketRecID = -1

        Try
            If TakeABreak = False Then
                sql = "SELECT TOP 1 ISNULL(RecID, 1000000000) + 1 AS NextRecID FROM C_Tickets WHERE RECID < 5000000000 ORDER BY RecID DESC"
                If Debug_Mode Then
                    sEventEntry = "CIM Service: Function GetNextTicketId, returns " & Convert.ToString(GetNextTicketRecID) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                        & "SQL to execute: " & sql
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
                With DManager
                    ds = .GetDataSet(sql)
                    If Not ds Is Nothing Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            GetNextTicketRecID = ds.Tables(0).Rows(0).Item("NextRecID")
                        Else
                            GetNextTicketRecID = 1000000000
                        End If
                    Else
                        GetNextTicketRecID = 1000000000
                    End If
                End With
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function GetNextTicketRecID, returns " _
                & Convert.ToString(GetNextTicketRecID) & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                & "SQL: " & sql & vbCrLf & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Function
        End Try
    End Function

    Public Function GetNextTicketId(ByVal DataAreaID As String, ByRef TakeABreak As Boolean) As String
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        REM Dim sql As String = ""

        Dim sNextTicketIdNum As String = ""
        Dim sNextTicketIdZeros As String = ""
        Dim ZeroCt As Integer = 1
        Dim x As Integer = 1

        GetNextTicketId = ""
        Try

            ' gets and increments
            sNextTicketIdNum = Convert.ToString(GetTicketNumberSequence(TakeABreak))
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function GetNextTicketId #01" & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                    & "sNextTicketIdNum: " & sNextTicketIdNum
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
            If Len(sNextTicketIdNum) > 0 Then

                'sql = "SELECT TOP 1 TicketID_Display as NextRecID FROM C_Tickets WHERE DataAreaID = '" & DataAreaID & "' order by TicketID_Display DESC"

                'Try
                '    With DManager
                '        ds = .GetDataSet(sql)
                '        If Not ds Is Nothing Then
                '            If ds.Tables(0).Rows.Count > 0 Then
                '                NextTicketIdNum = CInt(Replace(ds.Tables(0).Rows(0).Item("NextRecID"), "AGST", "")) + 1
                '            End If
                '        End If
                '    End With

                'Catch ex As Exception
                '    Write("Witness Quote Process - " & "Procedure:GetNextTicketId ", ex.Message, 2)
                'End Try

                ZeroCt = 6 - Len(sNextTicketIdNum)
                'Write("Witness Quote Process - " & "Procedure:GetNextTicketId ", "IdNum: " & NextTicketIdNum & " - IdNum len: " & Len(NextTicketIdNum) & " - ZeroCt: " & ZeroCt, 0)
                While x <= ZeroCt
                    sNextTicketIdZeros = sNextTicketIdZeros + "0"
                    x = x + 1
                End While

                GetNextTicketId = "AGST" & sNextTicketIdZeros & sNextTicketIdNum
                If Debug_Mode Then
                    sEventEntry = "CIM Service: Function GetNextTicketId #02, returns " & GetNextTicketId & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                        & "NextTicketIdZeros: " & sNextTicketIdZeros & vbCrLf _
                        & "NextTicketIdNum: " & sNextTicketIdNum & vbCrLf _
                        & "x: " & Convert.ToString(x) & vbCrLf _
                        & "ZeroCt: " & Convert.ToString(ZeroCt)
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function GetNextTicketId #03, returns " & GetNextTicketId & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "NextTicketIdZeros: " & sNextTicketIdZeros & vbCrLf _
                & "NextTicketIdNum: " & sNextTicketIdNum & vbCrLf _
                & "x: " & Convert.ToString(x) & vbCrLf _
                & "ZeroCt: " & Convert.ToString(ZeroCt) & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Function
        End Try
    End Function

    Public Function GetTicketNumberSequence(ByRef TakeABreak As Boolean) As Long
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim sql As String = "SELECT NextRec FROM NumberSequenceTable WHERE DataAreaID = '" & DataAreaID & "' AND NumberSequence = 'CC_Tickets'"

        GetTicketNumberSequence = -1
        Try
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function GetTicketNumberSequence #01" & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                    & "SQL to execute: " & sql
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
            With DManager
                If Debug_Mode Then WriteToEventLog("CIM Service: Function GetTicketNumberSequence #02", EventLogEntryType.Information)
                ds = .GetDataSet(sql)
                If Debug_Mode Then WriteToEventLog("CIM Service: Function GetTicketNumberSequence #03", EventLogEntryType.Information)
                If Not ds Is Nothing Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        GetTicketNumberSequence = ds.Tables(0).Rows(0).Item("nextrec")
                        'sEventEntry = "CIM Service : Function GetTicketNumberSequence - Ticket Number [NextRec]: " & CStr(GetTicketNumberSequence) & " and [NextRec] will increment it by 1"
                        'WriteToEventLog(sEventEntry, EventLogEntryType.Information)

                        'sql = "UPDATE NumberSequenceTable SET NextRec = " & GetTicketNumberSequence + 1 & " WHERE DataAreaID = '" & DataAreaID & "' AND numbersequence = 'CC_Tickets'"
                        '.ExecuteSQL(sql)
                        WriteToEventLog("CIM Service: Function GetTicketNumberSequence #04", EventLogEntryType.Information)
                    End If
                End If
            End With
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function GetTicketNumberSequence #05, returns " _
                    & Convert.ToString(GetTicketNumberSequence) & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                    & "SQL to execute: " & sql
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function GetTicketNumberSequence #06, returns " _
                & Convert.ToString(GetTicketNumberSequence) & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "SQL: " & sql & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Function
        End Try
    End Function
    Public Function IncrementTicketNumberSequence(ByRef TakeABreak As Boolean) As Boolean
        Dim TicketNum As Integer
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim sql As String = "SELECT NextRec FROM NumberSequenceTable WHERE DataAreaID = '" & DataAreaID & "' AND NumberSequence = 'CC_Tickets'"

        IncrementTicketNumberSequence = False
        Try
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function IncrementTicketNumberSequence #01" & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                    & "SQL to execute: " & sql
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
            With DManager
                If Debug_Mode Then WriteToEventLog("CIM Service: Function IncrementTicketNumberSequence #02", EventLogEntryType.Information)
                ds = .GetDataSet(sql)
                If Debug_Mode Then WriteToEventLog("CIM Service: Function IncrementTicketNumberSequence #03", EventLogEntryType.Information)
                If Not ds Is Nothing Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        TicketNum = ds.Tables(0).Rows(0).Item("NextRec") + 1
                        sql = "UPDATE NumberSequenceTable SET NextRec = " & TicketNum _
                            & " WHERE DataAreaID = '" & DataAreaID & "' AND numbersequence = 'CC_Tickets'"
                        .ExecuteSQL(sql)
                        If Debug_Mode Then WriteToEventLog("CIM Service: Function IncrementTicketNumberSequence #04", EventLogEntryType.Information)
                        IncrementTicketNumberSequence = True
                    End If
                End If
            End With
            If Debug_Mode Or Inform_Mode Then
                sEventEntry = "CIM Service: Function IncrementTicketNumberSequence #05, returns " _
                    & Convert.ToString(IncrementTicketNumberSequence) & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                    & "TicketNum: " & Convert.ToString(TicketNum) & vbCrLf & vbCrLf _
                    & "SQL executed: " & sql
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function IncrementTicketNumberSequence #06, returns " _
                & Convert.ToString(IncrementTicketNumberSequence) & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "TicketNum: " & Convert.ToString(TicketNum) & vbCrLf & vbCrLf _
                & "SQL: " & sql & vbCrLf & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Function

    Public Function GetNextTicketEventRecID(ByRef TakeABreak As Boolean) As Long
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        Dim sql As String = ""

        GetNextTicketEventRecID = -1
        Try
            If TakeABreak = False Then
                sql = "SELECT TOP 1 RecID + 1 as NextRecID FROM C_TicketEvents ORDER BY RecID DESC"
                If Debug_Mode Then
                    sEventEntry = "CIM Service: Function GetNextTicketEventRecId #01" & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                        & "SQL to execute: " & sql
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
                With DManager
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function GetNextTicketEventRecID #02", EventLogEntryType.Information)
                    ds = .GetDataSet(sql)
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function GetNextTicketEventRecID #03", EventLogEntryType.Information)
                    If Not ds Is Nothing Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            GetNextTicketEventRecID = ds.Tables(0).Rows(0).Item("NextRecID")
                        End If
                        If Debug_Mode Then WriteToEventLog("CIM Service: Function GetNextTicketEventRecID #04", EventLogEntryType.Information)
                    End If
                End With
                If Debug_Mode Then
                    sEventEntry = "CIM Service: Function GetNextTicketEventRecId #05, returns " _
                        & Convert.ToString(GetNextTicketEventRecID) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                        & "SQL executed: " & sql
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function GetNextTicketEventRecId #06, returns " _
                & Convert.ToString(GetNextTicketEventRecID) & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                & "SQL: " & sql & vbCrLf & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Function

    Sub SendUpdateTicketNotFoundEmail(ByVal msg As MailMessage, ByVal Partner As String, ByVal PartnerTicketNo As String, ByVal DataAreaID As String)
        Dim mailMsg As MailMessage = New MailMessage
        mailMsg.From.EMail = "Support@adtechglobal.com"

        Dim client As SMTP = New SMTP
        client.Host = "inrelay.it.adtech.net"
        client.Port = 25

        Try
            'will need another IF THEN block for 'UK'                               Adam Ip 
            If DataAreaID = "US" Then
                mailMsg.To.Add("Arnett Kelly", "akelly@adtech.net")
                mailMsg.To.Add("Brian Brazil", "bbrazil@adtech.net")
                mailMsg.Subject = "DROPPED " & Partner & " Update Email (incident #: " & PartnerTicketNo & ")"
                mailMsg.Body = Partner & " incident # " & Chr(96) & PartnerTicketNo & Chr(39) & " was not found in the AGS support ticket system. " _
                        & vbCrLf & vbCrLf & "AGS Automated System Message" & vbCrLf & vbCrLf _
                        & "---- Original message ----" & vbCrLf & vbCrLf & msg.PlainMessage.Body.ToString()
                mailMsg.Headers.Add("Reply-To", "noreply@adtech.net")    ' Adam Ip 2011-03-17 
                client.SendMessage(mailMsg)
            End If
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub SendUpdateTicketNotFoundEmail" & vbCrLf _
                    & "Message ID: " & msg.MessageID & vbCrLf _
                    & "From: " & mailMsg.From.ToString & vbCrLf _
                    & "To: " & mailMsg.To.ToString & vbCrLf _
                    & "Subject: " & mailMsg.Subject & vbCrLf _
                    & "Message Body: " & mailMsg.Body
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            sEventEntry = "CIM Service Exception: Sub SendUpdateTicketNotFoundEmail" & vbCrLf _
                & "Message ID: " & msg.MessageID & vbCrLf _
                & "From: " & mailMsg.From.ToString & vbCrLf _
                & "To: " & mailMsg.To.ToString & vbCrLf _
                & "Subject: " & mailMsg.Subject & vbCrLf _
                & "Message Body: " & mailMsg.Body & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Sub

    Sub SendReplyEmail(ByVal msg As MailMessage)
        Dim mailMsg As MailMessage = New MailMessage
        mailMsg.From.EMail = "support@adtech.net"

        Dim client As SMTP = New SMTP
        client.Host = "inrelay.it.adtech.net"
        client.Port = 25

        Try
            mailMsg.To.Add(msg.From.Name.ToString(), msg.From.EMail.ToString())
            mailMsg.Subject = "We've received your message: " & msg.Subject
            mailMsg.Body = "---- Original message ----" & vbCrLf & vbCrLf & msg.PlainMessage.Body.ToString()
            mailMsg.Headers.Add("Reply-To", "noreply@adtech.net")    ' Adam Ip 2011-03-17 

            client.SendMessage(mailMsg)
            If Debug_Mode Then
                'Write("CIM Service", "Reply email sent for Message ID: " & msg.MessageID & " has been sent to: " & mailMsg.To.ToString() & ".", 0)
                sEventEntry = "CIM Service: Sub SendReplyEmail" & vbCrLf _
                    & "Message ID: " & msg.MessageID & vbCrLf _
                    & "From: " & mailMsg.From.ToString & vbCrLf _
                    & "To: " & mailMsg.To.ToString & vbCrLf _
                    & "Subject: " & mailMsg.Subject & vbCrLf _
                    & "Message Body: " & mailMsg.Body
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            sEventEntry = "CIM Service Exception: Sub SendReplyEmail" & vbCrLf _
                & "Message ID: " & msg.MessageID & vbCrLf _
                & "From: " & mailMsg.From.ToString & vbCrLf _
                & "To: " & mailMsg.To.ToString & vbCrLf _
                & "Subject: " & mailMsg.Subject & vbCrLf _
                & "Message Body: " & mailMsg.Body & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Sub

    Sub SendNotificationEmail(ByVal msg As MailMessage)
        Dim mailMsg As MailMessage = New MailMessage
        mailMsg.From.EMail = "support@adtech.net"

        Dim client As SMTP = New SMTP
        client.Host = "inrelay.it.adtech.net"
        client.Port = 25

        Try
            mailMsg.To.Add("Matt Pearson", "mpearson@adtech.net")
            mailMsg.Subject = "New DotProject Email Ticket: " & msg.Subject
            mailMsg.Body = "---- Original message ----" & vbCrLf & vbCrLf & msg.PlainMessage.Body.ToString()

            client.SendMessage(mailMsg)

            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub SendNotificationEmail" & vbCrLf _
                    & "Message ID: " & msg.MessageID & vbCrLf _
                    & "From: " & mailMsg.From.ToString & vbCrLf _
                    & "To: " & mailMsg.To.ToString & vbCrLf _
                    & "Subject: " & mailMsg.Subject & vbCrLf _
                    & "Message Body: " & mailMsg.Body
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            sEventEntry = "CIM Service Exception: Sub SendNotificationEmail" & vbCrLf _
                & "Message ID: " & msg.MessageID & vbCrLf _
                & "From: " & mailMsg.From.ToString & vbCrLf _
                & "To: " & mailMsg.To.ToString & vbCrLf _
                & "Subject: " & mailMsg.Subject & vbCrLf _
                & "Message Body: " & mailMsg.Body & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Sub

    Public Function FindMessageID(ByVal MessageID As String, ByRef TakeABreak As Boolean) As Boolean
        Dim sql As String = ""
        Dim strMsgID As String = ""

        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        FindMessageID = False
        Try
            MessageID = Trim(MessageID)
            If Len(MessageID) > 0 And TakeABreak = False Then
                sql = "SELECT MessageID FROM MessageIDs WHERE MessageID = '" & MessageID & "'"
                With DManager
                    ds = .GetDataSet(sql)
                    If Not ds Is Nothing Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            strMsgID = ds.Tables(0).Rows(0).Item("MessageID")
                            If Len(Trim(strMsgID)) > 0 Then
                                FindMessageID = True
                            End If
                        End If
                    End If
                End With
            End If
            If Debug_Mode Then
                If TakeABreak = True Then
                    sEventEntry = "CIM Service: Function FindMessageID #01 has an error occurs" & vbCrLf _
                        & "FindMessageID returns " & Convert.ToString(FindMessageID) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                        & "Message ID: '" & MessageID & "'" & vbCrLf _
                        & "strMsgID: '" & strMsgID & "'" & vbCrLf & vbCrLf _
                        & "SQL: " & sql
                Else
                    sEventEntry = "CIM Service: Function FindMessageID #02 returns " & Convert.ToString(FindMessageID) & vbCrLf & "This is a "
                    If FindMessageID = False Then
                        sEventEntry = sEventEntry & "NEW e-mail"
                    Else
                        sEventEntry = sEventEntry & "previously read e-mail"
                    End If
                    sEventEntry = sEventEntry & vbCrLf & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                        & "Message ID: '" & MessageID & "'" & vbCrLf _
                        & "strMsgID: '" & strMsgID & "'" & vbCrLf & vbCrLf _
                        & "SQL executed: " & sql
                End If
                REM display in system log only when this is a NEW e-mail; otherwise, several hundreds of previously read e-mails will be displayed
                If (Debug_Mode And FindMessageID = False) Or TakeABreak = True Then WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function FindMessageID #03" & vbCrLf _
                & "FindMessageID returns " & Convert.ToString(FindMessageID) & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "Message ID: '" & MessageID & "'" & vbCrLf _
                & "strMsgID: '" & strMsgID & "'" & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message & vbCrLf & vbCrLf _
                & "SQL: " & sql
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Function

    Public Function UpdateMessageID(ByVal MessageID As String, ByRef TakeABreak As Boolean) As Boolean
        Dim sql As String = ""

        UpdateMessageID = False
        Try
            If TakeABreak = False Then
                MessageID = Trim(MessageID)
                With DManager
                    sql = "INSERT INTO MessageIDs " & _
                                "(MessageID, CreationDateTime) " & _
                                "VALUES " & _
                                "(" & _
                                "'" & MessageID & "', " & _
                                "GetDate() " & _
                                ");"
                    If Debug_Mode Or Inform_Mode Then
                        sEventEntry = "CIM Service: Function UpdateMessageID #01" & vbCrLf _
                            & "Parameter MessageID passed ByVal: " & MessageID & vbCrLf _
                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                            & "SQL to execute: " & sql
                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    End If
                    .ExecuteSQL(sql)
                End With
                UpdateMessageID = True
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function UpdateMessageID #02 returns " & Convert.ToString(UpdateMessageID) & vbCrLf _
                & "Parameter MessageID passed ByVal: " & MessageID & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                & "SQL: " & sql & vbCrLf & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Function

    Public Function UpdateMessageIDinProgress(ByVal MessageID As String, ByVal sAction As Char, ByRef TakeABreak As Boolean) As Long
        Dim sql(2) As String
        Dim ds As DataSet = New DataSet

        UpdateMessageIDinProgress = CLng(0)
        sql(0) = ""
        sql(1) = ""

        Try
            If TakeABreak = False Then
                MessageID = Trim(MessageID)
                With DManager
                    Select Case sAction
                        Case "A"
                            sql(0) = "SELECT COUNT(MessageID) AS MsgCount FROM MessageIDinProgress WHERE MessageID LIKE '" & MessageID & "' GROUP BY MessageID;"
                            ds = .GetDataSet(sql(0))
                            If Not ds Is Nothing Then
                                If ds.Tables(0).Rows.Count > 0 Then
                                    UpdateMessageIDinProgress = CLng(ds.Tables(0).Rows(0).Item("MsgCount"))
                                End If
                            End If
                            sql(1) = "INSERT INTO MessageIDinProgress " _
                                        & "(MessageID, CreationDateTime) " _
                                        & "VALUES " _
                                        & "(" _
                                        & "'" & MessageID & "', " _
                                        & "GetDate() " _
                                        & ");"
                            .ExecuteSQL(sql(1))
                        Case "D"
                            sql(0) = "DELETE FROM MessageIDinProgress " _
                                        & "WHERE MessageID = '" & MessageID & "';"
                            .ExecuteSQL(sql(0))
                    End Select
                    If Debug_Mode Then
                        sEventEntry = "CIM Service: Function UpdateMessageIDinProgress #01 " & vbCrLf _
                            & "UpdateMessageIDinProgress returns: " & Convert.ToString(UpdateMessageIDinProgress) & vbCrLf _
                            & "Parameter MessageID passed ByVal: " & MessageID & vbCrLf _
                            & "Parameter sAction passed ByVal: " & Convert.ToString(sAction) & vbCrLf _
                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                            & "SQL(0) to execute: " & sql(0) & vbCrLf & vbCrLf _
                            & "SQL(1) to execute: " & sql(1)
                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    End If
                End With
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function UpdateMessageIDinProgress #02 returns " & Convert.ToString(UpdateMessageIDinProgress) & vbCrLf _
                & "Parameter MessageID passed ByVal: " & MessageID & vbCrLf _
                & "Parameter sAction passed ByVal: " & Convert.ToString(sAction) & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                & "SQL(0): " & sql(0) & vbCrLf & vbCrLf _
                & "SQL(1): " & sql(1) & vbCrLf & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Function
    Private Function OpenDB(ByRef TakeABreak As Boolean) As Boolean
        OpenDB = False
        Dim ConnString As String = ""
        Try
            'ODBC Connection String - SAVE THIS!
            ConnString = "Driver={SQL Server};" & "Server=" & DatabaseServer & ";" & _
                       "Database=" & DatabaseName & ";" & "Uid=" & DatabaseUserName & ";" & "Pwd=" & DatabasePassword & "; "

            DManager = New DataManager(ConnString)
            With DManager
                If .OpenConnection Then
                    OpenDB = True
                End If
            End With
            REM If Debug_Mode Then WriteToEventLog("CIM Service: Function OpenDB, returns " & Convert.ToString(OpenDB), EventLogEntryType.Information)
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function OpenDB, returns " & Convert.ToString(OpenDB) & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                & "SQL: " & ConnString & vbCrLf & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Function

    Private Function CloseDB() As Boolean
        With DManager
            .CloseConnection()
        End With
    End Function

    ' ValidateEmailSource - to validate the FROM email address.  In this application only 
    '           the following designated 3 email addresses will be further processed  
    Private Function ValidateEmailSource(ByVal MsgID As String, ByVal MsgFrom As String, ByRef TakeABreak As Boolean) As Boolean
        ValidateEmailSource = False
        Try
            If TakeABreak = False Then
                MsgFrom = LCase(MsgFrom)    ' convert to lower case
                If InStr(MsgFrom, "prodwfmailer@verint.com", CompareMethod.Text) > 0 Or InStr(MsgFrom, "onyxmail@verint.com", CompareMethod.Text) > 0 Or _
                    InStr(MsgFrom, "aip@adtechglobal.com", CompareMethod.Text) > 0 Or InStr(MsgFrom, "aip@adtech.net", CompareMethod.Text) > 0 Then
                    ValidateEmailSource = True
                End If
            End If
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function ValidateEmailSource, returns " & Convert.ToString(ValidateEmailSource) & vbCrLf _
                    & "Message From (input): " & MsgFrom & vbCrLf _
                    & "Message ID (input): " & MsgID & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function ValidateEmailSource, returns " & Convert.ToString(ValidateEmailSource) & vbCrLf _
                & "Message From (input): " & MsgFrom & vbCrLf _
                & "Message ID (input): " & MsgID & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Function
    Private Function FindMin(ByVal first As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Integer
        Dim iter As Integer, Arr(3) As Integer
        FindMin = first
        Arr(0) = x
        Arr(1) = y
        Arr(2) = z
        Try
            For iter = 0 To 2
                If FindMin > Arr(iter) Then
                    FindMin = Arr(iter)
                End If
            Next iter
        Catch ex As Exception
            sEventEntry = "CIM Service Exception: Function FindMin" & vbCrLf & "The 4 input parameters are " _
                & CStr(first) & ", " & CStr(x) & ", " & CStr(y) & ", " & CStr(x) & "." & vbCrLf _
                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString()
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            Exit Function
        End Try
    End Function

    Private Sub ClearMemory()
    End Sub

    Private Property DatabaseServer() As Object
        Get
            DatabaseServer = mDatabaseServer
        End Get
        Set(ByVal Value As Object)
            mDatabaseServer = Value
        End Set
    End Property

    Private Property DatabaseName() As Object
        Get
            DatabaseName = mDatabaseName
        End Get
        Set(ByVal Value As Object)
            mDatabaseName = Value
        End Set
    End Property

    Private Property DatabaseUserName() As Object
        Get
            DatabaseUserName = mDatabaseUsername
        End Get
        Set(ByVal Value As Object)
            mDatabaseUsername = Value
        End Set
    End Property

    Private Property DatabasePassword() As Object
        Get
            DatabasePassword = mDatabasePassword
        End Get
        Set(ByVal Value As Object)
            mDatabasePassword = Value
        End Set
    End Property
End Class
