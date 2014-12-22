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


    Public Sub New()
        If Not System.IO.Directory.Exists(Application.StartupPath & "\Attachments\") Then
            System.IO.Directory.CreateDirectory(Application.StartupPath & "\Attachments\")
        End If
        sFolder = Application.StartupPath & "\Attachments\"

        DatabaseServer = "db0.adtech.net" ' "jackson"
        DatabaseName = "DynamicsAxProd"
        DatabaseUserName = "Integration"
        DatabasePassword = "1ntegr@te"
    End Sub

    ''' <summary>
    ''' This Function is the entry point called by Windows Service             Adam Ip
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ProcessEmailInbox() As Boolean
        Const MAX_AREA As Integer = 2
        Dim i As Integer = 0, SuccessCount As Integer = 0
        Dim DataAreaIDcode(MAX_AREA - 1) As String

        Dim msg As MailMessage
        Dim Inbox(MAX_AREA - 1) As POP3

        DataAreaIDcode(0) = "UK"
        Inbox(0) = New POP3()
        Inbox(0).Username = "emeasupport"
        Inbox(0).Password = "zZ3k.D&tk9bew"
        Inbox(0).Host = "INPOP.it.adtech.net"

        DataAreaIDcode(1) = "US"
        Inbox(1) = New POP3()
        Inbox(1).Username = "support"
        Inbox(1).Password = "zZ3k.D&tk9bew"
        Inbox(1).Host = "INPOP.it.adtech.net"

        ProcessEmailInbox = False
        SuccessCount = 0
        'Write("CIM Service - Started processing", "UK US", 0)
        WriteToEventLog("CIM Service Started processing - UK US", EventLogEntryType.Information)

        For i = 0 To MAX_AREA - 1
            DataAreaID = DataAreaIDcode(i)
            Try
                If Inbox(i).Connect() = False Then
                    'Write("CIM Service: " & DataAreaID, "Connect to mailbox failed", 2)
                    WriteToEventLog("CIM Service: (" & DataAreaID & ") Procedure: ProcessEmailInbox - Connect to mailbox failed", EventLogEntryType.Error)
                Else
                    'Write("CIM Service - Connected to inbox (" & Inbox(i).Count & ")", "", 0)
                    'Write("CIM_PreVersion: " & DataAreaID & " service cycle started", Inbox(i).Count & " " & DataAreaID & " emails found", 0)
                    sEventEntry = "CIM_PreVersion: (" & DataAreaID & ") Procedure: ProcessEmailInbox - Services cycle started" _
                            & vbCrLf & Str(Inbox(i).Count) & " (" & DataAreaID & ") emails found"
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)

                    If OpenDB() Then
                        'Write("CIM Service - Opened DB", "", 0)
                        For Each msg In Inbox(i)
                            ' make sure this email is from the past week at least!
                            Dim TodaysDate As DateTime = DateTime.Now
                            'If msg.Date() Then

                            'End If

                            Application.DoEvents()

                            If Not FindMessageID(msg.MessageID) Then 'No Message ID found, then this is a new message, so process.  If Message ID found, then no process.

                                sEventEntry = "CIM Service: (" & DataAreaID & ") Procedure: ProcessEmailInbox - Started processing Message ID: " & msg.MessageID & vbCrLf _
                                        & "Message date: " & msg.Date() & vbCrLf _
                                        & "Message body: " & vbCrLf & "-- beginning of line --" & vbCrLf _
                                        & msg.PlainMessage.Body.ToString() & vbCrLf _
                                        & "-- end of line --" & vbCrLf
                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)

                                'if e-mail is NOT from Verint nor aip@adtech.net then ValidateEmailSource(msg.From.ToString()) returns false
                                '   then simply only update the MessageIDs table through UpdateMessageID()
                                'if ValidateEmailSource(msg.From.ToString()) returns true, then email is either from  
                                '    from Verint nor aip@adtech.net, then proceed with ProcessSupportMessage()
                                'Don't combine these 2 following IF's into one IF
                                '   If Not a Or B Then ...    the logic is different
                                If Not ValidateEmailSource(msg.MessageID, msg.From.ToString()) Then
                                    UpdateMessageID(msg.MessageID)
                                ElseIf ProcessSupportMessage(msg, DataAreaID) Then
                                    UpdateMessageID(msg.MessageID)
                                End If
                            End If
                        Next msg
                        SuccessCount = SuccessCount + 1
                    Else
                        'Write("CIM Service: " & DataAreaID, "Connect to database failed", 2)
                        WriteToEventLog("CIM Service: (" & DataAreaID & ") Procedure: ProcessEmailInbox - Connect to database failed", EventLogEntryType.Error)
                        'ProcessEmailInbox = False
                        Inbox(i).Disconnect()
                        'Do not Exit Function.  Proceed to the NEXT iteration
                    End If
                    Inbox(i).Disconnect()
                End If
            Catch ex As Exception
                'Write("CIM Service: " & DataAreaID, ex.Message, 2)
                sEventEntry = "CIM Service : (" & DataAreaID & ") Procedure: ProcessEmailInbox - " & vbCrLf & "Exception: " & ex.Message
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            End Try
        Next i
        If SuccessCount = MAX_AREA Then ProcessEmailInbox = True Else ProcessEmailInbox = False
        CloseDB()
        ClearMemory()
    End Function

    ' ProcessSupportMessage returns whether the message has been successfully processed
    Private Function ProcessSupportMessage(ByVal msg As MailMessage, ByVal DataAreaID As String) As Boolean
        Dim SQLString As String = ""

        Dim TicketEmailBody As String = msg.PlainMessage.Body.ToString()
        Dim TicketID As String = ""
        Dim IntegrationRequestType As String = "NEW"

        ProcessSupportMessage = False

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

        'Write("CIM Service", "From: " & msg.From.ToString() & " - To: " & msg.To.ToString(), 0)
        sEventEntry = "CIM Service - Function ProcessSupportMessage #01 - From: " & msg.From.ToString() & " - To: " & msg.To.ToString()
        WriteToEventLog(sEventEntry, EventLogEntryType.Information)

        Try
            If (msg.To.ToString().ToLower().Contains("cellstacksupport") Or msg.Cc.ToString().ToLower().Contains("cellstacksupport")) Then
                'DataAreaID = "US"
                Partner3_ = "OTH"
                TicketQueue = "CellStack"
                'Caller = "CON012643"
                'CallerName = msg.To.ToString()
                'CallerPhone = "1-800-494-8637"
                'CallerEmail = "feg-amer@witness.com"
                RequestStatus = ""
            ElseIf (InStr(UCase(msg.Subject), "CUSTOMER INCIDENT") > 0) And ((InStr(UCase(msg.Subject), "WITNESS") > 0) Or (InStr(UCase(msg.Subject), "VERINT") > 0)) Then
                If (InStr(UCase(msg.Subject), "UPDATED") > 0) Then
                    IntegrationRequestType = "UPDATE"
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
                    'Write("CIM Service - Found an Oracle Integration Request", "Subject: " & msg.Subject & vbCrLf & vbCrLf & "Body: " & TicketEmailBody, 0)
                    sEventEntry = "CIM Service - Function ProcessSupportMessage #02 - Found an Oracle Integration Request" & vbCrLf & "Subject: " & msg.Subject.ToString() & vbCrLf & "Body: " & TicketEmailBody
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    PartnerTicketSource = "Oracle"
                    ParseVerintOracleEmail(TicketEmailBody, msg.MessageID, IntegrationRequestType)
                Else
                    'Write("CIM Service - Found an Onyx Integration Request", "Subject: " & msg.Subject & vbCrLf & vbCrLf & "Body: " & TicketEmailBody, 0)
                    sEventEntry = "CIM Service - Function ProcessSupportMessage #03 - Found an Onyx Integration Request" & vbCrLf & "Subject: " & msg.Subject.ToString() & vbCrLf & "Body: " & TicketEmailBody
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    PartnerTicketSource = "Onyx"
                    ParseVerintOnyxEmail(TicketEmailBody, msg.MessageID, IntegrationRequestType)
                End If
            End If
        Catch ex As Exception
            sEventEntry = "CIM Service : Exception - Function ProcessSupportMessage #04: " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try

        ' last level of defense!  Can't update w/o these values
        If (Partner3_ = "" Or PartnerTicketNo = "") Then
            UpdatePartner = 0
        End If
        If (ContactEmail = "") Then
            UpdateContact = 0
        End If
        If (CallerEmail = "") Then
            UpdateCaller = 0
        End If

        'Write("CIM Service - Customer Name: " & CustomerName, "blah", 0)
        'sEventEntry = "CIM Service : Function ProcessSupportMessage #05: CustomerName = " & CustomerName
        'WriteToEventLog(sEventEntry, EventLogEntryType.Information)

        'Dim eventViewerString As String = ""
        If (Partner3_ <> "") Then ' partner integration needs to have a partner to process!
            If (PartnerTicketNo <> "") Then ' Verint integration needs to have a ticket number specified to process!
                Try
                    ' by this point, you should have parsed all applicable partner integration email messages...new or update.  The insert/update commands to follow are partner indepedent

                    If IntegrationRequestType = "NEW" Then ' only insert if this is a new ticket

                        ' Get the next ticket id.  
                        TicketID = GetNextTicketId(DataAreaID)
                        'If CInt(Replace(TicketID, "AGST", "")) + 1 > GetTicketNumberSequence() Then ' Check if the number sequence table needs to be synched up too...
                        '    With DManager
                        '        SQLString = "UPDATE NumberSequenceTable SET NextRec = " & CInt(Replace(TicketID, "AGST", "")) + 2 & " Where DataAreaID = '" & DataAreaID & "' and numbersequence = 'CC_Tickets'"
                        '        .ExecuteSQL(SQLString)
                        '        'Write("CIM Service - Updated numbersequenceTable", "SQL: " & SQLString, 0)
                        '    End With
                        'End If

                        'C_Tickets
                        WriteToEventLog("CIM Service - Function ProcessSupportMessage #06 - Next Ticket ID: " & TicketID, EventLogEntryType.Information)
                        With DManager
                            'SQLString = "INSERT INTO C_Tickets (" & _
                            '    "   CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, TicketID_Display, Source, Status, TicketQueue, Tier, Partner3_, PartnerTicketSource, UpdatePartner, CallerPhone, " & _
                            '    "   CallerEmail, RecID, RecVersion, DataAreaID, UpdateContact, UpdateCaller, Caller, CallerName, severityLevel, CustomerName, " & _
                            '    "   PartnerTicketNo, SiteName, Product, SN, ContactName, ContactPhone, ContactEmail, ProblemDescription) " & _
                            '    "" & _
                            '    "VALUES (" & _
                            '    "   'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & TicketID & "', 'Email', 'Open', '" & TicketQueue & "', 'Tier 1', '" & Partner3_ & "', '" & PartnerTicketSource & "', 1, '" & Replace(CallerPhone, "'", "''") & "', " & _
                            '    "   '" & Replace(CallerEmail, "'", "''") & "', '" & GetNextTicketRecId() & "', 1, '" & DataAreaID & "', 1, 1, '" & Replace(Caller, "'", "''") & "', '" & Replace(Mid(CallerName, 1, 60), "'", "''") & "', '" & SeverityLevel & "', '" & Replace(Mid(CustomerName, 1, 60), "'", "''") & "', " & _
                            '    "   '" & PartnerTicketNo & "', '" & Replace(Mid(SiteName, 1, 60), "'", "''") & "', '" & Replace(Product, "'", "''") & "', '" & Replace(SN, "'", "''") & "', '" & Replace(Mid(ContactName, 1, 60), "'", "''") & "', '" & Replace(ContactPhone, "'", "''") & "', '" & Replace(ContactEmail, "'", "''") & "', '" & Replace(ProblemDescription, "'", "''") & "')"

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
                                        & " 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & TicketID & "', 'Email', 'Open', '" & TicketQueue & "', 'Tier 1', '" _
                                        & Partner3_ & "', '" & PartnerTicketSource & "', 1, '" & Replace(CallerPhone, "'", "''") & "', " & "   '" & Replace(CallerEmail, "'", "''") & "', '" _
                                        & GetNextTicketRecID() & "', 1, '" & DataAreaID & "', 1, 1, '" _
                                        & Replace(Caller, "'", "''") & "', '" & Replace(Mid(CallerName, 1, 60), "'", "''") & "', '" _
                                        & SeverityLevel & "', '" & Replace(Mid(CustomerName, 1, 60), "'", "''") & "', '" _
                                        & PartnerTicketNo & "', '" & Replace(Mid(SiteName, 1, 60), "'", "''") & "', '" _
                                        & Replace(Product, "'", "''") & "', '" & Replace(SN, "'", "''") & "', '" _
                                        & Replace(Mid(ContactName, 1, 60), "'", "''") & "', '" & Replace(ContactPhone, "'", "''") & "', '" _
                                        & Replace(ContactEmail, "'", "''") & "', '" _
                                        & Replace(ProblemDescription, "'", "''") & "' )"

                            'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "'" & vbCrLf & "SQL: " & SQLString & vbCrLf
                            ' sEventEntry = "CIM Service - Function ProcessSupportMessage #07 Ticket ID: " & TicketID & vbCrLf & "SQL: " & SQLString 
                            sEventEntry = "CIM Service - Function ProcessSupportMessage #07 Ticket ID: " & TicketID & vbCrLf & "SQL: " & SQLString _
                            & vbCrLf & "SN: " & Replace(Mid(ContactName, 1, 60), "'", "''") _
                            & vbCrLf & "Contact Name: " & Replace(SN, "'", "''") _
                            & vbCrLf & "Contact Phone: " & Replace(ContactPhone, "'", "''") _
                            & vbCrLf & "Contact Email: " & Replace(ContactEmail, "'", "''") _
                            & vbCrLf & "Problem Desc: " & Replace(ProblemDescription, "'", "''")
                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            .ExecuteSQL(SQLString)
                            'Write("CIM Service - Inserted Ticket ID: " & TicketID, "SQL: " & SQLString, 0)
                            'sEventEntry = "CIM Service - Function ProcessSupportMessage #08 - Inserted Ticket ID: " & TicketID & vbCrLf & "SQL: " & SQLString
                            'WriteToEventLog(sEventEntry, EventLogEntryType.Information)

                            IncrementTicketNumberSequence()
                        End With

                        'C_TicketEvents
                        With DManager
                            SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " & _
                                        "VALUES ('New Ticket', 0, '" & TicketID & "', 'New " & Partner3_ & " Integration Email received: " & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"

                            'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                            'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                            sEventEntry = "CIM Service - Function ProcessSupportMessage #08 - Inserted Ticket ID: " & TicketID & " Event" & vbCrLf & "SQL: " & SQLString & vbCrLf
                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            .ExecuteSQL(SQLString)
                        End With
                    Else ' update ticket
                        ' get AGS ticket from partner's ticket number
                        Dim ds As DataSet = New DataSet
                        Dim dt As DataTable = New DataTable
                        Dim Status As String = ""

                        SQLString = "SELECT TOP 1 TicketID_Display, Status FROM C_Tickets WHERE PartnerTicketNo = '" & PartnerTicketNo & "' order by TicketID_Display desc"

                        With DManager
                            sEventEntry = "CIM Service - Function ProcessSupportMessage #09 - Partner Ticket ID: " & PartnerTicketNo & vbCrLf & "SQL: " & SQLString
                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            ds = .GetDataSet(SQLString)
                            If Not ds Is Nothing Then
                                dt = ds.Tables(0)
                                If dt.Rows.Count > 0 Then
                                    TicketID = dt.Rows(0).Item("TicketID_Display")
                                    Status = dt.Rows(0).Item("Status")
                                End If
                            End If
                        End With

                        If (TicketID = "") Then ' send email to Arnett Kelly and Brian Brazil that an update request was recieved but no ticket was found internally
                            SendUpdateTicketNotFoundEmail(msg, Partner3_, PartnerTicketNo, DataAreaID)
                        Else ' process accordingly
                            If Status = "Closed" And RequestStatus <> "Close" Then ' Re-Open...change value on header, update appropriate fields and add Event entry.
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), Status = 'Open', IsRead = 0 WHERE TicketID_Display = '" & TicketID & "'"
                                    ' " & SQLTicketChangesString & "
                                    'eventViewerString &= vbCrLf & vbCrLf & "Re-Opened Ticket ID '" & TicketID & "' SQL: " & SQLString
                                    'Write("CIM Service - Ticket ID: " & TicketID & " Re-Opened", "SQL: " & SQLString, 0)
                                    sEventEntry = "CIM Service - Function ProcessSupportMessage #10 - Ticket ID: " & TicketID & " Re-opened" & vbCrLf & "SQL: " & SQLString
                                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    .ExecuteSQL(SQLString)
                                End With

                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Ticket Re-Opened', 0, '" & TicketID & "', 'Ticket Re-Opened due to an update from the partner.  Please review and process accordingly." _
                                        & vbCrLf & vbCrLf & Replace(SQLTicketChangesEventString, "'", "''") & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf _
                                        & Replace(TicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" _
                                        & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"
                                    'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                                    'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                    sEventEntry = "CIM Service - Function ProcessSupportMessage #11 - Inserted Ticket ID: " & TicketID & " Event" & vbCrLf & "SQL: " & SQLString
                                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    .ExecuteSQL(SQLString)
                                End With
                            ElseIf Status <> "Closed" And RequestStatus = "Close" Then ' add Event entry.  Notify ticket owner that partner has closed the ticket
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), IsRead = 0 WHERE TicketID_Display = '" & TicketID & "'"
                                    'eventViewerString &= vbCrLf & vbCrLf & "Received Closed Ticket '" & TicketID & "' request, Ticket updated SQL: " & SQLString
                                    'Write("CIM Service - Ticket ID: " & TicketID & " Closed per partners request", "SQL: " & SQLString, 0)
                                    sEventEntry = "CIM Service - Function ProcessSupportMessage #12 - Ticket ID: " & TicketID & " Closed per partners request" & vbCrLf & "SQL: " & SQLString
                                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    .ExecuteSQL(SQLString)
                                End With

                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Comment', 1, '" & TicketID & "', 'Please note that it appears that this ticket may have been closed by the partner.  Please review and process accordingly." _
                                        & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") _
                                        & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel _
                                        & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"
                                    'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                                    'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                    sEventEntry = "CIM Service - Function ProcessSupportMessage #13 - Inserted Ticket ID: " & TicketID & " Event" & vbCrLf & "SQL: " & SQLString
                                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    .ExecuteSQL(SQLString)
                                End With
                            ElseIf Status = "Closed" And RequestStatus = "Close" Then ' Close request when our ticket was already closed...probably just a receipt confirmation from our partner.  No notification of ticket owner
                                ' no update needed to ticket header (C_Tickets)
                                'Write("CIM Service - Ticket ID: " & TicketID & " Closed per partners request", "Ticket was already closed internally", 0)
                                WriteToEventLog("CIM Service - Function ProcessSupportMessage #14 - Ticket ID: " & TicketID & " Closed per partners request" & vbCrLf & "Ticket was already closed internally", EventLogEntryType.Information)

                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                                & "VALUES ('Comment', 1, '" & TicketID & "', 'Please note that this ticket was closed by the partner (internal ticket already closed).  Review not necessary." _
                                                & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") _
                                                & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID _
                                                & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"
                                    'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event (close request on a closed ticket) SQL: " & SQLString
                                    'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                    sEventEntry = "CIM Service - Function ProcessSupportMessage #15 - Inserted Ticket ID: " & TicketID & " Event" & vbCrLf & "SQL: " & SQLString
                                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    .ExecuteSQL(SQLString)
                                End With
                            Else ' regular update request...insert and update accordingly
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET " & _
                                                "   ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), IsRead = 0 WHERE TicketID_Display = '" & TicketID & "' "
                                    ' " & SQLTicketChangesString & "

                                    'eventViewerString &= vbCrLf & vbCrLf & "Received Update Ticket '" & TicketID & "' request, Ticket updated SQL: " & SQLString
                                    'Write("CIM Service - Updated Ticket ID: " & TicketID, SQLString, 0)
                                    sEventEntry = "CIM Service - Function ProcessSupportMessage #16 - Updated Ticket ID: " & TicketID & vbCrLf & "SQL: " & SQLString
                                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    .ExecuteSQL(SQLString)
                                End With

                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Comment', 1, '" & TicketID & "', 'Partner Integration Update Email Received.  Please review and process accordingly." _
                                        & vbCrLf & vbCrLf & Replace(SQLTicketChangesEventString, "'", "''") & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf _
                                        & Replace(TicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription _
                                        & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"
                                    'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                                    'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                    sEventEntry = "CIM Service - Function ProcessSupportMessage #17 - Inserted Ticket ID: " & TicketID & " Event" & vbCrLf & "SQL: " & SQLString
                                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    .ExecuteSQL(SQLString)
                                End With
                            End If
                        End If
                    End If
                    'Write("CIM Service - received ticket integration email", eventViewerString, 0)
                    sEventEntry = "CIM Service - Function ProcessSupportMessage #18 Overall" & vbCrLf & "SQL : " & SQLString
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)

                    ProcessSupportMessage = True    'returns
                Catch ex As Exception
                    'Write("CIM Service", ex.Message & vbCrLf & vbCrLf & vbCrLf & "Current status(es): " & eventViewerString, 2)
                    'Write("CIM Service", ex.Message & vbCrLf & vbCrLf & "Possible SQL: " & SQLString, 2)
                    sEventEntry = "CIM Service : Exception - Function ProcessSupportMessage #19: " & vbCrLf _
                        & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message & vbCrLf _
                        & "Possible SQL: " & SQLString & vbCrLf
                    '& "Current status(es): " & eventViewerString & vbCrLf
                    WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                End Try
            Else ' process other partner ticket emails (non integration)
                Try
                    ' Get the next ticket id.  Check if the number sequence table needs to be synched up too...
                    TicketID = GetNextTicketId(DataAreaID)
                    'If CInt(Replace(TicketID, "AGST", "")) + 1 > GetTicketNumberSequence() Then
                    '    With DManager
                    '        SQLString = "UPDATE NumberSequenceTable SET NextRec = " & CInt(Replace(TicketID, "AGST", "")) + 2 & " Where DataAreaID = '" & DataAreaID & "' and numbersequence = 'CC_Tickets'"
                    '        .ExecuteSQL(SQLString)
                    '    End With
                    'End If

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

                    'DataAreaID = "US"
                    'Partner3_ = "OTH"
                    'TicketQueue = "CellStack"

                    'C_Tickets
                    With DManager
                        SQLString = "INSERT INTO C_Tickets (" & _
                                    " CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, TicketID_Display, Source, Status, TicketQueue, Tier, Partner3_," & _
                                    " PartnerTicketSource, UpdatePartner, CallerPhone," & _
                                    " CallerEmail, RecID, RecVersion, DataAreaID, UpdateContact, UpdateCaller, Caller, CallerName, severityLevel, CustomerName," & _
                                    " PartnerTicketNo, SiteName, Product, SN, ContactName, ContactPhone, ContactEmail, ProblemDescription" & _
                                    ")" & _
                                    "VALUES (" & _
                                    "   'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & TicketID & "', 'Email', 'Open', '" & _
                                    TicketQueue & "', 'Tier 1', '" & Partner3_ & "', 'email', 1, '', " & _
                                    "   '" & Replace(msg.From.EMail.ToString(), "'", "''") & "', '" & GetNextTicketRecId() & "', 1, '" & DataAreaID & "', 1, 1, '" & _
                                    Replace(msg.From.Name.ToString(), "'", "''") & "', '" & Replace(Mid(msg.From.Name.ToString(), 1, 60), "'", "''") & _
                                    "', '" & SeverityLevel & "', '', " & _
                                    "   '', '', '', '', '" & Replace(msg.From.Name.ToString(), "'", "''") & "', '', '" & Replace(msg.From.EMail.ToString(), "'", "''") & _
                                    "', '" & Replace(msg.PlainMessage.Body.ToString(), "'", "''") & "')"

                        'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' SQL: " & SQLString
                        'Write("CIM Service - Inserted Ticket ID: " & TicketID, "SQL: " & SQLString, 0)
                        sEventEntry = "CIM Service - Function ProcessSupportMessage #20 - Inserted Ticket ID: " & TicketID & vbCrLf & "SQL: " & SQLString
                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                        .ExecuteSQL(SQLString)

                        IncrementTicketNumberSequence()
                    End With

                    'C_TicketEvents
                    With DManager
                        SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " & _
                            "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " & _
                            "VALUES ('New Ticket', 0, '" & TicketID & "', 'New " & Partner3_ & " Integration Email received: " & vbCrLf & vbCrLf & _
                            "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") & _
                            "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & _
                            SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"
                        'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                        'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                        WriteToEventLog("CIM Service - Function ProcessSupportMessage #21 - Inserted Ticket ID: " & TicketID & " Event" & vbCrLf & "SQL: " & SQLString, EventLogEntryType.Information)
                        .ExecuteSQL(SQLString)
                    End With
                    'Write("CIM Service - received ticket integration email", eventViewerString, 0)
                    'WriteToEventLog("CIM Service - Received ticket integration email" & vbCrLf & "Event Viewer String : " & eventViewerString, EventLogEntryType.Information)
                    ProcessSupportMessage = True    'returns
                Catch ex As Exception
                    'Write("CIM Service", ex.Message & vbCrLf & vbCrLf & vbCrLf & "Current status(es): " & eventViewerString, 2)
                    'Write("CIM Service", ex.Message & vbCrLf & vbCrLf & "Possible SQL: " & SQLString, 2)
                    sEventEntry = "CIM Service : Exception - Function ProcessSupportMessage #22: " & vbCrLf _
                        & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message & vbCrLf _
                        & "Possible SQL: " & SQLString & vbCrLf
                    '& "Current status(es): " & eventViewerString & vbCrLf
                    WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                    ProcessSupportMessage = False   'returns
                End Try
            End If
        Else
            'Write("CIM Service - Email received that was not an integration request", TicketEmailBody, 0)
            sEventEntry = "CIM Service - Function ProcessSupportMessage #23 - Email received that was not an integration request" & vbCrLf _
                & "Ticket Email Body: " & vbCrLf & "-- beginning of line --" & vbCrLf _
                & TicketEmailBody & vbCrLf _
                & "-- end of line --" & vbCrLf
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            ProcessSupportMessage = True    'returns
        End If
    End Function

    Sub ParseVerintOnyxEmail(ByVal TicketEmailBody As String, ByVal MsgID As String, ByVal IntegrationRequestType As String)
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

        ' parse message body...Verint's emails are formatted the same regardless of them being a new or an update request

        'Priority:  4 - Low
        PriorityStart = InStr(TicketEmailBody, "Priority:") + 9
        If (PriorityStart > 9) Then
            PriorityEnd = FindMin(InStr(PriorityStart, TicketEmailBody, vbCr), _
                InStr(PriorityStart, TicketEmailBody, vbLf), _
                InStr(PriorityStart, TicketEmailBody, vbCrLf), _
                InStr(PriorityStart, TicketEmailBody, vbNewLine))
            If (PriorityEnd > 0) Then
                Priority = Trim(Mid(TicketEmailBody, PriorityStart, PriorityEnd - PriorityStart))
            End If
        End If

        'Company Name:  Hartford, The (Server #1) - Charlotte, NC
        CompanyNameStart = InStr(TicketEmailBody, "Company Name:") + 13
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

        'Onyx Customer Number:  159048
        OnyxCustomerNumberStart = InStr(TicketEmailBody, "Onyx Customer Number:") + 21
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

        'Onyx Incident Number:  3810615
        OnyxIncidentNumberStart = InStr(TicketEmailBody, "Onyx Incident Number:") + 21
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

        'Customer Primary Contact:  Frank Beatty
        CustomerPrimaryContactStart = InStr(TicketEmailBody, "Customer Primary Contact:") + 25
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

        'Location:  Charlotte Service Ctr, 8711 University East Dr, Charlotte, North Carolina, United States
        LocationStart = InStr(TicketEmailBody, "Location:") + 9
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

        'Postal Code:  28213
        PostalCodeStart = InStr(TicketEmailBody, "Postal Code:") + 12
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

        'Incident Contact:  Steven Hudak
        IncidentContactStart = InStr(TicketEmailBody, "Incident Contact:") + 17
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

        'Business Phone:  8609197122
        BusinessPhoneStart = InStr(TicketEmailBody, "Business Phone:") + 15
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

        'Cell Phone:  8609197122
        CellPhoneStart = InStr(TicketEmailBody, "Cell Phone:") + 11
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

        'Fax: 8609197122
        FaxStart = InStr(TicketEmailBody, "Fax:") + 4
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

        'Pager: 8609197122
        PagerStart = InStr(TicketEmailBody, "Pager:") + 6
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

        'Email Address:  steven.hudak@thehartford.com
        EmailAddressStart = InStr(TicketEmailBody, "Email Address:") + 14
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

        'Support Level:  Advantage
        SupportLevelStart = InStr(TicketEmailBody, "Support Level:") + 14
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

        'Incident Type:  Web - Unknown
        IncidentTypeStart = InStr(TicketEmailBody, "Incident Type:") + 14
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

        'Product:  eQContactStore
        ProductStart = InStr(TicketEmailBody, "Product:") + 8
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

        'Version:  CS - 7.2
        VersionStart = InStr(TicketEmailBody, "Version:") + 8
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

        'Serial Number: RN-XXXXXX
        SerialNumberStart = InStr(TicketEmailBody, "Serial Number:") + 14
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

        'Submitted By:  Janice Wells
        SubmittedByStart = InStr(TicketEmailBody, "Submitted By:") + 13
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

        'WorkNotes: 
        WorkNotesStart = InStr(TicketEmailBody, "WorkNotes:") + 10
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

        'Write("CIM Service - Parsed an Onyx integration email Message (ID: " & MsgID & ")", "Priority: '" & Priority & "'" & vbCrLf & "Company Name: '" & CompanyName & "'" & vbCrLf & "Site Name: '" & SiteName & "'" & vbCrLf & "Onyx Customer Number: '" & OnyxCustomerNumber & "'" & vbCrLf & "Onyx Incident Number: '" & OnyxIncidentNumber & "'" & vbCrLf & "Customer Primary Contact: '" & CustomerPrimaryContact & "'" & vbCrLf & "Location: '" & Location & "'" & vbCrLf & "Postal Code: '" & PostalCode & "'" & vbCrLf & "Incident Contact: '" & IncidentContact & "'" & vbCrLf & "Business Phone: '" & BusinessPhone & "'" & vbCrLf & "Cell Phone: '" & CellPhone & "'" & vbCrLf & "Fax: '" & Fax & "'" & vbCrLf & "Pager: '" & Pager & "'" & vbCrLf & "Email Address: '" & EmailAddress & "'" & vbCrLf & "Support Level: '" & SupportLevel & "'" & vbCrLf & "Incident Type: '" & IncidentType & "'" & vbCrLf & "Product: '" & Product & "'" & vbCrLf & "Version: '" & Version & "'" & vbCrLf & "Serial Number: '" & SerialNumber & "'" & vbCrLf & "Submitted By: '" & SubmittedBy & "'" & vbCrLf & "WorkNotes: '" & WorkNotes & "'", 0)
        sEventEntry = "CIM Service - Parsed an Onyx integration email Message (ID: " & MsgID & ")" & vbCrLf & "Priority: '" & Priority & "'" & vbCrLf _
                & "Company Name: '" & CompanyName & "'" & vbCrLf & "Site Name: '" & SiteName & "'" & vbCrLf & "Onyx Customer Number: '" & OnyxCustomerNumber _
                & "'" & vbCrLf & "Onyx Incident Number: '" & OnyxIncidentNumber & "'" & vbCrLf & "Customer Primary Contact: '" & CustomerPrimaryContact & "'" & vbCrLf _
                & "Location: '" & Location & "'" & vbCrLf & "Postal Code: '" & PostalCode & "'" & vbCrLf & "Incident Contact: '" & IncidentContact & "'" & vbCrLf _
                & "Business Phone: '" & BusinessPhone & "'" & vbCrLf & "Cell Phone: '" & CellPhone & "'" & vbCrLf & "Fax: '" & Fax & "'" & vbCrLf _
                & "Pager: '" & Pager & "'" & vbCrLf & "Email Address: '" & EmailAddress & "'" & vbCrLf & "Support Level: '" & SupportLevel & "'" & vbCrLf _
                & "Incident Type: '" & IncidentType & "'" & vbCrLf & "Product: '" & Product & "'" & vbCrLf & "Version: '" & Version & "'" & vbCrLf _
                & "Serial Number: '" & SerialNumber & "'" & vbCrLf & "Submitted By: '" & SubmittedBy & "'" & vbCrLf & "WorkNotes: '" & WorkNotes & "'"
        WriteToEventLog(sEventEntry, EventLogEntryType.Information)

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

        Dim SQLString As String = ""
        If IntegrationRequestType <> "NEW" Then
            ' Get old stuff to compare...need to compare against what was originally submitted by the Partner
            Dim ds As DataSet = New DataSet, dt As DataTable = New DataTable
            Dim OldCustomerName As String = "", OldStatus As String = "", OldSiteName As String = "", OldSN As String = "", Old As String = "", OldSeverity_Priority As String = "", OldProduct As String = "", OldCallerName As String = ""
            Dim OldCallerPhone As String = "", OldCallerEmail As String = "", OldContactName As String = "", OldContactPhone As String = "", OldContactEmail As String = ""

            SQLString = "SELECT TOP 1 * FROM PartnerIntegrationEmailData WHERE PartnerTicketNo = '" & PartnerTicketNo & "' order by LastChange desc"
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
                SQLTicketChangesEventString = "The following fields were changed from their originally submitted values:" & vbCrLf & vbCrLf & SQLTicketChangesEventString
            End If
        End If

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

            .ExecuteSQL(SQLString)
            'Write("CIM Service - Inserted PartnerIntegrationEmailData Record: " & PartnerTicketNo & " Event", SQLString, 0)
            sEventEntry = "CIM Service - Procedure: PartnerIntegrationEmailData - Partner Ticket Number: " & PartnerTicketNo & " Event" & vbCrLf & "SQL: " & SQLString
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        End With
    End Sub

    Sub ParseVerintOracleEmail(ByVal TicketEmailBody As String, ByVal MsgID As String, ByVal IntegrationRequestType As String)
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

        sEventEntry = "CIM Service : Sub ParseVerintOracleEmail - TicketEmailBody: " & TicketEmailBody
        WriteToEventLog(sEventEntry, EventLogEntryType.Information)

        Try
            'Parse message body...Verint's emails are formatted the same regardless of them being a new or an update request
            'SR Severity         : VRNT P3
            SRSeverityStart = InStr(TicketEmailBody, "SR Severity         :") + 21
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
            sEventEntry = "CIM Service : Sub ParseVerintOracleEmail - SRSeverity: " & SRSeverity & vbCrLf & "SRSeverityStart: " & Str(SRSeverityStart) & vbCrLf & "SRSeverityEnd: " & Str(SRSeverityEnd)
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)

            'Customer Name       : VERINT GMBH
            CustomerNameStart = InStr(TicketEmailBody, "Customer Name       :") + 21
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - CustomerName: " & CustomerName, EventLogEntryType.Information)

            'Site Name           : VERINT GMBH KARLSRUHE,  D-76139 DE_Site
            SiteNameStart = InStr(TicketEmailBody, "Site Name           :") + 21
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - SiteName: " & SiteName, EventLogEntryType.Information)

            'Customer Number     : 295464
            CustomerNumberStart = InStr(TicketEmailBody, "Customer Number     :") + 21
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - CustomerNumber: " & CustomerNumber, EventLogEntryType.Information)

            'Incident Number     : 1445980
            IncidentNumberStart = InStr(TicketEmailBody, "Incident Number     :") + 21
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - IncidentNumber: " & IncidentNumber, EventLogEntryType.Information)

            'Location            : AM STORRENACKER 2, , KARLSRUHE, D-76139  DE
            LocationStart = InStr(TicketEmailBody, "Location            :") + 21
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - Location: " & Location, EventLogEntryType.Information)

            'Postal Code         : D-76139
            PostalCodeStart = InStr(TicketEmailBody, "Postal Code         :") + 21
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - PostalCode: " & PostalCode, EventLogEntryType.Information)

            'Incident Primary Contact First name    : 
            IncidentContactStart = InStr(TicketEmailBody, "Incident Primary Contact First name    :") + 40
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - IncidentContact First Name: " & IncidentContact, EventLogEntryType.Information)

            'Incident Primary Contact Last name     : 
            IncidentContactStart = InStr(TicketEmailBody, "Incident Primary Contact Last name     :") + 40
            If (IncidentContactStart > 40) Then
                'IncidentContactEnd = InStr(IncidentContactStart, TicketEmailBody, vbCrLf)
                IncidentContactEnd = FindMin(InStr(IncidentContactStart, TicketEmailBody, vbCr), _
                    InStr(IncidentContactStart, TicketEmailBody, vbLf), _
                    InStr(IncidentContactStart, TicketEmailBody, vbCrLf), _
                    InStr(IncidentContactStart, TicketEmailBody, vbNewLine))
                If (IncidentContactEnd > 0) Then
                    IncidentContact = IncidentContact & " " & Trim(Mid(TicketEmailBody, IncidentContactStart, IncidentContactEnd - IncidentContactStart))
                End If
            End If
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - IncidentContact First Name & Last Name: " & IncidentContact, EventLogEntryType.Information)

            'Incident Primary Contact Phone         : 
            IncidentPrimaryContactPhoneStart = InStr(TicketEmailBody, "Incident Primary Contact Phone         :") + 40
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - IncidentPrimaryContactPhone: " & IncidentPrimaryContactPhone, EventLogEntryType.Information)

            'Incident Primary Contact Email         : 
            IncidentPrimaryContactEmailStart = InStr(TicketEmailBody, "Incident Primary Contact Email         :") + 40
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - IncidentPrimaryContactEmail: " & IncidentPrimaryContactEmail, EventLogEntryType.Information)

            'Support Level       : 
            SupportLevelStart = InStr(TicketEmailBody, "Support Level       :") + 21
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - SupportLevel: " & SupportLevel, EventLogEntryType.Information)

            'SR Type             : VRNT System Malfunction
            SRTypeStart = InStr(TicketEmailBody, "SR Type             :") + 21
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - SRType: " & SRType, EventLogEntryType.Information)

            'Service Product Vertical : Next Generation (NG) WAS ServiceProductVertical Version : VI360
            ServiceProductVerticalStart = InStr(TicketEmailBody, "Service Product Vertical :") + 26
            If (ServiceProductVerticalStart > 26) Then
                ServiceProductVerticalEnd = FindMin(InStr(ServiceProductVerticalStart, TicketEmailBody, vbCr), _
                    InStr(ServiceProductVerticalStart, TicketEmailBody, vbLf), _
                    InStr(ServiceProductVerticalStart, TicketEmailBody, vbCrLf), _
                    InStr(ServiceProductVerticalStart, TicketEmailBody, vbNewLine))
                If (ServiceProductVerticalEnd > 0) Then
                    ServiceProductVertical = Trim(Mid(TicketEmailBody, ServiceProductVerticalStart, ServiceProductVerticalEnd - ServiceProductVerticalStart))
                End If
            End If
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - ServiceProductVertical: " & ServiceProductVertical, EventLogEntryType.Information)

            'Serial Number       : 
            SerialNumberStart = InStr(TicketEmailBody, "Serial Number       :") + 21
            If (SerialNumberStart > 21) Then
                SerialNumberEnd = FindMin(InStr(SerialNumberStart, TicketEmailBody, vbCr), _
                    InStr(SerialNumberStart, TicketEmailBody, vbLf), _
                    InStr(SerialNumberStart, TicketEmailBody, vbCrLf), _
                    InStr(SerialNumberStart, TicketEmailBody, vbNewLine))
                If (SerialNumberEnd > 0) Then
                    SerialNumber = Trim(Mid(TicketEmailBody, SerialNumberStart, SerialNumberEnd - SerialNumberStart))
                End If
            End If
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - SerialNumber: " & SerialNumber, EventLogEntryType.Information)

            'Task Status         : VRNT Closed
            TaskStatusStart = InStr(TicketEmailBody, "Task Status         :") + 21
            If (TaskStatusStart > 21) Then
                TaskStatusEnd = FindMin(InStr(TaskStatusStart, TicketEmailBody, vbCr), _
                    InStr(TaskStatusStart, TicketEmailBody, vbLf), _
                    InStr(TaskStatusStart, TicketEmailBody, vbCrLf), _
                    InStr(TaskStatusStart, TicketEmailBody, vbNewLine))
                If (TaskStatusEnd > 0) Then
                    TaskStatus = Trim(Mid(TicketEmailBody, TaskStatusStart, TaskStatusEnd - TaskStatusStart))
                End If
            End If
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - TaskStatus: " & TaskStatus, EventLogEntryType.Information)

            'Logged By           : NSINGH
            LoggedByStart = InStr(TicketEmailBody, "Logged By           :") + 21
            If (LoggedByStart > 21) Then
                LoggedByEnd = FindMin(InStr(LoggedByStart, TicketEmailBody, vbCr), _
                    InStr(LoggedByStart, TicketEmailBody, vbLf), _
                    InStr(LoggedByStart, TicketEmailBody, vbCrLf), _
                    InStr(LoggedByStart, TicketEmailBody, vbNewLine))
                If (LoggedByEnd > 0) Then
                    LoggedBy = Trim(Mid(TicketEmailBody, LoggedByStart, LoggedByEnd - LoggedByStart))
                End If
            End If
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - LoggedBy: " & LoggedBy, EventLogEntryType.Information)

            'Work Notes          : 
            WorkNotesStart = InStr(TicketEmailBody, "Work Notes          :") + 21
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
            WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail - WorkNotes: " & WorkNotes, EventLogEntryType.Information)

            'Write("CIM Service --- Parsed an Oracle integration email Message (ID: " & MsgID & ")", "SR Severity: '" & SRSeverity & "'" & vbCrLf & "Customer Name: '" & CustomerName & "'" & vbCrLf & "Site Name: '" & SiteName & "'" & vbCrLf & "Customer Number: '" & CustomerNumber & "'" & vbCrLf & "Incident Number: '" & IncidentNumber & "'" & vbCrLf & "Location: '" & Location & "'" & vbCrLf & "Postal Code: '" & PostalCode & "'" & vbCrLf & "Incident Contact: '" & IncidentContact & "'" & vbCrLf & "Incident Primary Contact Phone: '" & IncidentPrimaryContactPhone & "'" & vbCrLf & "Incident Primary Contact Email: '" & IncidentPrimaryContactEmail & "'" & vbCrLf & "Support Level: '" & SupportLevel & "'" & vbCrLf & "SR Type: '" & SRType & "'" & vbCrLf & "Service Product Vertical: '" & ServiceProductVertical & "'" & vbCrLf & "Serial Number: '" & SerialNumber & "'" & vbCrLf & "Task Status: '" & TaskStatus & "'" & vbCrLf & "Logged By: '" & LoggedBy & "'" & vbCrLf & "Work Notes: '" & WorkNotes & "'", 0)
            sEventEntry = "CIM Service - Sub ParseVerintOracleEmail #01" & vbCrLf _
                    & "Parsed an Oracle integration email Message ID: " & Chr(96) & MsgID & Chr(39) & vbCrLf _
                    & "SR Severity: " & Chr(96) & SRSeverity & Chr(39) & vbCrLf _
                    & "Customer Name: " & Chr(96) & CustomerName & Chr(39) & vbCrLf _
                    & "Site Name: " & Chr(96) & SiteName & Chr(39) & vbCrLf _
                    & "Customer Number: " & Chr(96) & CustomerNumber & Chr(39) & vbCrLf _
                    & "Incident Number: " & Chr(96) & IncidentNumber & Chr(39) & vbCrLf _
                    & "Location: " & Chr(96) & Location & Chr(39) & vbCrLf _
                    & "Postal Code: " & Chr(96) & PostalCode & Chr(39) & vbCrLf _
                    & "Incident Contact: " & Chr(96) & IncidentContact & Chr(39) & vbCrLf _
                    & "Incident Primary Contact Phone: " & Chr(96) & IncidentPrimaryContactPhone & Chr(39) & vbCrLf _
                    & "Incident Primary Contact Email: " & Chr(96) & IncidentPrimaryContactEmail & Chr(39) & vbCrLf _
                    & "Support Level: " & Chr(96) & SupportLevel & Chr(39) & vbCrLf _
                    & "SR Type: " & Chr(96) & SRType & Chr(39) & vbCrLf _
                    & "Service Product Vertical: " & Chr(96) & ServiceProductVertical & Chr(39) & vbCrLf _
                    & "Serial Number: " & Chr(96) & SerialNumber & Chr(39) & vbCrLf _
                    & "Task Status: " & Chr(96) & TaskStatus & Chr(39) & vbCrLf _
                    & "Logged By: " & Chr(96) & LoggedBy & Chr(39) & vbCrLf _
                    & "Work Notes: " & Chr(96) & WorkNotes & Chr(39) & vbCrLf & vbCrLf & vbCrLf
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        Catch ex As Exception
            sEventEntry = "CIM Service : Exception - Sub ParseVerintOracleEmail #02: " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
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
        Catch ex As Exception
            sEventEntry = "CIM Service : Exception - Sub ParseVerintOracleEmail #03: " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try

        Dim SQLString As String = ""
        Dim ds As DataSet = New DataSet, dt As DataTable = New DataTable
        Dim OldCustomerName As String = "", OldStatus As String = "", OldSiteName As String = "", OldSN As String = "", Old As String = "", OldSeverity_Priority As String = "", OldProduct As String = "", OldCallerName As String = ""
        Dim OldCallerPhone As String = "", OldCallerEmail As String = "", OldContactName As String = "", OldContactPhone As String = "", OldContactEmail As String = ""

        If IntegrationRequestType <> "NEW" Then
            Try
                ' Get old stuff to compare...need to compare against what was originally submitted by the Partner

                SQLString = "SELECT TOP 1 * FROM PartnerIntegrationEmailData WHERE PartnerTicketNo = '" & PartnerTicketNo & "' order by LastChange desc"
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
                sEventEntry = "CIM Service : Exception - Sub ParseVerintOracleEmail #04: " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            End Try

            Try
                WriteToEventLog("CIM Service : Sub ParseVerintOracleEmail #05 IF..THEN block begins", EventLogEntryType.Information)
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
                sEventEntry = "CIM Service : Exception - Sub ParseVerintOracleEmail #06: " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
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
                'Write("CIM Service - Inserted PartnerIntegrationEmailData Record: " & PartnerTicketNo & " Event", SQLString, 0)
                sEventEntry = "CIM Service : Sub ParseVerintOracleEmail #07 - Partner Ticket Number: " & PartnerTicketNo & " Event" & vbCrLf & "SQL : " & SQLString
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                .ExecuteSQL(SQLString)
            End With
        Catch ex As Exception
            sEventEntry = "CIM Service : Exception - Sub ParseVerintOracleEmail #08: " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Sub

    Public Function GetNextTicketRecID() As Long
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        Dim sql As String = ""

        GetNextTicketRecID = 1

        sql = "SELECT TOP 1 ISNULL(RecID, 1000000000) + 1 as NextRecID FROM C_Tickets WHERE RECID < 5000000000 order by recid desc"

        Try
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

        Catch ex As Exception
            'Write("Witness Quote Process - " & "Procedure:GetNextTicketRecID ", ex.Message, 2)
            sEventEntry = "CIM Service : Exception in Function: GetNextTicketRecID " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try

    End Function

    Public Function GetNextTicketId(ByVal DataAreaID As String) As String
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        Dim sql As String = ""

        Dim NextTicketIdNum As String = ""
        Dim NextTicketIdZeros As String = ""
        Dim ZeroCt As Integer = 1
        Dim x As Integer = 1

        GetNextTicketId = ""

        NextTicketIdNum = GetTicketNumberSequence() ' gets and increments

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

        ZeroCt = 6 - Len(NextTicketIdNum)
        'Write("Witness Quote Process - " & "Procedure:GetNextTicketId ", "IdNum: " & NextTicketIdNum & " - IdNum len: " & Len(NextTicketIdNum) & " - ZeroCt: " & ZeroCt, 0)
        While x <= ZeroCt
            NextTicketIdZeros = NextTicketIdZeros + "0"
            x += 1
        End While

        GetNextTicketId = "AGST" & NextTicketIdZeros & NextTicketIdNum
        'Write("Witness Quote Process - " & "Procedure:GetNextTicketId ", "Zeros: " & NextTicketIdZeros & " - IdNum: " & NextTicketIdNum & " = " & GetNextTicketId, 0)
    End Function

    Public Function GetTicketNumberSequence() As Integer
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim sql As String = "SELECT nextrec FROM NumberSequenceTable WHERE DataAreaID = '" & DataAreaID & "' AND numbersequence = 'CC_Tickets'"

        GetTicketNumberSequence = 1
        Try
            With DManager
                ds = .GetDataSet(sql)
                If Not ds Is Nothing Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        GetTicketNumberSequence = ds.Tables(0).Rows(0).Item("nextrec")

                        'sEventEntry = "CIM Service : Function GetTicketNumberSequence - Ticket Number [NextRec]: " & CStr(GetTicketNumberSequence) & " and [NextRec] will increment it by 1"
                        'WriteToEventLog(sEventEntry, EventLogEntryType.Information)

                        'sql = "UPDATE NumberSequenceTable SET NextRec = " & GetTicketNumberSequence + 1 & " WHERE DataAreaID = '" & DataAreaID & "' AND numbersequence = 'CC_Tickets'"
                        '.ExecuteSQL(sql)
                    End If
                End If
            End With

        Catch ex As Exception
            'Write("Witness Quote Process - " & "Procedure:GetTicketNumberSequence ", ex.Message, 2)
            sEventEntry = "CIM Service : Exception in Function: GetTicketNumberSequence " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try

    End Function
    Public Function IncrementTicketNumberSequence() As Boolean
        Dim TicketNum As Integer
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable
        Dim sql As String = "SELECT nextrec FROM NumberSequenceTable WHERE DataAreaID = '" & DataAreaID & "' AND numbersequence = 'CC_Tickets'"

        IncrementTicketNumberSequence = False
        Try
            With DManager
                ds = .GetDataSet(sql)
                If Not ds Is Nothing Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        TicketNum = ds.Tables(0).Rows(0).Item("nextrec") + 1
                        sql = "UPDATE NumberSequenceTable SET NextRec = " & TicketNum & " WHERE DataAreaID = '" & DataAreaID & "' AND numbersequence = 'CC_Tickets'"
                        .ExecuteSQL(sql)
                        IncrementTicketNumberSequence = True
                    End If
                End If
            End With

        Catch ex As Exception
            'Write("Witness Quote Process - " & "Procedure:GetTicketNumberSequence ", ex.Message, 2)
            sEventEntry = "CIM Service : Exception in Function: IncrementTicketNumberSequence " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try

    End Function

    Public Function GetNextTicketEventRecId() As Long
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        Dim sql As String = ""

        GetNextTicketEventRecId = 1

        sql = "SELECT TOP 1 RecID + 1 as NextRecID FROM C_TicketEvents order by recid desc"

        Try
            With DManager
                ds = .GetDataSet(sql)
                If Not ds Is Nothing Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        GetNextTicketEventRecId = ds.Tables(0).Rows(0).Item("NextRecID")
                    End If
                End If
            End With

        Catch ex As Exception
            'Write("Witness Quote Process - " & "Procedure:GetNextTicketEventRecId ", ex.Message, 2)
            sEventEntry = "CIM Service : Exception in Function GetNextTicketEventRecId " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
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
                mailMsg.To.Add("Arnett Kelly", "akelly@adtechglobal.com")
                mailMsg.To.Add("Brian Brazil", "bbrazil@adtechglobal.com")
                mailMsg.Subject = "DROPPED " & Partner & " Update Email (incident #: " & PartnerTicketNo & ")"
                mailMsg.Body = Partner & " incident # " & Chr(96) & PartnerTicketNo & Chr(39) & " was not found in the AGS support ticket system. " _
                        & vbCrLf & vbCrLf & "AGS Automated System Message" & vbCrLf & vbCrLf _
                        & "---- Original message ----" & vbCrLf & vbCrLf & msg.PlainMessage.Body.ToString()
                mailMsg.Headers.Add("Reply-To", "noreply@dtechglobal.com")    ' Adam Ip 2011-03-17 
                client.SendMessage(mailMsg)

                'Write("CIM Service - " & mailMsg.Subject, mailMsg.Body, 0)
                sEventEntry = "CIM Service : Sub SendUpdateTicketNotFoundEmail" & vbCrLf _
                    & "To: " & mailMsg.To.ToString & vbCrLf _
                    & "Subject: " & mailMsg.Subject & vbCrLf & "Message body: " & mailMsg.Body & vbCrLf
                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
            End If
        Catch ex As Exception
            'Write("CIM Service:SendReplyEmail function", ex.Message, 2)
            sEventEntry = "CIM Service : Exception Sub SendUpdateTicketNotFoundEmail" & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Sub

    Sub SendReplyEmail(ByVal msg As MailMessage)
        Dim mailMsg As MailMessage = New MailMessage
        mailMsg.From.EMail = "Support@adtechglobal.com"

        Dim client As SMTP = New SMTP
        client.Host = "inrelay.it.adtech.net"
        client.Port = 25

        Try
            mailMsg.To.Add(msg.From.Name.ToString(), msg.From.EMail.ToString())
            mailMsg.Subject = "We've received your message: " & msg.Subject
            mailMsg.Body = "---- Original message ----" & vbCrLf & vbCrLf & msg.PlainMessage.Body.ToString()
            mailMsg.Headers.Add("Reply-To", "noreply@dtechglobal.com")    ' Adam Ip 2011-03-17 

            client.SendMessage(mailMsg)

            'Write("CIM Service", "Reply email sent for Message ID: " & msg.MessageID & " has been sent to: " & mailMsg.To.ToString() & ".", 0)
            sEventEntry = "CIM Service : Sub SendReplyEmail" & vbCrLf _
                & "Message ID: " & msg.MessageID & vbCrLf _
                & "To: " & mailMsg.To.ToString & vbCrLf _
                & "Subject: " & mailMsg.Subject & vbCrLf _
                & "Message body: " & mailMsg.Body & vbCrLf
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        Catch ex As Exception
            'Write("CIM Service:SendReplyEmail function", ex.Message, 2)
            sEventEntry = "CIM Service : Exception Sub SendReplyEmail " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Sub

    Sub SendNotificationEmail(ByVal msg As MailMessage)
        Dim mailMsg As MailMessage = New MailMessage
        mailMsg.From.EMail = "Support@adtechglobal.com"

        Dim client As SMTP = New SMTP
        client.Host = "inrelay.it.adtech.net"
        client.Port = 25

        Try
            mailMsg.To.Add("Matt Pearson", "mpearson@adtechglobal.com")
            mailMsg.Subject = "New DotProject Email Ticket: " & msg.Subject
            mailMsg.Body = "---- Original message ----" & vbCrLf & vbCrLf & msg.PlainMessage.Body.ToString()

            client.SendMessage(mailMsg)

            'Write("CIM Service", "Notification email sent for Message ID: " & msg.MessageID & " has been sent to: " & mailMsg.To.ToString() & ".", 0)
            sEventEntry = "CIM Service : Sub SendNotificationEmail" & vbCrLf _
                & "Message ID: " & msg.MessageID & vbCrLf & "To: " & mailMsg.To.ToString & vbCrLf _
                & "Subject: " & mailMsg.Subject & vbCrLf & "Message body: " & mailMsg.Body & vbCrLf
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        Catch ex As Exception
            'Write("CIM Service:SendReplyEmail function", ex.Message, 2)
            sEventEntry = "CIM Service : Exception Sub SendNotificationEmail " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Sub

    Public Function FindMessageID(ByVal MessageID As String) As Boolean
        Dim sql As String = ""
        Dim strMsgID As String = ""

        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        FindMessageID = False

        sql = "SELECT MessageID FROM MessageIDs WHERE MessageID = '" & MessageID & "'"
        Try
            With DManager
                ds = .GetDataSet(sql)
                If Not ds Is Nothing Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        strMsgID = ds.Tables(0).Rows(0).Item("MessageID")
                        If Trim(strMsgID) <> "" Then
                            FindMessageID = True
                        End If
                    End If
                End If
            End With

            If Not FindMessageID Then ' if not found, then this is a NEW message
                sEventEntry = "CIM Service : Function FindMessageID returns False" & vbCrLf _
                    & "Message ID: " & Chr(96) & strMsgID & Chr(39) & vbCrLf _
                    & "SQL: " & sql & vbCrLf
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

        Catch ex As Exception
            'Write("CIM Service - " & "Procedure:FindMessageID ", ex.Message, 2)
            sEventEntry = "CIM Service : Exception Function FindMessageID " & vbCrLf _
                & "Message ID: " & Chr(96) & MessageID & Chr(39) & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
        Return FindMessageID
    End Function

    Public Function UpdateMessageID(ByVal MessageID As String) As Boolean
        Dim sql As String = ""
        Try
            With DManager
                sql = "INSERT INTO MessageIDs " & _
                            "(MessageID, CreationDateTime) " & _
                            "VALUES " & _
                            "(" & _
                            "'" & MessageID & "', " & _
                            "GetDate() " & _
                            ");"
                .ExecuteSQL(sql)
                sEventEntry = "CIM Service - Function: UpdateMessageID" & vbCrLf & "SQL: " & sql
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End With
        Catch ex As Exception
            'Write("CIM Service - " & "Procedure:UpdateMessageID ", ex.Message, 2)
            sEventEntry = "CIM Service : Exception Function UpdateMessageID " & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Function

    Private Function OpenDB() As Boolean
        OpenDB = False
        Dim ConnString As String = ""
        Try
            'ODBC Connection String - SAVE THIS!
            ConnString = "Driver={SQL Server};" & "Server=" & DatabaseServer & ";" & _
                       "Database=" & DatabaseName & ";" & "Uid=" & DatabaseUserName & ";" & _
                       "Pwd=" & DatabasePassword & "; "

            DManager = New DataManager(ConnString)
            With DManager
                If .OpenConnection Then
                    OpenDB = True
                End If
            End With
        Catch ex As Exception
            'Write("CIM Service", ex.Message, 2)
            sEventEntry = "CIM Service : Exception Function OpenDB" & vbCrLf & "Connecting string: " & ConnString & vbCrLf _
                & "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
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
    Private Function ValidateEmailSource(ByVal msgid As String, ByVal msgfr As String) As Boolean
        Try
            msgfr = LCase(msgfr)    ' convert to lower case
            If InStr(msgfr, "prodwfmailer@verint.com", CompareMethod.Text) > 0 Or InStr(msgfr, "onyxmail@verint.com", CompareMethod.Text) > 0 Or _
                InStr(msgfr, "aip@adtechglobal.com", CompareMethod.Text) > 0 Or InStr(msgfr, "aip@adtech.net", CompareMethod.Text) > 0 Then
                ValidateEmailSource = True
            Else
                ValidateEmailSource = False
            End If
            sEventEntry = "CIM Service: Function ValidateEmailSource validates and returns " & Str(ValidateEmailSource) & _
                " on" & vbCrLf & "Message ID: " & msgid & vbCrLf & "Message From: " & msgfr
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)

        Catch ex As Exception
            sEventEntry = "CIM Service : Exception Function ValidateEmailSource.  The input parameters are" & vbCrLf & _
                "Message From: " & msgfr & vbCrLf & "Message ID: " & msgid & vbCrLf & _
                "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
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
            sEventEntry = "CIM Service : Exception Function FindMin.  The 4 parameters are " & CStr(first) & ", " & CStr(x) & ", " & CStr(y) & ", " & CStr(x) & "." & vbCrLf & _
                "Target: " & ex.TargetSite.ToString() & vbCrLf & "Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Function

    Private Sub ClearMemory()
    End Sub

    Private Property DatabaseServer()
        Get
            DatabaseServer = mDatabaseServer
        End Get
        Set(ByVal Value)
            mDatabaseServer = Value
        End Set
    End Property

    Private Property DatabaseName()
        Get
            DatabaseName = mDatabaseName
        End Get
        Set(ByVal Value)
            mDatabaseName = Value
        End Set
    End Property

    Private Property DatabaseUserName()
        Get
            DatabaseUserName = mDatabaseUsername
        End Get
        Set(ByVal Value)
            mDatabaseUsername = Value
        End Set
    End Property

    Private Property DatabasePassword()
        Get
            DatabasePassword = mDatabasePassword
        End Get
        Set(ByVal Value)
            mDatabasePassword = Value
        End Set
    End Property
End Class

