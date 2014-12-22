Imports System.IO
Imports System.Xml
Imports devBiz.Net.Mail
Imports System.Windows.Forms
Imports System.Data.SqlClient

Imports CIM_PreVersion.Constants
Imports CIM_PreVersion.EventLog

Public Class IntegrationParser
    Private objlogger As EventLog = New EventLog
    Private DManager As DataManager
    Private mQuoteKeyString As String
    Private mDatabaseServer As String
    Private mDatabaseName As String
    Private mDatabasePassword As String
    Private mDatabaseUsername As String
    Private mQuoteKey As String
    Private mInfoKey As String
    Private mCustKey As String
    Private NumberOfSites As Integer
    Private CurrentSite As Integer
    Private BaseQuoteKey As String
    Private SolutionType As String
    Private strFormattedXML As String
    Private RecordEffected As Long
    Private sFolder As String
    Private Subject As String
    Private Company As String
    Private Owner As String

    Dim TicketQueue As String, DataAreaID As String, Partner3_ As String, Caller As String, CallerName As String, CallerPhone As String, CallerEmail As String, PartnerTicketSource As String
    Dim SeverityLevel As String, SeverityLevelDescription As String, Customer As String, CustomerName As String, PartnerTicketNo As String, SiteName As String
    Dim Product As String, SN As String, ContactName As String, ContactPhone As String, ContactEmail As String, ProblemDescription As String
    Dim UpdatePartner As Integer, UpdateCaller As Integer, UpdateContact As Integer

    Dim RequestStatus As String, SQLTicketChangesString As String, SQLTicketChangesEventString As String


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

    Public Function ProcessEmailInbox() As Boolean
        Dim msg As MailMessage
        Dim Inbox As POP3 = New POP3

        Inbox.Username = "support"
        Inbox.Password = "zZ3k.D&tk9bew"
        Inbox.Host = "INPOP.it.adtech.net"

        ProcessEmailInbox = False
        'Write("CIM Service - Started processing", "", 0)
        Try
            If Inbox.Connect() = True Then
                Write("CIM_PreVersion: service cycle started", Inbox.Count & " emails found", 0)
                'Write("CIM Service - Connected to inbox (" & Inbox.Count & ")", "", 0)

                If OpenDB() Then
                    'Write("CIM Service - Opened DB", "", 0)

                    For Each msg In Inbox
                        ' make sure this email is from the past week at least!
                        Dim TodaysDate As DateTime = DateTime.Now
                        'If msg.Date() Then

                        'End If

                        Application.DoEvents()

                        If Not FindMessageID(msg.MessageID) Then 'No Message ID found so process.
                            Write("CIM Service - Started processing Message ID: " & msg.MessageID & " Date: " & msg.Date(), msg.PlainMessage.Body.ToString(), 0)
                            If ProcessSupportMessage(msg) Then
                                UpdateMessageID(msg.MessageID)
                            End If
                        End If
                    Next
                Else
                    Write("CIM Service", "Connect to database failed", 2)

                    ProcessEmailInbox = False
                    ClearMemory()
                    Inbox.Disconnect()

                    Exit Function
                End If
                ProcessEmailInbox = True
                CloseDB()
                ClearMemory()
                Inbox.Disconnect()
            Else
                Write("CIM Service", "Connect to mailbox failed", 2)
            End If
        Catch ex As Exception
            Write("CIM Service", ex.Message, 2)
        End Try
    End Function

    Private Function ProcessSupportMessage(ByVal msg As MailMessage) As Boolean
        Dim SQLString As String
        Dim ProcessOK As Boolean = False

        Dim TicketEmailBody As String = msg.PlainMessage.Body.ToString()

        Dim TicketID As String = "", IntegrationRequestType As String = "NEW"

        Partner3_ = ""
        PartnerTicketNo = ""
        TicketQueue = ""
        CallerName = ""
        CallerPhone = ""
        CallerEmail = ""
        DataAreaID = "US"
        SeverityLevel = ""
        SeverityLevelDescription = ""
        UpdatePartner = 0
        UpdateCaller = 0
        UpdateContact = 0
        SQLTicketChangesEventString = ""

        'Write("CIM Service", "From: " & msg.From.ToString() & " - To: " & msg.To.ToString(), 0)
        If (msg.To.ToString().ToLower().Contains("cellstacksupport") Or msg.Cc.ToString().ToLower().Contains("cellstacksupport")) Then
            DataAreaID = "US"
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

            DataAreaID = "US"
            Partner3_ = "VerintDir"
            TicketQueue = "Support Ct"
            Caller = "CON012643"
            CallerName = msg.To.ToString()
            CallerPhone = "1-800-494-8637"
            CallerEmail = "feg-amer@witness.com"
            RequestStatus = ""

            If (InStr(UCase(msg.Subject), "VERINT") > 0) Then ' spoof where it came from for testing...Oracle emails come with Verint in subject, Onyx with Witness
                'Write("CIM Service - Found an Oracle Integration Request", "Subject: " & msg.Subject & vbCrLf & vbCrLf & "Body: " & TicketEmailBody, 0)
                PartnerTicketSource = "Oracle"
                ParseVerintOracleEmail(TicketEmailBody, msg.MessageID, IntegrationRequestType)
            Else
                'Write("CIM Service - Found an Onyx Integration Request", "Subject: " & msg.Subject & vbCrLf & vbCrLf & "Body: " & TicketEmailBody, 0)
                PartnerTicketSource = "Onyx"
                ParseVerintOnyxEmail(TicketEmailBody, msg.MessageID, IntegrationRequestType)
            End If
        End If

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

        Dim eventViewerString As String = ""
        If (Partner3_ <> "") Then ' partner integration needs to have a partner to process!
            If (PartnerTicketNo <> "") Then ' Verint integration needs to have a ticket number specified to process!
                Try
                    ' by this point, you should have parsed all applicable partner integration email messages...new or update.  The insert/update commands to follow are partner indepedent

                    If IntegrationRequestType = "NEW" Then ' only insert if this is a new ticket

                        ' Get the next ticket id.  
                        TicketID = GetNextTicketId(DataAreaID)
                        'If CInt(Replace(TicketID, "AGST", "")) + 1 > GetTicketNumberSequence() Then ' Check if the number sequence table needs to be synched up too...
                        '    With DManager
                        '        SQLString = "UPDATE NumberSequenceTable SET NextRec = " & CInt(Replace(TicketID, "AGST", "")) + 2 & " Where dataareaid = '" & DataAreaID & "' and numbersequence = 'CC_Tickets'"
                        '        .ExecuteSQL(SQLString)
                        '        'Write("CIM Service - Updated numbersequenceTable", "SQL: " & SQLString, 0)
                        '    End With
                        'End If

                        'C_Tickets
                        With DManager
                            SQLString = "INSERT INTO C_Tickets (" & _
                                        "   CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, TicketID_Display, Source, Status, TicketQueue, Tier, Partner3_, PartnerTicketSource, UpdatePartner, CallerPhone, " & _
                                        "   CallerEmail, RecID, RecVersion, DataAreaID, UpdateContact, UpdateCaller, Caller, CallerName, severityLevel, CustomerName, " & _
                                        "   PartnerTicketNo, SiteName, Product, SN, ContactName, ContactPhone, ContactEmail, ProblemDescription) " & _
                                        "" & _
                                        "VALUES (" & _
                                        "   'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & TicketID & "', 'Email', 'Open', '" & TicketQueue & "', 'Tier 1', '" & Partner3_ & "', '" & PartnerTicketSource & "', 1, '" & Replace(CallerPhone, "'", "''") & "', " & _
                                        "   '" & Replace(CallerEmail, "'", "''") & "', '" & GetNextTicketRecId() & "', 1, '" & DataAreaID & "', 1, 1, '" & Replace(Caller, "'", "''") & "', '" & Replace(Mid(CallerName, 1, 60), "'", "''") & "', '" & SeverityLevel & "', '" & Replace(Mid(CustomerName, 1, 60), "'", "''") & "', " & _
                                        "   '" & PartnerTicketNo & "', '" & Replace(Mid(SiteName, 1, 60), "'", "''") & "', '" & Replace(Product, "'", "''") & "', '" & Replace(SN, "'", "''") & "', '" & Replace(Mid(ContactName, 1, 60), "'", "''") & "', '" & Replace(ContactPhone, "'", "''") & "', '" & Replace(ContactEmail, "'", "''") & "', '" & Replace(ProblemDescription, "'", "''") & "')"

                            eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' SQL: " & SQLString
                            .ExecuteSQL(SQLString)
                            'Write("CIM Service - Inserted Ticket ID: " & TicketID, "SQL: " & SQLString, 0)
                        End With

                        'C_TicketEvents
                        With DManager
                            SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " & _
                                        "VALUES ('New Ticket', 0, '" & TicketID & "', 'New " & Partner3_ & " Integration Email received: " & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"

                            eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                            .ExecuteSQL(SQLString)
                            'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                        End With
                    Else ' update ticket
                        ' get AGS ticket from partner's ticket number
                        Dim ds As DataSet = New DataSet
                        Dim dt As DataTable = New DataTable
                        Dim Status As String

                        SQLString = "SELECT TOP 1 TicketID_Display, Status FROM C_Tickets WHERE PartnerTicketNo = '" & PartnerTicketNo & "' order by TicketID_Display desc"

                        With DManager
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
                            SendUpdateTicketNotFoundEmail(msg, Partner3_, PartnerTicketNo)
                        Else ' process accordingly
                            If Status = "Closed" And RequestStatus <> "Close" Then ' Re-Open...change value on header, update appropriate fields and add Event entry.
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), Status = 'Open', IsRead = 0 WHERE TicketID_Display = '" & TicketID & "'"
                                    ' " & SQLTicketChangesString & "

                                    eventViewerString &= vbCrLf & vbCrLf & "Re-Opened Ticket ID '" & TicketID & "' SQL: " & SQLString
                                    .ExecuteSQL(SQLString)
                                    'Write("CIM Service - Ticket ID: " & TicketID & " Re-Opened", "SQL: " & SQLString, 0)
                                End With

                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " & _
                                                "VALUES ('Ticket Re-Opened', 0, '" & TicketID & "', 'Ticket Re-Opened due to an update from the partner.  Please review and process accordingly." & vbCrLf & vbCrLf & Replace(SQLTicketChangesEventString, "'", "''") & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"

                                    eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                                    .ExecuteSQL(SQLString)
                                    'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                End With
                            ElseIf Status <> "Closed" And RequestStatus = "Close" Then ' add Event entry.  Notify ticket owner that partner has closed the ticket
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), IsRead = 0 WHERE TicketID_Display = '" & TicketID & "'"

                                    eventViewerString &= vbCrLf & vbCrLf & "Received Closed Ticket '" & TicketID & "' request, Ticket updated SQL: " & SQLString
                                    .ExecuteSQL(SQLString)
                                    'Write("CIM Service - Ticket ID: " & TicketID & " Closed per partners request", "SQL: " & SQLString, 0)
                                End With

                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " & _
                                                "VALUES ('Comment', 1, '" & TicketID & "', 'Please note that it appears that this ticket may have been closed by the partner.  Please review and process accordingly." & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"

                                    eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                                    .ExecuteSQL(SQLString)
                                    'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                End With
                            ElseIf Status = "Closed" And RequestStatus = "Close" Then ' Close request when our ticket was already closed...probably just a receipt confirmation from our partner.  No notification of ticket owner
                                ' no update needed to ticket header (C_Tickets)
                                'Write("CIM Service - Ticket ID: " & TicketID & " Closed per partners request", "Ticket was already closed internally", 0)

                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " & _
                                                "VALUES ('Comment', 1, '" & TicketID & "', 'Please note that this ticket was closed by the partner (internal ticket already closed).  Review not necessary." & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"

                                    eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event (close request on a closed ticket) SQL: " & SQLString
                                    .ExecuteSQL(SQLString)
                                    'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                End With
                            Else ' regular update request...insert and update accordingly
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET " & _
                                                "   ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), IsRead = 0 WHERE TicketID_Display = '" & TicketID & "' "
                                    ' " & SQLTicketChangesString & "

                                    eventViewerString &= vbCrLf & vbCrLf & "Received Update Ticket '" & TicketID & "' request, Ticket updated SQL: " & SQLString
                                    .ExecuteSQL(SQLString)
                                    'Write("CIM Service - Updated Ticket ID: " & TicketID, SQLString, 0)
                                End With

                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " & _
                                                "VALUES ('Comment', 1, '" & TicketID & "', 'Partner Integration Update Email Received.  Please review and process accordingly." & vbCrLf & vbCrLf & Replace(SQLTicketChangesEventString, "'", "''") & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"

                                    eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                                    .ExecuteSQL(SQLString)
                                    'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                End With
                            End If
                        End If
                    End If
                    Write("CIM Service - received ticket integration email", eventViewerString, 0)

                    ProcessOK = True
                Catch ex As Exception
                    Write("CIM Service", ex.Message & vbCrLf & vbCrLf & vbCrLf & "Current status(es): " & eventViewerString, 2)
                    'Write("CIM Service", ex.Message & vbCrLf & vbCrLf & "Possible SQL: " & SQLString, 2)
                    ProcessOK = False
                End Try
            Else ' process other partner ticket emails (non integration)
                Try
                    ' Get the next ticket id.  Check if the number sequence table needs to be synched up too...
                    TicketID = GetNextTicketId(DataAreaID)
                    'If CInt(Replace(TicketID, "AGST", "")) + 1 > GetTicketNumberSequence() Then
                    '    With DManager
                    '        SQLString = "UPDATE NumberSequenceTable SET NextRec = " & CInt(Replace(TicketID, "AGST", "")) + 2 & " Where dataareaid = '" & DataAreaID & "' and numbersequence = 'CC_Tickets'"
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
                                    "   CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, TicketID_Display, Source, Status, TicketQueue, Tier, Partner3_, PartnerTicketSource, UpdatePartner, CallerPhone, " & _
                                    "   CallerEmail, RecID, RecVersion, DataAreaID, UpdateContact, UpdateCaller, Caller, CallerName, severityLevel, CustomerName, " & _
                                    "   PartnerTicketNo, SiteName, Product, SN, ContactName, ContactPhone, ContactEmail, ProblemDescription) " & _
                                    "" & _
                                    "VALUES (" & _
                                    "   'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & TicketID & "', 'Email', 'Open', '" & TicketQueue & "', 'Tier 1', '" & Partner3_ & "', 'email', 1, '', " & _
                                    "   '" & Replace(msg.From.EMail.ToString(), "'", "''") & "', '" & GetNextTicketRecId() & "', 1, '" & DataAreaID & "', 1, 1, '" & Replace(msg.From.Name.ToString(), "'", "''") & "', '" & Replace(Mid(msg.From.Name.ToString(), 1, 60), "'", "''") & "', '" & SeverityLevel & "', '', " & _
                                    "   '', '', '', '', '" & Replace(msg.From.Name.ToString(), "'", "''") & "', '', '" & Replace(msg.From.EMail.ToString(), "'", "''") & "', '" & Replace(msg.PlainMessage.Body.ToString(), "'", "''") & "')"

                        eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' SQL: " & SQLString
                        .ExecuteSQL(SQLString)
                        'Write("CIM Service - Inserted Ticket ID: " & TicketID, "SQL: " & SQLString, 0)
                    End With

                    'C_TicketEvents
                    With DManager
                        SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " & _
                                    "VALUES ('New Ticket', 0, '" & TicketID & "', 'New " & Partner3_ & " Integration Email received: " & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(TicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & GetNextTicketEventRecId() & "', 1, 1, 0)"

                        eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event SQL: " & SQLString
                        .ExecuteSQL(SQLString)
                        'Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                    End With
                    Write("CIM Service - received ticket integration email", eventViewerString, 0)

                    ProcessOK = True
                Catch ex As Exception
                    Write("CIM Service", ex.Message & vbCrLf & vbCrLf & vbCrLf & "Current status(es): " & eventViewerString, 2)
                    'Write("CIM Service", ex.Message & vbCrLf & vbCrLf & "Possible SQL: " & SQLString, 2)
                    ProcessOK = False
                End Try
            End If
        Else
            Write("CIM Service - Email received that was not an integration request", TicketEmailBody, 0)
            ProcessOK = True
        End If

        Return ProcessOK
    End Function

    Sub ParseVerintOnyxEmail(ByVal TicketEmailBody As String, ByVal MsgID As String, ByVal IntegrationRequestType As String)
        Dim Priority As String, PriorityStart As Integer, PriorityEnd As Integer, PriorityDescription As String
        Dim CompanyName As String, CompanyNameStart As Integer, CompanyNameEnd As Integer, CompanyNameEndA As Integer, CompanyNameEndB As Integer
        Dim SiteNameStart As Integer, SiteNameEnd As Integer ' SiteName As String, 
        Dim OnyxCustomerNumber As String, OnyxCustomerNumberStart As Integer, OnyxCustomerNumberEnd As Integer
        Dim OnyxIncidentNumber As String, OnyxIncidentNumberStart As Integer, OnyxIncidentNumberEnd As Integer
        Dim CustomerPrimaryContact As String, CustomerPrimaryContactStart As Integer, CustomerPrimaryContactEnd As Integer
        Dim Location As String, LocationStart As Integer, LocationEnd As Integer
        Dim PostalCode As String, PostalCodeStart As Integer, PostalCodeEnd As Integer
        Dim IncidentContact As String, IncidentContactStart As Integer, IncidentContactEnd As Integer
        Dim BusinessPhone As String, BusinessPhoneStart As Integer, BusinessPhoneEnd As Integer
        Dim CellPhone As String, CellPhoneStart As Integer, CellPhoneEnd As Integer
        Dim Fax As String, FaxStart As Integer, FaxEnd As Integer
        Dim Pager As String, PagerStart As Integer, PagerEnd As Integer
        Dim EmailAddress As String, EmailAddressStart As Integer, EmailAddressEnd As Integer
        Dim SupportLevel As String, SupportLevelStart As Integer, SupportLevelEnd As Integer
        Dim IncidentType As String, IncidentTypeStart As Integer, IncidentTypeEnd As Integer
        Dim ProductStart As Integer, ProductEnd As Integer ' Product As String, 
        Dim Version As String, VersionStart As Integer, VersionEnd As Integer
        Dim SerialNumber As String, SerialNumberStart As Integer, SerialNumberEnd As Integer
        Dim SubmittedBy As String, SubmittedByStart As Integer, SubmittedByEnd As Integer
        Dim WorkNotes As String, WorkNotesStart As Integer, WorkNotesEnd As Integer

        ' parse message body...Verint's emails are formatted the same regardless of them being a new or an update request

        'Priority:  4 - Low
        PriorityStart = InStr(TicketEmailBody, "Priority:") + 9
        If (PriorityStart > 9) Then
            PriorityEnd = InStr(PriorityStart, TicketEmailBody, vbCrLf)
            If (PriorityEnd > 0) Then
                Priority = Trim(Mid(TicketEmailBody, PriorityStart, PriorityEnd - PriorityStart))
            End If
        End If

        'Company Name:  Hartford, The (Server #1) - Charlotte, NC
        CompanyNameStart = InStr(TicketEmailBody, "Company Name:") + 13
        If (CompanyNameStart > 13) Then
            ' if they don't submit the customer name and site name together (no " - " then), we have to handle it...
            CompanyNameEndA = InStr(CompanyNameStart, TicketEmailBody, " - ")
            CompanyNameEndB = InStr(CompanyNameStart, TicketEmailBody, vbCrLf)
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
                        SiteNameEnd = InStr(SiteNameStart, TicketEmailBody, vbCrLf)
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
            OnyxCustomerNumberEnd = InStr(OnyxCustomerNumberStart, TicketEmailBody, vbCrLf)
            If (OnyxCustomerNumberEnd > 0) Then
                OnyxCustomerNumber = Trim(Mid(TicketEmailBody, OnyxCustomerNumberStart, OnyxCustomerNumberEnd - OnyxCustomerNumberStart))
            End If
        End If

        'Onyx Incident Number:  3810615
        OnyxIncidentNumberStart = InStr(TicketEmailBody, "Onyx Incident Number:") + 21
        If (OnyxIncidentNumberStart > 21) Then
            OnyxIncidentNumberEnd = InStr(OnyxIncidentNumberStart, TicketEmailBody, vbCrLf)
            If (OnyxIncidentNumberEnd > 0) Then
                OnyxIncidentNumber = Trim(Mid(TicketEmailBody, OnyxIncidentNumberStart, OnyxIncidentNumberEnd - OnyxIncidentNumberStart))
            End If
        End If

        'Customer Primary Contact:  Frank Beatty
        CustomerPrimaryContactStart = InStr(TicketEmailBody, "Customer Primary Contact:") + 25
        If (CustomerPrimaryContactStart > 25) Then
            CustomerPrimaryContactEnd = InStr(CustomerPrimaryContactStart, TicketEmailBody, vbCrLf)
            If (CustomerPrimaryContactEnd > 0) Then
                CustomerPrimaryContact = Trim(Mid(TicketEmailBody, CustomerPrimaryContactStart, CustomerPrimaryContactEnd - CustomerPrimaryContactStart))
            End If
        End If

        'Location:  Charlotte Service Ctr, 8711 University East Dr, Charlotte, North Carolina, United States
        LocationStart = InStr(TicketEmailBody, "Location:") + 9
        If (LocationStart > 9) Then
            LocationEnd = InStr(LocationStart, TicketEmailBody, vbCrLf)
            If (LocationEnd > 0) Then
                Location = Trim(Mid(TicketEmailBody, LocationStart, LocationEnd - LocationStart))
            End If
        End If

        'Postal Code:  28213
        PostalCodeStart = InStr(TicketEmailBody, "Postal Code:") + 12
        If (PostalCodeStart > 12) Then
            PostalCodeEnd = InStr(PostalCodeStart, TicketEmailBody, vbCrLf)
            If (PostalCodeEnd > 0) Then
                PostalCode = Trim(Mid(TicketEmailBody, PostalCodeStart, PostalCodeEnd - PostalCodeStart))
            End If
        End If

        'Incident Contact:  Steven Hudak
        IncidentContactStart = InStr(TicketEmailBody, "Incident Contact:") + 17
        If (IncidentContactStart > 17) Then
            IncidentContactEnd = InStr(IncidentContactStart, TicketEmailBody, vbCrLf)
            If (IncidentContactEnd > 0) Then
                IncidentContact = Trim(Mid(TicketEmailBody, IncidentContactStart, IncidentContactEnd - IncidentContactStart))
            End If
        End If

        'Business Phone:  8609197122
        BusinessPhoneStart = InStr(TicketEmailBody, "Business Phone:") + 15
        If (BusinessPhoneStart > 15) Then
            BusinessPhoneEnd = InStr(BusinessPhoneStart, TicketEmailBody, vbCrLf)
            If (BusinessPhoneEnd > 0) Then
                BusinessPhone = Trim(Mid(TicketEmailBody, BusinessPhoneStart, BusinessPhoneEnd - BusinessPhoneStart))
            End If
        End If

        'Cell Phone:  8609197122
        CellPhoneStart = InStr(TicketEmailBody, "Cell Phone:") + 11
        If (CellPhoneStart > 11) Then
            CellPhoneEnd = InStr(CellPhoneStart, TicketEmailBody, vbCrLf)
            If (CellPhoneEnd > 0) Then
                CellPhone = Trim(Mid(TicketEmailBody, CellPhoneStart, CellPhoneEnd - CellPhoneStart))
            End If
        End If

        'Fax: 8609197122
        FaxStart = InStr(TicketEmailBody, "Fax:") + 4
        If (FaxStart > 4) Then
            FaxEnd = InStr(FaxStart, TicketEmailBody, vbCrLf)
            If (FaxEnd > 0) Then
                Fax = Trim(Mid(TicketEmailBody, FaxStart, FaxEnd - FaxStart))
            End If
        End If

        'Pager: 8609197122
        PagerStart = InStr(TicketEmailBody, "Pager:") + 6
        If (PagerStart > 6) Then
            PagerEnd = InStr(PagerStart, TicketEmailBody, vbCrLf)
            If (PagerEnd > 0) Then
                Pager = Trim(Mid(TicketEmailBody, PagerStart, PagerEnd - PagerStart))
            End If
        End If

        'Email Address:  steven.hudak@thehartford.com
        EmailAddressStart = InStr(TicketEmailBody, "Email Address:") + 14
        If (EmailAddressStart > 14) Then
            EmailAddressEnd = InStr(EmailAddressStart, TicketEmailBody, vbCrLf)
            If (EmailAddressEnd > 0) Then
                EmailAddress = Trim(Mid(TicketEmailBody, EmailAddressStart, EmailAddressEnd - EmailAddressStart))
            End If
        End If

        'Support Level:  Advantage
        SupportLevelStart = InStr(TicketEmailBody, "Support Level:") + 14
        If (SupportLevelStart > 14) Then
            SupportLevelEnd = InStr(SupportLevelStart, TicketEmailBody, vbCrLf)
            If (SupportLevelEnd > 0) Then
                SupportLevel = Trim(Mid(TicketEmailBody, SupportLevelStart, SupportLevelEnd - SupportLevelStart))
            End If
        End If

        'Incident Type:  Web - Unknown
        IncidentTypeStart = InStr(TicketEmailBody, "Incident Type:") + 14
        If (IncidentTypeStart > 14) Then
            IncidentTypeEnd = InStr(IncidentTypeStart, TicketEmailBody, vbCrLf)
            If (IncidentTypeEnd > 0) Then
                IncidentType = Trim(Mid(TicketEmailBody, IncidentTypeStart, IncidentTypeEnd - IncidentTypeStart))
            End If
        End If

        'Product:  eQContactStore
        ProductStart = InStr(TicketEmailBody, "Product:") + 8
        If (ProductStart > 8) Then
            ProductEnd = InStr(ProductStart, TicketEmailBody, vbCrLf)
            If (ProductEnd > 0) Then
                Product = Trim(Mid(TicketEmailBody, ProductStart, ProductEnd - ProductStart))
            End If
        End If

        'Version:  CS - 7.2
        VersionStart = InStr(TicketEmailBody, "Version:") + 8
        If (VersionStart > 8) Then
            VersionEnd = InStr(VersionStart, TicketEmailBody, vbCrLf)
            If (VersionEnd > 0) Then
                Version = Trim(Mid(TicketEmailBody, VersionStart, VersionEnd - VersionStart))
            End If
        End If

        'Serial Number: RN-XXXXXX
        SerialNumberStart = InStr(TicketEmailBody, "Serial Number:") + 14
        If (SerialNumberStart > 14) Then
            SerialNumberEnd = InStr(SerialNumberStart, TicketEmailBody, vbCrLf)
            If (SerialNumberEnd > 0) Then
                SerialNumber = Trim(Mid(TicketEmailBody, SerialNumberStart, SerialNumberEnd - SerialNumberStart))
            End If
        End If

        'Submitted By:  Janice Wells
        SubmittedByStart = InStr(TicketEmailBody, "Submitted By:") + 13
        If (SubmittedByStart > 13) Then
            SubmittedByEnd = InStr(SubmittedByStart, TicketEmailBody, vbCrLf)
            If (SubmittedByEnd > 0) Then
                SubmittedBy = Trim(Mid(TicketEmailBody, SubmittedByStart, SubmittedByEnd - SubmittedByStart))
            End If
        End If

        'WorkNotes: 
        WorkNotesStart = InStr(TicketEmailBody, "WorkNotes:") + 10
        If (WorkNotesStart > 10) Then
            WorkNotesEnd = Len(TicketEmailBody)
            If (WorkNotesEnd > 0) Then
                WorkNotes = Trim(Mid(TicketEmailBody, WorkNotesStart, WorkNotesEnd - WorkNotesStart))
                ' remove up to the first 3 carriage returns so that the problem description will show!
                If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                End If
                If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                End If
                If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                End If
            End If
        End If

        'Write("CIM Service - Parsed an Onyx integration email Message (ID: " & MsgID & ")", "Priority: '" & Priority & "'" & vbCrLf & "Company Name: '" & CompanyName & "'" & vbCrLf & "Site Name: '" & SiteName & "'" & vbCrLf & "Onyx Customer Number: '" & OnyxCustomerNumber & "'" & vbCrLf & "Onyx Incident Number: '" & OnyxIncidentNumber & "'" & vbCrLf & "Customer Primary Contact: '" & CustomerPrimaryContact & "'" & vbCrLf & "Location: '" & Location & "'" & vbCrLf & "Postal Code: '" & PostalCode & "'" & vbCrLf & "Incident Contact: '" & IncidentContact & "'" & vbCrLf & "Business Phone: '" & BusinessPhone & "'" & vbCrLf & "Cell Phone: '" & CellPhone & "'" & vbCrLf & "Fax: '" & Fax & "'" & vbCrLf & "Pager: '" & Pager & "'" & vbCrLf & "Email Address: '" & EmailAddress & "'" & vbCrLf & "Support Level: '" & SupportLevel & "'" & vbCrLf & "Incident Type: '" & IncidentType & "'" & vbCrLf & "Product: '" & Product & "'" & vbCrLf & "Version: '" & Version & "'" & vbCrLf & "Serial Number: '" & SerialNumber & "'" & vbCrLf & "Submitted By: '" & SubmittedBy & "'" & vbCrLf & "WorkNotes: '" & WorkNotes & "'", 0)

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

        Dim SQLString As String
        If IntegrationRequestType <> "NEW" Then
            ' Get old stuff to compare...need to compare against what was originally submitted by the Partner
            Dim ds As DataSet = New DataSet, dt As DataTable = New DataTable
            Dim OldCustomerName As String, OldStatus As String, OldSiteName As String, OldSN As String, Old As String, OldSeverity_Priority As String, OldProduct As String, OldCallerName As String
            Dim OldCallerPhone As String, OldCallerEmail As String, OldContactName As String, OldContactPhone As String, OldContactEmail As String

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
                        "   PartnerTicketNo, CustomerName, Status, SiteName, SN, Severity_Priority, Product, CallerName, CallerPhone, CallerEmail, ContactName, ContactPhone, ContactEmail, LastChange" & _
                        ") VALUES (" & _
                        "   '" & Replace(PartnerTicketNo, "'", "''") & "', '" & Replace(CustomerName, "'", "''") & "', '', '" & Replace(SiteName, "'", "''") & "', '" & Replace(SN, "'", "''") & "', '" & Priority & "', '" & Replace(Product, "'", "''") & "', '" & Replace(CallerName, "'", "''") & "'," & _
                        "   '" & Replace(CallerPhone, "'", "''") & "', '" & Replace(CallerEmail, "'", "''") & "', '" & Replace(ContactName, "'", "''") & "', '" & Replace(ContactPhone, "'", "''") & "', '" & Replace(ContactEmail, "'", "''") & "', GetDate()" & _
                        ")"

            .ExecuteSQL(SQLString)
            'Write("CIM Service - Inserted PartnerIntegrationEmailData Record: " & PartnerTicketNo & " Event", SQLString, 0)
        End With
    End Sub

    Sub ParseVerintOracleEmail(ByVal TicketEmailBody As String, ByVal MsgID As String, ByVal IntegrationRequestType As String)
        Dim SRSeverity As String, SRSeverityStart As Integer, SRSeverityEnd As Integer, SRSeverityDescription As String
        Dim CustomerNameStart As Integer, CustomerNameEnd As Integer ' CustomerName As String, 
        Dim SiteNameStart As Integer, SiteNameEnd As Integer ' SiteName As String, 
        Dim CustomerNumber As String, CustomerNumberStart As Integer, CustomerNumberEnd As Integer
        Dim IncidentNumber As String, IncidentNumberStart As Integer, IncidentNumberEnd As Integer
        Dim Location As String, LocationStart As Integer, LocationEnd As Integer
        Dim PostalCode As String, PostalCodeStart As Integer, PostalCodeEnd As Integer
        Dim IncidentContact As String, IncidentContactStart As Integer, IncidentContactEnd As Integer
        Dim IncidentPrimaryContactPhone As String, IncidentPrimaryContactPhoneStart As Integer, IncidentPrimaryContactPhoneEnd As Integer
        Dim IncidentPrimaryContactEmail As String, IncidentPrimaryContactEmailStart As Integer, IncidentPrimaryContactEmailEnd As Integer
        Dim SupportLevel As String, SupportLevelStart As Integer, SupportLevelEnd As Integer
        Dim SRType As String, SRTypeStart As Integer, SRTypeEnd As Integer
        Dim ServiceProductVertical As String, ServiceProductVerticalStart As Integer, ServiceProductVerticalEnd As Integer
        Dim SerialNumber As String, SerialNumberStart As Integer, SerialNumberEnd As Integer
        Dim TaskStatus As String, TaskStatusStart As Integer, TaskStatusEnd As Integer
        Dim LoggedBy As String, LoggedByStart As Integer, LoggedByEnd As Integer
        Dim WorkNotes As String, WorkNotesStart As Integer, WorkNotesEnd As Integer

        ' parse message body...Verint's emails are formatted the same regardless of them being a new or an update request
        'SR Severity         : VRNT P3
        SRSeverityStart = InStr(TicketEmailBody, "SR Severity         :") + 21
        If (SRSeverityStart > 21) Then
            SRSeverityEnd = InStr(SRSeverityStart, TicketEmailBody, vbCrLf)
            If (SRSeverityEnd > 0) Then
                SRSeverity = Trim(Mid(TicketEmailBody, SRSeverityStart, SRSeverityEnd - SRSeverityStart))
            End If
        End If

        'Customer Name       : VERINT GMBH
        CustomerNameStart = InStr(TicketEmailBody, "Customer Name       :") + 21
        If (CustomerNameStart > 21) Then
            CustomerNameEnd = InStr(CustomerNameStart, TicketEmailBody, vbCrLf)
            If (CustomerNameEnd > 0) Then
                CustomerName = Trim(Mid(TicketEmailBody, CustomerNameStart, CustomerNameEnd - CustomerNameStart))
            End If
        End If

        'Site Name           : VERINT GMBH KARLSRUHE,  D-76139 DE_Site
        SiteNameStart = InStr(TicketEmailBody, "Site Name           :") + 21
        If (SiteNameStart > 21) Then
            SiteNameEnd = InStr(SiteNameStart, TicketEmailBody, vbCrLf)
            If (SiteNameEnd > 0) Then
                SiteName = Trim(Mid(TicketEmailBody, SiteNameStart, SiteNameEnd - SiteNameStart))
            End If
        End If

        'Customer Number     : 295464
        CustomerNumberStart = InStr(TicketEmailBody, "Customer Number     :") + 21
        If (CustomerNumberStart > 21) Then
            CustomerNumberEnd = InStr(CustomerNumberStart, TicketEmailBody, vbCrLf)
            If (CustomerNumberEnd > 0) Then
                CustomerNumber = Trim(Mid(TicketEmailBody, CustomerNumberStart, CustomerNumberEnd - CustomerNumberStart))
            End If
        End If

        'Incident Number     : 1445980
        IncidentNumberStart = InStr(TicketEmailBody, "Incident Number     :") + 21
        If (IncidentNumberStart > 21) Then
            IncidentNumberEnd = InStr(IncidentNumberStart, TicketEmailBody, vbCrLf)
            If (IncidentNumberEnd > 0) Then
                IncidentNumber = Trim(Mid(TicketEmailBody, IncidentNumberStart, IncidentNumberEnd - IncidentNumberStart))
            End If
        End If

        'Location            : AM STORRENACKER 2, , KARLSRUHE, D-76139  DE
        LocationStart = InStr(TicketEmailBody, "Location            :") + 21
        If (LocationStart > 21) Then
            LocationEnd = InStr(LocationStart, TicketEmailBody, vbCrLf)
            If (LocationEnd > 0) Then
                Location = Trim(Mid(TicketEmailBody, LocationStart, LocationEnd - LocationStart))
            End If
        End If

        'Postal Code         : D-76139
        PostalCodeStart = InStr(TicketEmailBody, "Postal Code         :") + 21
        If (PostalCodeStart > 21) Then
            PostalCodeEnd = InStr(PostalCodeStart, TicketEmailBody, vbCrLf)
            If (PostalCodeEnd > 0) Then
                PostalCode = Trim(Mid(TicketEmailBody, PostalCodeStart, PostalCodeEnd - PostalCodeStart))
            End If
        End If

        'Incident Primary Contact First name    : 
        IncidentContactStart = InStr(TicketEmailBody, "Incident Primary Contact First name    :") + 40
        If (IncidentContactStart > 40) Then
            IncidentContactEnd = InStr(IncidentContactStart, TicketEmailBody, vbCrLf)
            If (IncidentContactEnd > 0) Then
                IncidentContact = Trim(Mid(TicketEmailBody, IncidentContactStart, IncidentContactEnd - IncidentContactStart))
            End If
        End If

        'Incident Primary Contact Last name     : 
        IncidentContactStart = InStr(TicketEmailBody, "Incident Primary Contact Last name     :") + 40
        If (IncidentContactStart > 40) Then
            IncidentContactEnd = InStr(IncidentContactStart, TicketEmailBody, vbCrLf)
            If (IncidentContactEnd > 0) Then
                IncidentContact = IncidentContact & " " & Trim(Mid(TicketEmailBody, IncidentContactStart, IncidentContactEnd - IncidentContactStart))
            End If
        End If

        'Incident Primary Contact Phone         : 
        IncidentPrimaryContactPhoneStart = InStr(TicketEmailBody, "Incident Primary Contact Phone         :") + 40
        If (IncidentPrimaryContactPhoneStart > 40) Then
            IncidentPrimaryContactPhoneEnd = InStr(IncidentPrimaryContactPhoneStart, TicketEmailBody, vbCrLf)
            If (IncidentPrimaryContactPhoneEnd > 0) Then
                IncidentPrimaryContactPhone = Trim(Mid(TicketEmailBody, IncidentPrimaryContactPhoneStart, IncidentPrimaryContactPhoneEnd - IncidentPrimaryContactPhoneStart))
            End If
        End If

        'Incident Primary Contact Email         : 
        IncidentPrimaryContactEmailStart = InStr(TicketEmailBody, "Incident Primary Contact Email         :") + 40
        If (IncidentPrimaryContactEmailStart > 40) Then
            IncidentPrimaryContactEmailEnd = InStr(IncidentPrimaryContactEmailStart, TicketEmailBody, vbCrLf)
            If (IncidentPrimaryContactEmailEnd > 0) Then
                IncidentPrimaryContactEmail = Trim(Mid(TicketEmailBody, IncidentPrimaryContactEmailStart, IncidentPrimaryContactEmailEnd - IncidentPrimaryContactEmailStart))
            End If
        End If

        'Support Level       : 
        SupportLevelStart = InStr(TicketEmailBody, "Support Level       :") + 21
        If (SupportLevelStart > 21) Then
            SupportLevelEnd = InStr(SupportLevelStart, TicketEmailBody, vbCrLf)
            If (SupportLevelEnd > 0) Then
                SupportLevel = Trim(Mid(TicketEmailBody, SupportLevelStart, SupportLevelEnd - SupportLevelStart))
            End If
        End If

        'SR Type             : VRNT System Malfunction
        SRTypeStart = InStr(TicketEmailBody, "SR Type             :") + 21
        If (SRTypeStart > 21) Then
            SRTypeEnd = InStr(SRTypeStart, TicketEmailBody, vbCrLf)
            If (SRTypeEnd > 0) Then
                SRType = Trim(Mid(TicketEmailBody, SRTypeStart, SRTypeEnd - SRTypeStart))
            End If
        End If

        'Service ServiceProductVertical Vertical : Next Generation (NG) WAS ServiceProductVertical Version : VI360
        ServiceProductVerticalStart = InStr(TicketEmailBody, "Service ServiceProductVertical Vertical :") + 21
        If (ServiceProductVerticalStart > 21) Then
            ServiceProductVerticalEnd = InStr(ServiceProductVerticalStart, TicketEmailBody, vbCrLf)
            If (ServiceProductVerticalEnd > 0) Then
                ServiceProductVertical = Trim(Mid(TicketEmailBody, ServiceProductVerticalStart, ServiceProductVerticalEnd - ServiceProductVerticalStart))
            End If
        End If

        'Serial Number       : 
        SerialNumberStart = InStr(TicketEmailBody, "Serial Number       :") + 21
        If (SerialNumberStart > 21) Then
            SerialNumberEnd = InStr(SerialNumberStart, TicketEmailBody, vbCrLf)
            If (SerialNumberEnd > 0) Then
                SerialNumber = Trim(Mid(TicketEmailBody, SerialNumberStart, SerialNumberEnd - SerialNumberStart))
            End If
        End If

        'Task Status         : VRNT Closed
        TaskStatusStart = InStr(TicketEmailBody, "Task Status         :") + 21
        If (TaskStatusStart > 21) Then
            TaskStatusEnd = InStr(TaskStatusStart, TicketEmailBody, vbCrLf)
            If (TaskStatusEnd > 0) Then
                TaskStatus = Trim(Mid(TicketEmailBody, TaskStatusStart, TaskStatusEnd - TaskStatusStart))
            End If
        End If

        'Logged By           : NSINGH
        LoggedByStart = InStr(TicketEmailBody, "Logged By           : ") + 21
        If (LoggedByStart > 13) Then
            LoggedByEnd = InStr(LoggedByStart, TicketEmailBody, vbCrLf)
            If (LoggedByEnd > 0) Then
                LoggedBy = Trim(Mid(TicketEmailBody, LoggedByStart, LoggedByEnd - LoggedByStart))
            End If
        End If

        'Work Notes          : 
        WorkNotesStart = InStr(TicketEmailBody, "Work Notes          :") + 21
        If (WorkNotesStart > 10) Then
            WorkNotesEnd = Len(TicketEmailBody)
            If (WorkNotesEnd > 0) Then
                WorkNotes = Trim(Mid(TicketEmailBody, WorkNotesStart, WorkNotesEnd - WorkNotesStart))
                ' remove up to the first 3 carriage returns so that the problem description will show!
                If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                End If
                If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                End If
                If (Mid(WorkNotes, 1, 2) = vbCrLf) Then
                    WorkNotes = Mid(WorkNotes, 3, Len(WorkNotes))
                End If
            End If
        End If

        Write("CIM Service - Parsed an Oracle integration email Message (ID: " & MsgID & ")", "SR Severity: '" & SRSeverity & "'" & vbCrLf & "Customer Name: '" & CustomerName & "'" & vbCrLf & "Site Name: '" & SiteName & "'" & vbCrLf & "Customer Number: '" & CustomerNumber & "'" & vbCrLf & "Incident Number: '" & IncidentNumber & "'" & vbCrLf & "Location: '" & Location & "'" & vbCrLf & "Postal Code: '" & PostalCode & "'" & vbCrLf & "Incident Contact: '" & IncidentContact & "'" & vbCrLf & "Incident Primary Contact Phone: '" & IncidentPrimaryContactPhone & "'" & vbCrLf & "Incident Primary Contact Email: '" & IncidentPrimaryContactEmail & "'" & vbCrLf & "Support Level: '" & SupportLevel & "'" & vbCrLf & "SR Type: '" & SRType & "'" & vbCrLf & "Service Product Vertical: '" & ServiceProductVertical & "'" & vbCrLf & "Serial Number: '" & SerialNumber & "'" & vbCrLf & "Task Status: '" & TaskStatus & "'" & vbCrLf & "Logged By: '" & LoggedBy & "'" & vbCrLf & "Work Notes: '" & WorkNotes & "'", 0)

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

        Dim SQLString As String
        If IntegrationRequestType <> "NEW" Then
            ' Get old stuff to compare...need to compare against what was originally submitted by the Partner
            Dim ds As DataSet = New DataSet, dt As DataTable = New DataTable
            Dim OldCustomerName As String, OldStatus As String, OldSiteName As String, OldSN As String, Old As String, OldSeverity_Priority As String, OldProduct As String, OldCallerName As String
            Dim OldCallerPhone As String, OldCallerEmail As String, OldContactName As String, OldContactPhone As String, OldContactEmail As String

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
        End If

        With DManager
            SQLString = "INSERT INTO PartnerIntegrationEmailData (" & _
                        "   PartnerTicketNo, CustomerName, Status, SiteName, SN, Severity_Priority, Product, CallerName, CallerPhone, CallerEmail, ContactName, ContactPhone, ContactEmail, LastChange" & _
                        ") VALUES (" & _
                        "   '" & Replace(PartnerTicketNo, "'", "''") & "', '" & Replace(CustomerName, "'", "''") & "', '" & Replace(TaskStatus, "'", "''") & "', '" & Replace(SiteName, "'", "''") & "', '" & Replace(SN, "'", "''") & "', '" & SRSeverity & "', '" & Replace(Product, "'", "''") & "', '" & Replace(CallerName, "'", "''") & "'," & _
                        "   '" & Replace(CallerPhone, "'", "''") & "', '" & Replace(CallerEmail, "'", "''") & "', '" & Replace(ContactName, "'", "''") & "', '" & Replace(ContactPhone, "'", "''") & "', '" & Replace(ContactEmail, "'", "''") & "', GetDate() " & _
                        ")"

            .ExecuteSQL(SQLString)
            'Write("CIM Service - Inserted PartnerIntegrationEmailData Record: " & PartnerTicketNo & " Event", SQLString, 0)
        End With
    End Sub

    Public Function GetNextTicketRecId() As Long
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        Dim sql As String

        GetNextTicketRecId = 1

        sql = "SELECT TOP 1 ISNULL(RecID, 1000000000) + 1 as NextRecID FROM C_Tickets WHERE RECID < 5000000000 order by recid desc"

        Try
            With DManager
                ds = .GetDataSet(sql)
                If Not ds Is Nothing Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        GetNextTicketRecId = ds.Tables(0).Rows(0).Item("NextRecID")
                    Else
                        GetNextTicketRecId = 1000000000
                    End If
                Else
                    GetNextTicketRecId = 1000000000
                End If
            End With

        Catch ex As Exception
            Write("Witness Quote Process - " & "Procedure:GetNextTicketRecId ", ex.Message, 2)
        End Try

    End Function

    Public Function GetNextTicketId(ByVal DataAreaID As String) As String
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        Dim sql As String

        Dim NextTicketIdNum As String
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
        Dim sql As String = "select nextrec From NumberSequenceTable where dataareaid = '" & DataAreaID & "' and numbersequence = 'CC_Tickets'"

        GetTicketNumberSequence = 1
        Try
            With DManager
                ds = .GetDataSet(sql)
                If Not ds Is Nothing Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        GetTicketNumberSequence = ds.Tables(0).Rows(0).Item("nextrec")

                        sql = "UPDATE NumberSequenceTable SET NextRec = " & GetTicketNumberSequence + 1 & " Where dataareaid = '" & DataAreaID & "' and numbersequence = 'CC_Tickets'"
                        .ExecuteSQL(sql)
                    End If
                End If
            End With
        Catch ex As Exception
            Write("Witness Quote Process - " & "Procedure:GetTicketNumberSequence ", ex.Message, 2)
        End Try

    End Function

    Public Function GetNextTicketEventRecId() As Long
        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        Dim sql As String

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
            Write("Witness Quote Process - " & "Procedure:GetNextTicketEventRecId ", ex.Message, 2)
        End Try

    End Function

    Sub SendUpdateTicketNotFoundEmail(ByVal msg As MailMessage, ByVal Partner As String, ByVal PartnerTicketNo As String)
        Dim mailMsg As MailMessage = New MailMessage
        mailMsg.From.EMail = "Support@adtechglobal.com"

        Dim client As SMTP = New SMTP
        client.Host = "inrelay.it.adtech.net"
        client.Port = 25

        Try
            mailMsg.To.Add("Arnett Kelly", "akelly@adtechglobal.com")
            mailMsg.To.Add("Brian Brazil", "bbrazil@adtechglobal.com")
            mailMsg.Subject = "DROPPED " & Partner & " Update Email (incident #: " & PartnerTicketNo & ")"
            mailMsg.Body = Partner & " incident # '" & PartnerTicketNo & "' was not found in the AGS support ticket system. " & vbCrLf & vbCrLf & _
                            "AGS Automated System Message" & vbCrLf & vbCrLf & _
                            "---- Original message ----" & vbCrLf & vbCrLf & msg.PlainMessage.Body.ToString()

            client.SendMessage(mailMsg)

            'Write("CIM Service - " & mailMsg.Subject, mailMsg.Body, 0)
        Catch ex As Exception
            Write("CIM Service:SendReplyEmail function", ex.Message, 2)
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

            client.SendMessage(mailMsg)

            'Write("CIM Service", "Reply email sent for Message ID: " & msg.MessageID & " has been sent to: " & mailMsg.To.ToString() & ".", 0)
        Catch ex As Exception
            Write("CIM Service:SendReplyEmail function", ex.Message, 2)
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
        Catch ex As Exception
            Write("CIM Service:SendReplyEmail function", ex.Message, 2)
        End Try


    End Sub

    Public Function FindMessageID(ByVal MessageID As String) As Boolean
        Dim sql As String
        Dim strMsgID As String

        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        FindMessageID = False

        sql = "Select MessageID from MessageIDs " & _
              "where " & _
              "MessageID = '" & MessageID & "'"

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
        Catch ex As Exception
            Write("CIM Service - " & "Procedure:FindMessageID ", ex.Message, 2)
        End Try

        Return FindMessageID
    End Function

    Public Function UpdateMessageID(ByVal MessageID As String) As Boolean
        Dim sql As String
        Try
            With DManager
                sql = "Insert into MessageIDs " & _
                            "(MessageID, CreationDateTime) " & _
                            "Values " & _
                            "(" & _
                            "'" & MessageID & "', " & _
                            "GetDate() " & _
                            ");"

                .ExecuteSQL(sql)
            End With
        Catch ex As Exception
            Write("CIM Service - " & "Procedure:UpdateMessageID ", ex.Message, 2)
        End Try
    End Function

    Private Function OpenDB() As Boolean
        OpenDB = False
        Try
            'ODBC Connection String - SAVE THIS!
            Dim ConnString As String = "Driver={SQL Server};" & "Server=" & DatabaseServer & ";" & _
                             "Database=" & DatabaseName & ";" & "Uid=" & DatabaseUserName & ";" & _
                             "Pwd=" & DatabasePassword & "; "

            DManager = New DataManager(ConnString)
            With DManager
                If .OpenConnection Then
                    OpenDB = True
                End If
            End With
        Catch ex As Exception
            Write("CIM Service", ex.Message, 2)
        End Try
    End Function

    Private Function CloseDB() As Boolean
        With DManager
            .CloseConnection()
        End With
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

