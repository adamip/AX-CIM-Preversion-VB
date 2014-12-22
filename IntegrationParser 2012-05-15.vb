REM Major Modification                                                                                                  Adam Ip 2011-09-28 
REM     Rewrite all data field extraction by creating
REM         Function ExtractSingleLineData
REM         Private Function ExtractWorkNotes
REM     Write Function FindNonZeroMin 
REM
REM
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
    Public Shared Inform_Mode As Boolean = False
    Public Shared Summary_Inform_Mode As Boolean = True

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
        REM Username cannot contains upper case character
        Inbox(0).Username = "emeasupport"
        Inbox(0).Password = "zZ3k.D&tk9bew"
        Inbox(0).Host = "INPOP.it.adtech.net"
        TakeABreak(0) = False
        DataAreaIDcode(1) = "US"
        Inbox(1) = New POP3()
        REM Username cannot contains upper case character
        Inbox(1).Username = "supportcim"
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
                                                & "Data Area ID: " & DataAreaID & vbCrLf & vbCrLf _
                                                & "Message ID: " & msg.MessageID.ToString() & vbCrLf _
                                                & "From: " & msg.From.ToString() & vbCrLf _
                                                & "To: " & msg.To.ToString() & vbCrLf _
                                                & "Subject: " & msg.Subject.ToString() & vbCrLf _
                                                & "Date: " & msg.Date.ToString() & vbCrLf _
                                                & "Body: " & vbCrLf & vbTab & "---- Beginning of lines ----" & vbCrLf _
                                                & msg.PlainMessage.Body.ToString() & vbCrLf & vbTab _
                                                & "---- End of lines ----"
                                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                        End If

                                        'if e-mail is NOT from Verint nor aip@adtech.net then ValidateEmailSource(msg.From.ToString()) returns false
                                        '   then simply only update the MessageIDs table through UpdateMessageID()
                                        'if ValidateEmailSource(msg.From.ToString()) returns true, then email is either from  
                                        '    from Verint nor aip@adtech.net, then proceed with ProcessSupportMessage()
                                        'Don't combine these 2 following IF's into one IF
                                        '   If A = false Or B Then ...    the logic is different
                                        If ValidateEmailSource(msg.MessageID, msg.From.ToString(), TakeABreak(i)) = False Then
                                            UMI = UpdateMessageID(msg.MessageID, TakeABreak(i))
                                            If Debug_Mode Then
                                                sEventEntry = "CIM Service: [" & DataAreaID & "] #06 Function ProcessEmailInbox, UpdateMessageID returns " & Convert.ToString(UMI)
                                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                            End If
                                        Else
                                            UMIIP = UpdateMessageIDinProgress(msg.MessageID, msg.From.ToString(), "A", TakeABreak(i))
                                            If Debug_Mode Then
                                                sEventEntry = "CIM Service: [" & DataAreaID & "] #07 Function ProcessEmailInbox" & vbCrLf _
                                                    & "UpdateMessageIDinProgress returns " & Convert.ToString(UMIIP)
                                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                            End If
                                            If UMIIP = 0 And TakeABreak(i) = False Then
                                                REM most processing mechanism happens in ProcessSupportMessage( )
                                                PSM = ProcessSupportMessage(msg, DataAreaID, TakeABreak(i))
                                                If Debug_Mode Then
                                                    sEventEntry = "CIM Service: [" & DataAreaID & "] #08 Function ProcessEmailInbox" & vbCrLf _
                                                        & "ProcessSupportMessage returns " & Convert.ToString(PSM)
                                                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                                End If
                                                If PSM = True Then
                                                    If TakeABreak(i) = False Then UMI = UpdateMessageID(msg.MessageID, TakeABreak(i))
                                                    If TakeABreak(i) = False Then UpdateMessageIDinProgress(msg.MessageID, msg.From.ToString(), "D", TakeABreak(i))
                                                    If Debug_Mode Then
                                                        sEventEntry = "CIM Service: [" & DataAreaID & "] #09 Function ProcessEmailInbox" & vbCrLf _
                                                         & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak(i)) & vbCrLf _
                                                         & "UMI: " & Convert.ToString(UMI) & vbCrLf _
                                                         & "UMIIP: " & Convert.ToString(UMIIP)
                                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                                    End If
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
                & "Data Area ID: " & DataAreaID & vbCrLf & vbCrLf _
                & "Message ID: " & msg.MessageID.ToString() & vbCrLf _
                & "From: " & msg.From.ToString() & vbCrLf _
                & "To: " & msg.To.ToString() & vbCrLf _
                & "Subject: " & msg.Subject.ToString() & vbCrLf _
                & "Date: " & msg.Date.ToString() & vbCrLf _
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
                If String.Compare(sIntegrationRequestType, "NEW", True) = 0 Then    'equals to zero when strA equals strB
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
                        If Debug_Mode Then
                            sEventEntry = "CIM Service: Function ProcessSupportMessage #11" & vbCrLf & "sTicketID: " & sTicketID
                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                        End If
                    Catch ex As Exception
                        TakeABreak = True
                        sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #12 returns " _
                            & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                            & "sTicketID: " & sTicketID & vbCrLf _
                            & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                            & "Exception Target: " & ex.TargetSite.ToString()
                        WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                        Exit Function
                    End Try

                    If TakeABreak Then Exit Function

                    'C_Tickets
                    If Debug_Mode Then
                        sEventEntry = "CIM Service: Function ProcessSupportMessage #13" & vbCrLf & "sTicketID: " & sTicketID
                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    End If

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
                        Catch ex As Exception
                            TakeABreak = True
                            sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #16 returns " _
                                & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                & "Exception Target: " & ex.TargetSite.ToString()
                            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                            Exit Function
                        End Try

                        Try
                            SQLString = "INSERT INTO C_Tickets (" _
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
                                        & " VALUES (" _
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
                        Catch ex As Exception
                            TakeABreak = True
                            sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #17 returns " _
                                & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                & "Exception Target: " & ex.TargetSite.ToString()
                            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                            Exit Function
                        End Try

                        Try
                            If Debug_Mode Or Inform_Mode Then
                                sEventEntry = "CIM Service: Function ProcessSupportMessage #18" & vbCrLf _
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
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #19", EventLogEntryType.Information)
                        Catch ex As Exception
                            TakeABreak = True
                            sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #20 returns " _
                                & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                & "Exception Target: " & ex.TargetSite.ToString()
                            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                            Exit Function
                        End Try
                        If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #21", EventLogEntryType.Information)
                        IncrementTicketNumberSequence(TakeABreak)
                    End With

                    If TakeABreak Then Exit Function
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #22", EventLogEntryType.Information)

                    'C_TicketEvents
                    NextTicketEventRecID = GetNextTicketEventRecID(TakeABreak)
                    If TakeABreak Then Exit Function

                    If Debug_Mode Then
                        sEventEntry = "CIM Service: Function ProcessSupportMessage #23" & vbCrLf _
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
                                sEventEntry = "CIM Service: Function ProcessSupportMessage #24" & vbCrLf _
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
                        sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #25 returns " _
                            & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                            & "SQL: " & SQLString & vbCrLf & vbCrLf _
                            & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                            & "Exception Target: " & ex.TargetSite.ToString()
                        WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                        Exit Function
                    End Try
                Else ' If String.Compare(sIntegrationRequestType, "NEW", True) <> 0 Then  
                    ' update(ticket)
                    ' get AGS ticket from partner's ticket number
                    Dim ds As DataSet = New DataSet
                    Dim dt As DataTable = New DataTable
                    Dim Status As String = ""
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #26", EventLogEntryType.Information)
                    Try

                        SQLString = "SELECT TOP 1 TicketID_Display, Status FROM C_Tickets WHERE PartnerTicketNo = '" & PartnerTicketNo _
                            & "' ORDER BY TicketID_Display DESC"

                        If Debug_Mode Then
                            sEventEntry = "CIM Service: Function ProcessSupportMessage #27" & vbCrLf _
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
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #28", EventLogEntryType.Information)
                                    sTicketID = dt.Rows(0).Item("TicketID_Display")
                                    Status = dt.Rows(0).Item("Status")
                                End If
                            End If
                            If Debug_Mode Then
                                sEventEntry = "CIM Service: Function ProcessSupportMessage #29" & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                    & "sTicketID: " & sTicketID & vbCrLf _
                                    & "Status: " & Status & vbCrLf & vbCrLf _
                                    & "SQL executed: " & SQLString
                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            End If
                        End With
                    Catch ex As Exception
                        TakeABreak = True
                        sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #30 returns " _
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
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #31", EventLogEntryType.Information)
                    If Len(sTicketID) = 0 Then ' send email to Brian Brazil that an update request was recieved but no ticket was found internally
                        SendUpdateTicketNotFoundEmail(msg, Partner3_, PartnerTicketNo, DataAreaID)
                    Else ' process accordingly
                        NextTicketEventRecID = GetNextTicketEventRecID(TakeABreak)
                        If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #32", EventLogEntryType.Information)
                        If TakeABreak = True Or NextTicketEventRecID <= 0 Then Exit Function
                        If Status = "Closed" And RequestStatus <> "Close" Then ' Re-Open...change value on header, update appropriate fields and add Event entry.
                            Try
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), Status = 'Open', IsRead = 0 WHERE TicketID_Display = '" & sTicketID & "'"
                                    ' " & SQLTicketChangesString & "
                                    'eventViewerString &= vbCrLf & vbCrLf & "Re-Opened Ticket ID '" & TicketID & "' SQL: " & SQLString
                                    'Write("CIM Service - Ticket ID: " & TicketID & " Re-Opened", "SQL: " & SQLString, 0)
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #33" & vbCrLf _
                                            & "Ticket ID: " & sTicketID & " Re-opened" & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #34", EventLogEntryType.Information)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #35 returns " _
                                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #36", EventLogEntryType.Information)
                            Try
                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Ticket Re-Opened', 0, '" & sTicketID & "', 'Ticket Re-Opened due to an update from the partner.  Please review and process accordingly." _
                                        & vbCrLf & vbCrLf & Replace(SQLTicketChangesEventString, "'", "''") & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf _
                                        & Replace(sTicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" _
                                        & SeverityLevel & "', '" & DataAreaID & "', '" & NextTicketEventRecID & "', 1, 1, 0)"
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #37" & vbCrLf _
                                            & "Ticket Event ID: " & sTicketID & vbCrLf _
                                            & "Next Ticket Event Rec ID: " & Convert.ToString(NextTicketEventRecID) & " Re-opened" & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #38 returns " _
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
                                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #39", EventLogEntryType.Information)
                                With DManager
                                    SQLString = "UPDATE C_Tickets SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), IsRead = 0 WHERE TicketID_Display = '" & sTicketID & "'"
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #40 Ticket ID: " & sTicketID _
                                            & " Receive Closed Ticket per partner's request" & vbCrLf _
                                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #41", EventLogEntryType.Information)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #42 returns " _
                                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                                    & "Ticket ID: " & sTicketID & " Receive Closed Ticket per partner's request" & vbCrLf & vbCrLf _
                                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #43", EventLogEntryType.Information)
                            Try
                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Comment', 1, '" & sTicketID & "', 'Please note that it appears that this ticket may have been closed by the partner.  Please review and process accordingly." _
                                        & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(sTicketEmailBody, "'", "''") _
                                        & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel _
                                        & "', '" & DataAreaID & "', '" & NextTicketEventRecID & "', 1, 1, 0)"
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #44" & vbCrLf _
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
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #45 returns " _
                                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                    & "Ticket Event ID: " & Convert.ToString(NextTicketEventRecID) & " Receive Closed Ticket per partner's request" & vbCrLf & vbCrLf _
                                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #46", EventLogEntryType.Information)
                            REM Close request when our ticket was already closed...probably just a receipt confirmation from our partner.  No notification of ticket owner
                        ElseIf Status = "Closed" And RequestStatus = "Close" Then
                            Try
                                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #47", EventLogEntryType.Information)
                                ' no update needed to ticket header (C_Tickets)
                                'Write("CIM Service - Ticket ID: " & TicketID & " Closed per partners request", "Ticket was already closed internally", 0)
                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Comment', 1, '" & sTicketID & "', 'Please note that this ticket was closed by the partner, i.e. internal ticket already closed.  Review isn''t necessary." _
                                        & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf & Replace(sTicketEmailBody, "'", "''") _
                                        & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription & "', '" & SeverityLevel _
                                        & "', '" & DataAreaID _
                                        & "', '" & NextTicketEventRecID & "', 1, 1, 0)"
                                    REM eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' Event (close request on a closed ticket) SQL: " & SQLString
                                    REM Write("CIM Service - Inserted Ticket ID: " & TicketID & " Event", "SQL: " & SQLString, 0)
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #48" & vbCrLf _
                                            & "Closed per partner's request" & vbCrLf _
                                            & "Ticket ID: " & sTicketID & vbCrLf _
                                            & "Next Ticket Event Rec ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf _
                                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                            & "Ticket was already closed internally" & vbCrLf & vbCrLf & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #49", EventLogEntryType.Information)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #50 returns " _
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
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #51", EventLogEntryType.Information)
                        Else ' regular update request...insert and update accordingly
                            Try
                                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #52", EventLogEntryType.Information)
                                With DManager
                                    SQLString = "UPDATE C_Tickets " _
                                            & "SET ModifiedBy = 'MS AX', ModifiedDateTime = DATEADD(hh,4,getdate()), IsRead = 0 " _
                                            & "WHERE TicketID_Display = '" & sTicketID & "'"

                                    'eventViewerString &= vbCrLf & vbCrLf & "Received Update Ticket '" & TicketID & "' request, Ticket updated SQL: " & SQLString
                                    'Write("CIM Service - Updated Ticket ID: " & TicketID, SQLString, 0)
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #53" & vbCrLf _
                                            & "Receive updated Ticket ID: " & sTicketID & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #54 returns " _
                                & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                & "Updated Ticket ID: " & sTicketID & vbCrLf & vbCrLf _
                                & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #55", EventLogEntryType.Information)
                            Try
                                With DManager
                                    SQLString = "INSERT INTO C_TicketEvents (EventType, Private, TicketID, Comment_, CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, " _
                                        & "SeverityDescription, SeverityLevel, DataAreaID, RecID, RecVersion, Completed, PartnerUpdateSent) " _
                                        & "VALUES ('Comment', 1, '" & sTicketID & "', 'Partner Integration Update Email Received.  Please review and process accordingly." _
                                        & vbCrLf & vbCrLf & Replace(SQLTicketChangesEventString, "'", "''") & vbCrLf & vbCrLf & "---- Original message ----" & vbCrLf & vbCrLf _
                                        & Replace(sTicketEmailBody, "'", "''") & "', 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & SeverityLevelDescription _
                                        & "', '" & SeverityLevel & "', '" & DataAreaID & "', '" & NextTicketEventRecID & "', 1, 1, 0)"
                                    If Debug_Mode Or Inform_Mode Then
                                        sEventEntry = "CIM Service: Function ProcessSupportMessage #56 Inserted Ticket Event ID: " _
                                            & Convert.ToString(NextTicketEventRecID) & vbCrLf & vbCrLf _
                                            & "SQL to execute: " & SQLString
                                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                                    End If
                                    .ExecuteSQL(SQLString)
                                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #57", EventLogEntryType.Information)
                                End With
                            Catch ex As Exception
                                TakeABreak = True
                                sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #58 returns " _
                                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                                    & "Inserted Ticket Event ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf & vbCrLf _
                                    & "SQL: " & SQLString & vbCrLf & vbCrLf _
                                    & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                                    & "Exception Target: " & ex.TargetSite.ToString()
                                WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                                Exit Function
                            End Try
                            If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #59", EventLogEntryType.Information)
                        End If
                    End If
                End If
                ProcessSupportMessage = True    'returns

            Else ' process other partner ticket emails (non integration)
                Try
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #60", EventLogEntryType.Information)
                    ' Get the next ticket id.  Check if the number sequence table needs to be synched up too...
                    sTicketID = GetNextTicketId(DataAreaID, TakeABreak)
                    'If CInt(Replace(TicketID, "AGST", "")) + 1 > GetTicketNumberSequence() Then
                    '    With DManager
                    '        SQLString = "UPDATE NumberSequenceTable SET NextRec = " & CInt(Replace(TicketID, "AGST", "")) + 2 & " Where DataAreaID = '" & DataAreaID & "' and numbersequence = 'CC_Tickets'"
                    '        .ExecuteSQL(SQLString)
                    '    End With
                    'End If
                    If TakeABreak = True Or Len(sTicketID) = 0 Then Exit Function
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #61", EventLogEntryType.Information)
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
                        sEventEntry = "CIM Service: Function ProcessSupportMessage #62" & vbCrLf _
                            & "Ticket ID: " & sTicketID & vbCrLf _
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
                        sEventEntry = "CIM Service: Function ProcessSupportMessage #63" & vbCrLf _
                            & "Next Ticket Rec ID: " & Convert.ToString(NextTicketRecID)
                        WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                    End If
                    If TakeABreak = True Or NextTicketRecID <= 0 Then Exit Function
                Catch ex As Exception
                    TakeABreak = True
                    sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #64 returns " _
                        & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                        & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                        & "Exception Target: " & ex.TargetSite.ToString()
                    WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                    Exit Function
                End Try
                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #65", EventLogEntryType.Information)

                Try
                    If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #66", EventLogEntryType.Information)
                    With DManager
                        SQLString = "INSERT INTO C_Tickets (" _
                            & " CreatedBy, CreatedDateTime, ModifiedBy, ModifiedDateTime, TicketID_Display, Source, Status, TicketQueue, Tier, Partner3_," _
                            & " PartnerTicketSource, UpdatePartner, CallerPhone," _
                            & " CallerEmail, RecID, RecVersion, DataAreaID, UpdateContact, UpdateCaller, Caller, CallerName, severityLevel, CustomerName," _
                            & " PartnerTicketNo, SiteName, Product, SN, ContactName, ContactPhone, ContactEmail, ProblemDescription" _
                            & " )" _
                            & " VALUES (" _
                            & " 'MS AX', DATEADD(hh,4,getdate()), 'MS AX', DATEADD(hh,4,getdate()), '" & sTicketID & "', 'Email', 'Open', '" _
                            & TicketQueue & "', 'Tier 1', '" & Partner3_ & "', 'email', 1, '', '" _
                            & Replace(msg.From.EMail.ToString(), "'", "''") & "', '" & NextTicketRecID & "', 1, '" & DataAreaID & "', 1, 1, '" _
                            & Replace(msg.From.Name.ToString(), "'", "''") & "', '" _
                            & Replace(Mid(msg.From.Name.ToString(), 1, 60), "'", "''") & "', '" _
                            & SeverityLevel & "', '', '', '', '', '', '" _
                            & Replace(msg.From.Name.ToString(), "'", "''") & "', '', '" _
                            & Replace(msg.From.EMail.ToString(), "'", "''") & "', '" _
                            & Replace(msg.PlainMessage.Body.ToString(), "'", "''") & "')"

                        'eventViewerString &= vbCrLf & vbCrLf & "Inserted Ticket '" & TicketID & "' SQL: " & SQLString
                        'Write("CIM Service - Inserted Ticket ID: " & TicketID, "SQL: " & SQLString, 0)
                        If Debug_Mode Or Inform_Mode Then
                            sEventEntry = "CIM Service: Function ProcessSupportMessage #67" & vbCrLf _
                                & "Inserting Ticket ID: " & sTicketID & vbCrLf & vbCrLf & "SQL to execute: " & SQLString
                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                        End If
                        .ExecuteSQL(SQLString)
                    End With
                Catch ex As Exception
                    TakeABreak = True
                    sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #68 returns " _
                        & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                        & "Ticket ID: " & sTicketID & vbCrLf & vbCrLf _
                        & "SQL: " & SQLString & vbCrLf & vbCrLf _
                        & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                        & "Exception Target: " & ex.TargetSite.ToString()
                    WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                    Exit Function
                End Try

                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #69", EventLogEntryType.Information)
                IncrementTicketNumberSequence(TakeABreak)
                If TakeABreak = True Then Exit Function
                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #70", EventLogEntryType.Information)

                'C_TicketEvents
                REM Cast a long number to an integer 
                NextTicketEventRecID = GetNextTicketEventRecID(TakeABreak)
                If TakeABreak = True Or NextTicketEventRecID <= 0 Then Exit Function
                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #71", EventLogEntryType.Information)

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
                            sEventEntry = "CIM Service: Function ProcessSupportMessage #72" & vbCrLf _
                                & "Ticket ID: " & sTicketID & vbCrLf _
                                & "Next Ticket Events Rec ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf _
                                & vbCrLf & vbCrLf & "SQL to execute: " & SQLString
                            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                        End If
                        .ExecuteSQL(SQLString)
                        If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #73", EventLogEntryType.Information)
                    End With
                Catch ex As Exception
                    TakeABreak = True
                    sEventEntry = "CIM Service Exception: Function ProcessSupportMessage #74 returns " _
                        & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                        & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                        & "Ticket Events ID: " & Convert.ToString(NextTicketEventRecID) & vbCrLf & vbCrLf _
                        & "SQL: " & SQLString & vbCrLf & vbCrLf _
                        & "Exception Message: " & ex.Message.ToString() & vbCrLf _
                        & "Exception Target: " & ex.TargetSite.ToString()
                    WriteToEventLog(sEventEntry, EventLogEntryType.Error)
                    Exit Function
                End Try

                If Debug_Mode Then WriteToEventLog("CIM Service: Function ProcessSupportMessage #75", EventLogEntryType.Information)
                'Write("CIM Service - received ticket integration email", eventViewerString, 0)
                'WriteToEventLog("CIM Service - Received ticket integration email" & vbCrLf & "Event Viewer String: " & eventViewerString, EventLogEntryType.Information)
                ProcessSupportMessage = True    'returns
            End If
        Else
            ProcessSupportMessage = True    'returns
            If Debug_Mode Then
                sEventEntry = "CIM Service: Function ProcessSupportMessage #76, returns" _
                    & Convert.ToString(ProcessSupportMessage) & vbCrLf _
                    & "Email received that was not an integration request" & vbCrLf _
                    & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf _
                    & "Message ID: " & msg.MessageID.ToString() & vbCrLf _
                    & "From: " & msg.From.ToString() & vbCrLf _
                    & "To: " & msg.To.ToString() & vbCrLf _
                    & "Subject: " & msg.Subject.ToString() & vbCrLf _
                    & "Date: " & msg.Date.ToString() & vbCrLf _
                    & "Body: " & vbCrLf & vbTab & "-- beginning of line --" & vbCrLf _
                    & sTicketEmailBody & vbCrLf & vbTab _
                    & "-- end of line --"
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        End If
    End Function

    Sub ParseVerintOnyxEmail(ByVal TicketEmailBody As String, ByVal MsgID As String, ByVal IntegrationRequestType As String, ByRef TakeABreak As Boolean)
        Dim Priority As String = ""
        Dim CompanyName As String = "", PositionDash As Integer = 0
        Dim OnyxCustomerNumber As String = ""
        Dim OnyxIncidentNumber As String = ""
        Dim CustomerPrimaryContact As String = ""
        Dim Location As String = ""
        Dim PostalCode As String = ""
        Dim IncidentContact As String = ""
        Dim BusinessPhone As String = ""
        Dim CellPhone As String = ""
        Dim Fax As String = ""
        Dim Pager As String = ""
        Dim EmailAddress As String = ""
        Dim SupportLevel As String = ""
        Dim IncidentType As String = ""
        Dim Version As String = ""
        Dim SerialNumber As String = ""
        Dim SubmittedBy As String = ""
        Dim WorkNotes As String = ""
        Dim iStart As Integer = 0, iEnd As Integer = 0

        If Debug_Mode Then
            sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #01" & vbCrLf _
                & "Msg ID: " & MsgID & vbCrLf _
                & "Integration Request Type: " & IntegrationRequestType & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf _
                & "Ticket Email Body: " & vbCrLf & vbTab & "---- Beginnning of lines ----" & vbCrLf _
                & TicketEmailBody & vbCrLf & vbTab & "---- End of lines ----"
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        End If

        REM Parse message body...Verint's emails are formatted the same regardless of them being a new or an update request

        Try
            'Priority:  1 - Critical
            ExtractSingleLineData(TicketEmailBody, "Priority:", Priority, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #02" _
                    & vbCrLf & "Priority: " & Priority & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Company Name:  Hartford, The (Server #1) - Charlotte, NC
            ExtractSingleLineData(TicketEmailBody, "Company Name:", CompanyName, iStart, iEnd)

            PositionDash = InStr(CompanyName, "-")
            REM if a dash exists then after the dash there is the Site Name.  Before the dash there is the Company Name
            If PositionDash > 0 Then
                If Len(CompanyName) <= PositionDash Then
                    SiteName = ""
                Else
                    SiteName = Right(CompanyName, Len(CompanyName) - PositionDash)
                    SiteName = SiteName.Trim()
                End If
                If PositionDash <= 1 Then
                    CompanyName = ""
                Else
                    CompanyName = Left(CompanyName, PositionDash - 1)
                    CompanyName = CompanyName.Trim()
                End If
            End If
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #03" & vbCrLf _
                    & "Company Name: " & CompanyName & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd) & vbCrLf _
                    & "Site Name: " & SiteName
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Onyx Customer Number:  159048
            ExtractSingleLineData(TicketEmailBody, "Onyx Customer Number:", OnyxCustomerNumber, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #04" & vbCrLf _
                    & "Onyx Customer Number: " & OnyxCustomerNumber & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Onyx Incident Number:  3810615
            ExtractSingleLineData(TicketEmailBody, "Onyx Incident Number:", OnyxIncidentNumber, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #05" & vbCrLf _
                    & "Onyx Incident Number: " & OnyxIncidentNumber & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Customer Primary Contact:  Frank Beatty
            ExtractSingleLineData(TicketEmailBody, "Customer Primary Contact:", CustomerPrimaryContact, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #06" & vbCrLf _
                    & "Customer Primary Contact: " & CustomerPrimaryContact & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Location:  Charlotte Service Ctr, 8711 University East Dr, Charlotte, North Carolina, United States
            ExtractSingleLineData(TicketEmailBody, "Location:", Location, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #07" & vbCrLf _
                    & "Location: " & Location & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Postal Code:  28213
            ExtractSingleLineData(TicketEmailBody, "Postal Code:", PostalCode, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #08" & vbCrLf _
                    & "Postal Code: " & PostalCode & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Incident Contact:  Steven Hudak
            ExtractSingleLineData(TicketEmailBody, "Incident Contact:", IncidentContact, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #09" & vbCrLf _
                    & "Incident Contact: " & IncidentContact & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Business Phone:  8609197122
            ExtractSingleLineData(TicketEmailBody, "Business Phone:", BusinessPhone, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #10" & vbCrLf _
                    & "Business Phone: " & BusinessPhone & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Cell Phone:  8609197122
            ExtractSingleLineData(TicketEmailBody, "Cell Phone:", CellPhone, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #11" & vbCrLf _
                    & "Cell Phone: " & CellPhone & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Fax: 8609197122
            ExtractSingleLineData(TicketEmailBody, "Fax:", Fax, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #12" & vbCrLf _
                    & "Fax: " & Fax & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Pager: 8609197122
            ExtractSingleLineData(TicketEmailBody, "Pager:", Pager, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #13" & vbCrLf _
                    & "Pager: " & Pager & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Email Address:  steven.hudak@thehartford.com
            ExtractSingleLineData(TicketEmailBody, "Email Address:", EmailAddress, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #14" & vbCrLf _
                    & "EmailAddress: " & EmailAddress & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Support Level:  Advantage
            ExtractSingleLineData(TicketEmailBody, "Support Level:", SupportLevel, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #15" & vbCrLf _
                    & "SupportLevel: " & SupportLevel & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Incident Type:  Web - Unknown
            ExtractSingleLineData(TicketEmailBody, "Incident Type:", IncidentType, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #16" & vbCrLf _
                    & "IncidentType: " & IncidentType & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Product:  eQContactStore
            ExtractSingleLineData(TicketEmailBody, "Product:", Product, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #17" & vbCrLf _
                    & "Product: " & Product & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Version:  CS - 7.2
            ExtractSingleLineData(TicketEmailBody, "Version:", Version, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #18" & vbCrLf _
                    & "Version: " & Version & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Serial Number: RN-XXXXXX
            ExtractSingleLineData(TicketEmailBody, "Serial Number:", SerialNumber, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #19" & vbCrLf _
                    & "Serial Number: " & SerialNumber & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Submitted By:  Janice Wells
            ExtractSingleLineData(TicketEmailBody, "Submitted By:", SubmittedBy, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #20" & vbCrLf _
                    & "SubmittedBy: " & SubmittedBy & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'WorkNotes: 
            REM do not put a colon on the 2nd parameter of ExtractWorkNotes( )
            ExtractWorkNotes(TicketEmailBody, "WorkNotes", WorkNotes, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #21" & vbCrLf & "WorkNotes: " _
                    & vbCrLf & vbTab & "---- Beginning of lines ----" & vbCrLf _
                    & WorkNotes & vbCrLf & vbTab & "---- End of lines ----" _
                    & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
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
                sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #22" & vbCrLf _
                    & "SeverityLevel: " & SeverityLevel & vbCrLf _
                    & "SeverityLevelDescription: " & SeverityLevelDescription
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

            If Debug_Mode Or Inform_Mode Or Summary_Inform_Mode Then
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
                If Debug_Mode Then
                    sEventEntry = "CIM Service: Sub ParseVerintOnyxEmail #27" & vbCrLf & "IF..THEN block begins"
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
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
        Dim SRSeverity As String = ""
        Dim CustomerNumber As String = ""
        Dim IncidentNumber As String = ""
        Dim Location As String = ""
        Dim PostalCode As String = ""
        Dim IncidentContact As String = "", IncidentContactFirstName As String = "", IncidentContactLastName As String = ""
        Dim IncidentPrimaryContactPhone As String = ""
        Dim IncidentPrimaryContactEmail As String = ""
        Dim SupportLevel As String = ""
        Dim SRType As String = ""
        Dim ServiceProductVertical As String = ""
        Dim SerialNumber As String = ""
        Dim TaskStatus As String = ""
        Dim LoggedBy As String = ""
        Dim WorkNotes As String = ""
        Dim iStart As Integer = 0, iEnd As Integer = 0

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
            'SR Severity         :    
            ExtractSingleLineData(TicketEmailBody, "SR Severity", SRSeverity, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #02" & vbCrLf _
                    & "SRSeverity: " & SRSeverity & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If


            'Customer Name       : VERINT GMBH          
            ExtractSingleLineData(TicketEmailBody, "Customer Name", CustomerName, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #03" & vbCrLf _
                    & "CustomerName: " & CustomerName & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Site Name           : VERINT GMBH KARLSRUHE,  D-76139 DE_Site
            ExtractSingleLineData(TicketEmailBody, "Site Name", SiteName, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #04" & vbCrLf _
                    & "SiteName: " & SiteName & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Customer Number     : 295464
            ExtractSingleLineData(TicketEmailBody, "Customer Number", CustomerNumber, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #05" & vbCrLf _
                    & "CustomerNumber: " & CustomerNumber & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Incident Number     : 1445980
            ExtractSingleLineData(TicketEmailBody, "Incident Number", IncidentNumber, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #06" & vbCrLf _
                    & "IncidentNumber: " & IncidentNumber & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Location            : AM STORRENACKER 2, , KARLSRUHE, D-76139  DE
            ExtractSingleLineData(TicketEmailBody, "Location", Location, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #07" & vbCrLf _
                    & "Location: " & Location & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Postal Code         : D-76139
            ExtractSingleLineData(TicketEmailBody, "Postal Code", PostalCode, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #08" & vbCrLf _
                    & "PostalCode: " & PostalCode & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Incident Primary Contact First name    : 
            ExtractSingleLineData(TicketEmailBody, "Incident Primary Contact First name", IncidentContactFirstName, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #09" & vbCrLf _
                    & "IncidentContact First Name: " & IncidentContactFirstName & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Incident Primary Contact Last name     : 
            ExtractSingleLineData(TicketEmailBody, "Incident Primary Contact Last name", IncidentContactLastName, iStart, iEnd)
            If Len(IncidentContactFirstName) > 0 Then
                If Len(IncidentContactLastName) > 0 Then
                    IncidentContact = IncidentContactFirstName & " " & IncidentContactLastName
                Else
                    IncidentContact = IncidentContactFirstName
                End If
            Else
                If Len(IncidentContactLastName) > 0 Then
                    IncidentContact = IncidentContactLastName
                Else
                    IncidentContact = ""
                End If
            End If
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #10" & vbCrLf _
                    & "IncidentContact Last Name: " & IncidentContactLastName & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd) & vbCrLf _
                    & "IncidentContact First + Last Name: " & IncidentContact
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Incident Primary Contact Email         : 
            ExtractSingleLineData(TicketEmailBody, "Incident Primary Contact Email", IncidentPrimaryContactEmail, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #11" & vbCrLf _
                    & "IncidentPrimaryContactEmail: " & IncidentPrimaryContactEmail & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Incident Primary Contact Phone         : 
            ExtractSingleLineData(TicketEmailBody, "Incident Primary Contact Phone", IncidentPrimaryContactPhone, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #12" & vbCrLf _
                    & "IncidentPrimaryContactPhone: " & IncidentPrimaryContactPhone & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Support Level       : 
            ExtractSingleLineData(TicketEmailBody, "Support Level", SupportLevel, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #13" & vbCrLf _
                    & "SupportLevel: " & SupportLevel & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'SR Type             : VRNT System Malfunction
            ExtractSingleLineData(TicketEmailBody, "SR Type", SRType, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #14" & vbCrLf _
                    & "SRType: " & SRType & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Service Product Vertical : Next Generation (NG) WAS ServiceProductVertical Version : VI360
            ExtractSingleLineData(TicketEmailBody, "Service Product Vertical", ServiceProductVertical, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #15" & vbCrLf _
                    & "ServiceProductVertical: " & ServiceProductVertical & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Serial Number       : 
            ExtractSingleLineData(TicketEmailBody, "Serial Number", SerialNumber, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #16" & vbCrLf _
                    & "SerialNumber: " & SerialNumber & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Task Status         : VRNT Closed
            ExtractSingleLineData(TicketEmailBody, "Task Status", TaskStatus, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #17" & vbCrLf _
                    & "TaskStatus: " & TaskStatus & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Logged By           : NSINGH
            ExtractSingleLineData(TicketEmailBody, "Logged By", LoggedBy, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #18" & vbCrLf & "LoggedBy: " & LoggedBy _
                    & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If

            'Work Notes          : 
            REM do not put a colon on the 2nd parameter of ExtractWorkNotes( )
            ExtractWorkNotes(TicketEmailBody, "Work Notes", WorkNotes, iStart, iEnd)
            If Debug_Mode Then
                sEventEntry = "CIM Service: Sub ParseVerintOracleEmail #19" & vbCrLf & "WorkNotes: " & vbCrLf _
                    & vbTab & "---- Beginning of lines ----" & vbCrLf & WorkNotes & vbCrLf _
                    & vbTab & "---- End of lines ----" _
                    & vbCrLf & vbTab & CStr(iStart) & ", " & CStr(iEnd)
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
            If Debug_Mode Or Inform_Mode Or Summary_Inform_Mode Then
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
                    & "Problem Description: " & vbCrLf & vbTab & "---- Beginning of lines ----" _
                    & vbCrLf & ProblemDescription & vbCrLf & vbTab & "---- End of lines ----"
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
                            OldStatus = dt.Rows(0).Item("Status")
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
                sEventEntry = "CIM Service Exception: Sub ParseVerintOracleEmail #24: " & vbCrLf _
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
                If (OldStatus <> TaskStatus) Then
                    SQLTicketChangesEventString = vbCrLf & vbCrLf & "Status" & vbCrLf & "OLD: " & OldStatus & vbCrLf & "NEW: " & TaskStatus
                    SQLTicketChangesString = ", Status = '" & TaskStatus & "'"
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
        If Debug_Mode Then WriteToEventLog("CIM Service: Sub SendUpdateTicketNotFoundEmail #01", EventLogEntryType.Information)

        Try
            'will need another IF THEN block for 'UK'                               Adam Ip 
            If DataAreaID = "US" Then
                mailMsg.To.Add("Brian Brazil", "bbrazil@adtech.net")
                mailMsg.Subject = "DROPPED " & Partner & " Update Email (incident #: " & PartnerTicketNo & ")"
                mailMsg.Body = Partner & " incident # " & Chr(96) & PartnerTicketNo & Chr(39) & " was not found in the AGS support ticket system. " _
                        & vbCrLf & vbCrLf & "AGS Automated System Message" & vbCrLf & vbCrLf _
                        & "---- Original message ----" & vbCrLf & vbCrLf & msg.PlainMessage.Body.ToString()
                mailMsg.Headers.Add("Reply-To", "noreply@adtech.net")    ' Adam Ip 2011-03-17 
                client.SendMessage(mailMsg)
                If Debug_Mode Then
                    sEventEntry = "CIM Service: Sub SendUpdateTicketNotFoundEmail #02" & vbCrLf _
                        & "Message ID: " & msg.MessageID & vbCrLf _
                        & "From: " & mailMsg.From.ToString() & vbCrLf _
                        & "To: " & mailMsg.To.ToString.ToString() & vbCrLf _
                        & "Subject: " & mailMsg.Subject.ToString() & vbCrLf _
                        & "Date: " & mailMsg.Date.ToString() & vbCrLf _
                        & "Body: " & mailMsg.Body.ToString()
                    WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                End If
            End If
        Catch ex As Exception
            sEventEntry = "CIM Service Exception: Sub SendUpdateTicketNotFoundEmail #03" & vbCrLf _
                & "Message ID: " & msg.MessageID & vbCrLf _
                & "From: " & mailMsg.From.ToString() & vbCrLf _
                & "To: " & mailMsg.To.ToString() & vbCrLf _
                & "Subject: " & mailMsg.Subject.ToString() & vbCrLf _
                & "Date: " & mailMsg.Date.ToString() & vbCrLf _
                & "Message Body: " & mailMsg.Body.ToString() & vbCrLf _
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
                    & "From: " & mailMsg.From.ToString() & vbCrLf _
                    & "To: " & mailMsg.To.ToString() & vbCrLf _
                    & "Subject: " & mailMsg.Subject.ToString() & vbCrLf _
                    & "Date: " & mailMsg.Date.ToString() & vbCrLf _
                    & "Message Body: " & mailMsg.Body
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            sEventEntry = "CIM Service Exception: Sub SendReplyEmail" & vbCrLf _
                & "Message ID: " & msg.MessageID & vbCrLf _
                & "From: " & mailMsg.From.ToString() & vbCrLf _
                & "To: " & mailMsg.To.ToString() & vbCrLf _
                & "Subject: " & mailMsg.Subject.ToString() & vbCrLf _
                & "Date: " & mailMsg.Date.ToString() & vbCrLf _
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
                    & "From: " & mailMsg.From.ToString() & vbCrLf _
                    & "To: " & mailMsg.To.ToString() & vbCrLf _
                    & "Subject: " & mailMsg.Subject.ToString() & vbCrLf _
                    & "Date: " & mailMsg.Date.ToString() & vbCrLf _
                    & "Message Body: " & mailMsg.Body
                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
            End If
        Catch ex As Exception
            sEventEntry = "CIM Service Exception: Sub SendNotificationEmail" & vbCrLf _
                & "Message ID: " & msg.MessageID & vbCrLf _
                & "From: " & mailMsg.From.ToString() & vbCrLf _
                & "To: " & mailMsg.To.ToString() & vbCrLf _
                & "Subject: " & mailMsg.Subject.ToString() & vbCrLf _
                & "Date: " & mailMsg.Date.ToString() & vbCrLf _
                & "Message Body: " & mailMsg.Body & vbCrLf _
                & "Exception Target: " & ex.TargetSite.ToString() & vbCrLf _
                & "Exception Message: " & ex.Message
            WriteToEventLog(sEventEntry, EventLogEntryType.Error)
        End Try
    End Sub

    Public Function FindMessageID(ByRef MessageID As String, ByRef TakeABreak As Boolean) As Boolean
        Dim sql As String = ""
        Dim strMsgID As String = ""

        Dim ds As DataSet = New DataSet
        Dim dt As DataTable = New DataTable

        FindMessageID = False
        Try
            REM Trim MessageID here, will Return By Reference
            MessageID = Trim(MessageID)
            If Len(MessageID) > 0 And TakeABreak = False Then
                sql = "SELECT * FROM MessageIDs WHERE MessageID LIKE '" & MessageID & "';"
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

    Public Function UpdateMessageIDinProgress(ByVal MessageID As String, ByVal MessageFrom As String, ByVal sAction As Char, ByRef TakeABreak As Boolean) As Long
        Dim sql(2) As String
        Dim ds As DataSet = New DataSet

        UpdateMessageIDinProgress = CLng(0)
        sql(0) = ""
        sql(1) = ""

        Try
            If TakeABreak = False Then
                MessageID = Trim(MessageID)
                With DManager
                    If Debug_Mode Then
                        sEventEntry = "CIM Service: Function UpdateMessageIDinProgress #01 " & vbCrLf _
                            & "Parameter MessageID passed ByVal: " & MessageID & vbCrLf _
                            & "Parameter sAction passed ByVal: " & Convert.ToString(sAction) & vbCrLf _
                            & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) & vbCrLf & vbCrLf
                    End If
                    Select Case sAction
                        Case "A"
                            sql(0) = "SELECT COUNT(MessageID) AS MsgCount FROM MessageIDinProgress WHERE MessageID LIKE '" & MessageID & "';"
                            ds = .GetDataSet(sql(0))
                            If Not ds Is Nothing Then
                                If ds.Tables(0).Rows.Count > 0 Then
                                    UpdateMessageIDinProgress = CLng(ds.Tables(0).Rows(0).Item("MsgCount"))
                                End If
                            End If
                            sql(1) = "INSERT INTO MessageIDinProgress " _
                                        & "(MessageID, MessageFrom, CreationDateTime) " _
                                        & "VALUES " _
                                        & "(" _
                                        & "'" & MessageID.Trim() & "', " _
                                        & "'" & MessageFrom.Trim() & "', " _
                                        & "GetDate() " _
                                        & ");"
                            If Debug_Mode Then
                                sEventEntry = sEventEntry & "UpdateMessageIDinProgress returns: " _
                                    & Convert.ToString(UpdateMessageIDinProgress) & vbCrLf & vbCrLf _
                                    & "SQL(0) executed: " & sql(0) & vbCrLf & vbCrLf _
                                    & "SQL(1) to execute: " & sql(1)
                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            End If
                            .ExecuteSQL(sql(1))
                        Case "D"
                            sql(0) = "DELETE FROM MessageIDinProgress WHERE MessageID LIKE '" & MessageID.Trim() & "';"
                            If Debug_Mode Then
                                sEventEntry = sEventEntry & "UpdateMessageIDinProgress returns: " _
                                    & Convert.ToString(UpdateMessageIDinProgress) & vbCrLf & vbCrLf & "SQL(0) to execute: " & sql(0)
                                WriteToEventLog(sEventEntry, EventLogEntryType.Information)
                            End If
                            .ExecuteSQL(sql(0))
                    End Select
                End With
            End If
        Catch ex As Exception
            TakeABreak = True
            sEventEntry = "CIM Service Exception: Function UpdateMessageIDinProgress #02" & vbCrLf _
                & "Parameter MessageID passed ByVal: " & MessageID & vbCrLf _
                & "Parameter sAction passed ByVal: " & Convert.ToString(sAction) & vbCrLf _
                & "Parameter TakeABreak passed ByRef: " & Convert.ToString(TakeABreak) _
                & "UpdateMessageIDinProgress returns: " & Convert.ToString(UpdateMessageIDinProgress) & vbCrLf & vbCrLf _
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

    Private Function ExtractSingleLineData(ByVal sBody As String, ByVal sFind As String, ByRef sFound As String, ByRef iStart As Integer, ByRef iEnd As Integer) As Boolean
        Dim i As Integer, j As Integer, k As Integer, l As Integer, colon As Integer
        Static Bookmark As Integer
        REM Dim sNext As String

        REM InStr( ) 1st parameter
        REM     Specifies the starting position for each search. 
        REM     Search begins at the first character position (1) by default
        REM InStr( ) 4th parameter
        REM     0 = vbBinaryCompare - Perform a binary comparison
        REM     1 = vbTextCompare - Perform a textual comparison

        REM Initialize Bookmark
        If iStart = 0 And iEnd = 0 Then Bookmark = 1

        If Debug_Mode Then
            sEventEntry = "CIM Service: Function ExtractSingleLineData #01 " & vbCrLf _
                & "sBody, starting from Bookmark: " & vbCrLf & vbTab & "---- Beginning of lines ----" & vbCrLf _
                & Right(sBody, Len(sBody) - Bookmark + 1) & vbCrLf & vbTab & "---- End of lines ----" & vbCrLf & vbCrLf _
                & "sFind (input): " & sFind & vbCrLf _
                & "Bookmark, iStart, iEnd: " & Convert.ToString(Bookmark) & ", " & Convert.ToString(iStart) & ", " & Convert.ToString(iEnd) _
                & vbCrLf & vbCrLf & vbCrLf
        End If

        iStart = InStr(Bookmark, sBody, sFind, 1)
        If iStart > 0 Then
            colon = InStr(1, sFind, ":")
            REM if sFind is not embedded with a colon
            If colon = 0 Then
                iStart = InStr(iStart, sBody, ":", 1) + 1
            Else
                iStart = iStart + colon
            End If
            If iStart > 0 Then
                While Mid(sBody, iStart, 1) = " " Or Mid(sBody, iStart, 1) = vbTab
                    iStart = iStart + 1
                End While
                i = InStr(iStart, sBody, vbCr, 0)
                j = InStr(iStart, sBody, vbLf, 0)
                k = InStr(iStart, sBody, vbCrLf, 0)
                l = InStr(iStart, sBody, vbNewLine, 0)
                iEnd = FindNonZeroMin(i, j, k, l)
                If iEnd > 0 Then
                    REM iEnd == iStart means Cr, Lf, Crlf, or NewLine immediately after colon 
                    If iEnd > iStart Then
                        REM RTrim white space, i.e. RTrim space bar, Tab, Cr, Lf, Crlf, and NewLine 
                        While Char.IsWhiteSpace(Mid(sBody, iEnd, 1))
                            iEnd = iEnd - 1
                        End While
                    End If
                Else
                    REM iEnd == 0 means reaching the End Of File
                    iEnd = Len(sBody)
                End If
                If iEnd - iStart + 1 > 0 Then
                    REM iStart: Space bar and Tab characters were just skipped
                    REM     so here is to determine whether iStart is pointing to a Cr, Lf, Crlf, or Newline
                    If Char.IsWhiteSpace(Mid(sBody, iStart, 1)) Then
                        sFound = ""
                    Else
                        sFound = Mid(sBody, iStart, iEnd - iStart + 1)
                        iEnd = iEnd + 1
                    End If
                Else
                    REM this statement is reduntant
                    sFound = ""
                End If
            End If
        End If
        If iStart > 0 And iEnd > 0 And (iEnd - iStart > 0 Or (iStart = iEnd And sFound = "")) Then
            Bookmark = iEnd + 1
            REM sNext = Right(sBody, Len(sBody) - iEnd)
            ExtractSingleLineData = True
        Else
            ExtractSingleLineData = False
        End If
        If Debug_Mode Then
            sEventEntry = sEventEntry & "CIM Service: Function ExtractSingleLineData #02 returns " _
                & Convert.ToString(ExtractSingleLineData) & vbCrLf _
                & "sBody, starting from Bookmark: " & vbCrLf & vbTab & "---- Beginning of lines ----" & vbCrLf _
                & Right(sBody, Len(sBody) - Bookmark + 1) & vbCrLf & vbTab & "---- End of lines ----" & vbCrLf & vbCrLf _
                & "sFind: " & sFind & vbCrLf _
                & "sFound: " & sFound & vbCrLf _
                & "Bookmark, iStart, iEnd: " & Convert.ToString(Bookmark) & ", " & Convert.ToString(iStart) & ", " & Convert.ToString(iEnd)
            WriteToEventLog(sEventEntry, EventLogEntryType.Information)
        End If
    End Function

    Private Function ExtractWorkNotes(ByVal sBody As String, ByVal sFind As String, ByRef sFound As String, ByRef iStart As Integer, ByRef iEnd As Integer) As Boolean
        Dim colon As Integer

        sFound = ""
        If iEnd < 1 Then iEnd = 1
        iStart = InStr(iEnd, sBody, sFind)
        iEnd = Len(sBody)
        If iStart > 0 Then
            colon = InStr(1, sFind, ":")
            REM if sFind is not embedded with a colon
            If colon = 0 Then
                iStart = InStr(iStart, sBody, ":", 1)
                REM if cannot locate a colon, then returns sFound as empty string
                REM If iStart = 0 Then Exit Function
            Else
                iStart = iStart + colon
            End If
            If iStart > 0 Then
                iStart = iStart + 1
                If iEnd >= iStart Then
                    sFound = Right(sBody, iEnd - iStart)
                    If Not sFound = Nothing Then sFound = sFound.Trim()
                Else
                    sFound = ""
                End If
            End If
        End If
    End Function

    Private Function FindNonZeroMin(ByVal iFirst As Integer, ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As Integer
        Dim iter As Integer, Arr(3) As Integer
        If iFirst > 0 Then
            FindNonZeroMin = iFirst
        Else
            REM -1 means this function fails to find a non-zero minimum because all parameters are zero or negative
            FindNonZeroMin = -1
        End If
        Arr(0) = x
        Arr(1) = y
        Arr(2) = z
        For iter = 0 To 2
            If (FindNonZeroMin > 0 And Arr(iter) > 0 And FindNonZeroMin > Arr(iter)) Or (FindNonZeroMin < 0 And Arr(iter) > 0) Then
                FindNonZeroMin = Arr(iter)
            End If
        Next iter
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
