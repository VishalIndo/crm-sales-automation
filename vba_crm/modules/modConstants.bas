Attribute VB_Name = "modConstants"
Option Explicit

Public Const APP_TITLE As String = "Excel VBA CRM"
Public Const APP_COMPANY As String = "Northwind Sales"

Public Const SHEET_HOME As String = "Home"
Public Const SHEET_LEADS As String = "Leads"
Public Const SHEET_ACTIVITIES As String = "Activities"
Public Const SHEET_QUOTES As String = "Quotations"
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_LISTS As String = "Lists"
Public Const SHEET_SETTINGS As String = "Settings"

Public Const TABLE_LEADS As String = "tblLeads"
Public Const TABLE_ACTIVITIES As String = "tblActivities"
Public Const TABLE_QUOTES As String = "tblQuotes"

Public Const BUTTON_PREFIX As String = "crmBtn_"
Public Const DECOR_PREFIX As String = "crmDecor_"

Public Function LeadHeaders() As Variant
    LeadHeaders = Array( _
        "LeadID", "CreatedDate", "CompanyName", "ContactName", "Phone", "Email", _
        "LeadSource", "Industry", "AssignedTo", "Stage", "ExpectedValue", _
        "Probability", "NextFollowUp", "Status", "LastContactDate", "Notes" _
    )
End Function

Public Function ActivityHeaders() As Variant
    ActivityHeaders = Array( _
        "ActivityID", "LeadID", "ActivityDate", "ActivityType", "Summary", _
        "NextAction", "NextFollowUp", "Owner", "Outcome" _
    )
End Function

Public Function QuoteHeaders() As Variant
    QuoteHeaders = Array( _
        "QuoteID", "LeadID", "QuoteDate", "Amount", "QuoteStatus", "ValidUntil", "Notes" _
    )
End Function

Public Function StageList() As Variant
    StageList = Array("New", "Contacted", "Qualified", "Quoted", "Negotiation", "Won", "Lost")
End Function

Public Function SourceList() As Variant
    SourceList = Array("Referral", "Website", "Email", "Call", "Walk-in", "Campaign", "Partner", "Other")
End Function

Public Function ActivityTypeList() As Variant
    ActivityTypeList = Array("Call", "Email", "Meeting", "Demo", "Quotation", "WhatsApp", "Other")
End Function

Public Function QuoteStatusList() As Variant
    QuoteStatusList = Array("Draft", "Sent", "Accepted", "Rejected", "Expired")
End Function

Public Function StatusList() As Variant
    StatusList = Array("Open", "Closed")
End Function
