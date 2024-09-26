VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolboxCourseConditionQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Properties
Private m_Detail() As String
Private m_Plane As String
Private m_Shinkansen As String
Private m_ShinkansenNozomi As String
Private m_SleeperTrain As String
Private m_LimitedExpress As String
Private m_LocalBus As String
Private m_MidnightBus As String
Private m_Liner As String
Private m_HighwayBus As String
Private m_ConnectionBus As String
Private m_CommunityBus As String

Private m_Ship As String
Private m_Walk As String
Private m_UseJr As String
Private m_Transfer As String
Private m_WaitAverageTime As Boolean
Private m_ExpressStartingStation As String
Private m_LocalBusOnly As Boolean
Private m_TransferTime As String
Private m_FuzzyLine As Boolean
Private m_EntryPathBehavior As Boolean
Private m_SurchargeKind As String
Private m_TeikiKind As String
Private m_JrSeasonalRate As Boolean
Private m_JrReservation As String
Private m_ShinkansenETicket As String
Private m_StudentDiscount As Boolean
Private m_TicketSystemType As String
Private m_PreferredTicketOrder As String
Private m_NikukanTeiki As Boolean
Private m_OffPeakTeiki As Boolean

Private m_Client As Ekispert

Property Let Client(ByVal Client As Ekispert)
    Set m_Client = Client
End Property

' Property Let methods
Property Let Detail(i As Long, ByVal value As String)
    ReDim Preserve m_Detail(i) As String
    m_Detail(i) = value
End Property

Public Property Let Plane(ByVal value As String)
    m_Plane = value
End Property

Public Property Let Shinkansen(ByVal value As String)
    m_Shinkansen = value
End Property

Public Property Let ShinkansenNozomi(ByVal value As String)
    m_ShinkansenNozomi = value
End Property

Public Property Let SleeperTrain(ByVal value As String)
    m_SleeperTrain = value
End Property

Public Property Let LimitedExpress(ByVal value As String)
    m_LimitedExpress = value
End Property

Public Property Let HighwayBus(ByVal value As String)
    m_HighwayBus = value
End Property

Public Property Let ConnectionBus(ByVal value As String)
    m_ConnectionBus = value
End Property

Public Property Let LocalBus(ByVal value As String)
    m_LocalBus = value
End Property

Public Property Let CommunityBus(ByVal value As CommunityBusEnum)
    m_CommunityBus = m_Client.GetCommunityBus(value)
End Property

Public Property Let MidnightBus(ByVal value As String)
    m_MidnightBus = value
End Property

Public Property Let Ship(ByVal value As String)
    m_Ship = value
End Property

Public Property Let Liner(ByVal value As String)
    m_Liner = value
End Property

Public Property Let WaitAverageTime(ByVal value As Boolean)
    m_WaitAverageTime = value
End Property

Public Property Let FuzzyLine(ByVal value As Boolean)
    m_FuzzyLine = value
End Property

Public Property Let EntryPathBehavior(ByVal value As Boolean)
    m_EntryPathBehavior = value
End Property

Public Property Let JrSeasonalRate(ByVal value As Boolean)
    m_JrSeasonalRate = value
End Property

Public Property Let StudentDiscount(ByVal value As Boolean)
    m_StudentDiscount = value
End Property

Public Property Let NikukanTeiki(ByVal value As Boolean)
    m_NikukanTeiki = value
End Property

Public Property Let OffPeakTeiki(ByVal value As Boolean)
    m_OffPeakTeiki = value
End Property

Public Property Let Walk(ByVal value As String)
    m_Walk = value
End Property

Public Property Let Transfer(ByVal value As String)
    m_Transfer = value
End Property

Public Property Let UseJr(ByVal value As String)
    m_UseJr = value
End Property

Public Property Let ExpressStartingStation(ByVal value As String)
    m_ExpressStartingStation = value
End Property

Public Property Let LocalBusOnly(ByVal value As Boolean)
    m_LocalBusOnly = value
End Property

Public Property Let TransferTime(ByVal value As String)
    m_TransferTime = value
End Property

Public Property Let SurchargeKind(ByVal value As String)
    m_SurchargeKind = value
End Property

Public Property Let TeikiKind(ByVal value As String)
    m_TeikiKind = value
End Property

Public Property Let JrReservation(ByVal value As String)
    m_JrReservation = value
End Property

Public Property Let ShinkansenETicket(ByVal value As String)
    m_ShinkansenETicket = value
End Property

Public Property Let TicketSystemType(ByVal value As String)
    m_TicketSystemType = value
End Property

Public Property Let PreferredTicketOrder(ByVal value As String)
    m_PreferredTicketOrder = value
End Property

' Initialize default values
Private Sub Class_Initialize()
    m_Plane = "normal"
    m_Shinkansen = "normal"
    m_ShinkansenNozomi = "normal"
    m_SleeperTrain = "never"
    m_LimitedExpress = "normal"
    m_HighwayBus = "normal"
    m_ConnectionBus = "normal"
    m_LocalBus = "normal"
    m_CommunityBus = "contain"
    m_MidnightBus = "never"
    m_Ship = "normal"
    m_Liner = "normal"
    m_Walk = "normal"
    m_UseJr = "normal"
    m_Transfer = "normal"
    m_WaitAverageTime = True
    m_ExpressStartingStation = "normal"
    m_LocalBusOnly = False
    m_TransferTime = "normal"
    m_FuzzyLine = False
    m_EntryPathBehavior = False
    m_SurchargeKind = "free"
    m_TeikiKind = "bussiness" ' TYPO
    m_JrSeasonalRate = True
    m_JrReservation = "none"
    m_ShinkansenETicket = "none"
    m_StudentDiscount = False
    m_TicketSystemType = "normal"
    m_PreferredTicketOrder = "none"
    m_NikukanTeiki = False
    m_OffPeakTeiki = False
End Sub

Public Function Find() As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/toolbox/course/condition"
    Dim Result As Dictionary
    Set Result = m_Client.Fetch(ApiPath, BuildQuery())
    
    Dim Response As ResultSet
    
    If m_Client.Success = False Then
        Response.Success = False
        Response.Error = m_Client.Error
        Find = Response
        Exit Function
    End If
    Response.Condition = Result("Condition")
    Response.Success = True
    Find = Response

    ' Debug.Print JsonConverter.ConvertToJson(Result)
End Function


Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    m_Client.ArrayToString BuildQuery, "detail", m_Detail
    BuildQuery.Add "waitAverageTime", IIf(m_WaitAverageTime, "true", "false")
    BuildQuery.Add "entryPathBehavior", IIf(m_EntryPathBehavior, "true", "false")
    BuildQuery.Add "localBusOnly", IIf(m_LocalBusOnly, "true", "false")
    BuildQuery.Add "JRSeasonalRate", IIf(m_JrSeasonalRate, "true", "false")
    BuildQuery.Add "studentDiscount", IIf(m_StudentDiscount, "true", "false")
    BuildQuery.Add "fuzzyLine", IIf(m_FuzzyLine, "true", "false")
    BuildQuery.Add "nikukanteiki", IIf(m_NikukanTeiki, "true", "false")
    BuildQuery.Add "offpeakTeiki", IIf(m_OffPeakTeiki, "true", "false")
    
    m_Client.StringParam BuildQuery, "plane", m_Plane
    m_Client.StringParam BuildQuery, "shinkansen", m_Shinkansen
    m_Client.StringParam BuildQuery, "shinkansenNozomi", m_ShinkansenNozomi
    m_Client.StringParam BuildQuery, "sleeperTrain", m_SleeperTrain
    m_Client.StringParam BuildQuery, "limitedExpress", m_LimitedExpress
    m_Client.StringParam BuildQuery, "highwayBus", m_HighwayBus
    m_Client.StringParam BuildQuery, "connectionBus", m_ConnectionBus
    m_Client.StringParam BuildQuery, "localBus", m_LocalBus
    m_Client.StringParam BuildQuery, "communityBus", m_CommunityBus
    m_Client.StringParam BuildQuery, "midnightBus", m_MidnightBus
    m_Client.StringParam BuildQuery, "ship", m_Ship
    m_Client.StringParam BuildQuery, "liner", m_Liner
    m_Client.StringParam BuildQuery, "walk", m_Walk
    m_Client.StringParam BuildQuery, "useJR", m_UseJr
    m_Client.StringParam BuildQuery, "transfer", m_Transfer
    m_Client.StringParam BuildQuery, "expressStartingStation", m_ExpressStartingStation
    m_Client.StringParam BuildQuery, "transferTime", m_TransferTime
    m_Client.StringParam BuildQuery, "surchargeKind", m_SurchargeKind
    m_Client.StringParam BuildQuery, "teikiKind", m_TeikiKind
    m_Client.StringParam BuildQuery, "JRReservation", m_JrReservation
    m_Client.StringParam BuildQuery, "shinkansenETicket", m_ShinkansenETicket
    m_Client.StringParam BuildQuery, "ticketSystemType", m_TicketSystemType
    m_Client.StringParam BuildQuery, "preferredTicketOrder", m_PreferredTicketOrder

End Function




