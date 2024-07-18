VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StationLightQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Name As String
Private m_NameMatchType As NameMatchTypeEnum
Private m_Code As Long
Private m_TransportType() As String
Private m_PrefectureCode() As String
Private m_CorporationBind As String
Private m_CommunityBus As CommunityBusEnum

Private m_Client As Ekispert

Private Sub Class_Initialize()
    m_NameMatchType = Forward
    m_CommunityBus = Contain
End Sub

Public Function Find() As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/station/light"
    Dim Query As Dictionary
    Set Query = BuildQuery()
    Dim Result As Dictionary
    Set Result = m_Client.Fetch(ApiPath, Query)
    Dim dicPoints() As Dictionary
    Dim Response As ResultSet
    
    If m_Client.Success = False Then
        Response.Success = False
        Response.Error = m_Client.Error
        Find = Response
        Exit Function
    End If
    m_Client.SetArray dicPoints, Result, "Point"
    Response.Max = Result("max")
    Response.Offset = Result("offset")
    Response.Points = CreatePoints(dicPoints)
    Response.Success = True
    Find = Response
End Function

Function CreatePoints(dicPoints() As Dictionary) As Point()
    Dim i As Long
    Dim Points() As Point
    ReDim Points(UBound(dicPoints) - 1) As Point
    
    For i = 0 To UBound(dicPoints) - 1
        Dim dicPoint As Dictionary
        Set dicPoint = dicPoints(i)
        ' Debug.Print JsonConverter.ConvertToJson(dicPoint)
        Dim tPoint  As Point
        ' Station
        Dim dicStation As Dictionary
        Set dicStation = dicPoint("Station")
        tPoint.Station = m_Client.CreateStation(dicStation)
        
        ' Prefecture
        Dim dicPrefecture As Dictionary
        Set dicPrefecture = dicPoint("Prefecture")
        tPoint.Prefecture = m_Client.CreatePrefecture(dicPrefecture)
        Points(i) = tPoint
    Next i
    CreatePoints = Points
End Function


Property Let Client(ByVal Client As Ekispert)
    Set m_Client = Client
End Property

Property Let Name(ByVal Name As String)
    m_Name = Name
End Property

Property Get Name() As String
    Name = m_Name
End Property

Property Let code(ByVal code As Long)
    m_Code = code
End Property

Property Let TransportType(i As Long, ByVal TransportType As TransportTypeEnum)
    ReDim Preserve m_TransportType(i) As String
     m_TransportType(i) = m_Client.GetTransportType(TransportType)
End Property

Property Let PrefectureCode(i As Long, code As PrefectureEnum)
    ReDim Preserve m_PrefectureCode(i) As String
    m_PrefectureCode(i) = CStr(code)
End Property

Property Let CorporationBind(ByVal CorporationBind As String)
    m_CorporationBind = CorporationBind
End Property

Property Let CommunityBus(ByVal CommunityBus As CommunityBusEnum)
    m_CommunityBus = CommunityBus
End Property

Property Let NameMatchType(ByVal NameMatchType As NameMatchTypeEnum)
    m_NameMatchType = NameMatchType
End Property


Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    If m_Name <> "" Then
        BuildQuery.Add "name", m_Name
    End If
    If m_Code > 0 Then
        BuildQuery.Add "code", m_Code & ""
    End If
    BuildQuery.Add "nameMatchType", IIf(m_NameMatchType = Forward, "forward", "partial")
    If (Not m_TransportType) <> -1 Then
        BuildQuery.Add "type", Join(m_TransportType, ":")
    End If
    If (Not m_PrefectureCode) <> -1 Then
        BuildQuery.Add "prefectureCode", Join(m_PrefectureCode, ":")
    End If
    If m_CorporationBind <> "" Then
        BuildQuery.Add "corporationBind", m_CorporationBind
    End If
    BuildQuery.Add "communityBus", IIf(m_CommunityBus = Contain, "contain", "except")
End Function


