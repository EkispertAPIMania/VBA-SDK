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
Private m_NameMatchType As eNameMatchType
Private m_Code As Long
Private m_TransportType As String
Private m_PrefectureCode As Integer
Private m_CorporationBind As String
Private m_CommunityBus As String

Private m_Client As Ekispert

Private Sub Class_Initialize()
    m_NameMatchType = eNameMatchType.Forward
    m_CommunityBus = "contain"
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


Property Let client(ByVal client As Ekispert)
    Set m_Client = client
End Property

Property Let Name(ByVal Name As String)
    m_Name = Name
End Property

Property Get Name() As String
    Name = m_Name
End Property

Property Let Code(ByVal Code As Long)
    m_Code = Code
End Property

Property Let TransportType(ByVal TransportType As String)
     m_TransportType = TransportType
End Property

Property Let PrefectureCode(ByVal PrefectureCode As Integer)
    m_PrefectureCode = PrefectureCode
End Property

Property Let CorporationBind(ByVal CorporationBind As String)
    m_CorporationBind = CorporationBind
End Property

Property Let CommunityBus(ByVal CommunityBus As String)
    m_CommunityBus = CommunityBus
End Property

Property Let NameMatchType(ByVal NameMatchType As eNameMatchType)
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
    If m_TransportType <> "" Then
        BuildQuery.Add "type", m_TransportType
    End If
    If m_PrefectureCode > 0 Then
        BuildQuery.Add "prefectureCode", m_PrefectureCode
    End If
    If m_CorporationBind <> "" Then
        BuildQuery.Add "corporationBind", m_CorporationBind
    End If
    If m_CommunityBus <> "" Then
        BuildQuery.Add "communityBus", m_CommunityBus
    End If
End Function


