VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StationQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Name As String
Private m_OldName As String
Private m_Code As Long
Private m_CorporationName As String
Private m_RailName As String
Private m_OperationLineCode As String
Private m_TransportType() As String
Private m_PrefectureCode() As String
Private m_Offset As Integer ' Default: 1
Private m_Limit As Integer ' Default: 100
Private m_Direction As String ' Default: up
Private m_CorporationBind As String
Private m_AddGateGroup As Boolean ' Default: false
Private m_CommunityBus  As CommunityBusEnum ' Default: contain
Private m_Gcs As GCSEnum

Private m_Client As Ekispert

Private Sub Class_Initialize()
    m_Offset = 1
    m_Limit = 100
    m_Direction = "up"
    m_AddGateGroup = False
    m_CommunityBus = Contain
    m_Gcs = GCSEnum.Tokyo
End Sub

Public Function Find(Optional Light As Boolean = False) As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/station"
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
    ' 往復の種類を表す要素。railNameを指定した時のみ付加します。
    If Result.Exists("RoundTripType") = True Then
        Response.RoundTripType = Result("RoundTripType")
    End If
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
        If dicPoint.Exists("getOff") = True Then
            tPoint.GetOff = dicPoint("getOff")
        End If
        If dicPoint.Exists("getOn") = True Then
            tPoint.GetOn = dicPoint("getOn")
        End If
        If dicPoint.Exists("onRoute") = True Then
            tPoint.OnRoute = dicPoint("onRoute")
        End If
        If dicPoint.Exists("onRouteEdge") = True Then
            tPoint.OnRouteEdge = dicPoint("onRouteEdge")
        End If
        ' Station
        Dim dicStation As Dictionary
        Set dicStation = dicPoint("Station")
        tPoint.Station = m_Client.CreateStation(dicStation)
        
        ' Prefecture
        Dim dicPrefecture As Dictionary
        Set dicPrefecture = dicPoint("Prefecture")
        tPoint.Prefecture = m_Client.CreatePrefecture(dicPrefecture)
        
        ' GeoPoint
        Dim dicGeoPoint As Dictionary
        Set dicGeoPoint = dicPoint("GeoPoint")
        tPoint.GeoPoint = m_Client.CreateGeoPoint(dicGeoPoint)
        
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

Property Let OldName(ByVal OldName As String)
    m_OldName = OldName
End Property

Property Let code(ByVal code As Long)
    m_Code = code
End Property

Property Let CorporationName(ByVal CorporationName As String)
    m_CorporationName = CorporationName
End Property

Property Let RailName(ByVal RailName As String)
    m_RailName = RailName
End Property

Property Let OperationLineCode(ByVal OperationLineCode As String)
    m_OperationLineCode = OperationLineCode
End Property

Property Let TransportType(i As Long, ByVal TransportType As TransportTypeEnum)
    ReDim Preserve m_TransportType(i) As String
     m_TransportType(i) = m_Client.GetTransportType(TransportType)
End Property

Property Let PrefectureCode(i As Long, code As Integer)
    ReDim Preserve m_PrefectureCode(i) As String
    m_PrefectureCode(i) = CStr(code)
End Property

Property Let Offset(ByVal Offset As Integer)
    m_Offset = Offset
End Property

Property Let Limit(ByVal Limit As Integer)
    m_Limit = Limit
End Property

Property Let Direction(ByVal Direction As String)
    m_Direction = Direction
End Property

Property Let CorporationBind(ByVal CorporationBind As String)
    m_CorporationBind = CorporationBind
End Property

Property Let AddGateGroup(ByVal AddGateGroup As Boolean)
    m_AddGateGroup = AddGateGroup
End Property

Property Let CommunityBus(ByVal CommunityBus As CommunityBusEnum)
    m_CommunityBus = CommunityBus
End Property

Property Let Gcs(ByVal value As GCSEnum)
    m_Gcs = value
End Property


Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    If m_Name <> "" Then
        BuildQuery.Add "name", m_Name
    End If
    If m_OldName <> "" Then
        BuildQuery.Add "oldName", m_OldName
    End If
    If m_Code > 0 Then
        BuildQuery.Add "code", m_Code & ""
    End If
    If m_CorporationName <> "" Then
        BuildQuery.Add "corporationName", m_CorporationName
    End If
    If m_RailName <> "" Then
        BuildQuery.Add "railName", m_RailName
    End If
    If m_OperationLineCode <> "" Then
        BuildQuery.Add "operationLineCode", m_OperationLineCode
    End If
    If (Not m_TransportType) <> -1 Then
        BuildQuery.Add "type", Join(m_TransportType, ":")
    End If
    If (Not m_PrefectureCode) <> -1 Then
        BuildQuery.Add "prefectureCode", Join(m_PrefectureCode, ":")
    End If
    If m_Offset > 0 Then
        BuildQuery.Add "offset", m_Offset
    End If
    If m_Limit > 0 Then
        BuildQuery.Add "limit", m_Limit
    End If
    If m_Direction <> "" Then
        BuildQuery.Add "direction", m_Direction
    End If
    If m_CorporationBind <> "" Then
        BuildQuery.Add "corporationBind", m_CorporationBind
    End If
    BuildQuery.Add "addGateGroup", IIf(m_AddGateGroup, "true", "false")
    BuildQuery.Add "communityBus", IIf(m_CommunityBus = Contain, "contain", "except")
    If m_Client.GCSString(m_Gcs) <> "" Then
        BuildQuery.Add "gcs", m_Client.GCSString(m_Gcs)
    End If
    
End Function

