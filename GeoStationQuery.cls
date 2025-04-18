VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoStationQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_GeoPoint As String
Private m_TransportType() As String
Private m_CorporationBind() As String
Private m_AddGateGroup As Boolean
Private m_ExcludeSameLineStation As Boolean
Private m_StaionCount As Long
Private m_CommunityBus As CommunityBusEnum
Private m_Gcs As GCSEnum

Private m_Client As Ekispert

' Class Initialize method to set default values
Private Sub Class_Initialize()
    m_Gcs = GCSEnum.Tokyo
    m_CommunityBus = Contain
    m_ExcludeSameLineStation = False
    m_AddGateGroup = False
    
End Sub

Public Function Find() As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/geo/station"
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
    Response.Points = m_Client.CreatePoints(dicPoints)
    Response.Success = True
    Find = Response

End Function

Property Let client(ByVal client As Ekispert)
    Set m_Client = client
End Property

Public Sub GeoPoint(Latitude As String, Longitude As String, Optional Gcs As GCSEnum, Optional Radius As Long = 0)
    m_GeoPoint = Latitude & "," & Longitude
    If m_Client.GCSString(Gcs) <> "" Then
        m_GeoPoint = m_GeoPoint & "," & m_Client.GCSString(Gcs)
    End If
    
    If Radius > 0 Then
        m_GeoPoint = m_GeoPoint & "," & Radius
    End If
End Sub

Property Let TransportType(i As Long, ByVal TransportType As TransportTypeEnum)
    ReDim Preserve m_TransportType(i) As String
    m_TransportType(i) = m_Client.GetTransportType(TransportType)
End Property

Public Property Let CorporationBind(i As Long, ByVal value As String)
    ReDim Preserve m_CorporationBind(i) As String
    m_CorporationBind(i) = value
End Property

Property Let AddGateGroup(ByVal AddGateGroup As Boolean)
    m_AddGateGroup = AddGateGroup
End Property

Public Property Let ExcludeSameLineStation(ByVal value As Boolean)
    m_ExcludeSameLineStation = value
End Property

Public Property Let StaionCount(ByVal value As Long)
    m_StaionCount = value
End Property

Property Let CommunityBus(ByVal CommunityBus As CommunityBusEnum)
    m_CommunityBus = CommunityBus
End Property

Public Property Let Gcs(value As GCSEnum)
    m_Gcs = value
End Property


Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    BuildQuery.Add "geoPoint", m_GeoPoint
    m_Client.ArrayToString BuildQuery, "type", m_TransportType
    m_Client.ArrayToString BuildQuery, "corporationBind", m_CorporationBind
    BuildQuery.Add "addGateGroup", IIf(m_AddGateGroup, "true", "false")
    BuildQuery.Add "excludeSameLineStation", IIf(m_ExcludeSameLineStation, "true", "false")
    BuildQuery.Add "communityBus", IIf(m_CommunityBus = Contain, "contain", "except")
    If m_StaionCount > 0 Then
        BuildQuery.Add "stationCount", m_StaionCount
    End If
    If m_Client.GCSString(m_Gcs) <> "" Then
        BuildQuery.Add "gcs", m_Client.GCSString(m_Gcs)
    End If
End Function



