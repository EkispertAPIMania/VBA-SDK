Attribute VB_Name = "EkispertType"
Type GeoPoint
    Longitude_DMS As String
    Latitude_DMS As String
    Longitude_DD As String
    Latitude_DD As String
    Gcs As String
End Type

Type Prefecture
    Code As String
    Name As String
End Type

Type Gate
    Code As String
    Name As String
    GeoPoint As GeoPoint
End Type

Type TransportType
    Text As String
    Detail As String
End Type

Type Station
    Code As String
    Name As String
    OldName As String
    Yomi As String
    TransportType As TransportType
    Gate() As Gate
End Type

Type Point
    GetOn As Boolean
    GetOff As Boolean
    OnRoute As Boolean
    OnRouteEdge As Boolean
    GeoPoint As GeoPoint
    Prefecture As Prefecture
    Station As Station
End Type

Type EkispertError
    ApiVersion As String
    EngineVersion As String
    Code As String
    Message As String
End Type

Type ResultSet
    Max As Long
    Offset As Long
    RoundTripType As String
    Points() As Point
    Success As Boolean
    Error As EkispertError
End Type

Public Enum eNameMatchType
    Forward = 1
    Partial = 2
End Enum


