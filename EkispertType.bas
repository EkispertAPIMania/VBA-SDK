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

Public Enum NameMatchTypeEnum
    Forward = 1
    Partial = 2
End Enum

Public Enum PrefectureEnum
    hokkaido = 1
    Aomori = 2
    Iwate = 3
    Miyagi = 4
    Akita = 5
    Yamagata = 6
    Fukushima = 7
    Ibaraki = 8
    Tochigi = 9
    Gunma = 10
    Saitama = 11
    Chiba = 12
    Tokyo = 13
    Kanagawa = 14
    Niigata = 15
    Toyama = 16
    Ishikawa = 17
    Fukui = 18
    Yamanashi = 19
    Nagano = 20
    Gifu = 21
    Shizuoka = 22
    Aichi = 23
    Mie = 24
    Shiga = 25
    Kyoto = 26
    Osaka = 27
    Hyogo = 28
    Nara = 29
    Wakayama = 30
    Tottori = 31
    Shimane = 32
    Okayama = 33
    Hiroshima = 34
    Yamaguchi = 35
    Tokushima = 36
    Kagawa = 37
    Ehime = 38
    Kochi = 39
    Fukuoka = 40
    Saga = 41
    Nagasaki = 42
    Kumamoto = 43
    Oita = 44
    Miyazaki = 45
    Kagoshima = 46
    Okinawa = 47
End Enum

Enum TransportTypeEnum
    Train = 0
    Plane = 1
    Ship = 2
    Bus = 3
    Walk = 4
    Strange = 5
    LocalBus = 6
    ConnectionBus = 7
    HighwayBus = 8
    MidnightBus = 9
End Enum

Enum CommunityBusEnum
    Contain = 1
    Except = 2
End Enum
