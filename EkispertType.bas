Attribute VB_Name = "EkispertType"
Type GeoPoint
    Longitude_DMS As String
    Latitude_DMS As String
    Longitude_DD As String
    Latitude_DD As String
    Gcs As String
End Type

Type Prefecture
    code As String
    Name As String
End Type

Type Gate
    code As String
    Name As String
    GeoPoint As GeoPoint
End Type

Type TransportType
    Text As String
    Detail As String
End Type

Type Station
    code As String
    Name As String
    OldName As String
    Yomi As String
    TransportType As TransportType
    Gate() As Gate
End Type

Type Cost
    Minute As Long
    TransferCount As Long
    BaseIndex As Long
End Type

Type Point
    GetOn As Boolean
    GetOff As Boolean
    OnRoute As Boolean
    OnRouteEdge As Boolean
    GeoPoint As GeoPoint
    Prefecture As Prefecture
    Station As Station
    Costs() As Cost
End Type

Type EkispertError
    ApiVersion As String
    EngineVersion As String
    code As String
    Message As String
End Type

Type DateTime
    Text As String
    operation As String
End Type


Type ArrivalState
    no As String
    Type As String
    DateTime As DateTime
End Type

Type DepartureState
    no As String
    Type As String
    DateTime As DateTime
    isStarting As String
End Type

Type LineSymbol
    code As String
    Name As String
End Type


Type Line
    stopStationCount As String
    teiki3Index As String
    teiki6Index As String
    timeOnBoard As String
    track As String
    exhaustCO2 As String
    fareIndex As String
    exhaustCO2atPassengerCar As String
    distance As String
    trainID As String
    teiki1Index As String
    Name As String
    Type As String
    ArrivalState As ArrivalState
    Destination As String
    TimeReliability As String
    DepartureState As DepartureState
    LineSymbol As LineSymbol
    Color As String
End Type

Type Remark
    Text As Long
    Remark As String
    FullRemark As String
End Type


Type Price
    FareRevisionStatus As String
    ToLineIndex As Long
    FromLineIndex As Long
    Kind As String
    Index As String
    Selected As Boolean
    Type As String
    Oneway As Long
    OnewayRemark As Remark
    RevisionStatus As String
    Round As Long
    RoundRemark As Remark
End Type

Type Teiki
    SerializeData As String
    DisplayRoute As String
End Type

Type Route
    timeOther As String
    timeOnBoard As String
    exhaustCO2 As String
    exhaustCO2atPassengerCar As String
    distance As String
    timeWalk As String
    TransferCount As String
    Lines() As Line
    Points() As Point
End Type

Type Course
    searchType As String
    dataType As String
    SerializeData As String
    Prices() As Price
    Teiki As Teiki
    Route As Route
End Type

Type Base
    Point As Point
End Type

Type ResultSet
    Max As Long
    Offset As Long
    RoundTripType As String
    Points() As Point
    Courses() As Course
    Bases() As Base
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

Enum GCSEnum
    Tokyo = 1
    Wgs84 = 2
End Enum

Enum SearchTypeEnum
    Departure = 1
    Arrival = 2
    LastTrain = 3
    FirstTrain = 4
    Plain = 5
End Enum

Enum SortEnum
    Ekispert = 1
    Price = 2
    Time = 3
    Teiki = 4
    Transfer = 5
    Co2 = 6
    Teiki1 = 7
    Teiki3 = 8
    Teiki6 = 9
End Enum

Enum OffpeakTeikiModeEnum
    OffpeakTime = 0
    PeakTime = 1
End Enum

