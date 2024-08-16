VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ekispert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_ApiKey As String    ' APIキー
Private m_EndPoint As String  ' エンドポイントドメイン
Private m_Secure As Boolean   ' TrueならHTTPS/FalseならHTTP
Private m_Error As EkispertError
Private m_Success As Boolean

Private Sub Class_Initialize()
    m_ApiKey = ""
    m_EndPoint = "api.ekispert.jp"
    m_secture = True
    m_Success = False
End Sub

Property Let Secure(ByVal Secure As Boolean)
    m_Secure = Secure
End Property

Property Let ApiKey(ByVal ApiKey As String)
    m_ApiKey = ApiKey
End Property

Property Get Secure() As Boolean
    Secure = m_Secure
End Property

Property Get ApiKey() As String
    ApiKey = m_ApiKey
End Property

' 駅検索クエリーオブジェクトを返す
Function StationQuery() As StationQuery
    Set StationQuery = New StationQuery
    StationQuery.Client = Me
End Function

' 簡易駅検索クエリーオブジェクトを返す
Function StationLightQuery() As StationLightQuery
    Set StationLightQuery = New StationLightQuery
    StationLightQuery.Client = Me
End Function

' 経路探索クエリーオブジェクトを返す
Function CourseExtremeQuery() As CourseExtremeQuery
    Set CourseExtremeQuery = New CourseExtremeQuery
    CourseExtremeQuery.Client = Me
End Function

' 経路探索クエリーオブジェクトを返す
Function CoursePlainQuery() As CoursePlainQuery
    Set CoursePlainQuery = New CoursePlainQuery
    CoursePlainQuery.Client = Me
End Function

' 緯度経度からの周辺駅検索クエリーオブジェクトを返す
Function GeoStationQuery() As GeoStationQuery
    Set GeoStationQuery = New GeoStationQuery
    GeoStationQuery.Client = Me
End Function

Function SearchMultipleRangeQuery() As SearchMultipleRangeQuery
    Set SearchMultipleRangeQuery = New SearchMultipleRangeQuery
    SearchMultipleRangeQuery.Client = Me
End Function


Property Get Url() As String
    If m_Secure = True Then
        Url = "https"
    Else
        Url = "http"
    End If
    Url = Url & "://" & m_EndPoint
End Property

Function Fetch(ApiPath As String, Query As Dictionary) As Dictionary
    Dim e As EkispertError
    m_Error = e
    m_Success = False
    Dim Client As New WebClient
    Client.BaseUrl = Url
    
    Dim DirectionsRequest As New WebRequest
    DirectionsRequest.Method = WebMethod.HttpGet
    Query.Add "key", m_ApiKey
    
    Dim QueryString() As String
    ReDim QueryString(UBound(Query.Items))
    
    Dim i As Integer
    For i = 0 To UBound(Query.Items)
        QueryString(i) = Query.Keys(i) & "=" & WorksheetFunction.EncodeURL(Query.Item(Query.Keys(i)))
    Next i
    
    DirectionsRequest.Resource = ApiPath & "?" & Join(QueryString, "&")
    Dim Response As WebResponse
    Set Response = Client.Execute(DirectionsRequest)
    ' Debug.Print Client.GetFullUrl(DirectionsRequest)
    
    If Response.StatusCode <> WebStatusCode.Ok Then
        Dim Error As Dictionary
        Set Error = Response.Data("ResultSet")
        BuildError Error
        Exit Function
    End If
    m_Success = True
    Set Fetch = Response.Data("ResultSet")
End Function

Sub BuildError(Error As Dictionary)
    m_Error.ApiVersion = Error("apiVersion")
    m_Error.EngineVersion = Error("engineVersion")
    Dim e As Dictionary
    Set e = Error("Error")
    m_Error.code = e("code")
    m_Error.Message = e("Message")
End Sub

Property Get Error() As EkispertError
    Error = m_Error
End Property

Property Get Success() As Boolean
    Success = m_Success
End Property

Function CreateCourses(dicCourses() As Dictionary) As Course()
    Dim i As Long
    Dim Courses() As Course
    ReDim Courses(UBound(dicCourses)) As Course
    Dim dicPrices() As Dictionary
    For i = 0 To UBound(dicCourses)
        Dim dicCourse As Dictionary
        Set dicCourse = dicCourses(i)
        Dim tCourse As Course
        tCourse.SerializeData = dicCourse("SerializeData")
        tCourse.dataType = dicCourse("dataType")
        tCourse.searchType = dicCourse("searchType")
        
        ' Price
        SetArray dicPrices, dicCourse, "Price"
        tCourse.Prices = CreatePrices(dicPrices)

        ' Route
        tCourse.Route = CreateRoute(dicCourse)
        
        ' Teiki
        tCourse.Teiki = CreateTeiki(dicCourse)
        Courses(i) = tCourse
    Next i
    CreateCourses = Courses
End Function

Sub LngOrDict(var As Variant, ByRef num As Long, ByRef dic As Dictionary)
        If IsNumeric(var) Then
           num = CLng(var)
        Else
            num = -1
            Set dic = var
        End If
End Sub

Function CreatePrices(dicPrices() As Dictionary) As Price()
    Dim i As Long
    Dim Prices() As Price
    ReDim Prices(UBound(dicPrices)) As Price
    Dim num As Long
    Dim dicOneway As Dictionary
    Dim dicRound As Dictionary
    Dim OnewayRemark As Remark
    Dim RoundRemark As Remark
    
    Dim dicPrice As Dictionary
    For i = 0 To UBound(dicPrices)
        Set dicPrice = dicPrices(i)
        Dim tPrice As Price
        
        LngOrDict dicPrice("Oneway"), num, dicOneway
        If num <> -1 Then
            tPrice.Oneway = num
        Else
            OnewayRemark.Text = CLng(dicOneway("text"))
            OnewayRemark.Remark = dicOneway("remark")
            OnewayRemark.FullRemark = dicOneway("fullRemark")
            tPrice.OnewayRemark = OnewayRemark
        End If
        
        tPrice.RevisionStatus = dicPrice("RevisionStatus")
        LngOrDict dicPrice("Round"), num, dicRound
        If num <> -1 Then
            tPrice.Round = num
        Else
            RoundRemark.Text = CLng(dicRound("text"))
            RoundRemark.Remark = dicRound("remark")
            RoundRemark.FullRemark = dicRound("fullRemark")
            tPrice.RoundRemark = RoundRemark
        End If
        tPrice.Type = dicPrice("Type")
        tPrice.FareRevisionStatus = dicPrice("fareRevisionStatus")
        tPrice.FromLineIndex = CLng(dicPrice("fromLineIndex"))
        tPrice.Index = CLng(dicPrice("index"))
        tPrice.Kind = dicPrice("kind")
        tPrice.Selected = IIf(dicPrice("selected") = "true", True, False)
        tPrice.ToLineIndex = CLng(dicPrice("toLineIndex"))
        Prices(i) = tPrice
    Next i
    CreatePrices = Prices
End Function

Function CreateTeiki(dicCourse As Dictionary) As Teiki
    Dim dicTeiki As Dictionary
    If dicCourse.Exists("Teiki") Then
        Set dicTeiki = dicCourse("Teiki")
        CreateTeiki.DisplayRoute = dicTeiki("DisplayRoute")
        CreateTeiki.SerializeData = dicTeiki("SerializeData")
    End If
End Function

Function CreateRoute(dicCourse As Dictionary) As Route
    Dim dicRoute As Dictionary
    Set dicRoute = dicCourse("Route")
    CreateRoute.distance = dicRoute("distance")
    CreateRoute.exhaustCO2 = dicRoute("exhaustCO2")
    CreateRoute.exhaustCO2atPassengerCar = dicRoute("exhaustCO2atPassengerCar")
    CreateRoute.timeOnBoard = dicRoute("timeOnBoard")
    CreateRoute.timeOther = dicRoute("timeOther")
    CreateRoute.timeWalk = dicRoute("timeWalk")
    CreateRoute.TransferCount = dicRoute("transferCount")
    ' Point
    Dim dicPoints() As Dictionary
    SetArray dicPoints, dicRoute, "Point"
    CreateRoute.Points = CreatePoints(dicPoints)
    
    ' Line
    Dim dicLines() As Dictionary
    SetArray dicLines, dicRoute, "Line"
    
    CreateRoute.Lines = CreateLines(dicLines)
End Function

Function CreateLines(dicLines() As Dictionary) As Line()
    Dim i As Long
    Dim Lines() As Line
    ReDim Lines(UBound(dicLines)) As Line
    
    Dim dicLine As Dictionary
    For i = 0 To UBound(dicLines)
        Set dicLine = dicLines(i)
        
        Lines(i) = CreateLine(dicLine)
    Next i
    CreateLines = Lines
End Function

Function CreateLine(dicLine As Dictionary) As Line
    Dim dicArrivalState As Dictionary
    Set dicArrivalState = dicLine("ArrivalState")
    Dim dicDateTime As Dictionary
    Set dicDateTime = dicArrivalState("Datetime")
    Dim DateTime As DateTime
    DateTime.operation = dicDateTime("operation")
    DateTime.Text = dicDateTime("text")
    Dim ArrivalState As ArrivalState
    ArrivalState.DateTime = DateTime
    ArrivalState.no = dicArrivalState("no")
    ArrivalState.Type = dicArrivalState("Type")
    CreateLine.ArrivalState = ArrivalState
    
    Dim dicDepartureState As Dictionary
    Set dicDepartureState = dicLine("DepartureState")
    Set dicDateTime = dicDepartureState("Datetime")
    Dim DepartureState As DepartureState
    DateTime.operation = dicDateTime("operation")
    DateTime.Text = dicDateTime("text")
    DepartureState.DateTime = DateTime
    DepartureState.isStarting = dicDepartureState("isStarting")
    DepartureState.no = dicDepartureState("no")
    DepartureState.Type = dicDepartureState("Type")
    
    
    Dim dicLineSymbol As Dictionary
    If dicLine.Exists("LineSymbol") Then
        Set dicLineSymbol = dicLine("LineSymbol")
        Dim LineSymbol As LineSymbol
        LineSymbol.code = dicLineSymbol("code")
        LineSymbol.Name = dicLineSymbol("Name")
        CreateLine.LineSymbol = LineSymbol
    End If
    CreateLine.Color = dicLine("Color")
    CreateLine.Destination = dicLine("Destination")
    CreateLine.Name = dicLine("Name")
    CreateLine.TimeReliability = dicLine("TimeReliability")
    CreateLine.Type = dicLine("Type")
    CreateLine.distance = CLng(dicLine("distance"))
    CreateLine.exhaustCO2 = CLng(dicLine("exhaustCO2"))
    CreateLine.exhaustCO2atPassengerCar = CLng(dicLine("exhaustCO2atPassengerCar"))
    CreateLine.fareIndex = CLng(dicLine("fareIndex"))
    CreateLine.stopStationCount = CLng(dicLine("stopStationCount"))
    CreateLine.teiki1Index = CLng(dicLine("teiki1Index"))
    CreateLine.teiki3Index = CLng(dicLine("teiki3Index"))
    CreateLine.teiki6Index = CLng(dicLine("teiki6Index"))
    CreateLine.timeOnBoard = CLng(dicLine("timeOnBoard"))
    CreateLine.track = dicLine("track")
    CreateLine.trainID = dicLine("trainID")

End Function

Function CreatePoints(dicPoints() As Dictionary) As Point()
    Dim i As Long
    Dim Points() As Point
    ReDim Points(UBound(dicPoints)) As Point
    
    Dim dicPoint As Dictionary
    For i = 0 To UBound(dicPoints)
        Set dicPoint = dicPoints(i)
        Points(i) = CreatePoint(dicPoint)
    Next i
    CreatePoints = Points
End Function

Function CreatePoint(dicPoint As Dictionary) As Point
    If dicPoint.Exists("GeoPoint") = True Then
        Dim GeoPoint As Dictionary
        Set GeoPoint = dicPoint("GeoPoint")
        CreatePoint.GeoPoint = CreateGeoPoint(GeoPoint)
    End If
    If dicPoint.Exists("Prefecture") = True Then
        Dim Prefecture As Dictionary
        Set Prefecture = dicPoint("Prefecture")
        CreatePoint.Prefecture = CreatePrefecture(Prefecture)
    End If
    If dicPoint.Exists("Station") = True Then
        Dim Station As Dictionary
        Set Station = dicPoint("Station")
        CreatePoint.Station = CreateStation(Station)
    End If
    If dicPoint.Exists("Cost") = True Then
        Dim dicCosts() As Dictionary
        SetArray dicCosts, dicPoint, "Cost"
        CreatePoint.Costs = CreateCosts(dicCosts)
    End If
End Function

Function CreateStation(dicStation As Dictionary) As Station
    Dim dicType As Dictionary
    If dicStation.Exists("code") = True Then
        CreateStation.code = dicStation("code")
    End If
    If dicStation.Exists("Name") = True Then
        CreateStation.Name = dicStation("Name")
    End If
    If dicStation.Exists("OldName") = True Then
        CreateStation.OldName = dicStation("OldName")
    End If
    If dicStation.Exists("Yomi") = True Then
        CreateStation.Yomi = dicStation("Yomi")
    End If
    If dicStation.Exists("Type") = True Then
        If VarType(dicStation("Type")) = VbVarType.vbString Then
            Dim tTransportType As TransportType
            tTransportType.Text = dicStation("Type")
            CreateStation.TransportType = tTransportType
        Else
            Set dicType = dicStation("Type")
            CreateStation.TransportType = CreateTransportType(dicType)
        End If
    End If
End Function

Function CreateCosts(dicCosts() As Dictionary) As Cost()
    Dim i As Long
    Dim Costs() As Cost
    ReDim Costs(UBound(dicCosts)) As Cost
    
    Dim dicCost As Dictionary
    For i = 0 To UBound(dicCosts)
        Set dicCost = dicCosts(i)
        Costs(i) = CreateCost(dicCost)
    Next i
    CreateCosts = Costs
End Function

Function CreateCost(dicCost As Dictionary) As Cost
    If dicCost.Exists("Minute") = True Then
        CreateCost.Minute = CLng(dicCost("Minute"))
    End If
    If dicCost.Exists("TransferCount") = True Then
        CreateCost.TransferCount = CLng(dicCost("TransferCount"))
    End If
    If dicCost.Exists("baseIndex") = True Then
        CreateCost.BaseIndex = CLng(dicCost("baseIndex"))
    End If
End Function

Function CreateBases(dicBases() As Dictionary) As Base()
    Dim i As Long
    Dim Bases() As Base
    Debug.Print UBound(dicBases)
    ReDim Bases(UBound(dicBases)) As Base
    
    Dim dicBase As Dictionary
    For i = 0 To UBound(dicBases)
        Set dicBase = dicBases(i)
        Bases(i) = CreateBase(dicBase)
    Next i
    CreateBases = Bases
End Function

Function CreateBase(dicBase As Dictionary) As Base
    Dim dicPoint As Dictionary
    If dicBase.Exists("Point") Then
        Set dicPoint = dicBase("Point")
        CreateBase.Point = CreatePoint(dicPoint)
    End If
End Function

' GeoPoint構造体の作成
Function CreateGeoPoint(dicGeoPoint As Dictionary) As GeoPoint
    If dicGeoPoint.Exists("gcs") = True Then
        CreateGeoPoint.Gcs = dicGeoPoint("gcs")
    End If
    If dicGeoPoint.Exists("lati_d") = True Then
        CreateGeoPoint.Latitude_DD = dicGeoPoint("lati_d")
    End If
    If dicGeoPoint.Exists("longi_d") = True Then
        CreateGeoPoint.Longitude_DD = dicGeoPoint("longi_d")
    End If
    If dicGeoPoint.Exists("lati") = True Then
        CreateGeoPoint.Latitude_DMS = dicGeoPoint("lati")
    End If
    If dicGeoPoint.Exists("longi") = True Then
        CreateGeoPoint.Longitude_DMS = dicGeoPoint("longi")
    End If
End Function


' Prefecture構造体の作成
Function CreatePrefecture(dicPrefecture As Dictionary) As Prefecture
    If dicPrefecture.Exists("code") = True Then
        CreatePrefecture.code = dicPrefecture("code")
    End If
    If dicPrefecture.Exists("Name") = True Then
        CreatePrefecture.Name = dicPrefecture("Name")
    End If
End Function


Function CreateTransportType(dictType As Dictionary) As TransportType
    If dictType.Exists("text") = True Then
        CreateTransportType.Text = dictType("text")
    End If
    If dictType.Exists("detail") = True Then
        CreateTransportType.Detail = dictType("detail")
    End If
End Function

Sub ArrayToString(ByRef dict As Dictionary, strKey As String, aryString() As String, Optional strJoin As String = ":")
    If (Not aryString) <> -1 Then
        dict.Add strKey, Join(aryString, strJoin)
    End If
End Sub

Public Property Get GCSString(m_Gcs As GCSEnum) As String
    Select Case m_Gcs
    Case GCSEnum.Tokyo
        GCSString = "tokyo"
    Case GCSEnum.Wgs84
        GCSString = "wgs84"
    Case Else
        GCSString = ""
    End Select
End Property

' 指定したキーが配列ではない場合、配列にする
' 配列の場合は、配列で返す
Sub SetArray(ByRef ary() As Dictionary, Data As Dictionary, Key As String)
On Error GoTo NotArray

    Dim Index As Long
    ReDim ary(Data(Key).Count - 1) As Dictionary
    Dim Params As Dictionary
    
    For Each Params In Data(Key)
        Set ary(Index) = Params
        Index = Index + 1
    Next Params
    Exit Sub
NotArray:
    ' 配列じゃない場合
    ReDim ary(0) As Dictionary
    Set ary(0) = Data(Key)
End Sub

Function GetTransportType(TransportType As TransportTypeEnum) As String
    Select Case TransportType
    Case TransportTypeEnum.Train
        GetTransportType = "train"
    Case TransportTypeEnum.Plane
        GetTransportType = "plane"
    Case TransportTypeEnum.Ship
        GetTransportType = "ship"
    Case TransportTypeEnum.Bus
        GetTransportType = "bus"
    Case TransportTypeEnum.Walk
        GetTransportType = "walk"
    Case TransportTypeEnum.Strange
        GetTransportType = "strange"
    Case TransportTypeEnum.LocalBus
        GetTransportType = "bus.local"
    Case TransportTypeEnum.ConnectionBus
        GetTransportType = "bus.connection"
    Case TransportTypeEnum.HighwayBus
        GetTransportType = "bus.highway"
    Case TransportTypeEnum.MidnightBus
        GetTransportType = "bus.midnight"
    End Select
End Function
