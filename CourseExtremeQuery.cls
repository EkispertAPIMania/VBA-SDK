VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CourseExtremeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_ViaList() As String
Private m_ExcludeSameLineStation As Boolean
Private m_FixedRailList() As String
Private m_FixedRailDirectionList() As String
Private m_Date As Date
Private m_SearchType As SearchTypeEnum
Private m_Sort As SortEnum
Private m_AnswerCount As Long
Private m_SearchCount As Long
Private m_ConditionDetail() As String
Private m_CorporationBind() As String
Private m_InterruptCorporationList() As String
Private m_InterruptRailList() As String
Private m_InterruptOperationLineCodeList() As String
Private m_InterruptTransferStationCodeList() As String
Private m_ResultDetail As String
Private m_AddOperationLinePattern As Boolean
Private m_CheckEngineVersion As Boolean
Private m_AssignTeikiSerializeData As String
Private m_AssignRoute() As String
Private m_AssignDetailRoute() As String
Private m_OffpeakTeikiMode As OffpeakTeikiModeEnum
Private m_AssignPassClassIndex As Long
Private m_Coupon As String
Private m_AddAssignStatus As Boolean
Private m_AddChange As Boolean
Private m_AddStop As Boolean
Private m_Gcs As GCSEnum

Private m_Client As Ekispert

Private Sub Class_Initialize()
    m_ExcludeSameLineStation = False
    m_Date = Now
    m_SearchType = Departure
    m_Sort = Ekispert
    
    m_AnswerCount = 5
    m_SearchCount = 20
    m_ResultDetail = ""
    m_AddOperationLinePattern = False
    m_CheckEngineVersion = True
    m_AssignTeikiSerializeData = ""
    m_OffpeakTeikiMode = OffpeakTime
    m_Coupon = ""
    m_AddAssignStatus = False
    m_AddChange = False
    m_AddStop = False
    m_Gcs = GCSEnum.Tokyo
End Sub

Public Function Find() As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/search/course/extreme"
    Dim Query As Dictionary
    Set Query = BuildQuery()
    Dim Result As Dictionary
    Set Result = m_Client.Fetch(ApiPath, Query)
    
    Dim dicCourses() As Dictionary
    Dim Response As ResultSet
    
    If m_Client.Success = False Then
        Response.Success = False
        Response.Error = m_Client.Error
        Find = Response
        Exit Function
    End If
    m_Client.SetArray dicCourses, Result, "Course"
    Response.Courses = CreateCourses(dicCourses)
    Response.Success = True
    Find = Response

    ' Debug.Print JsonConverter.ConvertToJson(Result)
End Function

Function CreateCourses(dicCourses() As Dictionary) As Course()
    Dim i As Long
    Dim Courses() As Course
    ReDim Courses(UBound(dicCourses) - 1) As Course
    Dim dicPrices() As Dictionary
    For i = 0 To UBound(dicCourses) - 1
        Dim dicCourse As Dictionary
        Set dicCourse = dicCourses(i)
        Dim tCourse As Course
        tCourse.SerializeData = dicCourse("SerializeData")
        tCourse.dataType = dicCourse("dataType")
        tCourse.searchType = dicCourse("searchType")
        
        ' Price
        m_Client.SetArray dicPrices, dicCourse, "Price"
        tCourse.Prices = CreatePrices(dicPrices)

        ' Route
        tCourse.Route = CreateRoute(dicCourse)
        
        ' Teiki
        tCourse.Teiki = CreateTeiki(dicCourse)
        Courses(i) = tCourse
    Next i
    CreateCourses = Courses
End Function

Function CreatePrices(dicPrices() As Dictionary) As Price()
    Dim i As Long
    Dim Prices() As Price
    ReDim Prices(UBound(dicPrices) - 1) As Price
    
    Dim dicPrice As Dictionary
    For i = 0 To UBound(dicPrices) - 1
        Set dicPrice = dicPrices(i)
        Dim tPrice As Price
        tPrice.Oneway = CLng(dicPrice("Oneway"))
        tPrice.RevisionStatus = dicPrice("RevisionStatus")
        tPrice.Round = CLng(dicPrice("Round"))
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
    Set dicTeiki = dicCourse("Teiki")
    CreateTeiki.DisplayRoute = dicTeiki("DisplayRoute")
    CreateTeiki.SerializeData = dicTeiki("SerializeData")
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
    CreateRoute.transferCount = dicRoute("transferCount")
    ' Point
    Dim dicPoints() As Dictionary
    m_Client.SetArray dicPoints, dicRoute, "Point"
    CreateRoute.Points = CreatePoints(dicPoints)
    
    ' Line
    Dim dicLines() As Dictionary
    m_Client.SetArray dicLines, dicRoute, "Line"
    
    CreateRoute.Lines = CreateLines(dicLines)
End Function

Function CreateLines(dicLines() As Dictionary) As Line()
    Dim i As Long
    Dim Lines() As Line
    ReDim Lines(UBound(dicLines) - 1) As Line
    
    Dim dicLine As Dictionary
    For i = 0 To UBound(dicLines) - 1
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
    DateTime.text = dicDateTime("text")
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
    DateTime.text = dicDateTime("text")
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
    ReDim Points(UBound(dicPoints) - 1) As Point
    
    Dim dicPoint As Dictionary
    For i = 0 To UBound(dicPoints) - 1
        Set dicPoint = dicPoints(i)
        Dim tPoint As Point
        Dim GeoPoint As Dictionary
        Set GeoPoint = dicPoint("GeoPoint")
        Dim Prefecture As Dictionary
        Set Prefecture = dicPoint("Prefecture")
        Dim Station As Dictionary
        Set Station = dicPoint("Station")
        
        tPoint.GeoPoint = m_Client.CreateGeoPoint(GeoPoint)
        tPoint.Prefecture = m_Client.CreatePrefecture(Prefecture)
        tPoint.Station = m_Client.CreateStation(Station)
        Points(i) = tPoint
    Next i
    CreatePoints = Points
End Function

Property Let Client(ByVal Client As Ekispert)
    Set m_Client = Client
End Property

Public Property Let ViaList(i As Long, ByVal value As String)
    ReDim Preserve m_ViaList(i) As String
    m_ViaList(i) = value
End Property

Public Property Let ExcludeSameLineStation(ByVal value As Boolean)
    m_ExcludeSameLineStation = value
End Property

Public Property Let FixedRailList(i As Long, ByVal value As String)
    ReDim Preserve m_FixedRailList(i) As String
    m_FixedRailList(i) = value
End Property

Public Property Let FixedRailDirectionList(i As Long, ByVal value As String)
    ReDim Preserve m_FixedRailDirectionList(i) As String
    m_FixedRailDirectionList(i) = value
End Property

Public Property Let DateProp(ByVal value As Date)
    m_Date = value
End Property

Public Property Let searchType(ByVal value As SearchTypeEnum)
    m_SearchType = value
End Property

Public Property Let Sort(ByVal value As SortEnum)
    m_Sort = value
End Property

Public Property Let AnswerCount(ByVal value As Long)
    m_AnswerCount = value
End Property

Public Property Let SearchCount(ByVal value As Long)
    m_SearchCount = value
End Property

Public Property Let ConditionDetail(i As Long, ByVal value As String)
    ReDim Preserve m_ConditionDetail(i) As String
    m_ConditionDetail(i) = value
End Property

Public Property Let CorporationBind(i As Long, ByVal value As String)
    ReDim Preserve m_CorporationBind(i) As String
    m_CorporationBind(i) = value
End Property

Public Property Let InterruptCorporationList(i As Long, ByVal value As String)
    ReDim Preserve m_InterruptCorporationList(i) As String
    m_InterruptCorporationList(i) = value
End Property

Public Property Let InterruptRailList(i As Long, ByVal value As String)
    ReDim Preserve m_InterruptRailList(i) As String
    m_InterruptRailList(i) = value
End Property

Public Property Let InterruptOperationLineCodeList(i As Long, ByVal value As String)
    ReDim Preserve m_InterruptOperationLineCodeList(i) As String
    m_InterruptOperationLineCodeList(i) = value
End Property

Public Property Let InterruptTransferStationCodeList(i As Long, ByVal value As String)
    ReDim Preserve m_InterruptTransferStationCodeList(i) As String
    m_InterruptTransferStationCodeList(i) = value
End Property

Public Property Let ResultDetail(ByVal value As String)
    m_ResultDetail = value
End Property

Public Property Let AddOperationLinePattern(ByVal value As Boolean)
    m_AddOperationLinePattern = value
End Property

Public Property Let CheckEngineVersion(ByVal value As Boolean)
    m_CheckEngineVersion = value
End Property

Public Property Let AssignTeikiSerializeData(ByVal value As String)
    m_AssignTeikiSerializeData = value
End Property

Public Property Let AssignRoute(i As Long, ByVal value As String)
    ReDim Preserve m_AssignRoute(i) As String
    m_AssignRoute(i) = value
End Property

Public Property Let AssignDetailRoute(i As Long, ByVal value As String)
    ReDim Preserve m_AssignDetailRoute(i) As String
    m_AssignDetailRoute(i) = value
End Property

Public Property Let OffpeakTeikiMode(ByVal value As OffpeakTeikiModeEnum)
    m_OffpeakTeikiMode = value
End Property

Public Property Let AssignPassClassIndex(ByVal value As Long)
    m_AssignPassClassIndex = value
End Property

Public Property Let Coupon(ByVal value As String)
    m_Coupon = value
End Property

Public Property Let BringAssignmentError(ByVal value As Boolean)
    m_BringAssignmentError = value
End Property

Public Property Let AddAssignStatus(ByVal value As Boolean)
    m_AddAssignStatus = value
End Property

Public Property Let AddChange(ByVal value As Boolean)
    m_AddChange = value
End Property

Public Property Let AddStop(ByVal value As Boolean)
    m_AddStop = value
End Property

Public Property Let GCS(ByVal value As GCSEnum)
    m_Gcs = value
End Property

Public Property Get SearchTypeString() As String
    Select Case m_SearchType
    Case SearchTypeEnum.Arrival
        SearchTypeString = "arrival"
    Case SearchTypeEnum.Departure
        SearchTypeString = "departure"
    Case SearchTypeEnum.LastTrain
        SearchTypeString = "lastTrain"
    Case SearchTypeEnum.FirstTrain
        SearchTypeString = "firstTrain"
    Case SearchTypeEnum.Plain
        SearchTypeString = "plain"
    Case Else
        SearchTypeString = ""
    End Select
End Property

Public Property Get SortString() As String
    Select Case m_Sort
    Case SortEnum.Ekispert
        SortString = "ekispert"
    Case SortEnum.Price
        SortString = "price"
    Case SortEnum.Time
        SortString = "time"
    Case SortEnum.Teiki
        SortString = "teiki"
    Case SortEnum.Transfer
        SortString = "transfer"
    Case SortEnum.Co2
        SortString = "co2"
    Case SortEnum.Teiki1
        SortString = "teiki1"
    Case SortEnum.Teiki3
        SortString = "teiki3"
    Case SortEnum.Teiki6
        SortString = "teiki6"
    End Select
End Property

Public Property Get OffpeakTeikiModeString() As String
    Select Case m_OffpeakTeikiMode
    Case OffpeakTeikiModeEnum.OffpeakTime
        OffpeakTeikiModeString = "offpeakTime"
    Case OffpeakTeikiModeEnum.OffpeakTime
        OffpeakTeikiModeString = "peakTime"
    End Select
End Property

Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    BuildQuery.Add "excludeSameLineStation", IIf(m_ExcludeSameLineStation, "true", "false")
    BuildQuery.Add "date", Format(m_Date, "yyyymmdd")
    BuildQuery.Add "time", Format(m_Date, "hnn")
    BuildQuery.Add "searchType", SearchTypeString()
    BuildQuery.Add "sort", SortString()
    BuildQuery.Add "answerCount", m_AnswerCount
    BuildQuery.Add "searchCount", m_SearchCount
    If m_ResultDetail <> "" Then
        BuildQuery.Add "resultDetail", m_ResultDetail
    End If
    BuildQuery.Add "addOperationLinePattern", IIf(m_AddOperationLinePattern, "true", "false")
    BuildQuery.Add "checkEngineVersion", IIf(m_CheckEngineVersion, "true", "false")
    If m_AssignTeikiSerializeData <> "" Then
        BuildQuery.Add "assignTeikiSerializeData", m_AssignTeikiSerializeData
    End If
    If (Not m_AssignRoute) <> -1 Or (Not m_AssignDetailRoute) <> -1 Or m_AssignTeikiSerializeData <> "" Then
        BuildQuery.Add "offpeakTeikiMode", OffpeakTeikiModeString()
    End If
    BuildQuery.Add "assignPassClassIndex", m_AssignPassClassIndex
    If m_Coupon <> "" Then
        BuildQuery.Add "coupon", m_Coupon
    End If
    If m_BringAssignmentError = True Then
        BuildQuery.Add "bringAssignmentError", IIf(m_BringAssignmentError, "true", "false")
    ElseIf m_AddAssignStatus = True Then
        BuildQuery.Add "addAssignStatus", IIf(m_AddAssignStatus, "true", "false")
    End If
    BuildQuery.Add "addChange", IIf(m_AddChange, "true", "false")
    BuildQuery.Add "addStop", IIf(m_AddStop, "true", "false")
    If m_Client.GCSString(m_Gcs) <> "" Then
        BuildQuery.Add "gcs", m_Client.GCSString(m_Gcs)
    End If
    
    m_Client.ArrayToString BuildQuery, "viaList", m_ViaList
    m_Client.ArrayToString BuildQuery, "fixedRailList", m_FixedRailList
    m_Client.ArrayToString BuildQuery, "fixedRailDirectionList", m_FixedRailDirectionList
    m_Client.ArrayToString BuildQuery, "conditionDetail", m_ConditionDetail
    m_Client.ArrayToString BuildQuery, "corporationBind", m_CorporationBind
    m_Client.ArrayToString BuildQuery, "interruptCorporationList", m_InterruptCorporationList
    m_Client.ArrayToString BuildQuery, "interruptRailList", m_InterruptRailList
    m_Client.ArrayToString BuildQuery, "interruptOperationLineCodeList", m_InterruptOperationLineCodeList
    m_Client.ArrayToString BuildQuery, "interruptTransferStationCodeList", m_InterruptTransferStationCodeList
    m_Client.ArrayToString BuildQuery, "assignRoute", m_AssignRoute
    m_Client.ArrayToString BuildQuery, "assignDetailRoute", m_AssignDetailRoute

End Function


