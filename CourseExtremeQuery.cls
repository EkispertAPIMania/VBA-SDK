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
    m_AssignPassClassIndex = 0
    Erase m_ViaList
    Erase m_FixedRailList
    Erase m_FixedRailDirectionList
    Erase m_ConditionDetail
    Erase m_CorporationBind
    Erase m_InterruptCorporationList
    Erase m_InterruptRailList
    Erase m_InterruptOperationLineCodeList
    Erase m_InterruptTransferStationCodeList
    Erase m_AssignRoute
    Erase m_AssignDetailRoute

End Sub

Public Sub Reset()
    Class_Initialize
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
    ' Debug.Print JsonConverter.ConvertToJson(dicCourses)
    Response.Courses = m_Client.CreateCourses(dicCourses)
    Response.Success = True
    Find = Response

    ' Debug.Print JsonConverter.ConvertToJson(Result)
End Function

Property Let client(ByVal client As Ekispert)
    Set m_Client = client
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

Public Property Let SearchType(ByVal value As SearchTypeEnum)
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

Public Property Let Gcs(ByVal value As GCSEnum)
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
    BuildQuery.Add "time", Format(m_Date, "hhnn")
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


