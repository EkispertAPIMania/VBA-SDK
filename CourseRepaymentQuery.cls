VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CourseRepaymentQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_SerializeData As String
Private m_CheckEngineVersion As Boolean
Private m_StartDate As Date
Private m_BuyDate As Date
Private m_RepaymentDate As Date
Private m_ValidityPeriod As Integer
Private m_ChangeSection As Boolean
Private m_Separator() As String
Private m_Client As Ekispert

' Class Initialize method to set default values
Private Sub Class_Initialize()
    m_CheckEngineVersion = True
    m_StartDate = Now
    m_RepaymentDate = Now
    m_ValidityPeriod = 6
    m_ChangeSection = False
End Sub

Public Function Find() As ResultSet
    Dim ApiPath As String
    ApiPath = "/v1/json/course/repayment"
    Dim Query As Dictionary
    Set Query = BuildQuery()
    Dim Result As Dictionary
    Set Result = m_Client.Fetch(ApiPath, Query)
    
    Dim dicRepaymentList As Dictionary
    Dim Response As ResultSet
    
    If m_Client.Success = False Then
        Response.Success = False
        Response.Error = m_Client.Error
        Find = Response
        Exit Function
    End If
    Response.RepaymentList = m_Client.CreateRepaymentList(Result("RepaymentList"))
    Response.TeikiRoute = m_Client.CreateTeikiRoute(Result("TeikiRoute"))
    
    Response.Success = True
    Find = Response

End Function

Property Let client(ByVal client As Ekispert)
    Set m_Client = client
End Property

Property Let Separator(i As Long, ByVal value As String)
    ReDim Preserve m_Separator(i) As String
    m_Separator(i) = value
End Property

Property Let SerializeData(ByVal value As String)
    m_SerializeData = value
End Property

Property Let CheckEngineVersion(ByVal value As Boolean)
    m_CheckEngineVersion = value
End Property

Property Let ChangeSection(ByVal value As Boolean)
    m_ChangeSection = value
End Property

Property Let StartDate(ByVal value As Date)
    m_StartDate = value
End Property

Property Let BuyDate(ByVal value As Date)
    m_BuyDate = value
End Property

Property Let RepaymentDate(ByVal value As Date)
    m_RepaymentDate = value
End Property

Property Let ValidityPeriod(ByVal value As Date)
    m_ValidityPeriod = value
End Property

Function BuildQuery() As Dictionary
    Set BuildQuery = New Dictionary
    BuildQuery.Add "serializeData", m_SerializeData
    BuildQuery.Add "checkEngineVersion", IIf(m_CheckEngineVersion, "true", "false")
    BuildQuery.Add "changeSection", IIf(m_ChangeSection, "true", "false")
    BuildQuery.Add "startDate", Format(m_StartDate, "yyyymmdd")
    BuildQuery.Add "repaymentDate", Format(m_RepaymentDate, "yyyymmdd")
    If m_BuyDate > 0 Then
        BuildQuery.Add "buyDate", Format(m_BuyDate, "yyyymmdd")
    End If
    BuildQuery.Add "validityPeriod", m_ValidityPeriod
    m_Client.ArrayToString BuildQuery, "separator", m_Separator
End Function




