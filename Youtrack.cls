VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Youtrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ytUrl As String
Private ytUser As String
Private ytToken As String

Public Sub Connect(url As String, user As String, password As String, authorizationToken As String, scope As String)
    Dim hubUrl As String
    ytUrl = url
    ytUser = user
    hubUrl = url + "/hub/api/rest/oauth2/token"
    Dim HTTPReq As Object
    Set HTTPReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTPReq.Open "POST", hubUrl, False
    HTTPReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    HTTPReq.setRequestHeader "Authorization", "Basic " + authorizationToken
    HTTPReq.send ("grant_type=password&username=" + user + "&password=" + password + "&scope=" + scope)
    Dim ResponseJson As Object
    Set ResponseJson = JsonConverter.ParseJson(HTTPReq.responseText)
    ytToken = ResponseJson("access_token")
End Sub

Public Function GetIssues(ByVal searchString As String) As IXMLDOMNodeList
    Dim issuesUrl As String
    Dim issuesXml As MSXML2.DOMDocument60
    issuesUrl = ytUrl + "/youtrack/rest/issue/?filter=" + URLEncode(searchString) + "&max=1000"
    Set issuesXml = ResponseXml(CallYoutrackRest(issuesUrl, "GET"))
    Dim issues As IXMLDOMNode
    Set issues = issuesXml.SelectSingleNode("//issueCompacts")
    Set GetIssues = issues.ChildNodes
End Function

Public Function GetSearches() As IXMLDOMNodeList
    Dim searchesUrl As String
    Dim searchesXml As MSXML2.DOMDocument60
    searchesUrl = ytUrl + "/youtrack/rest/user/" + ytUser + "/filter"
    Set searchesXml = ResponseXml(CallYoutrackRest(searchesUrl, "GET"))
    Dim searches As IXMLDOMNode
    Set searches = searchesXml.SelectSingleNode("//queries")
    Set GetSearches = searches.ChildNodes
End Function

Public Function GetWorkItems(ByVal issue As String) As IXMLDOMNodeList
    Dim workItemsUrl As String
    Dim workItemsXml As MSXML2.DOMDocument60
    workItemsUrl = ytUrl + "/youtrack/rest/issue/" + issue + "/timetracking/workitem/"
    Set workItemsXml = ResponseXml(CallYoutrackRest(searchesUrl, "GET"))
    Dim workItems As IXMLDOMNode
    Set workItems = workItemsXml.SelectSingleNode("//workItems")
    Set GetWorkItems = workItems.ChildNodes
End Function

Private Function ResponseXml(ByVal response As String) As MSXML2.DOMDocument60
    Dim oXml As MSXML2.DOMDocument60
    Set oXml = New MSXML2.DOMDocument60
    oXml.LoadXML response
    Set ResponseXml = oXml
End Function

Private Function CallYoutrackRest(ByVal url As String, ByVal requestType As String) As String
    Dim HTTPReq As Object
    Set HTTPReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTPReq.Open requestType, url, False
    HTTPReq.setRequestHeader "Authorization", "Bearer " + ytToken
    HTTPReq.send
    CallYoutrackRest = HTTPReq.responseText
End Function

Private Function URLEncode(ByVal str As String) As String
        Dim intLen As Integer
        Dim x As Integer
        Dim curChar As Long
        Dim newStr As String
        intLen = Len(str)
        newStr = ""

        For x = 1 To intLen
            curChar = Asc(Mid$(str, x, 1))

            If (curChar < 48 Or curChar > 57) And _
                (curChar < 65 Or curChar > 90) And _
                (curChar < 97 Or curChar > 122) Then
                                newStr = newStr & "%" & Hex(curChar)
            Else
                newStr = newStr & Chr(curChar)
            End If
        Next x

        URLEncode = newStr
End Function


