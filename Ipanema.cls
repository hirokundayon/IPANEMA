VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'--------------------------------------
'JSON���p�[�X�����I�u�W�F�N�g�𐶐�
'--------------------------------------
Public Function ParseJson(JsonText As String, Key As String) As String
    Dim jsonObj As Object

    ' Parse json to Dictionary
    ' "values" is parsed as Collection
    ' each item in "values" is parsed as Dictionary
    Set jsonObj = JsonConverter.ParseJson(JsonText)
    ParseJson = CStr(jsonObj(Key))
End Function

' Selenium Standalone Server �N��
Function runSeleniumStandaloneServer() As Integer
    Dim pid As Integer
    pid = _
      Shell( _
        "java -Dwebdriver.chrome.driver=C:\Selenium\bin\chromedriver.exe -Dwebdriver.gecko.driver=C:\Selenium\bin\geckodriver.exe -Dwebdriver.edge.driver=C:\Selenium\bin\MicrosoftWebDriver.exe -Dwebdriver.ie.driver=C:\Selenium\bin\IEDriverServer.exe -jar C:\Selenium\bin\selenium-server-standalone-3.8.1.jar" _
    )
    'pid = _
    '  Shell( _
    '         "C:\Selenium\bin\chromedriver.exe" _
    ')
    runSeleniumStandaloneServer = pid
End Function

' Selenium Standalone Server ��~
Sub stopSeleniumStandaloneServer(ByVal pid As Integer)
    Dim strComputer As String
    Dim objWMIService
    Dim colProcessList
    Dim objProcess
    
    strComputer = "."
    Set objWMIService = getObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where ProcessID = " & CStr(pid))
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
End Sub

' JSON��POST
Function postJSON(ByVal url As Variant, ByVal Json As String) As String
    postJSON = useAPI(url, "POST", Json)
End Function

' API��@��
Function useAPI(ByVal url As Variant, ByVal method As String, ByVal Json As String) As String
    Dim objHTTP As Object

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    objHTTP.Open method, url, False
    objHTTP.setRequestHeader "Content-type", "application/json"
    objHTTP.send Json
    
    useAPI = objHTTP.responseText
    
    Set objHTTP = Nothing
End Function

Function newSession() As String
    Dim url As String
    Dim response As String
    Dim result As Object
    Dim Json As String
    
    url = "http://localhost:4444/wd/hub/session"
    'url = "http://localhost:9515/"
    
    'Json = "{""desiredCapabilities"":{}, ""requiredCapabilities"":{}}"
    Json = "{""desiredCapabilities"":{""browserName"":""chrome""},""requiredCapabilities"":{}}"
    'Json = "{""desiredCapabilities"":{""browserName"":""firefox""},""requiredCapabilities"":{}}"
    'Json = "{""desiredCapabilities"":{""browserName"":""internet explorer""},""requiredCapabilities"":{}}"
    'Json = "{""desiredCapabilities"":{""browserName"":""MicrosoftEdge""},""requiredCapabilities"":{}}"
    
    response = postJSON(url, Json)
    
    newSession = ParseJson(response, "sessionId")
End Function

Sub deleteSession(ByVal sessionId As String)
    Dim objHTTP As Object
    Dim Json As String

    Dim url As Variant
    Dim result As Object
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    url = "http://localhost:4444/wd/hub/session/" & sessionId
    
    objHTTP.Open "DELETE", url, False
    objHTTP.send
    Set objHTTP = Nothing
End Sub

Sub goURL(ByVal sessionId As String, ByVal url As String)
    Dim Json As String

    Dim sessionURL As Variant
    Dim result As Object
    
    sessionURL = "http://localhost:4444/wd/hub/session/" & sessionId & "/url"
    
    Json = "{""url"":""" & url & """}"
    
    postJSON sessionURL, Json
End Sub

Function getTitle(ByVal sessionId As String) As String
    Dim objHTTP As Object

    Dim url As Variant
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    url = "http://localhost:4444/wd/hub/session/" & sessionId & "/title"
    
    getTitle = ParseJson(useAPI(url, "GET", ""), "value")

End Function

Sub waitByTitle(ByVal sessionId As String, ByVal targetTitle As String)
    Dim title As String
    
    Do
        title = getTitle(sessionId)
    Loop Until title = targetTitle
End Sub

' ���i�擾
Function getElement(ByVal sessionId As String, ByVal selector As String, ByVal locator As String) As String
    Dim objHTTP As Object
    
    Dim Json As String
    Dim url As String
    Dim result As Integer
    Dim element As Object
    Dim text As String
    Dim jsonObj As Object
        
    Json = "{""using"":""" & selector & """" & ",""value"":""" & locator & """}"
    url = "http://localhost:4444/wd/hub/session/" & sessionId & "/element"
    
    Do
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
        objHTTP.Open "POST", url, False

        objHTTP.setRequestHeader "Content-type", "application/json"
        objHTTP.send (Json)
    
        text = objHTTP.responseText
        result = objHTTP.status
        
        Set objHTTP = Nothing
    Loop Until result = 200
    
    Set jsonObj = JsonConverter.ParseJson(text)
    getElement = CStr(jsonObj("value")("ELEMENT"))
End Function

' �L�[����
Sub sendKeys(ByVal sessionId As String, ByVal elementId As String, ByVal text As String)
    Dim objHTTP As Object
    
    Dim Json As String
    Dim url As String
    
    Json = "{""value"":[""" & text & """]}"
    url = "http://localhost:4444/wd/hub/session/" & sessionId & "/element/" & elementId & "/value"

    postJSON url, Json

End Sub

' �N���b�N
Sub click(ByVal sessionId As String, ByVal elementId As String)
    Dim objHTTP As Object
    
    Dim url As String
    
    url = "http://localhost:4444/wd/hub/session/" & sessionId & "/element/" & elementId & "/click"
    postJSON url, ""

End Sub

Sub ipanema()
    Dim sessionId As String
    Dim url As String
    Dim pid As Integer
    Dim title As String
    Dim selector As String
    Dim locator As String
    Dim elementId As String
    Dim text As String
        
    ' Selenium Standalone Server �N��
    pid = runSeleniumStandaloneServer()
    
    ' 1�b�҂�
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Web�u���E�U�N��
    sessionId = newSession()
    
    ' Google�֔��
    url = "http://www.google.co.jp/"
    Call goURL(sessionId, url)
    
    title = "Google"
    Call waitByTitle(sessionId, title)
    
    ' "OSC ���É� 2012 ���|�[�g"������
    ' �������擾
    selector = "name"
    locator = "q"
    elementId = getElement(sessionId, selector, locator)
    
    ' "OSC ���É� 2012 ���|�[�g"�����
    text = "OSC ���É� 2012 ���|�[�g\n"
    sendKeys sessionId, elementId, text
    
    title = "OSC ���É� 2012 ���|�[�g - Google ����"
    Call waitByTitle(sessionId, title)
    
    ' OSC2012 Nagoya �̃��|�[�g��\��������
    selector = "partial link text"
    locator = "OSC2012 Nagoya�ւ����ꂠ�肪�Ƃ��������܂����I�I"
    elementId = getElement(sessionId, selector, locator)
    
    click sessionId, elementId
    title = "OSC2012 Nagoya�ւ����ꂠ�肪�Ƃ��������܂����I�I"
    Call waitByTitle(sessionId, title)
    
    ' �ʐ^���N���b�N
    selector = "css selector"
    locator = "img[title='�������[�J�[�h�G���[�I']"
    elementId = getElement(sessionId, selector, locator)
    Call click(sessionId, elementId)
    
    title = "�������[�J�[�h�G���[�I"
    Call waitByTitle(sessionId, title)
    
    ' �X�Ɏʐ^���N���b�N
    locator = "p[class='attachment'] > a"
    elementId = getElement(sessionId, selector, locator)
    Call click(sessionId, elementId)
    
    ' 10�b�҂�
    Application.Wait Now + TimeValue("00:00:10")
    
    ' �u���E�U��~
    Call deleteSession(sessionId)
    
    ' 1�b�҂�
    'Application.Wait Now + TimeValue("00:00:10")
    
    ' Selenium Standalone Server ��~
    Call stopSeleniumStandaloneServer(pid)
End Sub