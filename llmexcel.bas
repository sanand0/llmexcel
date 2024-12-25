'''
' LLM Excel - Talk to LLMs (like ChatGPT) in Excel

' @author root.node@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''
Function LLM(prompt As String, Optional ByVal model As String = "gpt-4o-mini", Optional ByVal refresh As Boolean = False, Optional ByVal temperature As Single = 1) As Variant
    ' Don't re-calculate unless the prompt changes
    Application.Volatile False

    ' Set up request
    Dim jsonRequest As String
    jsonRequest = "{""response_format"":{""type"":""json_object""},""model"":""" + model + """"
    jsonRequest = jsonRequest + ",""messages"":[{""role"":""system"",""content"":""Return only 1 JSON array""},{""role"":""user"",""content"":""" + Replace(prompt, """", """""") + """}]"
    If temperature <> 1 Then
        jsonRequest = jsonRequest + ",""temperature"":" + ConvertToJson(temperature)
    End If
    jsonRequest = jsonRequest + "}"

    ' Get the API Key from OPENAI_API_KEY environment variable
    Dim apiKey As String
    If Environ("OPENAI_API_KEY") <> "" Then
        apiKey = Environ("OPENAI_API_KEY")
    Else
        LLM = "#ERROR Missing environment variable OPENAI_API_KEY"
        Exit Function
    End If

    ' Send the HTTP request
    Dim http As Object
    Set http = CreateObject("Msxml2.XMLHTTP.6.0")
    ' Do not set timeouts. XMLHTTP doesn't support it. ServerXMLHTTP doesn't work
    ' http.setTimeouts 5000, 5000, 5000, 30000
    http.Open "POST", "https://api.openai.com/v1/chat/completions", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    If Not refresh Then
        http.setRequestHeader "Cache-Control", "max-age=86400"
    End If
    http.Send jsonRequest

    ' Parse the JSON response and return the content
    Dim responseJSON As Object
    On Error GoTo CannotParseResponseAsJSON
    Set responseJSON = JsonConverter.ParseJson(http.responseText)

    ' If error, show error and exit
    If Not IsEmpty(responseJSON("error")) Then
        LLM = "#ERROR: " & responseJSON("error")("message")
        Exit Function
    End If

    Dim content As String
    On Error GoTo CannotGetContent
    content = responseJSON("choices")(1)("message")("content")

    Dim answerJSON As Variant
    On Error GoTo CannotParseContentAsJSON
    Set answerJSON = JsonConverter.ParseJson(content)

    Dim answerType As String
    answerType = TypeName(answerJSON)
    On Error GoTo CannotExtractResult
    If answerType = "Dictionary" Or answerType = "Collection" Then
        ' If answerJSON is an object, get the values
        Dim result() As String
        Dim i As Integer
        Dim item As Variant
        ReDim result(answerJSON.Count - 1)
        i = 0
        For Each item In answerJSON
            ' ParseJson returns a Dictionary yields an empty key at the end. Handle that
            If IsEmpty(item) Then Exit For
            ' Reconvert to JSON if value is not a string
            If answerType = "Dictionary" Then
                If TypeName(answerJSON(item)) = "String" Then
                    result(i) = answerJSON(item)
                Else
                    result(i) = ConvertToJson(answerJSON(item))
                End If
            Else
                If TypeName(item) = "String" Then
                    result(i) = item
                Else
                    result(i) = ConvertToJson(item)
                End If
            End If
            i = i + 1
        Next item
        LLM = result
    Else
        LLM = "#ERROR: Not JSON: " & content
    End If
    Exit Function


CannotParseResponseAsJSON:
    LLM = "#ERROR: Cannot parse response as JSON: " & http.responseText
    Exit Function

CannotGetContent:
    LLM = "#ERROR: Cannot get choices[0].message.content: " & responseJSON
    Exit Function

CannotParseContentAsJSON:
    LLM = "#ERROR: Cannot parse content as JSON: " & content
    Exit Function

CannotExtractResult:
    LLM = "#ERROR: Cannot extract result from: " & content

End Function
