Attribute VB_Name = "AquaplotAPIDemo"

Private Function AquaplotBaseUrl() As String
    Dim aquaplot_api_domain As String
    Dim aquaplot_api_version As String
    Dim protocol As String
    
    protocol = "https"
    aquaplot_api_domain = "api.aquaplot.com"
    aquaplot_api_version = "v1"
    AquaplotBaseUrl = protocol + "://" + aquaplot_api_domain + "/" + aquaplot_api_version + "/"
End Function

Private Function EncodeBase64(ByVal text As String) As String
  'method taken from http://stackoverflow.com/a/169945
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Set objXML = CreateObject("MSXML2.DOMDocument")
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

Private Function SanitizeUrl(ByVal url As String) As String
    SanitizeUrl = url + IIf(Right(url, 1) <> "/", "/", "")
End Function

Private Function requestUrl(ByVal baseUrl As String, ByVal params As Variant) As String
    'join basUrl with params to form complete url
    Dim paramString As String
    paramString = Join(params, "&")
    requestUrl = SanitizeUrl(baseUrl) + "?" + paramString
End Function


Private Function SynchronousHttpGetUsingBasicAuth(ByVal url As String, ByVal user As String, ByVal pass As String) As String
    'Base64 encoding for authentication
    Dim encodedSecret As String
    encodedSecret = EncodeBase64(user + ":" + pass)
     
    'get request handler
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP")

    httpRequest.Open "GET", url, False  ' <-- "Asynchronous=False is set so the program waits until the request is fully processed (=response from server arrived)
    
    'set request header
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.setRequestHeader "Accept", "application/json"
    httpRequest.setRequestHeader "Authorization", ("Basic " + encodedSecret) '<-- authentication
    
    httpRequest.send
    
    SynchronousHttpGetUsingBasicAuth = httpRequest.responseText

    Set httpRequest = Nothing
End Function

Private Function AquaplotRequestUrl(ByVal reqType As String, ByVal fromLoc As Variant, ByVal toLoc As Variant, ByVal params As Variant) As String
    Dim baseUrl As String
    baseUrl = AquaplotBaseUrl() _
            + reqType + "/" _
            + "from/" + CStr(fromLoc(0)) + "/" + CStr(fromLoc(1)) + "/" _
            + "to/" + CStr(toLoc(0)) + "/" + CStr(toLoc(1)) + "/"
    AquaplotRequestUrl = requestUrl(baseUrl, params)
End Function


Function ShowAquaplotDistanceRequestResponse(ByRef user As Variant, ByRef pass As Variant, ByRef params As Range, ByRef fromLoc As Range, ByRef toLoc As Range) As String
    Dim paramsArray()
    ReDim paramsArray(params.Cells.Count)
    
    For i = 1 To params.Cells.Count
        paramsArray(i - 1) = params(i).Value
    Next i
    
    Dim passStr As String
    userStr = user.Value
    passStr = pass.Value
    
    Dim fromLocArray(2) As Double, toLocArray(2) As Double
    
    For i = 1 To 2
        fromLocArray(i - 1) = fromLoc(i).Value
        toLocArray(i - 1) = toLoc(i).Value
    Next i
    
    Dim requestUrl As String
    requestUrl = AquaplotRequestUrl("distance", fromLocArray, toLocArray, paramsArray)
    ShowAquaplotDistanceRequestResponse = SynchronousHttpGetUsingBasicAuth(requestUrl, userStr, passStr)
End Function
