VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EuVatAPIClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private webCli As WebClient
Private checkVatReq As WebRequest
Private Const supportedCountries As String = "AT|BE|BG|CY|CZ|DE|DK|EE|EL|ES|FI|FR|HR|HU|IE|IT|LT|LU|LV|MT|NL|PL|PT|RO|SE|SI|SK|XI"

Private Sub Class_Initialize()
  Set webCli = New WebClient
  With webCli
    .BaseUrl = "https://ec.europa.eu/taxation_customs/vies/services/checkVatService"
    .FollowRedirects = False
    .Insecure = False
  End With
  'VERSION 4.O of VBAWeb doesn't support XML so we need register our own converter and parser for xml
  WebHelpers.RegisterConverter "soap", "application/xml", "XmlConverter.ConvertToXml", "XmlConverter.ParseXml"
  Set checkVatReq = New WebRequest
  checkVatReq.Method = HttpPost
  checkVatReq.CustomRequestFormat = "soap"
  checkVatReq.CustomResponseFormat = "soap"
End Sub

Private Function SetXMLBody(ByVal country As String, ByVal vatNum As String) As DOMDocument
  Dim xmlString As String
  xmlString = "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" _
          & "<soap:Body>" _
          & "   <checkVat xmlns=""urn:ec.europa.eu:taxud:vies:services:checkVat:types"">" _
          & "      <countryCode>" & country & "</countryCode> " _
          & "     <vatNumber>" & vatNum & "</vatNumber>" _
          & "   </checkVat>" _
          & " </soap:Body>" _
          & "</soap:Envelope>"
  Set SetXMLBody = New DOMDocument
  
  With SetXMLBody
    .Async = False
    .LoadXML xmlString
  End With

End Function

Private Function IsValidCountry(ByVal countryCode As String) As Boolean
  IsValidCountry = (InStr(1, supportedCountries, countryCode))
End Function

Public Sub SetUpProxy(ByVal username As String, ByVal password As String)
  webCli.EnableAutoProxy = True
  webCli.ProxyUsername = username
  webCli.ProxyPassword = password
End Sub

Public Function CheckVat(ByVal country As String, ByVal vatNum As String) As VatValidationInfo
  
  If Not IsValidCountry(country) Then
    CheckVat.ValidationMessage = """" & country & """ is not valid country or not supported by VIES service"
    Exit Function
  End If
  
  Dim soapEnv As String
  Dim resp As WebResponse
  Dim respData As MSXML2.DOMDocument
  Set checkVatReq.Body = SetXMLBody(country, vatNum)
  Set resp = webCli.Execute(checkVatReq)
  
  If resp.StatusCode = Ok Then
    Set respData = resp.Data
    CheckVat.IsValid = CBool(respData.getElementsByTagName("valid")(0).Text)
    CheckVat.Address = respData.getElementsByTagName("address")(0).Text
    CheckVat.CompanyName = respData.getElementsByTagName("name")(0).Text
    CheckVat.ValidationDate = respData.getElementsByTagName("requestDate")(0).Text
    CheckVat.ValidationMessage = IIf(CheckVat.IsValid, "VAT number is valid", "Invalid VAT number")
  Else
    CheckVat.IsValid = False
    CheckVat.ValidationMessage = "Web request returned error: " & resp.StatusDescription
  End If
End Function

