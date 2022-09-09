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
    .BaseUrl = "https://ec.europa.eu/taxation_customs/vies/rest-api/"
    .FollowRedirects = False
    .Insecure = False
  End With
  
  Set checkVatReq = New WebRequest
  With checkVatReq
    .Method = HttpGet
    .Format = Json
    .Resource = "ms/{country_code}/vat/{vat_number}"
  End With
End Sub

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
  
  checkVatReq.AddUrlSegment "country_code", country
  checkVatReq.AddUrlSegment "vat_number", vatNum
  
  Dim resp As New WebResponse
  Set resp = webCli.Execute(checkVatReq)
  
  If resp.StatusCode = Ok Then
    Dim respData As New Dictionary
    Set respData = resp.Data
    CheckVat.IsValid = respData("isValid")
    CheckVat.ValidationDate = respData("requestDate")
    CheckVat.CompanyName = respData("name")
    CheckVat.Address = respData("address")
    
    If CheckVat.IsValid Then
      CheckVat.ValidationMessage = "VAT number is valid"
    Else
      CheckVat.ValidationMessage = StrConv(respData("userError"), vbProperCase)
    End If
  
  Else
    CheckVat.ValidationMessage = resp.StatusDescription
  End If
  
End Function

