VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UkVatAPIClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private webCli As WebClient
Private checkVatReq As WebRequest

Private Sub Class_Initialize()
  Set webCli = New WebClient
  With webCli
    .BaseUrl = "https://api.service.hmrc.gov.uk//organisations/vat/check-vat-number/lookup/"
    .Insecure = False
    .FollowRedirects = False
  End With
  Set checkVatReq = New WebRequest
  
  With checkVatReq
    .Resource = "{vatNum}"
    .Format = Json
    .Method = HttpGet
  End With

End Sub

Private Sub Class_Terminate()
  Set webCli = Nothing
  Set checkVatReq = Nothing
End Sub

Public Sub SetUpProxy(ByVal username As String, ByVal password As String)
  webCli.EnableAutoProxy = True
  webCli.ProxyUsername = username
  webCli.ProxyPassword = password
End Sub

Public Function CheckVat(ByVal vatNum As String) As VatValidationInfo

  checkVatReq.AddUrlSegment "vatNum", vatNum
  
  Dim resp As New WebResponse
  Set resp = webCli.Execute(checkVatReq)
  
  If resp.StatusCode = Ok Then
    Dim respData As New Dictionary
    Set respData = resp.Data
      CheckVat.IsValid = True
      CheckVat.CompanyName = respData("target")("name")
      CheckVat.Address = respData("target")("address")("line1") _
        & vbNewLine & respData("target")("address")("line2") _
        & vbNewLine & respData("target")("address")("line3") _
        & vbNewLine & respData("target")("address")("line4") _
        & vbNewLine & respData("target")("address")("postcode") _
        & " " & respData("target")("address")("countryCode")
      CheckVat.ValidationDate = respData("processingDate")
      CheckVat.ValidationMessage = "VAT number is valid"
  ElseIf resp.StatusCode = BadRequest Or resp.StatusCode = NotFound Then
    CheckVat.ValidationMessage = resp.Data("message")
  Else
    CheckVat.ValidationMessage = resp.StatusDescription
  End If
End Function
