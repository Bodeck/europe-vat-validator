Attribute VB_Name = "VatValidator"
Option Explicit
'Version: 0.0.1
Public Type VatValidationInfo
  IsValid As Boolean
  CompanyName As String
  Address As String
  ValidationMessage As String
  ValidationDate As String
End Type

Sub DisplayUserForm()
  UserForm.Show
End Sub

Sub RunValidator(Optional ByVal useProxy As Boolean, Optional ByVal proxyUser As String, Optional ByVal proxyPasword As String)
  
  Dim EuVat As New EuVatAPIClass
  Dim UkVat As New UkVatAPIClass
  
  If useProxy Then
    EuVat.SetUpProxy proxyUser, proxyPasword
    UkVat.SetUpProxy proxyUser, proxyPasword
  End If
  
  Dim VatNumList As Range
  Set VatNumList = ThisWorkbook.Worksheets("Validator").UsedRange
  Dim vatRow As Range
  For Each vatRow In VatNumList.Rows
    If vatRow.Row > 1 Then
      Dim country As String
      country = Left(vatRow.Cells(1, 1).Value, 2)
      Dim vatNum As String
      vatNum = Mid(vatRow.Cells(1, 1), 3)
      Dim vatInfo As VatValidationInfo
      
      If country = "GB" Then
        vatInfo = UkVat.CheckVat(vatNum)
      Else
        vatInfo = EuVat.CheckVat(country, vatNum)
      End If
      'write data
      vatRow.Cells(1, 2) = vatInfo.IsValid
      vatRow.Cells(1, 3) = vatInfo.ValidationMessage
      vatRow.Cells(1, 4) = vatInfo.CompanyName
      vatRow.Cells(1, 5) = vatInfo.Address
      vatRow.Cells(1, 6) = vatInfo.ValidationDate
    End If
  Next vatRow
  MsgBox "Validation completed!"
  Set EuVat = Nothing
  Set UkVat = Nothing
End Sub

