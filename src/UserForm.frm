VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Europe VAT Number Validator"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
  Me.Hide
End Sub

Private Sub btnStart_Click()
  If chkEnableProxy.Value Then
    Me.Hide
    RunValidator True, Me.txtProxyUser.Value, Me.txtProxyPassword.Value
  Else
    Me.Hide
    RunValidator
  End If
  
End Sub

Private Sub chkEnableProxy_Click()
  frmProxyDetails.Visible = chkEnableProxy.Value
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Debug.Print CloseMode
  If CloseMode = 0 Then
    Cancel = vbFalse
    Me.Hide
  End If
End Sub
