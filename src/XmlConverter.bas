Attribute VB_Name = "XmlConverter"
Option Explicit

Public Function ParseXml(xmlString As String) As Object
    Set ParseXml = CreateObject("MSXML2.DOMDocument")
    ParseXml.Async = False
    ParseXml.LoadXML xmlString
End Function

Public Function ConvertToXml(XmlDomDoc As Variant) As String
    ConvertToXml = Trim(Replace(XmlDomDoc.XML, vbCrLf, ""))
End Function
