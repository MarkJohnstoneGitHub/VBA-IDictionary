VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("VBA-IScriptingDictionary.Tests.Unit Testing")
Option Explicit

' --------------------------------------------- '
' Private Variables
' --------------------------------------------- '
Private Type TCustomer
    pvtName As String
    pvtDateOfBirth As Date
    pvtPostcode As Long
End Type

Private this As TCustomer

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public Property Get Customer() As Customer
    Set Customer = Me
End Property


Public Property Get Name() As String
    Name = this.pvtName
End Property
Public Property Let Name(ByVal customerName As String)
    this.pvtName = customerName
End Property

Public Property Get DateOfBirth() As Date
    DateOfBirth = this.pvtDateOfBirth
End Property
Public Property Let DateOfBirth(ByVal customerDateOfBirth As Date)
    this.pvtDateOfBirth = customerDateOfBirth
End Property

Public Property Get Postcode() As Long
    Postcode = this.pvtPostcode
End Property
Public Property Let Postcode(ByVal customerPostcode As Long)
    this.pvtPostcode = customerPostcode
End Property

Public Property Get Self() As Customer
    Set Self = Me
End Property

Public Function Create(ByVal customerName As String, _
                    ByVal customerDateOfBirth As Date, _
                    ByVal customerPostcode As Long) _
                    As Customer
    With New Customer
        .Name = customerName
        .DateOfBirth = customerDateOfBirth
        .Postcode = customerPostcode
        Set Create = .Self
    End With
End Function


Public Function Instance() As Customer
    With New Customer
        Set Instance = .Self
    End With
End Function

Public Function IsTypeOf(ByRef dataValue As Variant) As Boolean
On Error GoTo ErrHandler
    Dim pvtIsTypeOf As Boolean
      
    pvtIsTypeOf = False
    
    If TypeOf dataValue Is Customer Then
        pvtIsTypeOf = True
    End If
CleanExit:
    IsTypeOf = pvtIsTypeOf
Exit Function
ErrHandler:
    pvtIsTypeOf = False
    Resume CleanExit
End Function


Public Function IsTypeSafe(ByRef dataValue As Variant) As Boolean
On Error GoTo ErrHandler
    Dim pvtIsTypeSafe As Boolean
      
    pvtIsTypeSafe = False
    
    If TypeOf dataValue Is Customer Then
        pvtIsTypeSafe = True
    End If
CleanExit:
    IsTypeSafe = pvtIsTypeSafe
Exit Function
ErrHandler:
    pvtIsTypeSafe = False
    Resume CleanExit
End Function


