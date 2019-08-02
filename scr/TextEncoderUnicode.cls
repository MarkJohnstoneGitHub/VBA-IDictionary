VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextEncoderUnicode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Unicode text encoder that converts a string to a unicode hex string using a managed variant read-only integer array overlay. \r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''
'Rubberduck annotations
'@Folder("VBA-IScriptingDictionary.TextEncoding")
'@ModuleDescription "A Unicode text encoder that converts a string to a unicode hex string using a managed variant read-only integer array overlay. \r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''

''
'@Version VBA-IScriptingDictionary v2.0 (July 28, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified July 28, 2019

'@Description Converts a string to a unicode hex string using a read-only integer array overlay.

'@Usage   Eg. Dim textEncoder As UnicodeTextEncoder
'             Set textEncoder = New UnicodeTextEncoder
'             Debug.Print textEncoder.ToHexString("ABCabc")
'             Set textEncoder = Nothing
'@Remarks The
' Termination or resetting of the application while running the ToHexString
' function may cause the application to crash, if the managed variant
' hasn't been disposed of.
' Under normal circumstances the managed variant is disposed of
' in the Class_Terminate,when this object goes out of scope or set to nothing.
' Untested on Mac and VBA 6.
'------------------------------------------------------------'

Option Explicit
Implements ITextEncoding

'============================================='
'API Declarations
'============================================='
''
'@API_Declaration
'API's declarations for copy memory by pointer for Windows and Mac, with VBA6 and VBA7 compatibility.
'See https://stackoverflow.com/questions/45756170/how-to-read-write-memory-on-mac-os-x-with-vba
'------------------------------------------------------------'
#If Mac Then
  #If Win64 Then
    Private Declare PtrSafe Function CopyMemory_byPtr Lib "libc.dylib" Alias "memmove" _
            (ByVal dest As LongPtr, _
             ByVal src As LongPtr, _
             ByVal size As Long) _
             As LongPtr
  #Else
    Private Declare Function CopyMemory_byPtr Lib "libc.dylib" Alias "memmove" _
            (ByVal dest As Long, _
             ByVal src As Long, _
             ByVal size As Long) _
             As Long
  #End If
#ElseIf VBA7 Then
  #If Win64 Then
    Private Declare PtrSafe Sub CopyMemory_byPtr Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal dest As LongPtr, _
             ByVal src As LongPtr, _
             ByVal size As LongLong)
  #Else
    Private Declare PtrSafe Sub CopyMemory_byPtr Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal dest As LongPtr, _
             ByVal src As LongPtr, _
             ByVal size As Long)
  #End If
#Else
  Private Declare Sub CopyMemory_byPtr Lib "kernel32" Alias "RtlMoveMemory" _
          (ByVal dest As Long, _
           ByVal src As Long, _
           ByVal size As Long)
#End If

''
'@API_Declaration
'API's declarations for fill memory.
'------------------------------------------------------------'
#If VBA7 Then
  Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" _
          (ByRef dest As Any, _
           ByVal size As Long, _
           ByVal fill As Byte)
#Else
  Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" _
          (ByRef dest As Any, _
           ByVal size As Long, _
           ByVal fill As Byte)
#End If

'============================================='
'Constants
'============================================='
Private Const INTEGER_BYTESIZE As Long = 2

'============================================='
'Types
'============================================='
Private Type SAFEARRAYBOUND
    cElements    As Long 'The number of elements in the dimension.
    lLbound      As Long
End Type

#If VBA7 Then
    Private Type TSafeArrayReadOnly
        cDims As Integer            ' nr of dimensions for the array
        fFeatures As Integer        ' extra information about the array contents
        cbElements As Long          ' nr of bytes per array element. Possible Examples: 1=byte,2=integer,4=long,8=currency
        cLocks As Long              ' nr of times array was locked w/o being unlocked
        pvData As LongPtr              ' address to 1st array item, can be a pointer to another structure/address
        rgsabound(0) As SAFEARRAYBOUND ' set of SafeArrayBound structures for each dimension. i.e. One dimension
    End Type
#Else
    Private Type TSafeArrayReadOnly
        cDims As Integer            ' nr of dimensions for the array
        fFeatures As Integer        ' extra information about the array contents
        cbElements As Long          ' nr of bytes per array element. Possible Examples: 1=byte,2=integer,4=long,8=currency
        cLocks As Long              ' nr of times array was locked w/o being unlocked
        pvData As Long              ' address to 1st array item, can be a pointer to another structure/address
        rgsabound(0) As SAFEARRAYBOUND ' set of SafeArrayBound structures for each dimension. i.e. One dimension
    End Type
#End If

#If VBA7 Then
    Private Type tagVARIANT
        vt As Integer               'Variable Type
        wReserved1 As Integer
        wReserved2 As Integer
        wReserved3 As Integer
        pSomething As LongPtr       'Pointer to string or array data
    End Type
#Else
    Private Type tagVARIANT
        vt As Integer               'Variable Type
        wReserved1 As Integer
        wReserved2 As Integer
        wReserved3 As Integer
        pSomething As Long          'Pointer to string or array data
    End Type
#End If

Private Type TUnicodeTextEncoder
    integerArrayOverlay As TSafeArrayReadOnly
    unmanagedVariant As tagVARIANT   'unmanaged variant
    managedVariant As Variant        'manaaged variant
End Type

'============================================='
'Private Variables
'============================================='
Private this As TUnicodeTextEncoder

'============================================='
'Constructors and destructors
'============================================='
Private Sub Class_Initialize()
    InitializeManagedVariant
End Sub

Private Sub Class_Terminate()
    EraseSafeArray
    DisposeManagedVariant  'destroy managed variant
End Sub

'============================================='
'Public Methods
'============================================='
''
'@Description("Converts a string to unicode hex string.")
'@param inText The unicode string to be converted to a unicode hex string.
'@Remarks Only for strings. For performance validation of input
' being a string is the responsibility of the calling module.
'
' Interuption or reset while running this function may cause
' the application to crash if the managed variant isn't disposed.
' Th
'------------------------------------------------------------'
Public Function ToHexString(ByRef inText As String) As String
Attribute ToHexString.VB_Description = "Converts a string to unicode hex string."
    Const HEX_STRING_LENGTH As Long = 4
    Const HEX_STRING_PADDING As Long = 65536 'Ensures that four hex characters are obtained with 0 padding
    
    'SafeArrayRedim inText 'Moved ToHexString for performance
    With this.integerArrayOverlay
        .pvData = StrPtr(inText)                    'Allocate Data, the array data points to the string pointer
        .rgsabound(0).cElements = LenB(inText) / INTEGER_BYTESIZE 'i.e. 2 bytes for an integer
    End With
    
    'Allocate the ToHexString buffer according to specified string i.e. each character is four hex characters
    ToHexString = VBA.Space$(HEX_STRING_LENGTH * (UBound(this.managedVariant) - LBound(this.managedVariant) + 1))
    
    Dim hexStrIndex As Long
    hexStrIndex = 1
    Dim char As Variant
    For Each char In this.managedVariant
        Mid$(ToHexString, hexStrIndex, HEX_STRING_LENGTH) = VBA.Right$(VBA.Hex(HEX_STRING_PADDING Or char), HEX_STRING_LENGTH)
        hexStrIndex = hexStrIndex + HEX_STRING_LENGTH
    Next
    'EraseSafeArray 'Removed for performance as each call ToHexString redims the SafeArray
End Function

''
'@Description("Creates and initializes a managed variant for an Integer SafeArray.")
'------------------------------------------------------------'
Private Sub InitializeManagedVariant()
Attribute InitializeManagedVariant.VB_Description = "Creates and initializes a managed variant for an Integer SafeArray."
    CreateIntegerSafeArray
    CreateManagedVariant
End Sub

''
'@Description "Copies the unmanaged variant to the unsafe managed variant"
'@Remarks
'
'------------------------------------------------------------'
Private Sub CreateManagedVariant()
Attribute CreateManagedVariant.VB_Description = "Copies the unmanaged variant to the unsafe managed variant"
    this.unmanagedVariant.vt = VBA.vbArray + VBA.vbInteger              ' Declare variant contains an integer array
    this.unmanagedVariant.pSomething = VarPtr(this.integerArrayOverlay) ' Set variant array pointer to integer array overlay
    'Create managed variant from unmanaged variant
    CopyMemory_byPtr VarPtr(this.managedVariant), VarPtr(this.unmanagedVariant), LenB(this.unmanagedVariant)
End Sub

''
'@Description("Destroys the managed variant by zero filling.")
'@Remarks
' The byte size of a variant is 16 btyes which must be zero
' filled to destroy the managed variant.
'------------------------------------------------------------'
Private Sub DisposeManagedVariant()
Attribute DisposeManagedVariant.VB_Description = "Destroys the managed variant by zero filling."
    Const VARIANT_BYTESIZE = 16
#If Not Mac Then
    FillMemory this.managedVariant, VARIANT_BYTESIZE, 0
#Else
    Dim zeroFilledVariant As Variant 'i.e. uninitialized and is nothing.
    CopyMemory_byPtr VarPtr(this.managedVariant), VarPtr(zeroFilledVariant), VARIANT_BYTESIZE 'Copy zero filled variant to managed one to destroy it.
#End If
End Sub


''
'@Description("Creates an unsafe managed SafeArray variant for a one-dimensional integer array.")
'------------------------------------------------------------'
Private Sub CreateIntegerSafeArray()
Attribute CreateIntegerSafeArray.VB_Description = "Creates an unsafe managed SafeArray variant for a one-dimensional integer array."
    With this.integerArrayOverlay
        .cbElements = INTEGER_BYTESIZE                                  ' Integer size is two bytes
        .cDims = 1                                                      ' View as one dimensional array of integers
    End With
End Sub

''
'@Description("Redims and allocates the data pointer for the one-dimensional integer array for the specified string.")
'------------------------------------------------------------'
Private Sub SafeArrayRedim(ByRef inText As Variant)
Attribute SafeArrayRedim.VB_Description = "Redims and allocates the data pointer for the one-dimensional integer array for the specified string."
    With this.integerArrayOverlay
        .pvData = StrPtr(inText)                                        'Allocate Data, the array data points to the string pointer
        .rgsabound(0).lLbound = 0
        .rgsabound(0).cElements = LenB(inText) / INTEGER_BYTESIZE
    End With
End Sub

''
'@Description("Erases the integer array by clearing the data pointer and number of elements.")
'------------------------------------------------------------'
Private Sub EraseSafeArray()
Attribute EraseSafeArray.VB_Description = "Erases the integer array by clearing the data pointer and number of elements."
    With this.integerArrayOverlay
        .pvData = 0                'Array data is the string pointer
        .rgsabound(0).cElements = 0
    End With
End Sub

'============================================='
'Interfaces
'============================================='
' --------------------------------------------- '
' Interface ITextEncoding
' --------------------------------------------- '
Private Function ITextEncoding_ToHexString(ByRef inText As String) As String
    ITextEncoding_ToHexString = Me.ToHexString(inText)
End Function
