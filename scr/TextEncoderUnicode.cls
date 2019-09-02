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
'@TODO Rename ToHexString to Encode and some for Interface functions

''
'Rubberduck annotations
'@Folder("VBA-IScriptingDictionary.TextEncoding")
'@ModuleDescription "A Unicode text encoder that converts a string to a unicode hex string using a managed variant read-only integer array overlay. \r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''

''
'@Version VBA-IScriptingDictionary v2.1 (September 02, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified September 02, 2019
'@Description Converts a string to a unicode hex string using a read-only integer array overlay.
''@Dependencies
'   ManagedCharSafeArray.cls
'   TypeSafeArray.bas
'
'@Usage   Eg. Dim textEncoder As UnicodeTextEncoder
'             Set textEncoder = New UnicodeTextEncoder
'             Debug.Print textEncoder.ToHexString("ABCabc")
'             Set textEncoder = Nothing
'@Remarks
'------------------------------------------------------------'

Option Explicit
Implements ITextEncoding

Private Type TUnicodeTextEncoder
    managedCharDescriptor As ManagedCharSafeArray 'managed safe array descriptor for managedChar
    managedChar() As Integer                 'managed character array containing unicode of string in integer form.
End Type

'============================================='
'Private Variables
'============================================='
Private this As TUnicodeTextEncoder

'============================================='
'Constructors and destructors
'============================================='
Private Sub Class_Initialize()
    Set this.managedCharDescriptor = New ManagedCharSafeArray
    this.managedChar = this.managedCharDescriptor.ManagedCharArray
End Sub

Private Sub Class_Terminate()
    this.managedCharDescriptor.FreeManagedCharArray this.managedChar
End Sub

'============================================='
'Public Methods
'============================================='

''
'@Description("Converts a string to unicode hex string.")
'@param inText The unicode string to be converted to a unicode hex string.
'@Remarks
' Each string character is converted into four hex characters.
' Interuption or resetting while running this function may cause
' the application to crash if the managed chars array hasn't been freed.
'------------------------------------------------------------'
Public Function ToHexString(ByRef inText As String) As String
Attribute ToHexString.VB_Description = "Converts a string to unicode hex string."
    Const HEX_CHARACTER_LENGTH As Long = 4      'Each string character is represented by four UTF-16 hexidemial characters.
    Const HEX_CHARACTER_PADDING As Long = 65536 'Ensures that four hex characters are obtained with 0 padding
    
    this.managedCharDescriptor.AllocateCharData inText 'Allocate string data to managedChar array
    'Allocate the hex string buffer according to specified string size, where each string character converts to four hex characters
    ToHexString = VBA.Space$(HEX_CHARACTER_LENGTH * (UBound(this.managedChar) - LBound(this.managedChar) + 1))
    
    Dim charIndex As Long
    For charIndex = LBound(this.managedChar) To UBound(this.managedChar)
        Mid$(ToHexString, (charIndex * HEX_CHARACTER_LENGTH) + 1, HEX_CHARACTER_LENGTH) = VBA.Right$(VBA.Hex(HEX_CHARACTER_PADDING Or this.managedChar(charIndex)), HEX_CHARACTER_LENGTH)
    Next
End Function

'============================================='
'Interfaces
'============================================='
' --------------------------------------------- '
' Interface ITextEncoding
' --------------------------------------------- '
Private Function ITextEncoding_Encode(ByRef inText As String) As String
    ITextEncoding_Encode = Me.ToHexString(inText)
End Function
