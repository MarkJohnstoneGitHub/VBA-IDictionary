VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextEncoderASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An ASCII text encoder for converting a string to a ASCII hex string. \r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''
'Rubberduck annotations
'@Folder("VBA-IScriptingDictionary.TextEncoding")
'@ModuleDescription "An ASCII text encoder for converting a string to a ASCII hex string. \r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''

''
'@Version VBA-IScriptingDictionary v2.0 (July 28, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified July 28, 2019
' ------------------------------------------------------------'

Option Explicit
Implements ITextEncoding

'@Description("Returns an string of Ascii hex characters for a specified string.")
'@param inText The string to be encoded into an ASCII hex string.
'@returns (String) Ascii Hex characters.
'@Remarks
' Unicode bytes are in little-endian order, for ASCII conversion only require every second byte
' from the first byte.
' Eg. The hex characters of "ABC" = 41 00 42 00 43 00
' The ASCII hex string returned would be "414243".
' ------------------------------------------------------------'
Public Function ToHexString(ByRef inText As String) As String
Attribute ToHexString.VB_Description = "Returns an string of Ascii hex characters for a specified string."
    Const CHAR_LENGTH As Long = 2
    
    Dim bytes() As Byte
    bytes = inText                      'Assign string to unicode byte array.

    'Allocate string size the required length of the hex string i.e. For ASCII each character is converted to two hex characters.
    ToHexString = VBA.Space$(CHAR_LENGTH * (UBound(bytes) - LBound(bytes) + 1))
    Dim byteIndex As Long
    For byteIndex = 0 To UBound(bytes) Step 2
        Mid$(ToHexString, byteIndex + 1, CHAR_LENGTH) = VBA.Right$(VBA.Hex(256 Or bytes(byteIndex)), CHAR_LENGTH)
    Next byteIndex

End Function

'============================================='
'Interfaces
'============================================='

' --------------------------------------------- '
' Interface ITextEncoding
' --------------------------------------------- '
Private Function ITextEncoding_ToHexString(ByRef inText As String) As String
    ITextEncoding_ToHexString = Me.ToHexString(inText)
End Function

