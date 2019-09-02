VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITextEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "The ITextEncoding interface for text encoding implementations.\r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''
'Rubberduck annotations
'@Folder("VBA-IScriptingDictionary.TextEncoding")
'@Interface
'@ModuleDescription "The ITextEncoding interface for text encoding implementations.\r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''

''
'@Version VBA-IScriptingDictionary v2.0 (July 28, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description The ITextEncoding interface used for text encoding implementations.
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified July 28, 2019

Option Explicit

'List of text encoding implemenations
Public Enum TextEncodingMethod
    temUnicode = 1
    temAscii = 2
    [_First] = temUnicode
    [_Last] = temAscii
End Enum

'@Description("Returns an string of hex characters for a specified string.")
'@param inText The string to be encoded into a hex string.
'@returns (String) Hex string characters.
' ------------------------------------------------------------'
Public Function ToHexString(ByRef inText As String) As String
Attribute ToHexString.VB_Description = "Returns an string of hex characters for a specified string."
End Function
