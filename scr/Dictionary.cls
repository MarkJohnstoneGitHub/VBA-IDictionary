VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "IScriptingDictionary factory which is a VBA dictionary alternative for the Microsoft Scripting.Dictionary which is Mac compatible.\r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''
'Rubberduck annotations
'@Folder("VBA-IScriptingDictionary")
'@PredeclaredId
'@ModuleDescription "IScriptingDictionary factory which is a VBA dictionary alternative for the Microsoft Scripting.Dictionary which is Mac compatible.\r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@REFERENCEADDIN Optional Microsoft Scripting Runtime Scripting scrrun.dll

''
'@Version VBA-IDictionary v2.0 (July 28, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description IScriptingDictionary factory which is a VBA dictionary alternative for the Microsoft Scripting.Dictionary which is Mac compatible.
'@Author markjohnstone@hotmail.com
'@LastModified July 28, 2019
'@Dependencies
'   IScriptingDictionary.cls
'   Dictionary.cls
'   DictionaryKeyValuePair.cls
'   ScriptingDictionary.cls
'   ITextEncoding.cls
'   TextEncoderUnicode.cls
'   TextEncoderASCII.cls
'
'   Optional: Reference addin Microsoft Scripting Runtime Scripting, scrrun.dll dependant on compiler constant settings.
'
'@Usage
' Eg. Dim myDictionary As IScriptingDictionary
'     Set myDictionary = Dictionary.Create
'     Set myDictionary = Dictionary.Create(IScriptingDictionaryType.isdtScriptingDictionary, VBA.vbTextCompare)
'     Set myDictionary = Dictionary.Create(IScriptingDictionaryType.isdtDictionaryKeyValuePair, VBA.vbBinaryCompare, TextEncodingMethod.temAscii)
'
'@Remarks
' The Dictionary class can only be created using the Dictionary.Create method.
' The default IScriptingDictionary returned is the ScriptingDictionary unless not available and the
' DictionaryKeyValuePair will be returned as the alternative.
'
' Compiler Constants.
' The compiler constants SCRIPTING_REFERENCE and SCRIPTING_LATEBINDING are
' used to determine at compile time whether the Microsoft Scripting Runtime library is referenced or if not.
'
' These compiler constants are required to be appropriately set in the
' ScriptingDictionary.cls and the Dictionary.cls
' If not set appropriately they may cause compile errors.
' If the reference is not available it will attempt to create using late binding.
'
' Updating the compiler constants manually in ScriptingDictionary.cls and the Dictionary.cls
' If the Scripting Runtime library is referenced set to
'     SCRIPTING_REFERENCE = True
' If the Scripting is not referenced and is available for late binding set
'     SCRIPTING_REFERENCE = False, SCRIPTING_LATEBINDING = True
' For Mac set
'     SCRIPTING_REFERENCE = False, SCRIPTING_LATEBINDING = False
'
' For the Mac the DictionaryKeyValuePair will be used as an alternative dictionary.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
Option Explicit

'============================================='
'Compiler Constants
'============================================='
'@CompilerConstants
#Const SCRIPTING_REFERENCE = True
#Const SCRIPTING_LATEBINDING = True

'============================================='
'Types
'============================================='
'List of available IScriptingDictionary implementations
Public Enum IScriptingDictionaryType
    isdtScriptingDictionary = 1
    isdtDictionaryKeyValuePair = 3
    [_First] = isdtScriptingDictionary
    [_Last] = isdtDictionaryKeyValuePair
End Enum

'============================================='
'Constructors and destructors
'============================================='
''
'@Description Creates a Singleton class which cannot be instantiated.
'------------------------------------------------------------'
Private Sub Class_Initialize()
    If Not Me Is Dictionary Then
        VBA.Err.Raise 429, "Dictionary", "Cannot create object Dictionary."
    End If
End Sub

'============================================='
'Public Methods
'============================================='
''
'@Description("Creates a IScriptingDictionary for the specified implementation and according to the compilier constants.")
'@param dictionaryType (Optional) The dictionary implementation specified to implement a IScriptingDictionary, default is ScriptingDictionary.
'@param compareMethod  (Optional) The comparison mode for comparing string keys in a Dictionary object, default is vbBinaryCompare.
'@param encodingMethod (Optional) The encoding mode (unicode or ascii) for case senstive keys, default is temUnicode.
'@return (IScriptingDictionary)
'@Error 9  Subscript out of range.
'          Raised for an invalid dictionaryType specified when outside of the range of accepted values of the IScriptingDictionaryType.
'          Raised of an invalid compareMethod specified when outside of the range of accepted values of the VBA.VbCompareMethod.
'          Raised of an invalid encodingMethod specified when outside of the range of accepted values of the TextEncodingMethod.
'@Error 13 Type Mismatch.
'          Raised in the calling code for an invalid data type for the the following: dictionaryType , compareMethod, encodingMethod
'@Remarks
' The alternative compatible IScriptingDictionary used is DictionaryKeyValuePair.
' Unless invalid parameters are supplied a IScriptingDictionary will be returned even if the Scripting.Dictionary reference
' is broken or the compiler constants, SCRIPTING_REFERENCE and SCRIPTING_LATEBINDING are both set to false.
' I.e. If Mac or the Scripting.Dictionary isn't referenced, broken link or not available for late binding the alternative
' DictionaryKeyValuePair is used.
'
' The encodingMethod is only applicable to DictionaryKeyValuePair, and only can be changed through this create method or
' when cast the IScriptingDictionary to a DictionaryKeyValuePair object before items are added to the dictionary.
'------------------------------------------------------------'
Public Function Create(Optional ByVal dictionaryType As IScriptingDictionaryType = IScriptingDictionaryType.isdtScriptingDictionary, _
                       Optional ByVal compareMethod As VBA.VbCompareMethod = VBA.vbBinaryCompare, _
                       Optional ByVal encodingMethod As TextEncodingMethod = TextEncodingMethod.temUnicode) _
                       As IScriptingDictionary
Attribute Create.VB_Description = "Creates a IScriptingDictionary for the specified implementation and according to the compilier constants."
    On Error GoTo ErrorHandler
    Dim result As IScriptingDictionary
    Select Case dictionaryType
        Case IScriptingDictionaryType.isdtScriptingDictionary
            #If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
                Set result = ScriptingDictionary.Create(compareMethod)
            #ElseIf Mac Or Not (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
                Set result = DictionaryKeyValuePair.Create(compareMethod, encodingMethod)
            #Else
                Set result = DictionaryKeyValuePair.Create(compareMethod, encodingMethod)
            #End If
        Case IScriptingDictionaryType.isdtDictionaryKeyValuePair
                Set result = DictionaryKeyValuePair.Create(compareMethod, encodingMethod)
        Case Else
            ' illeagal value
            VBA.Err.Raise 9 '<- Subscript out of range
    End Select
CleanExit:
    Set Create = result
    Exit Function
ErrorHandler:
    On Error GoTo 0 'Disable any previous VBA error handling
    'Run-time error '429': ActiveX component can't create object
    If Err.Number = 429 Then
        Err.Clear 'reset error trapping.
        'Use DictionaryKeyValuePair as the alternative IScriptingDictionary implemenation
        Set result = DictionaryKeyValuePair.Create(compareMethod, encodingMethod)
    Else
        'Bubble up any other errors
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Resume CleanExit
End Function


