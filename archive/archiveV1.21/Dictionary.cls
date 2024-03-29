VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "IDictionary factory which is a VBA dictionary alternative for the Microsoft Scripting.Dictionary which is Mac compatible.\r\n\r\nVBA-IDictionary v1.21 (July 15, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@ModuleDescription "IDictionary factory which is a VBA dictionary alternative for the Microsoft Scripting.Dictionary which is Mac compatible.\r\n\r\nVBA-IDictionary v1.21 (July 15, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@PredeclaredId
'@Folder("VBA-IDictionary")
'@REFERENCEADDIN Microsoft Scripting Runtime Scripting scrrun.dll

''
'@Version VBA-IDictionary v1.21 (July 15, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description IDictionary factory which is a VBA dictionary alternative for the Microsoft Scripting.Dictionary which is Mac compatible.
'@Author markjohnstone@hotmail.com
'@LastModified July 20, 2019
'@Dependencies ("IDictionary.cls, Dictionary.cls, DictionaryKeyValuePair.cls, ScriptingDictionary.cls")
'              Optional: Reference addin Microsoft Scripting Runtime Scripting, scrrun.dll dependant on compiler constant settings.
'@Package IDictionary.cls
'         Dictionary.cls
'         ScriptingDictionary.cls
'         DictionaryKeyValuePair.cls
'
'@Usage   Eg. Dim myDictionary As IDictionary
'             Set myDictionary = Dictionary.Create
'             Set myDictionary = Dictionary.Create(IDictionaryType.IDict_Scripting, VBA.vbTextCompare)
'             Set myDictionary = Dictionary.Create(IDictionaryType.IDict_KeyValuePair, VBA.vbBinaryCompare, TextEncodingMethod.temAscii)
'
'@Remarks
'  The default IDictionary returned is the ScriptingDictionary unless not available and the
'  DictionaryKeyValuePair will be returned.
'
'  Compiler Constants.
'  The compiler constants SCRIPTING_REFERENCE and SCRIPTING_LATEBINDING are
'  used to determine at compile time whether the library is referenced or if not.
'  If the reference is not available it will attempt to create using late binding.
'  These compiler constants are required to be appropriately set in the
'  ScriptingDictionary.cls and the Dictionary.cls
'  If not set appropriately they may cause compile errors.
'
'  Updating the compiler constants manually in ScriptingDictionary.cls and the Dictionary.cls
'  If the Scripting Runtime library is referenced set to
'      SCRIPTING_REFERENCE = True
'  If the Scripting is not referenced and is available for late binding set
'      SCRIPTING_REFERENCE = False, SCRIPTING_LATEBINDING = True
'  For Mac set
'      SCRIPTING_REFERENCE = False, SCRIPTING_LATEBINDING = False
'
'  For the Mac the DictionaryKeyValuePair will be used as an alternative dictionary.
'
'@VersionNotes VBA-IDictionary v1.1 (July 14, 2019)
' Modified the method Create adding optional parameters, compareMethod and encodingMethod.
' The text encoding parameter provides the option for increased performance for the encoding of
' case sensitive string keys, using ASCII encoding.
' Its default value is Unicode encoding and only applicable to the DictionaryKeyValuePair.cls.
' It can only be changed if there are no items in the dictionary object.
' When using ASCII encoding the calling code is responsible for ensuring ASCII compatible
' string keys are provided or indetermined results may occur.
'@VersionNotes VBA-IDictionary v1.21 (July 15, 2019)
' Modified ScriptingDictionary.cls and DictionaryKeyValuePair.cls to Predeclared added a Create method.
' Removed private function for CreateDictionaryKeyValue pair and moved to the create method of DictionaryKeyValuePair.
' Modified naming of enums in IDictionaryType to avoid any conflicts with class names.
' Modified naming of enums in TextEncodingMethod DictionaryKeyValuePair.cls.
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

'List of available IDictionary implementations
Public Enum IDictionaryType
    idtScriptingDictionary = 0
    idtDictionaryKeyValuePair = 2
    [_First] = idtScriptingDictionary
    [_Last] = idtDictionaryKeyValuePair
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
'@Description("Creates a IDictionary for the specified IDictionary implementation and according to the compilier constants.")
'@param dictionaryType (Optional) The dictionary implementation specified to implement a IDictionary, default is ScriptingDictionary.
'@param compareMethod  (Optional) The comparison mode for comparing string keys in a Dictionary object, default is vbBinaryCompare.
'@param encodingMethod (Optional) The encoding mode (unicode or ascii) for case senstive keys, default is temUnicode.
'@return (IDictionary)
'@Error 9  Subscript out of range.
'          Raised for an invalid dictionaryType specified when outside of the range of accepted values of the IDictionaryType.
'          Raised of an invalid compareMethod specified when outside of the range of accepted values of the VBA.VbCompareMethod.
'          Raised of an invalid encodingMethod specified when outside of the range of accepted values of the TextEncodingMethod.
'@Error 13 Type Mismatch.
'          Raised in the calling code for an invalid data type for the the following: dictionaryType , compareMethod, encodingMethod
'@Remarks
' The alternative compatible IDictionary used is DictionaryKeyValuePair.
' Unless invalid parameters are supplied a IDictionary will be returned even if the Scripting.Dictionary reference
' is broken or the compiler constants, SCRIPTING_REFERENCE and SCRIPTING_LATEBINDING are both set to false.
' I.e. If Mac or the Scripting.Dictionary isn't referenced, broken link or not available for late binding the alternative
' DictionaryKeyValuePair is used.
'
' The encodingMethod is only applicable to DictionaryKeyValuePair, and only can be changed through this create method or
' when cast the IDictionary to a DictionaryKeyValuePair object before items are added to the dictionary.
'------------------------------------------------------------'
Public Function Create(Optional ByVal dictionaryType As IDictionaryType = IDictionaryType.idtScriptingDictionary, _
                       Optional ByVal compareMethod As VBA.VbCompareMethod = VBA.vbBinaryCompare, _
                       Optional ByVal encodingMethod As TextEncodingMethod = TextEncodingMethod.temUnicode) _
                       As IDictionary
Attribute Create.VB_Description = "Creates a IDictionary for the specified IDictionary implementation and according to the compilier constants."
    On Error GoTo ErrorHandler
    Dim result As IDictionary
    Select Case dictionaryType
        Case IDictionaryType.idtScriptingDictionary
            #If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
                Set result = ScriptingDictionary.Create(compareMethod)
            #ElseIf Mac Or Not (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
                Set result = DictionaryKeyValuePair.Create(compareMethod, encodingMethod)
            #Else
                Set result = DictionaryKeyValuePair.Create(compareMethod, encodingMethod)
            #End If
        Case IDictionaryType.idtDictionaryKeyValuePair
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
        'Use DictionaryKeyValuePair as the alternative IDictionary implemenation
        Set result = DictionaryKeyValuePair.Create(compareMethod, encodingMethod)
    Else
        'Bubble up any other errors
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Resume CleanExit
End Function







