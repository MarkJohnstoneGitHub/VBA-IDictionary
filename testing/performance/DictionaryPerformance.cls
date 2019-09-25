VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "IScriptingDictionary factory for testing performance of various implementations, not for general use.\r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@Folder("VBA-IScriptingDictionary.Tests.Performance")
'@PredeclaredId
'@ModuleDescription "IScriptingDictionary factory for testing performance of various implementations, not for general use.\r\n\r\nVBA-IDictionary v2.0 (July 28, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"

'@REFERENCEADDIN Microsoft Scripting Runtime Scripting scrrun.dll
''
'@Version VBA-IDictionary v2.0 (July 28, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description IDictionary factory for testing performance of various IDictionary implementations.
'@Dependencies IDictionary.cls, DictionaryPerformance.cls, ScriptingDictionary.cls, ScriptingDictionaryLateBinding,
'              DictionaryKeyValuePair.cls, DictionaryVBATools
'              Optional: Reference addin Microsoft Scripting Runtime Scripting, scrrun.dll dependant on compiler constant settings.
'@Author markjohnstone@hotmail.com
'@LastModified July 28, 2019
'
''@Usage Eg.
' Dim myDictionary As IDictionary
' myDictionary = DictionaryPerformance.Create(IScriptingDictionaryPerformanceType.isdtDictionaryKeyValuePair,VBA.vbBinaryCompare,TextEncodingMethod.temUnicode)
'
'@Remarks
' Used only for testing for performance of DictionaryKeyValuePair compared to the following IDictionary implementations:
' A modified version of Tim Hall 's Dictionary at https://github.com/timhall/VBA-Dictionary see DictionaryVBATools.cls
' Microsoft Scripting Runtime Scripting early binding see ScriptingDictionary.cls
' Microsoft Scripting Runtime Scripting late binding see ScriptingDictionaryLateBinding.cls
' On the Mac the performance testing for the Scripting.Dictionary cannot be performed.
'
' Compiler Constants.
' The compiler constants SCRIPTING_REFERENCE and SCRIPTING_LATEBINDING are
' used to determine at compile time whether the library is referenced or if not.
' If the reference is not available it will attempt to create using late binding.
' These compiler constants are required to be appropriately set in the
' ScriptingDictionary.cls and the DictionaryFactoryTesting.cls
' If not set appropriately they may cause compile errors.
'
' Updating the compiler constants manually in ScriptingDictionary.cls and the Dictionary.cls
' If the Scripting Runtime library is referenced set to
'    SCRIPTING_REFERENCE = True
' If the Scripting is not referenced and is available for late binding set
'    SCRIPTING_REFERENCE = False, SCRIPTING_LATEBINDING = True
' For Mac set
'    SCRIPTING_REFERENCE = False, SCRIPTING_LATEBINDING = False
' For the Mac if ScriptingDictionary is selected error will be raised
'
'@VersionNotes VBA-IDictionary v1.21 (July 15, 2019)
' Rename DictionaryTesting.cls to DictionaryPerformance.cls
' Removed private function for CreateDictionaryKeyValue pair.
' Modified naming of enum IDictionaryTypeTesting to IScriptingDictionaryPerformanceType.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit

'============================================='
'Compiler Constants
'============================================='
'@CompilerConstants
#Const SCRIPTING_REFERENCE = True
#Const SCRIPTING_LATEBINDING = True

'List of available IDictionary implementations
Public Enum IScriptingDictionaryPerformanceType
    isdtScriptingDictionary = 1
    isdtScriptingLateBinding = 2
    isdtDictionaryKeyValuePair = 3
    isdtDictionaryVBATools = 4
    [_First] = isdtScriptingDictionary
    [_Last] = isdtDictionaryVBATools
End Enum

'============================================='
'Constructors and destructors
'============================================='
''
'@Description Creates a Singleton class which cannot be instantiated.
'------------------------------------------------------------'
Private Sub Class_Initialize()
    If Not Me Is DictionaryPerformance Then
        Err.Raise 429, "DictionaryPerformance", "Cannot create an object instance of the DictionaryPerformance."
    End If
End Sub

'============================================='
'Public Methods
'============================================='
''
'@Description("Creates a IScriptingDictionary for the specified IDictionary implementation and according to the compilier constants.")
'@param dictionaryType     (Optional) The dictionary implementation specified to implement a IDictionary, default is ScriptingDictionary.
'@param compareMethod      (Optional) The comparison mode for comparing string keys in a Dictionary object, default is vbBinaryCompare.
'@param textEncodingMethod (Optional) The encoding mode (unicode or ascii) for case senstive keys, default is temUnicode.
'@return (IScriptingDictionary)
'@Error 9   Subscript out of range.
'           Raised for an invalid dictionaryType specified when outside of the range of accepted values of a IDictionaryType.
'@Error 13  Type Mismatch.
'           Raised in the calling code for an invalid data type for the dictionaryType specified.
'@Error 429 Run-time error '429': ActiveX component can't create object VBA
'           Raised when cannot create object as Scripting.Dictionary is not available.
'------------------------------------------------------------'
Public Function Create(Optional ByVal dictionaryType As IScriptingDictionaryPerformanceType = IScriptingDictionaryPerformanceType.isdtScriptingDictionary, _
                        Optional ByVal compareMethod As VBA.VbCompareMethod = VBA.vbBinaryCompare, _
                        Optional ByVal encodingMethod As TextEncodingMethod = TextEncodingMethod.temUnicode) _
                        As IScriptingDictionary
Attribute Create.VB_Description = "Creates a IScriptingDictionary for the specified IDictionary implementation and according to the compilier constants."
    Dim result As IScriptingDictionary
    Select Case dictionaryType
        Case IScriptingDictionaryPerformanceType.isdtScriptingDictionary
            #If Mac Then
                VBA.Err.Raise 429, , "Cannot create ScriptingDictionary using early binding as not available on Mac."
            #ElseIf SCRIPTING_REFERENCE Then
                Set result = ScriptingDictionary.Create(compareMethod)
            #Else
                VBA.Err.Raise 429, , "Cannot create ScriptingDictionary using early binding due to compiler constants."
            #End If
        Case IScriptingDictionaryPerformanceType.isdtScriptingLateBinding
            #If Mac Then
                VBA.Err.Raise 429, , "Cannot create ScriptingDictionary using late binding as not available on Mac."
            #ElseIf SCRIPTING_LATEBINDING Then
                Set result = ScriptingDictionaryLateBinding.Create(compareMethod)
            #Else
                VBA.Err.Raise 429, , "Cannot create ScriptingDictionary using late binding due to compiler constants."
            #End If
        Case IScriptingDictionaryPerformanceType.isdtDictionaryKeyValuePair
            Set result = DictionaryKeyValuePair.Create(compareMethod, encodingMethod)
        Case IScriptingDictionaryPerformanceType.isdtDictionaryVBATools
            Set result = New DictionaryVBATools
            result.CompareMode = compareMethod
        Case Else
            ' illeagal value
            VBA.Err.Raise 9 '<- Subscript out of range
    End Select
    Set Create = result
End Function



