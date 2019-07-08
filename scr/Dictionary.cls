VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "IDictionary factory which is a VBA dictionary alternative for the Microsoft Scripting.Dictionary which is Mac compatible.\r\n\r\nIDictionary v1.0 (July 08, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@PredeclaredId
'@Folder("VBA-IDictionary")
'@REFERENCEADDIN Microsoft Scripting Runtime Scripting scrrun.dll

''
'@Version IDictionary v1.0 (July 08, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description IDictionary factory which is a VBA dictionary alternative for the Microsoft Scripting.Dictionary which is Mac compatible.
'@Author markjohnstone@hotmail.com
'@LastModified July 08, 2019
'@Required IDictionary.cls, Dictionary.cls, DictionaryKeyValuePair.cls, ScriptingDictionary.cls
'@Optional Reference addin Microsoft Scripting Runtime Scripting, scrrun.dll dependant on Compiler Constants settings.
'@Package IDictionary.cls
'         Dictionary.cls
'         ScriptingDictionary.cls
'         DictionaryKeyValuePair.cls
'
'@Usage   Eg. Dim myDictionary As IDictionary
'             myDictionary = Dictionary.Create
'             myDictionary = Dictionary.Create(IDictionaryType.ScriptingDictionary)
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
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
Option Explicit

'============================================='
'Compiler Constants
'============================================='
'@CompilerConstants
#Const SCRIPTING_REFERENCE = False
#Const SCRIPTING_LATEBINDING = False

'List of available IDictionary implementations
Public Enum IDictionaryType
    ScriptingDictionary = 0
    DictionaryKeyValuePair = 2
    [_First] = ScriptingDictionary
    [_Last] = DictionaryKeyValuePair
End Enum

'============================================='
'Constants
'============================================='
Const SCRIPTING_DICTIONARY As String = "Scripting"

'============================================='
'Constructors and destructors
'============================================='

''
'@Description Creates a Singleton class which cannot be instantiated.
'------------------------------------------------------------'
Private Sub Class_Initialize()
    If Not Me Is Dictionary Then
        Err.Raise VBA.vbObjectError, "Dictionary", "Cannot create an object instance of the Dictionary."
    End If
End Sub

'============================================='
'Public Methods
'============================================='

''
'@Description("Creates a IDictionary for the specified IDictionary implementation and according to the compilier constants.")
'@param dictionaryType (Optional) The dictionary implementation specified which implements a IDictionary.
'@return (IDictionary)
'@Error 9  Subscript out of range.
'          Raised for an invalid dictionaryType specified when outside of the range of accepted values of a IDictionaryType.
'@Error 13 Type Mismatch.
'          Raised in the calling code for an invalid data type for the dictionaryType specified.
'------------------------------------------------------------'
Public Function Create(Optional ByVal dictionaryType As IDictionaryType = IDictionaryType.ScriptingDictionary) As IDictionary
Attribute Create.VB_Description = "Creates a IDictionary for the specified IDictionary implementation and according to the compilier constants."
    Dim result As IDictionary
    Select Case dictionaryType
        Case IDictionaryType.ScriptingDictionary
            #If Mac Then
                Set result = New DictionaryKeyValuePair
            #ElseIf SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING Then
                Set result = New ScriptingDictionary
            #Else
                Set result = New DictionaryKeyValuePair
            #End If
        Case IDictionaryType.DictionaryKeyValuePair
            Set result = New DictionaryKeyValuePair
        Case Else
            ' illeagal value
            VBA.Err.Raise 9 '<- Subscript out of range
    End Select
    Set Create = result
End Function
