VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A VBA wrapper for the Microsoft Scripting.Dictionary, scrrun.dll library, implementing an IDictionary interface.\r\n\r\nVBA-IDictionary v1.21 (July 15, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@PredeclaredId
'@ModuleDescription "A VBA wrapper for the Microsoft Scripting.Dictionary, scrrun.dll library, implementing an IDictionary interface.\r\n\r\nVBA-IDictionary v1.21 (July 15, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@Folder("VBA-IDictionary")
'@REFERENCEADDIN Microsoft Scripting Runtime Scripting scrrun.dll

'@Version VBA-IDictionary v1.21 (July 15, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description A VBA wrapper for the Microsoft Scripting.Dictionary, scrrun.dll library
'@Dependencies IDictionary.cls, ScriptingDictionary.cls
'              Optional: Reference addin Microsoft Scripting Runtime Scripting, scrrun.dll dependant on compiler constant settings.
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified July 12, 2019
'@Package IDictionary.cls
'         Dictionary.cls
'         ScriptingDictionary.cls
'         DictionaryKeyValuePair.cls
'@Errors 429 Cannot Create Object as Scripting.Dictionary is not available.
'            Raised for Mac or Scripting.Dictionary isn't available from compiler arguments.
'            It can also be raised for late binding where the call to CreateObject fails.
'@Remarks
' The compiler constants SCRIPTING_REFERENCE and SCRIPTING_LATEBINDING are
' used to determine whether the library is referenced or if not. If not available
' to created using late binding.
' These compiler constants are required to be appropriately set in the
' ScriptingDictionary.cls and the Dictionary.cls
' If not set appropriately they may cause compile errors.
'
' Updating the compiler constants
' If the Microsoft Scripting Runtime is referenced set:
'    SCRIPTING_REFERENCE = True
' If the Microsoft Scripting Runtime is not referenced and is available for late binding set
'    SCRIPTING_REFERENCE = False, SCRIPTING_LATEBINDING = True
' For Mac, set to:
'    SCRIPTING_REFERENCE = False, SCRIPTING_LATEBINDING = True
'
'@VersionNotes VBA-IDictionary v1.21 (July 15, 2019)
' Modified to Predeclared added a Create method.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
Option Explicit

'============================================='
'Compiler Constants
'============================================='
'@CompilerConstants
#Const SCRIPTING_REFERENCE = True
#Const SCRIPTING_LATEBINDING = True

Implements IDictionary

'============================================='
'Types
'============================================='
#If Not Mac And SCRIPTING_REFERENCE Then
    Private Type TScriptingDictionary
        scriptDictionary As Scripting.Dictionary
    End Type

#ElseIf Not Mac And SCRIPTING_LATEBINDING Then
    Private Type TScriptingDictionary
        scriptDictionary As Object
    End Type
#End If

'============================================='
'Constants and Private Variables
'============================================='
Const SCRIPTING_DICTIONARY As String = "Scripting.Dictionary"

'============================================='
'Private Variables
'============================================='
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Private this As TScriptingDictionary
#End If

    
'============================================='
'Constructors and destructors
'============================================='

'@Description Initializes the ScriptingDictionary with either early or late binding according to the constant compiler arguments.
'@Errors 429 Run-time error '429': ActiveX component can't create object
'            Raised for Mac or Scripting.Dictionary isn't available from compiler arguments.
'            It can also be raised for late binding where the call to CreateObject fails.
Private Sub Class_Initialize()
    #If Not Mac And SCRIPTING_REFERENCE Then
        Set this.scriptDictionary = New Scripting.Dictionary
    #ElseIf Not Mac And SCRIPTING_LATEBINDING Then
        Set this.scriptDictionary = CreateObject(SCRIPTING_DICTIONARY)
    #Else
        VBA.Err.Raise 429 'Cannot create object as Scripting.Dictionary is not available."
    #End If
End Sub

Private Sub Class_Terminate()
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Set this.scriptDictionary = Nothing
#End If
End Sub

'============================================= '
'Properties
'============================================= '

''
'@Description("Sets or returns the comparison mode for comparing keys in a Dictionary object.")
'------------------------------------------------------------'
Public Property Get CompareMode() As VBA.VbCompareMethod
Attribute CompareMode.VB_Description = "Sets or returns the comparison mode for comparing keys in a Dictionary object."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    CompareMode = this.scriptDictionary.CompareMode
#End If
End Property

''
'@Description("Sets the comparison mode for comparing keys in a Dictionary object.")
'@param compareMethod Value representing the comparison mode used by functions such as StrComp.
'@Error 5 Invalid procedure call or argument
'         Raised for an invalid data type for compareMethod.
'@Error 9 Subscript out of range
'         Raised for an invalid value for compareMethod.
'@Remarks The comparison mode cannot be changed once the dictionary contains items
'------------------------------------------------------------'
Public Property Let CompareMode(ByVal compareMethod As VBA.VbCompareMethod)
Attribute CompareMode.VB_Description = "Sets the comparison mode for comparing keys in a Dictionary object."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.CompareMode = compareMethod
#End If
End Property

''
'@Description("Returns the number of key, item pairs in a Dictionary object.")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of key, item pairs in a Dictionary object."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Count = this.scriptDictionary.Count
#End If
End Property

''
'@Enumerator
'@Description("Enumerator for dictionary which returns dictionary keys.")
'@Remarks Attribute NewEnum.VB_UserMemId = -4 declares NewEnum as the enumerator
'         Attribute NewEnum.VB_MemberFlags = "40" declares NewEnum as a hidden property.
' ------------------------------------------------------------'
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator for dictionary which returns dictionary keys."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Set NewEnum = this.scriptDictionary.Keys.[_NewEnum]
#End If
End Property

''
'@DefaultMember
'@Description("Sets or returns the value of an item in a Dictionary object.")
'@param key The key associated with the item being retrieved.
'@Error 5 Invalid procedure call or argument.
'         Raised for invalid key data type.
'@Remarks
'  The syntax for setting an item is: dictionaryobject.Item(key) = newItem
'  If you try to set item to a nonexistent key, a new key, item pair is added
'  to the dictionary, and its associated item is left empty, a sort of "implicit add".
'  Attribute Item.VB_UserMemId = 0 Declares property Item as the default property.
'------------------------------------------------------------'
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_Description = "Sets or returns the value of an item in a Dictionary object."
Attribute Item.VB_UserMemId = 0
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Dim result As Variant
    result = Array(this.scriptDictionary.Item(Key)) 'The get Item property result is wrapped in the Array to avoid running twice
    If VBA.IsObject(result(0)) Then
        Set Item = result(0)
    Else
        Item = result(0)
    End If
#End If
End Property
''
'@DefaultMember
'@Description Sets a scalar item value for a specified key in a Dictionary object
'@param key  The key associated with the item being retrieved or added.
'@param Item The new value associated with the specified key
'@Error 5    Invalid procedure call or argument
'            Raised for an invalid key data type
'@Error 450  Wrong number of arguments of invalid property assignment
'            Raised when an item is an object when expecting a scalar value
'@Remarks
' The syntax for setting an item is: dictionaryobject.Item(key) = newItem
' If you try to set item to a nonexistent key, the key is added to the dictionary,
' and the item is linked to it, a sort of "implicit add."
'------------------------------------------------------------'
Public Property Let Item(ByRef Key As Variant, ByRef Item As Variant)
Attribute Item.VB_Description = "Sets or returns the value of an item in a Dictionary object."
Attribute Item.VB_UserMemId = 0
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.Item(Key) = Item
#End If
End Property
''
'@DefaultMember
'@Description Sets the object value of an item in a Dictionary object.
'@param key  Key associated with the item being retrieved or added.
'@param Item The new value associated with the specified key
'@Error 5    Invalid procedure call or argument
'            Raised for invalid key data type.
'@Error 424  Object required
'            Raised when the item is a scalar value when expecting an object.
'@Remarks
' The syntax for setting an item is: Set dictionaryobject.Item(key) = newItem
' If you try to set item to a nonexistent key, the key is added to the dictionary,
' and the item is linked to it, a sort of an "implicit add."
'------------------------------------------------------------'
Public Property Set Item(ByRef Key As Variant, ByRef Item As Object)
Attribute Item.VB_Description = "Sets or returns the value of an item in a Dictionary object."
Attribute Item.VB_UserMemId = 0
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Set this.scriptDictionary.Item(Key) = Item
#End If
End Property

''
'@Description("Sets a new key for an existing key in a Dictionary object.")
'@param key    The key value being changed.
'@param newKey The new key value that replaces the specified key.
'@Error 5     Invalid procedure call or argument
'             Raised for invalid key data type
'@Error 457   This key is already associated with an element of this collection
'             Raised when new key already exists in the dictionary object
'@Error 32811 Application-defined or object-defined error
'             Raised when previous key doesn't exist in the dictionary object
'------------------------------------------------------------'
Public Property Let Key(ByRef Key As Variant, ByRef newKey As Variant)
Attribute Key.VB_Description = "Sets a new key for an existing key in a Dictionary object."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.Key(Key) = newKey
#End If
End Property

'============================================='
'Public Methods
'============================================='

''
'@Description("Creates a new instance of a ScriptingDictionary.")
'@param compareMethod (Optional) The comparison mode for comparing string keys in a Dictionary object, default is vbBinaryCompare.
'@return (ScriptingDictionary)
Public Function Create(Optional ByVal compareMethod As VBA.VbCompareMethod = VBA.vbBinaryCompare) As ScriptingDictionary
Attribute Create.VB_Description = "Creates a new instance of a ScriptingDictionary."
    Dim newDictionary As ScriptingDictionary
    Set newDictionary = New ScriptingDictionary
    newDictionary.CompareMode = compareMethod
    Set Create = newDictionary
End Function

''
'@Description("Adds a key and item pair to a Dictionary object.")
'@param key  The key associated with the item being added.
'@param item The new item associated with the key being added.
'@Error 5    Invalid procedure call or argument
'            Raised for an invalid/unsupported key data type
'@Error 457  This key is already associated with an element of this collection
'------------------------------------------------------------'
Public Sub Add(ByRef Key As Variant, ByRef Item As Variant)
Attribute Add.VB_Description = "Adds a key and item pair to a Dictionary object."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.Add Key, Item
#End If
End Sub

''
'@Description("Returns True if a specified key exists in the Dictionary object; False if it does not.")
'@param key The key value being searched for in the Dictionary object.
'@returns {Boolean}
'@Error 5   Invalid procedure call or argument
'           Raised for invalid key data type
'------------------------------------------------------------'
Public Function Exists(ByRef Key As Variant) As Boolean
Attribute Exists.VB_Description = "Returns True if a specified key exists in the Dictionary object; False if it does not."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Exists = this.scriptDictionary.Exists(Key)
#End If
End Function

''
'@Description("Returns an array of items in a dictionary object.")
'@return Variant Array of items
'------------------------------------------------------------'
Public Function Items() As Variant
Attribute Items.VB_Description = "Returns an array of items in a dictionary object."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Items = this.scriptDictionary.Items
#End If
End Function

''
'@Description("Returns an array of all the keys in a dictionary object.")
'@returns {Variant} Array of keys
'------------------------------------------------------------'
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Returns an array of all the keys in a dictionary object."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Keys = this.scriptDictionary.Keys
#End If
End Function

''
'@Description("Removes a key, item pair from a Dictionary object.")
'@param key  The key associated with the key/item pair to remove from the Dictionary object.
'@Error 5  Invalid procedure call or argument
'          Raised for invalid/unsupported key data type
'------------------------------------------------------------'
Public Sub Remove(ByRef Key As Variant)
Attribute Remove.VB_Description = "Removes a key, item pair from a Dictionary object."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.Remove Key
#End If
End Sub

''
'@Description("Removes all key, item pairs from a Dictionary object.")
'------------------------------------------------------------'
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Removes all key, item pairs from a Dictionary object."
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.RemoveAll
#End If
End Sub


'============================================='
'Interfaces
'============================================='

'------------------------------------------------------------'
'IDictionary Interface Properties
'------------------------------------------------------------'
Private Property Get IDictionary_CompareMode() As VBA.VbCompareMethod
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    IDictionary_CompareMode = this.scriptDictionary.CompareMode
#End If
End Property
Private Property Let IDictionary_CompareMode(ByVal compareMethod As VBA.VbCompareMethod)
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.CompareMode = compareMethod
#End If
End Property

Private Property Get IDictionary_Count() As Long
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    IDictionary_Count = this.scriptDictionary.Count
#End If
End Property

Private Property Get IDictionary_NewEnum() As IUnknown
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Set IDictionary_NewEnum = Me.NewEnum
#End If
End Property

Private Property Get IDictionary_Item(ByRef Key As Variant) As Variant
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Dim result As Variant
    result = Array(this.scriptDictionary.Item(Key)) 'The get Item property result is wrapped in the Array to avoid running twice
    If IsObject(result(0)) Then
        Set IDictionary_Item = result(0)
    Else
        IDictionary_Item = result(0)
    End If
#End If
End Property
Private Property Let IDictionary_Item(ByRef Key As Variant, ByRef Item As Variant)
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.Item(Key) = Item
#End If
End Property
Private Property Set IDictionary_Item(ByRef Key As Variant, ByRef Item As Object)
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    Set this.scriptDictionary.Item(Key) = Item
#End If
End Property

Private Property Let IDictionary_Key(ByRef Key As Variant, ByRef newKey As Variant)
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.Key(Key) = newKey
#End If
End Property

'------------------------------------------------------------'
'IDictionary Interface Methods
'------------------------------------------------------------'
Private Sub IDictionary_Add(ByRef Key As Variant, ByRef Item As Variant)
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.Add Key, Item
#End If
End Sub

Private Function IDictionary_Exists(ByRef Key As Variant) As Boolean
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    IDictionary_Exists = this.scriptDictionary.Exists(Key)
#End If
End Function

Private Function IDictionary_Items() As Variant
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    IDictionary_Items = this.scriptDictionary.Items
#End If
End Function

Private Function IDictionary_Keys() As Variant
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    IDictionary_Keys = this.scriptDictionary.Keys
#End If
End Function

Private Sub IDictionary_Remove(Key As Variant)
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.Remove Key
#End If
End Sub

Private Sub IDictionary_RemoveAll()
#If Not Mac And (SCRIPTING_REFERENCE Or SCRIPTING_LATEBINDING) Then
    this.scriptDictionary.RemoveAll
#End If
End Sub

