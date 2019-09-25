VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptingDictionaryLateBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
'Rubberduck annotations
'@Folder("VBA-IScriptingDictionary.Tests.Performance")
'@PredeclaredId
''

''
'@Version IDictionary v2.0 (July 28, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description VBA wrapper for the Scripting.Dictionary, scrrun.dll library using late binding.
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified July 28, 2019
'@Usage   Eg. Dim myDictionary As ScriptingDictionaryLateBinding
'             myDictionary = New ScriptingDictionaryLateBinding
'
'@Remarks If the Scripting.Dictionary scrrun.dl isn't avaiable an error 429 is raised.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Implements IScriptingDictionary

'============================================='
'Types
'============================================='

Private Type TScriptingDictionary
    scriptDictionary As Object
End Type

'============================================= '
'Constants and Private Variables
'============================================= '
Const SCRIPTING_DICTIONARY As String = "Scripting.Dictionary"

'============================================='
'Private Variables
'============================================='
Private this As TScriptingDictionary

'============================================='
'Constructors and destructors
'============================================='
Private Sub Class_Initialize()
#If Not Mac Then
    Set this.scriptDictionary = CreateObject(SCRIPTING_DICTIONARY)
#Else
    VBA.Err.Raise 429, "Cannot Create Object as Scripting.Dictionary is not available on Mac"
#End If
End Sub

Private Sub Class_Terminate()
    Set this.scriptDictionary = Nothing
End Sub

'============================================= '
'Properties
'============================================= '
'@Description("Creates a new instance of a ScriptingDictionaryLateBinding.")
'@param compareMethod (Optional) The comparison mode for comparing string keys in a Dictionary object, default is vbBinaryCompare.
'@return (ScriptingDictionaryLateBinding)
Public Function Create(Optional ByVal compareMethod As VBA.VbCompareMethod = VBA.vbBinaryCompare) As ScriptingDictionaryLateBinding
Attribute Create.VB_Description = "Creates a new instance of a ScriptingDictionaryLateBinding."
    Dim newDictionary As ScriptingDictionaryLateBinding
    Set newDictionary = New ScriptingDictionaryLateBinding
    newDictionary.CompareMode = compareMethod
    Set Create = newDictionary
End Function

''
'@Description("Sets or returns the comparison mode for comparing keys in a Dictionary object.")
'------------------------------------------------------------'
Public Property Get CompareMode() As VBA.VbCompareMethod
Attribute CompareMode.VB_Description = "Sets or returns the comparison mode for comparing keys in a Dictionary object."
    CompareMode = this.scriptDictionary.CompareMode
End Property

''
'@Description Sets the comparison mode for comparing keys in a Dictionary object
'@Remarks The comparison mode cannot be changed once the dictionary contains items
'@param compareMethod Value representing the comparison mode used by functions such as StrComp.
'@Error 5 Invalid procedure call or argument
'         Raised for an invalid data type for compareMethod
'@Error 9 Subscript out of range
'         Raised for an invalid value for compareMethod
'------------------------------------------------------------'
Public Property Let CompareMode(ByVal compareMethod As VBA.VbCompareMethod)
    this.scriptDictionary.CompareMode = compareMethod
End Property

''
'@Description("Returns the number of key/item pairs in a Dictionary object.")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of key/item pairs in a Dictionary object."
    Count = this.scriptDictionary.Count
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
    Set NewEnum = this.scriptDictionary.Keys.[_NewEnum]
End Property

''
'@Description("Sets or returns the value of an item in a Dictionary object.")
'@Remarks If key is not found when attempting to return an existing item, a new key is created
'         and its corresponding item is left empty.
'         Is the default property
'@param key The key associated with the item being retrieved.
'@Error 5 Invalid procedure call or argument
'         Raised for invalid key data type
'------------------------------------------------------------'
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_Description = "Sets or returns the value of an item in a Dictionary object."
    Dim result As Variant
    result = Array(this.scriptDictionary.Item(Key)) 'The get Item property result is wrapped in the Array to avoid running twice
    If VBA.IsObject(result(0)) Then
        Set Item = result(0)
    Else
        Item = result(0)
    End If
End Property

''
'@Description Sets a scalar item value for a specified key in a Dictionary object
'@Remarks The syntax for setting an item is: dictionaryobject.Item(key) = newItem
'         If you try to set item to a nonexistent key, the key is added to the dictionary,
'         and the item is linked to it, a sort of "implicit add."
'@param key  The key associated with the item being retrieved or added.
'@param Item The new value associated with the specified key
'@Error 5    Invalid procedure call or argument
'            Raised for an invalid key data type
'@Error 450  Wrong number of arguments of invalid property assignment
'            Raised when an item is an object when expecting a scalar value
'------------------------------------------------------------'
Public Property Let Item(ByRef Key As Variant, ByRef Item As Variant)
    this.scriptDictionary.Item(Key) = Item
End Property

''
'@Description Sets the object value of an item in a Dictionary object.
'@Remarks The syntax for setting an item is: Set dictionaryobject.Item(key) = newItem
'         If you try to set item to a nonexistent key, the key is added to the dictionary,
'         and the item is linked to it, a sort of an "implicit add."
'@param key  Key associated with the item being retrieved or added.
'@param Item The new value associated with the specified key
'@Error 5    Invalid procedure call or argument
'            Raised for invalid key data type.
'@Error 424  Object required
'            Raised when the item is a scalar value when expecting an object
'------------------------------------------------------------'
Public Property Set Item(ByRef Key As Variant, ByRef Item As Variant)
    Set this.scriptDictionary.Item(Key) = Item
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
    this.scriptDictionary.Key(Key) = newKey
End Property

'============================================='
'Public Methods
'============================================='
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
    this.scriptDictionary.Add Key, Item
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
    Exists = this.scriptDictionary.Exists(Key)
End Function

''
'@Description("Returns an array of items in a dictionary object.")
'@return Variant Array of items
'------------------------------------------------------------'
Public Function Items() As Variant
Attribute Items.VB_Description = "Returns an array of items in a dictionary object."
    Items = this.scriptDictionary.Items
End Function

''
'@Description("Returns an array of all the keys in a dictionary object.")
'@returns {Variant} Array of keys
'------------------------------------------------------------'
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Returns an array of all the keys in a dictionary object."
    Keys = this.scriptDictionary.Keys
End Function

''
'@Description("Removes a key/item pair from a Dictionary object.")
'@param key  The key associated with the key/item pair to remove from the Dictionary object.
'@Error 5  Invalid procedure call or argument
'          Raised for invalid/unsupported key data type
'------------------------------------------------------------'
Public Sub Remove(ByRef Key As Variant)
Attribute Remove.VB_Description = "Removes a key/item pair from a Dictionary object."
    this.scriptDictionary.Remove Key
End Sub

''
'@Description("Removes all key,item pairs from a Dictionary object.")
'------------------------------------------------------------'
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Removes all key,item pairs from a Dictionary object."
    this.scriptDictionary.RemoveAll
End Sub

'============================================='
'Interfaces
'============================================='
'------------------------------------------------------------'
'IScriptingDictionary Properties
'------------------------------------------------------------'
Private Property Get IScriptingDictionary_CompareMode() As VBA.VbCompareMethod
    IScriptingDictionary_CompareMode = this.scriptDictionary.CompareMode
End Property
Private Property Let IScriptingDictionary_CompareMode(ByVal compareMethod As VBA.VbCompareMethod)
    this.scriptDictionary.CompareMode = compareMethod
End Property

Private Property Get IScriptingDictionary_Count() As Long
    IScriptingDictionary_Count = this.scriptDictionary.Count
End Property

Private Property Get IScriptingDictionary_NewEnum() As IUnknown
    Set IScriptingDictionary_NewEnum = this.scriptDictionary.Keys.[_NewEnum]
End Property

Private Property Get IScriptingDictionary_Item(ByRef Key As Variant) As Variant
    Dim result As Variant
    result = Array(this.scriptDictionary.Item(Key)) 'The get Item property result is wrapped in the Array to avoid running twice
    If IsObject(result(0)) Then
        Set IScriptingDictionary_Item = result(0)
    Else
        IScriptingDictionary_Item = result(0)
    End If
End Property
Private Property Let IScriptingDictionary_Item(ByRef Key As Variant, ByRef Item As Variant)
    this.scriptDictionary.Item(Key) = Item
End Property
Private Property Set IScriptingDictionary_Item(ByRef Key As Variant, ByRef Item As Object)
    Set this.scriptDictionary.Item(Key) = Item
End Property

Private Property Let IScriptingDictionary_Key(ByRef Key As Variant, ByRef newKey As Variant)
    this.scriptDictionary.Key(Key) = newKey
End Property

'------------------------------------------------------------'
'IScriptingDictionary Methods
'------------------------------------------------------------'
Private Sub IScriptingDictionary_Add(ByRef Key As Variant, ByRef Item As Variant)
    this.scriptDictionary.Add Key, Item
End Sub

Private Function IScriptingDictionary_Exists(ByRef Key As Variant) As Boolean
    IScriptingDictionary_Exists = this.scriptDictionary.Exists(Key)
End Function

Private Function IScriptingDictionary_Items() As Variant
    IScriptingDictionary_Items = this.scriptDictionary.Items
End Function

Private Function IScriptingDictionary_Keys() As Variant
    IScriptingDictionary_Keys = this.scriptDictionary.Keys
End Function

Private Sub IScriptingDictionary_Remove(Key As Variant)
    this.scriptDictionary.Remove Key
End Sub

Private Sub IScriptingDictionary_RemoveAll()
    this.scriptDictionary.RemoveAll
End Sub


