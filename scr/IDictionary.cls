VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "The IDictionary interface used for dictionary implementations.\r\n\r\nIDictionary v1.1 (July 12, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@ModuleDescription "The IDictionary interface used for dictionary implementations.\r\n\r\nIDictionary v1.1 (July 12, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@Folder("VBA-IDictionary")

''
'@Version IDictionary v1.1 (July 12, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description The IDictionary interface used for dictionary implementations.
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified July 12, 2019
'@Package IDictionary.cls
'         Dictionary.cls
'         ScriptingDictionary.cls
'         DictionaryKeyValuePair.cls
'------------------------------------------------------------'

Option Explicit

'============================================='
'Properties
'============================================='

''
'@Description("Sets or returns the comparison mode for comparing string keys.")
'------------------------------------------------------------'
Public Property Get CompareMode() As VBA.VbCompareMethod
Attribute CompareMode.VB_Description = "Sets or returns the comparison mode for comparing string keys."
End Property
''
'@Description("Sets or returns the comparison mode for comparing string keys.")
'------------------------------------------------------------'
Public Property Let CompareMode(ByVal compareMethod As VBA.VbCompareMethod)
Attribute CompareMode.VB_Description = "Sets or returns the comparison mode for comparing string keys."
End Property

''
'@Description("Returns the number of items in the dictionary.")
'------------------------------------------------------------'
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the dictionary."
End Property

''
'@Enumerator
'@Description("Enumerator for the dictionary object.")
'@Remarks Attribute NewEnum.VB_UserMemId = -4 declares NewEnum as the enumerator
'         Attribute NewEnum.VB_MemberFlags = "40" declares NewEnum as a hidden property.
' ------------------------------------------------------------'
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator for the dictionary object."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property

''
'@DefaultMember
'@Description("Set or returns the item for a specified key.")
'         Attribute Item.VB_UserMemId = 0 Declares property Item as the default property.
'@Remarks The syntax for setting an item is: dictionaryobject.Item(key) = newItem
'         If you try to set item to a nonexistent key, a new key, item pair is added
'         to the dictionary, and the item is linked to it, a sort of "implicit add".
'         Attribute Item.VB_UserMemId = 0 Declares property Item as the default property.
'------------------------------------------------------------'
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_Description = "Set or returns the item for a specified key."
Attribute Item.VB_UserMemId = 0
End Property
''
'@DefaultMember
'@Description("Set or returns the item for a specified key.")
'------------------------------------------------------------'
Public Property Let Item(ByRef Key As Variant, ByRef Item As Variant)
Attribute Item.VB_Description = "Set or returns the item for a specified key."
Attribute Item.VB_UserMemId = 0
End Property
''
'@DefaultMember
'@Description("Set or returns the item for a specified key.")
'------------------------------------------------------------'
Public Property Set Item(ByRef Key As Variant, ByRef Item As Object)
Attribute Item.VB_Description = "Set or returns the item for a specified key."
Attribute Item.VB_UserMemId = 0
End Property

''
'@Description("Changes the key of an existing key, value pair to a new key value.")
'------------------------------------------------------------'
Public Property Let Key(ByRef Key As Variant, ByRef newKey As Variant)
Attribute Key.VB_Description = "Changes the key of an existing key, value pair to a new key value."
End Property

'============================================='
'Public Methods
'============================================='

''
'@Description("Adds a new key, item pair to the dictionary.")
'@param Key  The key associated with the item being added.
'@param Item The newItem associated with the key being added.
' ------------------------------------------------------------'
Public Sub Add(ByRef Key As Variant, ByRef Item As Variant)
Attribute Add.VB_Description = "Adds a new key, item pair to the dictionary."
End Sub

''
'@Description("Returns True if a specified key exists in the dictionary; False if it does not.")
'@param Key The key value being searched for in the Dictionary object.
'@return {Boolean}
'------------------------------------------------------------'
Public Function Exists(ByRef Key As Variant) As Boolean
Attribute Exists.VB_Description = "Returns True if a specified key exists in the dictionary; False if it does not."
End Function

''
'@Description("Returns an array of all items in the dictionary.")
'@return {Variant} Array of keys.
' ------------------------------------------------------------'
Public Function Items() As Variant
Attribute Items.VB_Description = "Returns an array of all items in the dictionary."
End Function

''
'@Description("Returns an array of all keys in the dictionary.")
'@returns {Variant}
' ------------------------------------------------------------'
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Returns an array of all keys in the dictionary."
End Function

''
'@Description("Removes a key,item pair from the dictionary for the key specified.")
'@param {Variant} Key
'------------------------------------------------------------'
Public Sub Remove(ByRef Key As Variant)
Attribute Remove.VB_Description = "Removes a key,item pair from the dictionary for the key specified."
End Sub

''
'@Description("Removes all key, item pairs from the dictionary.")
'------------------------------------------------------------'
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Removes all key, item pairs from the dictionary."
End Sub


