VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryVBATools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Drop-in replacement for Scripting.Dictionary on Mac\r\n\r\nDictionary v1.4.0\r\n(c) Tim Hall - https://github.com/timhall/VBA-Dictionary\r\nAuthor: tim.hall.engr@gmail.com\r\nLicense: MIT (http://www.opensource.org/licenses/mit-license.php)\r\n"
'@ModuleDescription "Drop-in replacement for Scripting.Dictionary on Mac\r\n\r\nDictionary v1.4.0\r\n(c) Tim Hall - https://github.com/timhall/VBA-Dictionary\r\nAuthor: tim.hall.engr@gmail.com\r\nLicense: MIT (http://www.opensource.org/licenses/mit-license.php)\r\n"
'@Folder("VBA-IScriptingDictionary.Tests.Performance")

''
'Dictionary v1.4.1
'(c) Tim Hall - https://github.com/timhall/VBA-Dictionary
'
'Drop-in replacement for Scripting.Dictionary on Mac
'@OriginalAuthor: tim.hall.engr@gmail.com
'@license: MIT (http://www.opensource.org/licenses/mit-license.php
'
'@Usage Eg. Dim myDictionary As DictionaryVBATools
'           myDictionary = New DictionaryVBATools
'
'@ModifiedBy Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Modificatons
'    Separated the custom dictionary code and code related to the Scripting.Dictionary
'    Removed all code related to the Scripting.Dictionary and moved the functionality to ScriptingDictionary.cls
'    Removed the compiler constants which were moved to Dictionary.cls and ScriptingDictionary.cls
'    Added a IScriptingDictionary interface
'    Moved class member variables to type TDictionaryVBATools and using the private member this of type TDictionaryVBATools
'    Removed CompareMethod type and changed to use VBA.VbCompareMethod.
'@LastModified July 28, 2019
'
'@Remarks
' Used only for testing for performance of DictionaryKeyValuePair compared to the following IScriptingDictionary implementations:
' A modified version of Tim Hall 's Dictionary at https://github.com/timhall/VBA-Dictionary see DictionaryVBATools.cls
' Microsoft Scripting Runtime Scripting early binding see ScriptingDictionary.cls
' Microsoft Scripting Runtime Scripting late binding see ScriptingDictionaryLateBinding.cls
' For testing performance purposes use the following IScriptingDictionary factory class:
' DictionaryFactoryTesting.cls
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Implements IScriptingDictionary

' --------------------------------------------- '
' Types
' --------------------------------------------- '
' dict_KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
Private Type TDictionaryVBATools
    keyValues As VBA.Collection
    dictKeys() As Variant
    dictItems() As Variant
    dictObjectKeys As VBA.Collection
    mCompareMode As VBA.VbCompareMethod
End Type

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '
Private this As TDictionaryVBATools

' --------------------------------------------- '
' Constructors and destructors
' --------------------------------------------- '
Private Sub Class_Initialize()
    Set this.keyValues = New VBA.Collection

    Erase this.dictKeys
    Erase this.dictItems
    Set this.dictObjectKeys = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Set this.keyValues = Nothing
    Set this.dictObjectKeys = Nothing
End Sub

' --------------------------------------------- '
' Properties
' --------------------------------------------- '
Public Property Get CompareMode() As VBA.VbCompareMethod
Attribute CompareMode.VB_Description = "Set or get the string comparison method."
    CompareMode = this.mCompareMode
End Property
Public Property Let CompareMode(ByVal compareMethod As VBA.VbCompareMethod)
Attribute CompareMode.VB_Description = "Set or get the string comparison method."
    If Me.Count > 0 Then
        ' Can't change CompareMode for Dictionary that contains data
        ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
        Err.Raise 5 ' Invalid procedure call or argument
    End If

    this.mCompareMode = compareMethod
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Get the number of items in the dictionary.\n"
    Count = this.keyValues.Count
End Property

''
'@Description("Enumerator for dictionary which returns variant array of key, item pair .")
'@Enumerator
' ------------------------------------------------------------'
'@TODO CHECK what is returned since changes
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator for dictionary which returns variant array of key, item pair ."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = Me.Keys.[_NewEnum]
End Property

Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_Description = "Set or get the item for a given key."
Attribute Item.VB_UserMemId = 0
    Dim dict_KeyValue As Variant
    dict_KeyValue = GetKeyValue(Key)

    If Not IsEmpty(dict_KeyValue) Then
        If VBA.IsObject(dict_KeyValue(2)) Then
            Set Item = dict_KeyValue(2)
        Else
            Item = dict_KeyValue(2)
        End If
    Else
        ' Not found -> Returns Empty
    End If
End Property
Public Property Let Item(ByRef Key As Variant, ByRef Item As Variant)
Attribute Item.VB_Description = "Set or get the item for a given key."
Attribute Item.VB_UserMemId = 0
    If Me.Exists(Key) Then
        ReplaceKeyValue GetKeyValue(Key), Key, Item
    Else
        AddKeyValue Key, Item
    End If
End Property
Public Property Set Item(ByRef Key As Variant, ByRef Item As Object)
Attribute Item.VB_Description = "Set or get the item for a given key."
Attribute Item.VB_UserMemId = 0
    If Me.Exists(Key) Then
        ReplaceKeyValue GetKeyValue(Key), Key, Item
    Else
        AddKeyValue Key, Item
    End If
End Property

Public Property Let Key(ByRef Previous As Variant, ByRef Updated As Variant)
Attribute Key.VB_Description = "Change a key to a different key."
    Dim dict_KeyValue As Variant
    dict_KeyValue = GetKeyValue(Previous)

    If Not VBA.IsEmpty(dict_KeyValue) Then
        ReplaceKeyValue dict_KeyValue, Updated, dict_KeyValue(2)
    End If
End Property

' --------------------------------------------- '
' Public Methods
' --------------------------------------------- '

''
' Add an item with the given key
'
' @param {Variant} Key
' @param {Variant} Item
' --------------------------------------------- '
Public Sub Add(ByRef Key As Variant, ByRef Item As Variant)
Attribute Add.VB_Description = "Add a new key and item to the dictionary."
    If Not Me.Exists(Key) Then
        AddKeyValue Key, Item
    Else
        ' This key is already associated with an element of this collection
        Err.Raise 457
    End If
End Sub

''
' Check if an item exists for the given key
'
' @param {Variant} Key
' @return {Boolean}
' --------------------------------------------- '
Public Function Exists(ByRef Key As Variant) As Boolean
Attribute Exists.VB_Description = "Determine if a given key is in the dictionary."
    Exists = Not IsEmpty(GetKeyValue(Key))
End Function

''
' Get an array of all items
'
' @return {Variant}
' --------------------------------------------- '
Public Function Items() As Variant
Attribute Items.VB_Description = "Get an array containing all items in the dictionary."
    If Me.Count > 0 Then
        Items = this.dictItems
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        Items = VBA.Split(vbNullString)
    End If
End Function

''
' Get an array of all keys
'
' @return {Variant}
' --------------------------------------------- '
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Get an array containing all keys in the dictionary."
    If Me.Count > 0 Then
        Keys = this.dictKeys
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        Keys = VBA.Split(vbNullString)
    End If
End Function

''
' Remove an item for the given key
'
' @param {Variant} Key
' --------------------------------------------- '
Public Sub Remove(ByRef Key As Variant)
Attribute Remove.VB_Description = "Remove a given key from the dictionary."
    Dim dict_KeyValue As Variant
    dict_KeyValue = GetKeyValue(Key)

    If Not VBA.IsEmpty(dict_KeyValue) Then
        RemoveKeyValue dict_KeyValue
    Else
        ' Application-defined or object-defined error
        Err.Raise 32811
    End If
End Sub

''
' Remove all items
' --------------------------------------------- '
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Remove all information from the dictionary."
    Set this.keyValues = New VBA.Collection

    Erase this.dictKeys
    Erase this.dictItems
End Sub

' ============================================= '
' Private Functions
' ============================================= '
Private Function GetKeyValue(ByRef dict_Key As Variant) As Variant
    On Error Resume Next
    GetKeyValue = this.keyValues(GetFormattedKey(dict_Key))
    Err.Clear
End Function

Private Sub AddKeyValue(ByRef dict_Key As Variant, ByRef dict_Value As Variant, Optional dict_Index As Long = -1)
    If Me.Count = 0 Then
        ReDim this.dictKeys(0 To 0)
        ReDim this.dictItems(0 To 0)
    Else
        ReDim Preserve this.dictKeys(0 To UBound(this.dictKeys) + 1)
        ReDim Preserve this.dictItems(0 To UBound(this.dictItems) + 1)
    End If

    Dim dict_FormattedKey As String
    dict_FormattedKey = GetFormattedKey(dict_Key)

    If dict_Index >= 0 And dict_Index < this.keyValues.Count Then
        ' Shift keys/items after + including index into empty last slot
        Dim dict_i As Long
        For dict_i = UBound(this.dictKeys) To dict_Index + 1 Step -1
            this.dictKeys(dict_i) = this.dictKeys(dict_i - 1)
            If VBA.IsObject(this.dictItems(dict_i - 1)) Then
                Set this.dictItems(dict_i) = this.dictItems(dict_i - 1)
            Else
                this.dictItems(dict_i) = this.dictItems(dict_i - 1)
            End If
        Next dict_i

        ' Add key/item at index
        this.dictKeys(dict_Index) = dict_Key
        If VBA.IsObject(dict_Value) Then
            Set this.dictItems(dict_Index) = dict_Value
        Else
            this.dictItems(dict_Index) = dict_Value
        End If

        ' Add key-value at proper index
        this.keyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey, Before:=dict_Index + 1
    Else
        ' Add key-value as last item
        If VBA.IsObject(dict_Key) Then
            Set this.dictKeys(UBound(this.dictKeys)) = dict_Key
        Else
            this.dictKeys(UBound(this.dictKeys)) = dict_Key
        End If
        If VBA.IsObject(dict_Value) Then
            Set this.dictItems(UBound(this.dictItems)) = dict_Value
        Else
            this.dictItems(UBound(this.dictItems)) = dict_Value
        End If

        this.keyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey
    End If
End Sub

Private Sub ReplaceKeyValue(ByRef dict_KeyValue As Variant, ByRef dict_Key As Variant, ByRef dict_Value As Variant)
    Dim dict_Index As Long
    Dim dict_i As Integer

    dict_Index = GetKeyIndex(dict_KeyValue(1))

    ' Remove existing dict_Value
    RemoveKeyValue dict_KeyValue, dict_Index

    ' Add new dict_Key dict_Value back
    AddKeyValue dict_Key, dict_Value, dict_Index
End Sub

Private Sub RemoveKeyValue(ByRef dict_KeyValue As Variant, Optional ByVal dict_Index As Long = -1)
    Dim dict_i As Long
    If dict_Index = -1 Then
        dict_Index = GetKeyIndex(dict_KeyValue(1))
    End If

    If dict_Index >= 0 And dict_Index <= UBound(this.dictKeys) Then
        ' Shift keys/items after index down
        For dict_i = dict_Index To UBound(this.dictKeys) - 1
            this.dictKeys(dict_i) = this.dictKeys(dict_i + 1)

            If VBA.IsObject(this.dictItems(dict_i + 1)) Then
                Set this.dictItems(dict_i) = this.dictItems(dict_i + 1)
            Else
                this.dictItems(dict_i) = this.dictItems(dict_i + 1)
            End If
        Next dict_i

        ' Resize keys/items to remove empty slot
        If UBound(this.dictKeys) = 0 Then
            Erase this.dictKeys
            Erase this.dictItems
        Else
            ReDim Preserve this.dictKeys(0 To UBound(this.dictKeys) - 1)
            ReDim Preserve this.dictItems(0 To UBound(this.dictItems) - 1)
        End If
    End If

    this.keyValues.Remove dict_KeyValue(0)
    RemoveObjectKey dict_KeyValue(1)
End Sub

Private Function GetFormattedKey(ByRef dict_Key As Variant) As String
    If VBA.IsObject(dict_Key) Then
        GetFormattedKey = GetObjectKey(dict_Key)
    ElseIf VarType(dict_Key) = VBA.vbBoolean Then
        GetFormattedKey = IIf(dict_Key, "-1__-1", "0__0")
    ElseIf VarType(dict_Key) = VBA.vbString Then
        GetFormattedKey = dict_Key

        If Me.CompareMode = VBA.vbBinaryCompare Then
            ' Collection does not have method of setting key comparison
            ' So case-sensitive keys aren't supported by default
            ' -> Approach: Append lowercase characters to original key
            '    AbC -> AbC___b_, abc -> abc__abc, ABC -> ABC_____
            Dim dict_Lowercase As String
            dict_Lowercase = vbNullString

            Dim dict_i As Integer
            Dim dict_Char As String
            Dim dict_Ascii As Integer
            For dict_i = 1 To VBA.Len(GetFormattedKey)
                dict_Char = VBA.Mid$(GetFormattedKey, dict_i, 1)
                dict_Ascii = VBA.Asc(dict_Char)
                If dict_Ascii >= 97 And dict_Ascii <= 122 Then
                    dict_Lowercase = dict_Lowercase & dict_Char
                Else
                    dict_Lowercase = dict_Lowercase & "_"
                End If
            Next dict_i

            If dict_Lowercase <> vbNullString Then
                GetFormattedKey = GetFormattedKey & "__" & dict_Lowercase
            End If
        End If
    Else
        ' For numbers, add duplicate to distinguish from strings
        ' ->  123  -> "123__123"
        '    "123" -> "123"
        GetFormattedKey = VBA.CStr(dict_Key) & "__" & CStr(dict_Key)
    End If
End Function

Private Function GetObjectKey(ByRef dict_ObjKey As Variant) As String
    Dim dict_i As Integer
    For dict_i = 1 To this.dictObjectKeys.Count
        If this.dictObjectKeys.Item(dict_i) Is dict_ObjKey Then
            GetObjectKey = "__object__" & dict_i
            Exit Function
        End If
    Next dict_i

    this.dictObjectKeys.Add dict_ObjKey
    GetObjectKey = "__object__" & this.dictObjectKeys.Count
End Function

Private Sub RemoveObjectKey(ByRef dict_ObjKey As Variant)
    Dim dict_i As Integer
    For dict_i = 1 To this.dictObjectKeys.Count
        If this.dictObjectKeys.Item(dict_i) Is dict_ObjKey Then
            this.dictObjectKeys.Remove dict_i
            Exit Sub
        End If
    Next dict_i
End Sub

Private Function GetKeyIndex(ByRef dict_Key As Variant) As Long
    Dim dict_i As Long
    For dict_i = 0 To UBound(this.dictKeys)
        If VBA.IsObject(this.dictKeys(dict_i)) And VBA.IsObject(dict_Key) Then
            If this.dictKeys(dict_i) Is dict_Key Then
                GetKeyIndex = dict_i
                Exit For
            End If
        ElseIf VBA.IsObject(this.dictKeys(dict_i)) Or VBA.IsObject(dict_Key) Then
            ' Both need to be objects to check equality, skip
        ElseIf this.dictKeys(dict_i) = dict_Key Then
            GetKeyIndex = dict_i
            Exit For
        End If
    Next dict_i
End Function

' --------------------------------------------- '
' Interface Properties
' --------------------------------------------- '
Private Property Get IScriptingDictionary_CompareMode() As VBA.VbCompareMethod
    IScriptingDictionary_CompareMode = Me.CompareMode
End Property
Private Property Let IScriptingDictionary_CompareMode(ByVal compareMethodValue As VBA.VbCompareMethod)
    Me.CompareMode = compareMethodValue
End Property

Private Property Get IScriptingDictionary_Count() As Long
    IScriptingDictionary_Count = Me.Count
End Property

Private Property Get IScriptingDictionary_NewEnum() As IUnknown
    Set IScriptingDictionary_NewEnum = Me.NewEnum
End Property

Private Property Get IScriptingDictionary_Item(ByRef Key As Variant) As Variant
    IScriptingDictionary_Item = Me.Item(Key)
End Property
Private Property Let IScriptingDictionary_Item(ByRef Key As Variant, ByRef Item As Variant)
    Me.Item(Key) = Item
End Property
Private Property Set IScriptingDictionary_Item(ByRef Key As Variant, ByRef Item As Object)
    Set Me.Item(Key) = Item
End Property

Private Property Let IScriptingDictionary_Key(ByRef Previous As Variant, ByRef Updated As Variant)
    Me.Key(Previous) = Updated
End Property

' --------------------------------------------- '
' Interface Methods
' --------------------------------------------- '
Private Sub IScriptingDictionary_Add(ByRef Key As Variant, Item As Variant)
    Me.Add Key, Item
End Sub

Private Function IScriptingDictionary_Exists(ByRef Key As Variant) As Boolean
    IScriptingDictionary_Exists = Me.Exists(Key)
End Function

Private Function IScriptingDictionary_Items() As Variant
    IScriptingDictionary_Items = Me.Items
End Function

Private Function IScriptingDictionary_Keys() As Variant
    IScriptingDictionary_Keys = Me.Keys
End Function

Private Sub IScriptingDictionary_Remove(ByRef Key As Variant)
    Me.Remove Key
End Sub

Private Sub IScriptingDictionary_RemoveAll()
    Me.RemoveAll
End Sub



