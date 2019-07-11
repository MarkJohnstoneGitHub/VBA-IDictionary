VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryKeyValuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A VBA alternative for the Scripting.Dictionary which is Mac compatible using the IDictionary interface.\r\n\r\nIDictionary v1.1 (July 11, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@Folder("VBA-IDictionary")

'@TODO For enumeration of the DictionaryKeyItemPair consider changing to on Keys for consistency.
'@TODO Or provide a property to choose what to enumerate on. i.e. Keys, Items or key,item pair

''
'@Version IDictionary v1.1 (July 11, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description A VBA alternative for the Scripting.Dictionary which is Mac compatible using the IDictionary interface.
'@Required IDictionary.cls, DictionaryKeyValuePair.cls
'@Usage   Eg. Dim myDictionary As DictionaryKeyValuePair
'             myDictionary = New DictionaryKeyValuePair
'         May be created as above or using the Dictionary.cls factory class using
'         the Dictionary.Create function which returns a IDictionary type.
'
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified July 11, 2019
'@Package IDictionary.cls
'         Dictionary.cls
'         ScriptingDictionary.cls
'         DictionaryKeyValuePair.cls
'
'@Remarks
'  Implemented using a VBA.Collection of key, value pairs which are added into
'  Collection using collectionObject.Add Item:=Array(Key, Item), Key:=encodedKey
'
'  The encodedKey consists of the following:
'  a) First character reserved to represent its data type.
'  b) Subsequent characters are the string conversion of the variant Key specified.
'  For String Keys and if the CompareMode is vbBinaryCompare they are
'  encoded to a unicode hex String, where each character from the variant Key is
'  converted into four hexidecimal characters.
'  Ascii encoding could be used thou undetermined results may occur for
'  non compatible Ascii String keys.  Ascii encoding, however has significant
'  performance verses unicode encoding.
'
'  Enumeration of the DictionaryKeyItemPair returns a variant which is a
'  one dimensional array where the first array item is the dictionary Key
'  and the second is the dictionary Item.
'
'  Added support for LongLong keys which the Scripting.Dictionary doesn't support.
'
'  Adding large datasets i.e. 100,000+ with the Key and/or Item containing an object
'  may result in a noticeable slow cleaning up of the dictionary when set to nothing.
'  This may appear to the user that the application has hanged, thou infact it's performing
'  VBA's object garbage collection.
'  This appears to be related to the VBA's garbage collection of objects for all data structures.
'
'@VersionNotes IDictionary v1.1 (July 11, 2019)
'  Added the Property TexEncodingMode
'  Added private method BytesToAsciiHexString
'  The default is unicode encoding. (Slower performance however should cater for most languages.)
'  The TexEncodingMode provides the option for increased performance for the encoding of
'  case sensItive string keys using ASCII encoding.
'  The calling code is responsibile for ensuring ASCII compatible string keys are provided or
'  indeterminate results may occur.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit
Implements IDictionary

'============================================='
'Types
'============================================='
'List of text encoding methods available
Public Enum TextEncodingMethod
    TextEncodingUnicode = 0
    TextEncodingAscii = 1
    [_First] = TextEncodingUnicode
    [_Last] = TextEncodingAscii
End Enum

'List of available IDictionary implementations
Private Type TDictionaryKeyValuePair
    keyItemPairs As VBA.Collection
    compareMethod As VBA.VbCompareMethod
    keyList() As Variant
    itemList() As Variant
    'For increased efficiency the status for repopulating the array Keys() and Items() are used only if a change has occurred.
    updateKeys As Boolean
    updateItems As Boolean
    mTextEncodingMethod As TextEncodingMethod 'The text encoding method used of case sensitive text keys. i.e. Unicode or Ascii
End Type

'============================================='
'Constants
'============================================='
'Key, value pairs are stored in the collection item using Array(Key,Item)
Const KEYVALUEPAIR_KEY = 0
Const KEYVALUEPAIR_VALUE = 1

'Used to reserve the first character of the encoded dictionary key for its data type.
Const DATATYPE_IDENTIFIER_LENGTH = 1
Const EMPTY_KEY As String = "0"
Const NULL_KEY As String = "1"
Const OBJECT_NOTHING_KEY As String = "2"
Const STRING_KEY As String = "5"
Const OBJECT_KEY As String = "6"
Const BOOLEAN_KEY As String = "7"
Const NUMBER_KEY As String = "9"    'reserved for the number types: vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte, vbDate types


'============================================='
'Private Variables
'============================================='
Private this As TDictionaryKeyValuePair

'============================================='
'Constructors and destructors
'============================================='
Private Sub Class_Initialize()
    Set this.keyItemPairs = New VBA.Collection
    this.compareMethod = VBA.vbBinaryCompare                                'The default text comparsion method used is case sensitive.
    this.mTextEncodingMethod = TextEncodingMethod.TextEncodingUnicode  'The default text encoding method is unicode.
    this.updateKeys = False
    this.updateItems = False
End Sub

Private Sub Class_Terminate()
    Set this.keyItemPairs = Nothing
End Sub

'============================================='
'Properties
'============================================='

''
'@Description("Sets or returns the comparison mode for comparing string keys.")
' ------------------------------------------------------------'
Public Property Get CompareMode() As VBA.VbCompareMethod
Attribute CompareMode.VB_Description = "Sets or returns the comparison mode for comparing string keys."
    CompareMode = this.compareMethod
End Property

''
'@Description("Sets or returns the comparison mode for comparing string keys.")
'@param compareMethod Value representing the comparison mode used by functions such as StrComp.
'@Error 5  Invalid procedure call or argument.
'          Raised for changing the CompareMode for a dictionary that contains data.
'@Error 9  Subscript out of range.
'          Raised for an invalid value range for compareMethod.
'@Error 13 Type Mismatch.
'          Raised in the calling code for an invalid data type for compareMethod.
'@Remarks The comparison mode cannot be changed once the dictionary contains items.
' ------------------------------------------------------------'
Public Property Let CompareMode(ByVal compareMethod As VBA.VbCompareMethod)
Attribute CompareMode.VB_Description = "Sets or returns the comparison mode for comparing string keys."
    If Me.Count > 0 Then
        ' Can't change the CompareMode for a dictionary that contains data
        ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
        VBA.Err.Raise 5 ' Invalid procedure call or argument
    Else
        If (compareMethod = vbBinaryCompare) Or _
           (compareMethod = vbTextCompare) Or _
           (compareMethod = vbDatabaseCompare) _
        Then
            this.compareMethod = compareMethod
        Else
            ' illeagal value
            VBA.Err.Raise 9 '<- Subscript out of range
        End If
    End If
End Property

''
'@Description("Returns the number of key, item pairs in a Dictionary object.")
' ------------------------------------------------------------'
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of key, item pairs in a Dictionary object."
   Count = this.keyItemPairs.Count
End Property

''
'@Enumerator
'@Description("Enumerator for dictionary which returns a one dimensional variant array of a key,item pair.")
'@Remarks Attribute NewEnum.VB_UserMemId = -4 declares NewEnum as the enumerator
'         Attribute NewEnum.VB_MemberFlags = "40" declares NewEnum as a hidden property.
' ------------------------------------------------------------'
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator for dictionary which returns a one dimensional variant array of a key,item pair."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = this.keyItemPairs.[_NewEnum]
End Property

''
'@DefaultMember
'@Description("Set or returns the item for a specified key.")
'         The key and item can be a scalar value or an object.
'@param key  The key associated with the item being retrieved or added.
'@param Item The new value associated with the specified key.
'@Error 5    Invalid procedure call or argument.
'            Raised for invalid key data type from encoding the collection key.
'@Error 13   Type mismatch
'            Raised in calling code when an item is a scalar value when expecting an object.
'@Error 450  Wrong number of arguments or invalid property assignment.
'            Raised in calling code when an item is an object when expecting a scalar value.
'@Remarks
'  The syntax for setting an item is: dictionaryobject.Item(key) = newItem
'  If you try to set item to a nonexistent key, a new key, item pair is added
'  to the dictionary, and its associated item is left empty, a sort of "implicit add".
'  Attribute Item.VB_UserMemId = 0 Declares property Item as the default property.
' ------------------------------------------------------------'
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_Description = "Set or returns the item for a specified key."
Attribute Item.VB_UserMemId = 0
    Dim encodedKey As String
    encodedKey = EncodeKey(Key)
    'retrieving an array with key, item
    Dim keyValuePair As Variant
    GetKeyValuePair encodedKey, keyValuePair
    If Not IsEmpty(keyValuePair) Then
        If IsObject(keyValuePair(KEYVALUEPAIR_VALUE)) Then
            Set Item = keyValuePair(KEYVALUEPAIR_VALUE)
        Else
            Item = keyValuePair(KEYVALUEPAIR_VALUE)
        End If
    Else
        'If key is not found when attempting to return an existing item, a new key is created and its corresponding item value is left empty.
        this.keyItemPairs.Add Item:=Array(Key, Empty), Key:=encodedKey
        this.updateItems = True
        this.updateKeys = True
        Item = Empty ' Not found -> Returns Empty
    End If
End Property

''
'@Description("Set or returns the item for a specified key.")
'@param key  Key associated with the item being retrieved or added.
'@param Item The new value associated with the specified key.
'@Error 5    Invalid procedure call or argument.
'            Raised for an invalid key data type from encoding the collection key.
'@Error 450  Wrong number of arguments of invalid property assignment.
'            Raised in calling code when an item is an object when expecting a scalar value.
'@Remarks
' The syntax for setting an item is: dictionaryobject.Item(key) = newItem .
' If you try to set item to a nonexistent key, the key is added to the dictionary,
' and the item is linked to it, a sort of "implicit add".
' The key can be scalar or an object, the item must be a scalar value.
' ------------------------------------------------------------'
Public Property Let Item(ByRef Key As Variant, ByRef Item As Variant)
Attribute Item.VB_Description = "Set or returns the item for a specified key."
    'Item must be a scalar value only
    If Not IsObject(Item) Then
        Dim encodedKey As String
        encodedKey = EncodeKey(Key)
        'retrieving an array with key, item
        Dim keyValuePair As Variant
        GetKeyValuePair encodedKey, keyValuePair
        If IsEmpty(keyValuePair) Then
            this.keyItemPairs.Add Item:=Array(Key, Item), Key:=encodedKey
        Else
            'Note When a basic data type is stored in a Collection it is read-only, to update the item require to delete then readd
            this.keyItemPairs.Remove encodedKey
            this.keyItemPairs.Add Item:=Array(Key, Item), Key:=encodedKey
        End If
    Else
        'Raised when creating a key,item pair where a item is an object when expecting a scalar value.
        VBA.Err.Raise 450 ' Wrong number of arguments of invalid property assignment.
    End If
    this.updateItems = True
    this.updateKeys = True
End Property

''
'@Description("Set or returns the item for a specified key.")
'         The key can be scalar or an object, the item must be an object
'@param key  Key associated with the item being retrieved or added.
'@param Item The new value associated with the specified key.
'@Error 5    Invalid procedure call or argument.
'            Raised for invalid key data type from encoding the collection key.
'@Error 424  Object required.
'            Raised when creating a key,item pair where a item is a scalar value when expecting an object.
'@Remarks
' The syntax for setting an item is: Set dictionaryobject.Item(key) = newItem.
' If you try to set item to a nonexistent key, the key is added to the dictionary,
' and the item is linked to it, a sort of an "implicit add".
' ------------------------------------------------------------'
Public Property Set Item(ByRef Key As Variant, ByRef Item As Variant)
Attribute Item.VB_Description = "Set or returns the item for a specified key."
    'Item must be an object value only
    If IsObject(Item) Then
        Dim encodedKey As String
        encodedKey = EncodeKey(Key)
        'retrieving an array with key, item
        Dim keyValuePair As Variant
        GetKeyValuePair encodedKey, keyValuePair
        If IsEmpty(keyValuePair) Then
            'Add new object item for its associated key to the dictionary
            this.keyItemPairs.Add Item:=Array(Key, Item), Key:=encodedKey
        Else
            'Update dictionary object item for its existing associated key
            'Note When a basic data type is stored in a Collection it is read-only, to update the item require to delete then readd
            this.keyItemPairs.Remove encodedKey
            this.keyItemPairs.Add Item:=Array(Key, Item), Key:=encodedKey
        End If
    Else
        'Raised when creating a key,item pair where a item is a scalar value when expecting an object.
        VBA.Err.Raise 424 ' Object required.
    End If
    this.updateItems = True
    this.updateKeys = True
End Property

''
'@Description("Changes the key of an existing key, value pair to a new key value.")
'@param key    The existing key value being changed.
'@param newKey The new key value that replaces the specified key.
'@Error 5     Invalid procedure call or argumentv
'             Raised for invalid key data type from encoding the collection key.
'@Error 457   This key is already associated with an element of this collection.
'             Raised when new key already exists in the dictionary.
'@Error 32811 Application-defined or object-defined error.
'             Raised when the key specifed to be changed doesn't exist in the dictionary.
' ------------------------------------------------------------'
Public Property Let Key(ByRef Key As Variant, ByRef newKey As Variant)
Attribute Key.VB_Description = "Changes the key of an existing key, value pair to a new key value."
    Dim encodedKey As String
    encodedKey = EncodeKey(Key)
    'retrieving an array with a key, item pair when the first item is the key and second is the value
    Dim keyValuePair As Variant
    GetKeyValuePair encodedKey, keyValuePair
    If Not IsEmpty(keyValuePair) Then
        Dim newEncodedKey As String
        newEncodedKey = EncodeKey(newKey)
        If Not EncodedKeyExists(newEncodedKey) Then
            'Cannont update collection keys must remove and add back the key,item pair
            this.keyItemPairs.Remove encodedKey
            this.keyItemPairs.Add Item:=Array(newKey, keyValuePair(KEYVALUEPAIR_VALUE)), Key:=newEncodedKey
            this.updateItems = True
            this.updateKeys = True
        Else
            'Error raised when the new key specified already exists in the dictionary.
            'This key is already associated with an element of this collection.
            VBA.Err.Raise 457
        End If
    Else
        'Error raised when the key specified to be changed doesn't exist in the dictionary.
        VBA.Err.Raise 32811     'Application-defined or object-defined error
    End If
End Property

''
'@Description("Sets or returns the text encoding mode for encoding case senstive string keys.")
' ------------------------------------------------------------'
Public Property Get TextEncodingMode() As TextEncodingMethod
Attribute TextEncodingMode.VB_Description = "Sets or returns the text encoding mode for encoding case senstive string keys."
    TextEncodingMode = this.mTextEncodingMethod
End Property
''
'@Description("Sets or returns the text encoding mode for encoding case senstive string keys.")
'@param textEncodingMethod Value representing the text encoding method used i.e. Unicode or Ascii.
'@Error 5  Invalid procedure call or argument.
'          Raised for changing the  textEncodingMethod for a dictionary that contains items.
'@Error 9  Subscript out of range.
'          Raised for an invalid value range for textEncodingMethod.
'@Error 13 Type Mismatch.
'          Raised in the calling code for an invalid data type for textEncodingMethod.
'@Remarks The text encoding mode cannot be changed once the dictionary contains items.
' ------------------------------------------------------------'
Public Property Let TextEncodingMode(ByVal keyEncodingMethod As TextEncodingMethod)
    If Me.Count > 0 Then
        ' Can't change the TextEncodingMode for a dictionary that contains data
        VBA.Err.Raise 5 ' Invalid procedure call or argument
    Else
        If (keyEncodingMethod = TextEncodingMethod.TextEncodingUnicode) Or _
           (keyEncodingMethod = TextEncodingMethod.TextEncodingAscii) _
        Then
            this.mTextEncodingMethod = keyEncodingMethod
        Else
            ' illeagal value
            VBA.Err.Raise 9 '<- Subscript out of range
        End If
    End If
End Property

'============================================='
'Public Methods
'============================================='

''
'@Description("Adds a new key, item pair to the dictionary.")
'@param key  The key associated with the item being added.
'@param Item The newItem associated with the key being added.
'@Error 5    Invalid procedure call or argument
'            Raised for invalid key data type from encoding the collection key.
'@Error 457  This key is already associated with an element of this collection.
'            Raised when the specified new key already exists in the dictionary.
'@Remarks The key and item can be a scalar value or an object.
' ------------------------------------------------------------'
Public Sub Add(ByRef Key As Variant, ByRef Item As Variant)
Attribute Add.VB_Description = "Adds a new key, item pair to the dictionary."
    Dim encodedKey As String
    encodedKey = EncodeKey(Key)
    this.keyItemPairs.Add Item:=Array(Key, Item), Key:=encodedKey
    this.updateItems = True
    this.updateKeys = True
End Sub

''
'@Description("Returns True if a specified key exists in the dictionary; False if it does not.")
'@param Key The key value being searched for in the Dictionary object.
'@return {Boolean}
'@Error 5  Invalid procedure call or argument.
'          Raised for invalid/unsupported key data type from encoding the collection key.
' ------------------------------------------------------------'
Public Function Exists(ByRef Key As Variant) As Boolean
Attribute Exists.VB_Description = "Returns True if a specified key exists in the dictionary; False if it does not."
    Dim keyExists As Boolean
    keyExists = False
    Dim encodedKey As String
    encodedKey = EncodeKey(Key)
    If EncodedKeyExists(encodedKey) Then
        keyExists = True
    End If
    Exists = keyExists
End Function

''
'@Description("Returns an array of all items in the dictionary.")
'@return (Variant) Array of items.
'@Remarks Repopulates items array only if changes to items are made.
' ------------------------------------------------------------'
Public Function Items() As Variant
Attribute Items.VB_Description = "Returns an array of all items in the dictionary."
    If Me.Count > 0 Then
        If this.updateItems Then
            Erase this.itemList
            ReDim this.itemList(0 To this.keyItemPairs.Count - 1)
            Dim keyValuePair As Variant
            Dim itemsIndex As Long
            itemsIndex = 0
            For Each keyValuePair In this.keyItemPairs
                If IsObject(keyValuePair(KEYVALUEPAIR_VALUE)) Then
                    Set this.itemList(itemsIndex) = keyValuePair(KEYVALUEPAIR_VALUE)
                Else
                    this.itemList(itemsIndex) = keyValuePair(KEYVALUEPAIR_VALUE)
                End If
                itemsIndex = itemsIndex + 1
            Next keyValuePair
        End If
        Items = this.itemList
    Else
        Items = VBA.Split(vbNullString)
    End If
    this.updateItems = False
End Function

''
'@Description("Returns an array of all keys in the dictionary.")
'@return {Variant} Array of keys.
'@Remarks Repopulates the keys array only if changes to keys are made.
' ------------------------------------------------------------'
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Returns an array of all keys in the dictionary."
    If Me.Count > 0 Then
        If this.updateKeys Then
            Erase this.keyList
            ReDim this.keyList(0 To this.keyItemPairs.Count - 1)
            Dim keyValuePair As Variant
            Dim keysIndex As Long
            keysIndex = 0
            For Each keyValuePair In this.keyItemPairs
                If IsObject(keyValuePair(KEYVALUEPAIR_KEY)) Then
                    Set this.keyList(keysIndex) = keyValuePair(KEYVALUEPAIR_KEY)
                Else
                    this.keyList(keysIndex) = keyValuePair(KEYVALUEPAIR_KEY)
                End If
                keysIndex = keysIndex + 1
            Next keyValuePair
            
        End If
        Keys = this.keyList
    Else
        Keys = VBA.Split(vbNullString)
    End If
    this.updateKeys = False
End Function

''
'@Description("Removes a key,item pair from the dictionary for the key specified.")
'@param key The key associated with the key/item pair to remove from the dictionary.
'@Error 5     Invalid procedure call or argument.
'             Raised for invalid/unsupported key data type from encoding the key.
'@Error 32811 Method 'Remove' of object IDictionary failed
'             Raised when key doesn't exist to be removed
' ------------------------------------------------------------'
Public Sub Remove(ByRef Key As Variant)
Attribute Remove.VB_Description = "Removes a key,item pair from the dictionary for the key specified."
    Dim encodedKey As String
    encodedKey = EncodeKey(Key)
    RemoveKeyItemPair encodedKey
    this.updateKeys = True
    this.updateItems = True
End Sub

''
'@Description("Removes all key, item pairs from the dictionary.")
' ------------------------------------------------------------'
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Removes all key, item pairs from the dictionary."
    Set this.keyItemPairs = New VBA.Collection
    Erase this.keyList
    Erase this.itemList
    this.updateItems = True
    this.updateKeys = True
End Sub

'============================================='
'Private Methods
'============================================='

''
'@Description Returns True if a specified key exists in the collection object; False if it does not.
'@param encodedKey Endcoded string key value being searched for in the collection object.
'@returns (Boolean) True if a specified key exists in the collection object.
' ------------------------------------------------------------'
Private Function EncodedKeyExists(ByRef encodedKey As String) As Boolean
    On Error GoTo ErrorHandler
    this.keyItemPairs.Item encodedKey
    EncodedKeyExists = True
CleanExit:
    Exit Function
ErrorHandler:
    EncodedKeyExists = False
    Err.Clear 'reset error trapping.
    Resume CleanExit
End Function

''
'@Description Returns a string collection key for a variant dictionary key.
'@param key The key to be encoded into a string collection key.
'@returns (String) The encoded string collection key.
'@Error 5     Invalid procedure call or argument.
'             Raised for arrays or an unsupported data type for a key.
'@Error 457   This key is already associated with an element of this collection.
'             Attempting to add dictionary keys of different numeric data types with the same value results in Error 457.
'@Remarks
'  The encoded string key consists of the following:
'  a) First character reserved to represent its data type.
'  b) Subsequent characters are the string conversion of the variant Key specified.
'
'  For String dictionary Keys and if the CompareMode is vbBinaryCompare they are
'  encoded to a unicode hex String, where each character from the variant Key is
'  converted into four hexidecimal characters.
'  Ascii encoding could be used thou undetermined results may occur for
'  non Ascii compatible String Keys.  Ascii encoding, however has a significant
'  performance verses unicode encoding.
'
'  Numeric data types are all grouped together. i.e. Integer, Long, Single, Double, Currency, Byte, Date
'  Added support 64-Bit types such as LongLong for keys which the Scripting.Dictionary doesn't support.
'  LongPtr keys are converted by VBA to LongLong or Long which are is supported.
'  A dictionary key specified may be Null, Empty or Nothing, thou can only be added once as a key.
' ------------------------------------------------------------'
Private Function EncodeKey(ByRef Key As Variant) As String
    Dim keyVarType As VBA.VbVarType
    keyVarType = VarType(Key)
    Select Case keyVarType
        'The Scripting.Dictionary groups all numeric format data types together
        Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbByte, VBA.vbDate
            EncodeKey = NUMBER_KEY & VBA.str$(Key)
        'Added support for LongLong Scripting.Dictionary doesn't allow 64-Bit data types for keys
        #If VBA7 Then
        Case VBA.vbLongLong
            EncodeKey = NUMBER_KEY & VBA.str$(Key)
        #End If
        Case VBA.vbString
            EncodeKey = EncodeStringKey(Key)
        Case VBA.vbObject
            If Key Is Nothing Then
                EncodeKey = OBJECT_NOTHING_KEY
             Else
                EncodeKey = OBJECT_KEY & VBA.str$(ObjPtr(Key))
            End If
        Case VBA.vbBoolean
            EncodeKey = BOOLEAN_KEY & IIf(Key, "-1", "0")
        Case VBA.vbNull
            EncodeKey = NULL_KEY
        Case VBA.vbEmpty
            EncodeKey = EMPTY_KEY
        Case Else
            VBA.Err.Raise 5
    End Select
End Function

''
'@Description Encodes string Keys into encoded string collection keys according to the CompareMode and TextEncodingMode specified.
'@param key The string key to be encoded into a string collection key.
'@returns String The string collection key.
'@Remarks
'  For case sensitive string keys each character is converted into two or four hex characters
'  using either ascii or unicode according to the TextEncodingMode specified.
' ------------------------------------------------------------'
Private Function EncodeStringKey(ByRef Key As Variant) As String
    If this.compareMethod = vbBinaryCompare Then
        'Allow for upper and lower case string keys by encoding to a Unicode hex string
        Dim bytes() As Byte
        bytes = Key  'Assign Unicode string to unicode byte array.
        If this.mTextEncodingMethod = TextEncodingMethod.TextEncodingUnicode Then
            EncodeStringKey = STRING_KEY & BytesToUnicodeHexString(bytes) 'Each character is converted into four hex characters.
        Else
            EncodeStringKey = STRING_KEY & BytesToAsciiHexString(bytes)   'Each character is converted into two hex characters.
        End If
    Else
        EncodeStringKey = STRING_KEY & CStr(Key)
    End If
End Function

''
'@Description Returns an string of unicode hex characters from a byte array.
'@param Key The byte array to be encoded into an hex string.
'@returns (String) Unicode Hex characters.
' ------------------------------------------------------------'
Private Function BytesToUnicodeHexString(ByRef bytes() As Byte) As String
    Const CHAR_LENGTH As Long = 2
    'Allocate string for the required length of the hex string i.e. Each character is converted to 4 hex characters
    BytesToUnicodeHexString = VBA.Space$(CHAR_LENGTH * (UBound(bytes) - LBound(bytes) + 1))
    Dim byteIndex As Long
    For byteIndex = 0 To UBound(bytes)
        Mid$(BytesToUnicodeHexString, (byteIndex * CHAR_LENGTH) + 1, CHAR_LENGTH) = VBA.Right$(VBA.Hex(256 Or bytes(byteIndex)), CHAR_LENGTH)
    Next byteIndex
End Function

''
'@Description Returns an string of Ascii hex characters from a byte array.
'@param Key The byte array to be encoded into an hex string.
'@returns (String) Ascii Hex characters.
' ------------------------------------------------------------'
Private Function BytesToAsciiHexString(ByRef bytes() As Byte) As String
    Const CHAR_LENGTH As Long = 2
    'Allocate string for the required length of the hex string i.e. Each character is converted to 4 hex characters
    BytesToAsciiHexString = VBA.Space$(CHAR_LENGTH * (UBound(bytes) - LBound(bytes) + 1))
    Dim byteIndex As Long
    For byteIndex = 0 To UBound(bytes) Step 2
        Mid$(BytesToAsciiHexString, byteIndex + 1, CHAR_LENGTH) = VBA.Right$(VBA.Hex(256 Or bytes(byteIndex)), CHAR_LENGTH)
    Next byteIndex
End Function

''
'@Description Returns the key item pair for its associated key. If not found returns nothing.
'@param encodedKey      Endcoded string key value being searched for in the collection object.
'@param outKeyValuePair The key, value pair in variant array format for the associated key
' ------------------------------------------------------------'
Private Function GetKeyValuePair(ByRef encodedKey As String, ByRef outKeyValuePair As Variant)
    On Error GoTo ErrorHandler
    outKeyValuePair = this.keyItemPairs.Item(encodedKey)
CleanExit:
    Exit Function
ErrorHandler:
    Err.Clear 'reset error trapping.
    Resume CleanExit
End Function

''
'@Description Removes a Key Value Pair from the collection for it's associated key.
'@TODO Do you need to get the key value pair to destroy to avoid memory leaks?
'@param encodedKey Endcoded string key value to remove from the collection object.
'@Error 32811 Method 'Remove' of object IDictionary failed.
'             Raised when key doesn't exist to be removed.
' ------------------------------------------------------------'
Private Sub RemoveKeyItemPair(ByRef encodedKey As String)
    On Error GoTo ErrorHandler
    this.keyItemPairs.Remove encodedKey
CleanExit:
    Exit Sub
ErrorHandler:
    'Match Scripting.Dictionary error for when key doesn't exist when being removed
     VBA.Err.Raise 32811, "Method 'Remove' of object IDictionary failed"
     Resume CleanExit
End Sub

'============================================='
'Interfaces
'============================================='

' ------------------------------------------------------------'
'IDictionary Properties
' ------------------------------------------------------------'
Private Property Get IDictionary_CompareMode() As VBA.VbCompareMethod
    IDictionary_CompareMode = this.compareMethod
End Property
Private Property Let IDictionary_CompareMode(ByVal compareMethod As VBA.VbCompareMethod)
    Me.CompareMode = compareMethod
End Property

Private Property Get IDictionary_Count() As Long
    IDictionary_Count = this.keyItemPairs.Count
End Property

Private Property Get IDictionary_NewEnum() As IUnknown
    Set IDictionary_NewEnum = Me.NewEnum
End Property

Private Property Get IDictionary_Item(ByRef Key As Variant) As Variant
    Dim result As Variant
    result = Array(Me.Item(Key)) 'The get Item property result is wrapped in the Array to avoid running twice
    If IsObject(result(0)) Then
        Set IDictionary_Item = result(0)
    Else
        IDictionary_Item = result(0)
    End If
End Property
Private Property Let IDictionary_Item(ByRef Key As Variant, ByRef Item As Variant)
    Me.Item(Key) = Item
End Property
Private Property Set IDictionary_Item(ByRef Key As Variant, ByRef Item As Variant)
    Set Me.Item(Key) = Item
End Property

Private Property Let IDictionary_Key(ByRef Key As Variant, ByRef newKey As Variant)
    Me.Key(Key) = newKey
End Property

' ------------------------------------------------------------'
'IDictionary Methods
' ------------------------------------------------------------'
Private Sub IDictionary_Add(ByRef Key As Variant, ByRef Item As Variant)
    Me.Add Key, Item
End Sub

Private Function IDictionary_Exists(ByRef Key As Variant) As Boolean
    IDictionary_Exists = Me.Exists(Key)
End Function

Private Function IDictionary_Items() As Variant
    IDictionary_Items = Me.Items
End Function

Private Function IDictionary_Keys() As Variant
    IDictionary_Keys = Me.Keys
End Function

Private Sub IDictionary_Remove(Key As Variant)
    Me.Remove Key
End Sub

Private Sub IDictionary_RemoveAll()
    Me.RemoveAll
End Sub
