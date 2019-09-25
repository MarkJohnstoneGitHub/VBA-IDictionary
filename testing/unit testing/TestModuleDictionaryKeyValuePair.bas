Attribute VB_Name = "TestModuleDictionaryKeyValuePair"
Option Explicit
Option Private Module

'@TestModule
'@Folder("VBA-IScriptingDictionary.Tests.Unit Testing")

'@TODO Check that all errors raised are exactly the same for all IScriptingDictionary implementations

'@Version IScriptingDictionary v2.0 (July 28, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description Test module for unit testing the DictionaryKeyValuePair using the Rubberduck addin
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified July 28, 2019
'@Dependencies
'   The Rubberduck addin, See http://rubberduckvba.com/
'   IScriptingDictionary.cls
'   Dictionary.cls
'   DictionaryKeyValuePair.cls
'   ScriptingDictionary.cls
'   ITextEncoding.cls
'   TextEncoderUnicode.cls
'   TextEncoderASCII.cls
'   For Testing also
'       ArrayFunctions.bas
'       Customer.cls
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'============================================='
'Methods Dictionary.Create
'============================================='

'@TestMethod("Dictionary Create")
Private Sub TestDictionaryCreateErrorCannotCreateObject()
    Const ExpectedError As Long = 429
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New Dictionary
        
    'Act:
    dict.CompareMode = "A"
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'============================================='
'Methods DictionaryKeyValuePair.Create
'============================================='

'@TestMethod("Methods Create")
Private Sub TestMethodsCreate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As DictionaryKeyValuePair
    
    Set dict = DictionaryKeyValuePair.Create(vbBinaryCompare, temUnicode)
    
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
    dict.Add "E", Array(1, 2, 3)
    dict.Add "F", dict
        
    'Act:

    'Assert:
    Assert.IsTrue dict("A") = 123
    Assert.IsTrue dict("B") = 3.14
    Assert.IsTrue dict("C") = "ABC"
    Assert.IsTrue dict("D") = True
    Assert.IsTrue dict("E")(1) = 2
    Assert.IsTrue dict("F")("C") = "ABC"
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Methods Create")
Private Sub TestMethodsDictionaryCreate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    
    Set dict = Dictionary.Create(IScriptingDictionaryType.isdtDictionaryKeyValuePair, vbTextCompare, temAscii)
    
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
    dict.Add "E", Array(1, 2, 3)
    dict.Add "F", dict
        
    'Act:

    'Assert:
    Assert.IsTrue dict("A") = 123
    Assert.IsTrue dict("B") = 3.14
    Assert.IsTrue dict("C") = "ABC"
    Assert.IsTrue dict("D") = True
    Assert.IsTrue dict("E")(1) = 2
    Assert.IsTrue dict("F")("C") = "ABC"
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'============================================='
'Property Count
'============================================='

'@TestMethod("Property Count")
Private Sub TestPropertyCount()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    
    'Act:

    'Assert:
    Dim Actual As Long
    Actual = 3
    Assert.IsTrue dict.Count = Actual
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'============================================='
'Property CompareMode
'============================================='

'The default CompareMode is VBA.vbBinaryCompare i.e 0
'@TestMethod("Property CompareMode")
Private Sub TestPropertyCompareModeDefault()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
        
    'Act:

    'Assert:
    Assert.IsTrue dict.CompareMode = VBA.vbBinaryCompare
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Property CompareMode")
Private Sub TestPropertyCompareModeBinaryCompare()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.CompareMode = 0
    dict.Add "A", 123
    dict("a") = 456
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
        
    'Act:

    'Assert:
    Assert.IsTrue dict.CompareMode = VBA.vbBinaryCompare
    Assert.IsTrue dict("A") = 123
    Assert.IsTrue dict("a") = 456
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property CompareMode")
Private Sub TestPropertyCompareModeTextCompare()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair

    dict.CompareMode = 1
    dict.Add "A", 123
    dict("a") = 456
    dict.Add "B", 3.14
    dict.Add "C", "ABC"

    'Act:

    'Assert:
    Assert.IsTrue dict.CompareMode = VBA.vbTextCompare
    Assert.IsTrue dict("A") = 456
    Assert.IsTrue dict("a") = 456
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property CompareMode")
Private Sub TestPropertyCompareModeErrors()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
        
    'Act:
    dict.CompareMode = vbTextCompare
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Property CompareMode")
Private Sub TestPropertyCompareModeErrorsSubscriptOutOfRange()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
        
    'Act:
    dict.CompareMode = 3
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Property CompareMode")
Private Sub TestPropertyCompareModeErrorsTypeMismatch()
    Const ExpectedError As Long = 13
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
        
    'Act:
    dict.CompareMode = "A"
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'============================================='
'Property Item Get
'============================================='

'@TestMethod("Property Item Get")
Private Sub TestPropertyItemGetByKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    
    'Act:

    'Assert:
    Assert.IsTrue dict.Item("B") = 3.14
    Assert.IsTrue dict.Item("D") = VBA.vbEmpty
    Assert.IsTrue dict("B") = 3.14
    Assert.IsTrue dict("D") = VBA.vbEmpty
    Assert.IsTrue dict.Count = 4
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property Item Get")
Private Sub TestPropertyItemGet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    dict.Item("A") = 123
    'Act:
    Dim dictItem As Variant
    dictItem = dict.Item("A")
    
    Dim dictItemB As Variant
    dictItemB = dict("A")
    
    'Assert:
    Assert.IsTrue dictItem = 123
    Assert.IsTrue dictItemB = 123

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property Item Get")
'@Aim To test getting an item that doesn't exist which creates a new key, item pair with an Empty item.
Private Sub TestPropertyItemGetNotExistsAddEmptyItem()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    'Act:
    Dim dictItem As Variant
    dictItem = dict.Item("A")
    
    'Assert:
    Assert.IsTrue dict.Item("A") = Empty

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'============================================='
'Errors Property Item Get
'============================================='
'@Error 5    Invalid procedure call or argument.
'            Raised for invalid key data type from encoding the collection key.
'@Error 13   Type mismatch
'            Raised in calling code when an item is a scalar value when expecting an object.
'@Error 450  Wrong number of arguments or invalid property assignment.
'            Raised in calling code when an item is an object value when expecting a scalar value.

'Aim Test Item Get for error 5 Invalid procedure call or argument.
'@Error 5    Invalid procedure call or argument
'            Raised for invalid key data type from encoding the collection key
'@TestMethod("Property Item Get")
Private Sub TestPropertyItemGetErrorsInvalidProcedureCall()
    Const ExpectedError As Long = 5 'Invalid procedure call or argument
    On Error GoTo TestFail
    
    Dim invalidDataTypeArray() As String
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Item("A") = 123
    
    Assert.IsTrue IsArray(invalidDataTypeArray)
    
    'Act:
    Dim dictItem As Variant
    dictItem = dict.Item(invalidDataTypeArray)
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 13 Type mismatch
'          Raised when an item is a scalar value when expecting an object from calling code.
'@TestMethod("Property Item Get")
Private Sub TestPropertyItemGetErrorsTypeMismatch()
    Const ExpectedError As Long = 13 'Object required.
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    Dim itemScalar As String
    itemScalar = "ABC"
    dict.Item("A") = itemScalar
    
    'Act:
    Dim dictItem As Variant
    Set dictItem = dict.Item("A") 'Item("A") contains a string scalar value
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 450  Wrong number of arguments or invalid property assignment.
'            Raised when an item is an object value when expecting a scalar value.
'@TestMethod("Property Item Get")
Private Sub TestPropertyItemGetErrorsInvalidProperty()
    Const ExpectedError As Long = 450 'Object required.
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    Dim itemObject As VBA.Collection
    Set itemObject = New VBA.Collection
    Set dict.Item("A") = itemObject
    'Act:
    Dim dictItem As Variant
    dictItem = dict.Item("A") 'Item("A") contains the object Collection which requires Set
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'============================================='
'Property Item Let
'============================================='

'@TestMethod("Property Item Let")
Private Sub TestPropertyItemLetByKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"

    'Act:
    ' Let + New
    dict("D") = True
        
    ' Let + Replace
    dict("A") = 456
    dict("B") = 3.14159

    'Assert:
    Assert.IsTrue dict("A") = 456
    Assert.IsTrue dict("B") = 3.14159
    Assert.IsTrue dict("C") = "ABC"
    Assert.IsTrue dict("D") = True
    Assert.IsTrue dict.Count = 4
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property Item Let")
Private Sub TestPropertyItemLetKeyVariant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair

    Dim dictKeyA As Variant
    dictKeyA = "A"
    dict(dictKeyA) = 123

    Dim dictKeyB As Variant
    dictKeyB = "B"
    Set dict(dictKeyB) = New DictionaryKeyValuePair

    'Act:

    'Assert:
    Assert.IsTrue dict(dictKeyA) = 123
    Assert.IsTrue dict(dictKeyB).Count = 0
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'============================================='
'Errors Property Item Let
'============================================='
'Errors Raised
'@Error 5    Invalid procedure call or argument
'            Raised for an invalid key data type
'@Error 450  Wrong number of arguments of invalid property assignment
'            Raised when an item is an object when expecting a scalar value


'@Error 5    Invalid procedure call or argument
'            Raised for invalid key data type from encoding the collection key
'@TestMethod("Property Item Let")
Private Sub TestPropertyItemLetErrorsInvalidProcedureCall()
    Const ExpectedError As Long = 5 'Invalid procedure call or argument
    On Error GoTo TestFail
    
    Dim invalidDataTypeArray() As String
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    Assert.IsTrue IsArray(invalidDataTypeArray)
    
    'Act:
    dict.Item(invalidDataTypeArray) = 123
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 450  Wrong number of arguments of invalid property assignment
'            Raised when an item is an object when expecting a scalar value
'@TestMethod("Property Item Let")
Private Sub TestPropertyItemLetErrorsNeItemInvalidProperty()
    Const ExpectedError As Long = 450 'Wrong number of arguments of invalid property assignment
    On Error GoTo TestFail
    
    Dim invalidDataTypeArray() As String
    Dim itemInvalidObject As Collection
    
    Set itemInvalidObject = New Collection
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    'Act:
    '@Ignore ObjectVariableNotSet
    dict.Item("A") = itemInvalidObject
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 450  Wrong number of arguments of invalid property assignment
'            Raised when an item is an object when expecting a scalar value
'@TestMethod("Property Item Let")
Private Sub TestPropertyItemLetErrorsUpdateItemInvalidProperty()
    Const ExpectedError As Long = 450 'Wrong number of arguments of invalid property assignment
    On Error GoTo TestFail
    
    Dim itemInvalidObject As Collection
    
    Set itemInvalidObject = New Collection
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Item("A") = 123
    
    'Act:
    
    '@Ignore ObjectVariableNotSet
    dict.Item("A") = itemInvalidObject
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'============================================='
'Property Item Set
'============================================='


'============================================='
'Errors Property Item Set
'============================================='

'@Error 424  Object required
'            Raised when an item is an scalar when expecting an object
'@TestMethod("Property Item Set")
Private Sub TestPropertyItemSetErrorsNewItemObjectRequired()
    Const ExpectedError As Long = 424 'Object required
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    Dim invalidScalarItem As Variant
    invalidScalarItem = 123

    'Act:
    Set dict.Item("A") = invalidScalarItem
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 424  Object required
'            Raised when an item is an scalar when expecting an object
'TestMethod("Property Item Set")
Private Sub TestPropertyItemSetErrorsNewItemObjectRequiredScripting()
    Const ExpectedError As Long = 424 'Object required
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New ScriptingDictionary
    Dim invalidScalarItem As Variant
    invalidScalarItem = 123

    'Act:
    Set dict.Item("A") = invalidScalarItem
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'============================================='
'Property Key Let
'============================================='

'@TestMethod("Property Key")
Private Sub TestPropertyKeyScalar()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
        
    'Act:
    dict.Key("B") = "PI"

    'Assert:
    Assert.IsTrue dict("PI") = 3.14
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property Key")
Private Sub TestPropertyKeyObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    
    Dim myCustomer As Customer
    Set myCustomer = Customer.Create("Bob", #1/1/2000#, 9999)
    dict.Add myCustomer, "ABC"
    
    'Act:
    Dim myCustomerNew As Customer
    Set myCustomerNew = Customer.Create("Mary", #5/5/1999#, 9999)

    dict.Key(myCustomer) = myCustomerNew
    
    'Assert:
    Assert.IsTrue dict(myCustomerNew) = "ABC"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'============================================='
'Property Key Let Errors
'============================================='
'@Error 5     Invalid procedure call or argument
'             Raised for invalid key data type from encoding the collection key
'@Error 457   This key is already associated with an element of this collection
'             Raised when new key already exists in the dictionary
'@Error 32811 Application-defined or object-defined error
'             Raised when the key specifed to be changed doesn't exist in the dictionary


'@Error 5     Invalid procedure call or argument
'             Raised for invalid key data type from encoding the collection key
'@TestMethod("Property Key")
Private Sub TestPropertyKeyLetErrorsInvalidArgument()
    Const ExpectedError As Long = 5 'Invalid procedure call or argument
    On Error GoTo TestFail

    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    
    Dim invalidKeyArray() As Variant

    'Act:
    dict.Key(invalidKeyArray) = "EFG"
    'Arrange:
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 457   This key is already associated with an element of this collection
'             Raised when new key already exists in the dictionary
'@TestMethod("Property Key")
Private Sub TestPropertyKeyLetErrorsKeyAlreadyExists()
    Const ExpectedError As Long = 457 'This key is already associated with an element of this collection
    On Error GoTo TestFail

    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
        
    'Arrange:
    'Act:
    dict.Key("A") = "C"
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 32811 Application-defined or object-defined error
'             Raised when the key specifed to be changed doesn't exist in the dictionary
'@TestMethod("Property Key")
Private Sub TestPropertyKeyLetErrorsKeyNotExists()
    Const ExpectedError As Long = 32811 'Application-defined or object-defined error
    On Error GoTo TestFail

    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
        
    'Arrange:
    'Act:
    dict.Key("D") = "E"
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'============================================='
'Methods Add
'============================================='

'@TestMethod("Methods Add")
Private Sub TestMethodsAdd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As DictionaryKeyValuePair
    Set dict = New DictionaryKeyValuePair
    
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
    dict.Add "E", Array(1, 2, 3)
    dict.Add "F", dict
        
    'Act:

    'Assert:
    Assert.IsTrue dict("A") = 123
    Assert.IsTrue dict("B") = 3.14
    Assert.IsTrue dict("C") = "ABC"
    Assert.IsTrue dict("D") = True
    Assert.IsTrue dict("E")(1) = 2
    Assert.IsTrue dict("F")("C") = "ABC"
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Methods Add")
Private Sub TestMethodAddSetItem()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair

    
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
        
    'Act:
    ' Set + New
    Set dict("D") = New DictionaryKeyValuePair
    
    Dim dictItem As DictionaryKeyValuePair
    Set dictItem = dict.Item("D")
    
    dict.Item("D").Add "key", "D"
    

    ' Set + Replace
    Set dict("A") = New DictionaryKeyValuePair
    
    
    dict("A").Add "key", "A"
    Set dict("B") = New DictionaryKeyValuePair
    dict("B").Add "key", "B"

    'Assert:
    Assert.IsTrue dict.Item("A")("key") = "A"
    Assert.IsTrue dict.Item("B")("key") = "B"
    Assert.IsTrue dict("C") = "ABC"
    Assert.IsTrue dict.Item("D")("key") = "D"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Methods Add")
'Aim: Test adding numeric keys and string keys with the same "value"
Private Sub TestMethodsAddKeyNumeric()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair

    dict.Add 3, 1
    dict.Add 2, 2
    dict.Add 1, 3
    dict.Add "3", 4
    dict.Add "2", 5
    dict.Add "1", 6

    'Act:

    'Assert:
    Assert.IsTrue dict(3) = 1
    Assert.IsTrue dict(1) = 3
    Assert.IsTrue dict("3") = 4
    Assert.IsTrue dict("2") = 5
    Assert.IsTrue dict("1") = 6
    
    Assert.IsTrue VarType(dict.Keys()(0)) = VBA.vbInteger
    Assert.IsTrue VarType(dict.Keys()(1)) = VBA.vbInteger
    Assert.IsTrue VarType(dict.Keys()(2)) = VBA.vbInteger
    
    Assert.IsTrue dict.Count = 6
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Methods Add")
'Aim: Test adding numeric keys with the same integer value and with/without decimals
Private Sub TestMethodsAddKeyNumericWithDecimals()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    Dim dictKeyDouble As Double
    dictKeyDouble = 1.23
    
    dict.Add Key:=1, Item:="Item 1"
    dict.Add Key:=1.1, Item:="Item 1.1"
    dict.Add Key:=dictKeyDouble, Item:=dictKeyDouble

    'Act:

    'Assert:
    Assert.IsTrue dict(1) = "Item 1"
    Assert.IsTrue dict(1.1) = "Item 1.1"
    Assert.IsTrue dict(dictKeyDouble) = dictKeyDouble
    
    Assert.IsTrue dict.Count = 3
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Methods Add")
Private Sub TestMethodsAddKeyBoolean()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add True, 1
    dict.Add False, 2
        
    'Act:

    'Assert:
    Assert.IsTrue dict(True) = 1
    Assert.IsTrue dict(False) = 2
    Assert.IsTrue VarType(dict.Keys()(0)) = VBA.vbBoolean
    Assert.IsTrue VarType(dict.Keys()(1)) = VBA.vbBoolean
    
    Assert.IsTrue dict.Count = 2
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Methods Add")
Private Sub TestMethodsAddKeyDate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    Dim testDateKey As Date
    testDateKey = #3/3/2019#
    
    dict.Add testDateKey, 1
    dict.Add #10/3/2019#, 2
    dict.Add #3/10/2019#, 3
    'Act:

    'Assert:
    Assert.IsTrue dict(testDateKey) = 1
    Assert.IsTrue dict(#10/3/2019#) = 2
    Assert.IsTrue dict(#3/10/2019#) = 3
    
    Assert.IsTrue VarType(dict.Keys()(0)) = VBA.vbDate
    Assert.IsTrue VarType(dict.Keys()(1)) = VBA.vbDate
    Assert.IsTrue VarType(dict.Keys()(2)) = VBA.vbDate
    Assert.IsTrue dict.Count = 3
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Methods Add")
Private Sub TestMethodsAddKeyObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    Dim a As Collection
    Dim B As IScriptingDictionary
    Set a = New Collection
    Set B = New DictionaryKeyValuePair
    a.Add 123
    B.Add "a", 456
    dict.Add a, "123"
    dict.Add B, "456"
        
    'Act:

    'Assert:
    Assert.IsTrue dict(a) = "123"
    Assert.IsTrue dict(B) = "456"
    
    dict.Remove B
    dict.Key(a) = B
    Assert.IsTrue dict(B) = "123"
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Methods Add")
'Default CompareMode is case sensitive
Private Sub TestMethodsAddKeyStringCaseSensitive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    Dim TestStringKey As String
    
    TestStringKey = "mark"
    
    dict.Add TestStringKey, 1
    dict.Add "Bob", 2
    dict.Add "Mark", 3
    'Act:

    'Assert:
    Assert.IsTrue dict(TestStringKey) = 1
    Assert.IsTrue dict("Bob") = 2
    Assert.IsTrue dict("Mark") = 3
    
    Assert.IsTrue VarType(dict.Keys()(0)) = VBA.vbString
    Assert.IsTrue VarType(dict.Keys()(1)) = VBA.vbString
    Assert.IsTrue VarType(dict.Keys()(2)) = VBA.vbString
    Assert.IsTrue dict.Count = 3
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Methods Add")
Private Sub TestMethodsAddKeyStringCaseInsensitive()
    Const ExpectedError As Long = 457 'This key is already associated with an element of this collection
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    dict.CompareMode = vbTextCompare
    Dim TestStringKey As String
    TestStringKey = "mark"
    dict.Add TestStringKey, 1
    dict.Add "Bob", 2

'Act:
    dict.Add "Mark", 3 'Raises Error 457 as key already exists in the dictionary

Assert:
    Assert.IsTrue dict(TestStringKey) = 1
    Assert.IsTrue dict("Bob") = 2
    Assert.IsTrue dict.Exists(TestStringKey)
    Assert.IsTrue dict.Exists("mArK")
    
    Assert.IsTrue VarType(dict.Keys()(0)) = VBA.vbString
    Assert.IsTrue VarType(dict.Keys()(1)) = VBA.vbString
    Assert.IsTrue dict.Count = 2
        
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume Assert
    Else
        Assert.Fail "Expected error was not raised"
        Resume TestExit
    End If
End Sub


'@Descripting Test adding key with a value that is Null, Empty and Nothing
'@TestMethod("Methods Add")
Private Sub TestMethodsAddKeyNullEmptyNothing()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add Null, 1
    dict.Add Empty, 2
    dict.Add Nothing, 3
        
    'Act:

    'Assert:
    Assert.IsTrue dict(Null) = 1
    Assert.IsTrue dict(Empty) = 2
    Assert.IsTrue dict(Nothing) = 3
    Assert.IsTrue dict.Count = 3
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'============================================='
'Methods Add Errors
'============================================='

'@Error 5    Invalid procedure call or argument
'            Raised for invalid key data type from encoding the collection key
'@TestMethod("Methods Add")
Private Sub TestMethodsAddErrorsInvalidArgument()
    Const ExpectedError As Long = 5 'Invalid procedure call or argument
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict  As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    Dim invalidKeyArray() As Variant

    'Act:
    dict.Add invalidKeyArray, 123
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 457  This key is already associated with an element of this collection
'            Raised when the specified new key already exists in the dictionary
'@TestMethod("Methods Add")
Private Sub TestMethodsAddErrorsKeyExists()
    Const ExpectedError As Long = 457 'This key is already associated with an element of this collection
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict  As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    dict.Add "A", 123
    'Act:
    dict.Add "A", 456
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'============================================='
'Methods Exists
'============================================='

'@TestMethod("Methods Exists")
Private Sub TestMethodsExists()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "Exists", 123
        
    'Act:

    'Assert:
    Assert.IsTrue dict.Exists("Exists") = True
    Assert.IsTrue dict.Exists("Doesn't Exist") = False
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'============================================='
'Methods Exists Errors
'============================================='

'@Error 5   Invalid procedure call or argument
'           Raised for invalid/unsupported key data type from encoding the collection key
'@TestMethod("Methods Exists")
Private Sub TestMethodsExistsInvalidArgument()
    Const ExpectedError As Long = 5 'Invalid procedure call or argument
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict  As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    Dim invalidKeyArray() As Variant
    Dim keyExists As Boolean
    
    'Act:
    keyExists = dict.Exists(invalidKeyArray)
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'============================================='
'Methods Items
'============================================='

'@TestMethod("Methods Items")
Private Sub TestMethodsItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    'Assert:
    Assert.IsTrue IsArray(dict.Items)
    Assert.IsTrue IsArrayEmpty(dict.Items)
    
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
     
    Dim dictItems As Variant
    dictItems = dict.Items

    Assert.IsTrue UBound(dictItems) = 3
    Assert.IsTrue dictItems(0) = 123
    Assert.IsTrue dictItems(3) = True

    dict.Remove "A"
    dict.Remove "B"
    dict.Remove "C"
    dict.Remove "D"
    
    'Assert:
    Assert.IsTrue IsArray(dict.Items)
    Assert.IsTrue IsArrayEmpty(dict.Items)
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Methods Items")
Private Sub TestMethodsItemsEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    'Act:

    'Assert:
    Assert.IsTrue IsArray(dict.Items)
    Assert.IsTrue IsArrayEmpty(dict.Items)
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'============================================='
'Methods Keys
'============================================='

'@TestMethod("Methods Keys")
Private Sub TestMethodsKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    'Assert:
    Assert.IsTrue IsArray(dict.Keys)
    Assert.IsTrue IsArrayEmpty(dict.Keys)
    
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
     
    Dim dictKeys As Variant
    dictKeys = dict.Keys

    'Assert:
    Assert.IsTrue UBound(dictKeys) = 3
    Assert.IsTrue dictKeys(0) = "A"
    Assert.IsTrue dictKeys(3) = "D"

    dict.RemoveAll
    
    'Assert:
    Assert.IsTrue IsArray(dict.Keys)
    Assert.IsTrue IsArrayEmpty(dict.Keys)
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Methods Keys")
Private Sub TestMethodsKeysEnumerate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    Dim keysCollection As Collection
    Set keysCollection = New Collection
    Dim dictKey As Variant
    For Each dictKey In dict.Keys
        keysCollection.Add dictKey
    Next dictKey
    
    'Assert:
    Assert.IsTrue keysCollection.Count = 0
    
    'Arrange:
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
     
    Set keysCollection = New Collection
    For Each dictKey In dict.Keys
        keysCollection.Add dictKey
    Next dictKey
    
    'Assert:
    Assert.IsTrue keysCollection.Count = 4
    Assert.IsTrue keysCollection(1) = "A"
    Assert.IsTrue keysCollection(4) = "D"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'============================================='
'Methods Remove
'============================================='

'@TestMethod("Methods Remove")
Private Sub TestMethodsRemoveCaseSensitiveKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.CompareMode = vbTextCompare
    
    'Assert:
    Assert.IsTrue IsArray(dict.Keys)
    Assert.IsTrue IsArrayEmpty(dict.Keys)
    
    'Arrange:
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
     
    'Assert:
    Assert.IsTrue dict.Count = 4
    dict.Remove "C"
    Assert.IsTrue dict.Count = 3
    
    dict.Remove "b" 'Case sensitive key removed
    Assert.IsTrue dict.Count = 2
        
    Assert.IsFalse dict.Exists("C") = True
    Assert.IsFalse dict.Exists("b") = True
          
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'============================================='
'Methods Remove Errors
'============================================='
'@Error 5     Invalid procedure call or argument
'             Raised for invalid/unsupported key data type from encoding the key
'@Error 32811 Method 'Remove' of object IScriptingDictionary failed
'             Raised when key doesn't exist to be removed


'@Error 5     Invalid procedure call or argument
'             Raised for invalid/unsupported key data type from encoding the key
'@TestMethod("Methods Remove")
Private Sub TestMethodsRemoveErrorsInvalidArgument()
    Const ExpectedError As Long = 5 'Invalid procedure call or argument
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict  As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
    
    Dim invalidKeyArray() As Variant

    'Act:
    dict.Remove invalidKeyArray
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 32811 Method 'Remove' of object IScriptingDictionary failed
'             Raised when key doesn't exist to be removed
'@TestMethod("Methods Remove")
Private Sub TestMethodsRemoveErrorsKeyNotExists()
    Const ExpectedError As Long = 32811 'Method 'Remove' of object IScriptingDictionary failed
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict  As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
    
    'Act:
    dict.Remove "E"
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@Error 32811 Method 'Remove' of object IScriptingDictionary failed
'             Raised when key doesn't exist to be removed
'@TestMethod("Methods Remove")
Private Sub TestMethodsRemoveErrorsCaseSensitiveKeyNotExists()
    Const ExpectedError As Long = 32811 'Method 'Remove' of object IScriptingDictionary failed
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict  As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    dict.CompareMode = vbBinaryCompare
    
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
    
    'Act:
    dict.Remove "d"
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'============================================='
'Methods RemoveAll
'============================================='

'@TestMethod("Methods RemoveAll")
Private Sub TestMethodsRemoveAll()
    On Error GoTo TestFail
    
    'Arrange:
    Dim dict As IScriptingDictionary
    Set dict = New DictionaryKeyValuePair
    
    dict.Add "A", 123
    dict.Add "B", 3.14
    dict.Add "C", "ABC"
    dict.Add "D", True
     
    'Assert:
    Assert.IsTrue dict.Count = 4
    dict.RemoveAll
    Assert.IsTrue dict.Count = 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




