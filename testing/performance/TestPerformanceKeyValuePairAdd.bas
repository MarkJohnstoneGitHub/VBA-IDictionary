Attribute VB_Name = "TestPerformanceKeyValuePairAdd"
'@Folder("VBA-IScriptingDictionary.Tests.Performance")
'@TODO Create Array with test data and place into a collection or dictionary
'@TODO Feed into performance subs the test data which is an array of keys, items
'@TODO Place performance output into a collection to be possible used for analysis in Excel
'Three types to test performance numbers, string, objects
'Require: Key data type, Item data type, number of tests, number of reports (must be mod 0 of number of tests and < 100)


Option Explicit

'Change constants for the
Const NUMBER_OF_TEST As Long = 1000000 'Size of data set to test
Const REPORT_TEST As Long = 50000     'Frequency of reporting of performance i.e every 30,000 data items

'Peformance Testing subs

'Testing DictionaryKeyValuePair performance to see if changes produce benefits.
Private Sub TestComparePerformanceKeyStringItemLongBinaryCompareKVP()
    TestComparePerformanceKeyStringItemLongCompareMethodKVP VBA.vbBinaryCompare
End Sub

Private Sub TestComparePerformanceKeyLongItemString()
    Dim dictionaryType As IScriptingDictionaryPerformanceType
    
    Dim sut As IScriptingDictionary
    For dictionaryType = IScriptingDictionaryPerformanceType.[_First] To IScriptingDictionaryPerformanceType.[_Last]
        Set sut = DictionaryPerformance.Create(dictionaryType, vbBinaryCompare)
        TestDictionaryKeyLongItemString sut, NUMBER_OF_TEST, REPORT_TEST
    Next
End Sub

Private Sub TestPerformanceKeyLongItemLong()
    Dim dictionaryType As IScriptingDictionaryPerformanceType
    Dim sut As IScriptingDictionary
    For dictionaryType = IScriptingDictionaryPerformanceType.[_First] To IScriptingDictionaryPerformanceType.[_Last]
        Set sut = DictionaryPerformance.Create(dictionaryType, vbBinaryCompare)
        TestDictionaryKeyLongItemLong sut, NUMBER_OF_TEST, REPORT_TEST
    Next
End Sub

Private Sub TestComparePerformanceKeyStringItemLongTextCompare()
    TestComparePerformanceKeyStringItemLongCompareMethodV2 VBA.vbTextCompare
End Sub

Private Sub TestComparePerformanceKeyStringItemLongBinaryCompare()
    TestComparePerformanceKeyStringItemLongCompareMethodV2 VBA.vbBinaryCompare
End Sub

Private Sub TestPerformanceTextEncoding()
    Dim sut As IScriptingDictionary
    Set sut = DictionaryPerformance.Create(IScriptingDictionaryPerformanceType.isdtDictionaryKeyValuePair, VBA.vbBinaryCompare, temUnicode)
    TestDictionaryKeyStringItemLong sut, NUMBER_OF_TEST, REPORT_TEST
    Set sut = DictionaryPerformance.Create(IScriptingDictionaryPerformanceType.isdtDictionaryKeyValuePair, VBA.vbBinaryCompare, temAscii)
    TestDictionaryKeyStringItemLong sut, NUMBER_OF_TEST, REPORT_TEST
End Sub

'Warning not recommended to test for more then about 300,000 data items
'as takes approximately 30 secs to set each dictionary being tested to nothing.
'This appears to be due to VBA dereferencing of objects.
'Note doing four tests so will take at least 2 min to run for 300,000 data items.
Private Sub TestComparePerformanceKeyLongItemObject()
    Dim dictionaryType As IScriptingDictionaryPerformanceType
    For dictionaryType = IScriptingDictionaryPerformanceType.[_First] To IScriptingDictionaryPerformanceType.[_Last]
        TestDictionaryKeyLongItemObject NUMBER_OF_TEST, REPORT_TEST, dictionaryType
    Next
End Sub

'Warning not recommended to test for more then about 300,000 data items
'as takes approximately 30 secs to set each dictionary being tested to nothing.
'This appears to be due to VBA garbage collection of objects.
'Note doing four tests so will take at least 2 min to run for 300,000 data items.
Private Sub TestComparePerformanceKeyObjectItemObject()
    Dim dictionaryType As IScriptingDictionaryPerformanceType
    For dictionaryType = IScriptingDictionaryPerformanceType.[_First] To IScriptingDictionaryPerformanceType.[_Last]
        TestDictionaryKeyObjectItemObject NUMBER_OF_TEST, REPORT_TEST, dictionaryType
    Next
End Sub

'Warning not recommended to test for more then about 300,000 data items
'as takes approximately 30 secs to set each dictionary being tested to nothing.
'This appears to be due to VBA dereferencing of objects.
'Note doing four tests so will take at least 2 min to run for 300,000 data items.
Private Sub TestComparePerformanceKeyObjectItemLong()
    Dim dictionaryType As IScriptingDictionaryPerformanceType
    For dictionaryType = IScriptingDictionaryPerformanceType.[_First] To IScriptingDictionaryPerformanceType.[_Last]
        TestDictionaryKeyObjectItemLong NUMBER_OF_TEST, REPORT_TEST, dictionaryType
    Next
End Sub


'----------------------------------------------------------------------------------------------------
Private Sub TestComparePerformanceKeyStringItemLongCompareMethodV2(ByVal compareMethod As VBA.VbCompareMethod)
    Dim dictionaryType As IScriptingDictionaryPerformanceType
    
    Dim sut As IScriptingDictionary
    For dictionaryType = IScriptingDictionaryPerformanceType.[_First] To IScriptingDictionaryPerformanceType.[_Last]
        Set sut = DictionaryPerformance.Create(dictionaryType, compareMethod)
        TestDictionaryKeyStringItemLong sut, NUMBER_OF_TEST, REPORT_TEST
    Next
    
    If compareMethod = vbBinaryCompare Then
        Set sut = DictionaryPerformance.Create(IScriptingDictionaryPerformanceType.isdtDictionaryKeyValuePair, compareMethod, temAscii)
        TestDictionaryKeyStringItemLong sut, NUMBER_OF_TEST, REPORT_TEST
    End If
    
End Sub

Private Sub TestComparePerformanceKeyStringItemLongCompareMethodKVP(ByVal compareMethod As VBA.VbCompareMethod)
    Dim sut As IScriptingDictionary
    Set sut = DictionaryPerformance.Create(IScriptingDictionaryPerformanceType.isdtDictionaryKeyValuePair, compareMethod)
    TestDictionaryKeyStringItemLong sut, NUMBER_OF_TEST, REPORT_TEST
End Sub

Private Sub TestComparePerformanceKeyStringItemLongCompareMethod(ByVal compareMethod As VBA.VbCompareMethod)
    Dim dictionaryType As IScriptingDictionaryPerformanceType
    For dictionaryType = IScriptingDictionaryPerformanceType.[_First] To IScriptingDictionaryPerformanceType.[_Last]
        'TestDictionaryKeyStringItemLong NUMBER_OF_TEST, REPORT_TEST, dictionaryType, compareMethod
    Next
End Sub


Private Sub TestDictionaryKeyLongItemString(ByRef sut As IScriptingDictionary, ByVal numberOfItems As Long, ByVal reportFrequency As Long)
    Dim startTime As Double
    Dim endTime As Double
    Dim currTime As Double
        
    Dim objectName As String
    objectName = TypeName(sut)
    
    If sut.CompareMode = vbBinaryCompare Then
        Debug.Print objectName & " Key:Long, Item:String" & ", Binary Compare"
    Else
        Debug.Print objectName & " Key:Long, Item:String" & ", Text Compare"
    End If
        
    Dim i As Long
    i = 0
    startTime = MicroTimer()
    Do While i < numberOfItems
        i = i + 1
        sut.Add i, "abcdefghijklmnopqrstuvwxy" & Str$(i)
        If i Mod reportFrequency = 0 Then
             currTime = MicroTimer()
             Debug.Print i & " , " & Round(currTime - startTime, 3) & " , " & Round((i / (currTime - startTime)), 1)
        End If
    Loop
    endTime = MicroTimer()
    Debug.Print "ADD: " & numberOfItems & " items in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((sut.Count / (endTime - startTime)), 1) & " per second"

    startTime = MicroTimer()
    Set sut = Nothing
    endTime = MicroTimer()
    Debug.Print "Nothing: " & numberOfItems & " items removed in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((NUMBER_OF_TEST / (endTime - startTime)), 1) & " per second"
    Debug.Print
End Sub

Private Sub TestDictionaryKeyLongItemLong(ByRef sut As IScriptingDictionary, ByVal numberOfItems As Long, ByVal reportFrequency As Long)
    Dim startTime As Double
    Dim endTime As Double
    Dim currTime As Double
        
    Dim objectName As String
    objectName = TypeName(sut)
    
    If sut.CompareMode = vbBinaryCompare Then
        Debug.Print objectName & " Key:Long, Item:Long" & ", Binary Compare"
    Else
        Debug.Print objectName & " Key:Long, Item:Long" & ", Text Compare"
    End If
        
    Dim i As Long
    i = 0
    startTime = MicroTimer()
    Do While i < numberOfItems
        i = i + 1
        sut.Add i, i
        If i Mod reportFrequency = 0 Then
             currTime = MicroTimer()
             Debug.Print i & " , " & Round(currTime - startTime, 3) & " , " & Round((i / (currTime - startTime)), 1)
        End If
    Loop
    endTime = MicroTimer()
    Debug.Print "ADD: " & numberOfItems & " items in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((sut.Count / (endTime - startTime)), 1) & " per second"

    startTime = MicroTimer()
    Set sut = Nothing
    endTime = MicroTimer()
    Debug.Print "Nothing: " & numberOfItems & " items removed in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((NUMBER_OF_TEST / (endTime - startTime)), 1) & " per second"
    Debug.Print
End Sub


Private Sub TestDictionaryKeyStringItemLong(ByRef sut As IScriptingDictionary, ByVal numberOfItems As Long, ByVal reportFrequency As Long)

    'Heading
    Dim objectName As String
    objectName = TypeName(sut)
    Dim reportHeading As String
    reportHeading = objectName & " Key:Sring, Item:Long"
    If sut.CompareMode = vbBinaryCompare Then
        reportHeading = reportHeading & ", Binary Compare"
    Else
        reportHeading = reportHeading & ", Text Compare"
    End If
        
    If TypeOf sut Is DictionaryKeyValuePair Then
        Dim dictKeyValuePair As DictionaryKeyValuePair
        Set dictKeyValuePair = sut
        
        If dictKeyValuePair.TextEncodingMode = temUnicode Then
            reportHeading = reportHeading & ", Encoding: Unicode"
        Else
            reportHeading = reportHeading & ", Encoding: Ascii"
        End If
    End If
    Debug.Print reportHeading

    'Adding to Dictionary
    Dim startTime As Double
    Dim endTime As Double
    Dim currTime As Double
    
    Dim i As Long
    i = 0
    startTime = MicroTimer()
    Do While i < numberOfItems
        i = i + 1
        sut.Add "abcdefghijklmnopqrstuvwxy" & Str$(i), i
        If i Mod reportFrequency = 0 Then
             currTime = MicroTimer()
             Debug.Print i & " , " & Round(currTime - startTime, 3) & " , " & Round((i / (currTime - startTime)), 1)
        End If
    Loop
    endTime = MicroTimer()
    Debug.Print "ADD: " & numberOfItems & " items in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((sut.Count / (endTime - startTime)), 1) & " per second"

    
    'Setting Dictionary to Nothing
    startTime = MicroTimer()
    Set sut = Nothing
    endTime = MicroTimer()
    Debug.Print "Nothing: " & numberOfItems & " items removed in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((NUMBER_OF_TEST / (endTime - startTime)), 1) & " per second"
    Debug.Print
End Sub

Private Sub TestDictionaryKeyLongItemObject(ByVal numberOfItems As Long, ByVal reportFrequency As Long, Optional ByVal dictType As IScriptingDictionaryPerformanceType = IScriptingDictionaryPerformanceType.isdtScriptingDictionary, Optional ByVal compareMethod As VBA.VbCompareMethod = VBA.vbBinaryCompare)
    Dim startTime As Double
    Dim endTime As Double
    Dim currTime As Double
    
    Dim sut As IScriptingDictionary
    Set sut = DictionaryPerformance.Create(dictType)
    sut.CompareMode = compareMethod
    
    Dim objectName As String
    objectName = TypeName(sut)
    
    If compareMethod = vbBinaryCompare Then
        Debug.Print objectName & " Key:Long, Item: Object" & ", Binary Compare"
    Else
        Debug.Print objectName & " Key:Long, Item: Object" & ", Text Compare"
    End If

    Dim i As Long
    i = 0
    startTime = MicroTimer()
    Do While i < numberOfItems
        i = i + 1
        sut.Add i, Customer.Create("Mark", #8/4/1971#, 3182)
        If i Mod reportFrequency = 0 Then
             currTime = MicroTimer()
             Debug.Print i & " , " & Round(currTime - startTime, 3) & " , " & Round((i / (currTime - startTime)), 1)
        End If
    Loop
    endTime = MicroTimer()
    Debug.Print "ADD: " & numberOfItems & " items in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((sut.Count / (endTime - startTime)), 1) & " per second"
    
    Debug.Print "Setting Dictionary to Nothing."
    startTime = MicroTimer()
    Set sut = Nothing
    endTime = MicroTimer()
    Debug.Print "Nothing: " & numberOfItems & " items removed in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((NUMBER_OF_TEST / (endTime - startTime)), 1) & " per second"
    Debug.Print
End Sub


Private Sub TestDictionaryKeyObjectItemObject(ByVal numberOfItems As Long, ByVal reportFrequency As Long, Optional ByVal dictType As IScriptingDictionaryPerformanceType = IScriptingDictionaryPerformanceType.isdtScriptingDictionary, Optional ByVal compareMethod As VBA.VbCompareMethod = VBA.vbBinaryCompare)
    Dim startTime As Double
    Dim endTime As Double
    Dim currTime As Double
    
    Dim sut As IScriptingDictionary
    Set sut = DictionaryPerformance.Create(dictType)
    sut.CompareMode = compareMethod
    
    Dim objectName As String
    objectName = TypeName(sut)
    
    If compareMethod = vbBinaryCompare Then
        Debug.Print objectName & " Key:Object, Item: Object" & ", Binary Compare"
    Else
        Debug.Print objectName & " Key:Object, Item: Object" & ", Text Compare"
    End If

    Dim i As Long
    Dim custObject As Customer
    i = 0
    startTime = MicroTimer()
    Do While i < numberOfItems
        i = i + 1
        Set custObject = Customer.Create("Mark", #8/4/1971#, 3182)
        sut.Add custObject, custObject
        If i Mod reportFrequency = 0 Then
             currTime = MicroTimer()
             Debug.Print i & " , " & Round(currTime - startTime, 3) & " , " & Round((i / (currTime - startTime)), 1)
        End If
    Loop
    endTime = MicroTimer()
    Debug.Print "ADD: " & numberOfItems & " items in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((sut.Count / (endTime - startTime)), 1) & " per second"
    
    Debug.Print "Setting Dictionary to Nothing."
    startTime = MicroTimer()
    Set sut = Nothing
    endTime = MicroTimer()
    Debug.Print "Nothing: " & numberOfItems & " items removed in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((NUMBER_OF_TEST / (endTime - startTime)), 1) & " per second"
    Debug.Print
End Sub

Private Sub TestDictionaryKeyObjectItemLong(ByVal numberOfItems As Long, ByVal reportFrequency As Long, Optional ByVal dictType As IScriptingDictionaryPerformanceType = IScriptingDictionaryPerformanceType.isdtScriptingDictionary, Optional ByVal compareMethod As VBA.VbCompareMethod = VBA.vbBinaryCompare)
    Dim startTime As Double
    Dim endTime As Double
    Dim currTime As Double
    
    Dim sut As IScriptingDictionary
    Set sut = DictionaryPerformance.Create(dictType)
    sut.CompareMode = compareMethod
    
    Dim objectName As String
    objectName = TypeName(sut)
    
    If compareMethod = vbBinaryCompare Then
        Debug.Print objectName & " Key:Object, Item: Long" & ", Binary Compare"
    Else
        Debug.Print objectName & " Key:Object, Item: Long" & ", Text Compare"
    End If

    Dim i As Long
    i = 0
    startTime = MicroTimer()
    Do While i < numberOfItems
        i = i + 1
        sut.Add Customer.Create("Mark", #8/4/1971#, 3182), i
        If i Mod reportFrequency = 0 Then
             currTime = MicroTimer()
             Debug.Print i & " , " & Round(currTime - startTime, 3) & " , " & Round((i / (currTime - startTime)), 1)
        End If
    Loop
    
    endTime = MicroTimer()
    Debug.Print "ADD: " & numberOfItems & " items in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((sut.Count / (endTime - startTime)), 1) & " per second"
    
    Debug.Print "Setting Dictionary to Nothing."
    startTime = MicroTimer()
    Set sut = Nothing
    endTime = MicroTimer()
    Debug.Print "Nothing: " & numberOfItems & " items removed in " & Round(endTime - startTime, 3) & " seconds" & " At:" & Round((NUMBER_OF_TEST / (endTime - startTime)), 1) & " per second"
    Debug.Print
End Sub





