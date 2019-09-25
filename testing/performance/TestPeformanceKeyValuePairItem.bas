Attribute VB_Name = "TestPeformanceKeyValuePairItem"
'@Folder("VBA-IScriptingDictionary.Tests.Performance")
Option Explicit

'Change constants for the
Const NUMBER_OF_TEST As Long = 1000000 'Size of data set to test
Const REPORT_TEST As Long = 50000     'Frequency of reporting of performance i.e every 30,000 data items


Private Sub TestPerformanceKVPItemKeyLongItemLong()
    Dim sut As IScriptingDictionary
    Set sut = DictionaryPerformance.Create(IScriptingDictionaryPerformanceType.isdtDictionaryKeyValuePair)
    TestDictionaryKeyLongItemLong sut, NUMBER_OF_TEST, REPORT_TEST
End Sub


Private Sub TestPerformanceItemKeyLongItemLong()
    Dim dictionaryType As IScriptingDictionaryPerformanceType
    Dim sut As IScriptingDictionary
    For dictionaryType = IScriptingDictionaryPerformanceType.[_First] To IScriptingDictionaryPerformanceType.[_Last]
        Set sut = DictionaryPerformance.Create(dictionaryType, vbBinaryCompare)
        TestDictionaryKeyLongItemLong sut, NUMBER_OF_TEST, REPORT_TEST
    Next
End Sub

Private Sub TestDictionaryKeyLongItemLong(ByRef sut As IScriptingDictionary, ByVal numberOfItems As Long, ByVal reportFrequency As Long)

    'Heading
    Dim objectName As String
    objectName = TypeName(sut)
    Dim reportHeading As String
    reportHeading = objectName & ", dictObject.Item(keyLong) = itemLong" & ", Key:Long, Item:Long"
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
    
    
    'Adding to Dictionary using Item
    Dim startTime As Double
    Dim endTime As Double
    Dim currTime As Double
        
    Dim i As Long
    i = 0
    startTime = MicroTimer()
    Do While i < numberOfItems
        i = i + 1
        sut.Item(i) = i
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
