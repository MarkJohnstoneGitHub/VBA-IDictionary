Attribute VB_Name = "TestModuleManagedCharSafeArray"
'@Folder("VBA-IScriptingDictionary.Tests.Unit Testing")
Option Explicit

Private Sub TestManagedCharSafeArray()
    Dim managedChars() As Integer
    Dim managedCharsDescriptor As ManagedCharSafeArray
    
    Set managedCharsDescriptor = ManagedCharSafeArray.Create(managedChars)
    Dim text As String
    text = "ABCDabcd"
    managedCharsDescriptor.AllocateCharData text
    Dim index As Long
    Debug.Print "LBound: "; LBound(managedChars)
    Debug.Print "UBound: "; UBound(managedChars)
    For index = LBound(managedChars) To UBound(managedChars)
        Debug.Print managedChars(index)
    Next
    managedCharsDescriptor.Dispose
End Sub

'As the managed Char array is read-only i.e. locked if
'attempt to resize the array Runtime Error 10 is raised.
'Raises Runtime Error 10 "This array is fixed or temporary locked."
Private Sub TestManagedCharSafeArrayErrorResize()
On Error GoTo ErrorHandler
    Dim managedChars() As Integer
    Dim managedCharsDescriptor As ManagedCharSafeArray
    
    Set managedCharsDescriptor = ManagedCharSafeArray.Create(managedChars)
    Dim text As String
    text = "ABCDabcd"
    managedCharsDescriptor.AllocateCharData text
    Dim index As Long
    Debug.Print "LBound: "; LBound(managedChars)
    Debug.Print "UBound: "; UBound(managedChars)
    For index = LBound(managedChars) To UBound(managedChars)
        Debug.Print managedChars(index)
    Next
    
    ReDim managedChars(0 To 20) 'Raises Runtime Error 10 "This array is fixed or temporary locked."
CleanExit:
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number; Err.Description
    managedCharsDescriptor.Dispose
    Resume CleanExit
End Sub
