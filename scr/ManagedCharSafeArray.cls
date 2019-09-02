VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManagedCharSafeArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Managed array of characters stored within a Integer array by manipulating the SafeArrayDescriptor structure. \r\n\r\nVBA-IDictionary v2.1 (September 02, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
'@Todo Design more appropriate Err.Raise where required

''
'Rubberduck annotations
'@Folder("VBA-IScriptingDictionary.Data Types.SafeArray")
'@ModuleDescription "Managed array of characters stored within a Integer array by manipulating the SafeArrayDescriptor structure. \r\n\r\nVBA-IDictionary v2.1 (September 02, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''

''
'@Version VBA-IScriptingDictionary v2.1 (September 02, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description Creates a managed Integer Array to overlay a string into the SafeArray.pvData
'@Dependencies
'   TypeSafeArray.bas
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified September 02, 2019
'
'@Usage eg Display the Unicode characters of a string
'    Dim managedCharDescriptor As ManagedCharSafeArray
'    Set managedCharDescriptor = New ManagedCharSafeArray
'    Dim managedChars() As Integer
'    managedChars = managedCharDescriptor.ManagedCharArray
'    Dim text As String
'    text = "ABCDabcd"
'    managedCharDescriptor.AllocateCharData text
'    Dim index as long
'    For index = LBound(managedChars) To UBound(managedChars)
'        Debug.Print managedChars(Index)
'    Next
'    managedCharDescriptor.FreeCharArray managedChars

'@Remarks
' Each integer element in the Char Array represents a Unicode character
' for the specified string from AllocateCharData.
' FreeCharArray must be called before the managed integer array obtained
' from the Get Property ManagedCharArray goes out of scope.
' The managed Char Array manipulates a SafeArrayDescriptor without the use
' of the SafeArray API.
' For futher reading regarding the SafeArray descriptor see:
' https://doxygen.reactos.org/db/d60/dll_2win32_2oleaut32_2safearray_8c_source.html
' https://stackoverflow.com/questions/18784470/where-is-safearray-var-type-stored
'------------------------------------------------------------'

'@Todo possibly only require one lock at a time to exist as could be issues with multiple managed arrays to the same SafeArray descriptor??

Option Explicit

'============================================='
'Constants
'============================================='
#If VBA7 Then
    Private Const vbNullPtr As LongPtr = 0
#Else
    Private Const vbNullPtr As Long = 0
#End If

Private Const SIZEOF_VBAINTEGER As Long = 2
Private Const MAX_SAFEARRAY_cLock As Long = 16384  '/* Maximum of 16384 locks at a time */

'============================================='
'API Declarations
'============================================='
''
'@API_Declaration
'API's declarations for copy memory by pointer for Windows and Mac, with VBA6 and VBA7 compatibility.
'See https://stackoverflow.com/questions/45756170/how-to-read-write-memory-on-mac-os-x-with-vba
'------------------------------------------------------------'
#If Mac Then
  #If Win64 Then
    Private Declare PtrSafe Function CopyMemoryByPtr Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As LongPtr, _
             ByVal source As LongPtr, _
             ByVal size As Long) _
             As LongPtr
  #Else
    Private Declare Function CopyMemoryByPtr Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As Long, _
             ByVal src As Long, _
             ByVal size As Long) _
             As Long
  #End If
#ElseIf VBA7 Then
  #If Win64 Then
    Private Declare PtrSafe Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal destination As LongPtr, _
             ByVal source As LongPtr, _
             ByVal size As LongLong)
  #Else
    Private Declare PtrSafe Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal destination As LongPtr, _
             ByVal source As LongPtr, _
             ByVal size As Long)
  #End If
#Else
  Private Declare Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" _
          (ByVal destination As Long, _
           ByVal src As Long, _
           ByVal size As Long)
#End If

''
'@API_Declaration
'API's declarations for CopyAnyToMemory for Windows and Mac, with VBA6 and VBA7 compatibility.
'------------------------------------------------------------'
#If Mac Then
  #If Win64 Then
    Private Declare PtrSafe Function CopyAnyToMemory Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As LongPtr, _
             ByRef source As Any, _
             ByVal size As Long) _
             As LongPtr
  #Else
    Private Declare Function CopyMemoryByPtr Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As Long, _
             ByVal source As Any, _
             ByVal size As Long) _
             As Long
  #End If
#ElseIf VBA7 Then
  #If Win64 Then
    Private Declare PtrSafe Sub CopyAnyToMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
            (ByVal destination As LongPtr, _
             ByRef source As Any, _
             ByVal size As LongLong)
  #Else
    Private Declare PtrSafe Sub CopyAnyToMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
        (ByVal destination As LongPtr, _
         ByRef source As Any, _
         ByVal size As Long)
  #End If
#Else
    Private Declare Sub CopyAnyToMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
        (ByVal destination As Long, _
         ByRef source As Any, _
         ByVal size As Long)
#End If

''
'@API_Declaration for VarPtrArray
'------------------------------------------------------------'
#If VBA7 Then
    Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias _
            "VarPtr" (ByRef var() As Any) As LongPtr
#Else
    Private Declare PtrSafe Function VarPtrArray Lib "VBE6" Alias _
            "VarPtr" (ByRef var() As Any) As Long
#End If

'============================================='
'Types
'============================================='
Private Type TSafeArrayDescriptor
    vt(15) As Byte                  'Array variable type of 16 bytes which precedes the SafeArray pointer
    integerSafeArray As SafeArray1D 'SafeArray Structure for one dimension
End Type

'============================================='
'Private Variables
'============================================='
Private this As TSafeArrayDescriptor

'============================================='
'Constructors and destructors
'============================================='
Private Sub Class_Initialize()
    CreateSafeArrayDescriptor
End Sub

Private Sub Class_Terminate()
    DisposeData
    If this.integerSafeArray.cLocks > 0 Then
        Debug.Print "Err.Raise E_UNEXPECTED 'i.e. There are still arrays attached to the SafeArrayDescriptor which must be freed, which will cause memory leaks/crash"
        'Err.Raise E_UNEXPECTED 'i.e. There are still arrays attached to the SafeArrayDescriptor which must be freed.
    End If
End Sub

'============================================='
'Properties
'============================================='

''
'@Description("Returns an initialized, locked and empty, managed Integer Array which must be freed calling FreeManagedCharArray before going out of scope.")
'@Remarks The returned Integer code must be returned to an uninitialize Integer Array or
'         is required reinstate its original SafeArray pointer.
'         The calling code cann't resize the integer array returned, however the array contents may be altered.
'@return (Integer Array)
'------------------------------------------------------------'
Public Property Get ManagedCharArray() As Integer()
Attribute ManagedCharArray.VB_Description = "Returns an initialized, locked and empty, managed Integer Array which must be freed calling FreeManagedCharArray before going out of scope."
    #If VBA7 Then
        Dim pManagedIntegerArray As LongPtr
        Dim pSafeArray As LongPtr
    #Else
        Dim pManagedIntegerArray As Long
        Dim pSafeArray As Long
    #End If

    pManagedIntegerArray = VarPtrArray(ManagedCharArray)
    pSafeArray = VarPtr(this.integerSafeArray)
    SafeArrayLock 'increment SAFEARRAY.cLock
    CopyMemoryByPtr pManagedIntegerArray, VarPtr(pSafeArray), LenB(pManagedIntegerArray)
End Property

'============================================='
'Public Methods
'============================================='

''
'@Description("Allocatates a string to an Integer SafeArray.pvData") A
'------------------------------------------------------------'
Public Sub AllocateCharData(ByRef inText As String)
    this.integerSafeArray.pvData = StrPtr(inText)
    this.integerSafeArray.cElements = Len(inText)
End Sub

''
'@Description("Frees the managed characer array SafeArray descriptor.")
'Required to free any array obtained from the managed integer array by removing the SafeArray reference
'Only free integer arrays attached to SafeArrayDescriptor structure
'@Error 5  Invalid procedure call or argument.
'          Raise for an invalid Integer Array which doesn't match the managed SafeArray pointer.
'------------------------------------------------------------'
Public Sub FreeManagedCharArray(ByRef outIntegerArray() As Integer)
Attribute FreeManagedCharArray.VB_Description = "Frees the managed characer array SafeArray descriptor."
    #If VBA7 Then
        Dim pIntegerArray As LongPtr
        '@Ignore VariableNotAssigned
        Dim pSafeArray As LongPtr
    #Else
        Dim pIntegerArray As Long
        Dim pSafeArray As Long
    #End If
    
    pIntegerArray = VarPtrArray(outIntegerArray)
    CopyMemoryByPtr VarPtr(pSafeArray), pIntegerArray, LenB(pSafeArray)
    If pSafeArray = VarPtr(this.integerSafeArray) Then
        CopyAnyToMemory pIntegerArray, vbNullPtr, LenB(pIntegerArray)
        SafeArrayUnLock
    Else
        'Can only free arrays attached to the managed SafeArrayDescriptor structure."
        VBA.Err.Raise 5, "ManagedCharSafeArray FreeMangedCharArray", "Can only free an array which is managed for its own descriptor"
    End If
End Sub

''
'@Description("Returns an initialized and empty managed Integer SafeArray descriptor which must be freed calling FreeManagedCharArray before going out of scope.")
'@Remarks The returned Integer code must be returned to an uninitialize Integer Array or
'         is required reinstate its original SafeArray pointer.
'         The calling code cann't resize the integer array returned, however the array contents may be altered.
'------------------------------------------------------------'
Private Sub CreateSafeArrayDescriptor()
Attribute CreateSafeArrayDescriptor.VB_Description = "Returns an initialized and empty managed Integer SafeArray descriptor which must be freed calling FreeManagedCharArray before going out of scope."
    this.vt(12) = CByte(SAFEARRAY_VT.VT_INTEGER) 'The four byte DWord containts the Variable Type when  fFeatures = FADF_HAVEVARTYPE is set
    this.integerSafeArray.cbElements = SIZEOF_VBAINTEGER
    this.integerSafeArray.cDims = 1
    this.integerSafeArray.fFeatures = SafeArrayFeatures.FADF_HAVEVARTYPE
End Sub

''
'@Description("Disposes of the SafeArray.pvData by setting a null pointer and resets the number of elements to zero")
'------------------------------------------------------------'
Private Sub DisposeData()
Attribute DisposeData.VB_Description = "Disposes of the SafeArray.pvData by setting a null pointer and resets the number of elements to zero"
    this.integerSafeArray.pvData = vbNullPtr
    this.integerSafeArray.cElements = 0
End Sub

''
'@Description("Increments the locks the Safe Array to disallow resizing etc.")
'@Remarks
'@Error VBA.vbObjectError Raised when exceed the maximum allowed number of locks on an SafeArray.
'------------------------------------------------------------'
Private Sub SafeArrayLock()
Attribute SafeArrayLock.VB_Description = "Increments the locks the Safe Array to disallow resizing etc."
    If this.integerSafeArray.cLocks < MAX_SAFEARRAY_cLock Then
        this.integerSafeArray.cLocks = this.integerSafeArray.cLocks + 1
    Else
        '@Todo Err.Raise E_UNEXPECTED, "Out of locks"
        Err.Raise VBA.vbObjectError, "MamagedCharSafeArray, SafeArrayLock", "Out of locks"
    End If
End Sub

''
'@Description("Returns an initialized, locked and empty, managed Integer Array which must be freed calling FreeManagedCharArray before going out of scope.")
'@Remarks
'------------------------------------------------------------'
Private Sub SafeArrayUnLock()
Attribute SafeArrayUnLock.VB_Description = "Returns an initialized, locked and empty, managed Integer Array which must be freed calling FreeManagedCharArray before going out of scope."
    If this.integerSafeArray.cLocks > 0 Then
        this.integerSafeArray.cLocks = this.integerSafeArray.cLocks - 1
    End If
End Sub




