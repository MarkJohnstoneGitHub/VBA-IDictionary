VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManagedCharSafeArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Creates a SafeArray descriptor to be associated with a managed array of characters represented by an Integer array of Unicode values. \r\n\r\nVBA-IDictionary v2.1 (September 02, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''
'Rubberduck annotations
'@Folder("VBA-IScriptingDictionary.Data Types.SafeArray")
'@PredeclaredId
'@ModuleDescription "Creates a SafeArray descriptor to be associated with a managed array of characters represented by an Integer array of Unicode values. \r\n\r\nVBA-IDictionary v2.1 (September 02, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''

''
'@Version VBA-IScriptingDictionary v2.1 (September 02, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description Creates a managed Integer Array to overlay a string into the SafeArray.pvData
'@Dependencies
'   TypeSafeArray.bas
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified September 05, 2019
'
'@Usage eg Display the Unicode characters of a string
'    Dim managedChars() As Integer
'    Dim managedCharsDescriptor As ManagedCharSafeArray
'    Set managedCharsDescriptor = ManagedCharSafeArray.Create(managedChars)
'    Dim text As String
'    text = "ABCDabcd"
'    managedCharsDescriptor.AllocateCharData text
'    Dim index As Long
'    For index = LBound(managedChars) To UBound(managedChars)
'        Debug.Print managedChars(index)
'    Next
'    managedCharsDescriptor.Dispose
'
'@Remarks
' The Integer/managed Char array specified must be uninitialize to allow a SafeArray descriptor to be allocated.
' The calling code cann't resize the integer array returned as it is locked, however the array contents may be altered.
' The SafeArray descriptor should be freed from the array by calling the Dispose method, this is
' to ensure that it is freed before the array goes out of scope.
' Each integer element in the Char Array represents a Unicode character
' for the specified string using the method AllocateCharData.
' The managed Char Array manipulates a SafeArrayDescriptor without the use
' of the SafeArray API.
'
' For futher reading regarding the SafeArray descriptor see:
' https://doxygen.reactos.org/db/d60/dll_2win32_2oleaut32_2safearray_8c_source.html
' https://stackoverflow.com/questions/18784470/where-is-safearray-var-type-stored
'------------------------------------------------------------'

Option Explicit

'============================================='
'Constants
'============================================='
#If VBA7 Then
    Private Const NULL_PTR As LongPtr = 0
#Else
    Private Const NULL_PTR As Long = 0
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
    vt(15) As Byte                                   'Array variable type of 16 bytes which precedes the SafeArray pointer.
    charSafeArray As SafeArray1D                     'SafeArray Structure for one dimensional array.
End Type

#If VBA7 Then
    Private Type TManagedCharSafeArray
        SafeArrayDescriptor As TSafeArrayDescriptor  'SafeArray descriptor structure for a one dimensional Integer array.
        pManagedChars As LongPtr                     'Pointer to managed Char array associated with the SafeArray descriptor.
    End Type
#Else
    Private Type TManagedCharSafeArray
        SafeArrayDescriptor As TSafeArrayDescriptor  'SafeArray descriptor structure for a one dimensional Integer array.
        pManagedChars As Long                        'Pointer to managed Char array associated with the SafeArray descriptor.
    End Type
#End If

'============================================='
'Private Variables
'============================================='
Private this As TManagedCharSafeArray

'============================================='
'Constructors and destructors
'============================================='
Private Sub Class_Initialize()
    CreateSafeArrayDescriptor
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

'============================================='
'Public Methods
'============================================='

''
'@Description("Creates a new instance of a ManagedCharSafeArray for the managed Char array specified.")
'@Remarks
' The Integer array specified must be uninitialize to allow a SafeArray descriptor to be allocated.
'------------------------------------------------------------'
Public Function Create(ByRef outCharsArray() As Integer) As ManagedCharSafeArray
Attribute Create.VB_Description = "Creates a new instance of a ManagedCharSafeArray for the managed Char array specified."
    Dim managedChars As ManagedCharSafeArray
    Set managedChars = New ManagedCharSafeArray
    managedChars.SafeArrayAllocateToCharArray outCharsArray
    Set Create = managedChars
End Function

''
'@Description("Allocatate the specified string data to the SafeArray.pvData")
'------------------------------------------------------------'
Public Sub AllocateCharData(ByRef inText As String)
Attribute AllocateCharData.VB_Description = "Allocatate the specified string data to the SafeArray.pvData"
    this.SafeArrayDescriptor.charSafeArray.pvData = StrPtr(inText)
    this.SafeArrayDescriptor.charSafeArray.cElements = Len(inText)
End Sub

''
'@Description("Allocates the SafeArray descriptor to the specified managed Char array.")
'@param outManagedChars The uninitialised Integer array which is allocated a SafeArray descriptor.
'@Remarks
' The Integer array specified must be uninitialize to allow a SafeArray descriptor to be allocated.
' The calling code cann't resize the integer array returned as it is locked,
' however the array contents may be altered.
' The SafeArray descriptor should be freed from the manged array before it goes out of scope,
' by calling the Dispose method.
'@Error 5   Invalid procedure call or argument
'           Raised when the SafeArray descriptor has already been allocated to a managed Char array.
'           Raised when the managed Char array specified isn't an unitialised Integer Array.
'------------------------------------------------------------'
Public Sub SafeArrayAllocateToCharArray(ByRef outManagedChars() As Integer)
Attribute SafeArrayAllocateToCharArray.VB_Description = "Allocates the SafeArray descriptor to the specified managed Char array."
    #If VBA7 Then
        Dim pOutManagedChars As LongPtr
        '@Ignore VariableNotAssigned
        Dim pSafeArray As LongPtr
    #Else
        Dim pOutManagedChars As Long
        Dim pSafeArray As Long
    #End If

    'Obtain the array pointer of managedChars
    pOutManagedChars = VarPtrArray(outManagedChars())
    'Obtain the SafeArray pointer of managedChars
    CopyMemoryByPtr VarPtr(pSafeArray), pOutManagedChars, LenB(pSafeArray)
    
    'Assign the SafeArray descriptor to an unitialize integer array
    If pSafeArray = NULL_PTR Then
        'A SafeArray descriptor can only be associated with only one managed char array at a time
        If this.pManagedChars = NULL_PTR Then
            #If VBA7 Then
                Dim pNewSafeArray As LongPtr
            #Else
                Dim pNewSafeArray As Long
            #End If
            pNewSafeArray = VarPtr(this.SafeArrayDescriptor.charSafeArray)
            SafeArrayLock 'lock the array from being resized
            'Copy by reference the SafeArray pointer from the SA descriptor to the managed chars array pointer
            CopyMemoryByPtr pOutManagedChars, VarPtr(pNewSafeArray), LenB(this.pManagedChars)
            'set the managed Char array pointer that is associated with the SafeArray descriptor
            this.pManagedChars = pOutManagedChars
        Else
            VBA.Err.Raise 5, "ManagedCharSafeArray ManagedCharArray", "The SafeArray descriptor has already been assigned to a managed Char array."
        End If
    Else
        VBA.Err.Raise 5, "ManagedCharSafeArray ManagedCharArray", "Cannot assign an initialise array to the managed array."
    End If
End Sub

''
'@Description("Creates an initialized and empty SafeArray descriptor for a one dimensional Integer array.")
'@Remarks
' The four byte DWord preceding the SafeArray pointer containts the Variable Type
' of an Array when fFeatures = FADF_HAVEVARTYPE is set.
'------------------------------------------------------------'
Private Sub CreateSafeArrayDescriptor()
Attribute CreateSafeArrayDescriptor.VB_Description = "Creates an initialized and empty SafeArray descriptor for a one dimensional Integer array."
    this.SafeArrayDescriptor.vt(12) = CByte(SAFEARRAY_VT.VT_INTEGER)        'The four byte DWord contains the array variable type/
    this.SafeArrayDescriptor.charSafeArray.cbElements = SIZEOF_VBAINTEGER
    this.SafeArrayDescriptor.charSafeArray.cDims = 1
    this.SafeArrayDescriptor.charSafeArray.fFeatures = SafeArrayFeatures.FADF_HAVEVARTYPE
End Sub

''
'@Description("Disposes of the SafeArray descriptor from the managed Char array and clears the SafeArray.pvData and resizes to 0 elements.")
'------------------------------------------------------------'
Public Sub Dispose()
Attribute Dispose.VB_Description = "Disposes of the SafeArray descriptor from the managed Char array and clears the SafeArray.pvData and resizes to 0 elements."
    DisposeData
    DisposeManagedSafeArray
    SafeArrayUnLock
End Sub

''
'@Description("Disposes of the SafeArray.pvData by setting a null pointer and resets the number of elements to zero")
'------------------------------------------------------------'
Private Sub DisposeData()
Attribute DisposeData.VB_Description = "Disposes of the SafeArray.pvData by setting a null pointer and resets the number of elements to zero"
    this.SafeArrayDescriptor.charSafeArray.pvData = NULL_PTR
    this.SafeArrayDescriptor.charSafeArray.cElements = 0
End Sub

''
'@Description("Disposes of the SafeArray descriptor from its associated managed Char array.")
'------------------------------------------------------------'
Private Sub DisposeManagedSafeArray()
Attribute DisposeManagedSafeArray.VB_Description = "Disposes of the SafeArray descriptor from its associated managed Char array."
    If Not this.pManagedChars = NULL_PTR Then
        CopyAnyToMemory this.pManagedChars, NULL_PTR, LenB(this.pManagedChars)
        this.pManagedChars = NULL_PTR
    End If
End Sub

''
'@Description("Sets the SafeArray.clocks to disallow array resizing.")
'@Remarks
'@Error VBA.vbObjectError Raised when a lock already in use.
'------------------------------------------------------------'
Private Sub SafeArrayLock()
Attribute SafeArrayLock.VB_Description = "Sets the SafeArray.clocks to disallow array resizing."
    If this.SafeArrayDescriptor.charSafeArray.cLocks = 0 Then
        this.SafeArrayDescriptor.charSafeArray.cLocks = 1
    Else
        Err.Raise VBA.vbObjectError, "MamagedCharSafeArray, SafeArrayLock", "Managed Char array descriptor is already allocated"
    End If
End Sub

''
'@Description("Releases any locks on the Safe Array descriptor.")
'------------------------------------------------------------'
Private Sub SafeArrayUnLock()
Attribute SafeArrayUnLock.VB_Description = "Releases any locks on the Safe Array descriptor."
    If Not this.SafeArrayDescriptor.charSafeArray.cLocks = 0 Then
        this.SafeArrayDescriptor.charSafeArray.cLocks = 0
    End If
End Sub









