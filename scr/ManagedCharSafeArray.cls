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
'@Version VBA-IScriptingDictionary v2.1 (February 05, 2023)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description Creates a managed Char Array where a specified string is overlayed into the SafeArray.pvData
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified February 05, 2023
'@Dependencies
'   TypeSafeArray.bas
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
'Errors Raised:
'@Error 5   Invalid procedure call or argument
'           Raised when the SafeArray descriptor has already been allocated to a managed Char array.
'           Raised when the managed Char array specified isn't an unitialised Integer Array.
'@Error 10  This array is fixed or temporary locked.
'           Raised when the managedChars array is attempted to resized as it is locked.
'
'@Remarks
' The ManagedCharSafeArray manipulates a SafeArray descriptor without the use
' of the SafeArray API.
' Each integer element in the managed Char Array represents a Unicode character
' for the specified string using the method AllocateCharData.
' The Integer/managed Char array specified must be uninitialize,
' to allow a SafeArray descriptor to be allocated.
' The calling code cannot resize the integer array returned as it is locked,
' however the array contents may be altered.
' Attempting to resize the managed Chars array will result in the
' Runtime Error 10 This array is fixed or temporary locked.
' The SafeArray descriptor should be freed from the array by calling the
' Dispose method, this is to ensure that it is freed before
' the managed chars array goes out of scope.
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

Private Const SIZEOF_VBA_INTEGER As Long = 2

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
    Private Declare Function CopyAnyToMemory Lib "libc.dylib" Alias "memmove" _
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
    vt(15) As Byte                              'Array variable type of 16 bytes which precedes the SafeArray pointer.
    charSafeArray As SafeArray1D                'SafeArray Structure for one dimensional array.
End Type

#If VBA7 Then
    Private Type TManagedCharArray
        saDescriptor As TSafeArrayDescriptor    'SafeArray descriptor structure for a one dimensional Integer array.
        pManagedChars As LongPtr                'Pointer to managed Char array associated with the SafeArray descriptor.
    End Type
#Else
    Private Type TManagedCharArray
        saDescriptor As TSafeArrayDescriptor    'SafeArray descriptor structure for a one dimensional Integer array.
        pManagedChars As Long                   'Pointer to managed Char array associated with the SafeArray descriptor.
    End Type
#End If

'============================================='
'Private Variables
'============================================='
Private this As TManagedCharArray

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
'@param outManagedChars()  The uninitialised managed Char array to be allocated a SafeArray descriptor.
'@Remarks
' The Integer array specified must be uninitialize to allow a SafeArray descriptor to be allocated.
'------------------------------------------------------------'
Public Function Create(ByRef outManagedChars() As Integer) As ManagedCharSafeArray
Attribute Create.VB_Description = "Creates a new instance of a ManagedCharSafeArray for the managed Char array specified."
    Dim newManagedCharSafeArray  As ManagedCharSafeArray
    Set newManagedCharSafeArray = New ManagedCharSafeArray
    newManagedCharSafeArray.AllocateSafeArrayToCharArray outManagedChars
    Set Create = newManagedCharSafeArray
End Function

''
'@Description("Allocatates the specified string data to the managed Char array.")
'@param inText The specified string to be allocated to the managed Char array SafeArray.pvData
'@Remarks
' For the specified String updates the SafeArray.pvData and the SafeArray.cElements
' according to its character length.
'------------------------------------------------------------'
Public Sub AllocateCharData(ByRef inText As String)
Attribute AllocateCharData.VB_Description = "Allocatate the specified string data to the SafeArray.pvData"
    this.saDescriptor.charSafeArray.pvData = StrPtr(inText)
    this.saDescriptor.charSafeArray.cElements = Len(inText)
End Sub

''
'@Description("Allocates the SafeArray descriptor to the specified managed Char array.")
'@param outManagedChars The uninitialised managed Char array to be allocated a SafeArray descriptor.
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
Public Sub AllocateSafeArrayToCharArray(ByRef outManagedChars() As Integer)
Attribute AllocateSafeArrayToCharArray.VB_Description = "Allocates the SafeArray descriptor to the specified managed Char array."
    #If VBA7 Then
        Dim pOutManagedChars As LongPtr
        '@Ignore VariableNotAssigned
        Dim pSafeArray As LongPtr
    #Else
        Dim pOutManagedChars As Long
        Dim pSafeArray As Long
    #End If
    'Obtain the array pointer of outManagedChars
    pOutManagedChars = VarPtrArray(outManagedChars())
    'Obtain the SafeArray pointer of outManagedChars
    CopyMemoryByPtr VarPtr(pSafeArray), pOutManagedChars, LenB(pSafeArray)
    'Allocate the SafeArray descriptor to the specified unitialize integer array
    If pSafeArray = NULL_PTR Then
        'The SafeArray descriptor can only be allocted to one managed Char array at a time
        If this.pManagedChars = NULL_PTR Then
            #If VBA7 Then
                Dim pNewSafeArray As LongPtr
            #Else
                Dim pNewSafeArray As Long
            #End If
            pNewSafeArray = VarPtr(this.saDescriptor.charSafeArray)
            'Copy the new SafeArray pointer to the managed chars array
            CopyMemoryByPtr pOutManagedChars, VarPtr(pNewSafeArray), LenB(this.pManagedChars)
            'set the managed Char array pointer that is associated with the SafeArray descriptor
            this.pManagedChars = pOutManagedChars
        Else
            VBA.Err.Raise 5, "ManagedCharSafeArray.AllocateSafeArrayToCharArray", "The SafeArray descriptor has already been allocated to a managed Char array."
        End If
    Else
        VBA.Err.Raise 5, "ManagedCharSafeArray.AllocateSafeArrayToCharArray", "Cannot assign an initialise Integer array to the managed Char array."
    End If
End Sub

''
'@Description("Disposes of the SafeArray descriptor from the managed Char array.")
'------------------------------------------------------------'
Public Sub Dispose()
Attribute Dispose.VB_Description = "Disposes of the SafeArray descriptor from the managed Char array."
    DisposeData
    DisposeManagedCharArray
End Sub

'============================================='
'Private Methods
'============================================='

''
'@Description("Creates initialized and empty SafeArray descriptor for a one dimensional Integer array, which is locked.")
'@Remarks
' Creates the managed Chars SafeArray descriptor for a read-only, one dimensional Integer array, which
' is only resized and char data added calling the AllocateCharData method.
' The four byte DWord preceding the SafeArray pointer containts the Variable Type
' of an Array when fFeatures = FADF_HAVEVARTYPE is set.
'------------------------------------------------------------'
Private Sub CreateSafeArrayDescriptor()
Attribute CreateSafeArrayDescriptor.VB_Description = "Creates initialized and empty SafeArray descriptor for a one dimensional Integer array, which is locked."
    'The last four byte DWord contains the array variable type for fFeatures = FADF_HAVEVARTYPE
    this.saDescriptor.vt(12) = CByte(SAFEARRAY_VT.VT_INTEGER)
    this.saDescriptor.charSafeArray.cbElements = SIZEOF_VBA_INTEGER
    this.saDescriptor.charSafeArray.cDims = 1  'One dimensional
    this.saDescriptor.charSafeArray.fFeatures = SafeArrayFeatures.FADF_HAVEVARTYPE
    this.saDescriptor.charSafeArray.cLocks = 1 'lock the array from being resized
End Sub

''
'@Description("Disposes of the SafeArray.pvData by setting a null pointer and resets the number of elements to zero")
'------------------------------------------------------------'
Private Sub DisposeData()
Attribute DisposeData.VB_Description = "Disposes of the SafeArray.pvData by setting a null pointer and resets the number of elements to zero"
    this.saDescriptor.charSafeArray.pvData = NULL_PTR
    this.saDescriptor.charSafeArray.cElements = 0
End Sub

''
'@Description("Disposes of the SafeArray descriptor from its associated managed Char array.")
'------------------------------------------------------------'
Private Sub DisposeManagedCharArray()
Attribute DisposeManagedCharArray.VB_Description = "Disposes of the SafeArray descriptor from its associated managed Char array."
    If Not this.pManagedChars = NULL_PTR Then
        CopyAnyToMemory this.pManagedChars, NULL_PTR, LenB(this.pManagedChars)
        this.pManagedChars = NULL_PTR
    End If
End Sub
