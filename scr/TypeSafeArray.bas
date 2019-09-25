Attribute VB_Name = "TypeSafeArray"
Attribute VB_Description = "SafeArray and SafeArrayBound types and required enums for SafeArray manipulation.\r\n\r\nVBA-IDictionary  v2.1 (September 02, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''
'Rubberduck annotations
'@Folder("VBA-IScriptingDictionary.Data Types.SafeArray")
'@ModuleDescription "SafeArray and SafeArrayBound types and required enums for SafeArray manipulation.\r\n\r\nVBA-IDictionary  v2.1 (September 02, 2019)\r\n(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary\r\nAuthor: markjohnstone@hotmail.com\r\n"
''

''
'@Version VBA-IDictionary v2.1 (September 02, 2019)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/VBA-IDictionary
'@Description SafeArray and SafeArrayBound types and required enums for SafeArray manipulation.
'@Author Mark Johnstone markjohnstone@hotmail.com
'@LastModified September 25, 2019
'   Updated incorrect enum values in SAFEARRAY_VT for array types which are not currently used.
'@Remarks
' For futher reading regarding the SafeArray descriptor see:
' https://doxygen.reactos.org/db/d60/dll_2win32_2oleaut32_2safearray_8c_source.html
' https://stackoverflow.com/questions/18784470/where-is-safearray-var-type-stored
''

Option Explicit

#If VBA7 Then
    Type SafeArray1D
        cDims As Integer            ' nr of dimensions for the array
        fFeatures As Integer        ' extra information about the array contents
        cbElements As Long          ' nr of bytes per array element. Possible Examples: 1=byte,2=integer,4=long,8=currency
        cLocks As Long              ' nr of times array was locked w/o being unlocked
        pvData As LongPtr           ' Address to 1st array item, can be a pointer to another structure/address
        cElements    As Long        ' The number of elements in the dimension.
        lLbound      As Long        ' The lower bound of the array.
    End Type
#Else
    Type SafeArray1D
        cDims As Integer            ' nr of dimensions for the array
        fFeatures As Integer        ' extra information about the array contents
        cbElements As Long          ' nr of bytes per array element. Possible Examples: 1=byte,2=integer,4=long,8=currency
        cLocks As Long              ' nr of times array was locked w/o being unlocked
        pvData As Long              ' Address to 1st array item, can be a pointer to another structure/address
        cElements    As Long        ' The number of elements in the dimension.
        lLbound      As Long        ' The lower bound of the array.
    End Type
#End If

Public Type SAFEARRAYBOUND
    cElements    As Long            'The number of elements in the dimension.
    lLbound      As Long            'The lower bound of the array.
End Type

Public Enum SafeArrayFeatures
    FADF_AUTO = &H1             'An array that is allocated on the stack.
    FADF_STATIC = &H2           'An array that is statically allocated.
    FADF_EMBEDDED = &H4         'An array that is embedded in a structure.
    FADF_FIXEDSIZE = &H10       'An array that may not be resized or reallocated.
    FADF_RECORD = &H20          'An array that contains records. When set, there will be a pointer to the IRecordInfo interface at negative offset 4 in the array descriptor.
    FADF_HAVEIID = &H40         'An array that has an IID identifying interface. When set, there will be a GUID at negative offset 16 in the safe array descriptor. Flag is set only when FADF_DISPATCH or FADF_UNKNOWN is also set.
    FADF_HAVEVARTYPE = &H80     'An array that has a variant type. The variant type can be retrieved with SafeArrayGetVartype.
    FADF_BSTR = &H100           'An array of BSTRs.
    FADF_UNKNOWN = &H200        'An array of IUnknown*.
    FADF_DISPATCH = &H400       'An array of IDispatch*.
    FADF_VARIANT = &H800        'Array is of type Variant
    FADF_RESERVED = &HF008      'Bits reserved for future use.
    FADF_CREATEVECTOR = &H2000  '
End Enum

#If VBA7 Then
    Public Enum SAFEARRAY_VT
        VT_INTEGER = VBA.vbInteger                      'VT_I2 = 2
        VT_LONG = VBA.vbLong                            'VT_I4 = 3
        VT_SINGLE = VBA.vbSingle                        'VT_R4 = 4
        VT_DOUBLE = VBA.vbDouble                        'VT_R8 = 5
        VT_CURRENCY = VBA.vbCurrency                    'VT_CY = 6
        VT_DATE = VBA.vbDate                            'VT_DATE = 7
        VT_STRING = VBA.vbString                        'VT_BSTR = 8
        VT_OBJECT = VBA.vbObject                        'VT_DISPATCH = 9
        VT_BOOLEAN = VBA.vbBoolean                      'VT_BOOL = 11
        VT_VARIANT = VBA.vbVariant                      'VT_VARIANT = 12
        VT_DECIMAL = VBA.vbDecimal                      'VT_DECIMAL = 14
        VT_BYTE = VBA.vbByte                            'VT_UI1 = 17
        VT_LONGLONG = VBA.vbLongLong                    'VT_I8 = 20
        VT_USERDEFINEDTYPE = VBA.vbUserDefinedType      'VT_RECORD = 34
        [_First] = VT_INTEGER
        [_Last] = VT_USERDEFINEDTYPE
    End Enum
#Else
    Public Enum SAFEARRAY_VT
        VT_INTEGER = VBA.vbInteger                      'VT_I2 = 2
        VT_LONG = VBA.vbLong                            'VT_I4 = 3
        VT_SINGLE = VBA.vbSingle                        'VT_R4 = 4
        VT_DOUBLE = VBA.vbDouble                        'VT_R8 = 5
        VT_CURRENCY = VBA.vbCurrency                    'VT_CY = 6
        VT_DATE = VBA.vbDate                            'VT_DATE = 7
        VT_STRING = VBA.vbString                        'VT_BSTR = 8
        VT_OBJECT = VBA.vbObject                        'VT_DISPATCH = 9
        VT_BOOLEAN = VBA.vbBoolean                      'VT_BOOL = 11
        VT_VARIANT = VBA.vbVariant                      'VT_VARIANT = 12
        VT_DECIMAL = VBA.vbDecimal                      'VT_DECIMAL = 14
        VT_BYTE = VBA.vbByte                            'VT_UI1 = 17
        VT_USERDEFINEDTYPE = VBA.vbUserDefinedType      'VT_RECORD = 34
        [_First] = VT_INTEGER
        [_Last] = VT_USERDEFINEDTYPE
    End Enum
#End If

'Notes:

'https://docs.microsoft.com/en-us/cpp/atl/reference/ccomsafearray-class?view=vs-2019
'A CComSafeArray can contain the following subset of VARIANT data types:
'varType    Description
'VT_I1      char
'VT_I2      short
'VT_I4      int
'VT_I4      long
'VT_I8      longlong
'VT_UI1     byte
'VT_UI2     ushort
'VT_UI4     uint
'VT_UI4     ulong
'VT_UI8     ulonglong
'VT_R4      float
'VT_R8      double
'VT_DECIMAL decimal pointer
'VT_VARIANT variant pointer
'VT_CY      Currency data type


'https://doxygen.reactos.org/d5/db1/dll_2win32_2dbghelp_2compat_8h.html#af5ea9823ca4227b3883c737c5217ff28af39ede0dde0514d01c526e75abb12271
'{
'     VT_EMPTY = 0,
'     VT_NULL = 1,
'     VT_I2 = 2,
'     VT_I4 = 3,
'     VT_R4 = 4,
'     VT_R8 = 5,
'     VT_CY = 6,
'     VT_DATE = 7,
'     VT_BSTR = 8,
'     VT_DISPATCH = 9,
'     VT_ERROR = 10,
'     VT_BOOL = 11,
'     VT_VARIANT = 12,
'     VT_UNKNOWN = 13,
'     VT_DECIMAL = 14,
'     VT_I1 = 16,
'     VT_UI1 = 17,
'     VT_UI2 = 18,
'     VT_UI4 = 19,
'     VT_I8 = 20,
'     VT_UI8 = 21,
'     VT_INT = 22,
'     VT_UINT = 23,
'     VT_VOID = 24,
'     VT_HRESULT = 25,
'     VT_PTR = 26,
'     VT_SAFEARRAY = 27,
'     VT_CARRAY = 28,
'     VT_USERDEFINED = 29,
'     VT_LPSTR = 30,
'     VT_LPWSTR = 31,
'     VT_RECORD = 36,
'     VT_INT_PTR = 37,
'     VT_UINT_PTR = 38,
'     VT_FILETIME = 64,
'     VT_BLOB = 65,
'     VT_STREAM = 66,
'     VT_STORAGE = 67,
'     VT_STREAMED_OBJECT = 68,
'     VT_STORED_OBJECT = 69,
'     VT_BLOB_OBJECT = 70,
'     VT_CF = 71,
'     VT_CLSID = 72,
'     VT_VERSIONED_STREAM = 73,
'     VT_BSTR_BLOB = 0xfff,
'     VT_VECTOR = 0x1000,
'     VT_ARRAY = 0x2000,
'     VT_BYREF = 0x4000,
'     VT_RESERVED = 0x8000,
'     VT_ILLEGAL = 0xffff,
'     VT_ILLEGALMASKED = 0xfff,
'     VT_TYPEMASK = 0xfff
' };



