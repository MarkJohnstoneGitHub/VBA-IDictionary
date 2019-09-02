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
'@LastModified September 02, 2019
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

Public Enum SAFEARRAY_VT
    VT_BYTE = 1
    VT_BOOL = 2
    VT_INTEGER = 2
    VT_LONG = 4
    VT_SINGLE = 4
    VT_DOUBLE = 8
    VT_CURRENCY = 8
    VT_DECIMAL = 14
    VT_DATE = 8
    VT_OBJECT = 4
    VT_STRING = 10
    VT_VARIANT = 16
End Enum


