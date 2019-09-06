# VBA-IDictionary v2.1 (September 03, 2019)

VBA dictionary alternative to the [Scripting.Dictionary](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object) which is Mac compatible which implements a IScriptingDictionary interface to allow for various dictionary implementations. 

**Classes Required:**
 - [IScriptingDictionary.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/IScriptingDictionary.cls)
 
 - [Dictionary.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/Dictionary.cls)
 
 - [DictionaryKeyValuePair.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/DictionaryKeyValuePair.cls)
 
 - [ScriptingDictionary.cls]( https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/ScriptingDictionary.cls)
 
 - [ITextEncoding.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/ITextEncoding.cls)
 
 - [TextEncoderASCII.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/TextEncoderASCII.cls)
 
  - [TextEncoderUnicode.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/TextEncoderUnicode.cls)
  
  - [ManagedCharSafeArray.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/ManagedCharSafeArray.cls)

  - [TypeSafeArray.bas](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/TypeSafeArray.bas)
   
Optional Reference Addin:  Microsoft Scripting Runtime Scripting scrrun.dll
 
## Usage

### Creating an IScriptingDictionary

**Creating a IScriptingDictionary with Dictionary.Create()**

The Dictionary.cls is an IScriptingDictionary factory class.  It returns an available IScriptingDictionary implementation, according to the compiler constants and/or the implementation specified.  I.e. If the Scripting.Dictionary is specified and not available the DictionaryKeyValuePair will be returned.  Cannot use New and must use the Dictionary.Create method to create a IScriptingDictionary object. 

The Dictionary.Create has three optional parameters:

 - dictionaryType : Specifies which IScriptingDictionary implementation to create. i.e. ScriptingDictionary or DictionaryKeyValuePairs. Default is the IScriptingDictionaryType.isdtScriptingDictionary.

 - compareMethod  : Specifies how string keys are handled either case sensitive or ignore case. Default value is vbBinaryCompare.

 - encodingMethod : Specifies which encoding method to use on case sensitive string keys i.e. Unicode or ASCII.  It is only applicable to the DictionaryKeyValuePair class for improved performance, and  the option TextEncodingMethod.temAscii should only be used where string keys are ASCII compatible.  Default value is TextEncodingMethod.temUnicode.

Examples:

  Dim myDictionary As IScriptingDictionary
  
  Set myDictionary = Dictionary.Create
  
  Set myDictionary = Dictionary.Create(IScriptingDictionaryType.isdtScriptingDictionary, VBA.vbTextCompare)
  
  Set myDictionary = Dictionary.Create(IScriptingDictionaryType.isdtDictionaryKeyValuePair, VBA.vbBinaryCompare, TextEncodingMethod.temAscii)
  
**Creating using directly a IScriptingDictionary implementation provided**

Two IScriptingDictionary implementations have been provided, DictionaryKeyValuePair and ScriptingDictionary, they can be created using New or the Create method. 
 
The ScriptingDictionary.Create() has one optional parameter:

 - compareMethod : Specifies how string keys are handled either case sensitive or ignore case. Default value is vbBinaryCompare.
 
 Example
 
  Dim myDictionary As IScriptingDictionary 'or could use As ScriptingDictionary
  
  Set myDictionary = ScriptingDictionary.Create(VBA.vbTextCompare)
  
The DictionaryKeyValuePair.Create() has the two optional parameters:

 - compareMethod : Specifies how string keys are handled either case sensitive or ignore case.  Default value is vbBinaryCompare.
 
 - encodingMethod : Specifies which encoding text method to use on case sensitive string keys i.e. Unicode or ASCII. This parameter provides performance improvement for ASCII compatibile string keys. The default value is TextEncodingMethod.temUnicode
 
   Dim myDictionary As IScriptingDictionary 'or could use As DictionaryKeyValuePair
  
  Set myDictionary = DictionaryKeyValuePair.Create(VBA.vbTextCompare, TextEncodingMethod.temAscii)
 
 
**Add, CompareMode, Count, Exists, Item, Key, Items, Keys, Remove, RemoveAll**

The same behaviour as the [Scripting.Dictionary object](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object)
  
## Compiler Constants

Note the compiler constants in the [Dictionary.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/Dictionary.cls) and [ScriptingDictionary.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/ScriptingDictionary.cls)

These compiler constants may require updating to True or False according to the availablity of the Scripting.Dictionary reference and platform.

**For Mac:**

#Const SCRIPTING_REFERENCE = False

#Const SCRIPTING_LATEBINDING = False

**For Windows:**

**If the Microsoft Scripting Runtime Scripting is referenced:**

#Const SCRIPTING_REFERENCE = True

**If the Microsoft Scripting Runtime Scripting not referenced but wish to use it late bound:**

#Const SCRIPTING_REFERENCE = False

#Const SCRIPTING_LATEBINDING = True

For Mac as the Scripting.Dictionary isn't available the Dictionary.Create uses the DictionaryKeyValuePair as a compatible alternative. 

On Windows if both compiler constants are set to False the Dictionary.Create uses the DictionaryKeyValuePair as an alternative. 

## Advantages

The VBA-IDictionary provides interfaces for the dictionary implementations provided, allowing for the easy transition when switching between dictionary implementations.  i.e. It allows for programming to an interface instead to a particular implementation which can be advantageous for unit testing etc.

Added support for keys of LongLong data type which is only availablue using the DictionaryKeyVluePair.cls IDictionary implementation and compatibile with VBA 7.  This also allows using LongPtr as keys as they are converted to LongLong or Long data types for earlier versions.

Significant performance compared to other VBA dictionary implementations, and in some scenerios provides a significant performance over the Scripting.Dictionary. 


## Performance

Great consideration has been given to provide as good as possible performance while using the underlying VBA.Collection.

See the [Excel VBA-IDictionaryPerformance spreadhsheet](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/VBA-IDictionary%20V2-1%20Sep%2003%2C%202019%20Performance.xlsx) for a performance comparisions of the ScriptingDictionary, DictionaryKeyValuePair and other VBA dictionary implementations.  Performance results displayed in the graphs can be filtered to compare various IScriptingDictionary implementations, for the various key and item data types and string key encoding options provided.

Performance testing can be performed using  the modules in the MS Access database provided, TestPerformanceKeyValuePairAdd.bas and TestPeformanceKeyValuePairItem.bas and results are displayed in the immediate window.  The results are currently manually copied into the peformance Excel spreadsheet using the Text Import Wizard. Those modules are still under development.

The DictionaryKeyValuePairs compared to other VBA dictionary implementations its significant performance improvement, especially when adding items, is likely due to not constantly maintaining an array of Items and Keys, and only populating them when requested. On future requests for Items and Keys they are only repopulated if changes have been made to the dictionary keys, and/or items.  This tradeoff results in the first iteration of Items and Keys to be fractionally slower and any subsequent requests without changes are at similar performance as not repopulated.

For Adding items, compared to the MS Scripting.Dictionary for datasets at approximately 350,000 items it starts to outperform.

For the scenerio for string keys and the compare mode is vbTextCompare at approximately 100,000 items it starts to outperform.

Overall the performance of the DictionaryKeyValuePair has significant improvements over other VBA dictionary implementations at approximately twice the performance for small datasets, and numerous times for large ones. It appears to have a gradual linear degraduation in performance the larger the dataset. 

I haven't recently got around to pushing it limitations, thou should be reasonably capable of handling datasets of two million items within a reasonable timeframe.  The Scripting.Dictionary noticeably degrades in performance for datasets over 500,000+ items.

For datasets continaing key and/or items that are objects whatever the data structure used, VBA's appears slow at dereferening of objects and destroying them as it is single threaded.  For this requirement it's best to keep datasets under 300,000 items as it takes considerable time to clean up.  The only other work around is when have large datasets containing objects to give them global scope or keep alive and push the cleaning up process to when the application closes. 

As always it's a matter of selecting the appropriate data structure for your requirements and not one suits all purposes. 

## Testing
Unit testing can be found in TestModuleDictionaryKeyValuePair.bas within the Ms Access database provided.  The VBA Addin [Rubberduck](http://rubberduckvba.com/) is required to run the unit testing. 

Only tested on the Windows platform and would be appreciated if anyone can test it on the Mac platform. I don't anticipate any compatiblity issues if the compiler constants are appropriately set.

## Notes

Support for keys of LongLong data type which is only availablue under VBA 7 using the DictionaryKeyValuePair.

Untested on Mac and VBA 6. 

## Version 2.1 Modifications
For performance improvements removed the use of a managed variant containing an Integer array to use a managed Integer Array.  Overall result was a slight performance improvement of approximately 5% for processing case-sensitive string keys.

Updated TextEncoderUnicode  ToHexString function to use a managed Char array which is manually freed when the class goes out of scope.

Added ManagedCharSafeArray.cls 

Added TypeSafeArray.bas

## Version 2.1 Modifications September 05, 2019

  - Updated: [TextEncoderUnicode.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/TextEncoderUnicode.cls)
  
     According to the ManagedCharSafeArray.cls updates.
 
  - Updated: [ManagedCharSafeArray.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/ManagedCharSafeArray.cls)
  

MangedCharSafeArray was restructured to more appropriately associate the managed Char array with its SafeArray descriptor allocated. This was to ensure only that the SafeArray descriptor is allocated to only one managed Char array at a time.
  
  - Updated to predeclared class and added a Create method. ie. 
 
    Public Function Create(ByRef outCharsArray() As Integer) As ManagedCharSafeArray

  - The Get Property was removed and a replace with the method:
 
    Public Sub SafeArrayAllocateToCharArray(ByRef outManagedChars() As Integer)
    
 ## Future Modifications
From the initial release have improved the performance of processing case-senstive string keys by appropriately 20% and there are now limited opportunities for future performance improvements.  Any suggestions for performance improvements or issues are highly welcomed.

Also testing is required on the Mac platform and VBA 6 which I'm unable to test. 
