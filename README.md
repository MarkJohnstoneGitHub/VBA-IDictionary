# VBA-IDictionary v2.0 July 28

VBA dictionary which is Mac compatible which implements an IScriptingDictionary interface. 

**Classes Required:**
 - [IScriptingDictionary.cls]https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/IScriptingDictionary.cls)
 
 - [Dictionary.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/Dictionary.cls)
 
 - [DictionaryKeyValuePair.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/DictionaryKeyValuePair.cls)
 
 - [ScriptingDictionary.cls]( https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/ScriptingDictionary.cls)
 
 - [ITextEncoding.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/ITextEncoding.cls)
 
 - [TextEncoderASCII.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/TextEncoderASCII.cls)
 
  - [TextEncoderUnicode.cls](https://github.com/MarkJohnstoneGitHub/VBA-IDictionary/blob/master/scr/TextEncoderUnicode.cls)

   
Optional Reference Addin:  Microsoft Scripting Runtime Scripting scrrun.dll
 
## Usage

### Creating an IScriptingDictionary

**Creating a IScriptingDictionary with Dictionary.Create()**

The Dictionary.cls is an IScriptingDictionary factory class.  It returns an available IScriptingDictionary implementation, according to the compiler constants and/or the implemententation specified.  I.e. If the Scripting.Dictionary is specified and not available a DictionaryKeyValuePair will be returned.

The Dictionary.Create has three optional parameters:

 - dictionaryType : Specifies which IDictionary implementation to create. i.e. ScriptingDictionary or DictionaryKeyValuePairs. Default is IScriptingDictionaryType.isdtScriptingDictionary

 - compareMethod  : Specifies how string keys are handled either case sensitive or ignore case. Default value is vbBinaryCompare.

 - encodingMethod : Specifies which encoding method to use on case sensitive string keys i.e. Unicode or ASCII.  It is only applicable to the DictionaryKeyValuePairs class for improved performance, and  the option TextEncodingMethod.temAscii should only be used where string keys are ASCII compatible.  Default value is TextEncodingMethod.temUnicode.

Examples:

  Dim myDictionary As IScriptingDictionary
  
  Set myDictionary = Dictionary.Create
  
  Set myDictionary = Dictionary.Create(IScriptingDictionaryType.isdtScriptingDictionary, VBA.vbTextCompare)
  
  Set myDictionary = Dictionary.Create(IScriptingDictionaryType.isdtDictionaryKeyValuePair, VBA.vbBinaryCompare, TextEncodingMethod.temAscii)
  
**Creating using directly a IDictionary implementation provided**

Two IDictionary implementations have been provided, DictionaryKeyValuePair.cls and ScriptingDictionary.cls
 
The ScriptingDictionary.Create() has one optional parameter:

 - compareMethod : Specifies how string keys are handled either case sensitive or ignore case. Default value is vbBinaryCompare.
 
 Example
 
  Dim myDictionary As IDictionary 'or could use As ScriptingDictionary
  
  Set myDictionary = ScriptingDictionary.Create(VBA.vbTextCompare)
  
The DictionaryKeyValuePair.Create() has the two optional parameters:

 - compareMethod : Specifies how string keys are handled either case sensitive or ignore case.  Default value is vbBinaryCompare.
 
 - encodingMethod : Specifies which encoding text method to use on case sensitive string keys i.e. Unicode or ASCII. Default value is TextEncodingMethod.temUnicode
 
   Dim myDictionary As IDictionary 'or could use As DictionaryKeyValuePair
  
  Set myDictionary = DictionaryKeyValuePair.Create(VBA.vbTextCompare, TextEncodingMethod.temAscii)
 
 
**Add, CompareMode, Count, Exists, Item, Key, Items, Keys, Remove, RemoveAll**

The same as the [Scripting.Dictionary object](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object)
  

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

The VBA-IDictionary provides interfaces for the dictionary implementations provided, allowing for the easy transition when switching between dictionary implementations.  i.e. It allows for programming to an interface instead to a particular implementation which can be advantageous for unit testing.

Added support for keys of LongLong data type which is only availablue using the DictionaryKeyVluePair.cls IDictionary implementation and compatibile with VBA 7.  This also allows using LongPtr as keys as they are converted to LongLong or Long data types for earlier versions.


## Performance

Great consideration has been given to provide as good as possible performance while using the underlying VBA.Collection.

See the Excel VBA-IDictionaryPerformance spreadhsheet for a performance comparisions of the ScriptingDictionary, DictionaryKeyValuePair and other VBA dictionary implementations.  Performance results displayed in the graphs can be filtered to compare various IDictionary implements for the various key and item data types and string key processing options provided.

Performance testing can be performed using modules in the MS Access database provided, TestPerformanceKeyValuePairAdd.bas and TestPeformanceKeyValuePairItem.bas and results are displayed in the immediate window.  The results are currently manually copied into the peformance Excel spreadsheet using the Text Import Wizard. Those modules are still under development.

The DictionaryKeyValuePairs compared to other VBA dictionary implementations its significant performance improvement, especially when adding items, is likely due to not constantly maintaining an array of Items and Keys, and only populating them when requested. On future requests for Items and Keys they are only repopulated if changes have been made to the dictionary keys, and/or items.  This tradeoff results in the first iteration of Items and Keys to be fractionally slower and any subsequent requests without changes are at similar performance as not repopulated.

For Adding items, compared to the MS Scripting.Dictionary for datasets at approximately 350,000 items it starts to outperform.

For the scenerio for string keys and the compare mode is vbTextCompare at approximately 100,000 items it starts to outperform.

Overall the performance of the DictionaryKeyValuePair has significant improvements over other VBA dictionary implementations at approximately twice the performance for small datasets, and numerous times for large ones. It appears to have a gradual linear degraduation in performance in comparision to others that appear to expotientially degrade in performance. 

I haven't recently got around to pushing it limitations, thou should be reasonably capable of handling datasets of two million items within a reasonable timeframe.  The Scripting.Dictionary noticeable degrades in performance for datasets over 500,000+ items.

For datasets continaing key and/or items that are objects whatever the data structure used, VBA's appears slow at dereferening of objects and destroying them as it is single threaded.  For this requirement it's best to keep datasets under 300,000 items as it takes considerable time to clean up.  The only other work around is have large datasets containing objects to have global references and push the cleaning up process to when the application closes. 

As always it's a matter of selecting the appropriate data structure for your requirements and not one suits all purposes. 

## Testing
Unit testing can be found in TestModuleDictionaryKeyValuePair.bas within the Ms Access database provided.  The VBA Addin [Rubberduck](http://rubberduckvba.com/) is required to run the unit testing. 

Only tested on the Windows platform and would be appreciated if anyone can test it on the Mac platform. I don't anticipate any compatiblity issues if the compiler constants are appropriately set.


## Notes

Support for keys of LongLong data type which is only availablue using the DictionaryKeyVluePair.cls IScriptingDictionary implementation and compatibile with Mac, Windows, VBA 6, VBA 7.

Untested on Mac and VBA 6. 


## Future Extensions

For Version 2.0 possibly rename the IDictionary.cls to IScriptingDictionary.cls and updating all classes implementing the IDictionary interface according.  This is to better reflect that the interface conforms to the public interface of the MS Scripting Runtime Dictionary and avoid confusion with other IDictionary implementations published.

Improve performance of Unicode encoding of case-senstive keys using a read-only Integer Array Overlay.  Initial testing displays an approximately 15 percent improvement.  Require to make Mac compatibile by adding in the appropriate Mac memory API's. 

Will be examining adding an IList interface to integrate the behaviour of various data structures.

Possible explore typed a <TKey,TValue> Dictionary.  




