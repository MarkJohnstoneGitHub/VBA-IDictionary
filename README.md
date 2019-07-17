# VBA-IDictionary

VBA dictionary which is Mac compatible which implements an IDictionary interface. 

Classes Required:
  IDictionary.cls
  Dictionary.cls
  ScriptingDictionary.cls
  DictionaryKeyValuePair.cls
   
Optional Reference Addin:  Microsoft Scripting Runtime Scripting scrrun.dll

## Usage

  Dim myDictionary As IDictionary
  
  Set myDictionary = Dictionary.Create
  
  Set myDictionary = Dictionary.Create(IDictionaryType.ScriptingDictionary)

## Compiler Constants

Note the compiler constants in the Dictionary.cls and ScriptingDictionary.cls 
These compiler constants may require updating to True or False according to requirements and platform.

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

For VBA-IDictionary v2.1 the enumeration of DictionaryKeyValuePairs is on Keys, previously it had been on a key, value pair in a one dimensional array where the first element i.e. dictEntry(0) was the key and the second i.e. dictEntry(1) was the dictionary item.  This was changed for consistency with the Scripting.Dictionary behaviour.  Possibly in future will add an option to decided to enumerate on keys or key, value pairs.  General best practice is to specify to enumerate on Keys instead of enumerating on the dictionary object.

The performance of enumerating on the tuple of key,value pairs verses enumerating on Keys to obtain its associated item would expect to be quicker, as not requesting an item via the dictionary. 

Eg. *assuming only scalar values for dictionary keys and items

For Each dictkey in myDictionary.Keys

   dictItem = myDictionary.Item(dictkey)
  
Next

Verses the enumeration of key,value pairs 

Note only for DictionaryKeyValuePairs where if for the Property Get NewEnum() As IUnknown , Set NewEnum = this.keyItemPairs etc

For Each dictEntry in myDictionary

   dictKey = dictEntry(0)
  
   dictItem = dictEntry(1)
  
Next


## Future Extensions

For Version 2.0 possibly rename the IDictionary.cls to IScriptingDictionary.cls and updating all classes implementing the IDictionary interface according.  This is to better reflect that the interface conforms to the public interface of the MS Scripting Runtime Dictionary and avoid confusion with other IDictionary implementations published.

Will be examining adding an IList interface to integrate the behaviour of various datastructures.

Possible explore typed a <TKey,TValue> Dictionary.  




