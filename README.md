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

Compiler Constants:

Note the compiler constants in the Dictionary.cls and ScriptingDictionary.cls 
These compiler constants may require updating to True or False for user requirements and platform.

For Mac:

#Const SCRIPTING_REFERENCE = False

#Const SCRIPTING_LATEBINDING = False

For Windows:

If the Microsoft Scripting Runtime Scripting is referenced:

#Const SCRIPTING_REFERENCE = True

If the Microsoft Scripting Runtime Scripting not referenced but wish to use it late bound:

#Const SCRIPTING_REFERENCE = False

#Const SCRIPTING_LATEBINDING = True

For Mac as the Scripting.Dictionary isn't available the Dictionary.Create uses the DictionaryKeyValuePair as a compatible alternative. 

On Windows if both compiler constants are set to False the Dictionary.Create uses the DictionaryKeyValuePair as an alternative. 

Note: Only tested on Windows and would be appreciated if anyone can test it on the Mac platform. I don't anticipate any compatiblity issues if the compiler constants are appropriately set.


## Advantages

The VBA-IDictionary provides interfaces on all dictionary implementations provided, allowing for easy transition switching between the required dictionary implementation.  i.e. It allows for programming to an interface instead to a particular implementation. This can be advantageous for things such as unit testing etc.


## Performance

Great consideration has been given to provide as good as possible performance while using the underlying VBA.Collection.  See the Excel VBA-IDictionaryPerformance spreadhsheet for a performance comparisions of the ScriptingDictionary, DictionaryKeyValuePair and other VBA dictionary implementations.

The DictionaryKeyValuePairs compared to other VBA dictionary implementations its significant performance improvement, especially when adding items, is likely due to not constantly maintaining an array of Items and Keys and only populating them when requested. On future requests for Items and Keys they are only repopulated if changes have been made to the dictionary keys, and/or items.  This tradeoff results in the first iteration of Items and Keys to be fractionally slower and any subsequent requests without changes are at similar performance as not repopulated.

For Adding items:

Compared to the MS Scripting.Dictionary for datasets at approximately 350,000 items it starts to outperform. 
For the scenerio for string keys and the compare mode is vbTextCompare at approximatel 100,000 items it starts to outperform.

Overall the performance of the DictionaryKeyValuePair has significant improvements over other VBA dictionary implementations at approximately twice the performance for small datasets, and numerous times for large ones. It appears to have a gradual linear degraduation in performance in comparision to expotientially degrading. 

I haven't recently got around to pushing it limitations thou should be reasonably capable of handling datasets of two million items in a reasonable timeframe.  The Scripting.Dictionary noticeable degrades in performance for datasets over 500,000+ items.

For datasets continaing key and/or items that are objects note whatever the datastructure VBA's garbage collection is slow for cleaning up objects.

As always it's a matter of selecting the appropriate datastruce for your requirements and not one suits all purposes. 


## Notes

For VBA-IDictionary v2.1 the enumeration of DictionaryKeyValuePairs is on Keys, previously it had been on a key, value pair in a one dimensional array where the first element i.e. dictEntry(0) was the key and the second i.e. dictEntry(1) was the dictionary item.  This was changed for consistency with the Scripting.Dictionary behaviour.  Possibly in future will add an option to decided to enumerate on keys or key, value pairs.  General best practice is to specify to enumerate on Keys instead of enumerating on the dictionary object.

The performance of enumerating on the tuple of key,value pairs verses enumerating on Keys to obtain its associated item would expect to be quicker, as not requesting an item via the dictionary. 

Eg. *assuming only scalar values for dictionary keys and items

For each dictkey in myDictionary.Keys

  dictItem = Item(dictkey)
  
Next

Verses the enumeration of key,value pairs

For each dictEntry in myDictionary

  dictKey = dictEntry(0)
  
  dictItem = dictEntry(1)
  
Next


## Future Extensions

For Version 2.0 Possibly rename the IDictionary.cls to IScriptingDictionary.cls and updating all classes implementing the IDictionary interface according.  This is to better reflect that the interface conforms to the public interface of the MS Scripting Runtime Dictionary and avoid confusion with other IDictionary implementations published. 

Will be examining adding an IList interface to integrate the behaviour of various datastructures.

Possible explore typed a <TKey,TValue> Dictionary.  




