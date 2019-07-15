# VBA-IDictionary

VBA dictionary which is Mac compatible which implements an IDictionary interface. 

Classes Required:
  IDictionary.cls
  Dictionary.cls
  ScriptingDictionary.cls
  DictionaryKeyValuePair.cls
   
Optional Reference Addin:  Microsoft Scripting Runtime Scripting scrrun.dll

Usage:

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

Performance.

Great consideration has been given to provide as good as possible performance while using the underlying VBA.Collection.  See the Excel VBA-IDictionaryPerformance spreadhsheet for a performance comparisions of the ScriptingDictionary, DictionaryKeyValuePair and other VBA dictionary implementations.


Notes: 
For Version 2.0 Aiming to rename the IDictionary.cls to IScriptingDictionary.cls and updating all classes implementing the IDictionary interface according.  This is to better reflect that the interface conforms to the public interface of the MS Scripting Runtime Dictionary. 

Also as planning to use the IDictionary interface for extended features returning key, value pairs. 


