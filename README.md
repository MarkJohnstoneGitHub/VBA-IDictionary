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

Note the compiler constants in the Dictionary.cls and ScriptingDictionary.cls are set to False
The compiler constants may require updating to True for user requirements and platform.

See:

#Const SCRIPTING_REFERENCE = False

#Const SCRIPTING_LATEBINDING = False

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

Note: Only tested on Windows and would be appreciated if anyone can test it on the Mac platform. 


Performance.

Great consideration has been given to provide as good as possible performance while using the underlying VBA.Collection.  See the Excel VBA-IDictionaryPerformance spreadhsheet for a performance comparisions of the ScriptingDictionary, DictionaryKeyValuePair and other VBA dictionary implementations. 


