# VBA-IDictionary

VBA dictionary which is Mac compatible which implements an IDictionary interface. 

Classes Required:
  IDictionary.cls
  Dictionary.cls
  ScriptingDictionary.cls
  DictionaryKeyValuePair.cls
   
Optional Reference Addin:

  Microsoft Scripting Runtime Scripting scrrun.dll

Usage:
  Dim myDictionary As IDictionary
  
  myDictionary = Dictionary.Create
  
  myDictionary = Dictionary.Create(IDictionaryType.ScriptingDictionary)

Compiler Constants:

Note the compiler constants in the Dictionary.cls and ScriptingDictionary.cls are set to False
The compiler constants may require updating to True for user requirements and platform.

See
#Const SCRIPTING_REFERENCE = False
#Const SCRIPTING_LATEBINDING = False
