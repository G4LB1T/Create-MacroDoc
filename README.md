# Create-MacroDoc
Create-MacroDoc outputs a Mircosoft Office Word 1997 document. 
It allows the automation of multiple similar versions of files, allowing to test how slight differences will effect it.

## Parameters
### docPath
Full path for the output file.

### macroContentPath 
Full path of a text file containing macro to embed

### linkText 
Text to show in the URL embedded in the document body

### linkPath 
Address of the website (or local file) pointed by the link

### docText
Text to be shown in the document body

## Usage Example 
Create a document in the path c:\out\bla.doc with a macro stored in c:\macros\m.txt, in the document's body have the text "hello" and a link to google.com with the label "GOOGLE"
```
Create-MacroDoc -docPath "c:\out\bla.doc" -macroContentPath "c:\macros\m.txt" -docText "hello" -linkText "GOOGLE" -linkPath "https://www.google.com"
```

## Note
You need to have Microsoft Office installed in order to run this script.
Modern versions of Windows\Office require to set the following registry key in order to allow the Word COM object to edit VBA Object in
```
HKEY_CURRENT_USER\Software\Microsoft\Office\<OfficeVersion>\Word\Security\AccessVBOM
```

