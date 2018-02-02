<#

.SYNOPSIS
Creates a custom doc with text, link and macro.

.DESCRIPTION
Create-MacroDoc outputs a Mircosoft Office Word 1997 document. 
It allows the automation of multiple similar versions of files, allowing to test how slight differences will effect it.

.PARAMETER docPath
Full path for the output file.

.PARAMETER macroContentPath 
Full path of a text file containing macro to embed

.PARAMETER linkText 
Text to show in the URL embedded in the document body

.PARAMETER linkPath 
Address of the website (or local file) pointed by the link

.PARAMETER docText
Text to be shown in the document body

.EXAMPLE 
Create a document in the path c:\out\bla.doc with a macro stored in c:\macros\m.txt, in the document's body have the text "hello" and a link to google.com with the label "GOOGLE"
Create-MacroDoc -docPath "c:\out\bla.doc" -macroContentPath "c:\macros\m.txt" -docText "hello" -linkText "GOOGLE" -linkPath "https://www.google.com"

.NOTES
You need to have Microsoft Office installed in order to run this script.
#>


param (
    [Parameter(Mandatory=$true)][string]$docPath,
    [Parameter(Mandatory=$false)][string]$macroContentPath,
    [Parameter(Mandatory=$false)][string]$linkText = "Awesome free games inside, click me!",
    [Parameter(Mandatory=$false)][string]$linkPath = "https://www.facebook.com/groups/dc9723/",
    [Parameter(Mandatory=$false)][string]$docText = "Totally legit doc file, docx is for fools!"
 )

 $defaultMacro = @"
Private Sub Document_Open()
    MsgBox ("Hello .doc macro world")
End Sub
"@

 if (-Not ($macroContentPath)){
    $macroContent = $defaultMacro
} 
else {
    $macroContent = [IO.File]::ReadAllText($macroContentPath)
}

# create the Word COM object
$word = New-Object -ComObject word.application
$doc = $word.documents.add()

# add link
$range = $doc.Range()
$objLink = $doc.Hyperlinks.Add($range,$linkPath,"" , "", $linkText)

# add text
$selection = $word.selection
$selection.typeText($docText)
$selection.typeParagraph()

# saving the doc, last arg is reference to the enum type, doc
$doc.saveas([ref] $docPath, [ref] 0)
$word.quit()

# add macro, for some odd reason I needed to open it after it is saved, otherwise it did not work
$Word = New-Object -ComObject Word.Application
$Doc = $Word.Documents.Open($docPath)
$Doc.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString($macroContent)
$Doc.Close()
$Word.quit()
