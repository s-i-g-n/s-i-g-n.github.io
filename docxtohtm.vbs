

Option Explicit


'declare all variables
Dim objWord
Dim oDoc
Dim objFso
 
Const wdSaveFormat = 10 'for Filtered HTML output
 
'and proceed with the function
Set objFso = CreateObject("Scripting.FileSystemObject")

'create an instance of Word
Set objWord = CreateObject("Word.Application")
 
objWord.Visible = True

Dim absoluteFileName 
absoluteFileName = objFso.GetAbsolutePathName("index.docx")
 
objWord.Documents.Open absoluteFileName
'do all this in the background
' objWord.Visible = False

Dim baseFolder 
baseFolder = objFso.GetParentFolderName( absoluteFileName )

Dim htmOutputFile
htmOutputFile = baseFolder + "\" + "index.htm"



Set oDoc = objWord.ActiveDocument
oDoc.SaveAs htmOutputFile, wdSaveFormat
oDoc.Close
 
'close Word
objWord.Quit


