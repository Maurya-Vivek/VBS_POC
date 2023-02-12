
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
strOutFileName = InputBox("Enter File Name here (with extension like outfile.doc):")
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WScript.Echo "Path at which utility is present " & scriptdir & " Your Replacement will take place now, Please click OK to proceed and wait till success message!!"
CreateOutputfile
Dim Inputvalu
Set Inputvalu = fs.OpenTextFile(scriptdir & "\Inputs.csv")
Dim counter, line, INPUT_ARRAY, Key_array, Valu_array
counter = 0
Do While Not Inputvalu.AtEndOfStream
    line = Inputvalu.ReadLine
    counter = counter + 1

    If counter > 1 Then
        INPUT_ARRAY = Split(line, ",")
        GENERIC_TEXT_REPLACE INPUT_ARRAY(0), INPUT_ARRAY(1)
    end if
loop
WScript.Echo "Hey Maurya!! your file is Ready with Replacements at: " & scriptdir & "\" & strOutFileName & " You can now open your File."
Inputvalu.close
Set Inputvalu = Nothing
Set fs = Nothing

Function CreateOutputfile()
' Function will create a copy of template file which can be edited later.
Set objWord = CreateObject("Word.Application")
objWord.Visible = False

Set objDoc = objWord.Documents.Open("C:\MY_DOCS\Study_docs\Chena_utility\Template.doc")
objDoc.SaveAs(scriptdir & "\" & strOutFileName)
objDoc.Close
objWord.Quit

End Function

Function GENERIC_TEXT_REPLACE(find_text,replace_text)
' Fuction will Replace the word in the MSWd document

Set objWord2 = CreateObject("Word.Application")
Set objDoc2 = objWord2.Documents.Open(scriptdir & "\" & strOutFileName)
Set objSelection = objWord2.Selection
objWord2.Visible = False
    FindText = find_text
    MatchCase = False
    MatchWholeWord = true
    MatchWildcards = False
    MatchSoundsLike = False
    MatchAllWordForms = False
    Forward = True
    Wrap = wdFindContinue
    Format = False
    wdReplaceNone = 0
    ReplaceWith = replace_text
    wdFindContinue = 1
    wdReaplaceAll = 2

    a = objSelection.Find.Execute(FindText,MatchCase,MatchWholeWord,MatchWildcards,MatchSoundsLike,MatchAllWordForms,Forward,Wrap,Format,ReplaceWith,wdReaplaceAll)

objDoc2.Save
objDoc2.Close
objWord2.Quit

GENERIC_TEXT_REPLACE = "Success"

End Function