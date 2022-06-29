'********************************************************************************************
'   Count different items in a text file (Example: apple,banana,apple >>>> apple,2  banana,1)
'   Date:28/Jun/2022
'   Author: edcruces99@gmail.com
'********************************************************************************************

'Declare Array1
Set Array1 = CreateObject("System.Collections.ArrayList")

'Read input file
'filename = WScript.Arguments.Item(0)
filename="input.txt"

'Create file object
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set f = objFSO.OpenTextFile(filename)

'Load Array1
Do Until f.AtEndOfStream
   Array1.add f.ReadLine
Loop
f.Close

'Display Array1
'For Each item In Array1
'    WScript.Echo item
'Next 
'Wscript.Quit

'Dictionary
Set d = CreateObject("Scripting.Dictionary")
i=0
For Each item In Array1
    d(Array1(i)) = d(Array1(i)) + 1
    i=i+1
Next

'Dictionary Keys only
d1 = d.Keys()
Set Array2 = CreateObject("System.Collections.ArrayList")
For Each item In d1
      Array2.Add item
Next

'Sort
Array1.Sort
Array2.Sort

'Count each item
i=0
counter=0
aux=Array1.item(0)
n=0

Set objFSO=CreateObject("Scripting.FileSystemObject")
outFile="output.txt"
Set objFile=objFSO.CreateTextFile(outFile,True)

For Each item In Array1
    If(aux = Array1(i)) then
	   counter=counter+1
	Else
	   objFile.Write Array2(n) & "," & counter & vbCrLf
	   n=n+1
	   counter=1
	   aux=Array1.item(i)
	End if
	i=i+1
	If (i=Array1.count) Then
	    objFile.Write Array2(n) & "," & counter & vbCrLf
		Exit For
	End if
Next
objFile.Close
