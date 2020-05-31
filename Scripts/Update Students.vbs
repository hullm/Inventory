Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentFolder = objFSO.GetAbsolutePathName(".")
strCurrentFolder = strCurrentFolder & "\CSV\"
strCSV = strCurrentFolder & "Import.csv"

While txtSourceCSV.AtEndOfLine = False

Wend