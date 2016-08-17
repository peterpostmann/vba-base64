'========================================================== 
'== Base64 Encode/Decode
'== Based on: https://ghads.wordpress.com/2008/10/17/vbscript-readwrite-binary-encodedecode-base64/
'========================================================== 
option explicit 

' common consts
Const TypeBinary = 1, TypeText = 2
Const ForReading = 1, ForWriting = 2, ForAppending = 8

'========================================================== 
'== Get arguments 
'========================================================== 

If WScript.Arguments.Count = 0 Then
    WScript.Echo "Drag and drop a file onto the script."
    WScript.Quit
End If

Dim inputFile : inputFile = WScript.Arguments.Item(0)
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

If Not objFSO.FileExists(inputFile) Then
    WScript.Echo "Error: Input file '" & inputFile & "' not found!"
    WScript.Quit
End If

Dim inputData, outputData, outputFile

If LCase(objFSO.GetExtensionName(inputFile)) = "base64" Then
        
    inputData = readText(inputFile)
    outputData = decodeBase64(inputData)
    outputFile = objFSO.GetParentFolderName(inputFile) & "\" & objFSO.GetBaseName(inputFile)
    writeBytes outputFile, outputData

Else
    
    inputData = readBytes(inputFile)
    outputData = encodeBase64(inputData)
    outputFile = objFSO.GetAbsolutePathName(inputFile) & ".base64"
    writeText outputFile, outputData
    
End If


WScript.Quit

private function readBytes(file)
  dim inStream
  ' ADODB stream object used
  set inStream = WScript.CreateObject("ADODB.Stream")
  ' open with no arguments makes the stream an empty container 
  inStream.Open
  inStream.type= TypeBinary
  inStream.LoadFromFile(file)
  readBytes = inStream.Read()
end function

private function readText(file)
    Dim objFSO, oFile 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    Set oFile = objFSO.OpenTextFile(file, ForReading) 
    readText = oFile.ReadAll 
    oFile.Close 
end function
  
private function encodeBase64(bytes)
  dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set bytes, get encoded String
  EL.NodeTypedValue = bytes
  encodeBase64 = EL.Text
end function
  
private function decodeBase64(base64)
  dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set encoded String, get bytes
  EL.Text = base64
  decodeBase64 = EL.NodeTypedValue
end function
  
private function writeBytes(file, bytes)
  Dim binaryStream
  Set binaryStream = CreateObject("ADODB.Stream")
  binaryStream.Type = TypeBinary
  'Open the stream and write binary data
  binaryStream.Open
  binaryStream.Write bytes
  'Save binary data to disk
  binaryStream.SaveToFile file, ForWriting
End function

private function writeText(file, text)
    Dim objFSO, oFile 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    Set oFile = objFSO.OpenTextFile(file, ForWriting, True) 
    oFile.Write text 
    oFile.Close 
End function