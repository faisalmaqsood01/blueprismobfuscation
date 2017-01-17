Randomize 1
rndMax = 120
rndMin = -120
rndHgt = 60
rndWdt = 60
salt = "one"

StrFileName = "./object.xml"
result = ""
types = "|Data|Action|Decision|Recover|Exception|Resume|Calculation|Anchor|SubSheet|Note|"

Set ObjFso = CreateObject("Scripting.FileSystemObject")
Set ObjFile = ObjFso.OpenTextFile(StrFileName)
MyVar = ObjFile.ReadAll
ObjFile.Close

' WScript.Echo MyVar

set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.setProperty "SelectionLanguage", "XPath"
xmlDoc.async="false"
xmlDoc.load(StrFileName)

strXPath = "/process/stage"
Set processes = xmlDoc.documentElement.childNodes

set xmlDoc2 = CreateObject("Microsoft.XMLDOM")

Set stages = xmlDoc.selectNodes("/process/stage")


  For Each stage in stages

  xmlDoc2.loadXML( stage.xml )

  If InStr( types, ( "|" & stage.getAttribute("type") & "|" )) > 0  Then


    xmlDoc2.selectSingleNode("stage/displayx").Text = Int((rndMax-rndMin+1)*Rnd+rndMin)
    xmlDoc2.selectSingleNode("stage/displayy").Text = Int((rndMax-rndMin+1)*Rnd+rndMin)
    xmlDoc2.selectSingleNode("stage/displayheight").Text = rndHgt
    xmlDoc2.selectSingleNode("stage/displaywidth").Text = rndWdt
    xmlDoc2.selectSingleNode("stage/narrative").Text = ""
    name = xmlDoc2.selectSingleNode("stage").getAttribute("name")
    xmlDoc2.selectSingleNode("stage").setAttribute  "name", toAlphaOnly(hashString(name))


ElseIf Not InStr( "|Start|End|", ( "|" & stage.getAttribute("type") & "|" )) > 0  Then
result = result & stage.xml
End If

    ' Msgbox( xmlDoc2.xml )

    'xmlDoc2.attributes.getNamedItem("")
    ' MsgBox stage.getAttribute("stageid")

    If stage.getAttribute("type") = "Calculation" Then


      name = xmlDoc2.selectSingleNode("stage/calculation").getAttribute("stage")
      xmlDoc2.selectSingleNode("stage/calculation").setAttribute "stage", toAlphaOnly(hashString(name))
      expression = xmlDoc2.selectSingleNode("stage/calculation").getAttribute("expression")
      xmlDoc2.selectSingleNode("stage/calculation").setAttribute "expression", hashDataNames(expression)
End If


If stage.getAttribute("type") = "Decision" Then


    expression = xmlDoc2.selectSingleNode("stage/decision").getAttribute("expression")
    xmlDoc2.selectSingleNode("stage/decision").setAttribute "expression", hashDataNames(expression)
End If
If InStr( "|Start|Action|", ( "|" & stage.getAttribute("type") & "|" )) > 0  Then

  Set inputs = xmlDoc2.selectNodes("stage/inputs/input")

  For Each input in inputs

    stageName = input.getAttribute("stage")
    input.setAttribute "stage", toAlphaOnly( hashString( stageName ) )

    If stage.getAttribute("type") = "Action" Then
    expr = input.getAttribute("expr")
    input.setAttribute "expr", hashDataNames( expr )
    ' name = input.getAttribute("name")
    ' input.setAttribute "name", toAlphaOnly( hashString( name ) )
    input.setAttribute "narrative", ""

    End If



  Next
End If





If InStr( "|End|Action|", ( "|" & stage.getAttribute("type") & "|" )) > 0  Then


Set outputs = xmlDoc2.selectNodes("stage/outputs/output")

For Each output in outputs

  stageName = output.getAttribute("stage")
  output.setAttribute "stage", toAlphaOnly( hashString( stageName ) )


      If stage.getAttribute("type") = "Action" Then

      ' name = output.getAttribute("name")
      ' output.setAttribute "name", toAlphaOnly( hashString( name ) )
      output.setAttribute "narrative", ""

      End If
Next


End If

If stage.getAttribute("type") = "Exception" Then


  Set exception = xmlDoc2.selectSingleNode("stage/exception")



    detail = exception.getAttribute("detail")
    exception.setAttribute "detail", hashDataNames( detail )

End If
result = result & xmlDoc2.xml

Next

  Const ForReading = 1
  Const ForWriting = 2
  Const ForAppending = 8

  Set objTextFile = objFSO.OpenTextFile("./result.xml", ForWriting, True)

  ' Writes strText every time you run this VBScript
  objTextFile.WriteLine(result)
  objTextFile.Close
  Set ObjFso = Nothing





function toAlphaOnly( string )

  alphaChars = ""

  Set re = New RegExp
  With re
     .Pattern    = "[A-z]"
     .IgnoreCase = False
     .Global     = True
  End With

  Set chars = re.Execute( string )

  For Each char in chars
    alphaChars = alphaChars & char.Value
  Next

  toAlphaOnly = alphaChars
end function

function hashDataNames( string )

Set re = New RegExp
  With re
     .Pattern    = "\[[A-z0-1\s\-]+\]"
     .IgnoreCase = False
     .Global     = True
  End With

  Set data = re.Execute( string )

  For Each datum in data
  MsgBox (datum.Value)
  string = Replace( string, datum.Value, "[" & toAlphaOnly( hashString( datum.Value ) ) & "]" )
  Next

  hashDataNames = string

end function

function hashString( string )
  hashString = BytesToBase64(md5hashBytes(stringToUTFBytes(salt & string)))
end function

function md5hashBytes(aBytes)
    Dim MD5
    set MD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    MD5.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    md5hashBytes = MD5.ComputeHash_2( (aBytes) )
end function

function sha1hashBytes(aBytes)
    Dim sha1
    set sha1 = CreateObject("System.Security.Cryptography.SHA1Managed")

    sha1.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    sha1hashBytes = sha1.ComputeHash_2( (aBytes) )
end function

function sha256hashBytes(aBytes)
    Dim sha256
    set sha256 = CreateObject("System.Security.Cryptography.SHA256Managed")

    sha256.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    sha256hashBytes = sha256.ComputeHash_2( (aBytes) )
end function

function stringToUTFBytes(aString)
    Dim UTF8
    Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    stringToUTFBytes = UTF8.GetBytes_4(aString)
end function

function bytesToHex(aBytes)
    dim hexStr, x
    for x=1 to lenb(aBytes)
        hexStr= hex(ascb(midb( (aBytes),x,1)))
        if len(hexStr)=1 then hexStr="0" & hexStr
        bytesToHex=bytesToHex & hexStr
    next
end function

Function BytesToBase64(varBytes)
    With CreateObject("MSXML2.DomDocument").CreateElement("b64")
        .dataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = .Text
    End With
End Function

Function GetBytes(sPath)
    With CreateObject("Adodb.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .LoadFromFile sPath
        .Position = 0
        GetBytes = .Read
        .Close
    End With
End Function
