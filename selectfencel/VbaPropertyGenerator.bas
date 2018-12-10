Attribute VB_Name = "VbaPropertyGenerator"
Option Explicit

  '// Just make a big list of variable names and types, separated
  '// by a semicolon, ie:
  '//
  '// "Prop1;String;Prop2;Integer;Prop3;Double"
  '// GREG - Added class initializer for the objects initialization and Class name
                        
Private Const PropertyList$ = "LVStays;Collection;LVStruts;Collection;MVStays;Collection;MVStruts;Collection;TRFRCircles;Collection;TRFRNameplates;Collection;LVpoles;Collection;MVpoles;Collection;MVSharingPoles;Collection;KickerPoles;Collection;LVFlyingStays;Collection;LVShortStays;Collection;MVFlyingStays;Collection;MVShortStays;Collection;MVFuseIsolators;Collection;TRFRs;Collection;houseText;Collection;TrfrShapes;Collection;Airdac;Collection;LVConductor;Collection;MVConductor;Collection;elseCellElements;Collection;elseTextElements;Collection;elseLineElements;Collection;elseShapeElements;Collection;elseElements;Collection"
Private Const ClassName$ = "clsDataZone"
Private Const Indent$ = vbTab
Private Const NewLine$ = vbCrLf
Private Const PrivatePrefix$ = "m"
Private Const PrivateSuffix$ = ""
Private Const FirstLetterOfPrivateVarLowerCase = False


Private Const NativeTypes$ = "Boolean,Byte,Double,Integer,Long,Single,String,Element"

Sub VbaPropertyGenerator()

  '// Note: the functions called from this loop could be more
  '//       efficient by doing more in the loop, but I opted to
  '//       make the code more modular, so propsAndTypes, and the
  '//       prefix, suffix and other options are passed to many
  '//       of the other functions in entirety.
  
  Dim propsAndTypes As Variant
  propsAndTypes = Split(PropertyList$, ";")
  
  Dim sPropName As String, sTypeName As String
  Dim sPrivateItem As String, sPropertyItem As String, sPrivateItemDec As String
  Dim sPrivateBlock As String, sPropertyBlock As String
  
  '// Build all of the private variable declarations and public properties:
  Dim i As Integer
  For i = 0 To UBound(propsAndTypes) Step 2
  
    '// Get a prop name and data type from our name-type array:
    sPropName = propsAndTypes(i)
    sTypeName = propsAndTypes(i + 1)
    
    '// Build and append the private variable declaration:
    sPrivateItem = m_getPrivateVariableName(sPropName, sTypeName, PrivatePrefix, PrivateSuffix, FirstLetterOfPrivateVarLowerCase)
    sPrivateItemDec = "Private " & sPrivateItem & " As " & sTypeName
    sPrivateBlock = sPrivateBlock & NewLine & sPrivateItemDec
    
    '// Build and append the public property:
    sPropertyItem = m_getPublicProperty(sPropName, sPrivateItem, sTypeName)
    sPropertyBlock = sPropertyBlock & NewLine & sPropertyItem

  Next i
    
  '// Build the initializer for Set or object variables
  Dim sClassInitBlock As String
  sClassInitBlock = m_getClassInitializeBlock(propsAndTypes, PrivatePrefix, PrivateSuffix, FirstLetterOfPrivateVarLowerCase)
  '// Build the ToString function:
  Dim sToStringBlock As String
  sToStringBlock = m_getToStringFunction(propsAndTypes, PrivatePrefix, PrivateSuffix, FirstLetterOfPrivateVarLowerCase)
  
  '// Build the Class_Terminate block:
  Dim sClassTermBlock, outputStr As String
  sClassTermBlock = m_getClassTerminateBlock(propsAndTypes, PrivatePrefix, PrivateSuffix, FirstLetterOfPrivateVarLowerCase)
  
  '// Output general class header info:
  outputStr = "Option Explicit" & NewLine
  
  outputStr = outputStr & "'// Class: " & ClassName & NewLine
  outputStr = outputStr & "'// A class for..." & NewLine
    
  
  '// Output the private variable declarations, the public
  '// properties and the class termination procedure:
  outputStr = outputStr & sPrivateBlock & NewLine
  outputStr = outputStr & sClassInitBlock & NewLine
  outputStr = outputStr & sPropertyBlock & NewLine
  outputStr = outputStr & sToStringBlock & NewLine
  outputStr = outputStr & sClassTermBlock & NewLine
  
  'Debug.Print outputStr
  'printToFile outputStr
  DataScanner.outputBox.Text = outputStr
  copyToClipboard outputStr
  
End Sub

Sub printToFile(textOut As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile("C:\Users\Gregory Mavhunga\Documents\RSAT - Gareth Stanford")
    oFile.WriteLine textOut
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Sub

Sub copyToClipboard(str As String)
    Dim d As UserForm
    Dim dataObj As DataObject
    Set dataObj = New DataObject
    dataObj.SetText str
    dataObj.PutInClipboard
    Set dataObj = Nothing
    Set d = Nothing
End Sub
Private Function m_getPublicProperty(ByVal propertyName As String, _
                                     ByVal privateName As String, _
                                     ByVal dataTypeName As String) As String

  Dim sPropBlock As String
  Dim sLetOrSet  As String, sSet As String

  '// Set variables depending on if the data type requires a 'Let' or
  '// 'Set', and if the variables require 'Set' in order to be...well...set:
  If (m_isSetVariable(dataTypeName)) Then
    sSet = "Set "
    sLetOrSet = "Set "
  Else
    sSet = vbNullString
    sLetOrSet = "Let "
  End If
  
  '// Create the Get part of the property:
  sPropBlock = sPropBlock & "Public Property Get " & propertyName & "() As " & dataTypeName & NewLine
  sPropBlock = sPropBlock & Indent & sSet & propertyName & " = " & privateName & NewLine
  sPropBlock = sPropBlock & "End Property" & NewLine
  
  '// Append the Let/Set part of the property:
  sPropBlock = sPropBlock & "Public Property " & sLetOrSet & propertyName & "(value as " & dataTypeName & ")" & NewLine
  sPropBlock = sPropBlock & Indent & sSet & privateName & " = value" & NewLine
  sPropBlock = sPropBlock & "End Property" & NewLine
  
  m_getPublicProperty = sPropBlock
    
End Function
Private Function m_getPrivateVariableName(ByVal propertyName As String, _
                                          ByVal dataTypeName As String, _
                                          ByVal prefix As String, _
                                          ByVal suffix As String, _
                                          ByVal makeFirstLetterLowerCase As Boolean) As String

  Dim privateName As String
  privateName = propertyName
  
  If (makeFirstLetterLowerCase) Then
    privateName = LCase(VBA.Left(privateName, 1)) & Mid(privateName, 2)
  Else
    privateName = UCase(VBA.Left(privateName, 1)) & Mid(privateName, 2)
  End If
  
  privateName = prefix & privateName & suffix
  
  m_getPrivateVariableName = privateName

End Function



Private Function m_getClassTerminateBlock(ByRef propsAndTypes As Variant, _
                                          ByVal prefix As String, _
                                          ByVal suffix As String, _
                                          ByVal firstLetterLowerCase As Boolean) As String
                                          
  '// Creates a Class_Terminate block that sets all non-native-
  '// data-type variables to nothing. A handy clean-up subroutine
  '// that you don't have to do yourself!

  Dim sTermBlockStart As String, sTermBlockMiddle As String, sTermBlockEnd As String
  
  sTermBlockStart = "Private Sub Class_Terminate()"
  sTermBlockMiddle = vbNullString
  sTermBlockEnd = "End Sub"
    
  Dim sPropName As String, sTypeName As String, sPrivateName As String
  
  '// Build all of the private variable declarations and public properties:
  Dim i As Integer, iCt As Integer
  For i = 0 To UBound(propsAndTypes) Step 2
  
    '// Get a prop name and data type from our name-type array:
    sPropName = propsAndTypes(i)
    sTypeName = propsAndTypes(i + 1)
    
    '// See if the variable is of the 'Set' type:
    If (m_isSetVariable(sTypeName)) Then
      
      iCt = iCt + 1
      
      sPrivateName = m_getPrivateVariableName(sPropName, sTypeName, prefix, suffix, firstLetterLowerCase)
      
      '// Add a 'Set x = nothing' line:
      If (sTermBlockMiddle <> vbNullString) Then
        sTermBlockMiddle = sTermBlockMiddle & NewLine
      End If
      sTermBlockMiddle = sTermBlockMiddle & Indent & "Set " & sPrivateName & " = Nothing '//...(Type = " & sTypeName & ")"
      
    End If

  Next i
  
  If (iCt > 0) Then
    m_getClassTerminateBlock = sTermBlockStart & NewLine & sTermBlockMiddle & NewLine & sTermBlockEnd
  Else
    m_getClassTerminateBlock = vbNullString
  End If
  
End Function
Private Function m_getClassInitializeBlock(ByRef propsAndTypes As Variant, _
                                          ByVal prefix As String, _
                                          ByVal suffix As String, _
                                          ByVal firstLetterLowerCase As Boolean) As String
                                          
  '// Creates a Class_Terminate block that sets all non-native-
  '// data-type variables to nothing. A handy clean-up subroutine
  '// that you don't have to do yourself!

  Dim sTermBlockStart As String, sTermBlockMiddle As String, sTermBlockEnd As String
  
  sTermBlockStart = "Private Sub Class_Initialize()"
  sTermBlockMiddle = vbNullString
  sTermBlockEnd = "End Sub"
    
  Dim sPropName As String, sTypeName As String, sPrivateName As String
  
  '// Build all of the private variable declarations and public properties:
  Dim i As Integer, iCt As Integer
  For i = 0 To UBound(propsAndTypes) Step 2
  
    '// Get a prop name and data type from our name-type array:
    sPropName = propsAndTypes(i)
    sTypeName = propsAndTypes(i + 1)
    
    '// See if the variable is of the 'Set' type:
    If (m_isSetVariable(sTypeName)) Then
      
      iCt = iCt + 1
      
      sPrivateName = m_getPrivateVariableName(sPropName, sTypeName, prefix, suffix, firstLetterLowerCase)
      
      '// Add a 'Set x = nothing' line:
      If (sTermBlockMiddle <> vbNullString) Then
        sTermBlockMiddle = sTermBlockMiddle & NewLine
      End If
      sTermBlockMiddle = sTermBlockMiddle & Indent & "Set " & sPrivateName & " = New " & sTypeName & "'//...(Type = " & sTypeName & ")"
      
    End If

  Next i
  
  If (iCt > 0) Then
    m_getClassInitializeBlock = sTermBlockStart & NewLine & sTermBlockMiddle & NewLine & sTermBlockEnd
  Else
    m_getClassInitializeBlock = vbNullString
  End If
  
End Function
Private Function m_getToStringFunction(ByRef propsAndTypes As Variant, _
                                       ByVal prefix As String, _
                                       ByVal suffix As String, _
                                       ByVal firstLetterLowerCase As Boolean) As String
                                       
  '// Creates a VB.NET style ToString() function that will return
  '// a concatenated list of all of the properties in the object.
                                       
  Dim sItemSeparator As String
  sItemSeparator = "VbCrLf"
  
  Dim sFunctionStart As String
  sFunctionStart = "Public Function ToString() as String" & NewLine & NewLine & _
                    "'This function needs to be custom made in any case, this works for non -object members" & NewLine
                   'Indent & "'// Note: not all of the variables in this class will" & NewLine &
                   'Indent & "'//       necessarily be convertible to strings. In these" & NewLine &
                   'Indent & "'//       cases, you will get errors in this function." & NewLine &
                   'Indent & "'//       In these cases, either remove the line that contains" & NewLine &
                   'Indent & "'//       a variable that can't be converted to string, or" & NewLine &
                   'Indent & "'//       substitute some option that can be, such as" & NewLine &
                   'Indent & "'//       m_someObject.String or m_someObject.Name" & NewLine & NewLine
                   
  sFunctionStart = sFunctionStart & Indent & "Dim s As String" & NewLine & NewLine & Indent & "s = "

  Dim sFunctionEnd As String
  sFunctionEnd = NewLine & Indent & "ToString = s" & NewLine & NewLine & "End Function"
  
  Dim sToStringFunction As String

  Dim sFunctionMiddle As String
  sFunctionMiddle = vbNullString
  
  Dim sPropName As String, sTypeName As String, sPrivateName As String
  Dim i As Integer
  For i = 0 To UBound(propsAndTypes) Step 2
  
    '// Get a prop name and data type from our name-type array:
    sPropName = propsAndTypes(i)
    sTypeName = propsAndTypes(i + 1)
      
    '// Add new line and continuation stuff if the block is not empty:
    If (sFunctionMiddle <> vbNullString) Then
      sFunctionMiddle = sFunctionMiddle & " & " & sItemSeparator & " & _" & NewLine & Indent & "    "
    End If
    
    '// Get the private variable name:
    sPrivateName = m_getPrivateVariableName(sPropName, sTypeName, prefix, suffix, firstLetterLowerCase)
    
    sFunctionMiddle = sFunctionMiddle & Chr(34) & sPropName & " = " & Chr(34) & " & CStr(" & sPrivateName & ")"
    
  Next i
  
  '// Wrap the function with the begin and end parts of a function:
  sToStringFunction = sFunctionStart & sFunctionMiddle & NewLine & sFunctionEnd
                    
  '// Return the whole mess:
  m_getToStringFunction = sToStringFunction
  
End Function
Private Function m_isSetVariable(dataTypeName As String) As Boolean

  If (InStr(1, NativeTypes$, dataTypeName, vbTextCompare) > 0) Then
    m_isSetVariable = False
  Else
    m_isSetVariable = True
  End If

End Function








