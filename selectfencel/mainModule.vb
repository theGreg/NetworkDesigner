Attribute VB_Name = "mainModule"
'Dim gArray1 As Collection
'Set gArray1 = New Collection

Public Function selectedShapeOK() As Boolean
   Dim oElEnum As ElementEnumerator
   Dim oEl As Element
   Dim count As Integer
   selectedShapeOK = False
   Set oElEnum = ActiveModelReference.GetSelectedElements
   oElEnum.Reset
   Do While oElEnum.MoveNext
      Set oEl = oElEnum.Current
        If matchParams(oEl, "Level 11", 7) Then
            selectedShapeOK = True
            Exit Do
        End If
    Loop
End Function

Public Function validFenceDefined() As Boolean
    Dim valid, noSelectedElem As Boolean
    valid = False
    noSelectedElem = False
    If ActiveDesignFile.Fence.IsDefined Then
        valid = True
    Else
        'createFence
        If Not (ActiveModelReference.AnyElementsSelected) Then
           valid = False
           noSelectedElem = True
        Else
           'check if selected element fits criteria
           If selectedShapeOK Then
                'create fence from it
                createFenceFromShape
                'check again if fence has been created
                If ActiveDesignFile.Fence.IsDefined Then
                    valid = True
                End If
           Else
                MsgBox "No valid shape defined"
           End If
        End If
    End If
    
    validFenceDefined = valid
End Function

Public Function createFenceFromShapeElement(ByRef sElement) As Boolean
   
    createFenceFromShapeElement = False
    CadInputQueue.SendCommand "LOCK FENCE INSIDE"

    '   Set a variable associated with a dialog box
    SetCExpressionValue "tcb->msToolSettings.fence.placeMode", 3, ""

    Dim lastView As View
    Set lastView = CommandState.lastView

    Dim ofence As Fence
    Set ofence = ActiveDesignFile.Fence
    ofence.DefineFromElement lastView, sElement
    
    If ofence.IsDefined Then
        createFenceFromShapeElement = True
    End If
End Function
Public Sub createFenceFromShape()
    
    Dim oElEnum As ElementEnumerator
    Dim oelm As Element
    
    Set oElEnum = ActiveModelReference.GetSelectedElements
    oElEnum.Reset
    Do While oElEnum.MoveNext
       Set oelm = oElEnum.Current
    Loop

    CadInputQueue.SendCommand "LOCK FENCE INSIDE"

    '   Set a variable associated with a dialog box
    SetCExpressionValue "tcb->msToolSettings.fence.placeMode", 3, ""

    Dim lastView As View
    Set lastView = CommandState.lastView

    Dim ofence As Fence
    Set ofence = ActiveDesignFile.Fence
    ofence.DefineFromElement lastView, oelm
    
End Sub
Public Function matchParams(oElem As Element, lvlName As String, elmType As MsdElementType)
    
    Dim oSFlag, oLvlFlag, oColFlag, oTypeFlag As Boolean
    oSFlag = False
    oLvlFlag = False
    oColFlag = False
    oTypeFlag = False
    
    If oElem.type = elmType Then
        oTypeFlag = True
    End If
    If oElem.Level.Name = lvlName Then
        oTypeFlag = True
    End If
    
    matchParams = oTypeFlag 'And oLvlFlag And oColFlag
    
End Function
Sub selectAll()
    CadInputQueue.SendKeyin "LOCK FENCE INSIDE"
    CadInputQueue.SendKeyin "Fence Selection New"
   
End Sub
Function getAverage(ByVal elms As Double, ByVal Length As Double) As Double
    getAverage = (Length / elms)
End Function
Sub getStatsConductor()
    Dim maxLength As Double
    Dim minLength As Double
    Dim aveLength As Double
    Dim totalLength As Double
    Dim totalElements As Double
    Dim curLength As Double
    Dim msg As String
    minLength = 10000
    maxLength = 0
    
    
    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If oElem.color = 0 Then
        
        
        If (oElem.type = msdElementTypeLine Or oElem.type = msdElementTypeLineString) Then
                  If oElem.AsLineElement.Length = 0 Or oElem.AsLineElement.Length < 2 Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
            End If
           
            curLength = oElem.AsLineElement.Length
            totalElements = totalElements + 1
            totalLength = totalLength + oElem.AsLineElement.Length
            
            If oElem.AsLineElement.Length > maxLength Then
                maxLength = oElem.AsLineElement.Length
            ElseIf oElem.AsLineElement.Length < minLength And oElem.AsLineElement.Length > 0 Then
                minLength = oElem.AsLineElement.Length
            End If
        End If
        End If
       
    Loop
    
    MsgBox ("Max = " & maxLength & "Min Length : " & minLength)

End Sub


Sub selectOnlyLines()

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.type = msdElementTypeLine Or oElem.type = msdElementTypeLineString) Then
            oElem.Redraw msdDrawingModeHilite
            ActiveModelReference.SelectElement oElem
        End If
    Loop
End Sub

Sub selectOnlyText(ByVal nCons As Boolean)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.type = msdElementTypeText Or oElem.type = msdElementTypeTextNode) Then
            If (nCons) Then
                If (oElem.AsTextElement.Text = "H") Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                Else
                    
                End If
            Else
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
            End If
        End If
    Loop
End Sub


Sub selectOnlyPoles(ByVal LVOnly As Boolean)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim CellName As String
    ActiveModelReference.UnselectAllElements
    CellName = "009"
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.type = msdElementTypeCellHeader) Then
            If (LVOnly) Then
                If (oElem.AsCellElement.Name = "009") Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                End If
            Else
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
            End If
        End If
    Loop
End Sub
Public Sub GetPhaseAllocation(A_count As Integer, B_count As Integer, C_count As Integer)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim phaseA_count As Integer
    Dim phaseB_count As Integer
    Dim phaseC_count As Integer
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    phaseA_count = 0
    phaseB_count = 0
    phaseC_count = 0
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        'If oElem.Level.name = "Level 19" Then
            If ((oElem.type = msdElementTypeText)) Then 'Or oElem.Type = msdElementTypeTextNode)) Then
                'If (aPhase) Then
                    If (oElem.AsTextElement.Text = " A") Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                        oElem.Redraw msdDrawingModeHilite
                        ActiveModelReference.SelectElement oElem
                        phaseA_count = phaseA_count + 1
                    End If
               ' End If
               ' If (bPhase) Then
                    If (oElem.AsTextElement.Text = " B") Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                        oElem.Redraw msdDrawingModeHilite
                        ActiveModelReference.SelectElement oElem
                        phaseB_count = phaseB_count + 1
                    End If
                'End If
                'If (cPhase) Then
                    If (oElem.AsTextElement.Text = " C") Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                        oElem.Redraw msdDrawingModeHilite
                        ActiveModelReference.SelectElement oElem
                        phaseC_count = phaseC_count + 1
                    End If
                'End If
            End If
       
    Loop
    
    A_count = phaseA_count
    B_count = phaseB_count
    C_count = phaseC_count
    'If Not (ActiveModelReference.AnyElementsSelected) Then
    '    MsgBox "Phase does'nt exist"
    'End If
    
End Sub
Sub selectOnlyPhases(ByVal aPhase As Boolean, bPhase As Boolean, cPhase As Boolean)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim phaseA_count As Integer
    Dim phaseB_count As Integer
    Dim phaseC_count As Integer
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    phaseA_count = 0
    phaseB_count = 0
    phaseC_count = 0
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        'If oElem.Level.name = "Level 19" Then
            If ((oElem.type = msdElementTypeText)) Then 'Or oElem.Type = msdElementTypeTextNode)) Then
                If (aPhase) Then
                    If (oElem.AsTextElement.Text = " R") Then 'Or (oElem.AsTextElement.Text = " R") Then 'Or oElem.AsTextNodeElement.TextLine = "H") Then
                        oElem.Redraw msdDrawingModeHilite
                        ActiveModelReference.SelectElement oElem
                        phaseA_count = phaseA_count + 1
                    End If
                End If
                If (bPhase) Then
                    If (oElem.AsTextElement.Text = " W") Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                        oElem.Redraw msdDrawingModeHilite
                        ActiveModelReference.SelectElement oElem
                        phaseB_count = phaseB_count + 1
                    End If
                End If
                If (cPhase) Then
                    If (oElem.AsTextElement.Text = " B") Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                        oElem.Redraw msdDrawingModeHilite
                        ActiveModelReference.SelectElement oElem
                        phaseC_count = phaseC_count + 1
                    End If
                End If
            End If
       
    Loop
    
    If aPhase And bPhase And cPhase Then
         MsgBox ("Phase A :" & phaseA_count & vbCrLf & _
                "Phase B :" & phaseB_count & vbCrLf & _
                "Phase C :" & phaseC_count & vbCrLf)
    End If
    
    If Not (ActiveModelReference.AnyElementsSelected) Then
        MsgBox "Phase doesnt exist"
    End If
    
End Sub
Sub PopulateElementLists(ByRef oFenceEnum As ElementEnumerator, houseTextList, TrfrShapesList, TRFRList, AirdacList, LVpolesList, MVpolesList, MVSharingPolesListLVConductorList, MVConductorList, LVStaysAndAnchorsList, MVStaysAndAnchorsList, elseElements As Collection)
    ActiveModelReference.UnselectAllElements
    Set oFenceEnum = ActiveDesignFile.Fence.GetContents
    Dim oElem As Element
    Do While oFenceEnum.MoveNext
        Set oElem = oFenceEnum.Current
        Select Case oElem.Level.Name
            Case "Level 3" 'HouseTextLevelName
                houseTextList.Add oElem, oElem.ID
                
            'Case "Level 7" 'TrfrShapesLevelName
            '    TrfrShapesList.Add oElem, oElem.ID
           ' Case "Level 13" 'TRFRLevelName
           '     TRFRList.Add oElem, oElem.ID
           ' Case "Level 18" 'AirdacLevelName
          '      AirdacList.Add oElem, oElem.ID
          '  Case "Level 20" 'LVPolesLevelName
          '      LVpolesList.Add oElem, oElem.ID
           ' Case "Level 21" 'MVPolesLevelName
           '     MVpolesList.Add oElem, oElem.ID
          '  Case "Level 22" 'MVSharingPolesLevelName
         '       MVSharingPolesList.Add oElem, oElem.ID
         '   Case "Level 25" 'LVConductorLevelName
        '        LVConductorList.Add oElem, oElem.ID
         '   Case "Level 29" 'MVConductorLevelName
         '       MVConductorList.Add oElem, oElem.ID
          '  Case "Level 34" 'LVStaysAndAnchorsLevelName
          '      LVStaysAndAnchorsList.Add oElem, oElem.ID
          '  Case "Level 35" 'MVStaysAndAnchorsLevelName
          '      MVStaysAndAnchorsList.Add oElem, oElem.ID
           ' Case Else
           '     elseElements.Add oElem, oElem.ID
        End Select
    Loop

End Sub


Sub PopulatElementLists(ByVal oFenceEnum As ElementEnumerator, houseTextList, TrfrShapesList, TRFRList, AirdacList, LVpolesList, MVpolesList, MVSharingPolesList, LVConductorList, MVConductorList, LVStaysAndAnchorsList, MVStaysAndAnchorsList, elseElements As Collection)
    ActiveModelReference.UnselectAllElements
    'Set oFenceEnum = ActiveDesignFile.Fence.GetContents
    Dim oElem As Element
    Do While oFenceEnum.MoveNext
        Set oElem = oFenceEnum.Current
        Select Case oElem.Level.Name
            Case "Level 3" 'HouseTextLevelName
                houseTextList.Add oElem, oElem.ID
                
            'Case "Level 7" 'TrfrShapesLevelName
            '    TrfrShapesList.Add oElem, oElem.ID
           ' Case "Level 13" 'TRFRLevelName
           '     TRFRList.Add oElem, oElem.ID
           ' Case "Level 18" 'AirdacLevelName
          '      AirdacList.Add oElem, oElem.ID
          '  Case "Level 20" 'LVPolesLevelName
          '      LVpolesList.Add oElem, oElem.ID
           ' Case "Level 21" 'MVPolesLevelName
           '     MVpolesList.Add oElem, oElem.ID
          '  Case "Level 22" 'MVSharingPolesLevelName
         '       MVSharingPolesList.Add oElem, oElem.ID
         '   Case "Level 25" 'LVConductorLevelName
        '        LVConductorList.Add oElem, oElem.ID
         '   Case "Level 29" 'MVConductorLevelName
         '       MVConductorList.Add oElem, oElem.ID
          '  Case "Level 34" 'LVStaysAndAnchorsLevelName
          '      LVStaysAndAnchorsList.Add oElem, oElem.ID
          '  Case "Level 35" 'MVStaysAndAnchorsLevelName
          '      MVStaysAndAnchorsList.Add oElem, oElem.ID
           ' Case Else
           '     elseElements.Add oElem, oElem.ID
        End Select
    Loop

End Sub
Sub selectServiceLines()

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.type = msdElementTypeLine) Then ' Or oElem.Type = msdElementTypeLineString) Then
            If (oElem.Level.Name = "Level 18" And oElem.color = 0) Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
            End If
        End If
    Loop
End Sub
Sub selectOnlyHseText() 'ByVal nCons As Boolean)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.type = msdElementTypeText) Then ' Or oElem.Type = msdElementTypeTextNode) Then
            If (oElem.AsTextElement.Text = "_H") Then 'oElem.AsTextElement.Text = "HOUSE" Or oElem.AsTextElement.Text = "HSE" Or oElem.AsTextElement.Text = "_H") Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
            End If
        End If
    Loop
End Sub

Sub selectOnlyLVLines()

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.type = msdElementTypeLine Or oElem.type = msdElementTypeLineString) Then
            If (oElem.color = 2 And oElem.Level.Name = "Level 25") Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
            End If
        End If
    Loop
End Sub
Sub selectTrfZones(ByVal elementType As String)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    'Dim cellName As String
    'cellName = "001"
    'ActiveModelReference.UnselectAllElements
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    Select Case elementType
        Case "TRF SHAPE"
            elementType = "001"
        Case Else
            elementType = "000"
    End Select
     
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        
        If (oElem.type = msdTypeShape) Then
            'If (cellType = "All") Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
                'SendKeyin "mdl load cellcounter"
                'processStats()
            'Else
             '   If (oElem.AsCellElement.Name = cellName) Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                '    oElem.Redraw msdDrawingModeHilite
               '     ActiveModelReference.SelectElement oElem
              '  End If
            End If
        End If
    Loop
    
    
End Sub
Sub GetTrfCell(oCell As Element)
    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim CellName As String
    CellName = "TXTCS"
    ActiveModelReference.UnselectAllElements
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        
        If (oElem.type = msdElementTypeCellHeader) Then
            If (CellType = "All") Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
                'SendKeyin "mdl load cellcounter"
                'processStats()
            Else
                If (oElem.AsCellElement.Name = CellName) Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                    Set oCell = oElem
                    Exit Do
                End If
            End If
        End If
        
    Loop
    
End Sub

Sub selectTRFRzoneShapes() 'ByVal cellType As String)
    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    'Dim cellName As String
    'cellName = "001"
    ActiveModelReference.UnselectAllElements
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
              
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        
        If (oElem.type = msdElementTypeShape) Then
            If (oElem.Level.Name = "Level 11") Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
            End If
        End If
    Loop
    
End Sub
Sub selectZons(ByRef oEnum As ElementEnumerator, ByRef oSelZone As clsDataZone) ' As clsDataZone 'ByVal cellType As String)
    Dim oElem As Element


    Dim MVFlyingStays As Collection 'MVFlyingStay

    ActiveModelReference.UnselectAllElements
     
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        'Debug.Print CStr(oEnum.Current.ID.high) & CStr(oEnum.Current.ID.low)
        
        Select Case oElem.type
            Case msdElementTypeCellHeader 'Cells
            
                Select Case oElem.AsCellElement.Name
                    Case "001" 'LVStays
                        oSelZone.LVStays.Add oElem, CStr(oElem.ID.low)
                    Case "002" 'LVStruts
                        oSelZone.LVStruts.Add oElem, CStr(oElem.ID.low)
                    Case "003" 'MVStays
                        oSelZone.MVStays.Add oElem, CStr(oElem.ID.low)
                    Case "004" 'MVStruts
                        oSelZone.MVStruts.Add oElem, CStr(oElem.ID.low)
                    Case "005" 'TRFRCircles
                        oSelZone.TRFRCircles.Add oElem, CStr(oElem.ID.low)
                    Case "006" 'TRFRNameplates"
                        oSelZone.TRFRNameplates.Add oElem, CStr(oElem.ID.low)
                    Case "009" 'LVpoles"
                        oSelZone.LVpoles.Add oElem, CStr(oElem.ID.low)
                    Case "010" 'MVpoles"
                        oSelZone.mvpoles.Add oElem, CStr(oElem.ID.low)
                    Case "011"  'MVSharingPoles
                        oSelZone.MVSharingPoles.Add oElem, CStr(oElem.ID.low)
                    Case "012" 'KickerPoles"
                        oSelZone.KickerPoles.Add oElem, CStr(oElem.ID.low)
                    Case "015"  'LVFlyingStays"
                        oSelZone.LVFlyingStays.Add oElem, CStr(oElem.ID.low)
                    Case "016" 'LVShortStays
                        oSelZone.LVShortStays.Add oElem, CStr(oElem.ID.low)
                    Case "017" 'MVFlyingStays"
                        oSelZone.MVFlyingStays.Add oElem, CStr(oElem.ID.low)
                    Case "018" 'MVShortStays
                        oSelZone.MVShortStays.Add oElem, CStr(oElem.ID.low)
                    Case "035" 'MVFuseIsolators
                        oSelZone.MVFuseIsolators.Add oElem, CStr(oElem.ID.low)
                    Case "TXTCS" 'TRFRs
                        oSelZone.TRFRs.Add oElem, CStr(oElem.ID.low)
                    Case Else
                        oSelZone.elseCellElements.Add oElem, CStr(oElem.ID.low)
                End Select
                
            Case msdElementTypeLine
            
                Select Case oElem.Level.Name
                    Case "Level 18" ' Airdac
                         oSelZone.Airdac.Add oElem, CStr(oElem.ID.low)
                         'Debug.Print CStr(oEnum.Current.ID.high) & CStr(oEnum.Current.ID.low)
                    Case "Level 25" 'LVConductor
                        oSelZone.LVConductor.Add oElem, CStr(oElem.ID.low)
                    Case "Level 29" 'MVConductor
                        oSelZone.MVConductor.Add oElem, CStr(oElem.ID.low)
                    Case Else
                        oSelZone.elseLineElements.Add oElem, CStr(oElem.ID.low)
                End Select
                
            Case msdElementTypeShape
            
                Select Case oElem.Level.Name
                    Case "Level 11" 'TrfrShapes
                        oSelZone.TrfrShapes.Add oElem, CStr(oElem.ID.low)
                    Case Else
                        oSelZone.elseShapeElements.Add oElem, CStr(oElem.ID.low)
                End Select
                
            Case (msdElementTypeText)
            
                Select Case oElem.Level.Name
                    Case "Level 3" 'houseText
                        oSelZone.houseText.Add oElem, CStr(oElem.ID.low)
                    Case Else
                        oSelZone.elseTextElements.Add oElem, CStr(oElem.ID.low)
                        
                End Select
                
            Case Else
                oSelZone.elseElements.Add oElem, CStr(oElem.ID.low)
        End Select
   
    Loop
   

     Debug.Print oSelZone.ToString
     Debug.Print oSelZone.mergedMVPolesCollection.count
     
 Debug.Print oSelZone.printCoordinatesMVPoles
 Debug.Print oSelZone.printCoordinatesMVLines
    
End Sub
Sub selectNormalCells(ByVal CellType As String)
    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim CellName As String
    CellName = "001"
    ActiveModelReference.UnselectAllElements
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    Select Case CellType
        Case "LV Stays"
            CellName = "001"
        Case "LV Struts"
            CellName = "002"
        Case "MV Stays"
            CellName = "003"
        Case "MV Struts"
            CellName = "004"
        Case "LV Poles"
            CellName = "009"
        Case "MV Poles"
            CellName = "010"
        Case "MVLV Poles"
            CellName = "011"
        Case "LV Flying Stays"
            CellName = "015"
        Case "MV Flying Stays"
            CellName = "017"
        Case "TRFRS"
            CellName = "TXTCS"
        Case Else
            CellName = "000"
    End Select
     
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        
        If (oElem.type = msdElementTypeCellHeader) Then
            If (CellType = "All") Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
                'SendKeyin "mdl load cellcounter"
                'processStats()
            Else
                If (oElem.AsCellElement.Name = CellName) Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                End If
            End If
        End If
    Loop
    
End Sub

Public Sub QuickSort(vArray As Collection, inLow As Long, inHi As Long)

  Dim pivot   As Point3d
  Dim tmpSwap As Variant
  'Set pivot = New Collection
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot.X = vArray((inLow + inHi) \ 2).X
  pivot.Y = vArray((inLow + inHi) \ 2).Y
  While (tmpLow <= tmpHi)

     While (vArray(tmpLow).X < pivot.X And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot.X < vArray(tmpHi).X And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub
Sub GetServiceCable(Airdac As Integer) '' conType As String)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim color As Integer
    Dim levelName As String
    Dim stats As Boolean
    stats = False
    
    ''Select Case conType
        ''Case "Service Cable"
            color = 0
            levelName = "Level 18"
     ''   Case "LV Lines"
      ''      color = 2
        ''    levelName = "Level 25"
        ''Case "MV Lines"
         ''   color = 3
         ''   levelName = "Level 29"
        ''Case "Stats"
           stats = True
       '' Case Else
        
         '   MsgBox "Invalid conductor"
        'End Select
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.type = msdElementTypeLine Or oElem.type = msdElementTypeLineString) Then
            If Not stats Then
                If (oElem.color = color And oElem.Level.Name = levelName) Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                End If
            Else
                If (oElem.color = color And oElem.Level.Name = levelName) Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                End If
            End If
            
        End If
    Loop
    Airdac = UBound(ActiveModelReference.GetSelectedElements.BuildArrayFromContents) + 1
    
End Sub

Sub selectConductor(ByVal conType As String)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim color As Integer
    Dim levelName As String
    Dim stats As Boolean
    stats = False
    
    Select Case conType
        Case "Service Cable"
            color = 0
            levelName = "Level 18"
        Case "LV Lines"
            color = 2
            levelName = "Level 25"
        Case "MV Lines"
            color = 3
            levelName = "Level 29"
        Case "Stats"
            stats = True
        Case Else
        
            MsgBox "Invalid conductor"
        End Select
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.type = msdElementTypeLine Or oElem.type = msdElementTypeLineString) Then
            If Not stats Then
                If (oElem.color = color And oElem.Level.Name = levelName) Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                End If
            Else
                If (oElem.color = color And oElem.Level.Name = levelName) Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                End If
            End If
            
        End If
    Loop
End Sub

Sub getConductorStats()
    Dim maxLength As Double
    Dim minLength As Double
    Dim aveLength As Double
    Dim totalLength As Double
    Dim totalElements As Double
    Dim curLength As Double
    Dim msg As String
    minLength = 10000
    maxLength = 0
    
    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.color = 0 And oElem.Level.Name = "Level 18") Then
        
        If (oElem.type = msdElementTypeLine) Then  'Or oElem.Type = msdElementTypeLineString) Then
            If oElem.AsLineElement.Length = 0 Then 'Or oElem.AsLineElement.length < 12 Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
            End If
           
            curLength = oElem.AsLineElement.Length
            totalElements = totalElements + 1
            totalLength = totalLength + oElem.AsLineElement.Length
            aveLength = totalLength / totalElements
            If oElem.AsLineElement.Length > maxLength Then
                maxLength = oElem.AsLineElement.Length
            ElseIf oElem.AsLineElement.Length < minLength And oElem.AsLineElement.Length > 0 Then
                minLength = oElem.AsLineElement.Length
            End If
        End If
        End If
       
    Loop
    
    MsgBox ("Max = " & maxLength & vbCrLf & "Min Length : " & minLength & vbCrLf & "Average : " & aveLength & vbCrLf & " Total length : " & totalLength)
     
End Sub

Sub getSelConductorStats()
    Dim maxLength As Double
    Dim minLength As Double
    Dim aveLength As Double
    Dim totalLength As Double
    Dim totalElements As Double
    Dim curLength As Double
    Dim msg As String
    minLength = 10000
    maxLength = 0
    
    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.color = 2 And oElem.Level.Name = "Level 25") Then
        
        
        If (oElem.type = msdElementTypeLine Or oElem.type = msdElementTypeLineString) Then
            If oElem.AsLineElement.Length = 0 Or oElem.AsLineElement.Length < 12 Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
            End If
           
            curLength = oElem.AsLineElement.Length
            totalElements = totalElements + 1
            totalLength = totalLength + oElem.AsLineElement.Length
            
            If oElem.AsLineElement.Length > maxLength Then
                maxLength = oElem.AsLineElement.Length
            ElseIf oElem.AsLineElement.Length < minLength And oElem.AsLineElement.Length > 0 Then
                minLength = oElem.AsLineElement.Length
            End If
        End If
        End If
       
    Loop
    
    MsgBox ("Max = " & maxLength & " Min Length : " & minLength)
     
End Sub


Sub selectNormalCellsDupes() 'ByVal cellType As String)
    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim CellName As String
    CellName = "001"
    ActiveModelReference.UnselectAllElements
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    Select Case CellType
        Case "LVStays"
            CellName = "001"
        Case "LVStruts"
            CellName = "002"
        Case "MVStays"
            CellName = "003"
        Case "MVStruts"
            CellName = "004"
        Case "LVPoles"
            CellName = "009"
        Case "MVPoles"
            CellName = "010"
        Case "MVLVPoles"
            CellName = "011"
        Case "LVFlyingStays"
            CellName = "015"
        Case "MVFlyingStays"
            CellName = "017"
        Case "TRFRS"
            CellName = "TXTCS"
        Case Else
            CellName = "000"
    End Select
     
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        
        If (oElem.type = msdElementTypeCellHeader) Then
            If (CellType = "All") Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
                'SendKeyin "mdl load cellcounter"
                'processStats()
            Else
                If (oElem.AsCellElement.Name = CellName) Then  'Or oElem.AsTextNodeElement.TextLine = "H") Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                End If
            End If
        End If
    Loop
    
End Sub

Sub selectOnLevel(ByVal lvlName As String, ByVal addToSelSet As Boolean)
    'Function used to select only elements in the defined level. This works only on fence contents though
    If Not (addToSelSet) Then ActiveModelReference.UnselectAllElements
    ' select everything
    Dim oElem As Element
    Set oElem = elemIn
End Sub


Public Function matchLevel(oElem As Element, lvlName As String, LV As Boolean)
    
    Dim oSFlag, oLvlFlag, oColFlag, oTypeFlag As Boolean
    Dim style1, style2, style3 As String
    
    oLvlFlag = False
    oColFlag = False
    oTypeFlag = False
    
    
    If oElem.Level.Name = lvlName Then oLvlFlag = True
    
    If oElem.LineStyle = msdElementTypeShape Then oTypeFlag = True
    
    matchParams = oTypeFlag 'And oLvlFlag And oColFlag
    
End Function


Sub stats(ByVal LVStats As String)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim LVpoles, mvpoles, mvlvpoles, LVStays, LVStruts, LVFlyingStays, MVStays, txtLvl55, txtLvl56, txtLvl57, txtLvl58 As Collection
   
  
    Set TRFZone = New clsTrfZone
    Set TRFRZone = New clsTrfrZone
    Set txtLvl55 = New Collection
    Set txtLvl56 = New Collection
    Set txtLvl57 = New Collection
    Set txtLvl58 = New Collection
    
    Set LVpoles = New Collection
    Set mvpoles = New Collection
    Set mvlvpoles = New Collection
    Set LVStays = New Collection
    Set MVStays = New Collection
    
    
    Dim LV_SP_Structures() As String
    Dim LV_DP_Structures() As String
    Dim MV_Structures() As String
    
    'initialize the LV text prefixes
    LV_SP_Structures = InitializePrefixStrings(low:=1, high:=10, prefix:="115")
    LV_DP_Structures = InitializePrefixStrings(low:=1, high:=10, prefix:="114")
    MV_Structures = InitializePrefixStrings(low:=1, high:=10, prefix:="114")
    
    Dim levelName, output As String
    Dim length1, length2 As Double
    Dim ABC_1PH, ABC_2PH, ARVIDAL_35, OAK, ServiceCables As Collection
    Set ABC_1PH = New Collection
    Set ABC_2PH = New Collection
    Set ARVIDAL_35 = New Collection
    Set LVStruts = New Collection
    Set LVFlyingStays = New Collection
    Set OAK = New Collection
    Set txtLvl57 = New Collection
    Set ServiceCables = New Collection
    length1 = 0
    length2 = 0
    'Dim output As String
    
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        
            Select Case oElem.type
                Case msdElementTypeLine
                    TRFRZone.Conductors.Add oElem
                    If (oElem.AsLineElement.LineStyle.Name = "ABC35-R") Or _
                        (oElem.AsLineElement.LineStyle.Name = "ABC35-W") Or _
                        (oElem.AsLineElement.LineStyle.Name = "ABC35-B") Then
                        
                        ABC_1PH.Add oElem
                        TRFZone.Conductors.LV_1PHconductor.Add oElem
                        oElem.Redraw msdDrawingModeHilite
                        ActiveModelReference.SelectElement oElem
                       
                    ElseIf (oElem.AsLineElement.LineStyle.Name = "ABC35-RW") Or _
                        (oElem.AsLineElement.LineStyle.Name = "ABC35-WB") Or _
                        (oElem.AsLineElement.LineStyle.Name = "ABC35-RB") Then
                        
                        ABC_2PH.Add oElem
                         oElem.Redraw msdDrawingModeHilite
                         ActiveModelReference.SelectElement oElem
                    ElseIf (oElem.AsLineElement.LineStyle.Name = "35") Then
                        ARVIDAL_35.Add oElem
                        oElem.Redraw msdDrawingModeHilite
                        ActiveModelReference.SelectElement oElem
                    ElseIf (oElem.AsLineElement.LineStyle.Name = "4mm AD A") Or _
                        (oElem.AsLineElement.LineStyle.Name = "4mm AD B") Or _
                        (oElem.AsLineElement.LineStyle.Name = "4mm AD") Or _
                        (oElem.AsLineElement.LineStyle.Name = "4mm AD C") Then
                        
                        ServiceCables.Add oElem
                        oElem.Redraw msdDrawingModeHilite
                        ActiveModelReference.SelectElement oElem
                       
                    End If
                    
                Case msdElementTypeText
                    Select Case (oElem.AsTextElement.Level.Name)
                        Case ("Level 55") 'CBAzz
                            txtLvl55.Add oElem
                        Case ("Level 56") 'MV
                            txtLvl56.Add oElem
                        Case ("Level 57") 'LV
                            txtLvl57.Add oElem
                        Case ("Level 58") 'TRF
                            txtLvl58.Add oElem
                    End Select
                Case msdElementTypeCellHeader
                    Select Case oElem.AsCellElement.Name
                        Case "TXTCS"
                            'transformer cell
                        Case "009"
                            'LV poles
                            'LVpoleCol.LVCollection.Add oElem
                            LVpoles.Add oElem
                        Case "010"
                            'MV poles
                            mvpoles.Add oElem
                        Case "011"
                            'MVLV poles
                            mvlvpoles.Add oElem
                        Case "001"
                            'LV stays/struts
                            LVStays.Add oElem
                        Case "002"
                            'LV struts
                            LVStruts.Add oElem
                        Case "015"
                            'LV Flyingstays/struts
                            LVFlyingStays.Add oElem
                        Case "003", "004", "017"
                            'MV stays and struts
                            MVStays.Add oElem
                    End Select
                Case Else
            End Select
    Loop
    
    Set gABC_1PH = ABC_1PH
    Set gABC_2PH = ABC_2PH
    Set gLVPoles = LVpoles
    Set gMVLVPoles = mvlvpoles
    Set gLVStays = LVStays
    Set gLVStruts = LVStruts
    Set gLVFlyingStays = LVFlyingStays
    Set gServiceCables = ServiceCables
    Set gTxtLvl57 = txtLvl57
    
    
    'output = computeOutputLines(col_Line:=ABC_2PH)
    'output = output + computeOutputLines(col_Line:=ABC_2PH)
            
    'DataScanner.outputBox.Text = output
    If LVStats = "LV" Then
        computeOutputLV
    Else
        'computeOutputMV
    End If
End Sub

Function computeOutputLines(ByVal col_Line As Collection) As String
    Dim length1 As Double
    Dim cond As Element
    
    For Each cond In col_Line
        length1 = length1 + cond.AsLineElement.Length
    Next
    
    computeOutputLines = length1 & vbCrLf
    
End Function
Function computeServiceLines(ByVal col_Line As Collection) As String
    Dim length1 As Double
    Dim cond As Element
    
    For Each cond In col_Line
        length1 = length1 + cond.AsLineElement.Length
    Next
    
    computeServiceLines = length1 & vbCrLf
    
End Function

Function computeCells(ByVal cells As Collection) As String
    Dim oelm As Element
    For Each oelm In cells
        oelm.Redraw msdDrawingModeHilite
        ActiveModelReference.SelectElement oelm
    Next
    computeCells = cells.count & vbCrLf

End Function
Function ProcessText(ByVal textCol As Collection) As String
    Dim oText As String
    oText = ""
    Dim oelm As Element
    Dim ABC_1145 As Collection
    Dim ABC_1146 As Collection
    Dim ABC_1147 As Collection
    Dim ABC_1148 As Collection
    Dim ABC_1149 As Collection
    Dim ABC_1151 As Collection
    Dim ABC_1152 As Collection
    Dim ABC_1153 As Collection
    Dim ABC_1154 As Collection
    Dim ABC_1155 As Collection
    Dim ABC_1156 As Collection
    Dim ABC_1157 As Collection
    Dim ABC_1158 As Collection
    Dim ABC_1159 As Collection
    
    Set ABC_1145 = New Collection
    Set ABC_1146 = New Collection
    Set ABC_1147 = New Collection
    Set ABC_1148 = New Collection
    Set ABC_1149 = New Collection
    Set ABC_1151 = New Collection
    Set ABC_1152 = New Collection
    Set ABC_1153 = New Collection
    Set ABC_1154 = New Collection
    Set ABC_1155 = New Collection
    Set ABC_1156 = New Collection
    Set ABC_1157 = New Collection
    Set ABC_1158 = New Collection
    Set ABC_1159 = New Collection
    
    For Each oelm In textCol
        Select Case oelm.AsTextElement.Text
            Case "1145"
                ABC_1145.Add oelm
            Case "1146"
                ABC_1146.Add oelm
            Case "1147"
                ABC_1147.Add oelm
            Case "1148"
                ABC_1148.Add oelm
            Case "1149"
                ABC_1149.Add oelm
            Case "1151"
                ABC_1151.Add oelm
            Case "1152"
                ABC_1152.Add oelm
            Case "1153"
                ABC_1153.Add oelm
            Case "1154"
                ABC_1154.Add oelm
            Case "1155"
                ABC_1155.Add oelm
            Case "1156"
                ABC_1156.Add oelm
            Case "1157"
                ABC_1157.Add oelm
            Case "1158"
                ABC_1158.Add oelm
            Case "1159"
                ABC_1159.Add oelm
            Case "2x1154"
                ABC_1154.Add oelm
                ABC_1154.Add oelm
            Case "2x1146"
                ABC_1146.Add oelm
                ABC_1146.Add oelm
            Case "3x1154"
                ABC_1154.Add oelm
                ABC_1154.Add oelm
                ABC_1154.Add oelm
            Case "3x1146"
                ABC_1146.Add oelm
                ABC_1146.Add oelm
                ABC_1146.Add oelm
            Case "4x1154"
                ABC_1154.Add oelm
                ABC_1154.Add oelm
                ABC_1154.Add oelm
                ABC_1154.Add oelm
            Case "4x1146"
                ABC_1146.Add oelm
                ABC_1146.Add oelm
                ABC_1146.Add oelm
                ABC_1146.Add oelm
            Case Else
                
        End Select
        '1145
        oText = "LV Structures '1145' : " & ABC_1145.count & vbCrLf + _
                "LV Structures '1146' : " & ABC_1146.count & vbCrLf + _
                "LV Structures '1147' : " & ABC_1147.count & vbCrLf + _
                "LV Structures '1148' : " & ABC_1148.count & vbCrLf + _
                "LV Structures '1149' : " & ABC_1149.count & vbCrLf + _
                "LV Structures '1151' : " & ABC_1151.count & vbCrLf + _
                "LV Structures '1152' : " & ABC_1152.count & vbCrLf + _
                "LV Structures '1153' : " & ABC_1153.count & vbCrLf + _
                "LV Structures '1154' : " & ABC_1154.count & vbCrLf + _
                "LV Structures '1155' : " & ABC_1155.count & vbCrLf + _
                "LV Structures '1156' : " & ABC_1156.count & vbCrLf + _
                "LV Structures '1157' : " & ABC_1157.count & vbCrLf + _
                "LV Structures '1158' : " & ABC_1158.count & vbCrLf + _
                "LV Structures '1159' : " & ABC_1159.count & vbCrLf
        
        oelm.Redraw msdDrawingModeHilite
        ActiveModelReference.SelectElement oelm
    Next
    
    ProcessText = oText 'textCol.count & vbCrLf
    
End Function
Function getPoleTopBoxes(ByVal LVCol As Collection, ByVal MVLVCol As Collection) As String
    Dim str As String 'string(1 to 100) as s
    Dim ppole As Element
    Dim poleTopBoxes As Collection
    Set poleTopBoxes = New Collection
    For Each ppole In LVCol
        'Dim newBox As clsPoleTopBox
        'Set newBox = New clsPoleTopBox
        'Let newBox.pole = ppole
        'poleTopBoxes.Add (newBox)
        'Set newBox = Nothing
    Next
    
    
    
    
    
    
    getPoleTopBoxes = str
End Function
Sub computeOutputLV()
   Dim str As String
   'process the LV lines first
   str = "Total ABC35 1PH : " + computeOutputLines(col_Line:=gABC_1PH)
   str = str + "Total ABC35 2PH : " + computeOutputLines(col_Line:=gABC_2PH)
   'process text
   str = str + ProcessText(textCol:=gTxtLvl57)
   'process poletopboxes
   str = str + getPoleTopBoxes(LVCol:=gLVPoles, MVLVCol:=gMVLVPoles)
   'process poles
   str = str + "Total LV poles : " + computeCells(cells:=gLVPoles)
   str = str + "Total LV Stays : " + computeCells(cells:=gLVStays)
   str = str + "Total LV Struts : " + computeCells(cells:=gLVStruts)
   str = str + "Total LV Flying Stays : " + computeCells(cells:=gLVFlyingStays)
   'str = str + "Total LV poles : " + computeCells(cells:=gLVPoles)
   'str = str + "Total LV poles : " + computeCells(cells:=gLVPoles)
   'str = str + "Total LV poles : " + computeCells(cells:=gLVPoles)
   'str = str + "Total LV poles : " + computeCells(cells:=gLVPoles)
   '
   DataScanner.outputBox.Text = str
                                 
End Sub

'Type poleTopBox
 '       Public pole As Element 'Integer' public pole as element
'        Public freq As Integer
'End Type

    
Public Function getDenominations(list As Collection) As String
        Dim f As clsPoleTopBox
        'Dim g As poleTopBox
        'Set f = New clsPoleTopBox

        Dim outs(50) As Integer '= {} ') = New List(Of Integer)()
        Dim ptops(50) As Integer ' ={0}
        Dim intList As Collection
        Set intList = New Collection
        'Dim og As Collection = New Collection
        Dim index As Integer
        For index = 1 To list.count
            intList.Add (list(index))
        Next
        
       ' og = intList
        Dim listLen As Integer
        listLen = list.count
        Dim i As Integer
        Dim mO(5) As Integer
        Dim ii, ix, count As Integer
        i = 1
        count = 0
  
        While (list.count > 0)
            Dim j As Integer
            j = 1
            While (j <= intList.count)
                If list.Item(i) = intList.Item(j) Then
                    intList.Remove (j)
                    j = j - 1
                    count = count + 1
                End If
                j = j + 1
            Wend
            'Set list = intList
            'mO(count) = mO(count) + 1
           'count = 0
            i = i + 1
        Wend
        
        For ii = 0 To UBound(outs) - 1
            For ix = 1 To UBound(outs)
            
                If outs(ii) = ix Then
                    ptops(ix) = ptops(ix) + 1
                End If
            Next
        Next
        
        Dim str As String '
        Dim k As Integer
       
        For k = 1 To UBound(mO)
            If Not (mO(k) = 0) Then
                str = str + "Number of " & k & "'s pole top boxes is : " & mO(k) & vbCrLf
            End If
        Next k
        Debug.Print (str)
        getDenominations = str
End Function
Sub textfontchange()
    Dim tEle As TextElement
    Dim oFont As Font
    Set tEle = CreateTextElement1(Nothing, "hello", Point3dZero, Matrix3dIdentity)
    Set oFont = ActiveDesignFile.Fonts("Verdana")
    Set tEle.TextStyle.Font = oFont
    ActiveModelReference.AddElement tEle
End Sub
Public Sub ChangeCell104(ECell As CellElement) ', TrfrName As String, Rating As String)
    Dim oEE1     As ElementEnumerator
    Dim AElement1 As Element
    Dim Txt         As String
    Dim myString    As String
    Dim oFont As Font
     Set oEE1 = ECell.GetSubElements
     While oEE1.MoveNext
          Set AElement1 = oEE1.Current
          If (AElement1.IsCellElement) Then
               Call ChangeCell104(AElement1) ', TrfrName, Rating)
          ElseIf AElement1.IsTextElement Then
               Txt = AElement1.AsTextElement.Text
               'If InStr(UCase(Txt), UCase("Arvidal")) > 0 Then
                    Set oFont = ActiveDesignFile.Fonts("ISO")
                    AElement1.AsTextElement.TextStyle.Font = oFont 'ActiveDesignFile.Fonts("ISO")
                    'myString = Left(Rating, Len(Rating) - 3)
                   ' Rating = myString
                   
                   AElement1.Rewrite
                  ' Exit Sub
               'Else
                    'TrfrName = AElement1.AsTextElement.Text
                    
               'End If
          End If
     Wend
End Sub

Public Sub GetTransformerCellText(ECell As CellElement, TrfrName As String, Rating As String)
Dim oEE1     As ElementEnumerator
Dim AElement1 As Element
Dim Txt         As String
Dim myString    As String
     Set oEE1 = ECell.GetSubElements
     While oEE1.MoveNext
          Set AElement1 = oEE1.Current
          If (AElement1.IsCellElement) Then
               Call GetTransformerCellText(AElement1, TrfrName, Rating)
          ElseIf AElement1.IsTextElement Then
               Txt = AElement1.AsTextElement.Text
               If InStr(UCase(Txt), "KVA") > 0 Then
                    Rating = AElement1.AsTextElement.Text
                    myString = Left(Rating, Len(Rating) - 3)
                    Rating = myString
               Else
                    TrfrName = AElement1.AsTextElement.Text
               End If
          End If
     Wend
End Sub

Sub Macro2()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "CHOOSE ELEMENT"

'   Coordinates are in master units
    startPoint.X = -6341.161488
    startPoint.Y = -3533839.979168
    startPoint.Z = 0#

'   Send a data point to the current command
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 46.15319
    point.Y = startPoint.Y - 4.10721
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "PLACE FENCE ICON"

'   Set a variable associated with a dialog box
    SetCExpressionValue "tcb->msToolSettings.fence.placeMode", 2, ""

    CadInputQueue.SendCommand "PLACE FENCE ICON"

    SetCExpressionValue "tcb->msToolSettings.fence.placeMode", 3, ""

    CadInputQueue.SendCommand "PLACE FENCE ICON"

    CadInputQueue.SendCommand "LOCK FENCE INSIDE"

    point.X = startPoint.X - 2.051253
    point.Y = startPoint.Y + 39.018498
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 2.051253
    point.Y = startPoint.Y + 39.018498
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand DMSG ACTIVATETOOLBYPATH \Drawing Composition\Drawing\Linear\Place SmartLine"

    CadInputQueue.SendCommand "PLACE SMARTLINE"

    CadInputQueue.SendCommand "INPUTMANAGER RUNTOOL -609,2"

    CadInputQueue.SendCommand DMSG ACTIVATETOOLBYPATH \Drawing Composition\Drawing\Linear\Place SmartLine"

    CadInputQueue.SendCommand "PLACE SMARTLINE"

    CommandState.StartDefaultCommand
End Sub

Sub Macro6()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.X = -6075.195034
    startPoint.Y = -3534113.166033
    startPoint.Z = 0#

'   Send a data point to the current command
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Start a command
    CadInputQueue.SendCommand "PLACE FENCE ICON"

'   Set a variable associated with a dialog box
    SetCExpressionValue "tcb->msToolSettings.fence.placeMode", 4, ""

    CadInputQueue.SendCommand "PLACE FENCE ICON"

    SetCExpressionValue "tcb->msToolSettings.fence.placeMode", 3, ""

    CadInputQueue.SendCommand "PLACE FENCE ICON"

    CadInputQueue.SendCommand "LOCK FENCE CLIP"

    CadInputQueue.SendCommand "LOCK FENCE INSIDE"

    point.X = startPoint.X - 5.215948
    point.Y = startPoint.Y + 20.887717
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 5.215948
    point.Y = startPoint.Y - 10.443859
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X - 5.215948
    point.Y = startPoint.Y - 10.443859
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1


    CadInputQueue.SendCommand "Fence name save greg"
    
    CadInputQueue.SendCommand "SAVE DESIGN"

    point.X = startPoint.X + 3.095447
    point.Y = startPoint.Y + 289.326635
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    point.X = startPoint.X - 14.14638
    point.Y = startPoint.Y + 289.883409
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub
Sub Macro7()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.X = -6085.51114
    startPoint.Y = -3533641.571407
    startPoint.Z = 0#

'   Send a data point to the current command
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Send points to simulate a down-drag-up action
    point.X = startPoint.X + 153.72199
    point.Y = startPoint.Y - 323.125103
    point.Z = startPoint.Z
    point2.X = point.X - 0#
    point2.Y = point.Y - 0#
    point2.Z = point.Z
    CadInputQueue.SendDragPoints point, point2, 1

    point.X = startPoint.X + 188.826208
    point.Y = startPoint.Y - 248.140505
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 144.056468
    point.Y = startPoint.Y - 277.699641
    point.Z = startPoint.Z
    point2.X = point.X - 0.952548
    point2.Y = point.Y
    point2.Z = point.Z
    CadInputQueue.SendDragPoints point, point2, 1

    point.X = startPoint.X + 98.334181
    point.Y = startPoint.Y - 144.206771
    point.Z = startPoint.Z
    point2.X = point.X - 0#
    point2.Y = point.Y + 0.95352
    point2.Z = point.Z
    CadInputQueue.SendDragPoints point, point2, 1

    point.X = startPoint.X + 87.365303
    point.Y = startPoint.Y - 284.858987
    point.Z = startPoint.Z
    point2.X = point.X
    point2.Y = point.Y
    point2.Z = point.Z
    CadInputQueue.SendDragPoints point, point2, 1

'   Start a command
    CadInputQueue.SendCommand "PLACE FENCE ICON"

    point.X = startPoint.X + 190.716724
    point.Y = startPoint.Y - 241.473804
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 170.713223
    point.Y = startPoint.Y - 238.136483
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 170.713223
    point.Y = startPoint.Y - 238.136483
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Send a keyin that can be a command string
  

    CadInputQueue.SendKeyin "fence named save jhj"

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    point.X = startPoint.X + 69.02876
    point.Y = startPoint.Y - 171.390048
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub
Sub Macro1()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.X = -1755.97593576782
    startPoint.Y = -3535308.51078313
    startPoint.Z = 0#

'   Send points to simulate a down-drag-up action
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    point2.X = point.X
    point2.Y = point.Y - 1.88308132160455
    point2.Z = point.Z
    CadInputQueue.SendDragPoints point, point2, 1

'   Start a command
    CadInputQueue.SendCommand "PLACE FENCE ICON"

'   Send a data point to the current command
    point.X = startPoint.X - 75.1264822944843
    point.Y = startPoint.Y + 13.1815692540258
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    Dim modalHandler As New Macro1ModalHandler
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Information"

    CadInputQueue.SendCommand DMSG ACTIVATETOOLBYPATH PO Elec\PO Codes and Bill\Place Codes"

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    point.X = startPoint.X - 73.2483202371222
    point.Y = startPoint.Y + 7.53232528828084
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand DMSG ACTIVATETOOLBYPATH PO Elec\PO Codes and Bill\Place Codes"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
Sub GetSelectConductor(ByVal conType As String)

    Dim oElem As Element
    Dim ofence As Fence
    Dim oEnum As ElementEnumerator
    Dim color As Integer
    Dim levelName As String
    Dim stats As Boolean
    stats = False
    
    Select Case conType
        Case "Service Cable"
            color = 0
            levelName = "Level 18"
        Case "LV Lines"
            color = 2
            levelName = "Level 25"
        Case "MV Lines"
            color = 3
            levelName = "Level 29"
        Case "Stats"
            stats = True
        Case Else
        
            MsgBox "Invalid conductor"
        End Select
    ActiveModelReference.UnselectAllElements
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    Do While oEnum.MoveNext
        Set oElem = oEnum.Current
        If (oElem.type = msdElementTypeLine Or oElem.type = msdElementTypeLineString) Then
            If Not stats Then
                If (oElem.color = color And oElem.Level.Name = levelName) Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                End If
            Else
                If (oElem.color = color And oElem.Level.Name = levelName) Then
                    oElem.Redraw msdDrawingModeHilite
                    ActiveModelReference.SelectElement oElem
                End If
            End If
            
        End If
    Loop
End Sub

Sub SelectAllZones()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.X = 1648.35785518146
    startPoint.Y = -3530618.73250201
    startPoint.Z = 0#

'   Send points to simulate a down-drag-up action
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    point2.X = point.X + 6.8212103E-13
    point2.Y = point.Y - 30.9071755637415
    point2.Z = point.Z
    CadInputQueue.SendDragPoints point, point2, 1

'   Start a command
    CadInputQueue.SendCommand DIALOG CMDBROWSE"

'   Send a keyin that can be a command string

    CadInputQueue.SendCommand "MDL SILENTLOAD SELECTBY dialog"

    CadInputQueue.SendCommand DIALOG SELECTBY"

'   Set a variable associated with a dialog box

    SetCExpressionValue "selectorGlobals.colorButton", -1, "SELECTBY"

    SetCExpressionValue "selectorGlobals.color", 7, "SELECTBY"

    SetCExpressionValue "selectorGlobals.color", 7, "SELECTBY"

    CadInputQueue.SendCommand "SELECTBY LEVEL Level 11"

    Dim modalHandler As New SelectAllZonesModalHandler
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Alert"

    CadInputQueue.SendCommand "SELECTBY EXECUTE"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
    
    
End Sub
Sub GetAllZones()

    Dim oElem As Element
    Dim oEnum As ElementEnumerator
    
    Set ofence = ActiveDesignFile.Fence
    Set oEnum = ofence.GetContents
    
    While oEnum.MoveNext
        Set oElem = oEnum.Current
        If oElem.type = msdElementTypeShape Then
            If oElem.AsShapeElement.Level.Name = "Level 11" And oElem.AsShapeElement.color = 7 Then
                oElem.Redraw msdDrawingModeHilite
                ActiveModelReference.SelectElement oElem
                
            End If
        End If
    Wend
    
         
        
End Sub

Function IniytializePrefixStrings(low As Long, high As Long, prefix As String) As String()
    ''''''''''''''''''''''''''''''''''''''''''''''''', poletypes() As String
    ' LoadNumbers
    ' Returns an array of Longs containing the numbers
    ' between Low and High.
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Dim resultArray() As String
    'ReDim poletypes(1 To (High - Low + 1))
    'poletypes(1) = 8
    Dim Ndx As Long
    'Dim Val As Long
    If low > high Then
        MsgBox "Size of array cannot be negative"
        Exit Function
    End If
    ReDim resultArray(1 To (high - low + 1))
    'Val = Low
    For Ndx = LBound(resultArray) To UBound(resultArray)
        resultArray(Ndx) = prefix & CStr(Ndx - 1)
        'ResultArray(Ndx) = CStr(poletypes(1))
        'Val = Val + 1
    Next Ndx
    InitializePrefixStrings = resultArray()
End Function
Public Function ProcessFence(ee As ElementEnumerator) As Boolean
    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim elArray() As Element
    On Error GoTo ERROR_HANDLER
    ProcessText = False
    If ee Is Nothing Then
        Exit Function
    End If
    elArray = ee.BuildArrayFromContents
    iStart = LBound(elArray)
    iEnd = UBound(elArray)
    For i = iStart To iEnd
        Select Case elArray(i).type
            Case MsdElementType.msdElementTypeText
                With elArray(i).AsTextElement
                    '
                    ' Modify text element here
                    '
                End With
                elArray(i).Rewrite
            Case MsdElementType.msdElementTypeTextNode
                ' GetSubElement returns an ElementEnumerator containing text elements within text node
                ' use recursive function call to process new enumerator
                If Not ProcessText(elArray(i).AsTextNodeElement.GetSubElements) Then
                    Exit Function
                End If
            Case MsdElementType.msdElementTypeCellHeader
                ' GetSubElement returns an ElementEnumerator containing elements within cell
                ' use recursive function call to process new enumerator to process any text elements in cell
                If Not ProcessText(elArray(i).AsCellElement.GetSubElements) Then
                    Exit Function
                End If
        End Select
    Next
    ProcessFence = True
    Exit Function
ERROR_HANDLER:
    MsgBox "ProcessFence Error: " & CStr(Err.Number) & Err.Description
End Function

'Function isInsideCircle(oCircle As Element, ByVal test_pt As point3d) As Boolean
  '  Dim retval As Boolean
  '  If Abs(center - test_pt) < radius Then
  '      retval = True
  '  End If
   ' isInsideCircle = retval
    
'End Function
Sub GetAllStats(outputTxt As String)
    
    Dim trfZoneEnum As ElementEnumerator 'Store trf zones
    Dim LVpoles As Collection
    Dim mvlvpoles As Collection
    Dim LV_1Ph As Collection
    Dim LV_2Ph As Collection
    Dim MV_2Ph As Collection
    Dim MV_3Ph As Collection
    
    
    
    Dim trfZoneElem As Element
    Dim trfEnum As ElementEnumerator
    Dim trfElem As Element
    Dim headerText As String
    headerText = "TRF Name " + "," + "TRF Size" + "," + "Phase A" + "," + "Phase B" + "," + "Phase C"
    Dim tRat As String
    Dim Tname As String
    Dim textOut As String
    Dim currZoneInfo As String
    currZoneInfo = ""
    textOut = ""
    Dim connections As Integer
    Dim phaseACount As Integer
    Dim phaseBCount As Integer
    Dim phaseCCount As Integer
    'Dim trfShapesArray() As Shape
    connetions = 0
    phaseACount = 0
    phaseBCount = 0
    phaseCCount = 0
    If Not validFenceDefined Then
          MsgBox "No fence defined"
    End If
    'selectNormalCells "TRFRS"
    ''SelectAllZones
    GetAllZones 'selects all trf zones
    Set trfZoneEnum = ActiveModelReference.GetSelectedElements
    Dim elArray() As Element
    Dim elArr() As Element
    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    
    elArray = trfZoneEnum.BuildArrayFromContents
    iStart = LBound(elArray)
    iEnd = UBound(elArray)
   
    For i = iStart To iEnd
        ActiveModelReference.UnselectAllElements
        ActiveModelReference.SelectElement elArray(i)
        createFenceFromShape
        GetTrfCell trfElem
        'selectNormalCells "TRFRS"
        'Set trfEnum = ActiveModelReference.GetSelectedElements
        'trfEnum.Reset
        Dim ix As Long
        Dim iStartx As Long
        Dim iEndx As Long
        'If trfEnum.MoveNext Then
         '   trfElem = trfEnum.Current
        'End If
        'elArr = trfEnum.BuildArrayFromContents
        'iStartx = LBound(elArr)
        'iEndx = UBound(elArr)
        
        'trfElem = elArr(iStartx)
        GetTransformerCellText ECell:=trfElem.AsCellElement, TrfrName:=Tname, Rating:=tRat  '
        '3. Get number of connections
        ''selectConductor "Service Cable"
        GetServiceCable Airdac:=connections
        '4. get phase Allocation
        GetPhaseAllocation A_count:=phaseACount, B_count:=phaseBCount, C_count:=phaseCCount
        '5. Collate all information
        currZoneInfo = Tname + "," + tRat + "," + CStr(phaseACount) + "," + CStr(phaseBCount) + "," + CStr(phaseCCount) + "," + vbCrLf
        textOut = textOut + currZoneInfo
    Next
    
    outputTxt = textOut
End Sub

Function savetocsv()
    Dim fso As Object
    Set fso = CreateObject("scripting.filesystemobject")
    Dim oFile As Object
    oFile = fso.CreateTextFile(strPath)
    oFile.WriteLine "oui"
    Set fso = Nothing
    Set oFile = Nothing
End Function
End Function
Sub Macro3()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "PLACE FENCE ICON"

'   Coordinates are in master units
    startPoint.X = 3584.56714789442
    startPoint.Y = -3532302.15016276
    startPoint.Z = 0#

'   Send a data point to the current command
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 3.66763058241304
    point.Y = startPoint.Y - 1.83496068324894
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Send a keyin that can be a command string
    CadInputQueue.SendKeyin "fence selection new"

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    CommandState.StartDefaultCommand
End Sub
Sub Macro4()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    CommandState.StartDefaultCommand
End Sub
Sub Macro5()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand DIALOG REFERENCE TOGGLE"

    CadInputQueue.SendCommand "SAVE DESIGN"

    CommandState.StartDefaultCommand
End Sub
Sub Macro8()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand DIALOG REFERENCE TOGGLE"

'   Coordinates are in master units
    startPoint.X = 74727.9239282368
    startPoint.Y = -3662976.00726863
    startPoint.Z = 0#

'   Send a data point to the current command
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "SAVE DESIGN"

    CommandState.StartDefaultCommand
End Sub
Sub Macro9()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.X = 61344.2293941958
    startPoint.Y = -3664420.72486446
    startPoint.Z = 0#

'   Send a data point to the current command
    point.X = startPoint.X
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 10.128543015453
    point.Y = startPoint.Y + 8.25244911015034
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 10.128543015453
    point.Y = startPoint.Y + 8.25244911015034
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Send a keyin that can be a command string
    CadInputQueue.SendKeyin "edit text"

'   Send a message string to an application
'   Content is defined by the application
    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 31"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 moveCaretToPosition 31"

    CadInputQueue.SendKeyin "ACTIVE FONT 512"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 31"

    point.X = startPoint.X + 33.3471698294161
    point.Y = startPoint.Y + 50.0177269740961
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.X = startPoint.X + 32.1298024823191
    point.Y = startPoint.Y + 50.1917489655316
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub
Sub Macro10()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

    Dim modalHandler As New Macro10ModalHandler
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Copy Model"

'   Start a command
    CadInputQueue.SendCommand "MODEL COPY"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
