'
' Created by SharpDevelop.
' User: Gregory Mavhunga
' Date: 2018/12/13
' Time: 12:11 AM
' 
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'
Public Module Module1
	
Public Function GetTree(ByVal conductorList As Collection, polesList As Collection) As String

    Dim out As String
    
End Function
Public Function getRMDPoints() As Collection 'Matrix(fenceFlashOrigin As Point3d) As Matrix3d
    
    Dim polesarr As Collection
    Dim p, pp As Point3d
    p = getFlashOrigin
     polesarr = New Collection
    Dim n As Long
    n = 1
    Dim s As String
    Dim cell As CellElement
    Dim xCoords As Collection
     xCoords = New Collection
    For Each cell In mLVpoles
        
        pp.X = cell.Origin.X + (0 - p.X)
        pp.Y = cell.Origin.Y + (0 - p.Y)
        pp.Z = 0
        
        pp.X = pp.X / 4
        pp.Y = pp.Y / 4
        pp.Y = -pp.Y
        pp.Z = pp.Z
        
        polesarr.Add ( pp, CStr(cell.ID.low))
        xCoords.Add ( pp.X, CStr(cell.ID.low))
        
        
        
        'cell.Origin = p
        s = s & "N" & CStr(n) & " "
        n = n + 1
        s = s & "N" & CStr(n) & vbCrLf
    Next
    'Debug.Print s
    
    Dim rmd As ReticMaster.App
     rmd = New ReticMaster.App
    
     getRMDPoints = xCoords
End Function

Public Function getFlashOrigin() As Point3d
    Dim minX As Double
    Dim maxY As Double
    Dim i As Long
    Dim zcoord As Double
    zcoord = mZoneShape.Vertex(1).Z 'elm.Vertex(1).z
    minX = 1.7976931348623E+308
    maxY = -1.7976931348623E+308
    
    For i = 1 To mZoneShape.VerticesCount
        If mZoneShape.Vertex(i).X < minX Then minX = mZoneShape.Vertex(i).X
        If mZoneShape.Vertex(i).Y > maxY Then maxY = mZoneShape.Vertex(i).Y
    Next i
    
    Dim flashOrigin As Point3d
    flashOrigin.X = minX
    flashOrigin.Y = maxY
    flashOrigin.Z = zcoord
    Debug.Print (CStr(minX), CStr(maxY))
    getFlashOrigin = flashOrigin
End Function
Public Function mergedMVPolesCollection() As Collection '(ByVal col1 As Collection, ByVal col2 As Collection) As Collection
    'Add items from col2 to col1 and return the result
    'ByVal means we are only looking at copies of the collections (the values within them)
    'The function returns a NEW merged collection
    Dim mvpoles, mvlvpoles As Collection
     mvpoles = mMVpoles
     mvlvpoles = mMVSharingPoles
    Debug.Print (mvpoles.count)

    Dim el As Element
    For Each el In mvlvpoles ' col2.count
        mvpoles.Add (el, CStr(el.ID.low))
    Next
    Debug.Print (mvpoles.count)
    
     mergedMVPolesCollection = mvpoles 'set return value
End Function
Public Function ToString() As String

'This function needs to be custom made in any case, this works for non -object members
    Dim s As String

   

    s = "LVStays = " & CStr(mLVStays.count) & vbCrLf & "LVStruts = " & CStr(mLVStruts.count) & vbCrLf & _
        "MVStays = " & CStr(mMVStays.count) & vbCrLf & "MVStruts = " & CStr(mMVStruts.count) & vbCrLf & _
        "TRFRCircles = " & CStr(mTRFRCircles.count) & vbCrLf & _
        "TRFRNameplates = " & CStr(mTRFRNameplates.count) & vbCrLf & _
        "LVpoles = " & CStr(mLVpoles.count) & vbCrLf & _
        "MVpoles = " & CStr(mMVpoles.count) & vbCrLf & _
        "MVSharingPoles = " & CStr(mMVSharingPoles.count) & vbCrLf & _
        "KickerPoles = " & CStr(mKickerPoles.count) & vbCrLf & _
        "LVFlyingStays = " & CStr(mLVFlyingStays.count) & vbCrLf & _
        "LVShortStays = " & CStr(mLVShortStays.count) & vbCrLf & _
        "MVFlyingStays = " & CStr(mMVFlyingStays.count) & vbCrLf & _
        "MVShortStays = " & CStr(mMVShortStays.count) & vbCrLf & _
        "MVFuseIsolators = " & CStr(mMVFuseIsolators.count) & vbCrLf & _
        "TRFRs = " & CStr(mTRFRs.count) & vbCrLf & _
        "houseText = " & CStr(mHouseText.count) & vbCrLf & _
        "TrfrShapes = " & CStr(mTrfrShapes.count) & vbCrLf & _
        "Airdac = " & CStr(mAirdac.count) & vbCrLf & _
        "LVConductor = " & CStr(mLVConductor.count) & vbCrLf & _
        "MVConductor = " & CStr(mMVConductor.count) & vbCrLf & _
        "elseCellElements = " & CStr(mElseCellElements.count) & vbCrLf & _
        "elseTextElements = " & CStr(mElseTextElements.count) & vbCrLf & _
        "elseLineElements = " & CStr(mElseLineElements.count) & vbCrLf & _
        "elseShapeElements = " & CStr(mElseShapeElements.count) & vbCrLf & _
        "elseElements = " & CStr(mElseElements.count)
    ToString = s

End Function

Public Function printCoordinatesMVPoles() As String
    Dim s As String
    Dim delimiter, header As String
    delimiter = ","
    header = "Element ID,Cell Origin X,Cell origin Y,Cell Origin Z" & vbCrLf
    s = header
    Dim el As Element
    For Each el In mMVpoles
        s = s & el.ID.low & "," & el.AsCellElement.Origin.X & "," & el.AsCellElement.Origin.Y & "," & el.AsCellElement.Origin.Z & vbCrLf
    Next
    printCoordinatesMVPoles = s
End Function

Public Function printCoordinatesMVLines() As String
    Dim s, header As String
    Dim el As Element
    header = "Element ID,Start point X,Start point Y,End Point X,End Point Y" & vbCrLf
    s = header
    For Each el In mMVConductor
        s = s & el.ID.low & "," & el.AsLineElement.startPoint.X & "," & el.AsLineElement.startPoint.Y & "," & el.AsLineElement.EndPoint.X & "," & el.AsLineElement.EndPoint.Y & vbCrLf
    Next
    printCoordinatesMVLines = s
End Function

End Module
