
Option Explicit
'// Class: clsDataZone
'// A class for...
Private mZoneShape As ShapeElement
Private mLVStays As Collection
Private mLVStruts As Collection
Private mMVStays As Collection
Private mMVStruts As Collection
Private mTRFRCircles As Collection
Private mTRFRNameplates As Collection
Private mLVpoles As Collection
Private mMVpoles As Collection
Private mMVSharingPoles As Collection
Private mKickerPoles As Collection
Private mLVFlyingStays As Collection
Private mLVShortStays As Collection
Private mMVFlyingStays As Collection
Private mMVShortStays As Collection
Private mMVFuseIsolators As Collection
Private mTRFRs As Collection
Private mHouseText As Collection
Private mTrfrShapes As Collection
Private mAirdac As Collection
Private mLVConductor As Collection
Private mMVConductor As Collection
Private mElseCellElements As Collection
Private mElseTextElements As Collection
Private mElseLineElements As Collection
Private mElseShapeElements As Collection
Private mElseElements As Collection
Private Sub Class_Initialize()
    
     mLVStays = New Collection '//...(Type = Collection)
     mLVStruts = New Collection '//...(Type = Collection)
     mMVStays = New Collection '//...(Type = Collection)
     mMVStruts = New Collection '//...(Type = Collection)
     mTRFRCircles = New Collection '//...(Type = Collection)
     mTRFRNameplates = New Collection '//...(Type = Collection)
     mLVpoles = New Collection '//...(Type = Collection)
     mMVpoles = New Collection '//...(Type = Collection)
     mMVSharingPoles = New Collection '//...(Type = Collection)
     mKickerPoles = New Collection '//...(Type = Collection)
     mLVFlyingStays = New Collection '//...(Type = Collection)
     mLVShortStays = New Collection '//...(Type = Collection)
     mMVFlyingStays = New Collection '//...(Type = Collection)
     mMVShortStays = New Collection '//...(Type = Collection)
     mMVFuseIsolators = New Collection '//...(Type = Collection)
     mTRFRs = New Collection '//...(Type = Collection)
     mHouseText = New Collection '//...(Type = Collection)
     mTrfrShapes = New Collection '//...(Type = Collection)
     mAirdac = New Collection '//...(Type = Collection)
     mLVConductor = New Collection '//...(Type = Collection)
     mMVConductor = New Collection '//...(Type = Collection)
     mElseCellElements = New Collection '//...(Type = Collection)
     mElseTextElements = New Collection '//...(Type = Collection)
     mElseLineElements = New Collection '//...(Type = Collection)
     mElseShapeElements = New Collection '//...(Type = Collection)
     mElseElements = New Collection '//...(Type = Collection)
End Sub

Public Property Get LVStays() As Collection
     LVStays = mLVStays
End Property
Public Property  LVStays(value As Collection)
     mLVStays = value
End Property

Public Property Get LVStruts() As Collection
     LVStruts = mLVStruts
End Property
Public Property  LVStruts(value As Collection)
     mLVStruts = value
End Property

Public Property Get MVStays() As Collection
     MVStays = mMVStays
End Property
Public Property  MVStays(value As Collection)
     mMVStays = value
End Property

Public Property Get MVStruts() As Collection
     MVStruts = mMVStruts
End Property
Public Property  MVStruts(value As Collection)
     mMVStruts = value
End Property

Public Property Get TRFRCircles() As Collection
     TRFRCircles = mTRFRCircles
End Property
Public Property  TRFRCircles(value As Collection)
     mTRFRCircles = value
End Property

Public Property Get TRFRNameplates() As Collection
     TRFRNameplates = mTRFRNameplates
End Property
Public Property  TRFRNameplates(value As Collection)
     mTRFRNameplates = value
End Property

Public Property Get LVpoles() As Collection
     LVpoles = mLVpoles
End Property
Public Property  LVpoles(value As Collection)
     mLVpoles = value
End Property

Public Property Get mvpoles() As Collection
     mvpoles = mMVpoles
End Property
Public Property  mvpoles(value As Collection)
     mMVpoles = value
End Property

Public Property Get MVSharingPoles() As Collection
     MVSharingPoles = mMVSharingPoles
End Property
Public Property  MVSharingPoles(value As Collection)
     mMVSharingPoles = value
End Property

Public Property Get KickerPoles() As Collection
     KickerPoles = mKickerPoles
End Property
Public Property  KickerPoles(value As Collection)
     mKickerPoles = value
End Property

Public Property Get LVFlyingStays() As Collection
     LVFlyingStays = mLVFlyingStays
End Property
Public Property  LVFlyingStays(value As Collection)
     mLVFlyingStays = value
End Property

Public Property Get LVShortStays() As Collection
     LVShortStays = mLVShortStays
End Property
Public Property  LVShortStays(value As Collection)
     mLVShortStays = value
End Property

Public Property Get MVFlyingStays() As Collection
     MVFlyingStays = mMVFlyingStays
End Property
Public Property  MVFlyingStays(value As Collection)
     mMVFlyingStays = value
End Property

Public Property Get MVShortStays() As Collection
     MVShortStays = mMVShortStays
End Property
Public Property  MVShortStays(value As Collection)
     mMVShortStays = value
End Property

Public Property Get MVFuseIsolators() As Collection
     MVFuseIsolators = mMVFuseIsolators
End Property
Public Property  MVFuseIsolators(value As Collection)
     mMVFuseIsolators = value
End Property

Public Property Get TRFRs() As Collection
     TRFRs = mTRFRs
End Property
Public Property  TRFRs(value As Collection)
     mTRFRs = value
End Property

Public Property Get houseText() As Collection
     houseText = mHouseText
End Property
Public Property  houseText(value As Collection)
     mHouseText = value
End Property

Public Property Get TrfrShapes() As Collection
     TrfrShapes = mTrfrShapes
End Property
Public Property  TrfrShapes(value As Collection)
     mTrfrShapes = value
End Property

Public Property Get Airdac() As Collection
     Airdac = mAirdac
End Property
Public Property  Airdac(value As Collection)
     mAirdac = value
End Property

Public Property Get LVConductor() As Collection
     LVConductor = mLVConductor
End Property
Public Property  LVConductor(value As Collection)
     mLVConductor = value
End Property

Public Property Get MVConductor() As Collection
     MVConductor = mMVConductor
End Property
Public Property  MVConductor(value As Collection)
     mMVConductor = value
End Property

Public Property Get elseCellElements() As Collection
     elseCellElements = mElseCellElements
End Property
Public Property  elseCellElements(value As Collection)
     mElseCellElements = value
End Property

Public Property Get elseTextElements() As Collection
     elseTextElements = mElseTextElements
End Property
Public Property  elseTextElements(value As Collection)
     mElseTextElements = value
End Property

Public Property Get elseLineElements() As Collection
     elseLineElements = mElseLineElements
End Property
Public Property  elseLineElements(value As Collection)
     mElseLineElements = value
End Property

Public Property Get elseShapeElements() As Collection
     elseShapeElements = mElseShapeElements
End Property
Public Property  elseShapeElements(value As Collection)
     mElseShapeElements = value
End Property
Public Property Get elseElements() As Collection
     elseElements = mElseElements
End Property
Public Property  elseElements(value As Collection)
     mElseElements = value
End Property

Public Property Get ZoneShape() As ShapeElement
     ZoneShape = mZoneShape
End Property
Public Property  ZoneShape(value As ShapeElement)
     mZoneShape = value
End Property
Public Function GetTree(ByVal conductorList As Collection, polesList As Collection) As String

    Dim out As String
    for each inp
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
        
        polesarr.Add pp, CStr(cell.ID.low)
        xCoords.Add pp.X, CStr(cell.ID.low)
        
        
        
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
    Debug.Print CStr(minX), CStr(maxY)
    getFlashOrigin = flashOrigin
End Function
Public Function mergedMVPolesCollection() As Collection '(ByVal col1 As Collection, ByVal col2 As Collection) As Collection
    'Add items from col2 to col1 and return the result
    'ByVal means we are only looking at copies of the collections (the values within them)
    'The function returns a NEW merged collection
    Dim mvpoles, mvlvpoles As Collection
     mvpoles = mMVpoles
     mvlvpoles = mMVSharingPoles
    Debug.Print mvpoles.count

    Dim el As Element
    For Each el In mvlvpoles ' col2.count
        mvpoles.Add el, CStr(el.ID.low)
    Next
    Debug.Print mvpoles.count
    
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
Private Sub Class_Terminate()
     mLVStays = Nothing '//...(Type = Collection)
     mLVStruts = Nothing '//...(Type = Collection)
     mMVStays = Nothing '//...(Type = Collection)
     mMVStruts = Nothing '//...(Type = Collection)
     mTRFRCircles = Nothing '//...(Type = Collection)
     mTRFRNameplates = Nothing '//...(Type = Collection)
     mLVpoles = Nothing '//...(Type = Collection)
     mMVpoles = Nothing '//...(Type = Collection)
     mMVSharingPoles = Nothing '//...(Type = Collection)
     mKickerPoles = Nothing '//...(Type = Collection)
     mLVFlyingStays = Nothing '//...(Type = Collection)
     mLVShortStays = Nothing '//...(Type = Collection)
     mMVFlyingStays = Nothing '//...(Type = Collection)
     mMVShortStays = Nothing '//...(Type = Collection)
     mMVFuseIsolators = Nothing '//...(Type = Collection)
     mTRFRs = Nothing '//...(Type = Collection)
     mHouseText = Nothing '//...(Type = Collection)
     mTrfrShapes = Nothing '//...(Type = Collection)
     mAirdac = Nothing '//...(Type = Collection)
     mLVConductor = Nothing '//...(Type = Collection)
     mMVConductor = Nothing '//...(Type = Collection)
     mElseCellElements = Nothing '//...(Type = Collection)
     mElseTextElements = Nothing '//...(Type = Collection)
     mElseLineElements = Nothing '//...(Type = Collection)
     mElseShapeElements = Nothing '//...(Type = Collection)
     mElseElements = Nothing '//...(Type = Collection)
    
End Sub



