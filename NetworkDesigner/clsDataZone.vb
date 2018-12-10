VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    
    Set mLVStays = New Collection '//...(Type = Collection)
    Set mLVStruts = New Collection '//...(Type = Collection)
    Set mMVStays = New Collection '//...(Type = Collection)
    Set mMVStruts = New Collection '//...(Type = Collection)
    Set mTRFRCircles = New Collection '//...(Type = Collection)
    Set mTRFRNameplates = New Collection '//...(Type = Collection)
    Set mLVpoles = New Collection '//...(Type = Collection)
    Set mMVpoles = New Collection '//...(Type = Collection)
    Set mMVSharingPoles = New Collection '//...(Type = Collection)
    Set mKickerPoles = New Collection '//...(Type = Collection)
    Set mLVFlyingStays = New Collection '//...(Type = Collection)
    Set mLVShortStays = New Collection '//...(Type = Collection)
    Set mMVFlyingStays = New Collection '//...(Type = Collection)
    Set mMVShortStays = New Collection '//...(Type = Collection)
    Set mMVFuseIsolators = New Collection '//...(Type = Collection)
    Set mTRFRs = New Collection '//...(Type = Collection)
    Set mHouseText = New Collection '//...(Type = Collection)
    Set mTrfrShapes = New Collection '//...(Type = Collection)
    Set mAirdac = New Collection '//...(Type = Collection)
    Set mLVConductor = New Collection '//...(Type = Collection)
    Set mMVConductor = New Collection '//...(Type = Collection)
    Set mElseCellElements = New Collection '//...(Type = Collection)
    Set mElseTextElements = New Collection '//...(Type = Collection)
    Set mElseLineElements = New Collection '//...(Type = Collection)
    Set mElseShapeElements = New Collection '//...(Type = Collection)
    Set mElseElements = New Collection '//...(Type = Collection)
End Sub

Public Property Get LVStays() As Collection
    Set LVStays = mLVStays
End Property
Public Property Set LVStays(value As Collection)
    Set mLVStays = value
End Property

Public Property Get LVStruts() As Collection
    Set LVStruts = mLVStruts
End Property
Public Property Set LVStruts(value As Collection)
    Set mLVStruts = value
End Property

Public Property Get MVStays() As Collection
    Set MVStays = mMVStays
End Property
Public Property Set MVStays(value As Collection)
    Set mMVStays = value
End Property

Public Property Get MVStruts() As Collection
    Set MVStruts = mMVStruts
End Property
Public Property Set MVStruts(value As Collection)
    Set mMVStruts = value
End Property

Public Property Get TRFRCircles() As Collection
    Set TRFRCircles = mTRFRCircles
End Property
Public Property Set TRFRCircles(value As Collection)
    Set mTRFRCircles = value
End Property

Public Property Get TRFRNameplates() As Collection
    Set TRFRNameplates = mTRFRNameplates
End Property
Public Property Set TRFRNameplates(value As Collection)
    Set mTRFRNameplates = value
End Property

Public Property Get LVpoles() As Collection
    Set LVpoles = mLVpoles
End Property
Public Property Set LVpoles(value As Collection)
    Set mLVpoles = value
End Property

Public Property Get mvpoles() As Collection
    Set mvpoles = mMVpoles
End Property
Public Property Set mvpoles(value As Collection)
    Set mMVpoles = value
End Property

Public Property Get MVSharingPoles() As Collection
    Set MVSharingPoles = mMVSharingPoles
End Property
Public Property Set MVSharingPoles(value As Collection)
    Set mMVSharingPoles = value
End Property

Public Property Get KickerPoles() As Collection
    Set KickerPoles = mKickerPoles
End Property
Public Property Set KickerPoles(value As Collection)
    Set mKickerPoles = value
End Property

Public Property Get LVFlyingStays() As Collection
    Set LVFlyingStays = mLVFlyingStays
End Property
Public Property Set LVFlyingStays(value As Collection)
    Set mLVFlyingStays = value
End Property

Public Property Get LVShortStays() As Collection
    Set LVShortStays = mLVShortStays
End Property
Public Property Set LVShortStays(value As Collection)
    Set mLVShortStays = value
End Property

Public Property Get MVFlyingStays() As Collection
    Set MVFlyingStays = mMVFlyingStays
End Property
Public Property Set MVFlyingStays(value As Collection)
    Set mMVFlyingStays = value
End Property

Public Property Get MVShortStays() As Collection
    Set MVShortStays = mMVShortStays
End Property
Public Property Set MVShortStays(value As Collection)
    Set mMVShortStays = value
End Property

Public Property Get MVFuseIsolators() As Collection
    Set MVFuseIsolators = mMVFuseIsolators
End Property
Public Property Set MVFuseIsolators(value As Collection)
    Set mMVFuseIsolators = value
End Property

Public Property Get TRFRs() As Collection
    Set TRFRs = mTRFRs
End Property
Public Property Set TRFRs(value As Collection)
    Set mTRFRs = value
End Property

Public Property Get houseText() As Collection
    Set houseText = mHouseText
End Property
Public Property Set houseText(value As Collection)
    Set mHouseText = value
End Property

Public Property Get TrfrShapes() As Collection
    Set TrfrShapes = mTrfrShapes
End Property
Public Property Set TrfrShapes(value As Collection)
    Set mTrfrShapes = value
End Property

Public Property Get Airdac() As Collection
    Set Airdac = mAirdac
End Property
Public Property Set Airdac(value As Collection)
    Set mAirdac = value
End Property

Public Property Get LVConductor() As Collection
    Set LVConductor = mLVConductor
End Property
Public Property Set LVConductor(value As Collection)
    Set mLVConductor = value
End Property

Public Property Get MVConductor() As Collection
    Set MVConductor = mMVConductor
End Property
Public Property Set MVConductor(value As Collection)
    Set mMVConductor = value
End Property

Public Property Get elseCellElements() As Collection
    Set elseCellElements = mElseCellElements
End Property
Public Property Set elseCellElements(value As Collection)
    Set mElseCellElements = value
End Property

Public Property Get elseTextElements() As Collection
    Set elseTextElements = mElseTextElements
End Property
Public Property Set elseTextElements(value As Collection)
    Set mElseTextElements = value
End Property

Public Property Get elseLineElements() As Collection
    Set elseLineElements = mElseLineElements
End Property
Public Property Set elseLineElements(value As Collection)
    Set mElseLineElements = value
End Property

Public Property Get elseShapeElements() As Collection
    Set elseShapeElements = mElseShapeElements
End Property
Public Property Set elseShapeElements(value As Collection)
    Set mElseShapeElements = value
End Property
Public Property Get elseElements() As Collection
    Set elseElements = mElseElements
End Property
Public Property Set elseElements(value As Collection)
    Set mElseElements = value
End Property

Public Property Get ZoneShape() As ShapeElement
    Set ZoneShape = mZoneShape
End Property
Public Property Set ZoneShape(value As ShapeElement)
    Set mZoneShape = value
End Property
Public Function GetTree(ByVal conductorList As Collection, polesList As Collection) As String

    Dim out As String
    for each inp
End Function
Public Function getRMDPoints() As Collection 'Matrix(fenceFlashOrigin As Point3d) As Matrix3d
    
    Dim polesarr As Collection
    Dim p, pp As Point3d
    p = getFlashOrigin
    Set polesarr = New Collection
    Dim n As Long
    n = 1
    Dim s As String
    Dim cell As CellElement
    Dim xCoords As Collection
    Set xCoords = New Collection
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
    Set rmd = New ReticMaster.App
    
    Set getRMDPoints = xCoords
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
    Set mvpoles = mMVpoles
    Set mvlvpoles = mMVSharingPoles
    Debug.Print mvpoles.count

    Dim el As Element
    For Each el In mvlvpoles ' col2.count
        mvpoles.Add el, CStr(el.ID.low)
    Next
    Debug.Print mvpoles.count
    
    Set mergedMVPolesCollection = mvpoles 'set return value
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
    Set mLVStays = Nothing '//...(Type = Collection)
    Set mLVStruts = Nothing '//...(Type = Collection)
    Set mMVStays = Nothing '//...(Type = Collection)
    Set mMVStruts = Nothing '//...(Type = Collection)
    Set mTRFRCircles = Nothing '//...(Type = Collection)
    Set mTRFRNameplates = Nothing '//...(Type = Collection)
    Set mLVpoles = Nothing '//...(Type = Collection)
    Set mMVpoles = Nothing '//...(Type = Collection)
    Set mMVSharingPoles = Nothing '//...(Type = Collection)
    Set mKickerPoles = Nothing '//...(Type = Collection)
    Set mLVFlyingStays = Nothing '//...(Type = Collection)
    Set mLVShortStays = Nothing '//...(Type = Collection)
    Set mMVFlyingStays = Nothing '//...(Type = Collection)
    Set mMVShortStays = Nothing '//...(Type = Collection)
    Set mMVFuseIsolators = Nothing '//...(Type = Collection)
    Set mTRFRs = Nothing '//...(Type = Collection)
    Set mHouseText = Nothing '//...(Type = Collection)
    Set mTrfrShapes = Nothing '//...(Type = Collection)
    Set mAirdac = Nothing '//...(Type = Collection)
    Set mLVConductor = Nothing '//...(Type = Collection)
    Set mMVConductor = Nothing '//...(Type = Collection)
    Set mElseCellElements = Nothing '//...(Type = Collection)
    Set mElseTextElements = Nothing '//...(Type = Collection)
    Set mElseLineElements = Nothing '//...(Type = Collection)
    Set mElseShapeElements = Nothing '//...(Type = Collection)
    Set mElseElements = Nothing '//...(Type = Collection)
    
End Sub



