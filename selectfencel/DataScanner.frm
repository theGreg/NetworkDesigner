VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataScanner 
   Caption         =   "Data Scanner v1.03 rev 1 - July 2018"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8595
   OleObjectBlob   =   "DataScanner.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Abtn_Click()
  If validFenceDefined Then
        selectOnlyPhases True, False, False
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub AlgButton_Click()
    
    If Not validFenceDefined Then
        MsgBox "No fence defined"
        Exit Sub
    End If
    
    Dim curFence As Fence
    Set curFence = ActiveDesignFile.Fence
    Dim elm As ShapeElement

    Set elm = curFence.CreateElement()
 
    Dim oFenceEnum As ElementEnumerator
    Set oFenceEnum = curFence.GetContents
    
    Dim oDataZone As clsDataZone
    Set oDataZone = New clsDataZone
    Set oDataZone.ZoneShape = elm
    selectZons oEnum:=oFenceEnum, oSelZone:=oDataZone
    Dim parr As Collection
    Set parr = oDataZone.getRMDPoints
    
    
    Dim tempr As Collection
    Set tempr = parr
    
    QuickSort parr, 1, parr.count
    
 Dim rmd As ReticMaster.App
 Set rmd = New ReticMaster.App
 
 
 'Dim nnel As p
 'Set nnel = polesarr.Item(1)
 
 's = rmd.CreateLoadNode(polesarr.Item(1).Origin.X, polesarr.Item(1).Origin.Y, 25465, "nn1", 231, "25 ABC(ABC25)", "dual", "DP1", 32.5, 0, 0, 0, 1, 1, 1, 1, 1, "Domestic", 0, 0, 0, False, "", "general", "", "")
 
    'odatazone.getRotMatrix(minx miny minZ)
       
End Sub

Private Sub AuxCellsBtn_Click()
    If validFenceDefined Then
        selectNormalCells (auxCBox.Text)
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub auxStatsBtn_Click()
    CadInputQueue.SendCommand "2", True
    CadInputQueue.SendCommand "1", True
End Sub

Private Sub ConductorStatsBtn_Click()
    If validFenceDefined Then
        getConductorStats
    Else
        MsgBox "No fence defined"
    End If
End Sub


Private Sub Bbtn_Click()
 If validFenceDefined Then
        selectOnlyPhases False, True, False
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub Cbtn_Click()
 If validFenceDefined Then
        selectOnlyPhases False, False, True
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub clearBtn_Click()
    DataScanner.outputBox.Text = ""
End Sub




Private Sub CommandButton1_Click()

    Dim tEnum As ElementEnumerator
    Dim tRat As String
    Dim Tname As String
    Dim textOut As String
    textOut = ""
    'rat = "lkl"
    'tname = "lko"
    
    'if activemodelreference.
    
    Set tEnum = ActiveModelReference.GetSelectedElements
    
    While tEnum.MoveNext
    
        ChangeCell104 ECell:=tEnum.Current ', TrfrName:=Tname, Rating:=tRat  '
        'textOut = textOut + Tname + "," + tRat + vbCrLf
            
    Wend
    
    'outputBox.Text = textOut
    ActiveModelReference.UnselectAllElements
    
End Sub

Private Sub conductorSelect_Click()
    If validFenceDefined Then
        selectConductor condCBox.Text
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub CopyToCSV_Click()
    DataScanner.outputBox.Copy
    'CopyToCSV
    
End Sub

Private Sub DeselectAll_Click()
    ActiveModelReference.UnselectAllElements
End Sub

Private Sub Duplicates_Click()

    If validFenceDefined Then
        selectNormalCellsDups ("LVStruts")
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub LVStatsBtn_Click()
    If validFenceDefined Then
        selectNormalCells (ComboBox1.Text)
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub exportToXLBtn_Click()
Dim oXL As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim i As Integer, lvlCount As Long, colCount As Integer
    Dim lvlName As String
    Dim oLevels As Levels

    Dim oLvl As Level

    Set oXL = CreateObject("Excel.Application")
    oXL.Visible = True
    Set oBook = oXL.Workbooks.Add
    
    Set oSheet = oBook.Sheets(1)
    'oSheet.Name = CStr(vb.Date) '"greg"
    'oBook.FullName = "1731-Mqanduli PH2 2019_20" '&
    Dim lineRows() As String
    Dim columnItems() As String
    
    DataScanner.outputBox.SetFocus
    lvlCount = DataScanner.outputBox.LineCount
    lineRows = Split(DataScanner.outputBox.Text, vbNewLine)

    For i = 0 To UBound(lineRows)
        columnItems = Split(lineRows(i), ",")

        For colCount = 0 To UBound(columnItems)
            oSheet.cells(i + 1, colCount + 1).value = columnItems(colCount)
        
        Next colCount

    Next i
    
End Sub

Private Sub getTrfrCellText_Click()
    'If validFenceDefined Then
    '    selectOnlyHseText False
    'Else
    '    MsgBox "No fence defined"
    'End If
   
    Dim tEnum As ElementEnumerator
    Dim tRat As String
    Dim Tname As String
    Dim textOut As String
    textOut = ""
    'rat = "lkl"
    'tname = "lko"
    
    'if activemodelreference.
    
    Set tEnum = ActiveModelReference.GetSelectedElements
    
    While tEnum.MoveNext
        If tEnum.Current.type <> msdElementTypeCellHeader Then
            MsgBox "One or more selected items is not a transformer cell"
            Exit Sub
            
        End If
        
        
        GetTransformerCellText ECell:=tEnum.Current, TrfrName:=Tname, Rating:=tRat '
        textOut = textOut + Tname + "," + tRat + vbCrLf
        'gArray1.Add (textOut)
    Wend
    
    outputBox.Text = textOut

    
    
    
    
End Sub

Private Sub getTrfStats_Clickn()

    Dim tEnum As ElementEnumerator
    Dim trfEnum As ElementEnumerator
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
    
    connetions = 0
    phaseACount = 0
    phaseBCount = 0
    phaseCCount = 0
    If Not validFenceDefined Then
          MsgBox "No fence defined"
    End If
    selectNormalCells "TRFRS"
    ''SelectAllZones
    
    Set tEnum = ActiveModelReference.GetSelectedElements
    tEnum.Reset
    While tEnum.MoveNext
        '1. Create fence from shape GetTransformerCellText eCell:=tEnum.Current, TrfrName:=Tname, Rating:=tRat '
        'createFenceFromShape
        '2. get transformer - collect size and rating
        'If validFenceDefined Then
        '    selectNormalCells "TRFRS"
        'End If
        If ActiveModelReference.AnyElementsSelected Then
            'Dim trfEnum As ElementEnumerator
            Set trfEnum = ActiveModelReference.GetSelectedElements
            trfEnum.Reset
        End If
        
        'Set trfEnum = ActiveModelReference.GetSelectedElements
        GetTransformerCellText ECell:=trfEnum.Current, TrfrName:=Tname, Rating:=tRat  '
        '3. Get number of connections
        ''selectConductor "Service Cable"
        GetServiceCable Airdac:=connections
        '4. get phase Allocation
        GetPhaseAllocation A_count:=phaseACount, B_count:=phaseBCount, C_count:=phaseCCount
        '5. Collate all information
        currZoneInfo = Tname + tRat + phaseACount + phaseBCount + phaseCCount
        textOut = textOut + currZoneInfo + vbCrLf
    Wend
    
    outputBox.Text = textOut

End Sub
Sub NamedFence()
    ' Obtain an array of fences' names
    Dim fences() As String
    fences = ActiveDesignFile.Fence.GetFenceNames

    ' Go through the receive array of fence names
    Dim highBound As Integer
    highBound = UBound(fences)

    Dim counter As Integer
    For counter = 0 To highBound
        ProcessFence (fences(counter))
    Next
End Sub
Private Sub getTrfStats_Click()
    Dim stats As String, Tname As String, tRat As String, lineOut As String
    ActiveModelReference.UnselectAllElements
    Dim trfrZones() As ShapeElement
    Dim trfrZoneEnum As ElementEnumerator
    

    ' Enumerate fence content and do what you need
   
    If validFenceDefined Then
        'select all trfr zones in this fence
        'selectTRFRzones
        selectTRFRzoneShapes
        Set trfrZoneEnum = ActiveModelReference.GetSelectedElements 'ActiveDesignFile.Fence.GetContents
        trfrZoneEnum.Reset
        Dim trfrCell As CellElement
        Dim trfrCellEnum As ElementEnumerator
        Dim serviceCableEnum As ElementEnumerator
        ActiveModelReference.UnselectAllElements
        'trfrZones = enumerator.BuildArrayFromContents
        Dim startTime As Double, endTime As Double
        startTime = Timer
        Do While trfrZoneEnum.MoveNext
            'with each trfr zone
            ' 1. create fence from trfr zone
            If (Not createFenceFromShapeElement(sElement:=trfrZoneEnum.Current)) Then
               MsgBox "Something wronh, not sure yet"
               Exit Sub
            End If
            '2 Select trfr cell
            selectNormalCells "TRFRS"
            
            ' 3. Get TRFR Cell text
            Set trfrCellEnum = ActiveModelReference.GetSelectedElements
            Do While trfrCellEnum.MoveNext
                Set trfrCell = trfrCellEnum.Current
            Loop
            GetTransformerCellText ECell:=trfrCell, TrfrName:=Tname, Rating:=tRat
            
            '4. Get number of customers
             If (Not createFenceFromShapeElement(sElement:=trfrZoneEnum.Current)) Then
               MsgBox "Something wronh, not sure yet"
               Exit Sub
             End If
            
            selectConductor ("Service Cable")
            Set serviceCableEnum = ActiveModelReference.GetSelectedElements
            Dim scount As Integer
           '' scount = 0
            ''Do While serviceCableEnum.MoveNext
                ''Set trfrCell = trfrCellEnum.Current
              ''  scount = scount + 1
            ''Loop
            scount = UBound(serviceCableEnum.BuildArrayFromContents())
            lineOut = Tname + "," + tRat + "," + CStr(scount) + vbCrLf
            stats = stats + lineOut
        Loop
        'get all stats on the current selection
        
        'GetAllStats outputTxt:=stats
        outputBox.Text = outputBox.Text + stats + vbCrLf
    Else
        MsgBox "No fence defined"
    End If
    endTime = Timer
    
    outputBox.Text = outputBox.Text + CStr(endTime - startTime) + vbCrLf
End Sub

Private Sub hseTextBtn_Click()
    If validFenceDefined Then
        selectOnlyHseText 'True
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub phasesBTn_Click()

    If validFenceDefined Then
       ' Dim aPhase, bPhase, cPhase As Boolean
        
        selectOnlyPhases True, True, True
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub outputBox_Change()
    exportToXLBtn.Enabled = True
End Sub

Private Sub phaseStatsBtn_Click()
    If validFenceDefined Then
        selectOnlyPhases True, True, True
    Else
        MsgBox "No fence defined"
    End If
    
End Sub

Private Sub poleTopsbtn_Click()
    If validFenceDefined Then
    Dim s As Collection
    Set s = New Collection
    s.Add 1
    s.Add 1
    s.Add 2
    s.Add 3
    s.Add 5
    s.Add 6
    s.Add 3
    s.Add 1
    
    
        MsgBox getDenominations(list:=s)
    Else
        MsgBox "No fence defined ."
    End If
End Sub

Private Sub RunStats_Click()
    If validFenceDefined Then
        stats MVLVSCBox.Text
    Else
        MsgBox "No fence defined ."
    End If
End Sub

Private Sub selctTrfBtn_Click()
    If validFenceDefined Then
        selectNormalCells "TRFRS"
        'selectNormalCells "006"
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub SelectAllBtn_Click()

    If validFenceDefined Then
           selectAll
    Else
           MsgBox "No valid fence defined"
    End If
End Sub

Private Sub selectTRFRZonesShapes_Click()
    If validFenceDefined Then
        selectTRFRzoneShapes
    Else
        MsgBox "No fence defined"
    End If
End Sub

Private Sub SettingsBtn_Click()
    SettingsForm.Show
    
End Sub

Private Sub testExports_Click()

   Dim tEnum As ElementEnumerator
    'Dim tRat As String
    'Dim Tname As String
    Dim textOut As String
    textOut = ""
    Dim p3dStart As Point3d
    Dim p3dEnd As Point3d
       
    Set tEnum = ActiveModelReference.GetSelectedElements
    
    While tEnum.MoveNext
        p3dStart = tEnum.Current.AsLineElement.startPoint
        p3dEnd = tEnum.Current.AsLineElement.EndPoint
        
        textOut = textOut + CStr(p3dStart.X) + "," + CStr(p3dStart.Y) + " " + CStr(p3dEnd.X) + "," + CStr(p3dEnd.Y) + vbCrLf
            
    Wend
    
    outputBox.Text = textOut
End Sub

Private Sub UserForm_Activate()
   
    With phCBox
        .AddItem "All"
        .AddItem "Phase A"
        .AddItem "Phase B"
        .AddItem "Phase C"
        .Text = "All"
    End With
    
    With auxCBox
        .AddItem "LV Stays"
        .AddItem "LV Poles"
        .AddItem "LV Struts"
        .AddItem "LV Flying Stays"
        .AddItem "MV Stays"
        .AddItem "MV Poles"
        .AddItem "MVLV Poles"
        .AddItem "MV Struts"
        .AddItem "MV Flying Stays"
        .Text = "LVStays"
    End With
    
    With condCBox
        .AddItem "Service Cable"
        .AddItem "LV Lines"
        .AddItem "MV Lines"
        .Text = "Service Cable"
    End With
    
    With MVLVSCBox
        .AddItem "LV"
        .AddItem "MV"
        .Text = "LV"
    End With
    
    
        exportToXLBtn.Enabled = False
        'exportToXLBtn.Visible = False
   
    
End Sub
