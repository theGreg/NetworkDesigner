class SurroundingClass
{
    private void Abtn_Click()
    {
        if (validFenceDefined)
            selectOnlyPhases(true, false, false);
        else
            Interaction.MsgBox("No fence defined");
    }

    private void AlgButton_Click()
    {
        if (!validFenceDefined)
        {
            Interaction.MsgBox("No fence defined");
            return;
        }

        Fence curFence;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 362
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
   */
        Set curFence = ActiveDesignFile.Fence

 
        ShapeElement elm;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 434
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

    Set elm = curFence.CreateElement()

 */
        ElementEnumerator oFenceEnum;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 520
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    Set oFenceEnum = curFence.GetContents

 */
        clsDataZone oDataZone;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 603
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    Set oDataZone = New clsDataZone

 */
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 640
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    Set oDataZone.ZoneShape = elm

 */
        selectZons(oEnum: oFenceEnum, oSelZone: oDataZone);
        Collection parr;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 758
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    Set parr = oDataZone.getRMDPoints

 */

        Collection tempr;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 838
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    Set tempr = parr

 */
        QuickSort(parr, 1, parr.count);
        ReticMaster.App rmd;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 936
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
 Set rmd = New ReticMaster.App

 */
    }

    private void AuxCellsBtn_Click()
    {
        if (validFenceDefined)
            selectNormalCells(auxCBox.Text);
        else
            Interaction.MsgBox("No fence defined");
    }

    private void auxStatsBtn_Click()
    {
        CadInputQueue.SendCommand("2", true);
        CadInputQueue.SendCommand("1", true);
    }

    private void ConductorStatsBtn_Click()
    {
        if (validFenceDefined)
            getConductorStats();
        else
            Interaction.MsgBox("No fence defined");
    }


    private void Bbtn_Click()
    {
        if (validFenceDefined)
            selectOnlyPhases(false, true, false);
        else
            Interaction.MsgBox("No fence defined");
    }

    private void Cbtn_Click()
    {
        if (validFenceDefined)
            selectOnlyPhases(false, false, true);
        else
            Interaction.MsgBox("No fence defined");
    }

    private void clearBtn_Click()
    {
        DataScanner.outputBox.Text = "";
    }




    private void CommandButton1_Click()
    {
        ElementEnumerator tEnum;
        string tRat;
        string Tname;
        string textOut;
        textOut = "";
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 2358
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    'rat = "lkl"
    'tname = "lko"
    
    'if activemodelreference.
    
    Set tEnum = ActiveModelReference.GetSelectedElements

 */
        while (tEnum.MoveNext)

            ChangeCell104(ECell: tEnum.Current);
        // outputBox.Text = textOut
        ActiveModelReference.UnselectAllElements();
    }

    private void conductorSelect_Click()
    {
        if (validFenceDefined)
            selectConductor(condCBox.Text);
        else
            Interaction.MsgBox("No fence defined");
    }

    private void CopyToCSV_Click()
    {
        DataScanner.outputBox.Copy();
    }

    private void DeselectAll_Click()
    {
        ActiveModelReference.UnselectAllElements();
    }

    private void Duplicates_Click()
    {
        if (validFenceDefined)
            selectNormalCellsDups("LVStruts");
        else
            Interaction.MsgBox("No fence defined");
    }

    private void LVStatsBtn_Click()
    {
        if (validFenceDefined)
            selectNormalCells(ComboBox1.Text);
        else
            Interaction.MsgBox("No fence defined");
    }

    private void exportToXLBtn_Click()
    {
        object oXL;
        object oBook;
        object oSheet;
        int i;
        long lvlCount;
        int colCount;
        string lvlName;
        Levels oLevels;

        Level oLvl;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 3755
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 

    Set oXL = CreateObject("Excel.Application")

 */
        oXL.Visible = true;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 3830
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    Set oBook = oXL.Workbooks.Add

 */
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 3865
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    
    Set oSheet = oBook.Sheets(1)

 */    // oSheet.Name = CStr(vb.Date) '"greg"
       // oBook.FullName = "1731-Mqanduli PH2 2019_20" '&
        string[] lineRows;
        string[] columnItems;

        DataScanner.outputBox.SetFocus();
        lvlCount = DataScanner.outputBox.LineCount;
        lineRows = Split(DataScanner.outputBox.Text, Constants.vbNewLine);
        var loopTo = Information.UBound(lineRows);
        for (i = 0; i <= loopTo; i++)
        {
            columnItems = Microsoft.VisualBasic.Strings.Split(lineRows[i], ",");
            var loopTo1 = Information.UBound(columnItems);
            for (colCount = 0; colCount <= loopTo1; colCount++)
                oSheet.cells(i + 1, colCount + 1).value = columnItems[colCount];
        }
    }

    private void getTrfrCellText_Click()
    {
        // If validFenceDefined Then
        // selectOnlyHseText False
        // Else
        // MsgBox "No fence defined"
        // End If

        ElementEnumerator tEnum;
        string tRat;
        string Tname;
        string textOut;
        textOut = "";
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 4789
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    'rat = "lkl"
    'tname = "lko"
    
    'if activemodelreference.
    
    Set tEnum = ActiveModelReference.GetSelectedElements

 */
        while (tEnum.MoveNext)
        {
            if (tEnum.Current.type != msdElementTypeCellHeader)
            {
                Interaction.MsgBox("One or more selected items is not a transformer cell");
                return;
            }


            GetTransformerCellText(ECell: tEnum.Current, TrfrName: Tname, Rating: tRat); textOut = textOut + Tname + "," + tRat + Constants.vbCrLf;
        }

        outputBox.Text = textOut;
    }

    private void getTrfStats_Clickn()
    {
        ElementEnumerator tEnum;
        ElementEnumerator trfEnum;
        string tRat;
        string Tname;
        string textOut;
        string currZoneInfo;
        currZoneInfo = "";
        textOut = "";
        int connections;
        int phaseACount;
        int phaseBCount;
        int phaseCCount;

        connetions = 0;
        phaseACount = 0;
        phaseBCount = 0;
        phaseCCount = 0;
        if (!validFenceDefined)
            Interaction.MsgBox("No fence defined");
        selectNormalCells("TRFRS");
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 6020
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    ''SelectAllZones
    
    Set tEnum = ActiveModelReference.GetSelectedElements

 */
        tEnum.Reset();
        while (tEnum.MoveNext)
        {
            // 1. Create fence from shape GetTransformerCellText eCell:=tEnum.Current, TrfrName:=Tname, Rating:=tRat '
            // createFenceFromShape
            // 2. get transformer - collect size and rating
            // If validFenceDefined Then
            // selectNormalCells "TRFRS"
            // End If
            if (ActiveModelReference.AnyElementsSelected)
            {
                ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 6500
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
            'Dim trfEnum As ElementEnumerator
            Set trfEnum = ActiveModelReference.GetSelectedElements

 */
                trfEnum.Reset();
            }

            // Set trfEnum = ActiveModelReference.GetSelectedElements
            GetTransformerCellText(ECell: trfEnum.Current, TrfrName: Tname, Rating: tRat);          // 3. Get number of connections
                                                                                                    // 'selectConductor "Service Cable"
            GetServiceCable(Airdac: connections);
            // 4. get phase Allocation
            GetPhaseAllocation(A_count: phaseACount, B_count: phaseBCount, C_count: phaseCCount);
            // 5. Collate all information
            currZoneInfo = Tname + tRat + phaseACount + phaseBCount + phaseCCount;
            textOut = textOut + currZoneInfo + Constants.vbCrLf;
        }

        outputBox.Text = textOut;
    }

    public void NamedFence()
    {
        // Obtain an array of fences' names
        string[] fences;
        fences = ActiveDesignFile.Fence.GetFenceNames;

        // Go through the receive array of fence names
        int highBound;
        highBound = Information.UBound(fences);

        int counter;
        var loopTo = highBound;
        for (counter = 0; counter <= loopTo; counter++)
            ProcessFence(fences[counter]);
    }

    private void getTrfStats_Click()
    {
        string stats;
        string Tname;
        string tRat;
        string lineOut;
        ActiveModelReference.UnselectAllElements();
        ShapeElement[] trfrZones;
        ElementEnumerator trfrZoneEnum;


        // Enumerate fence content and do what you need

        if (validFenceDefined)
        {
            // select all trfr zones in this fence
            // selectTRFRzones
            selectTRFRzoneShapes();
            ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 8109
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
        Set trfrZoneEnum = ActiveModelReference.GetSelectedElements 'ActiveDesignFile.Fence.GetContents

 */
            trfrZoneEnum.Reset();
            CellElement trfrCell;
            ElementEnumerator trfrCellEnum;
            ElementEnumerator serviceCableEnum;
            ActiveModelReference.UnselectAllElements();
            double startTime;
            // trfrZones = enumerator.BuildArrayFromContents
            double endTime;
            startTime = DateTime.Timer;
            while (trfrZoneEnum.MoveNext)
            {
                // with each trfr zone
                // 1. create fence from trfr zone
                if ((!createFenceFromShapeElement(sElement: trfrZoneEnum.Current)))
                {
                    Interaction.MsgBox("Something wronh, not sure yet");
                    return;
                }
                // 2 Select trfr cell
                selectNormalCells("TRFRS");
                ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 8941
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
            
            ' 3. Get TRFR Cell text
            Set trfrCellEnum = ActiveModelReference.GetSelectedElements

 */
                while (trfrCellEnum.MoveNext)
                    ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 9109
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
                Set trfrCell = trfrCellEnum.Current

 */
                GetTransformerCellText(ECell: trfrCell, TrfrName: Tname, Rating: tRat);

                // 4. Get number of customers
                if ((!createFenceFromShapeElement(sElement: trfrZoneEnum.Current)))
                {
                    Interaction.MsgBox("Something wronh, not sure yet");
                    return;
                }

                selectConductor("Service Cable");
                ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 9568
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
            Set serviceCableEnum = ActiveModelReference.GetSelectedElements

 */
                int scount;
                // ' scount = 0
                // 'Do While serviceCableEnum.MoveNext
                // 'Set trfrCell = trfrCellEnum.Current
                // '  scount = scount + 1
                // 'Loop
                scount = UBound(serviceCableEnum.BuildArrayFromContents());
                lineOut = Tname + "," + tRat + "," + System.Convert.ToString(scount) + Constants.vbCrLf;
                stats = stats + lineOut;
            }
            // get all stats on the current selection

            // GetAllStats outputTxt:=stats
            outputBox.Text = outputBox.Text + stats + Constants.vbCrLf;
        }
        else
            Interaction.MsgBox("No fence defined");
        endTime = DateTime.Timer;

        outputBox.Text = outputBox.Text + System.Convert.ToString(endTime - startTime) + Constants.vbCrLf;
    }

    private void hseTextBtn_Click()
    {
        if (validFenceDefined)
            selectOnlyHseText(); // True
        else
            Interaction.MsgBox("No fence defined");
    }

    private void phasesBTn_Click()
    {
        if (validFenceDefined)
            // Dim aPhase, bPhase, cPhase As Boolean

            selectOnlyPhases(true, true, true);
        else
            Interaction.MsgBox("No fence defined");
    }

    private void outputBox_Change()
    {
        exportToXLBtn.Enabled = true;
    }

    private void phaseStatsBtn_Click()
    {
        if (validFenceDefined)
            selectOnlyPhases(true, true, true);
        else
            Interaction.MsgBox("No fence defined");
    }

    private void poleTopsbtn_Click()
    {
        if (validFenceDefined)
        {
            Collection s;
            ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 11136
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
    Set s = New Collection

 */
            s.Add(1);
            s.Add(1);
            s.Add(2);
            s.Add(3);
            s.Add(5);
            s.Add(6);
            s.Add(3);
            s.Add(1);
            MsgBox(getDenominations(list: s));
        }
        else
            Interaction.MsgBox("No fence defined .");
    }

    private void RunStats_Click()
    {
        if (validFenceDefined)
            stats(MVLVSCBox.Text);
        else
            Interaction.MsgBox("No fence defined .");
    }

    private void selctTrfBtn_Click()
    {
        if (validFenceDefined)
            selectNormalCells("TRFRS");
        else
            Interaction.MsgBox("No fence defined");
    }

    private void SelectAllBtn_Click()
    {
        if (validFenceDefined)
            selectAll();
        else
            Interaction.MsgBox("No valid fence defined");
    }

    private void selectTRFRZonesShapes_Click()
    {
        if (validFenceDefined)
            selectTRFRzoneShapes();
        else
            Interaction.MsgBox("No fence defined");
    }

    private void SettingsBtn_Click()
    {
        SettingsForm.Show();
    }

    private void testExports_Click()
    {
        ElementEnumerator tEnum;
        // Dim tRat As String
        // Dim Tname As String
        string textOut;
        textOut = "";
        Point3d p3dStart;
        Point3d p3dEnd;
        ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 12387
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
       
    Set tEnum = ActiveModelReference.GetSelectedElements

 */
        while (tEnum.MoveNext)
        {
            p3dStart = tEnum.Current.AsLineElement.startPoint;
            p3dEnd = tEnum.Current.AsLineElement.EndPoint;

            textOut = textOut + System.Convert.ToString(p3dStart.X) + "," + System.Convert.ToString(p3dStart.Y) + " " + System.Convert.ToString(p3dEnd.X) + "," + System.Convert.ToString(p3dEnd.Y) + Constants.vbCrLf;
        }

        outputBox.Text = textOut;
    }

    private void UserForm_Activate()
    {
        {
            var withBlock = phCBox;
            withBlock.AddItem("All");
            withBlock.AddItem("Phase A");
            withBlock.AddItem("Phase B");
            withBlock.AddItem("Phase C");
            withBlock.Text = "All";
        }

        {
            var withBlock1 = auxCBox;
            withBlock1.AddItem("LV Stays");
            withBlock1.AddItem("LV Poles");
            withBlock1.AddItem("LV Struts");
            withBlock1.AddItem("LV Flying Stays");
            withBlock1.AddItem("MV Stays");
            withBlock1.AddItem("MV Poles");
            withBlock1.AddItem("MVLV Poles");
            withBlock1.AddItem("MV Struts");
            withBlock1.AddItem("MV Flying Stays");
            withBlock1.Text = "LVStays";
        }

        {
            var withBlock2 = condCBox;
            withBlock2.AddItem("Service Cable");
            withBlock2.AddItem("LV Lines");
            withBlock2.AddItem("MV Lines");
            withBlock2.Text = "Service Cable";
        }

        {
            var withBlock3 = MVLVSCBox;
            withBlock3.AddItem("LV");
            withBlock3.AddItem("MV");
            withBlock3.Text = "LV";
        }


        exportToXLBtn.Enabled = false;
    }
}