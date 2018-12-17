using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

#region Bentley Namespaces

using BCOM = Bentley.Interop.MicroStationDGN;
using BM = Bentley.MicroStation;
using BMI = Bentley.MicroStation.InteropServices;

#endregion Bentley Namespaces

namespace NetworkDesigner
{
    public static class CommonUtils
    {
        public static void resetSelectBy()
        {
            Program.MSApp.CadInputQueue.SendKeyin("mdl silent load selectby dialog");

            Program.MSApp.CadInputQueue.SendKeyin("selectby type none");

            Program.MSApp.CadInputQueue.SendKeyin("selectby level none");

            Program.MSApp.CadInputQueue.SendKeyin("choose none");
        }

        public static bool CreateDefpointsLevel(BCOM.Point3d startpoint)
        {
            //bool createDefpointsLevel_ = false;

            // ActiveSettings.Level = ActiveDesignFile.Levels("Defpoints")

            if (!(LevelExists("Defpoints")))
            {
                BCOM.Point3d point;

                // Dim startpoint As BCOM.Point3d
                BCOM.Point3d point2;
                // Dim lngTemp As Long

                // Start a command
                Program.MSApp.CadInputQueue.SendCommand("LEVELMANAGER DIALOG TOGGLE");

                Program.MSApp.CadInputQueue.SendCommand("LEVELMANAGER LEVEL NEWFROMGUI");

                // Send a keyin that can be a command string
                Program.MSApp.CadInputQueue.SendKeyin("LEVEL SET NAME \"New Level (0)\" \"Defpoints\"");

                // Set a variable associated with a dialog box
                Program.MSApp.SetCExpressionValue("lvlmangrGlobals.wLevelName", "Defpoints", "LVLMANGR");

                Program.MSApp.CadInputQueue.SendKeyin("LEVEL SET PLOT OFF   \"Defpoints\"");

                point = startpoint;
                point2 = point;
                Program.MSApp.CadInputQueue.SendDragPoints(ref point, ref point2, 1);

                Program.MSApp.CadInputQueue.SendDataPoint(ref point, 1);

                Program.MSApp.CadInputQueue.SendDragPoints(ref point, ref point2, 1);

                Program.MSApp.CommandState.StartDefaultCommand();
            }

            if (LevelExists("Defpoints"))
            {
                if (Program.MSApp.ActiveDesignFile.Levels["Defpoints"].Plot == true)
                    Program.MSApp.ActiveDesignFile.Levels["Defpoints"].Plot = false;
                return true;
            }
            return false;
        }

        public static bool LevelExists(string levelName)
        {
            BCOM.Level oLevel;
            oLevel = Program.MSApp.ActiveModelReference.Levels[levelName];
            return (!oLevel.Equals(null));
        }

        public static bool ValidFenceDefined()
        {
            bool valid;//, noSelectedElem;
            valid = false;
            //noSelectedElem = false;
            if (Program.MSApp.ActiveDesignFile.Fence.IsDefined)
                return true;

            // createFence
            if (!(Program.MSApp.ActiveModelReference.AnyElementsSelected))
            {
                valid = false;
                //noSelectedElem = true;
            }
            else
            // check if selected element fits criteria
            {
                if (SelectedShapeOK())
                {
                    // create fence from it
                    CreateFenceFromShape();
                    // check again if fence has been created
                    if (Program.MSApp.ActiveDesignFile.Fence.IsDefined)
                        valid = true;
                }
                else

                    Program.MSApp.ShowTempMessage(BCOM.MsdStatusBarArea.Left, "No valid shape defined");
            }

            return valid;
        }

        public static void CreateFenceFromShape()
        {
            BCOM.ElementEnumerator oElEnum;
            BCOM.Element oelm;

            oElEnum = Program.MSApp.ActiveModelReference.GetSelectedElements();
            oElEnum.Reset();
            while (oElEnum.MoveNext())
            {
                oelm = oElEnum.Current;

                Program.MSApp.CadInputQueue.SendCommand("LOCK FENCE INSIDE");

                // Set a variable associated with a dialog box
                Program.MSApp.SetCExpressionValue("tcb->msToolSettings.fence.placeMode", 3, "");
                BCOM.View lastView;

                lastView = Program.MSApp.CommandState.LastView();// (lastView);

                BCOM.Fence ofence;

                ofence = Program.MSApp.ActiveDesignFile.Fence;
                ofence.DefineFromElement(lastView, oelm);
            }
        }

        public static bool SelectedShapeOK()
        {
            BCOM.ElementEnumerator oElEnum;// = new BCOM.ElementEnumerator();
            BCOM.Element oEl;
            //int count = 0;
            bool retval = false;
            oElEnum = Program.MSApp.ActiveModelReference.GetSelectedElements();
            oElEnum.Reset();
            while (oElEnum.MoveNext())
            {
                oEl = oElEnum.Current;
                if (MatchParams(oEl, "Level 11", oEl.Type))
                    return true;
            }
            return retval;
        }

        public static bool MatchParams(BCOM.Element oElem, string lvlName, BCOM.MsdElementType elmType)
        {
            bool oSFlag, oLvlFlag, oColFlag, oTypeFlag;
            oSFlag = oLvlFlag = oColFlag = oTypeFlag = false;
            if (oElem.Type == elmType)
                oTypeFlag = true;
            if (oElem.Level.Name == lvlName)
                oTypeFlag = true;

            return oTypeFlag;
        }

        public static BCOM.Element[] ScanByLevel(string levelName)
        {
            // Dim nElements As Long
            BCOM.Element[] arr = new BCOM.Element[1];
            // nElements = 0
            // Find named level
            BCOM.Level oLevel;
            oLevel = Program.MSApp.ActiveModelReference.Levels[levelName];
            if ((oLevel == null))
            {
                MessageBox.Show("Level '" + levelName + "' does not exist", "Nonexistent Level");
                return arr;
            }
            // Set up scan criteria
            BCOM.ElementScanCriteria oScanCriteria;
            oScanCriteria = new BCOM.ElementScanCriteriaClass();
            oScanCriteria.ExcludeAllLevels();
            oScanCriteria.IncludeLevel(oLevel);
            oScanCriteria.ExcludeAllColors();
            oScanCriteria.IncludeColor(7);
            // Perform the scan
            BCOM.Element[] arry = Program.MSApp.ActiveModelReference.Scan(oScanCriteria).BuildArrayFromContents();
            return arry;
        }

        public static void selectNormalCells(string cellType)
        {
            BCOM.Element oElem;
            BCOM.Fence oFence;
            BCOM.ElementEnumerator oEnum;
            string cellName;
            cellName = "001";
            Program.MSApp.ActiveModelReference.UnselectAllElements();
            oFence = Program.MSApp.ActiveDesignFile.Fence;
            oEnum = oFence.GetContents();
            switch (cellType)
            {
                case "LV Stays":
                    {
                        cellName = "001";
                        break;
                    }

                case "LV Struts":
                    {
                        cellName = "002";
                        break;
                    }

                case "MV Stays":
                    {
                        cellName = "003";
                        break;
                    }

                case "MV Struts":
                    {
                        cellName = "004";
                        break;
                    }

                case "LV Poles":
                    {
                        cellName = "009";
                        break;
                    }

                case "MV Poles":
                    {
                        cellName = "010";
                        break;
                    }

                case "MVLV Poles":
                    {
                        cellName = "011";
                        break;
                    }

                case "LV Flying Stays":
                    {
                        cellName = "015";
                        break;
                    }

                case "MV Flying Stays":
                    {
                        cellName = "017";
                        break;
                    }

                case "TRFRS":
                    {
                        cellName = "TXTCS";
                        break;
                    }

                default:
                    {
                        cellName = "000";
                        break;
                    }
            }

            while (oEnum.MoveNext())
            {
                oElem = oEnum.Current;

                if ((oElem.Type == BCOM.MsdElementType.CellHeader))
                {
                    if ((cellType == "All"))
                    {
                        oElem.Redraw(BCOM.MsdDrawingMode.Hilite);
                        Program.MSApp.ActiveModelReference.SelectElement(oElem);
                    }
                    else if ((oElem.AsCellElement().Name == cellName))
                    {
                        oElem.Redraw(BCOM.MsdDrawingMode.Hilite);
                        Program.MSApp.ActiveModelReference.SelectElement(oElem);
                    }
                }
            }
        }

        public static void selectOnlyPhases(string phaseName)
        {
            BCOM.Element oElem;
            BCOM.Fence ofence;
            BCOM.ElementEnumerator oEnum;
            bool aPhase, bPhase, cPhase;
            aPhase = bPhase = cPhase = false;
            int phaseA_count;
            int phaseB_count;
            int phaseC_count;

            if (phaseName == "Red")
                aPhase = true;
            else if (phaseName == "White")
                bPhase = true;
            else if (phaseName == "Blue")
                cPhase = true;
            else
                aPhase = bPhase = cPhase = true;

            Program.MSApp.ActiveModelReference.UnselectAllElements();

            ofence = Program.MSApp.ActiveDesignFile.Fence;
            oEnum = ofence.GetContents();

            phaseA_count = 0;
            phaseB_count = 0;
            phaseC_count = 0;
            while (oEnum.MoveNext())
            {
                oElem = oEnum.Current;

                if (((oElem.Type == BCOM.MsdElementType.Text)))// Program.MSApp.MsdElementTypeText)))
                {
                    if ((aPhase))
                    {
                        if ((oElem.AsTextElement().Text == " R"))
                        {
                            oElem.Redraw(BCOM.MsdDrawingMode.Hilite);
                            Program.MSApp.ActiveModelReference.SelectElement(oElem);
                            phaseA_count = phaseA_count + 1;
                        }
                    }
                    if ((bPhase))
                    {
                        if ((oElem.AsTextElement().Text == " W"))
                        {
                            oElem.Redraw(BCOM.MsdDrawingMode.Hilite);
                            Program.MSApp.ActiveModelReference.SelectElement(oElem);
                            phaseB_count = phaseB_count + 1;
                        }
                    }
                    if ((cPhase))
                    {
                        if ((oElem.AsTextElement().Text == " B"))
                        {
                            oElem.Redraw(BCOM.MsdDrawingMode.Hilite);
                            Program.MSApp.ActiveModelReference.SelectElement(oElem);
                            phaseC_count = phaseC_count + 1;
                        }
                    }
                }
            }

            if (aPhase & bPhase & cPhase)
                MessageBox.Show("Phase A :" + phaseA_count + Environment.NewLine + "Phase B :" + phaseB_count + Environment.NewLine + "Phase C :" + phaseC_count + Environment.NewLine);

            if (!(Program.MSApp.ActiveModelReference.AnyElementsSelected))
                MessageBox.Show("Phase doesnt exist");
        }

        public static void selectConductor(string conType)
        {
            BCOM.Element oElem;// = default(BCOM.Element);
            BCOM.Fence ofence;// = default(Fence);
            BCOM.ElementEnumerator oEnum = default(BCOM.ElementEnumerator);
            int color = 0;
            string levelName = null;
            bool stats = false;
            stats = false;

            switch (conType)
            {
                case "Service Cable":
                    color = 0;
                    levelName = "Level 18";
                    break;

                case "LV Lines":
                    color = 2;
                    levelName = "Level 25";
                    break;

                case "MV Lines":
                    color = 3;
                    levelName = "Level 29";
                    break;

                case "Stats":
                    stats = true;
                    break;

                default:

                    MessageBox.Show("Invalid conductor");
                    break;
            }
            Program.MSApp.ActiveModelReference.UnselectAllElements();

            ofence = Program.MSApp.ActiveDesignFile.Fence;
            oEnum = ofence.GetContents();

            while (oEnum.MoveNext())
            {
                oElem = oEnum.Current;
                if ((oElem.Type == BCOM.MsdElementType.Line | oElem.Type == BCOM.MsdElementType.LineString))
                {
                    if (!stats)
                    {
                        if ((oElem.Color == color & oElem.Level.Name == levelName))
                        {
                            oElem.Redraw(BCOM.MsdDrawingMode.Hilite);
                            Program.MSApp.ActiveModelReference.SelectElement(oElem);
                        }
                    }
                    else
                    {
                        if ((oElem.Color == color & oElem.Level.Name == levelName))
                        {
                            oElem.Redraw(BCOM.MsdDrawingMode.Hilite);
                            Program.MSApp.ActiveModelReference.SelectElement(oElem);
                        }
                    }
                }
            }
        }

        public static void GetTransformerCellText(BCOM.CellElement ECell, string TrfrName, string Rating)
        {
            BCOM.ElementEnumerator oEE1;// = default(ElementEnumerator);
            BCOM.Element AElement1;//= default(Element);
            string Txt = null;
            string myString = null;
            oEE1 = ECell.GetSubElements();
            while ((oEE1.MoveNext()))
            {
                AElement1 = oEE1.Current;
                if ((AElement1.IsCellElement()))
                {
                    GetTransformerCellText(AElement1.AsCellElement(), TrfrName, Rating);
                }
                else if (AElement1.IsTextElement())
                {
                    Txt = AElement1.AsTextElement().Text;
                    // if (Strings.InStr(Strings.UCase(Txt), "KVA") > 0)
                    if (Txt.Contains("KVA"))
                    {
                        Rating = AElement1.AsTextElement().Text;
                        //myString = Strings.Left(Rating, Strings.Len(Rating) - 3);
                        myString = Rating.TrimEnd("KVA".ToCharArray());
                        Rating = myString;
                    }
                    else
                    {
                        TrfrName = AElement1.AsTextElement().Text;
                    }
                }
            }
        }

        //ByVal cellType As String)
        public static void selectTRFRzoneShapes()
        {
            BCOM.Element oElem;//= default(Element);
            BCOM.Fence ofence;//= default(Fence);
            BCOM.ElementEnumerator oEnum;// = default(ElementEnumerator);
            //Dim cellName As String
            //cellName = "001"
            Program.MSApp.ActiveModelReference.UnselectAllElements();
            ofence = Program.MSApp.ActiveDesignFile.Fence;
            oEnum = ofence.GetContents();

            while (oEnum.MoveNext())
            {
                oElem = oEnum.Current;

                if ((oElem.Type == BCOM.MsdElementType.Shape))
                {
                    if ((oElem.Level.Name == "Level 11"))
                    {
                        oElem.Redraw(BCOM.MsdDrawingMode.Hilite);
                        Program.MSApp.ActiveModelReference.SelectElement(oElem);
                    }
                }
            }
        }

        public static void selectZons(ref BCOM.ElementEnumerator oEnum, ref DataZone oSelZone)
        {
            BCOM.Element oElem = default(BCOM.Element);

            List<BCOM.Element> MVFlyingStays = new List<BCOM.Element>();
            //MVFlyingStay

            Program.MSApp.ActiveModelReference.UnselectAllElements();

            while (oEnum.MoveNext())
            {
                oElem = oEnum.Current;
                //Debug.Print CStr(oEnum.Current.ID.high) & CStr(oEnum.Current.ID.low)

                switch (oElem.Type)
                {
                    case BCOM.MsdElementType.CellHeader:
                        //Cells

                        switch (oElem.AsCellElement().Name)
                        {
                            case "001":
                                //LVStays
                                oSelZone.LVStays.Add(oElem.ID,oElem);// (oElem, Convert.ToString(oElem.ID));
                                break;

                            case "002":
                                //LVStruts
                                oSelZone.LVStruts.Add(oElem.ID, oElem);
                                break;

                            case "003":
                                //MVStays
                                oSelZone.MVStays.Add(oElem.ID, oElem);
                                break;

                            case "004":
                                //MVStruts
                                oSelZone.MVStruts.Add(oElem.ID, oElem);
                                break;

                            case "005":
                                //TRFRCircles
                                oSelZone.TRFRCircles.Add(oElem.ID, oElem);
                                break;

                            case "006":
                                //TRFRNameplates"
                                oSelZone.TRFRNameplates.Add(oElem.ID, oElem);
                                break;

                            case "009":
                                //LVpoles"
                                oSelZone.LVpoles.Add(oElem.ID, oElem);
                                break;

                            case "010":
                                //MVpoles"
                                oSelZone.MVpoles.Add(oElem.ID, oElem);
                                break;

                            case "011":
                                //MVSharingPoles
                                oSelZone.MVSharingPoles.Add(oElem.ID, oElem);
                                break;

                            case "012":
                                //KickerPoles"
                                oSelZone.KickerPoles.Add(oElem.ID, oElem);
                                break;

                            case "015":
                                //LVFlyingStays"
                                oSelZone.LVFlyingStays.Add(oElem.ID, oElem);
                                break;

                            case "016":
                                //LVShortStays
                                oSelZone.LVShortStays.Add(oElem.ID, oElem);
                                break;

                            case "017":
                                //MVFlyingStays"
                                oSelZone.MVFlyingStays.Add(oElem.ID, oElem);
                                break;

                            case "018":
                                //MVShortStays
                                oSelZone.MVShortStays.Add(oElem.ID, oElem);
                                break;

                            case "035":
                                //MVFuseIsolators
                                oSelZone.MVFuseIsolators.Add(oElem.ID, oElem);
                                break;

                            case "TXTCS":
                                //TRFRs
                                oSelZone.TRFRs.Add(oElem.ID, oElem);
                                break;

                            default:
                                oSelZone.elseCellElements.Add(oElem.ID, oElem);
                                break;
                        }
                        break;

                    case BCOM.MsdElementType.Line:

                        switch (oElem.Level.Name)
                        {
                            case "Level 18":
                                // Airdac
                                oSelZone.Airdac.Add(oElem.ID, oElem);
                                break;
                            //Debug.Print CStr(oEnum.Current.ID.high) & CStr(oEnum.Current.ID.low)
                            case "Level 25":
                                //LVConductor
                                oSelZone.LVConductor.Add(oElem.ID, oElem);
                                break;

                            case "Level 29":
                                //MVConductor
                                oSelZone.MVConductor.Add(oElem.ID, oElem);
                                break;

                            default:
                                oSelZone.elseLineElements.Add(oElem.ID, oElem);
                                break;
                        }
                        break;

                    case BCOM.MsdElementType.Shape:

                        switch (oElem.Level.Name)
                        {
                            case "Level 11":
                                //TrfrShapes
                                oSelZone.TrfrShapes.Add(oElem.ID, oElem);
                                break;

                            default:
                                oSelZone.elseShapeElements.Add(oElem.ID, oElem);
                                break;
                        }
                        break;

                    case (BCOM.MsdElementType.Text):

                        switch (oElem.Level.Name)
                        {
                            case "Level 3":
                                //houseText
                                oSelZone.houseText.Add(oElem.ID, oElem);
                                break;

                            default:
                                oSelZone.elseTextElements.Add(oElem.ID, oElem);
                                break;
                        }
                        break;

                    default:
                        oSelZone.elseElements.Add(oElem.ID, oElem);
                        break;
                }
            }

            Debug.Print(oSelZone.ToString());
            Debug.Print(oSelZone.mergedMVPolesList().Count.ToString());

            Debug.Print(oSelZone.printCoordinatesMVPoles());
            Debug.Print(oSelZone.printCoordinatesMVLines());
        }

        //public void selectZons(ref BCOM.ElementEnumerator oEnum, ref DataZone oSelZone)
        //{
        //    BCOM.Element oElem = default(BCOM.Element);

        //    List<BCOM.Element> MVFlyingStays = new List<BCOM.Element>();
        //    //MVFlyingStay

        //    Program.MSApp.ActiveModelReference.UnselectAllElements();

        //    while (oEnum.MoveNext())
        //    {
        //        oElem = oEnum.Current;
        //        //Debug.Print CStr(oEnum.Current.ID.high) & CStr(oEnum.Current.ID.low)

        //        switch (oElem.Type)
        //        {
        //            case BCOM.MsdElementType.CellHeader:
        //                //Cells

        //                switch (oElem.AsCellElement().Name)
        //                {
        //                    case "001":
        //                        //LVStays
        //                        oSelZone.LVStays.Add(oElem.ID, oElem);// (oElem, Convert.ToString(oElem.ID));
        //                        break;

        //                    case "002":
        //                        //LVStruts
        //                        oSelZone.LVStruts.Add(oElem.ID, oElem);
        //                        break;

        //                    case "003":
        //                        //MVStays
        //                        oSelZone.MVStays.Add(oElem.ID, oElem);
        //                        break;

        //                    case "004":
        //                        //MVStruts
        //                        oSelZone.MVStruts.Add(oElem.ID, oElem);
        //                        break;

        //                    case "005":
        //                        //TRFRCircles
        //                        oSelZone.TRFRCircles.Add(oElem.ID, oElem);
        //                        break;

        //                    case "006":
        //                        //TRFRNameplates"
        //                        oSelZone.TRFRNameplates.Add(oElem.ID, oElem);
        //                        break;

        //                    case "009":
        //                        //LVpoles"
        //                        oSelZone.LVpoles.Add(oElem.ID, oElem);
        //                        break;

        //                    case "010":
        //                        //MVpoles"
        //                        oSelZone.MVpoles.Add(oElem.ID, oElem);
        //                        break;

        //                    case "011":
        //                        //MVSharingPoles
        //                        oSelZone.MVSharingPoles.Add(oElem.ID, oElem);
        //                        break;

        //                    case "012":
        //                        //KickerPoles"
        //                        oSelZone.KickerPoles.Add(oElem.ID, oElem);
        //                        break;

        //                    case "015":
        //                        //LVFlyingStays"
        //                        oSelZone.LVFlyingStays.Add(oElem.ID, oElem);
        //                        break;

        //                    case "016":
        //                        //LVShortStays
        //                        oSelZone.LVShortStays.Add(oElem.ID, oElem);
        //                        break;

        //                    case "017":
        //                        //MVFlyingStays"
        //                        oSelZone.MVFlyingStays.Add(oElem.ID, oElem);
        //                        break;

        //                    case "018":
        //                        //MVShortStays
        //                        oSelZone.MVShortStays.Add(oElem.ID, oElem);
        //                        break;

        //                    case "035":
        //                        //MVFuseIsolators
        //                        oSelZone.MVFuseIsolators.Add(oElem.ID, oElem);
        //                        break;

        //                    case "TXTCS":
        //                        //TRFRs
        //                        oSelZone.TRFRs.Add(oElem.ID, oElem);
        //                        break;

        //                    default:
        //                        oSelZone.elseCellElements.Add(oElem.ID, oElem);
        //                        break;
        //                }
        //                break;

        //            case BCOM.MsdElementType.Line:

        //                switch (oElem.Level.Name)
        //                {
        //                    case "Level 18":
        //                        // Airdac
        //                        oSelZone.Airdac.Add(oElem.ID, oElem);
        //                        break;
        //                    //Debug.Print CStr(oEnum.Current.ID.high) & CStr(oEnum.Current.ID.low)
        //                    case "Level 25":
        //                        //LVConductor
        //                        oSelZone.LVConductor.Add(oElem.ID, oElem);
        //                        break;

        //                    case "Level 29":
        //                        //MVConductor
        //                        oSelZone.MVConductor.Add(oElem.ID, oElem);
        //                        break;

        //                    default:
        //                        oSelZone.elseLineElements.Add(oElem.ID, oElem);
        //                        break;
        //                }
        //                break;

        //            case BCOM.MsdElementType.Shape:

        //                switch (oElem.Level.Name)
        //                {
        //                    case "Level 11":
        //                        //TrfrShapes
        //                        oSelZone.TrfrShapes.Add(oElem.AsShapeElement().get_Vertex(0), oElem);
        //                        break;

        //                    default:
        //                        oSelZone.elseShapeElements.Add(oElem.AsShapeElement().get_Vertex(0), oElem);
        //                        break;
        //                }
        //                break;

        //            case (BCOM.MsdElementType.Text):

        //                switch (oElem.Level.Name)
        //                {
        //                    case "Level 3":
        //                        //houseText
        //                        oSelZone.houseText.Add(oElem.AsTextElement().get_Origin(), oElem);
        //                        break;

        //                    default:
        //                        oSelZone.elseTextElements.Add(oElem.AsTextElement().get_Origin(), oElem);
        //                        break;
        //                }
        //                break;

        //            default:
        //                oSelZone.elseElements.Add(oElem.ID, oElem);
        //                break;
        //        }
        //    }

        //    Debug.Print(oSelZone.ToString());
        //    Debug.Print(oSelZone.mergedMVPolesList().Count.ToString());

        //    Debug.Print(oSelZone.printCoordinatesMVPoles);
        //    Debug.Print(oSelZone.printCoordinatesMVLines);
        //}
    }
}